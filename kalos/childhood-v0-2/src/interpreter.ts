import { CANON_CONSTRAINTS } from './canon.js';
import type {
  BeatDefinition,
  ChronicleState,
  InterpreterContext,
  OutcomeProposal,
  PatchOperation,
  ProposalValidation,
} from './schema.js';

export const ALLOWED_PATCH_PREFIXES = [
  'characters.',
  'kalos.body.',
  'kalos.tendencies.',
  'kalos.skills.',
  'kalos.attention',
  'world.resources.',
  'world.persistentDamage',
  'publicStories',
  'memories',
  'promises',
  'opportunities',
  'presentation.',
  'flags.',
];

const blockedChildContent = /\b(sex|sexual|nude|naked|erotic|climax|orgasm|rape|seduce)\b/i;
const severeEvent = /\b(kill|death|dies|dead|amputate|permanent exile|burns down|drowns)\b/i;

function pathAllowed(path: string): boolean {
  return ALLOWED_PATCH_PREFIXES.some((prefix) => path.startsWith(prefix));
}

function operationMagnitudeAllowed(operation: PatchOperation): boolean {
  if (operation.op !== 'increment') return true;
  if (operation.path.includes('.relationships.') || /\.(trust|affection|respect|ease|resentment|fear|dependence|obligation)$/.test(operation.path)) {
    return Math.abs(operation.value) <= 0.18;
  }
  return Math.abs(operation.value) <= 0.25;
}

export function validateProposal(
  proposal: OutcomeProposal,
  state: ChronicleState,
  knownBeatIds: ReadonlySet<string>,
): ProposalValidation {
  if (proposal.contractVersion !== '1.0') {
    return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Unsupported interpreter contract.' };
  }
  if (blockedChildContent.test(proposal.actionText) || blockedChildContent.test(proposal.novelOutcome?.body ?? '')) {
    return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Childhood content boundary violated.' };
  }
  if (proposal.nextBeatId && !knownBeatIds.has(proposal.nextBeatId)) {
    return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Unknown next beat.' };
  }
  if (proposal.novelOutcome && !knownBeatIds.has(proposal.novelOutcome.returnToBeatId)) {
    return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Novel outcome lacks a valid return beat.' };
  }
  if (proposal.objectiveEvent) {
    const combined = `${proposal.objectiveEvent.summary} ${proposal.objectiveEvent.facts.join(' ')}`;
    if (severeEvent.test(combined)) {
      return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Severe events require an authored event rule.' };
    }
    if (proposal.objectiveEvent.sceneId !== state.sceneId) {
      return { accepted: false, acceptedPatch: [], rejectedPatch: [], reason: 'Objective event belongs to another scene.' };
    }
  }

  const acceptedPatch: PatchOperation[] = [];
  const rejectedPatch: Array<{ operation: PatchOperation; reason: string }> = [];
  for (const operation of proposal.patch) {
    if (!pathAllowed(operation.path)) {
      rejectedPatch.push({ operation, reason: 'Path is outside the interpreter’s authority.' });
      continue;
    }
    if (!operationMagnitudeAllowed(operation)) {
      rejectedPatch.push({ operation, reason: 'Single-turn effect exceeds the allowed magnitude.' });
      continue;
    }
    acceptedPatch.push(operation);
  }

  const hasAction = Boolean(proposal.selectedChoiceId || proposal.nextBeatId || proposal.novelOutcome);
  return {
    accepted: hasAction,
    acceptedPatch,
    rejectedPatch,
    reason: hasAction ? undefined : 'Proposal does not produce an actionable outcome.',
  };
}

function scoreChoice(input: string, choice: BeatDefinition['choices'][number]): number {
  const normalized = input.toLowerCase();
  const words = `${choice.label} ${choice.description} ${choice.axes?.join(' ') ?? ''}`
    .toLowerCase()
    .split(/[^a-z]+/)
    .filter((word) => word.length > 2);
  let score = 0;
  for (const word of words) {
    if (normalized.includes(word)) score += word.length > 6 ? 3 : 1;
  }
  const rules: Array<[RegExp, string[]]> = [
    [/\b(confess|admit|truth|my fault|i did)\b/, ['confess', 'own', 'tell']],
    [/\b(joke|laugh|funny|mock|tease)\b/, ['joke', 'laugh', 'humor']],
    [/\b(help|protect|share|carry|comfort)\b/, ['help', 'protect', 'share']],
    [/\b(wait|finish|work|careful|bind)\b/, ['finish', 'work', 'careful']],
    [/\b(leave|run|race|shore|go)\b/, ['leave', 'race', 'shore']],
    [/\b(silent|nothing|hide|deny|pretend)\b/, ['silent', 'hide', 'deny']],
    [/\b(test|inspect|look|evidence|check)\b/, ['test', 'inspect', 'check']],
  ];
  for (const [pattern, hints] of rules) {
    if (pattern.test(normalized) && hints.some((hint) => words.includes(hint))) score += 5;
  }
  return score;
}

export function interpretOffline(input: string, beat: BeatDefinition): OutcomeProposal {
  const ranked = beat.choices
    .map((choice) => ({ choice, score: scoreChoice(input, choice) }))
    .sort((a, b) => b.score - a.score);
  const selected = ranked[0];
  const confidence = selected && selected.score > 0 ? Math.min(0.92, 0.38 + selected.score / 20) : 0.2;
  return {
    contractVersion: '1.0',
    proposalId: `offline-${Date.now().toString(36)}`,
    source: 'offline',
    confidence,
    intent: {
      disclosure: /confess|admit|truth/.test(input.toLowerCase()) ? 'full' : 'not-applicable',
      tactic: selected?.choice.id ?? 'unresolved',
      target: 'group',
      desiredEffect: input,
    },
    actionText: input,
    selectedChoiceId: confidence >= 0.35 ? selected?.choice.id : undefined,
    patch: [],
    rationaleTags: selected ? [`matched:${selected.choice.id}`] : ['ambiguous'],
  };
}

export async function interpretPlayerInput(
  input: string,
  beat: BeatDefinition,
  state: ChronicleState,
  knownBeatIds: ReadonlySet<string>,
): Promise<{ proposal: OutcomeProposal; validation: ProposalValidation; fallbackUsed: boolean }> {
  const deterministicProposal = interpretOffline(input, beat);
  const adapter = window.KALOS_LLM_ADAPTER;
  if (state.settings.interpreterMode !== 'hybrid' || !adapter) {
    return {
      proposal: deterministicProposal,
      validation: validateProposal(deterministicProposal, state, knownBeatIds),
      fallbackUsed: false,
    };
  }

  const context: InterpreterContext = {
    contractVersion: '1.0',
    input,
    scene: { id: beat.id, sceneId: beat.sceneId, title: beat.title, choices: beat.choices },
    state: structuredClone(state),
    deterministicProposal,
    allowedPatchPrefixes: ALLOWED_PATCH_PREFIXES,
    canonConstraints: CANON_CONSTRAINTS,
  };

  try {
    const proposal = await adapter.interpret(context);
    const validation = validateProposal(proposal, state, knownBeatIds);
    if (validation.accepted) return { proposal, validation, fallbackUsed: false };
  } catch {
    // The offline interpreter is a complete runtime, not a degraded emergency mode.
  }

  return {
    proposal: deterministicProposal,
    validation: validateProposal(deterministicProposal, state, knownBeatIds),
    fallbackUsed: true,
  };
}
