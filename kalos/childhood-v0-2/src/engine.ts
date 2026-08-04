import { createInitialState } from './canon.js';
import type {
  BeatDefinition,
  CharacterId,
  ChronicleState,
  HistoryRecord,
  ObjectiveEvent,
  OutcomeProposal,
  PatchOperation,
  RelationshipState,
} from './schema.js';

export const SAVE_KEY = 'kalos.childhood.v0.2';
export const LEGACY_SAVE_KEY = 'kalos-hidden-hook';

const clamp = (value: number, min = 0, max = 1): number => Math.max(min, Math.min(max, value));

function getAtPath(root: unknown, path: string): unknown {
  return path.split('.').reduce<unknown>((node, key) => {
    if (node && typeof node === 'object') return (node as Record<string, unknown>)[key];
    return undefined;
  }, root);
}

function setAtPath(root: unknown, path: string, value: unknown): void {
  const keys = path.split('.');
  let node = root as Record<string, unknown>;
  for (let index = 0; index < keys.length - 1; index += 1) {
    const key = keys[index];
    if (!key) continue;
    const child = node[key];
    if (!child || typeof child !== 'object') node[key] = {};
    node = node[key] as Record<string, unknown>;
  }
  const finalKey = keys.at(-1);
  if (finalKey) node[finalKey] = value;
}

export function applyPatch(state: ChronicleState, patch: PatchOperation[]): ChronicleState {
  const next = structuredClone(state);
  for (const operation of patch) {
    if (operation.op === 'set') {
      setAtPath(next, operation.path, operation.value);
      continue;
    }
    if (operation.op === 'increment') {
      const current = getAtPath(next, operation.path);
      if (typeof current !== 'number') continue;
      setAtPath(next, operation.path, clamp(current + operation.value));
      continue;
    }
    if (operation.op === 'append') {
      const current = getAtPath(next, operation.path);
      if (Array.isArray(current)) current.push(operation.value);
    }
  }
  return next;
}

export function appendObjectiveEvent(state: ChronicleState, event: Omit<ObjectiveEvent, 'immutable'>): ChronicleState {
  const next = structuredClone(state);
  next.events.push({ ...event, immutable: true });
  return next;
}

export function relationship(
  state: ChronicleState,
  observer: CharacterId,
  subject: CharacterId,
): RelationshipState | undefined {
  return state.characters[observer].relationships[subject];
}

export function adjustRelationship(
  state: ChronicleState,
  observer: CharacterId,
  subject: CharacterId,
  changes: Partial<Omit<RelationshipState, 'sharedHistory'>>,
  sharedHistory?: string,
): ChronicleState {
  const next = structuredClone(state);
  const existing = next.characters[observer].relationships[subject];
  if (!existing) return next;
  for (const [dimension, delta] of Object.entries(changes)) {
    const key = dimension as keyof Omit<RelationshipState, 'sharedHistory'>;
    const current = existing[key];
    if (typeof current === 'number' && typeof delta === 'number') existing[key] = clamp(current + delta);
  }
  if (sharedHistory) existing.sharedHistory.push(sharedHistory);
  return next;
}

export function recordHistory(
  state: ChronicleState,
  beat: BeatDefinition,
  choiceId: string,
  actionText: string,
  interpreter: HistoryRecord['interpreter'],
  proposal?: OutcomeProposal,
): ChronicleState {
  const next = structuredClone(state);
  next.history.push({
    id: `history-${next.history.length + 1}-${Date.now().toString(36)}`,
    sceneId: beat.sceneId,
    beatId: beat.id,
    choiceId,
    actionText,
    interpreter,
    acceptedProposal: proposal?.proposalId,
    timestamp: new Date().toISOString(),
  });
  return next;
}

export function applyAcceptedProposal(
  state: ChronicleState,
  proposal: OutcomeProposal,
  acceptedPatch: PatchOperation[],
): ChronicleState {
  let next = applyPatch(state, acceptedPatch);
  if (proposal.objectiveEvent) next = appendObjectiveEvent(next, proposal.objectiveEvent);
  if (proposal.nextBeatId) next.beatId = proposal.nextBeatId;
  if (proposal.novelOutcome) {
    next.flags.pendingNovelOutcome = proposal.novelOutcome.body;
    next.flags.pendingNovelReturnBeat = proposal.novelOutcome.returnToBeatId;
  }
  return next;
}

export function saveState(state: ChronicleState): void {
  localStorage.setItem(SAVE_KEY, JSON.stringify(state));
}

function migrateLegacy(raw: unknown): ChronicleState {
  const next = createInitialState();
  next.migratedFromLegacy = true;
  next.flags.oldSaveArchived = true;
  if (raw && typeof raw === 'object') {
    const legacy = raw as Record<string, unknown>;
    const path = Array.isArray(legacy.path) ? legacy.path.map(String) : [];
    const publicAccount = typeof legacy.publicAccount === 'string' ? legacy.publicAccount : '';
    next.memories.push({
      id: 'legacy-hidden-hook-memory',
      owner: 'kalos',
      text: publicAccount || `A prior telling of the Hidden Hook survived as choices: ${path.join(', ') || 'unknown'}.`,
      sensoryAnchor: 'a bone hook warming in his palm',
      salience: 0.64,
      accuracy: 0.62,
      meaning: 'a prior chronicle archived before the childhood expansion',
    });
  }
  return next;
}

export function loadState(): ChronicleState {
  const current = localStorage.getItem(SAVE_KEY);
  if (current) {
    try {
      const parsed = JSON.parse(current) as ChronicleState;
      if (parsed.schemaVersion === 2) return parsed;
    } catch {
      localStorage.removeItem(SAVE_KEY);
    }
  }
  const legacy = localStorage.getItem(LEGACY_SAVE_KEY);
  if (legacy) {
    try {
      return migrateLegacy(JSON.parse(legacy));
    } catch {
      return migrateLegacy(undefined);
    }
  }
  return createInitialState();
}

export function resetState(seed?: number): ChronicleState {
  localStorage.removeItem(SAVE_KEY);
  return createInitialState(seed);
}

export function exportState(state: ChronicleState): Blob {
  return new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
}

export function importState(text: string): ChronicleState {
  const parsed = JSON.parse(text) as ChronicleState;
  if (parsed.schemaVersion !== 2) throw new Error('This save uses an unsupported schema version.');
  return parsed;
}

export function qualitativeRelationship(value: number): string {
  if (value >= 0.82) return 'deep';
  if (value >= 0.66) return 'strong';
  if (value >= 0.52) return 'growing';
  if (value >= 0.38) return 'uncertain';
  if (value >= 0.22) return 'thin';
  return 'broken';
}
