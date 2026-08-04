export type CharacterId =
  | 'kalos'
  | 'eda'
  | 'ruvan'
  | 'seli'
  | 'pali'
  | 'veya'
  | 'mara'
  | 'neri'
  | 'oren';

export type SceneId =
  | 'smoke-before-daylight'
  | 'the-drying-racks'
  | 'race-beneath-the-fish'
  | 'the-eighth-hook'
  | 'the-work-nobody-saw'
  | 'season-close';

export type ChoiceTier = 'light' | 'standard' | 'hinge';
export type InterpreterMode = 'offline' | 'hybrid';
export type Weather = 'clear-cold' | 'low-cloud' | 'mist' | 'light-rain';
export type Tide = 'rising' | 'high' | 'falling' | 'low';

export interface RelationshipState {
  affection: number;
  trust: number;
  respect: number;
  ease: number;
  resentment: number;
  dependence: number;
  fear: number;
  obligation: number;
  sharedHistory: string[];
}

export interface BeliefState {
  proposition: string;
  confidence: number;
  source: 'witnessed' | 'inferred' | 'heard' | 'remembered';
  eventId?: string;
}

export interface CharacterState {
  id: CharacterId;
  name: string;
  age: number;
  role: string;
  present: boolean;
  currentGoal: string;
  schedule: string;
  relationships: Partial<Record<CharacterId, RelationshipState>>;
  beliefs: BeliefState[];
  knowledge: string[];
}

export interface BodyState {
  hunger: number;
  fatigue: number;
  cold: number;
  pain: number;
  arousal: number;
  injuries: string[];
}

export interface TendencyState {
  followThrough: number;
  humorUnderPressure: number;
  protectiveness: number;
  statusSensitivity: number;
  strategicOmission: number;
  curiosity: number;
  shameTolerance: number;
  practicalPatience: number;
}

export interface SkillState {
  hookBinding: number;
  fiberPreparation: number;
  socialReading: number;
  storytelling: number;
  balance: number;
  swimming: number;
  caregiving: number;
}

export interface KalosState {
  age: 8;
  body: BodyState;
  tendencies: TendencyState;
  skills: SkillState;
  attention: string[];
}

export interface WorldState {
  day: number;
  time: 'predawn' | 'morning' | 'midday' | 'afternoon' | 'dusk' | 'night';
  weather: Weather;
  tide: Tide;
  location: string;
  resources: {
    driedFish: number;
    lampOil: number;
    bindingFiber: number;
    intactHooks: number;
  };
  persistentDamage: string[];
}

export interface ObjectiveEvent {
  id: string;
  sceneId: SceneId;
  summary: string;
  facts: string[];
  witnesses: CharacterId[];
  tags: string[];
  immutable: true;
}

export interface PublicStory {
  id: string;
  text: string;
  source: CharacterId | 'household';
  reach: number;
  distortion: number;
  relatedEventId?: string;
}

export interface MemoryState {
  id: string;
  owner: CharacterId;
  text: string;
  sensoryAnchor?: string;
  salience: number;
  accuracy: number;
  meaning: string;
  relatedEventId?: string;
}

export interface PromiseState {
  id: string;
  speaker: CharacterId;
  beneficiary: CharacterId | 'household';
  text: string;
  status: 'open' | 'kept' | 'broken' | 'partially-repaired';
  witnesses: CharacterId[];
  repetitions: number;
}

export interface OpportunityState {
  id: string;
  label: string;
  available: boolean;
  reason: string;
  expiresAfterScene?: SceneId;
}

export interface VisualCharacterState {
  id: CharacterId;
  position: 'foreground-left' | 'foreground-right' | 'center' | 'background-left' | 'background-right';
  posture: string;
  gaze: CharacterId | 'object' | 'away';
  proximity: 'isolated' | 'near' | 'touching' | 'crowded';
}

export interface VisualState {
  plate: 'longhouse' | 'drying-racks' | 'race' | 'hook-close' | 'work-shed' | 'inlet';
  title: string;
  subtitle: string;
  palette: 'ember-smoke' | 'salt-daylight' | 'wind-shadow' | 'bone-fiber' | 'empty-tide' | 'low-fire';
  framing: 'wide' | 'medium' | 'close' | 'split';
  weather: Weather;
  tide: Tide;
  light: string;
  motion: string[];
  focalObjects: string[];
  characters: VisualCharacterState[];
  sensoryPriority: Array<'touch' | 'smell' | 'sound' | 'temperature' | 'balance' | 'taste' | 'sight'>;
  audio: string[];
  proseMode: 'warm-observational' | 'technical' | 'playful-fast' | 'narrowing' | 'quiet-aftermath';
}

export interface ChoiceDefinition {
  id: string;
  label: string;
  description: string;
  tier: ChoiceTier;
  axes?: string[];
}

export interface BeatDefinition {
  id: string;
  sceneId: SceneId;
  title: string;
  kicker: string;
  body: string;
  choicePrompt: string;
  choiceHint: string;
  choices: ChoiceDefinition[];
  freeText: boolean;
  visual: VisualState;
}

export interface HistoryRecord {
  id: string;
  sceneId: SceneId;
  beatId: string;
  choiceId: string;
  actionText: string;
  interpreter: 'menu' | 'offline' | 'llm';
  acceptedProposal?: string;
  timestamp: string;
}

export interface SettingsState {
  interpreterMode: InterpreterMode;
  readingFocus: boolean;
  textScale: 'small' | 'normal' | 'large';
  reducedMotion: boolean;
  showChoiceAxes: boolean;
}

export interface ChronicleState {
  schemaVersion: 2;
  chronicleId: string;
  seed: number;
  sceneId: SceneId;
  beatId: string;
  sceneOrdinal: number;
  completed: boolean;
  migratedFromLegacy: boolean;
  world: WorldState;
  kalos: KalosState;
  characters: Record<CharacterId, CharacterState>;
  events: ObjectiveEvent[];
  publicStories: PublicStory[];
  memories: MemoryState[];
  promises: PromiseState[];
  opportunities: OpportunityState[];
  presentation: VisualState;
  history: HistoryRecord[];
  flags: Record<string, boolean | number | string>;
  settings: SettingsState;
}

export type PatchOperation =
  | { op: 'increment'; path: string; value: number }
  | { op: 'set'; path: string; value: string | number | boolean | string[] }
  | { op: 'append'; path: string; value: unknown };

export interface InterpreterContext {
  contractVersion: '1.0';
  input: string;
  scene: Pick<BeatDefinition, 'id' | 'sceneId' | 'title' | 'choices'>;
  state: ChronicleState;
  deterministicProposal: OutcomeProposal;
  allowedPatchPrefixes: string[];
  canonConstraints: string[];
}

export interface OutcomeProposal {
  contractVersion: '1.0';
  proposalId: string;
  source: 'offline' | 'llm';
  confidence: number;
  intent: {
    disclosure: 'full' | 'partial' | 'none' | 'deceptive' | 'not-applicable';
    tactic: string;
    target: CharacterId | 'group' | 'object' | 'self';
    desiredEffect: string;
  };
  actionText: string;
  selectedChoiceId?: string;
  nextBeatId?: string;
  novelOutcome?: {
    body: string;
    returnToBeatId: string;
  };
  patch: PatchOperation[];
  objectiveEvent?: Omit<ObjectiveEvent, 'immutable'>;
  rationaleTags: string[];
}

export interface ProposalValidation {
  accepted: boolean;
  acceptedPatch: PatchOperation[];
  rejectedPatch: Array<{ operation: PatchOperation; reason: string }>;
  reason?: string;
}

declare global {
  interface Window {
    KALOS_LLM_ADAPTER?: {
      name: string;
      interpret(context: InterpreterContext): Promise<OutcomeProposal>;
    };
    Kalos?: {
      registerInterpreter(adapter: NonNullable<Window['KALOS_LLM_ADAPTER']>): void;
      getState(): ChronicleState;
    };
  }
}
