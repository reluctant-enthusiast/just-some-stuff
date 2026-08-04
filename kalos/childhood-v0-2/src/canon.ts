import type {
  CharacterId,
  CharacterState,
  ChronicleState,
  RelationshipState,
  VisualState,
} from './schema.js';

const rel = (overrides: Partial<RelationshipState> = {}): RelationshipState => ({
  affection: 0.45,
  trust: 0.45,
  respect: 0.35,
  ease: 0.4,
  resentment: 0,
  dependence: 0,
  fear: 0,
  obligation: 0,
  sharedHistory: [],
  ...overrides,
});

const character = (
  id: CharacterId,
  name: string,
  age: number,
  role: string,
  currentGoal: string,
  schedule: string,
): CharacterState => ({
  id,
  name,
  age,
  role,
  present: true,
  currentGoal,
  schedule,
  relationships: {},
  beliefs: [],
  knowledge: [],
});

export const CANON_CONSTRAINTS = [
  'Kalos is eight years old. Childhood must never be eroticized.',
  'No character may know facts they did not witness, infer, hear, or remember.',
  'Objective history may be appended but never rewritten retroactively.',
  'An LLM may create a new causal outcome only through a validated proposal.',
  'No single action may change a relationship dimension by more than 0.18.',
  'Death, severe injury, major resource loss, or permanent exile require an authored event rule.',
  'Kalos has strong tendencies, not fixed rails; positive counter-patterns remain possible.',
  'The world and other people continue outside Kalos’s attention.',
  'Prose and art render experience; they do not determine objective state.',
];

export const PROVISIONAL_CAST_NOTE =
  'Only Kalos is a locked name. Household names and exact kinship rules remain provisional canon.';

export const initialVisual: VisualState = {
  plate: 'longhouse',
  title: 'Smoke Before Daylight',
  subtitle: 'Late summer · before the household rises',
  palette: 'ember-smoke',
  framing: 'wide',
  weather: 'low-cloud',
  tide: 'falling',
  light: 'low fire under blue predawn',
  motion: ['roof smoke folding beneath the beams', 'sleeping bodies shifting'],
  focalObjects: ['banked fire', 'Pali’s tangled blanket', 'a strip of crisp fish skin'],
  characters: [
    { id: 'kalos', position: 'foreground-left', posture: 'awake beneath a shared blanket', gaze: 'pali', proximity: 'near' },
    { id: 'pali', position: 'center', posture: 'sleeping cold with one foot uncovered', gaze: 'away', proximity: 'near' },
    { id: 'seli', position: 'background-right', posture: 'already dressing for work', gaze: 'object', proximity: 'isolated' },
  ],
  sensoryPriority: ['temperature', 'smell', 'sound', 'touch', 'sight'],
  audio: ['small fire settling', 'sleep breathing', 'water under the house pilings'],
  proseMode: 'warm-observational',
};

export function createInitialState(seed = Date.now() % 2147483647): ChronicleState {
  const characters = {
    kalos: character('kalos', 'Kalos', 8, 'child of the Long Roof household', 'find a place among older children without losing adult trust', 'moves between household tasks, the drying yard, and the shore'),
    eda: character('eda', 'Eda', 33, 'Kalos’s mother; host and food-work organizer', 'stretch late-summer abundance without making the household feel poor', 'predawn food work, midday distribution, evening hosting'),
    ruvan: character('ruvan', 'Ruvan', 37, 'Kalos’s father; canoe and storage worker', 'finish an outer-inlet repair before the weather changes', 'away before dawn; expected after midday'),
    seli: character('seli', 'Seli', 11, 'Kalos’s older sister and increasingly trusted worker', 'earn adult work without becoming responsible for every younger child', 'fire, drying racks, errands, private time near dusk'),
    pali: character('pali', 'Pali', 4, 'younger foster cousin sleeping beside Kalos', 'remain close to familiar people and avoid being sent with strangers', 'follows household caregivers; sleeps unpredictably'),
    veya: character('veya', 'Veya', 58, 'Ruvan’s elder sister; household memory and practical authority', 'keep an old obligation from becoming a public dispute', 'hearth work, visitors, evening stories'),
    mara: character('mara', 'Mara', 35, 'fishing-gear supervisor and exacting teacher', 'prepare reliable gear before the morning tide', 'drying yard and gear shed'),
    neri: character('neri', 'Neri', 9, 'child of a dependent household; precise with fiber', 'be trusted for work without accepting humiliating favors', 'drying yard, water carrying, younger-child duty'),
    oren: character('oren', 'Oren', 11, 'older cousin seeking adult notice', 'look useful and socially central before his father returns', 'line baskets, peer games, shore errands'),
  } satisfies Record<CharacterId, CharacterState>;

  characters.kalos.relationships = {
    eda: rel({ affection: 0.72, trust: 0.56, ease: 0.68, dependence: 0.48 }),
    ruvan: rel({ affection: 0.61, trust: 0.55, respect: 0.58, ease: 0.41 }),
    seli: rel({ affection: 0.58, trust: 0.48, ease: 0.52, resentment: 0.08 }),
    pali: rel({ affection: 0.66, trust: 0.61, dependence: 0.44, ease: 0.62 }),
    veya: rel({ affection: 0.49, trust: 0.48, respect: 0.63, ease: 0.31 }),
    mara: rel({ affection: 0.47, trust: 0.54, respect: 0.55, ease: 0.34 }),
    neri: rel({ affection: 0.43, trust: 0.45, respect: 0.39, ease: 0.37, sharedHistory: ['Neri once gave Kalos the larger half of a roasted root.'] }),
    oren: rel({ affection: 0.49, trust: 0.36, respect: 0.38, ease: 0.58, resentment: 0.04 }),
  };

  for (const id of Object.keys(characters) as CharacterId[]) {
    if (id === 'kalos') continue;
    characters[id].relationships.kalos = rel();
  }
  characters.eda.relationships.kalos = rel({ affection: 0.78, trust: 0.57, dependence: 0.3 });
  characters.seli.relationships.kalos = rel({ affection: 0.57, trust: 0.44, ease: 0.49, resentment: 0.11 });
  characters.pali.relationships.kalos = rel({ affection: 0.73, trust: 0.66, dependence: 0.54 });
  characters.mara.relationships.kalos = rel({ affection: 0.46, trust: 0.55, respect: 0.34 });
  characters.neri.relationships.kalos = rel({ affection: 0.41, trust: 0.46, respect: 0.37, wariness: undefined } as never);
  characters.oren.relationships.kalos = rel({ affection: 0.52, trust: 0.34, ease: 0.61, resentment: 0.05 });

  return {
    schemaVersion: 2,
    chronicleId: `kalos-${seed.toString(36)}-${Date.now().toString(36)}`,
    seed,
    sceneId: 'smoke-before-daylight',
    beatId: 's1-waking',
    sceneOrdinal: 1,
    completed: false,
    migratedFromLegacy: false,
    world: {
      day: 1,
      time: 'predawn',
      weather: 'low-cloud',
      tide: 'falling',
      location: 'Long Roof household',
      resources: { driedFish: 72, lampOil: 41, bindingFiber: 63, intactHooks: 19 },
      persistentDamage: [],
    },
    kalos: {
      age: 8,
      body: { hunger: 0.48, fatigue: 0.2, cold: 0.37, pain: 0, arousal: 0.16, injuries: [] },
      tendencies: {
        followThrough: 0.45,
        humorUnderPressure: 0.63,
        protectiveness: 0.56,
        statusSensitivity: 0.65,
        strategicOmission: 0.36,
        curiosity: 0.68,
        shameTolerance: 0.38,
        practicalPatience: 0.36,
      },
      skills: {
        hookBinding: 0.3,
        fiberPreparation: 0.36,
        socialReading: 0.62,
        storytelling: 0.55,
        balance: 0.61,
        swimming: 0.48,
        caregiving: 0.39,
      },
      attention: ['Pali’s uncovered foot', 'the smell of crisp skin near the coals', 'Seli leaving before anyone asks her'],
    },
    characters,
    events: [],
    publicStories: [],
    memories: [
      {
        id: 'memory-root-smoke',
        owner: 'kalos',
        text: 'Predawn smoke made the roof disappear before the people beneath it did.',
        sensoryAnchor: 'cold air and fish fat near the coals',
        salience: 0.42,
        accuracy: 0.92,
        meaning: 'home before obligation had fully begun',
      },
    ],
    promises: [],
    opportunities: [
      { id: 'mara-apprenticeship', label: 'Mara may teach Kalos more exact gear work', available: false, reason: 'Not yet earned.' },
      { id: 'older-children-shore', label: 'Join the older children at the shore', available: true, reason: 'Oren enjoys Kalos’s timing.' },
    ],
    presentation: initialVisual,
    history: [],
    flags: {
      paliHelpedAtDawn: false,
      overheardAdultConcern: false,
      ateCrispSkin: false,
      sharedBreakfast: false,
      workMethod: 'unselected',
      raceRoute: 'unselected',
      neriHumiliated: false,
      hookOutcome: 'pending',
      publicHookAccount: 'none',
      oldSaveArchived: false,
    },
    settings: {
      interpreterMode: 'offline',
      readingFocus: false,
      textScale: 'normal',
      reducedMotion: false,
      showChoiceAxes: true,
    },
  };
}
