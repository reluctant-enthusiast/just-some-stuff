import type { ChronicleState, VisualState } from './schema.js';

const clamp = (value: number): number => Math.max(0, Math.min(1, value));

export interface PresentationTokens {
  plate: VisualState['plate'];
  palette: VisualState['palette'];
  framing: VisualState['framing'];
  warmth: number;
  visibility: number;
  vignette: number;
  crowding: number;
  movement: number;
  subjectiveAttention: string[];
  cssVariables: Record<string, string>;
}

export function composePresentation(state: ChronicleState): PresentationTokens {
  const visual = state.presentation;
  const shamePressure = clamp(
    state.kalos.tendencies.statusSensitivity * 0.4 +
      (state.flags.publicHookAccount === 'concealed' ? 0.35 : 0) +
      state.kalos.body.fatigue * 0.18,
  );
  const hungerWarmth = clamp(state.kalos.body.hunger * 0.34);
  const crowding = clamp(
    visual.characters.filter((character) => character.proximity === 'crowded').length / 3 +
      (state.sceneId === 'smoke-before-daylight' ? 0.35 : 0),
  );
  const movement = clamp(visual.motion.length / 5 + (visual.proseMode === 'playful-fast' ? 0.35 : 0));
  const objectiveWarmth = visual.palette === 'ember-smoke' || visual.palette === 'low-fire' ? 0.64 : 0.25;
  const warmth = clamp(objectiveWarmth + hungerWarmth - shamePressure * 0.14);
  const visibility = visual.weather === 'mist' ? 0.56 : visual.weather === 'light-rain' ? 0.7 : 0.88;

  return {
    plate: visual.plate,
    palette: visual.palette,
    framing: visual.framing,
    warmth,
    visibility,
    vignette: clamp(0.12 + shamePressure * 0.52),
    crowding,
    movement,
    subjectiveAttention: [...state.kalos.attention, ...visual.focalObjects].slice(0, 6),
    cssVariables: {
      '--scene-warmth': warmth.toFixed(3),
      '--scene-visibility': visibility.toFixed(3),
      '--scene-vignette': clamp(0.12 + shamePressure * 0.52).toFixed(3),
      '--scene-crowding': crowding.toFixed(3),
      '--scene-motion': movement.toFixed(3),
    },
  };
}

export function visualStatePacket(state: ChronicleState): string {
  return JSON.stringify(
    {
      scene: {
        sceneId: state.sceneId,
        beatId: state.beatId,
        location: state.world.location,
        day: state.world.day,
        time: state.world.time,
        season: 'late-summer',
        weather: state.world.weather,
        tide: state.world.tide,
      },
      environment: {
        persistentDamage: state.world.persistentDamage,
        resources: state.world.resources,
        motion: state.presentation.motion,
      },
      continuityObjects: state.presentation.focalObjects,
      characters: state.presentation.characters,
      composition: {
        plate: state.presentation.plate,
        framing: state.presentation.framing,
        palette: state.presentation.palette,
        light: state.presentation.light,
      },
      subjectiveLens: {
        attentionTargets: state.kalos.attention,
        sensoryPriority: state.presentation.sensoryPriority,
        proseMode: state.presentation.proseMode,
      },
      audio: state.presentation.audio,
      derived: composePresentation(state),
    },
    null,
    2,
  );
}
