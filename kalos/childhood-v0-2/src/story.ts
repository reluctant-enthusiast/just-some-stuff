import { BEATS } from './content.js';
import { adjustRelationship, appendObjectiveEvent, recordHistory } from './engine.js';
import type { ChronicleState, MemoryState, PromiseState, PublicStory, VisualState } from './schema.js';

const clamp = (value: number): number => Math.max(0, Math.min(1, value));
const nowId = (prefix: string): string => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;

function memory(text: string, meaning: string, salience = 0.5, sensoryAnchor?: string): MemoryState {
  return {
    id: nowId('memory'), owner: 'kalos', text, meaning, salience, accuracy: 0.9, sensoryAnchor,
  };
}

function story(text: string, source: PublicStory['source'], reach = 0.2, distortion = 0.08): PublicStory {
  return { id: nowId('story'), text, source, reach, distortion };
}

function promise(text: string, status: PromiseState['status'] = 'open'): PromiseState {
  return { id: nowId('promise'), speaker: 'kalos', beneficiary: 'household', text, status, witnesses: ['mara'], repetitions: 0 };
}

function setBeat(state: ChronicleState, beatId: string, ordinal?: number): ChronicleState {
  const next = structuredClone(state);
  const beat = BEATS[beatId];
  if (!beat) throw new Error(`Unknown beat ${beatId}`);
  next.beatId = beatId;
  next.sceneId = beat.sceneId;
  if (ordinal) next.sceneOrdinal = ordinal;
  next.presentation = structuredClone(beat.visual) as VisualState;
  next.world.weather = beat.visual.weather;
  next.world.tide = beat.visual.tide;
  next.world.location = beat.visual.title;
  return next;
}

function tendency(state: ChronicleState, key: keyof ChronicleState['kalos']['tendencies'], delta: number): ChronicleState {
  const next = structuredClone(state);
  next.kalos.tendencies[key] = clamp(next.kalos.tendencies[key] + delta);
  return next;
}

function skill(state: ChronicleState, key: keyof ChronicleState['kalos']['skills'], delta: number): ChronicleState {
  const next = structuredClone(state);
  next.kalos.skills[key] = clamp(next.kalos.skills[key] + delta);
  return next;
}

function body(state: ChronicleState, key: keyof Omit<ChronicleState['kalos']['body'], 'injuries'>, delta: number): ChronicleState {
  const next = structuredClone(state);
  next.kalos.body[key] = clamp(next.kalos.body[key] + delta);
  return next;
}

function addMemory(state: ChronicleState, item: MemoryState): ChronicleState {
  const next = structuredClone(state); next.memories.push(item); return next;
}
function addStory(state: ChronicleState, item: PublicStory): ChronicleState {
  const next = structuredClone(state); next.publicStories.push(item); return next;
}
function addPromise(state: ChronicleState, item: PromiseState): ChronicleState {
  const next = structuredClone(state); next.promises.push(item); return next;
}
function flag(state: ChronicleState, key: string, value: boolean | number | string): ChronicleState {
  const next = structuredClone(state); next.flags[key] = value; return next;
}

export interface Resolution {
  state: ChronicleState;
  transitionBody: string;
  echo?: string;
}

export function resolveChoice(current: ChronicleState, choiceId: string, actionText?: string, interpreter: 'menu' | 'offline' | 'llm' = 'menu'): Resolution {
  const beat = BEATS[current.beatId];
  if (!beat) throw new Error(`No content for ${current.beatId}`);
  let state = recordHistory(current, beat, choiceId, actionText ?? choiceId, interpreter);
  let transitionBody = '';
  let echo: string | undefined;

  switch (`${current.beatId}:${choiceId}`) {
    case 's1-waking:cover-pali':
      state = flag(state, 'paliHelpedAtDawn', true);
      state = skill(state, 'caregiving', 0.04);
      state = adjustRelationship(state, 'pali', 'kalos', { trust: 0.08, affection: 0.05 }, 'Kalos covered him without waking him.');
      state = addMemory(state, memory('Pali never woke when Kalos pulled the blanket back over his foot.', 'Care can be real without being witnessed.', 0.45, 'cold toes against his ribs'));
      transitionBody = '<p>Kalos caught the blanket with two fingers and drew it over Pali’s heel. The child pressed closer without waking. No one saw.</p>';
      echo = 'Pali sleeps on.';
      state = setBeat(state, 's1-hearth');
      break;
    case 's1-waking:follow-seli':
      state = adjustRelationship(state, 'seli', 'kalos', { ease: 0.05, resentment: -0.03 }, 'Kalos followed without asking her to wait.');
      state = body(state, 'cold', -0.08);
      transitionBody = '<p>Seli heard him before he reached the fire. She made room with one knee and did not tell him to go back.</p>';
      echo = 'Seli makes room.';
      state = setBeat(state, 's1-hearth');
      break;
    case 's1-waking:listen-still':
      state = flag(state, 'overheardAdultConcern', true);
      state = tendency(state, 'curiosity', 0.03);
      state.characters.kalos.knowledge.push('Fog damaged part of an outer household’s food racks; the visitor wants labor or food.');
      transitionBody = '<p>He kept his breathing slow. Near the door, Veya said someone had arrived too late to ask without owing. Another voice answered that hunger did not arrive by appointment.</p>';
      echo = 'Kalos hears more than the adults intend.';
      state = setBeat(state, 's1-hearth');
      break;
    case 's1-hearth:take-skin':
      state = flag(state, 'ateCrispSkin', true);
      state = body(state, 'hunger', -0.28);
      state = tendency(state, 'statusSensitivity', 0.02);
      transitionBody = '<p>He took the largest piece. The skin cracked, then softened against his tongue. Seli saw. Eda saw her see. No one asked him to put it back.</p>';
      echo = 'Warm food; a small public claim.';
      state = setBeat(state, 's1-road');
      break;
    case 's1-hearth:split-with-pali':
      state = flag(state, 'sharedBreakfast', true);
      state = body(state, 'hunger', -0.12);
      state = adjustRelationship(state, 'eda', 'kalos', { respect: 0.04, trust: 0.02 });
      state = adjustRelationship(state, 'pali', 'kalos', { affection: 0.06, dependence: 0.04 });
      transitionBody = '<p>Pali woke to the smell and took his half with both hands. Kalos’s own piece disappeared too quickly. Hunger remained, made sharper by watching someone else chew.</p>';
      echo = 'Pali remembers the food; Kalos keeps some hunger.';
      state = setBeat(state, 's1-road');
      break;
    case 's1-hearth:ask-about-fog':
      state = flag(state, 'overheardAdultConcern', true);
      state.characters.kalos.knowledge.push('Some of the visitor’s preserved fish spoiled in fog, and Veya is deciding what aid creates what debt.');
      state = adjustRelationship(state, 'veya', 'kalos', { respect: 0.03, ease: -0.02 });
      transitionBody = '<p>Veya looked at him long enough to decide whether the question was childish. “Food went soft where it should have dried,” she said. “Now people must decide whether help is a gift, a trade, or a chain.”</p>';
      echo = 'A household problem becomes visible.';
      state = setBeat(state, 's1-road');
      break;
    case 's1-hearth:save-for-seli':
      state = flag(state, 'sharedBreakfast', true);
      state = body(state, 'hunger', 0.03);
      state = adjustRelationship(state, 'seli', 'kalos', { affection: 0.08, trust: 0.03 }, 'He left her the largest crisp piece without announcing it.');
      transitionBody = '<p>He left it beside her belt. Seli covered it with one hand before Pali saw, then broke off a corner and pushed it back toward Kalos.</p>';
      echo = 'A private exchange survives the crowded room.';
      state = setBeat(state, 's1-road');
      break;
    case 's1-road:walk-mara':
      state = adjustRelationship(state, 'mara', 'kalos', { respect: 0.04, trust: 0.02 });
      state.characters.kalos.knowledge.push('Mara considers hook binding real work, not practice, because a failed hook can waste a tide.');
      transitionBody = '<p>Mara did not slow. “Work trusted to a child is still work,” she said. “That is what makes it trust.”</p>';
      echo = 'Mara answers the question behind the question.';
      state = setBeat(state, 's2-lesson', 2);
      break;
    case 's1-road:walk-neri':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.07, ease: 0.06 }, 'They carried the prepared fiber together.');
      state = body(state, 'fatigue', 0.02);
      transitionBody = '<p>Kalos took the nearer rim without asking. The bowl stopped tilting. Neri shifted his grip once and said nothing that could be mistaken for thanks.</p>';
      echo = 'The bowl travels level.';
      state = setBeat(state, 's2-lesson', 2);
      break;
    case 's1-road:walk-oren':
      state = adjustRelationship(state, 'oren', 'kalos', { ease: 0.08, affection: 0.04, resentment: 0.01 });
      state.opportunities = state.opportunities.map((opportunity) => opportunity.id === 'older-children-shore' ? { ...opportunity, reason: 'Oren expects Kalos to improve the afternoon game.' } : opportunity);
      transitionBody = '<p>Oren explained the race he had not yet invented. By the time they reached the yard, Kalos had supplied the black post, the wet beam, and the part Oren later called his own idea.</p>';
      echo = 'The afternoon begins forming early.';
      state = setBeat(state, 's2-lesson', 2);
      break;
    case 's1-road:walk-seli':
      state = adjustRelationship(state, 'seli', 'kalos', { trust: 0.08, resentment: -0.06 }, 'Kalos took Pali without making Seli ask.');
      state = skill(state, 'caregiving', 0.03);
      transitionBody = '<p>Seli transferred Pali with indecent speed. “Do not lose him,” she said, already moving ahead. Pali wound one fist into Kalos’s hair for security.</p>';
      echo = 'Seli’s path opens; Kalos’s narrows.';
      state = setBeat(state, 's2-lesson', 2);
      break;
    case 's2-lesson:ask-repeat':
      state = flag(state, 'workMethod', 'asked-mara');
      state = skill(state, 'hookBinding', 0.13);
      state = tendency(state, 'shameTolerance', 0.04);
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.05, respect: 0.04 });
      transitionBody = '<p>Mara showed him again. Oren heard. Nothing terrible happened except that Kalos learned the knot.</p>';
      echo = 'Precision rises; embarrassment passes.';
      state = setBeat(state, 's2-interruption');
      break;
    case 's2-lesson:copy-neri':
      state = flag(state, 'workMethod', 'copied-neri');
      state = skill(state, 'hookBinding', 0.09);
      state = skill(state, 'socialReading', 0.02);
      state = adjustRelationship(state, 'neri', 'kalos', { respect: 0.04, wariness: 0 } as never);
      transitionBody = '<p>Neri flattened each turn with the side of his thumbnail before beginning the next. Kalos copied the motion. Neri noticed on the fourth hook and shifted his hand so the method was easier to see.</p>';
      echo = 'Learning passes between children without becoming a lesson.';
      state = setBeat(state, 's2-interruption');
      break;
    case 's2-lesson:race-work':
      state = flag(state, 'workMethod', 'raced-neri');
      state = skill(state, 'hookBinding', 0.05);
      state = tendency(state, 'statusSensitivity', 0.04);
      state = body(state, 'fatigue', 0.08);
      state = adjustRelationship(state, 'neri', 'kalos', { ease: -0.04, resentment: 0.03 });
      transitionBody = '<p>Kalos finished first. Mara pulled his third binding apart with one finger. Neri finished later and lost none.</p>';
      echo = 'Speed becomes visible; so does the repair it creates.';
      state = setBeat(state, 's2-interruption');
      break;
    case 's2-lesson:invent-method':
      state = flag(state, 'workMethod', 'invented-crosswrap');
      state = skill(state, 'hookBinding', 0.03);
      state = tendency(state, 'curiosity', 0.04);
      state = tendency(state, 'practicalPatience', -0.02);
      transitionBody = '<p>The crossing wrap looked cleaner because it used less fiber. Mara looked at it as if beauty had arrived at the wrong task. “Pull it,” she said. It held once. The second pull loosened it.</p>';
      echo = 'A clever form meets a material test.';
      state = setBeat(state, 's2-interruption');
      break;
    case 's2-interruption:catch-pali':
      state = adjustRelationship(state, 'pali', 'kalos', { trust: 0.08, affection: 0.04 });
      state = adjustRelationship(state, 'mara', 'kalos', { trust: -0.01, respect: 0.03 });
      state.world.resources.bindingFiber = Math.max(0, state.world.resources.bindingFiber - 4);
      transitionBody = '<p>Kalos caught Pali under the arms. The bowl finished turning. Mara’s hooks disappeared under wet fiber, but Pali’s head missed the rack post.</p>';
      echo = 'The child is safe; the work must be recovered.';
      state = setBeat(state, 's2-assignment');
      break;
    case 's2-interruption:save-hooks':
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.04, respect: 0.05 });
      state = adjustRelationship(state, 'pali', 'kalos', { trust: -0.02 });
      transitionBody = '<p>The tray cleared the spill. Pali struck both palms on the mat and looked at Kalos before deciding whether the pain required crying.</p>';
      echo = 'The objects remain ordered; Pali notices the order of rescue.';
      state = setBeat(state, 's2-assignment');
      break;
    case 's2-interruption:make-pali-laugh':
      state = skill(state, 'storytelling', 0.03);
      state = tendency(state, 'humorUnderPressure', 0.04);
      state = adjustRelationship(state, 'pali', 'kalos', { ease: 0.07, affection: 0.04 });
      transitionBody = '<p>Kalos fell after him, more extravagantly. Pali’s cry broke into surprise. Neri saved the bowl. Mara saved the hooks. Kalos saved the moment and none of the objects.</p>';
      echo = 'The yard laughs; other hands repair the work.';
      state = setBeat(state, 's2-assignment');
      break;
    case 's2-interruption:blame-seli':
      state = adjustRelationship(state, 'seli', 'kalos', { trust: -0.08, resentment: 0.09 });
      state = tendency(state, 'strategicOmission', 0.04);
      transitionBody = '<p>“Seli was watching him,” Kalos said. The claim reached Seli before the water stopped moving. She arrived angry enough not to ask whether Pali was hurt.</p>';
      echo = 'Responsibility finds an absent person quickly.';
      state = setBeat(state, 's2-assignment');
      break;
    case 's2-assignment:promise-eight':
      state = addPromise(state, promise('Finish all eight hooks before sunset.'));
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.03 });
      transitionBody = '<p>“All eight,” Kalos said. Mara nodded once. A promise entered the day without changing its length.</p>';
      echo = 'The assignment now has witnesses.';
      state = setBeat(state, 's3-call', 3);
      break;
    case 's2-assignment:ask-neri-pair':
      state = flag(state, 'neriWorkingBesideKalos', true);
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.05, obligation: 0.03 });
      transitionBody = '<p>Neri said yes only after Mara said the hooks remained Kalos’s responsibility. They worked close enough to share water and far enough to preserve ownership.</p>';
      echo = 'Company enters the work; responsibility does not leave it.';
      state = setBeat(state, 's3-call', 3);
      break;
    case 's2-assignment:joke-about-tired':
      state = tendency(state, 'humorUnderPressure', 0.03);
      state = adjustRelationship(state, 'mara', 'kalos', { ease: 0.04, trust: 0.01 });
      transitionBody = '<p>Mara’s mouth altered but did not become a smile. “The hooks may become tired after they work,” she said. “Not before.”</p>';
      echo = 'The warning becomes bearable and remains a warning.';
      state = setBeat(state, 's3-call', 3);
      break;
    case 's2-assignment:say-nothing':
      state = flag(state, 'unnamedHookObligation', true);
      transitionBody = '<p>Kalos gathered the tray. Mara watched long enough to make silence feel less like freedom than he had intended.</p>';
      echo = 'The work is accepted without a clean promise.';
      state = setBeat(state, 's3-call', 3);
      break;
    case 's3-call:finish-before-race':
      state = flag(state, 'hookOutcome', 'careful');
      state = tendency(state, 'followThrough', 0.1);
      state = tendency(state, 'practicalPatience', 0.08);
      state = skill(state, 'hookBinding', 0.05);
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.06 });
      transitionBody = '<p>The first race began without him. Kalos heard his own name once, then heard the game continue. He laid the sixth turn flat and tested the knot until his fingers hurt.</p>';
      echo = 'The hook is sound. The first race is gone.';
      state = setBeat(state, 's3-race');
      break;
    case 's3-call:ask-neri-finish': {
      const trust = state.characters.neri.relationships.kalos?.trust ?? 0.45;
      state = flag(state, 'hookOutcome', trust >= 0.48 ? 'shared-careful' : 'shared-uncertain');
      state = adjustRelationship(state, 'neri', 'kalos', { obligation: 0.08, trust: trust >= 0.48 ? 0.03 : -0.03 });
      transitionBody = trust >= 0.48
        ? '<p>Neri took the hook. “Your work,” he said. “My hands.” He finished the binding and left it in the center of the tray.</p>'
        : '<p>Neri looked at the hook and then at the children calling. “You asked me after you chose,” he said. He gave the hook two careful turns before Mara called him elsewhere.</p>';
      echo = trust >= 0.48 ? 'The hook is shared; ownership remains unsettled.' : 'Help begins and does not finish.';
      state = setBeat(state, 's3-race');
      break;
    }
    case 's3-call:hurry-eighth':
      state = flag(state, 'hookOutcome', 'hurried');
      state = tendency(state, 'followThrough', -0.04);
      state = tendency(state, 'statusSensitivity', 0.04);
      transitionBody = '<p>Two turns. One pull. The binding held against a child’s hands in a warm afternoon. Kalos placed it among the others and ran.</p>';
      echo = 'The hook looks finished from several steps away.';
      state = setBeat(state, 's3-race');
      break;
    case 's3-call:leave-openly':
      state = flag(state, 'hookOutcome', 'openly-unfinished');
      state = tendency(state, 'strategicOmission', -0.02);
      transitionBody = '<p>He left the hook apart from the others, cord loose beside it. Anyone looking at the tray would know there were seven.</p>';
      echo = 'The task remains unfinished but not hidden.';
      state = setBeat(state, 's3-race');
      break;
    case 's3-race:take-wet-beam':
      state = flag(state, 'raceRoute', 'wet-beam');
      state = skill(state, 'balance', 0.05);
      state = body(state, 'fatigue', 0.07);
      state = adjustRelationship(state, 'oren', 'kalos', { respect: 0.07, resentment: 0.03 });
      transitionBody = '<p>The beam rolled once beneath his foot. Kalos used the movement instead of fighting it. He reached the post first and heard Oren begin explaining why the start had been uneven.</p>';
      echo = 'Nerve becomes public property.';
      state = setBeat(state, 's3-laugh');
      break;
    case 's3-race:rewrite-rule':
      state = flag(state, 'raceRoute', 'rewritten');
      state = skill(state, 'socialReading', 0.04);
      state = skill(state, 'storytelling', 0.03);
      state = adjustRelationship(state, 'oren', 'kalos', { ease: 0.04, resentment: 0.04 });
      state = adjustRelationship(state, 'neri', 'kalos', { respect: 0.05 });
      transitionBody = '<p>Kalos added that the maker of a new rule had to cross backward. The group accepted before Oren could refuse without looking afraid of his own invention.</p>';
      echo = 'The rule changes without anyone admitting why.';
      state = setBeat(state, 's3-laugh');
      break;
    case 's3-race:run-old-course':
      state = flag(state, 'raceRoute', 'old-course');
      state = tendency(state, 'shameTolerance', 0.04);
      state = adjustRelationship(state, 'oren', 'kalos', { respect: 0.02, resentment: 0.07 });
      transitionBody = '<p>Kalos ran the course they had named before Oren changed it. Half the children followed him. The other half followed the newest rule. Two winners reached the beach and each called the other afraid.</p>';
      echo = 'Refusal divides the game rather than ending it.';
      state = setBeat(state, 's3-laugh');
      break;
    case 's3-race:stay-with-neri':
      state = flag(state, 'raceRoute', 'stayed-neri');
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.09, affection: 0.05, ease: 0.04 }, 'Kalos let the race leave without them.');
      state = adjustRelationship(state, 'oren', 'kalos', { resentment: 0.05, ease: -0.03 });
      transitionBody = '<p>The others ran. Kalos and Neri stood beneath the racks listening to a game become exciting at a distance. “You could have won,” Neri said. It was not gratitude. It was a fact offered carefully.</p>';
      echo = 'The crowd moves away; one relationship changes nearby.';
      state = setBeat(state, 's3-laugh');
      break;
    case 's3-laugh:laugh-with-neri':
      state = flag(state, 'neriHumiliated', false);
      state = skill(state, 'storytelling', 0.05);
      state = adjustRelationship(state, 'neri', 'kalos', { ease: 0.08, trust: 0.05, respect: 0.04 });
      transitionBody = '<p>Kalos gave Neri a grand reason for entering the mud: he had been inspecting whether it was deep enough to swallow Oren. Neri added that it was, if Oren would stand still. The third laugh belonged to both of them.</p>';
      echo = 'The target becomes an author.';
      state = setBeat(state, 's4-return', 4);
      break;
    case 's3-laugh:turn-on-oren':
      state = flag(state, 'neriHumiliated', false);
      state = tendency(state, 'humorUnderPressure', 0.05);
      state = adjustRelationship(state, 'oren', 'kalos', { resentment: 0.1, trust: -0.04, respect: 0.03 });
      state = adjustRelationship(state, 'neri', 'kalos', { respect: 0.05, trust: 0.03 });
      transitionBody = '<p>Kalos copied Oren’s glance toward the adults before each laugh. The imitation was exact enough that even Oren’s friends recognized it. Oren laughed last and hardest.</p>';
      echo = 'The group sees the status check; Oren sees who exposed it.';
      state = setBeat(state, 's4-return', 4);
      break;
    case 's3-laugh:let-oren-have-it':
      state = flag(state, 'neriHumiliated', true);
      state = adjustRelationship(state, 'oren', 'kalos', { ease: 0.07, affection: 0.03 });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.08, resentment: 0.08 });
      transitionBody = '<p>Kalos laughed. Oren repeated the sound Neri had made and improved it through repetition. Neri climbed from the mud without taking any offered hand.</p>';
      echo = 'Belonging arrives with a cost paid elsewhere.';
      state = setBeat(state, 's4-return', 4);
      break;
    case 's3-laugh:help-neri-up':
      state = flag(state, 'neriHumiliated', false);
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.08, affection: 0.04 });
      state = tendency(state, 'humorUnderPressure', -0.01);
      transitionBody = '<p>Kalos put out his hand. Neri considered refusing it because everyone was watching. Then he took it and pulled hard enough that Kalos nearly joined him in the mud.</p>';
      echo = 'No joke closes the moment.';
      state = setBeat(state, 's4-return', 4);
      break;
    case 's4-return:report-exactly': {
      const outcome = String(state.flags.hookOutcome);
      state = flag(state, 'publicHookAccount', 'exact');
      transitionBody = outcome === 'careful'
        ? '<p>“Eight. I missed the first race for the last one.” Mara looked at the tray, not at the cost Kalos named.</p>'
        : outcome === 'shared-careful' || outcome === 'shared-uncertain'
          ? '<p>“Neri put his hands on the last one. It is still mine.” Neri heard the sentence from the path and did not correct it.</p>'
          : outcome === 'openly-unfinished'
            ? '<p>“Seven. The last is apart from them.” Mara turned then. “Good. Finish the sentence with the hook.”</p>'
            : '<p>“The last one is hurried.” The older children heard. The word did not sound like failure until Mara repeated it.</p>';
      echo = 'The account matches what Kalos knows.';
      state = outcome === 'careful' || outcome === 'shared-careful' ? setBeat(state, 's4-morning-clear') : setBeat(state, 's4-accusation');
      break;
    }
    case 's4-return:claim-finished':
      state = flag(state, 'publicHookAccount', 'claimed-finished');
      state = tendency(state, 'strategicOmission', 0.08);
      state = addPromise(state, promise('All eight hooks are finished.', String(state.flags.hookOutcome).includes('careful') ? 'kept' : 'broken'));
      transitionBody = '<p>“All eight.” The answer was shorter than the walk back to the tray. Mara accepted it because work requires accepting some things before testing them.</p>';
      echo = 'A completed answer enters the household.';
      state = String(state.flags.hookOutcome).includes('careful') ? setBeat(state, 's4-morning-clear') : setBeat(state, 's4-accusation');
      break;
    case 's4-return:check-before-answer': {
      state = flag(state, 'publicHookAccount', 'checked-before-answer');
      state = tendency(state, 'followThrough', 0.05);
      const outcome = String(state.flags.hookOutcome);
      if (outcome === 'hurried' || outcome === 'shared-uncertain' || outcome === 'openly-unfinished') {
        state = flag(state, 'hookOutcome', 'repaired-after-check');
        state = skill(state, 'hookBinding', 0.04);
        transitionBody = '<p>The delay cost him the last of the beach. He found the loose turns, sat down, and finished them while Oren performed impatience nearby.</p>';
        echo = 'Inspection changes the object before it becomes an incident.';
        state = setBeat(state, 's4-morning-clear');
      } else {
        transitionBody = '<p>He returned, pulled the binding once more, and answered from beside the tray. “Eight.” The older children had gone by the time he looked up.</p>';
        echo = 'Certainty acquires a visible cost.';
        state = setBeat(state, 's4-morning-clear');
      }
      break;
    }
    case 's4-return:answer-with-joke':
      state = flag(state, 'publicHookAccount', 'joked-and-delayed');
      state = tendency(state, 'humorUnderPressure', 0.05);
      state = adjustRelationship(state, 'mara', 'kalos', { ease: 0.03, trust: -0.03 });
      transitionBody = '<p>“Are you counting hooks or promises?” Kalos asked. Oren laughed. Mara finally turned. “The thing that can pull a fish from water,” she said. The joke had bought one breath and spent it.</p>';
      echo = 'The room changes temperature; the hook does not.';
      state = String(state.flags.hookOutcome).includes('careful') ? setBeat(state, 's4-morning-clear') : setBeat(state, 's4-accusation');
      break;
    case 's4-morning-clear:keep-quiet-success':
      state = tendency(state, 'followThrough', 0.07);
      state = tendency(state, 'practicalPatience', 0.05);
      state = addMemory(state, memory('The hook held and nobody praised the absence of failure.', 'Prevention can be real without becoming a story.', 0.63, 'cord singing between Mara’s hands'));
      transitionBody = '<p>Kalos reached for the next task before anyone found one for him. The morning continued without requiring his name.</p>';
      echo = 'Quiet competence becomes slightly easier to inhabit.';
      state = setBeat(state, 's5-prevention', 5);
      break;
    case 's4-morning-clear:tell-mara-cost':
      state = adjustRelationship(state, 'mara', 'kalos', { respect: 0.05, ease: -0.01 });
      state = tendency(state, 'statusSensitivity', 0.02);
      transitionBody = '<p>“I missed the race for it,” Kalos said. Mara gave him the next tray. “Then the cost has already been paid,” she answered.</p>';
      echo = 'Recognition arrives as more work.';
      state = setBeat(state, 's5-prevention', 5);
      break;
    case 's4-morning-clear:credit-neri':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.08, respect: 0.04 });
      transitionBody = '<p>“Neri showed me the turn that held.” If that was only partly true, Neri knew which part. Mara handed him the dry fiber first.</p>';
      echo = 'Credit changes who receives the next material.';
      state = setBeat(state, 's5-prevention', 5);
      break;
    case 's4-morning-clear:tease-oren':
      state = adjustRelationship(state, 'oren', 'kalos', { resentment: 0.07, respect: 0.03 });
      state = tendency(state, 'humorUnderPressure', 0.03);
      transitionBody = '<p>Kalos asked whether baskets were meant to hold lines or merely accompany Oren past adults. The yard laughed. Oren reopened his failed corner with his jaw set.</p>';
      echo = 'Competence becomes a weapon quickly.';
      state = setBeat(state, 's5-prevention', 5);
      break;
    case 's4-accusation:confess-hook':
      state = flag(state, 'publicHookAccount', 'confessed');
      state = appendObjectiveEvent(state, { id: nowId('event-hook-confession'), sceneId: 'the-eighth-hook', summary: 'Kalos publicly owned the failed binding before Neri was punished.', facts: ['Kalos identified his own work.', 'Neri was cleared.', 'The fishing party left without Kalos.'], witnesses: ['kalos', 'mara', 'neri', 'oren'], tags: ['confession', 'displaced-cost-prevented'] });
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.08, respect: 0.05 });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.14, resentment: -0.04 });
      state = tendency(state, 'shameTolerance', 0.07);
      transitionBody = '<p>“I bound it badly. Neri did not damage it.” Mara returned Kalos’s own claim from yesterday before sending the canoe away without him.</p>';
      echo = 'Neri is cleared. The desirable morning leaves.';
      state = setBeat(state, 's5-mara', 5);
      break;
    case 's4-accusation:test-hooks':
      state = flag(state, 'publicHookAccount', 'evidence-first');
      state = tendency(state, 'strategicOmission', 0.03);
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.04, resentment: 0.03 });
      transitionBody = '<p>Mara set the failed hook down and brought the tray to the mat. Neri was not released. He was made to wait while evidence approached the answer Kalos already had.</p>';
      echo = 'Blame pauses; suspicion acquires structure.';
      state = setBeat(state, 's5-evidence', 5);
      break;
    case 's4-accusation:joke-oren-basket':
      state = flag(state, 'publicHookAccount', 'concealed-by-humor');
      state = tendency(state, 'humorUnderPressure', 0.09);
      state = adjustRelationship(state, 'oren', 'kalos', { resentment: 0.12, respect: 0.03 });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.08, resentment: 0.08 });
      state = addStory(state, story('Oren accused Neri while his own basket sat unfinished.', 'household', 0.32, 0.14));
      transitionBody = '<p>“Perhaps the hook climbed out to finish Oren’s basket,” Kalos said. Adults laughed. Oren’s accusation weakened. Neri still remained behind with the tray.</p>';
      echo = 'The room changes target; the burden does not fully move.';
      state = setBeat(state, 's5-neri', 5);
      break;
    case 's4-accusation:silent-hook':
      state = flag(state, 'publicHookAccount', 'concealed');
      state = tendency(state, 'strategicOmission', 0.09);
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.18, resentment: 0.15 });
      state = addStory(state, story('Neri mishandled the hook tray and missed the morning tide.', 'oren', 0.45, 0.22));
      transitionBody = '<p>Kalos lowered his eyes. Mara sent Neri to the shed. The canoe still had room for Kalos, and the children at the shore still wanted him.</p>';
      echo = 'Immediate belonging; displaced cost.';
      state = setBeat(state, 's5-neri', 5);
      break;
    case 's5-mara:plain-want':
      state = tendency(state, 'shameTolerance', 0.05);
      state = tendency(state, 'strategicOmission', -0.02);
      transitionBody = '<p>“Then next time say which thing you are choosing,” Mara said. She did not make the feeling smaller. She pushed the first hook across the mat.</p>';
      echo = 'The motive is named; the work remains.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-mara:thought-held':
      state = tendency(state, 'strategicOmission', 0.02);
      state = skill(state, 'hookBinding', 0.06);
      transitionBody = '<p>Mara showed him where the fiber had never seated. “You thought after you stopped looking,” she said.</p>';
      echo = 'Technical truth exposes the gap in the explanation.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-mara:feared-left':
      state = addMemory(state, memory('The canoe left while he sat among hooks.', 'Concealment produced the exclusion it tried to prevent.', 0.72, 'a hull scraping away from stones'));
      state = tendency(state, 'shameTolerance', 0.04);
      transitionBody = '<p>“They would have left me,” Kalos said. Mara listened to the canoe disappear. “They did,” she answered.</p>';
      echo = 'The feared outcome arrives by another route.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-mara:fish-joke':
      state = tendency(state, 'humorUnderPressure', 0.07);
      state = adjustRelationship(state, 'mara', 'kalos', { ease: 0.05, trust: 0.01 });
      state = addMemory(state, memory('Mara almost smiled and did not release him from the work.', 'Laughter can make accountability survivable without ending it.', 0.62));
      transitionBody = '<p>Mara almost smiled. Kalos felt relief too quickly. “The fish did not promise me eight hooks,” she said.</p>';
      echo = 'Warmth rises; obligation remains.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-evidence:complete-truth':
      state = adjustRelationship(state, 'mara', 'kalos', { trust: 0.02, respect: 0.03 });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.07, resentment: 0.02 });
      state = tendency(state, 'shameTolerance', 0.04);
      transitionBody = '<p>Kalos told it from the call at the beach through the answer he gave Mara. Neri was cleared. He did not look grateful.</p>';
      echo = 'Late truth retains value and carries the delay inside it.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-evidence:admit-hurried':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.06, resentment: 0.04 });
      state = tendency(state, 'strategicOmission', 0.04);
      transitionBody = '<p>“I hurried them. Neri did not drop them.” Mara knew he had not answered what he told her yesterday. She did not force the second truth in public.</p>';
      echo = 'Neri is protected; part of the account remains sealed.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-evidence:implicate-lesson':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.16, resentment: 0.14 });
      state = adjustRelationship(state, 'mara', 'kalos', { trust: -0.09 });
      state = tendency(state, 'strategicOmission', 0.08);
      transitionBody = '<p>Neri explained, calmly, that the failed wrap was not the method he had shown. The true fact Kalos used became the sharpest part of the lie.</p>';
      echo = 'Partial truth becomes a tool of blame.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-evidence:let-mara-infer':
      state = adjustRelationship(state, 'mara', 'kalos', { trust: -0.11, respect: -0.02 });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.05, resentment: 0.05 });
      transitionBody = '<p>Mara completed the sequence without his help. Kalos learned that truth could arrive without being invited and give him no part in protecting anyone.</p>';
      echo = 'Evidence resolves the event; relationship cost remains.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-neri:late-public-confession':
      state = flag(state, 'publicHookAccount', 'late-confession');
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.12, resentment: -0.04 });
      state = tendency(state, 'shameTolerance', 0.05);
      state = addMemory(state, memory('He could still return after choosing badly.', 'Reversal does not erase the first choice.', 0.74));
      transitionBody = '<p>Neri let Kalos speak when they returned. Mara asked why he had needed Neri to know before she did.</p>';
      echo = 'The first failure remains; so does the return.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-neri:buy-silence':
      state = addPromise(state, { ...promise('Finish Neri’s repair in exchange for silence.'), beneficiary: 'neri' });
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.05, dependence: 0.08, resentment: 0.06 });
      state = tendency(state, 'strategicOmission', 0.08);
      transitionBody = '<p>Neri accepted because he wanted the morning back. The bargain reduced his work and increased what he knew about Kalos.</p>';
      echo = 'Labor purchases silence, not trust.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-neri:silent-repair':
      state = adjustRelationship(state, 'neri', 'kalos', { respect: 0.04, trust: 0.01, resentment: 0.07 });
      state = addMemory(state, memory('Doing the work felt almost like telling the truth.', 'Restitution and accountability can resemble each other from inside the body.', 0.67, 'wet fiber between both boys'));
      transitionBody = '<p>Kalos sat and began binding. After a time Neri asked, “Am I supposed to thank you?” Neither boy stopped working.</p>';
      echo = 'The burden is reduced; the lie remains public.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-neri:deny-to-neri':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.24, resentment: 0.16, affection: -0.08 });
      state = tendency(state, 'strategicOmission', 0.09);
      state = addMemory(state, memory('Neri looked at the dark groove the fiber had left across Kalos’s fingers.', 'Some evidence belongs only to the person being lied to.', 0.71, 'fiber mark across his fingers'));
      transitionBody = '<p>“Then go,” Neri said. Kalos reached the beach while the laughter was still good.</p>';
      echo = 'The concealment succeeds immediately.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-neri:leave-for-beach':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: -0.2, resentment: 0.18, affection: -0.07 });
      state = adjustRelationship(state, 'oren', 'kalos', { ease: 0.08, affection: 0.04 });
      state = addStory(state, story('Neri dropped the tray and spoiled several hooks.', 'oren', 0.58, 0.3));
      transitionBody = '<p>Kalos joined the older children and improved Oren’s version of Mara until even Seli laughed. Behind the shed, Neri finished the tray.</p>';
      echo = 'A good morning for Kalos; a story begins hardening elsewhere.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-prevention:continue-ordinary':
      state = tendency(state, 'followThrough', 0.06);
      state = tendency(state, 'practicalPatience', 0.06);
      transitionBody = '<p>He began the next hook. No one marked the moment, which was part of what made it difficult.</p>';
      echo = 'Prevention becomes practice rather than performance.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-prevention:seek-mara-credit':
      state = adjustRelationship(state, 'mara', 'kalos', { respect: 0.04, ease: 0.01 });
      state = tendency(state, 'statusSensitivity', 0.02);
      transitionBody = '<p>Mara said she had seen it because it held. Then she handed him a harder hook with a narrower shank.</p>';
      echo = 'Credit arrives as access to more demanding work.';
      state.opportunities = state.opportunities.map((opportunity) => opportunity.id === 'mara-apprenticeship' ? { ...opportunity, available: true, reason: 'Mara saw Kalos choose inspection over the race.' } : opportunity);
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-prevention:help-oren-basket':
      state = adjustRelationship(state, 'oren', 'kalos', { trust: 0.06, resentment: -0.03, respect: 0.05 });
      transitionBody = '<p>Kalos showed him the loosened corner without making his voice carry. Oren repaired it and later told the others he had noticed first.</p>';
      echo = 'Help changes the object; ownership of the help remains contested.';
      state = setBeat(state, 's6-close', 6);
      break;
    case 's5-prevention:share-with-neri':
      state = adjustRelationship(state, 'neri', 'kalos', { trust: 0.07, affection: 0.04, ease: 0.04 });
      transitionBody = '<p>Neri took the driest fiber and pushed the next-best bundle back. The exchange required no account.</p>';
      echo = 'Material generosity leaves little room for rhetoric.';
      state = setBeat(state, 's6-close', 6);
      break;
    default:
      throw new Error(`No transition for ${current.beatId}:${choiceId}`);
  }

  if (state.beatId === 's6-close') {
    state.completed = true;
    state.world.day = 24;
    state.world.time = 'night';
    state.world.weather = 'light-rain';
  }
  return { state, transitionBody, echo };
}
