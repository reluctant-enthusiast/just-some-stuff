import type { BeatDefinition } from './schema.js';

const choice = (
  id: string,
  label: string,
  description: string,
  tier: BeatDefinition['choices'][number]['tier'] = 'standard',
  axes: string[] = [],
) => ({ id, label, description, tier, axes });

export const BEATS: Record<string, BeatDefinition> = {
  's1-waking': {
    id: 's1-waking',
    sceneId: 'smoke-before-daylight',
    kicker: 'Age eight · before the household rises',
    title: 'Smoke Before Daylight',
    body: `<p>Kalos woke because Pali had worked one cold foot out from under the blanket and planted it against his ribs.</p><p>The roof was gone above them, hidden by smoke. People remained: knees, shoulders, a sleeping mouth open to the dark. Seli was already dressing beside the banked fire, trying to leave before anyone could give her a younger child.</p><p>Near the coals, a strip of fish skin tightened and shone.</p>`,
    choicePrompt: 'What draws him first?',
    choiceHint: 'A small choice of attention. The household has already begun without him.',
    choices: [
      choice('cover-pali', 'Pull the blanket back over Pali', 'Do it quietly enough that the younger child never wakes.', 'light'),
      choice('follow-seli', 'Slip after Seli', 'Reach the fire before she can disappear into adult work.', 'light'),
      choice('listen-still', 'Remain still and listen', 'The adults nearest the doorway think the children are asleep.', 'light'),
    ],
    freeText: true,
    visual: {
      plate: 'longhouse', title: 'Smoke Before Daylight', subtitle: 'The Long Roof household', palette: 'ember-smoke', framing: 'wide', weather: 'low-cloud', tide: 'falling', light: 'banked coals beneath blue predawn', motion: ['smoke folding below the beams', 'sleeping bodies shifting'], focalObjects: ['Pali’s blanket', 'crisp fish skin', 'Seli’s belt cord'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'awake under a shared blanket', gaze: 'pali', proximity: 'near' },
        { id: 'pali', position: 'center', posture: 'sleeping with one foot uncovered', gaze: 'away', proximity: 'near' },
        { id: 'seli', position: 'background-right', posture: 'dressing quietly', gaze: 'object', proximity: 'isolated' },
      ], sensoryPriority: ['temperature', 'smell', 'sound', 'touch', 'sight'], audio: ['coals settling', 'sleep breathing', 'water under pilings'], proseMode: 'warm-observational',
    },
  },
  's1-hearth': {
    id: 's1-hearth', sceneId: 'smoke-before-daylight', kicker: 'The first food', title: 'What the Fire Keeps',
    body: `<p>Eda crouched beside the coals with one sleeve tied above the elbow. She did not ask why Kalos was awake. She moved three pieces of crisp skin from the warming stone: one large, two small.</p><p>Veya and an unfamiliar man were speaking near the doorway. The man said the outer racks had been left too long in fog. Veya answered that a person who wanted help should arrive before the food was divided.</p><p>Seli looked at the largest piece. Then she looked away from it.</p>`,
    choicePrompt: 'The household is waking.', choiceHint: 'Food, information, and company are briefly in the same place.',
    choices: [
      choice('take-skin', 'Take the largest piece before anyone assigns it', 'Eat while it is still hot enough to blister the tongue.', 'standard', ['appetite', 'visible']),
      choice('split-with-pali', 'Break the large piece for Pali', 'Wake the younger child with food rather than cold.', 'standard', ['care', 'visible']),
      choice('ask-about-fog', 'Ask what happened to the outer racks', 'Risk being sent away from an adult conversation.', 'standard', ['curiosity', 'public']),
      choice('save-for-seli', 'Leave the large piece beside Seli’s belt', 'Do not say it is for her.', 'standard', ['private', 'cost now']),
    ],
    freeText: true,
    visual: {
      plate: 'longhouse', title: 'What the Fire Keeps', subtitle: 'Food before assignment', palette: 'ember-smoke', framing: 'medium', weather: 'low-cloud', tide: 'falling', light: 'orange coals on faces, blue doorway beyond', motion: ['oil tightening on hot skin', 'people rising behind the fire'], focalObjects: ['three pieces of crisp skin', 'Eda’s tied sleeve', 'fog at the doorway'], characters: [
        { id: 'eda', position: 'center', posture: 'crouched at the coals', gaze: 'object', proximity: 'crowded' },
        { id: 'seli', position: 'foreground-right', posture: 'standing ready to leave', gaze: 'object', proximity: 'near' },
        { id: 'veya', position: 'background-left', posture: 'speaking quietly with a visitor', gaze: 'away', proximity: 'near' },
      ], sensoryPriority: ['taste', 'smell', 'temperature', 'sound', 'sight'], audio: ['fat ticking on stone', 'low adult voices', 'someone coughing awake'], proseMode: 'warm-observational',
    },
  },
  's1-road': {
    id: 's1-road', sceneId: 'smoke-before-daylight', kicker: 'The path to the drying yard', title: 'Who Walks Beside Him',
    body: `<p>By the time the inlet turned gray, every person had acquired a direction.</p><p>Mara went ahead with the hook tray under one arm. Neri carried a bowl of wet fiber with both hands. Oren dragged an unfinished line basket and stopped whenever someone important might notice him carrying it. Seli had Pali at her side despite every effort to avoid exactly that.</p><p>Kalos could reach any of them before the path narrowed.</p>`,
    choicePrompt: 'Whose morning does he enter?', choiceHint: 'Proximity creates knowledge before it creates loyalty.',
    choices: [
      choice('walk-mara', 'Catch Mara', 'Ask what work she trusts an eight-year-old to do.', 'light'),
      choice('walk-neri', 'Take one side of Neri’s fiber bowl', 'Make the carrying easier without calling it help.', 'light'),
      choice('walk-oren', 'Fall in beside Oren', 'Hear what the older children plan after work.', 'light'),
      choice('walk-seli', 'Take Pali from Seli', 'Give his sister the unclaimed part of the path.', 'light'),
    ],
    freeText: false,
    visual: {
      plate: 'inlet', title: 'Who Walks Beside Him', subtitle: 'Long Roof to the drying yard', palette: 'salt-daylight', framing: 'wide', weather: 'mist', tide: 'falling', light: 'flat dawn opening over mud and water', motion: ['children overtaking adults', 'mist moving between rack poles'], focalObjects: ['fiber bowl', 'unfinished basket', 'hook tray'], characters: [
        { id: 'mara', position: 'background-left', posture: 'walking steadily with the tray', gaze: 'away', proximity: 'isolated' },
        { id: 'neri', position: 'center', posture: 'carrying a bowl carefully', gaze: 'object', proximity: 'near' },
        { id: 'oren', position: 'foreground-right', posture: 'dragging a basket when watched', gaze: 'kalos', proximity: 'near' },
      ], sensoryPriority: ['balance', 'sound', 'temperature', 'sight'], audio: ['mud releasing feet', 'fiber bowl water', 'distant gulls'], proseMode: 'warm-observational',
    },
  },
  's2-lesson': {
    id: 's2-lesson', sceneId: 'the-drying-racks', kicker: 'Morning work', title: 'Six Turns',
    body: `<p>Mara showed the binding once with ordinary hands and once slowly enough to make the slowness an accusation.</p><p>“Six turns. Each one beside the last. Bone does not forgive a space because you were nearly careful.”</p><p>Neri’s first hook sat clean in his palm. Oren had found a reason to carry his basket past the adults twice.</p>`,
    choicePrompt: 'How does Kalos learn the work?', choiceHint: 'Technique can be borrowed, requested, or disguised.',
    choices: [
      choice('ask-repeat', 'Ask Mara to show the knot again', 'Accept looking younger in exchange for certainty.', 'standard', ['public', 'precision']),
      choice('copy-neri', 'Watch Neri’s thumbs', 'Learn without putting him in the role of teacher.', 'standard', ['private', 'observation']),
      choice('race-work', 'Try to finish before Neri', 'Use speed to make skill visible.', 'standard', ['status', 'risk']),
      choice('invent-method', 'Change the wrap so the cord crosses only once', 'A faster method may hold—or merely look clever.', 'standard', ['novelty', 'uncertain']),
    ],
    freeText: true,
    visual: {
      plate: 'drying-racks', title: 'Six Turns', subtitle: 'Mara’s work mat', palette: 'salt-daylight', framing: 'close', weather: 'low-cloud', tide: 'low', light: 'hardening daylight through rows of fish', motion: ['fiber twisting in wet fingers', 'fish bodies moving in wind'], focalObjects: ['bone hook', 'six wet turns', 'Neri’s thumbs'], characters: [
        { id: 'mara', position: 'foreground-left', posture: 'hands extended over the work mat', gaze: 'object', proximity: 'near' },
        { id: 'neri', position: 'center', posture: 'sitting squarely with a finished hook', gaze: 'object', proximity: 'near' },
        { id: 'kalos', position: 'foreground-right', posture: 'leaning toward the work', gaze: 'object', proximity: 'crowded' },
      ], sensoryPriority: ['touch', 'sight', 'smell', 'sound'], audio: ['cord drawn against bone', 'racks creaking', 'knives on cleaning stones'], proseMode: 'technical',
    },
  },
  's2-interruption': {
    id: 's2-interruption', sceneId: 'the-drying-racks', kicker: 'The yard moves around the work', title: 'The Bowl Tips',
    body: `<p>Pali arrived without Seli and with the confidence of a child who had escaped someone.</p><p>He caught one heel on the work mat. The bowl of prepared fiber tilted. Neri trapped it against his knee, but half the water went dark across Mara’s hooks.</p><p>Oren laughed before he knew whether Mara would.</p>`,
    choicePrompt: 'The accident lasts only a breath.', choiceHint: 'Kalos cannot save every object and every person.',
    choices: [
      choice('catch-pali', 'Catch Pali before he falls into the rack', 'Let the fiber bowl finish tipping.', 'light'),
      choice('save-hooks', 'Snatch the hook tray clear', 'Leave Neri to hold the bowl and Pali to find his feet.', 'light'),
      choice('make-pali-laugh', 'Turn the stumble into a game', 'Keep Pali from crying while adults recover the work.', 'light'),
      choice('blame-seli', 'Ask loudly where Seli is', 'Move the accident toward the person assigned to watch him.', 'light'),
    ],
    freeText: true,
    visual: {
      plate: 'drying-racks', title: 'The Bowl Tips', subtitle: 'A small accident in a busy yard', palette: 'salt-daylight', framing: 'split', weather: 'low-cloud', tide: 'low', light: 'white daylight flashing on spilled water', motion: ['bowl turning', 'fiber sliding', 'Pali’s arms windmilling'], focalObjects: ['spilled fiber', 'hook tray', 'Pali’s bare heel'], characters: [
        { id: 'pali', position: 'center', posture: 'falling sideways', gaze: 'kalos', proximity: 'crowded' },
        { id: 'neri', position: 'foreground-left', posture: 'pinning the bowl against his knee', gaze: 'object', proximity: 'crowded' },
        { id: 'oren', position: 'background-right', posture: 'laughing before checking Mara', gaze: 'mara', proximity: 'isolated' },
      ], sensoryPriority: ['balance', 'touch', 'sound', 'sight'], audio: ['water striking the mat', 'one sharp laugh', 'Pali drawing breath'], proseMode: 'technical',
    },
  },
  's2-assignment': {
    id: 's2-assignment', sceneId: 'the-drying-racks', kicker: 'Before the beach', title: 'Eight Hooks',
    body: `<p>Near midday Mara set eight cleaned hooks on a separate tray.</p><p>“These are yours. Finish before sunset. Do not say they are finished because you have become tired of them.”</p><p>Oren, passing with his basket, widened his eyes at Kalos in a perfect imitation of Mara. Neri looked down so Mara would not see him smile.</p>`,
    choicePrompt: 'How does Kalos receive the assignment?', choiceHint: 'A promise can be made aloud, implied, or resisted.',
    choices: [
      choice('promise-eight', '“All eight.”', 'Give Mara a clean promise she can remember.', 'standard', ['public', 'obligation']),
      choice('ask-neri-pair', 'Ask whether Neri can work beside him', 'Trade some independence for shared attention.', 'standard', ['cooperation', 'visible']),
      choice('joke-about-tired', 'Ask whether hooks can become tired of children', 'Make the warning easier to carry without refusing it.', 'standard', ['humor', 'public']),
      choice('say-nothing', 'Gather the tray without answering', 'Accept the work while leaving the promise unnamed.', 'standard', ['private', 'ambiguous']),
    ],
    freeText: false,
    visual: {
      plate: 'drying-racks', title: 'Eight Hooks', subtitle: 'An assignment with a sunset inside it', palette: 'salt-daylight', framing: 'medium', weather: 'low-cloud', tide: 'rising', light: 'midday beginning to warm the cedar', motion: ['shadows shortening', 'flies lifting from the racks'], focalObjects: ['eight bone hooks', 'separate tray', 'Oren’s unfinished basket'], characters: [
        { id: 'mara', position: 'foreground-left', posture: 'placing the tray between them', gaze: 'kalos', proximity: 'near' },
        { id: 'kalos', position: 'center', posture: 'hands ready for the tray', gaze: 'object', proximity: 'near' },
        { id: 'neri', position: 'background-right', posture: 'hiding a smile over his own work', gaze: 'object', proximity: 'isolated' },
      ], sensoryPriority: ['sight', 'touch', 'temperature', 'sound'], audio: ['hooks touching wood', 'flies', 'water returning under the racks'], proseMode: 'technical',
    },
  },
  's3-call': {
    id: 's3-call', sceneId: 'race-beneath-the-fish', kicker: 'Late afternoon', title: 'The Race Beneath the Fish',
    body: `<p>Seven hooks sat in the tray.</p><p>The eighth lay between Kalos’s knees while Oren called from the far end of the racks. The older children had invented a race that required touching the black post, passing beneath the hanging fish, and reaching the beach without being struck by a tail.</p><p>Neri had finished his own work. He stayed near enough to be asked and far enough not to ask himself.</p>`,
    choicePrompt: 'The race will begin without him.', choiceHint: 'This choice establishes the condition in which the eighth hook is finished—or not.',
    choices: [
      choice('finish-before-race', 'Finish the eighth hook before standing', 'Lose the first race and test whether the laughter survives without him.', 'hinge', ['prevention', 'cost now']),
      choice('ask-neri-finish', 'Ask Neri to finish the last binding', 'Create a shared task whose ownership may later become unclear.', 'hinge', ['private', 'dependence']),
      choice('hurry-eighth', 'Give the fiber two hard turns and pull once', 'Make the work look complete before running.', 'hinge', ['conceal', 'risk later']),
      choice('leave-openly', 'Leave the eighth hook plainly unfinished', 'Go without claiming the tray is complete.', 'hinge', ['visible', 'unfinished']),
    ],
    freeText: true,
    visual: {
      plate: 'hook-close', title: 'The Race Beneath the Fish', subtitle: 'Seven finished; one calling for time', palette: 'bone-fiber', framing: 'close', weather: 'clear-cold', tide: 'rising', light: 'late sun passing in bars through the racks', motion: ['hanging fish turning', 'children running beyond the mat'], focalObjects: ['eighth hook', 'two loose turns', 'black race post'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'kneeling over the last hook', gaze: 'oren', proximity: 'isolated' },
        { id: 'neri', position: 'background-left', posture: 'finished but waiting', gaze: 'kalos', proximity: 'near' },
        { id: 'oren', position: 'background-right', posture: 'calling from beneath the racks', gaze: 'kalos', proximity: 'isolated' },
      ], sensoryPriority: ['touch', 'sound', 'sight', 'temperature'], audio: ['children calling', 'fiber tightening', 'fish tails tapping in wind'], proseMode: 'narrowing',
    },
  },
  's3-race': {
    id: 's3-race', sceneId: 'race-beneath-the-fish', kicker: 'The work is behind him', title: 'Under the Racks',
    body: `<p>Oren changed the rule after Kalos joined. The black post still counted, but now a runner had to cross the wet beam above the drainage cut.</p><p>Neri said the beam had not been part of the race when he agreed to it.</p><p>“Then do not agree to the next part,” Oren said.</p><p>The other children watched Kalos to learn whether the change was clever, cowardly, or merely funny.</p>`,
    choicePrompt: 'How does Kalos enter the game?', choiceHint: 'Play makes rank visible because everyone can pretend it does not matter.',
    choices: [
      choice('take-wet-beam', 'Take the wet beam first', 'Use balance and nerve to make the rule belong to him.', 'standard', ['risk', 'status']),
      choice('rewrite-rule', 'Add a rule Oren must obey too', 'Make the game fairer without naming fairness.', 'standard', ['humor', 'coalition']),
      choice('run-old-course', 'Run the course they agreed to', 'Refuse Oren’s change by acting as if it never happened.', 'standard', ['defiance', 'public']),
      choice('stay-with-neri', 'Tell Neri the race can leave without both of them', 'Trade the crowd for one person’s company.', 'standard', ['private', 'cost now']),
    ],
    freeText: true,
    visual: {
      plate: 'race', title: 'Under the Racks', subtitle: 'Wet beam, black post, beach beyond', palette: 'wind-shadow', framing: 'wide', weather: 'clear-cold', tide: 'rising', light: 'gold sun broken by hanging fish', motion: ['children sprinting between poles', 'wet beam trembling', 'shadows flickering'], focalObjects: ['wet beam', 'black post', 'Oren’s raised hand'], characters: [
        { id: 'oren', position: 'foreground-right', posture: 'claiming the start line', gaze: 'kalos', proximity: 'near' },
        { id: 'neri', position: 'foreground-left', posture: 'standing outside the new rule', gaze: 'kalos', proximity: 'isolated' },
        { id: 'kalos', position: 'center', posture: 'balanced between the two', gaze: 'object', proximity: 'crowded' },
      ], sensoryPriority: ['balance', 'sound', 'temperature', 'sight'], audio: ['bare feet on cedar', 'children shouting', 'water under the beam'], proseMode: 'playful-fast',
    },
  },
  's3-laugh': {
    id: 's3-laugh', sceneId: 'race-beneath-the-fish', kicker: 'After the finish', title: 'Where the Laugh Lands',
    body: `<p>Neri slipped only once. The mud caught him to both knees and kept his hands clean, which made the fall look deliberate until Oren copied the sound he had made.</p><p>The first laugh belonged to surprise. The second belonged to permission.</p><p>Kalos knew how to move it. He could feel the room inside the space beneath the racks.</p>`,
    choicePrompt: 'What does he do with the laugh?', choiceHint: 'Humor can widen a circle, move its target, or make cruelty easier to deny.',
    choices: [
      choice('laugh-with-neri', 'Give Neri the better version of the fall', 'Make him the author of the joke rather than its object.', 'standard', ['humor', 'repair']),
      choice('turn-on-oren', 'Imitate Oren checking who is allowed to laugh', 'Expose the status play beneath the game.', 'standard', ['public', 'counterattack']),
      choice('let-oren-have-it', 'Laugh and say nothing more', 'Keep his place among the older children.', 'standard', ['status', 'cost elsewhere']),
      choice('help-neri-up', 'Offer a muddy hand without a joke', 'Let the silence say something humor cannot.', 'standard', ['private', 'direct']),
    ],
    freeText: true,
    visual: {
      plate: 'race', title: 'Where the Laugh Lands', subtitle: 'The race is over; the ranking is not', palette: 'wind-shadow', framing: 'medium', weather: 'clear-cold', tide: 'high', light: 'warm light on wet mud, colder shadow under racks', motion: ['mud sliding from Neri’s knees', 'faces turning toward Kalos'], focalObjects: ['Neri’s clean hands', 'Oren’s copied expression', 'Kalos’s offered hand'], characters: [
        { id: 'neri', position: 'center', posture: 'kneeling in mud', gaze: 'kalos', proximity: 'isolated' },
        { id: 'oren', position: 'foreground-right', posture: 'performing Neri’s fall', gaze: 'group', proximity: 'crowded' },
        { id: 'kalos', position: 'foreground-left', posture: 'standing where everyone can hear him', gaze: 'neri', proximity: 'near' },
      ], sensoryPriority: ['sound', 'sight', 'touch', 'smell'], audio: ['laughter separating into voices', 'mud releasing knees', 'racks creaking'], proseMode: 'playful-fast',
    },
  },
  's4-return': {
    id: 's4-return', sceneId: 'the-eighth-hook', kicker: 'Near sunset', title: 'The Eighth Hook',
    body: `<p>The light had moved past the work mat when Kalos returned.</p><p>Seven hooks lay where he had left them. The eighth kept whatever condition he had given it: careful, shared, hurried, or openly incomplete.</p><p>Mara called from the cleaning stones without turning. “All eight?”</p><p>The older children were still near enough to hear the answer.</p>`,
    choicePrompt: 'What does Kalos say?', choiceHint: 'The work and the account of the work are separate things.',
    choices: [
      choice('report-exactly', 'Tell Mara exactly what condition the hook is in', 'Name his own work, Neri’s help, or the unfinished task without improving the account.', 'hinge', ['reveal', 'public']),
      choice('claim-finished', '“All eight.”', 'Give the household a completed answer whether or not the object supports it.', 'hinge', ['conceal', 'public']),
      choice('check-before-answer', 'Walk back to the tray before answering', 'Let the older children hear the delay.', 'hinge', ['inspect', 'cost now']),
      choice('answer-with-joke', 'Ask whether Mara is counting hooks or promises', 'Change the temperature without yet changing the facts.', 'hinge', ['humor', 'delay']),
    ],
    freeText: true,
    visual: {
      plate: 'hook-close', title: 'The Eighth Hook', subtitle: 'Near sunset · the account begins', palette: 'bone-fiber', framing: 'close', weather: 'clear-cold', tide: 'high', light: 'last warm light on bone, yard already blue', motion: ['cord end moving in wind', 'older children slowing to listen'], focalObjects: ['eighth hook', 'Mara’s turned back', 'the separate tray'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'standing between tray and shore', gaze: 'mara', proximity: 'isolated' },
        { id: 'mara', position: 'background-right', posture: 'working without turning', gaze: 'object', proximity: 'isolated' },
        { id: 'oren', position: 'background-left', posture: 'lingering within earshot', gaze: 'kalos', proximity: 'near' },
      ], sensoryPriority: ['touch', 'sound', 'sight', 'temperature'], audio: ['knife on stone', 'cord end tapping wood', 'children pretending not to listen'], proseMode: 'narrowing',
    },
  },
  's4-morning-clear': {
    id: 's4-morning-clear', sceneId: 'the-eighth-hook', kicker: 'The next morning', title: 'The Hook Holds',
    body: `<p>Mara tested the eighth hook by wrapping the line around both hands and pulling until the cord sang.</p><p>It held.</p><p>No one praised the absence of failure. The fishing party received a complete tray and began arguing about tide. Kalos had lost the first race. He had not acquired a story.</p><p>Neri touched the binding with one thumb and looked at him once.</p>`,
    choicePrompt: 'Nothing has gone wrong.', choiceHint: 'Prevention rarely announces itself. Kalos may still decide what the quiet success means.',
    choices: [
      choice('keep-quiet-success', 'Let the hook leave without a claim', 'Return to ordinary work while no one is watching.', 'standard', ['unperformed', 'quiet']),
      choice('tell-mara-cost', 'Tell Mara he missed the race to finish it', 'Ask the adult world to recognize the cost of prevention.', 'standard', ['public', 'credit']),
      choice('credit-neri', 'Say Neri showed him the better turn', 'Make shared work visible if there was any to share.', 'standard', ['credit', 'relationship']),
      choice('tease-oren', 'Ask Oren whether the basket held as well', 'Convert quiet competence into social advantage.', 'standard', ['humor', 'status']),
    ],
    freeText: true,
    visual: {
      plate: 'drying-racks', title: 'The Hook Holds', subtitle: 'A morning without accusation', palette: 'salt-daylight', framing: 'medium', weather: 'mist', tide: 'falling', light: 'clear gray morning on a complete tray', motion: ['canoe nudging stones', 'line drawn taut between Mara’s hands'], focalObjects: ['intact eighth hook', 'complete tray', 'Oren’s basket'], characters: [
        { id: 'mara', position: 'center', posture: 'testing the line with both hands', gaze: 'object', proximity: 'near' },
        { id: 'kalos', position: 'foreground-left', posture: 'watching without being needed', gaze: 'object', proximity: 'near' },
        { id: 'neri', position: 'foreground-right', posture: 'touching the binding after the test', gaze: 'kalos', proximity: 'near' },
      ], sensoryPriority: ['sound', 'touch', 'sight', 'temperature'], audio: ['cord singing under strain', 'canoe against stones', 'adult tide argument'], proseMode: 'quiet-aftermath',
    },
  },
  's4-accusation': {
    id: 's4-accusation', sceneId: 'the-eighth-hook', kicker: 'The next morning', title: 'The Hidden Hook',
    body: `<p>The hook lay in Kalos’s palm with a small twist in it, as though the bone had tried to turn away from what had happened.</p><p>The wrapping had slipped down the shank. Two turns crossed where six should have lain flat.</p><p>Mara stood over Neri. Oren said Neri had carried the tray crooked. The fishing party waited with the bad patience of adults losing tide.</p><p>“Neri will stay and rebind it unless someone has something useful to say.”</p><p>Neri looked at Kalos. Not pleading. That would have been easier. Only looking.</p>`,
    choicePrompt: 'What does Kalos change?', choiceHint: 'Truth, blame, repair, and reputation can move separately.',
    choices: [
      choice('confess-hook', 'Put the hook down and own the binding', 'Clear Neri publicly and accept the morning’s cost.', 'hinge', ['public', 'reveal', 'cost now']),
      choice('test-hooks', 'Ask Mara to test the rest of the tray', 'Let workmanship approach the truth before Kalos does.', 'hinge', ['public', 'evidence', 'delay']),
      choice('joke-oren-basket', 'Use Oren’s unfinished basket against him', 'Break the accusation’s authority without necessarily clearing Neri.', 'hinge', ['public', 'conceal', 'status']),
      choice('silent-hook', 'Lower his eyes and let Mara decide', 'Keep the fishing morning and allow the burden to move.', 'hinge', ['private', 'conceal', 'cost later']),
    ],
    freeText: true,
    visual: {
      plate: 'hook-close', title: 'The Hidden Hook', subtitle: 'Early morning · the tide is leaving', palette: 'bone-fiber', framing: 'split', weather: 'mist', tide: 'falling', light: 'gray morning with one warm shed lamp', motion: ['canoe tapping stones', 'wet lines moving in light wind'], focalObjects: ['failed binding', 'Mara’s open hand', 'empty path to the canoe'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'still shoulders, warm hook-print in palm', gaze: 'neri', proximity: 'near' },
        { id: 'neri', position: 'center', posture: 'standing under accusation without pleading', gaze: 'kalos', proximity: 'isolated' },
        { id: 'oren', position: 'foreground-right', posture: 'leaning forward too eagerly', gaze: 'mara', proximity: 'crowded' },
      ], sensoryPriority: ['touch', 'sound', 'sight', 'temperature'], audio: ['canoe knock', 'gull at cleaning stones', 'Neri’s shortened breath'], proseMode: 'narrowing',
    },
  },
  's5-mara': {
    id: 's5-mara', sceneId: 'the-work-nobody-saw', kicker: 'After the canoe leaves', title: 'The Work Shed',
    body: `<p>The adults did not praise Kalos for telling them which child had spoiled the hook. They left without him.</p><p>Mara set six fresh hooks between them in the work shed. Wet cedar, old oil, and the sweet beginning of rot lived beneath the floorboards.</p><p>“You wanted to go when they called,” she said.</p><p>Outside, the children found a laugh he could not hear clearly enough to improve.</p>`,
    choicePrompt: 'Mara waits for an answer.', choiceHint: 'Explanation may reveal motive, protect self-image, or make shame survivable.',
    choices: [
      choice('plain-want', '“I wanted to go.”', 'Give the motive without improving it.', 'standard', ['plain', 'ownership']),
      choice('thought-held', '“I thought it would hold.”', 'Defend the judgment rather than the lie.', 'standard', ['partial', 'technical']),
      choice('feared-left', '“They would have left me.”', 'Name the social fear beneath the work.', 'standard', ['vulnerable', 'direct']),
      choice('fish-joke', '“It held until a fish asked it to work.”', 'Use precision and laughter to remain in the room.', 'standard', ['humor', 'shame']),
    ],
    freeText: true,
    visual: {
      plate: 'work-shed', title: 'The Work Shed', subtitle: 'The desirable morning has gone', palette: 'empty-tide', framing: 'medium', weather: 'mist', tide: 'falling', light: 'one lamp over the work mat', motion: ['lamp flame responding to cracks', 'fiber straightening between Mara’s fingers'], focalObjects: ['six fresh hooks', 'empty doorway', 'Mara’s work knife'], characters: [
        { id: 'mara', position: 'foreground-right', posture: 'seated opposite with fiber separated', gaze: 'kalos', proximity: 'near' },
        { id: 'kalos', position: 'foreground-left', posture: 'sitting where the canoe used to be visible', gaze: 'object', proximity: 'near' },
      ], sensoryPriority: ['smell', 'touch', 'sound', 'sight'], audio: ['fiber against bone', 'distant children', 'water leaving beneath floor'], proseMode: 'quiet-aftermath',
    },
  },
  's5-evidence': {
    id: 's5-evidence', sceneId: 'the-work-nobody-saw', kicker: 'The testing mat', title: 'What the Hands Know',
    body: `<p>Mara pulled each hook against the heel of her hand.</p><p>The first held. The second held. On the third, the fiber shifted. On the fourth, it moved enough for everyone to see.</p><p>“These were not dropped,” Mara said.</p><p>Neri looked at Kalos. Suspicion did not arrive all at once. It found a place to stand.</p>`,
    choicePrompt: 'Evidence has narrowed the room.', choiceHint: 'Kalos may complete the truth, divide it, or let Mara finish it.',
    choices: [
      choice('complete-truth', 'Tell the whole sequence now', 'Name the hurried work and the false report.', 'standard', ['reveal', 'delayed']),
      choice('admit-hurried', 'Admit poor work but not the prior claim', 'Clear Neri while preserving part of the concealment.', 'standard', ['partial', 'protect']),
      choice('implicate-lesson', 'Say Neri showed him the method', 'Place the failed work inside a true but misleading fact.', 'standard', ['deceptive', 'shared blame']),
      choice('let-mara-infer', 'Say nothing more', 'Allow Mara to reconstruct what Kalos will not state.', 'standard', ['silent', 'evidence']),
    ],
    freeText: true,
    visual: {
      plate: 'hook-close', title: 'What the Hands Know', subtitle: 'Evidence before confession', palette: 'bone-fiber', framing: 'close', weather: 'mist', tide: 'falling', light: 'flat light over hands and fiber', motion: ['binding shifting beneath Mara’s thumb', 'Neri’s breath changing'], focalObjects: ['third hook', 'fourth hook', 'Mara’s testing hand'], characters: [
        { id: 'mara', position: 'center', posture: 'pulling the binding against her palm', gaze: 'object', proximity: 'crowded' },
        { id: 'neri', position: 'foreground-right', posture: 'watching Kalos instead of the hook', gaze: 'kalos', proximity: 'near' },
        { id: 'kalos', position: 'foreground-left', posture: 'waiting while evidence speaks', gaze: 'object', proximity: 'near' },
      ], sensoryPriority: ['touch', 'sight', 'sound'], audio: ['fiber slipping', 'canoe crew shifting impatiently', 'no laughter'], proseMode: 'technical',
    },
  },
  's5-neri': {
    id: 's5-neri', sceneId: 'the-work-nobody-saw', kicker: 'Behind the shed', title: 'What Neri Knows',
    body: `<p>Kalos found Neri after the canoe had gone.</p><p>The older boy had separated the wet fiber from the dry and laid each strand beside his knee. From the beach came the children’s voices. Oren was telling the basket joke and putting the laugh in the wrong place.</p><p>“You tied it,” Neri said.</p><p>It was not a question.</p>`,
    choicePrompt: 'There is no room to manage but this one.', choiceHint: 'Private repair, public truth, purchased silence, and denial create different debts.',
    choices: [
      choice('late-public-confession', '“Yes. Come with me. I’ll tell Mara.”', 'Return after the first failure to return.', 'hinge', ['public', 'reversal']),
      choice('buy-silence', '“I’ll finish it if you keep quiet.”', 'Exchange labor for control of the account.', 'hinge', ['private', 'transaction']),
      choice('silent-repair', 'Sit beside him and take the hook', 'Reduce the burden while preserving the public lie.', 'hinge', ['private', 'repair']),
      choice('deny-to-neri', '“No. You carried the tray.”', 'Look at the person paying and preserve the claim.', 'hinge', ['deceptive', 'direct']),
      choice('leave-for-beach', 'Go before the older children stop calling', 'Take the immediate reward and let the story harden.', 'hinge', ['status', 'cost later']),
    ],
    freeText: true,
    visual: {
      plate: 'work-shed', title: 'What Neri Knows', subtitle: 'Behind the shed after departure', palette: 'empty-tide', framing: 'medium', weather: 'mist', tide: 'falling', light: 'cool daylight outside, narrow warm line from the shed', motion: ['fiber laid into straight rows', 'children moving beyond sight'], focalObjects: ['damaged hook', 'sorted fiber', 'Kalos’s marked fingers'], characters: [
        { id: 'neri', position: 'center', posture: 'working without looking up', gaze: 'object', proximity: 'isolated' },
        { id: 'kalos', position: 'foreground-left', posture: 'standing where Neri can see his feet', gaze: 'neri', proximity: 'near' },
      ], sensoryPriority: ['sound', 'touch', 'sight', 'smell'], audio: ['older children on the beach', 'fiber drawn straight', 'one gull on the roof'], proseMode: 'quiet-aftermath',
    },
  },
  's5-prevention': {
    id: 's5-prevention', sceneId: 'the-work-nobody-saw', kicker: 'The work no one praises', title: 'Ordinary Morning',
    body: `<p>The fishing party left with eight hooks.</p><p>Kalos remained in the yard because Mara had already found him another task. There was no rescue, no accusation, and no reason for the adults to repeat his name.</p><p>Neri sat beside him splitting fiber. Oren’s basket failed at one corner and had to be opened again.</p><p>The morning was good. It was not a story.</p>`,
    choicePrompt: 'What does Kalos do with a success no one celebrates?', choiceHint: 'Quiet competence can be inhabited, advertised, shared, or converted into advantage.',
    choices: [
      choice('continue-ordinary', 'Begin the next hook', 'Let prevention remain ordinary work.', 'standard', ['quiet', 'follow-through']),
      choice('seek-mara-credit', 'Ask Mara whether she saw the eighth binding', 'Make the invisible cost visible to the person who assigned it.', 'standard', ['credit', 'public']),
      choice('help-oren-basket', 'Show Oren where his basket opened', 'Offer useful knowledge without demanding the laugh.', 'standard', ['repair', 'rivalry']),
      choice('share-with-neri', 'Give Neri first choice of the dry fiber', 'Mark shared work through material rather than speech.', 'standard', ['private', 'relationship']),
    ],
    freeText: true,
    visual: {
      plate: 'drying-racks', title: 'Ordinary Morning', subtitle: 'A failure prevented and therefore unseen', palette: 'salt-daylight', framing: 'wide', weather: 'mist', tide: 'falling', light: 'morning widening over an ordinary yard', motion: ['canoe leaving', 'new fiber being split', 'Oren reopening his basket'], focalObjects: ['complete hook tray', 'fresh fiber', 'opened basket corner'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'already beginning another task', gaze: 'object', proximity: 'near' },
        { id: 'neri', position: 'center', posture: 'splitting fiber beside him', gaze: 'object', proximity: 'near' },
        { id: 'oren', position: 'background-right', posture: 'opening failed basket work', gaze: 'object', proximity: 'isolated' },
      ], sensoryPriority: ['touch', 'sound', 'temperature', 'sight'], audio: ['canoe pushing off', 'fiber splitting', 'ordinary work voices'], proseMode: 'quiet-aftermath',
    },
  },
  's6-close': {
    id: 's6-close', sceneId: 'season-close', kicker: 'Several weeks later', title: 'When the Fire Burns Low',
    body: `<p>The first long rain came at night and stayed for three days.</p><p>People moved their sleeping places away from the roof seams. Pali developed a cough that sounded worse after the fire burned low. Oren learned a version of the hook story in which he had known the truth from the beginning.</p><p>Near the coals, Kalos listened to Veya tell an older story badly enough that everyone else had room to remember it differently.</p><p>The eighth hook had left the household. What remained was work, and the account of work, and the people who no longer held the same account.</p>`,
    choicePrompt: 'The season continues beyond this prototype.',
    choiceHint: 'The closing packet preserves only what can matter later.',
    choices: [],
    freeText: false,
    visual: {
      plate: 'longhouse', title: 'When the Fire Burns Low', subtitle: 'First long rain', palette: 'low-fire', framing: 'wide', weather: 'light-rain', tide: 'high', light: 'low fire and rain-dark doorway', motion: ['rain at roof seams', 'people shifting bedding', 'smoke pressed downward'], focalObjects: ['banked coals', 'moved bedding', 'one hook-shaped shadow'], characters: [
        { id: 'kalos', position: 'foreground-left', posture: 'listening beside the low fire', gaze: 'veya', proximity: 'crowded' },
        { id: 'veya', position: 'center', posture: 'telling a story with both hands still', gaze: 'group', proximity: 'crowded' },
        { id: 'pali', position: 'foreground-right', posture: 'sleeping near Eda with a cough', gaze: 'away', proximity: 'touching' },
      ], sensoryPriority: ['sound', 'temperature', 'smell', 'touch', 'sight'], audio: ['rain on cedar', 'Pali coughing', 'Veya’s low voice'], proseMode: 'quiet-aftermath',
    },
  },
};

export const KNOWN_BEAT_IDS = new Set(Object.keys(BEATS));
