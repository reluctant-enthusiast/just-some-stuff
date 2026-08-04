# Kalos — Milestones 0–1

A mobile-first interactive literary narrative following Kalos at age eight.

## Playable build

Open `index.html` from the hosted `/kalos/` route. The app is deterministic and works without a language model.

## Milestone 0

The prototype has been rebuilt around a structured TypeScript source architecture with separate modules for:

- canon and household data;
- scene content;
- deterministic adjudication;
- belief, relationship, memory, promise, rumor, and world state;
- versioned persistence and legacy-save migration;
- scene-reactive visual-state packets;
- optional validated language-model interpretation.

The hosted deployment contains the compiled browser bundle. The complete TypeScript source package is maintained as the release source artifact.

## Milestone 1

The Age-Eight opening now spans five anchor sequences:

1. **Smoke Before Daylight** — household atmosphere, breakfast, family attachment, and work distribution.
2. **The Drying Racks** — Mara’s instruction, Neri’s skill, Kalos’s attention, and practical responsibility.
3. **The Race Beneath the Fish** — play, humor, status, cruelty, loyalty, and physical confidence.
4. **The Eighth Hook** — the original damaged-hook incident, now conditioned by earlier conduct.
5. **The Work Nobody Saw** — a substantial, branch-responsive aftermath in which confession, repair, concealment, and social reward separate.

## Binding literary canon

The approved prose, focalization, dialogue, interaction, scene-architecture, and Age Eight Arc II standards are consolidated in [`docs/literary-prose-constitution.md`](docs/literary-prose-constitution.md).

That document is binding editorial canon for the Age Eight phase and the presumptive baseline for later phases. It requires, among other things:

- lexically mature but child-bounded focalization;
- delayed interpretation rather than immediate thematic explanation;
- longer, inhabited scenes with autonomous NPC action;
- restrained player interaction through attention, belief, action, narrative focus, memory, and occasional lingering;
- equal prose care for prevention, ordinary pleasure, and quiet outcomes;
- revised Arc II order: Low Water, The Sweet Oil, The Borrowed Bed, Words Through the Wall, The Cave at Half Tide or tide-pool route, and What Stayed.

## Optional LLM bridge

The offline interpreter remains authoritative by default. A later provider may register `globalThis.KALOS_LLM_INTERPRETER` and return either:

- an authored-choice proposal; or
- a bounded dynamic outcome proposal.

Provider proposals pass through the same validator as deterministic outcomes. The model cannot bypass causal state, invent arbitrary scene jumps, overwrite objective history, or apply unbounded relationship changes.

See `docs/llm-interpreter.md` for the contract.

## Development

The source build uses TypeScript with no runtime framework dependency. Tests cover scene integrity, authored routes, deterministic free-text handling, optional provider proposals, bounded mutations, persistence, migration, and visual-state continuity.