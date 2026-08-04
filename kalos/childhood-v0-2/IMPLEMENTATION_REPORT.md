# Implementation Report — Milestones 0 and 1

## Delivered

The proof-of-concept has been rebuilt as a structured, static TypeScript-oriented application with a dependency-free browser runtime.

The legacy `/kalos/` prototype is unchanged. This build is staged at `/kalos/childhood-v0-2/` pending review and physical iPhone playtesting.

## Milestone 0

### State model

The new chronicle distinguishes:

- objective events;
- individual knowledge and belief;
- public stories and distortion;
- multidimensional relationships;
- promises;
- memories and their current meanings;
- skills and behavioral tendencies;
- bodily condition;
- household resources;
- opportunities and access;
- presentation state;
- immutable action history.

### Deterministic engine

All authored choices resolve through a deterministic transition graph. The game requires no server, model, subscription, or network request after its static assets load.

### Optional AI interpreter

An external adapter may later:

- interpret free-text actions;
- select a non-obvious authored action;
- propose a novel local outcome;
- append a new objective event;
- change relationships, beliefs, resources, opportunities, memories, promises, or presentation;
- return the player to a different authored beat.

Every proposal is checked against allowed state paths, effect limits, known return beats, childhood content boundaries, immutable history, and authored-event requirements. A proposal may be partially accepted. Failure returns to the complete offline interpreter.

### Presentation

Every beat supplies structured visual metadata:

- plate;
- palette;
- framing;
- weather and tide;
- light and motion;
- focal objects;
- character placement, posture, gaze, and proximity;
- sensory priority;
- ambient sound;
- prose mode.

The current temporary art is generated locally from responsive SVG scene plates. It is designed to be replaced by synchronized painted art without changing the state contract.

## Milestone 1

The expanded slice contains five substantive scene groups and a closing bridge:

1. household predawn and first food;
2. the path to work, technical instruction, a small accident, and the eight-hook assignment;
3. the older children’s race, changing rules, physical risk, and the social use of laughter;
4. the eighth hook, including a genuine prevention route and several failure/account routes;
5. differentiated aftermath with Mara, evidence, Neri, or quiet competence;
6. a first-long-rain closing scene that preserves state for the next childhood expansion.

The Hidden Hook is no longer inevitable. A player who accepts the immediate social cost of finishing or inspecting the work may prevent the accusation entirely. Failure routes retain the original distinction among truth, blame, restitution, reputation, and social reward.

## Validation built into the repository

Automated tests cover:

- every authored choice having a deterministic transition;
- prevention avoiding the accusation;
- successful concealment producing immediate social benefit and delayed relational cost;
- partial acceptance of an AI proposal;
- rejection of unauthorized canon rewriting;
- rejection of severe un-authored AI events;
- JavaScript syntax;
- local static-asset references;
- absence of remote runtime dependencies.

## Known limits

- Other household members have goals and relational state but do not yet run through a full offscreen scheduler. That belongs to Milestone 2.
- Temporary SVG plates are atmospheric composition tests, not final painted art.
- The offline free-text interpreter is heuristic. It deliberately asks for clarification rather than pretending to understand every unusual action.
- The AI adapter contract is implemented, but no provider or model is bundled.
- The current content pool is a dense vertical slice, not yet the full Age Eight season described in the expansion plan.
- Physical-device Safari testing remains necessary before replacing the existing public `/kalos/` entry point.
