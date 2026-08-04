# Kalos architecture

## Runtime layers

1. **Canon and scene content** define characters, household roles, locations, authored choices, visual templates, and branch prose.
2. **Simulation state** distinguishes objective events, individual beliefs, public narratives, relationships, promises, memories, rumors, bodily condition, household resources, and presentation state.
3. **Deterministic adjudication** validates choices, applies bounded mutations, advances the scene graph, and records causal history before rendering prose.
4. **Presentation composition** derives a visual-state packet from objective conditions, staging, Kalos’s attention, continuity objects, and accumulated memory.
5. **Persistence** stores a versioned chronicle and performs best-effort migration from the original Hidden Hook prototype.
6. **Optional interpretation** may propose authored or dynamic outcomes through a validated adapter; the game remains complete without it.

## State principles

- Objective truth is separate from what each person witnessed, inferred, heard, or later repeated.
- Relationships retain distinct dimensions such as affection, trust, respect, resentment, dependence, fear, and ease.
- Memories preserve sensory anchors and interpretations rather than functioning as a transcript.
- Presentation is causal. Weather, staging, objects, posture, palette, sound cues, prose rhythm, and choice magnitude derive from state rather than a generic mood filter.
- Childhood choices alter tendencies, access, relationships, and later meaning; they do not deterministically define adult Kalos.

## Scene structure

Milestone 1 contains five authored anchor sequences, subdivided into data-driven scene nodes. Choices can be light, tactical, quiet, reflective, or hinge decisions. Each meaningful outcome changes some combination of:

- future affordances;
- relationship or belief state;
- public narrative or rumor;
- persistent memory;
- household or bodily conditions;
- visual and sensory presentation.

## Hosted bundle

The hosted build uses a small loader plus compressed payload chunks because the full compiled app contains embedded temporary scene plates. The loader joins the chunks, decompresses the bundle in-browser, and mounts the app. A static fallback remains readable if a viewer blocks JavaScript.
