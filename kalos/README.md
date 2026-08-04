# Kalos — Milestones 0–1

This directory hosts the first expanded childhood vertical slice of **Kalos**, an interactive literary life narrative.

## Play

Open `index.html` through a static host. The public route contains the validated standalone build directly, so it loads as an ordinary web page on mobile and desktop without runtime Base64 chunk assembly or browser-side decompression.

## Milestone 0: engine foundation

The build is organized around a data-driven TypeScript architecture compiled into a self-contained static application. The underlying source separates:

- canon, household members, and locations;
- scene content and choices;
- deterministic adjudication;
- objective events, beliefs, memories, relationships, rumors, and obligations;
- versioned persistence and migration;
- visual-state composition;
- responsive UI and developer inspection.

The game remains fully playable offline without a language model.

### Optional LLM interpreter

The engine exposes an optional interpretation boundary for free-text actions. A future model may propose an `OutcomeProposal` that changes later outcomes, including validated factual additions, relationship effects, and next-scene selection. Proposals are schema-checked and limited to permitted domains; invalid, unavailable, or rejected proposals fall back to deterministic adjudication. The model is therefore influential but not authoritative over canon.

## Milestone 1: expanded Age Eight sequence

The playable sequence now includes:

1. **Smoke Before Daylight** — household life, food, family tension, and the morning assignment.
2. **The Drying Racks** — practical work, Mara's instruction, Neri's competence, and Kalos's attention.
3. **The Race Beneath the Fish** — play, physical confidence, Oren's status strategy, and social improvisation.
4. **The Eighth Hook** — the existing hook incident, now conditioned by prior conduct and relationships.
5. **The Work Nobody Saw** — a playable aftermath in which public truth, private repair, reputation, and trust diverge.

The chapter contains multiple decision tiers, free-text actions, branch-dependent prose, responsive interludes, and a structured visual-state packet for every beat.

## Validation

The build was tested across complete authored routes, deterministic free-text handling, optional interpreter proposals, state migration, save/export/import, visual packet generation, and representative phone and desktop viewports. A physical-device Safari test remains desirable before treating the interface as production-ready.

## Legacy

`legacy-v1.html` preserves the earlier proof-of-concept route.

The pre-fix compressed-payload deployment is preserved in Git on branch `preserve/kalos-atob-pre-fix-20260803`. The historical `payload-*.txt` and `bundle-*.txt` files remain in the repository for traceability but are not loaded by the public route.
