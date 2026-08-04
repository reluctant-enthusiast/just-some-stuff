# Kalos — Childhood v0.2

This directory implements Milestones 0 and 1 of the Kalos childhood expansion.

## What is implemented

### Milestone 0 — structured engine

- versioned chronicle state;
- objective events separated from beliefs, public stories, promises, memories, and relationships;
- deterministic transition graph with no network dependency;
- best-effort migration from the legacy `kalos-hidden-hook` localStorage save;
- responsive visual-state packets for future art systems;
- optional, validated LLM interpreter boundary;
- export/import, replay, developer inspection, reading focus, and accessibility settings;
- pure runtime modules with automated transition and proposal-validation tests.

### Milestone 1 — expanded Age Eight vertical slice

1. **Smoke Before Daylight** — household attachment, food, overheard pressure, and affiliation.
2. **The Drying Racks** — practical learning, interruption, and work assignment.
3. **The Race Beneath the Fish** — play, peer status, humor, risk, and the social lure that competes with work.
4. **The Eighth Hook** — prevention, shared responsibility, inspection, concealment, evidence, confession, and displaced blame.
5. **The Work Nobody Saw** — route-specific public or private aftermath.
6. **When the Fire Burns Low** — a short persistence bridge showing the season continuing beyond the incident.

The slice is intentionally denser than the original proof of concept. It makes the Hidden Hook preventable on some routes, preserves the original accusation architecture on others, and gives quiet competence a playable consequence rather than treating dramatic failure as inevitable.

## Runtime

The deployed application is static and dependency-free:

```text
index.html
styles.css
dist/content.js
dist/engine.js
dist/app.js
```

Serve the directory over HTTP. Opening through an attachment preview may block JavaScript on iOS; a normally hosted URL is the supported experience.

## Optional LLM interpreter

The game is complete in `offline` mode. An LLM may later change outcomes—not only prose—by registering an adapter:

```js
window.Kalos.registerInterpreter({
  name: 'Bounded narrative interpreter',
  async interpret(context) {
    return {
      contractVersion: '1.0',
      proposalId: crypto.randomUUID(),
      source: 'llm',
      confidence: 0.84,
      intent: {
        disclosure: 'partial',
        tactic: 'private warning',
        target: 'neri',
        desiredEffect: 'protect Neri without public confession'
      },
      actionText: context.input,
      novelOutcome: {
        body: '<p>Kalos moves beside Neri before Mara finishes speaking.</p>',
        returnToBeatId: 's5-neri'
      },
      patch: [
        {
          op: 'increment',
          path: 'characters.neri.relationshipToKalos.trust',
          value: 0.05
        }
      ],
      rationaleTags: ['age-plausible', 'causal', 'bounded']
    };
  }
});
```

The deterministic validator may accept, partially accept, or reject the proposal. It enforces:

- allowed state paths;
- relationship-effect magnitude limits;
- known return beats;
- childhood content boundaries;
- immutable prior history;
- authored-event requirements for death, severe injury, major loss, and exile.

If the adapter is absent, fails, times out at the integration layer, or produces an invalid proposal, the offline interpreter remains fully functional.

## Canon status

Only **Kalos** is a locked name. Eda, Ruvan, Seli, Pali, Veya, Mara, Neri, Oren, the Long Roof household, and exact kinship rules remain provisional canon.

## Next milestone

Milestone 2 should add the event ecology and season scheduler only after playtesting this slice. It should introduce conditional weather, visitors, illness, minor injury, resource pressure, quiet-grace events, and offscreen NPC schedules without turning every ordinary day into an incident.
