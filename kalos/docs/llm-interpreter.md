# Optional LLM interpreter contract

Kalos does not require a language model. The local deterministic interpreter remains the default and guarantees a complete offline playthrough.

A provider can later register an adapter at:

```js
globalThis.KALOS_LLM_INTERPRETER = {
  async interpret(request) {
    // Return an authored or bounded dynamic proposal.
  }
};
```

## Request

The request contains a constrained snapshot rather than unrestricted internal state:

- current scene ID and player text;
- available authored choices and their intent tags;
- Kalos’s relevant condition and established tendencies;
- nearby characters, knowledge, and relationship summaries;
- world and household conditions;
- a causal history summary;
- the allowed outcome schema and mutation limits.

## Proposal types

### Authored proposal

```js
{
  kind: 'authored',
  choiceId: 'hook_confession',
  confidence: 0.93,
  interpretation: 'Kalos accepts public responsibility.'
}
```

### Dynamic proposal

```js
{
  kind: 'dynamic',
  confidence: 0.84,
  interpretation: 'Kalos admits the bad binding to Neri first, then returns with him.',
  nextSceneId: 'work_nobody_saw',
  outcome: {
    prose: '...',
    mutations: [
      { path: 'relationships.neri.trust', op: 'add', value: 0.06 }
    ],
    memories: ['...'],
    visualModifiers: {
      attentionAdd: ['Neri’s hands around the repaired hook']
    }
  }
}
```

## Validation

Every provider response is treated as a proposal. The validator rejects or constrains:

- unknown or unreachable scene IDs;
- writes to schema version, chronicle identity, or objective event history;
- arbitrary deletion of relationships, memories, promises, or public narratives;
- mutation paths outside the allowlist;
- relationship or tendency changes beyond scene-specific limits;
- unearned death, major injury, resource collapse, or political transformation;
- prose or conduct inconsistent with Kalos’s age and the content boundaries;
- claims that contradict established witnesses or physical conditions.

A rejected, timed-out, or unavailable provider silently falls back to deterministic interpretation. The state history records whether the result came from offline rules, an authored provider match, or a validated dynamic proposal.

## Design principle

The future model may affect outcomes, not merely paraphrase menu choices, but it must operate inside the simulation’s causal and canonical boundaries. Simulation remains authoritative; generated prose and interpretation remain revisable renderings and bounded proposals.
