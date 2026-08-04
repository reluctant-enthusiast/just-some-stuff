import test from 'node:test';
import assert from 'node:assert/strict';
import {BEATS} from '../dist/content.js';
import {initialState,resolveChoice,validateProposal,applyProposal} from '../dist/engine.js';

function stateAt(beatId){
  const state=initialState(42);
  const beat=BEATS[beatId];
  state.beatId=beatId;
  state.sceneId=beat.sceneId;
  state.presentation=structuredClone(beat.visual);
  if(beatId==='s4-return')state.flags.hookOutcome='hurried';
  return state;
}

test('every authored choice has a resolvable deterministic transition',()=>{
  for(const beat of Object.values(BEATS)){
    for(const choice of beat.choices){
      const state=stateAt(beat.id);
      if(beat.id==='s4-return'&&choice.id==='report-exactly')state.flags.hookOutcome='careful';
      assert.doesNotThrow(()=>resolveChoice(state,choice.id,choice.label,'menu'),`${beat.id}:${choice.id}`);
    }
  }
});

test('the prevention route avoids the accusation',()=>{
  let state=initialState(7);
  state.beatId='s3-call';state.sceneId=BEATS['s3-call'].sceneId;state.presentation=BEATS['s3-call'].visual;
  state=resolveChoice(state,'finish-before-race').state;
  state.beatId='s4-return';state.sceneId=BEATS['s4-return'].sceneId;state.presentation=BEATS['s4-return'].visual;
  const result=resolveChoice(state,'claim-finished');
  assert.equal(result.state.beatId,'s4-morning-clear');
});

test('concealment can reward Kalos immediately while harming Neri',()=>{
  const state=stateAt('s4-accusation');
  const before=state.characters.neri.relationshipToKalos.trust;
  const result=resolveChoice(state,'silent-hook');
  assert.equal(result.state.beatId,'s5-neri');
  assert.ok(result.state.characters.neri.relationshipToKalos.trust<before);
  assert.ok(result.state.publicStories.length>0);
});

test('LLM proposals can change allowed state but cannot rewrite canon',()=>{
  const state=stateAt('s4-accusation');
  const proposal={contractVersion:'1.0',proposalId:'p1',source:'llm',confidence:.9,intent:{disclosure:'partial',tactic:'private warning',target:'neri',desiredEffect:'protect him quietly'},actionText:'Warn Neri and move beside him.',novelOutcome:{body:'<p>Kalos moves before he speaks.</p>',returnToBeatId:'s5-neri'},patch:[
    {op:'increment',path:'characters.neri.relationshipToKalos.trust',value:.05},
    {op:'set',path:'events.0.summary',value:'rewritten history'}
  ],rationaleTags:['age-plausible']};
  const validation=validateProposal(proposal,state);
  assert.equal(validation.accepted,true);
  assert.equal(validation.acceptedPatch.length,1);
  assert.equal(validation.rejectedPatch.length,1);
  const changed=applyProposal(state,proposal,validation);
  assert.ok(changed.characters.neri.relationshipToKalos.trust>state.characters.neri.relationshipToKalos.trust);
});

test('severe un-authored AI events are rejected',()=>{
  const state=stateAt('s3-race');
  const proposal={contractVersion:'1.0',proposalId:'p2',source:'llm',confidence:.9,intent:{disclosure:'not-applicable',tactic:'push',target:'oren',desiredEffect:'win'},actionText:'Push Oren.',nextBeatId:'s3-laugh',patch:[],objectiveEvent:{id:'bad',sceneId:'race-beneath-the-fish',summary:'Oren dies in the fall.',facts:['Oren dies.'],witnesses:['kalos'],tags:['severe']},rationaleTags:['dramatic']};
  assert.equal(validateProposal(proposal,state).accepted,false);
});
