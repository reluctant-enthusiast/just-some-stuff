import { BEATS, KNOWN_BEAT_IDS } from './content.js';
import {
  applyAcceptedProposal,
  importState,
  loadState,
  recordHistory,
  resetState,
  saveState,
} from './engine.js';
import { interpretPlayerInput } from './interpreter.js';
import { visualStatePacket } from './presentation.js';
import { resolveChoice } from './story.js';
import type { ChronicleState, OutcomeProposal } from './schema.js';

export interface RenderSnapshot {
  state: ChronicleState;
  body: string;
  echo?: string;
  visualPacket: string;
}

export class KalosController extends EventTarget {
  #state: ChronicleState;

  constructor(initial = loadState()) {
    super();
    this.#state = initial;
  }

  get state(): ChronicleState {
    return structuredClone(this.#state);
  }

  choose(choiceId: string, actionText = choiceId): RenderSnapshot {
    const resolution = resolveChoice(this.#state, choiceId, actionText, 'menu');
    this.#state = resolution.state;
    saveState(this.#state);
    this.dispatchEvent(new CustomEvent('change', { detail: this.state }));
    return {
      state: this.state,
      body: resolution.transitionBody,
      echo: resolution.echo,
      visualPacket: visualStatePacket(this.#state),
    };
  }

  async act(input: string): Promise<RenderSnapshot | { ambiguous: true }> {
    const beat = BEATS[this.#state.beatId];
    if (!beat) throw new Error(`Missing beat ${this.#state.beatId}.`);
    const { proposal, validation } = await interpretPlayerInput(input, beat, this.#state, KNOWN_BEAT_IDS);
    if (!validation.accepted) return { ambiguous: true };

    this.#state = applyAcceptedProposal(this.#state, proposal, validation.acceptedPatch);
    if (proposal.novelOutcome) {
      this.#state = recordHistory(
        this.#state,
        beat,
        proposal.selectedChoiceId ?? 'novel-outcome',
        input,
        'llm',
        proposal,
      );
      saveState(this.#state);
      this.dispatchEvent(new CustomEvent('change', { detail: this.state }));
      return {
        state: this.state,
        body: proposal.novelOutcome.body,
        echo: 'A validated interpreter proposal changed the local outcome.',
        visualPacket: visualStatePacket(this.#state),
      };
    }

    if (!proposal.selectedChoiceId) return { ambiguous: true };
    const resolution = resolveChoice(
      this.#state,
      proposal.selectedChoiceId,
      input,
      proposal.source === 'llm' ? 'llm' : 'offline',
    );
    this.#state = resolution.state;
    saveState(this.#state);
    this.dispatchEvent(new CustomEvent('change', { detail: this.state }));
    return {
      state: this.state,
      body: resolution.transitionBody,
      echo: resolution.echo,
      visualPacket: visualStatePacket(this.#state),
    };
  }

  registerInterpreter(adapter: NonNullable<Window['KALOS_LLM_ADAPTER']>): void {
    window.KALOS_LLM_ADAPTER = adapter;
  }

  restart(): ChronicleState {
    this.#state = resetState();
    this.dispatchEvent(new CustomEvent('change', { detail: this.state }));
    return this.state;
  }

  import(text: string): ChronicleState {
    this.#state = importState(text);
    saveState(this.#state);
    this.dispatchEvent(new CustomEvent('change', { detail: this.state }));
    return this.state;
  }
}

const controller = new KalosController();
window.Kalos = {
  registerInterpreter: (adapter) => controller.registerInterpreter(adapter),
  getState: () => controller.state,
};

export { controller };
export type { OutcomeProposal };
