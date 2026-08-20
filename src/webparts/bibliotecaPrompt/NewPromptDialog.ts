import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { IPromptChoices, DEFAULT_PROMPT_CHOICES } from './promptChoices';

export interface INewPromptData {
  titulo: string;
  acao: string;
  prompt: string;
  segmento: string;
  categoria: string;
  funcionaCom: string;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c] as string);
}

function optionsHtml(items: string[]): string {
  return ['<option value="">-- Selecione --</option>', ...items.map(i => `<option value="${escapeHtml(i)}">${escapeHtml(i)}</option>`)].join('');
}

export default class NewPromptDialog extends BaseDialog {
  public result: INewPromptData | undefined;
  private _choices: IPromptChoices;
  private _theme: 'light' | 'dark';
  private _blockOutside: ((ev: Event) => void) | undefined;
  private _focusGuardOut: ((ev: FocusEvent) => void) | undefined;
  private _focusGuardIn: ((ev: FocusEvent) => void) | undefined;
  private _lastFocused: HTMLElement | undefined;
  private _cleanupDone: boolean = false;

  constructor(choices?: IPromptChoices, theme: 'light' | 'dark' = 'light') {
    super();
    this._choices = choices || DEFAULT_PROMPT_CHOICES;
    this._theme = theme;
  }

  public onBeforeClose(): Promise<void> {
    this._cleanupBlocker();
    return Promise.resolve();
  }

  public close(): Promise<void> {
    this._cleanupBlocker();
    return super.close();
  }

  private _cleanupBlocker(): void {
    if (this._cleanupDone) return;
    this._cleanupDone = true;
    if (this._blockOutside) {
      document.removeEventListener('mousedown', this._blockOutside, true);
      document.removeEventListener('pointerdown', this._blockOutside, true);
      this._blockOutside = undefined;
    }
    if (this._focusGuardOut) {
      document.removeEventListener('focusout', this._focusGuardOut, true);
      this._focusGuardOut = undefined;
    }
    if (this._focusGuardIn) {
      document.removeEventListener('focusin', this._focusGuardIn, true);
      this._focusGuardIn = undefined;
    }
    this._lastFocused = undefined;
  }

  private _installFocusGuard(root: HTMLElement): void {
    this._blockOutside = (ev: Event) => {
      if (root.contains(ev.target as Node)) ev.stopImmediatePropagation();
    };
    document.addEventListener('mousedown', this._blockOutside, true);
    document.addEventListener('pointerdown', this._blockOutside, true);

    this._focusGuardIn = (ev: FocusEvent) => {
      const tgt = ev.target as HTMLElement;
      if (root.contains(tgt) && tgt !== root) this._lastFocused = tgt;
    };
    document.addEventListener('focusin', this._focusGuardIn, true);

    this._focusGuardOut = (ev: FocusEvent) => {
      const wentTo = ev.relatedTarget as HTMLElement | null;
      const wasIn = root.contains(ev.target as Node);
      if (!wasIn) return;
      if (wentTo && root.contains(wentTo)) return;
      const restore = this._lastFocused;
      if (!restore || !root.contains(restore)) return;
      setTimeout(() => {
        if (this._lastFocused === restore && document.contains(restore) && root.contains(restore)) {
          restore.focus();
        }
      }, 0);
    };
    document.addEventListener('focusout', this._focusGuardOut, true);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <style>
        .np-wrap {
          --np-surface: #ffffff;
          --np-surface-alt: #faf9f8;
          --np-hover: #f3f2f1;
          --np-text: #323130;
          --np-text-2: #605e5c;
          --np-text-3: #a19f9d;
          --np-border: #edebe9;
          --np-border-2: #c8c6c4;
          --np-border-3: #8a8886;
          --np-primary: #0078d4;
          --np-primary-h: #106ebe;
          --np-on-primary: #ffffff;
          --np-danger: #a4262c;
          --np-focus-ring: rgba(0, 120, 212, 0.2);
          --np-arrow: '%23605e5c';

          font-family: 'Segoe UI', sans-serif;
          color-scheme: light;
          display: flex;
          flex-direction: column;
          width: 560px;
          max-width: 90vw;
          height: 80vh;
          max-height: 720px;
          box-sizing: border-box;
          background: var(--np-surface);
          color: var(--np-text);
        }
        .np-wrap.np-dark {
          --np-surface: #2b2b2b;
          --np-surface-alt: #262626;
          --np-hover: #3b3a39;
          --np-text: #f3f2f1;
          --np-text-2: #c8c6c4;
          --np-text-3: #8a8886;
          --np-border: #3b3a39;
          --np-border-2: #605e5c;
          --np-border-3: #8a8886;
          --np-primary: #4cb2ff;
          --np-primary-h: #6cc0ff;
          --np-on-primary: #1f1f1f;
          --np-danger: #f1707b;
          --np-focus-ring: rgba(76, 178, 255, 0.35);
          --np-arrow: '%23c8c6c4';
          color-scheme: dark;
        }
        .np-header {
          padding: 20px 24px 16px;
          border-bottom: 1px solid var(--np-border);
          flex-shrink: 0;
        }
        .np-header h2 { margin: 0; font-size: 20px; font-weight: 600; color: var(--np-text); }
        .np-body {
          flex: 1;
          overflow-y: auto;
          padding: 20px 24px;
        }
        .np-footer {
          display: flex;
          justify-content: flex-end;
          align-items: center;
          gap: 8px;
          padding: 14px 24px;
          border-top: 1px solid var(--np-border);
          flex-shrink: 0;
          background: var(--np-surface-alt);
        }
        .np-field { margin-bottom: 16px; }
        .np-label {
          display: block;
          font-weight: 600;
          margin-bottom: 6px;
          color: var(--np-text);
          font-size: 13px;
        }
        .np-label .np-req { color: var(--np-danger); margin-left: 2px; }
        .np-field input[type="text"],
        .np-field select,
        .np-field textarea {
          width: 100%;
          padding: 8px 12px;
          box-sizing: border-box;
          border: 1px solid var(--np-border-2);
          border-radius: 8px;
          font-family: inherit;
          font-size: 14px;
          color: var(--np-text);
          background: var(--np-surface);
          transition: border-color 0.15s ease, box-shadow 0.15s ease;
        }
        .np-field input[type="text"]:hover,
        .np-field select:hover,
        .np-field textarea:hover {
          border-color: var(--np-text-3);
        }
        .np-field input[type="text"]:focus,
        .np-field select:focus,
        .np-field textarea:focus {
          outline: none;
          border-color: var(--np-primary);
          box-shadow: 0 0 0 2px var(--np-focus-ring);
        }
        .np-field textarea {
          min-height: 140px;
          resize: vertical;
          line-height: 1.4;
        }
        .np-field select {
          appearance: none;
          -webkit-appearance: none;
          background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3e%3cpath fill='%23605e5c' d='M6 8L0 2l1.4-1.4L6 5.2 10.6.6 12 2z'/%3e%3c/svg%3e");
          background-repeat: no-repeat;
          background-position: right 12px center;
          padding-right: 32px;
        }
        .np-wrap.np-dark .np-field select {
          background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3e%3cpath fill='%23c8c6c4' d='M6 8L0 2l1.4-1.4L6 5.2 10.6.6 12 2z'/%3e%3c/svg%3e");
        }
        .np-error { color: var(--np-danger); font-size: 12px; margin-top: 4px; display: none; }
        .np-btn {
          padding: 8px 22px;
          border: 1px solid transparent;
          border-radius: 6px;
          cursor: pointer;
          font-size: 14px;
          font-family: inherit;
          font-weight: 600;
          transition: background-color 0.15s ease, border-color 0.15s ease;
        }
        .np-btn.primary { background: var(--np-primary); color: var(--np-on-primary); }
        .np-btn.primary:hover { background: var(--np-primary-h); }
        .np-btn.primary:disabled { background: var(--np-text-3); cursor: not-allowed; }
        .np-btn.secondary { background: var(--np-surface); border-color: var(--np-border-3); color: var(--np-text); }
        .np-btn.secondary:hover { background: var(--np-hover); }
        .np-status { font-size: 13px; margin-right: auto; color: var(--np-text-2); }
        .np-status.error { color: var(--np-danger); }
      </style>
      <div class="np-wrap${this._theme === 'dark' ? ' np-dark' : ''}">
        <div class="np-header">
          <h2>Novo prompt</h2>
        </div>
        <form id="np-form" class="np-body" novalidate>
          <div class="np-field">
            <label class="np-label" for="np-titulo">Título<span class="np-req">*</span></label>
            <input type="text" id="np-titulo" name="titulo" maxlength="255" autocomplete="off" />
            <div class="np-error" data-for="titulo">Informe o título.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-acao">Ação<span class="np-req">*</span></label>
            <select id="np-acao" name="acao">${optionsHtml(this._choices.acoes)}</select>
            <div class="np-error" data-for="acao">Selecione uma ação.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-prompt">Prompt<span class="np-req">*</span></label>
            <textarea id="np-prompt" name="prompt"></textarea>
            <div class="np-error" data-for="prompt">Informe o prompt.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-segmento">Segmento<span class="np-req">*</span></label>
            <select id="np-segmento" name="segmento">${optionsHtml(this._choices.segmentos)}</select>
            <div class="np-error" data-for="segmento">Selecione um segmento.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-categoria">Categoria<span class="np-req">*</span></label>
            <select id="np-categoria" name="categoria">${optionsHtml(this._choices.categorias)}</select>
            <div class="np-error" data-for="categoria">Selecione uma categoria.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-funcionaCom">Funciona com<span class="np-req">*</span></label>
            <select id="np-funcionaCom" name="funcionaCom">${optionsHtml(this._choices.funcionaCom)}</select>
            <div class="np-error" data-for="funcionaCom">Selecione uma opção.</div>
          </div>
        </form>
        <div class="np-footer">
          <div class="np-status" id="np-status"></div>
          <button type="button" class="np-btn secondary" id="np-cancel">Cancelar</button>
          <button type="button" class="np-btn primary" id="np-submit">Salvar</button>
        </div>
      </div>
    `;

    const root = this.domElement;
    this._installFocusGuard(root);

    const form = this.domElement.querySelector('#np-form') as HTMLFormElement;
    const cancelBtn = this.domElement.querySelector('#np-cancel') as HTMLButtonElement;
    const submitBtn = this.domElement.querySelector('#np-submit') as HTMLButtonElement;

    cancelBtn.addEventListener('click', () => {
      this.result = undefined;
      this.close().catch(() => { /* noop */ });
    });

    const submit = (): void => {
      const data = this._collectAndValidate();
      if (!data) return;
      this.result = data;
      this.close().catch(() => { /* noop */ });
    };

    submitBtn.addEventListener('click', submit);
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      submit();
    });
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: true };
  }

  private _collectAndValidate(): INewPromptData | undefined {
    const root = this.domElement;
    const titulo = (root.querySelector('#np-titulo') as HTMLInputElement).value.trim();
    const acao = (root.querySelector('#np-acao') as HTMLSelectElement).value;
    const prompt = (root.querySelector('#np-prompt') as HTMLTextAreaElement).value.trim();
    const segmento = (root.querySelector('#np-segmento') as HTMLSelectElement).value;
    const categoria = (root.querySelector('#np-categoria') as HTMLSelectElement).value;
    const funcionaCom = (root.querySelector('#np-funcionaCom') as HTMLSelectElement).value;

    const errors: Record<string, boolean> = {
      titulo: !titulo,
      acao: !acao,
      prompt: !prompt,
      segmento: !segmento,
      categoria: !categoria,
      funcionaCom: !funcionaCom
    };

    let hasError = false;
    root.querySelectorAll('.np-error').forEach((el) => {
      const key = (el as HTMLElement).dataset.for as string;
      const show = errors[key];
      (el as HTMLElement).style.display = show ? 'block' : 'none';
      if (show) hasError = true;
    });

    if (hasError) return undefined;
    return { titulo, acao, prompt, segmento, categoria, funcionaCom };
  }
}

