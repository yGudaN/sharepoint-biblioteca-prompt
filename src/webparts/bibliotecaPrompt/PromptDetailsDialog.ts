import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { IPromptChoices, DEFAULT_PROMPT_CHOICES } from './promptChoices';

export interface IPromptDetailsData {
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

function optionsHtml(items: string[], selected: string): string {
  const opts = ['<option value="">-- Selecione --</option>'];
  for (const i of items) {
    const sel = i === selected ? ' selected' : '';
    opts.push(`<option value="${escapeHtml(i)}"${sel}>${escapeHtml(i)}</option>`);
  }
  return opts.join('');
}

export default class PromptDetailsDialog extends BaseDialog {
  public result: IPromptDetailsData | undefined;
  private _initial: IPromptDetailsData;
  private _choices: IPromptChoices;
  private _theme: 'light' | 'dark';
  private _editing: boolean = false;

  constructor(initial: IPromptDetailsData, choices?: IPromptChoices, theme: 'light' | 'dark' = 'light') {
    super();
    this._initial = initial;
    this._choices = choices || DEFAULT_PROMPT_CHOICES;
    this._theme = theme;
  }

  public render(): void {
    const d = this._initial;
    this.domElement.innerHTML = `
      <style>
        .pd-wrap {
          --pd-surface: #ffffff;
          --pd-surface-alt: #faf9f8;
          --pd-hover: #f3f2f1;
          --pd-text: #323130;
          --pd-text-2: #605e5c;
          --pd-text-3: #a19f9d;
          --pd-border: #edebe9;
          --pd-border-2: #c8c6c4;
          --pd-border-3: #8a8886;
          --pd-primary: #0078d4;
          --pd-primary-h: #106ebe;
          --pd-on-primary: #ffffff;
          --pd-danger: #a4262c;
          --pd-focus-ring: rgba(0, 120, 212, 0.2);

          font-family: 'Segoe UI', sans-serif;
          display: flex; flex-direction: column;
          width: 560px; max-width: 90vw; height: 80vh; max-height: 720px;
          box-sizing: border-box; background: var(--pd-surface); color: var(--pd-text);
        }
        .pd-wrap.pd-dark {
          --pd-surface: #2b2b2b;
          --pd-surface-alt: #262626;
          --pd-hover: #3b3a39;
          --pd-text: #f3f2f1;
          --pd-text-2: #c8c6c4;
          --pd-text-3: #8a8886;
          --pd-border: #3b3a39;
          --pd-border-2: #605e5c;
          --pd-border-3: #8a8886;
          --pd-primary: #4cb2ff;
          --pd-primary-h: #6cc0ff;
          --pd-on-primary: #1f1f1f;
          --pd-danger: #f1707b;
          --pd-focus-ring: rgba(76, 178, 255, 0.35);
        }
        .pd-header { display: flex; align-items: center; justify-content: space-between; padding: 20px 24px 16px; border-bottom: 1px solid var(--pd-border); flex-shrink: 0; }
        .pd-header h2 { margin: 0; font-size: 20px; font-weight: 600; color: var(--pd-text); }
        .pd-body { flex: 1; overflow-y: auto; padding: 20px 24px; }
        .pd-footer { display: flex; justify-content: flex-end; align-items: center; gap: 8px; padding: 14px 24px; border-top: 1px solid var(--pd-border); flex-shrink: 0; background: var(--pd-surface-alt); }
        .pd-field { margin-bottom: 16px; }
        .pd-label { display: block; font-weight: 600; margin-bottom: 6px; color: var(--pd-text); font-size: 13px; }
        .pd-label .pd-req { color: var(--pd-danger); margin-left: 2px; }
        .pd-field input[type="text"], .pd-field select, .pd-field textarea {
          width: 100%; padding: 8px 12px; box-sizing: border-box; border: 1px solid var(--pd-border-2); border-radius: 8px;
          font-family: inherit; font-size: 14px; color: var(--pd-text); background: var(--pd-surface);
          transition: border-color 0.15s ease, box-shadow 0.15s ease;
        }
        .pd-field input[type="text"]:hover:not(:disabled), .pd-field select:hover:not(:disabled), .pd-field textarea:hover:not(:disabled) { border-color: var(--pd-text-3); }
        .pd-field input[type="text"]:focus, .pd-field select:focus, .pd-field textarea:focus {
          outline: none; border-color: var(--pd-primary); box-shadow: 0 0 0 2px var(--pd-focus-ring);
        }
        .pd-field input:disabled, .pd-field select:disabled, .pd-field textarea:disabled {
          background: var(--pd-surface-alt); color: var(--pd-text); opacity: 1; cursor: default; border-color: var(--pd-border);
        }
        .pd-field textarea { min-height: 140px; resize: vertical; line-height: 1.4; }
        .pd-field select {
          appearance: none; -webkit-appearance: none;
          background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3e%3cpath fill='%23605e5c' d='M6 8L0 2l1.4-1.4L6 5.2 10.6.6 12 2z'/%3e%3c/svg%3e");
          background-repeat: no-repeat; background-position: right 12px center; padding-right: 32px;
        }
        .pd-wrap.pd-dark .pd-field select {
          background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3e%3cpath fill='%23c8c6c4' d='M6 8L0 2l1.4-1.4L6 5.2 10.6.6 12 2z'/%3e%3c/svg%3e");
        }
        .pd-field select:disabled { background-image: none; }
        .pd-error { color: var(--pd-danger); font-size: 12px; margin-top: 4px; display: none; }
        .pd-btn { padding: 8px 22px; border: 1px solid transparent; border-radius: 6px; cursor: pointer; font-size: 14px; font-family: inherit; font-weight: 600; }
        .pd-btn.primary { background: var(--pd-primary); color: var(--pd-on-primary); }
        .pd-btn.primary:hover { background: var(--pd-primary-h); }
        .pd-btn.secondary { background: var(--pd-surface); border-color: var(--pd-border-3); color: var(--pd-text); }
        .pd-btn.secondary:hover { background: var(--pd-hover); }
        .pd-mode-badge { font-size: 12px; padding: 2px 10px; border-radius: 12px; background: var(--pd-border); color: var(--pd-text-2); font-weight: 600; }
      </style>
      <div class="pd-wrap${this._theme === 'dark' ? ' pd-dark' : ''}">
        <div class="pd-header">
          <h2 id="pd-title-mode">Detalhes do prompt</h2>
          <span class="pd-mode-badge" id="pd-mode-badge">Somente leitura</span>
        </div>
        <form id="pd-form" class="pd-body" novalidate>
          <div class="pd-field">
            <label class="pd-label" for="pd-titulo">Título<span class="pd-req">*</span></label>
            <input type="text" id="pd-titulo" maxlength="255" value="${escapeHtml(d.titulo)}" disabled />
            <div class="pd-error" data-for="titulo">Informe o título.</div>
          </div>
          <div class="pd-field">
            <label class="pd-label" for="pd-acao">Ação<span class="pd-req">*</span></label>
            <select id="pd-acao" disabled>${optionsHtml(this._choices.acoes, d.acao)}</select>
            <div class="pd-error" data-for="acao">Selecione uma ação.</div>
          </div>
          <div class="pd-field">
            <label class="pd-label" for="pd-prompt">Prompt<span class="pd-req">*</span></label>
            <textarea id="pd-prompt" disabled>${escapeHtml(d.prompt)}</textarea>
            <div class="pd-error" data-for="prompt">Informe o prompt.</div>
          </div>
          <div class="pd-field">
            <label class="pd-label" for="pd-segmento">Segmento<span class="pd-req">*</span></label>
            <select id="pd-segmento" disabled>${optionsHtml(this._choices.segmentos, d.segmento)}</select>
            <div class="pd-error" data-for="segmento">Selecione um segmento.</div>
          </div>
          <div class="pd-field">
            <label class="pd-label" for="pd-categoria">Categoria<span class="pd-req">*</span></label>
            <select id="pd-categoria" disabled>${optionsHtml(this._choices.categorias, d.categoria)}</select>
            <div class="pd-error" data-for="categoria">Selecione uma categoria.</div>
          </div>
          <div class="pd-field">
            <label class="pd-label" for="pd-funcionaCom">Funciona com<span class="pd-req">*</span></label>
            <select id="pd-funcionaCom" disabled>${optionsHtml(this._choices.funcionaCom, d.funcionaCom)}</select>
            <div class="pd-error" data-for="funcionaCom">Selecione uma opção.</div>
          </div>
        </form>
        <div class="pd-footer">
          <button type="button" class="pd-btn secondary" id="pd-close">Fechar</button>
          <button type="button" class="pd-btn primary" id="pd-edit">Editar</button>
          <button type="button" class="pd-btn primary" id="pd-save" style="display:none">Salvar</button>
        </div>
      </div>
    `;

    const root = this.domElement;
    (root.querySelector('#pd-close') as HTMLButtonElement).addEventListener('click', () => {
      this.result = undefined;
      this.close().catch(() => { /* noop */ });
    });
    (root.querySelector('#pd-edit') as HTMLButtonElement).addEventListener('click', () => {
      this._editing = true;
      this._syncMode();
    });
    (root.querySelector('#pd-save') as HTMLButtonElement).addEventListener('click', () => {
      const data = this._collectAndValidate();
      if (!data) return;
      this.result = data;
      this.close().catch(() => { /* noop */ });
    });
  }

  private _syncMode(): void {
    const root = this.domElement;
    const fields = root.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>('#pd-titulo, #pd-acao, #pd-prompt, #pd-segmento, #pd-categoria, #pd-funcionaCom');
    fields.forEach((f) => { f.disabled = !this._editing; });
    (root.querySelector('#pd-title-mode') as HTMLElement).textContent = this._editing ? 'Editar prompt' : 'Detalhes do prompt';
    (root.querySelector('#pd-mode-badge') as HTMLElement).textContent = this._editing ? 'Edição' : 'Somente leitura';
    (root.querySelector('#pd-edit') as HTMLElement).style.display = this._editing ? 'none' : 'inline-block';
    (root.querySelector('#pd-save') as HTMLElement).style.display = this._editing ? 'inline-block' : 'none';
  }

  private _collectAndValidate(): IPromptDetailsData | undefined {
    const root = this.domElement;
    const titulo = (root.querySelector('#pd-titulo') as HTMLInputElement).value.trim();
    const acao = (root.querySelector('#pd-acao') as HTMLSelectElement).value;
    const prompt = (root.querySelector('#pd-prompt') as HTMLTextAreaElement).value.trim();
    const segmento = (root.querySelector('#pd-segmento') as HTMLSelectElement).value;
    const categoria = (root.querySelector('#pd-categoria') as HTMLSelectElement).value;
    const funcionaCom = (root.querySelector('#pd-funcionaCom') as HTMLSelectElement).value;

    const errors: Record<string, boolean> = { titulo: !titulo, acao: !acao, prompt: !prompt, segmento: !segmento, categoria: !categoria, funcionaCom: !funcionaCom };
    let hasError = false;
    root.querySelectorAll('.pd-error').forEach((el) => {
      const key = (el as HTMLElement).dataset.for as string;
      const show = errors[key];
      (el as HTMLElement).style.display = show ? 'block' : 'none';
      if (show) hasError = true;
    });
    if (hasError) return undefined;
    return { titulo, acao, prompt, segmento, categoria, funcionaCom };
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: true };
  }
}
