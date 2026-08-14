import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

export interface INewPromptData {
  titulo: string;
  acao: string;
  prompt: string;
  segmento: string;
  categoria: string;
  funcionaCom: string;
}

const ACOES = ['Analisar', 'Perguntar', 'Resumir', 'Criar', 'Encontrar', 'Aprender', 'Otimizar', 'Se preparar', 'Entender'];
const SEGMENTOS = ['Comercial', 'DP', 'Financeiro', 'Infra', 'Projetos', 'RH', 'Analista Funcional', 'Desenvolvedor(a)', 'Gerente de Projetos'];
const CATEGORIAS = ['Área', 'Função'];
const FUNCIONA_COM = ['Outlook', 'Teams', 'OneNote', 'Word', 'Excel', 'PowerPoint', 'Power BI', 'M365 Copilot', 'Copilot Studio', 'D365 CCaaS / Customer Service', 'D365 Customer Insights - Journeys', 'D365 Sales', 'Fabric', 'Power Apps', 'Power Automate', 'Power Pages', 'Whiteboard'];

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c] as string);
}

function optionsHtml(items: string[]): string {
  return ['<option value="">-- Selecione --</option>', ...items.map(i => `<option value="${escapeHtml(i)}">${escapeHtml(i)}</option>`)].join('');
}

export default class NewPromptDialog extends BaseDialog {
  public result: INewPromptData | undefined;

  public render(): void {
    this.domElement.innerHTML = `
      <style>
        .np-wrap {
          font-family: 'Segoe UI', sans-serif;
          display: flex;
          flex-direction: column;
          width: 560px;
          max-width: 90vw;
          height: 80vh;
          max-height: 720px;
          box-sizing: border-box;
          background: #fff;
        }
        .np-header {
          padding: 20px 24px 16px;
          border-bottom: 1px solid #edebe9;
          flex-shrink: 0;
        }
        .np-header h2 { margin: 0; font-size: 20px; font-weight: 600; color: #323130; }
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
          border-top: 1px solid #edebe9;
          flex-shrink: 0;
          background: #faf9f8;
        }
        .np-field { margin-bottom: 16px; }
        .np-label {
          display: block;
          font-weight: 600;
          margin-bottom: 6px;
          color: #323130;
          font-size: 13px;
        }
        .np-label .np-req { color: #a4262c; margin-left: 2px; }
        .np-field input[type="text"],
        .np-field select,
        .np-field textarea {
          width: 100%;
          padding: 8px 12px;
          box-sizing: border-box;
          border: 1px solid #c8c6c4;
          border-radius: 8px;
          font-family: inherit;
          font-size: 14px;
          color: #323130;
          background: #fff;
          transition: border-color 0.15s ease, box-shadow 0.15s ease;
        }
        .np-field input[type="text"]:hover,
        .np-field select:hover,
        .np-field textarea:hover {
          border-color: #a19f9d;
        }
        .np-field input[type="text"]:focus,
        .np-field select:focus,
        .np-field textarea:focus {
          outline: none;
          border-color: #0078d4;
          box-shadow: 0 0 0 2px rgba(0, 120, 212, 0.2);
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
        .np-error { color: #a4262c; font-size: 12px; margin-top: 4px; display: none; }
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
        .np-btn.primary { background: #0078d4; color: #fff; }
        .np-btn.primary:hover { background: #106ebe; }
        .np-btn.primary:disabled { background: #a19f9d; cursor: not-allowed; }
        .np-btn.secondary { background: #fff; border-color: #8a8886; color: #323130; }
        .np-btn.secondary:hover { background: #f3f2f1; }
        .np-status { font-size: 13px; margin-right: auto; color: #605e5c; }
        .np-status.error { color: #a4262c; }
      </style>
      <div class="np-wrap">
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
            <select id="np-acao" name="acao">${optionsHtml(ACOES)}</select>
            <div class="np-error" data-for="acao">Selecione uma ação.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-prompt">Prompt<span class="np-req">*</span></label>
            <textarea id="np-prompt" name="prompt"></textarea>
            <div class="np-error" data-for="prompt">Informe o prompt.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-segmento">Segmento<span class="np-req">*</span></label>
            <select id="np-segmento" name="segmento">${optionsHtml(SEGMENTOS)}</select>
            <div class="np-error" data-for="segmento">Selecione um segmento.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-categoria">Categoria<span class="np-req">*</span></label>
            <select id="np-categoria" name="categoria">${optionsHtml(CATEGORIAS)}</select>
            <div class="np-error" data-for="categoria">Selecione uma categoria.</div>
          </div>

          <div class="np-field">
            <label class="np-label" for="np-funcionaCom">Funciona com<span class="np-req">*</span></label>
            <select id="np-funcionaCom" name="funcionaCom">${optionsHtml(FUNCIONA_COM)}</select>
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

