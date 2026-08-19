import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import NewPromptDialog, { INewPromptData } from './NewPromptDialog';
import PromptDetailsDialog, { IPromptDetailsData } from './PromptDetailsDialog';
import { IPromptChoices, DEFAULT_PROMPT_CHOICES } from './promptChoices';

export interface IBibliotecaPromptWebPartProps {
  targetListTitle: string;
  favoritesListTitle: string;
  promptIdField: string;
  copyFields: string;
  extraToolColors: string;
}

interface IPromptItem {
  Id: number;
  Title: string;
  Prompt?: string;
  Categoria?: string;
  Categoria0?: string;
  Funcionacom?: string;
  A_x00e7__x00e3_o?: string;
}

const FAV_FILLED = '#FFB900';
const FAV_EMPTY = '#605E5C';
const PROMPT_PREVIEW_LEN = 220;

const ACTION_ICON: { [k: string]: string } = {
  'Analisar': '🔍',
  'Perguntar': '❓',
  'Resumir': '📋',
  'Criar': '✏️',
  'Encontrar': '🔎',
  'Aprender': '📚',
  'Otimizar': '⚡',
  'Se preparar': '🎯',
  'Entender': '💡'
};
const ACTION_ICON_DEFAULT = '📝';

const TOOL_COLOR: { [k: string]: string } = {
  'Outlook':                              '#0078D4',
  'Teams':                                '#6264A7',
  'OneNote':                              '#7719AA',
  'Word':                                 '#2B579A',
  'Excel':                                '#217346',
  'PowerPoint':                           '#B7472A',
  'Power BI':                             '#D4A017',
  'M365 Copilot':                         '#8A8886',
  'Copilot Studio':                       '#6264A7',
  'D365 CCaaS / Customer Service':        '#5C2D91',
  'D365 Customer Insights - Journeys':    '#0078D4',
  'D365 Sales':                           '#0078D4',
  'Fabric':                               '#1E8C5C',
  'Power Apps':                           '#742774',
  'Power Automate':                       '#0066FF',
  'Power Pages':                          '#742774',
  'Whiteboard':                           '#0078D4'
};
const TOOL_COLOR_DEFAULT = '#605E5C';

const HEX_COLOR_RE = /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i;

function parseExtraToolColors(raw: string | undefined): { [k: string]: string } {
  const out: { [k: string]: string } = {};
  if (!raw) return out;
  for (const line of raw.split(/\r?\n/)) {
    const eq = line.indexOf('=');
    if (eq < 0) continue;
    const name = line.substring(0, eq).trim();
    const color = line.substring(eq + 1).trim();
    if (name && HEX_COLOR_RE.test(color)) out[name] = color;
  }
  return out;
}

function esc(s: unknown): string {
  const str = s === undefined || s === null ? '' : String(s);
  return str.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c] as string);
}

function stripHtml(s: string | undefined): string {
  if (!s) return '';
  const tmp = document.createElement('div');
  tmp.innerHTML = s;
  return tmp.textContent || tmp.innerText || '';
}

export default class BibliotecaPromptWebPart extends BaseClientSideWebPart<IBibliotecaPromptWebPartProps> {
  private _items: IPromptItem[] = [];
  private _favMap: Map<number, number> = new Map();
  private _search: string = '';
  private _viewMode: 'all' | 'favorites' | 'app' | 'area' | 'funcao' = 'all';
  private _pending: Set<number> = new Set();
  private _loaded: boolean = false;
  private _choices: IPromptChoices = DEFAULT_PROMPT_CHOICES;
  private _theme: 'light' | 'dark' = 'light';

  public render(): void {
    this.domElement.innerHTML = this._shellHtml();
    this._wire();
    if (!this._loaded) {
      this._loadAll().catch((err) => this._showError(err));
    } else {
      this._renderRows();
    }
  }

  protected onInit(): Promise<void> {
    this._theme = this._detectInitialTheme();
    return Promise.resolve();
  }

  private _detectInitialTheme(): 'light' | 'dark' {
    try {
      const stored = window.localStorage.getItem('sbp-theme');
      if (stored === 'light' || stored === 'dark') return stored;
    } catch { /* ignore */ }
    const teams = (this.context.sdks as unknown as { microsoftTeams?: { context?: { theme?: string } } })?.microsoftTeams;
    const t = teams?.context?.theme;
    if (t === 'dark' || t === 'contrast') return 'dark';
    if (t === 'default') return 'light';
    try {
      if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) return 'dark';
    } catch { /* ignore */ }
    return 'light';
  }

  private _toggleTheme(): void {
    this._theme = this._theme === 'dark' ? 'light' : 'dark';
    try { window.localStorage.setItem('sbp-theme', this._theme); } catch { /* ignore */ }
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _shellHtml(): string {
    return `
      <style>
        .bp-wrap {
          --bp-bg: #ffffff;
          --bp-surface: #ffffff;
          --bp-surface-alt: #faf9f8;
          --bp-hover: #f3f2f1;
          --bp-text: #323130;
          --bp-text-2: #605e5c;
          --bp-text-3: #a19f9d;
          --bp-border: #edebe9;
          --bp-border-2: #c8c6c4;
          --bp-primary: #0078d4;
          --bp-primary-h: #106ebe;
          --bp-primary-soft: rgba(0, 120, 212, 0.12);
          --bp-on-primary: #ffffff;
          --bp-danger: #a4262c;
          --bp-success: #107c10;
          --bp-shadow: rgba(0, 0, 0, 0.08);
          --bp-shadow-hover: rgba(0, 0, 0, 0.12);
          --bp-focus-ring: rgba(0, 120, 212, 0.2);

          font-family: 'Segoe UI', sans-serif;
          color: var(--bp-text);
          background: var(--bp-bg);
          padding: 16px;
          border-radius: 8px;
          min-height: 100vh;
          box-sizing: border-box;
          transition: background 0.2s ease, color 0.2s ease;
        }
        .bp-wrap.bp-dark {
          --bp-bg: #1f1f1f;
          --bp-surface: #2b2b2b;
          --bp-surface-alt: #262626;
          --bp-hover: #3b3a39;
          --bp-text: #f3f2f1;
          --bp-text-2: #c8c6c4;
          --bp-text-3: #8a8886;
          --bp-border: #3b3a39;
          --bp-border-2: #605e5c;
          --bp-primary: #4cb2ff;
          --bp-primary-h: #6cc0ff;
          --bp-primary-soft: rgba(76, 178, 255, 0.18);
          --bp-on-primary: #1f1f1f;
          --bp-danger: #f1707b;
          --bp-success: #6bb700;
          --bp-shadow: rgba(0, 0, 0, 0.4);
          --bp-shadow-hover: rgba(0, 0, 0, 0.55);
          --bp-focus-ring: rgba(76, 178, 255, 0.35);
        }
        .bp-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 8px; }
        .bp-title { font-size: 20px; font-weight: 600; margin: 0; color: var(--bp-text); }
        .bp-head-actions { display: flex; align-items: center; gap: 8px; }
        .bp-btn { padding: 8px 18px; border: 1px solid transparent; border-radius: 6px; cursor: pointer; font-size: 14px; font-family: inherit; font-weight: 600; }
        .bp-btn.primary { background: var(--bp-primary); color: var(--bp-on-primary); }
        .bp-btn.primary:hover { background: var(--bp-primary-h); }
        .bp-btn.primary:disabled { background: var(--bp-text-3); cursor: not-allowed; }
        .bp-theme-toggle {
          background: transparent; border: 1px solid var(--bp-border);
          border-radius: 999px; width: 36px; height: 36px;
          cursor: pointer; font-size: 16px; line-height: 1;
          display: inline-flex; align-items: center; justify-content: center;
          color: var(--bp-text);
        }
        .bp-theme-toggle:hover { background: var(--bp-hover); }
        .bp-toolbar { display: flex; gap: 12px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }
        .bp-search {
          flex: 1; min-width: 200px; padding: 8px 12px;
          border: 1px solid var(--bp-border-2); border-radius: 6px;
          font-family: inherit; font-size: 14px;
          background: var(--bp-surface); color: var(--bp-text);
        }
        .bp-search::placeholder { color: var(--bp-text-3); }
        .bp-search:focus { outline: none; border-color: var(--bp-primary); box-shadow: 0 0 0 2px var(--bp-focus-ring); }
        .bp-filters { display: flex; gap: 6px; flex-wrap: wrap; }
        .bp-filter {
          background: transparent;
          border: 1px solid var(--bp-border-2);
          padding: 6px 14px;
          border-radius: 999px;
          cursor: pointer;
          font-family: inherit;
          font-size: 13px;
          font-weight: 500;
          color: var(--bp-text-2);
          display: inline-flex;
          align-items: center;
          gap: 6px;
          transition: background 0.15s ease, color 0.15s ease, border-color 0.15s ease;
        }
        .bp-filter:hover { background: var(--bp-hover); color: var(--bp-text); }
        .bp-filter.active {
          background: var(--bp-primary-soft);
          border-color: var(--bp-primary);
          color: var(--bp-primary);
          font-weight: 600;
        }
        .bp-section-title {
          display: flex; align-items: baseline; gap: 8px;
          font-size: 15px; font-weight: 600;
          margin: 22px 0 10px;
          color: var(--bp-text);
        }
        .bp-section-title:first-of-type { margin-top: 4px; }
        .bp-section-count { font-size: 12px; color: var(--bp-text-3); font-weight: 500; }
        .bp-status { padding: 8px 12px; font-size: 13px; color: var(--bp-text-2); }
        .bp-status.error { color: var(--bp-danger); }
        .bp-empty { text-align: center; padding: 40px 20px; color: var(--bp-text-2); }

        .bp-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
          gap: 16px;
        }
        .bp-card {
          background: var(--bp-surface);
          border: 2px solid var(--bp-border);
          border-radius: 16px;
          padding: 20px;
          height: 280px;
          box-sizing: border-box;
          display: flex;
          flex-direction: column;
          transition: box-shadow 0.15s ease, transform 0.15s ease;
          position: relative;
          cursor: pointer;
        }
        .bp-card:hover {
          box-shadow: 0 4px 12px var(--bp-shadow-hover);
          transform: translateY(-1px);
        }
        .bp-card:focus { outline: 2px solid var(--bp-primary); outline-offset: 2px; }
        .bp-card-head {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 12px;
          gap: 8px;
        }
        .bp-action-icon {
          font-size: 22px;
          line-height: 1;
          flex-shrink: 0;
        }
        .bp-tool-name {
          display: inline-block;
          padding: 3px 10px;
          border-radius: 999px;
          background: #605E5C;
          color: #fff;
          font-size: 11px;
          font-weight: 600;
          line-height: 1.5;
          max-width: 160px;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .bp-toast {
          position: fixed;
          top: 24px;
          right: 24px;
          padding: 12px 20px;
          background: var(--bp-success);
          color: #fff;
          border-radius: 8px;
          box-shadow: 0 4px 16px var(--bp-shadow-hover);
          font-size: 14px;
          font-weight: 500;
          z-index: 9999;
          opacity: 0;
          transform: translateY(-8px);
          transition: opacity 0.2s ease, transform 0.2s ease;
          pointer-events: none;
        }
        .bp-toast.show { opacity: 1; transform: translateY(0); }
        .bp-toast.error { background: var(--bp-danger); }
        .bp-tool-badge {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          min-width: 28px;
          height: 28px;
          padding: 0 6px;
          border-radius: 999px;
          background: #605E5C;
          color: #fff;
          font-size: 11px;
          font-weight: 700;
          letter-spacing: 0.5px;
        }
        .bp-card-title {
          font-size: 15px;
          font-weight: 600;
          color: var(--bp-text);
          margin-bottom: 8px;
          line-height: 1.3;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .bp-card-body {
          font-size: 13px;
          color: var(--bp-text-2);
          line-height: 1.4;
          flex: 1;
          overflow: hidden;
          display: -webkit-box;
          -webkit-line-clamp: 4;
          -webkit-box-orient: vertical;
          margin-bottom: 12px;
        }
        .bp-card-foot {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding-top: 12px;
          border-top: 1px solid var(--bp-border);
          font-size: 12px;
          color: var(--bp-text-2);
        }
        .bp-card-meta {
          display: flex;
          align-items: center;
          gap: 6px;
          overflow: hidden;
          flex: 1;
          min-width: 0;
        }
        .bp-card-meta .bp-cat, .bp-card-meta .bp-act {
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .bp-card-meta .bp-sep { color: var(--bp-text-3); }
        .bp-card-actions {
          display: flex;
          align-items: center;
          gap: 6px;
          flex-shrink: 0;
          margin-left: 8px;
        }
        .bp-icon-btn {
          background: none;
          border: none;
          cursor: pointer;
          padding: 4px;
          font-size: 18px;
          line-height: 1;
          border-radius: 4px;
          color: var(--bp-text-2);
        }
        .bp-icon-btn:hover { background: var(--bp-hover); }
        .bp-icon-btn[disabled] { cursor: not-allowed; opacity: 0.5; }
        .bp-icon-btn.copied { color: var(--bp-success); }
      </style>
      <div class="bp-wrap${this._theme === 'dark' ? ' bp-dark' : ''}">
        <div class="bp-head">
          <h2 class="bp-title">Biblioteca de Prompts</h2>
          <div class="bp-head-actions">
            <button type="button" class="bp-theme-toggle" data-role="theme-toggle" title="Alternar tema">${this._theme === 'dark' ? '☀️' : '🌙'}</button>
            <button type="button" class="bp-btn primary" data-action="new">+ Novo prompt</button>
          </div>
        </div>
        <div class="bp-toolbar">
          <input type="text" class="bp-search" placeholder="Buscar por título, prompt, categoria..." data-role="search" />
          <div class="bp-filters" data-role="filters">
            <button type="button" class="bp-filter" data-view="all">Todos</button>
            <button type="button" class="bp-filter" data-view="favorites">⭐ Favoritos</button>
            <button type="button" class="bp-filter" data-view="app">App</button>
            <button type="button" class="bp-filter" data-view="area">Área</button>
            <button type="button" class="bp-filter" data-view="funcao">Função</button>
          </div>
        </div>
        <div class="bp-status" data-role="status">Carregando...</div>
        <div data-role="table-container"></div>
        <div class="bp-toast" data-role="toast"></div>
      </div>
    `;
  }

  private _wire(): void {
    const root = this.domElement;
    const newBtn = root.querySelector('[data-action="new"]') as HTMLButtonElement | null;
    if (newBtn) newBtn.addEventListener('click', () => this._openNewPromptDialog());

    const themeBtn = root.querySelector('[data-role="theme-toggle"]') as HTMLButtonElement | null;
    if (themeBtn) themeBtn.addEventListener('click', () => this._toggleTheme());

    const search = root.querySelector('[data-role="search"]') as HTMLInputElement | null;
    if (search) {
      search.value = this._search;
      search.addEventListener('input', () => {
        this._search = search.value.toLowerCase();
        this._renderRows();
      });
    }

    const filterBtns = root.querySelectorAll<HTMLButtonElement>('[data-role="filters"] button');
    filterBtns.forEach((b) => {
      const view = b.getAttribute('data-view');
      if (view === this._viewMode) b.classList.add('active');
      b.addEventListener('click', () => {
        this._viewMode = (view as 'all' | 'favorites' | 'app' | 'area' | 'funcao') || 'all';
        filterBtns.forEach((x) => x.classList.toggle('active', x === b));
        this._renderRows();
      });
    });
  }

  private _setStatus(text: string, isError: boolean = false): void {
    const el = this.domElement.querySelector('[data-role="status"]') as HTMLElement | null;
    if (!el) return;
    el.textContent = text;
    el.classList.toggle('error', isError);
    el.style.display = text ? 'block' : 'none';
  }

  private _showToast(text: string, type: 'success' | 'error' = 'success'): void {
    const el = this.domElement.querySelector('[data-role="toast"]') as HTMLElement | null;
    if (!el) return;
    el.textContent = text;
    el.classList.toggle('error', type === 'error');
    el.classList.add('show');
    setTimeout(() => { el.classList.remove('show'); }, 3000);
  }

  private _showError(err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[BibliotecaPrompt]', err);
    this._showToast(msg, 'error');
  }

  private _loadAll(): Promise<void> {
    if (!this.properties.targetListTitle || !this.properties.favoritesListTitle) {
      this._setStatus('Configure os títulos das listas no painel de propriedades da web part.', true);
      return Promise.resolve();
    }
    this._setStatus('Carregando...');
    return Promise.all([this._loadItems(), this._loadFavorites(), this._loadChoices()])
      .then(() => {
        this._loaded = true;
        this._setStatus('');
        this._renderRows();
      })
      .catch((err) => this._showError(err));
  }

  private _loadChoices(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.targetListTitle);
    const filter = "InternalName eq 'A_x00e7__x00e3_o' or InternalName eq 'Categoria' or InternalName eq 'Categoria0' or InternalName eq 'Funcionacom'";
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/fields?$filter=${encodeURIComponent(filter)}&$select=InternalName,Choices`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) throw new Error(`Falha ao carregar opções dos campos (HTTP ${r.status})`);
        return r.json();
      })
      .then((data: { value: Array<{ InternalName: string; Choices?: string[] }> }) => {
        const byField: { [k: string]: string[] } = {};
        for (const f of (data.value || [])) {
          byField[f.InternalName] = f.Choices || [];
        }
        const pick = (k: string, fallback: string[]): string[] =>
          (byField[k] && byField[k].length ? byField[k] : fallback);
        this._choices = {
          acoes: pick('A_x00e7__x00e3_o', DEFAULT_PROMPT_CHOICES.acoes),
          segmentos: pick('Categoria', DEFAULT_PROMPT_CHOICES.segmentos),
          categorias: pick('Categoria0', DEFAULT_PROMPT_CHOICES.categorias),
          funcionaCom: pick('Funcionacom', DEFAULT_PROMPT_CHOICES.funcionaCom)
        };
      })
      .catch((err) => {
        // não bloqueia a carga: usa defaults e loga
        console.warn('[BibliotecaPrompt] Falha ao carregar choices, usando defaults.', err);
        this._choices = DEFAULT_PROMPT_CHOICES;
      });
  }

  private _loadItems(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.targetListTitle);
    const select = ['Id', 'Title', 'A_x00e7__x00e3_o', 'Prompt', 'Categoria', 'Categoria0', 'Funcionacom'].join(',');
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/items?$select=${select}&$top=5000&$orderby=Title`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Falha ao carregar prompts (HTTP ${r.status}): ${t}`)));
        return r.json();
      })
      .then((data: { value: IPromptItem[] }) => {
        this._items = data.value || [];
      });
  }

  private _loadFavorites(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const favTitle = encodeURIComponent(this.properties.favoritesListTitle);
    const promptIdField = this.properties.promptIdField || 'PromptID';
    const userId = (this.context.pageContext as unknown as { legacyPageContext?: { userId?: number } }).legacyPageContext?.userId;
    const filter = userId ? `AuthorId eq ${userId}` : `Author/EMail eq '${this.context.pageContext.user.email}'`;
    const url = `${webUrl}/_api/web/lists/getByTitle('${favTitle}')/items?$filter=${encodeURIComponent(filter)}&$select=Id,${promptIdField}&$top=5000`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Falha ao carregar favoritos (HTTP ${r.status}): ${t}`)));
        return r.json();
      })
      .then((data: { value: Array<Record<string, unknown>> }) => {
        this._favMap.clear();
        for (const it of data.value || []) {
          const pid = Number(it[promptIdField]);
          const fid = Number(it.Id);
          if (!isNaN(pid) && !isNaN(fid)) this._favMap.set(pid, fid);
        }
      });
  }

  private _renderRows(): void {
    const container = this.domElement.querySelector('[data-role="table-container"]') as HTMLElement | null;
    if (!container) return;

    const filtered = this._items.filter((it) => {
      if (this._viewMode === 'favorites' && !this._favMap.has(it.Id)) return false;
      if (!this._search) return true;
      const hay = [it.Title, it.A_x00e7__x00e3_o, stripHtml(it.Prompt), it.Categoria, it.Categoria0, it.Funcionacom]
        .map((v) => (v || '').toString().toLowerCase())
        .join(' ');
      return hay.indexOf(this._search) >= 0;
    });

    if (filtered.length === 0) {
      const emptyMsg = this._viewMode === 'favorites'
        ? 'Você ainda não favoritou nenhum prompt.'
        : 'Nenhum prompt encontrado.';
      container.innerHTML = `<div class="bp-empty">${emptyMsg}</div>`;
      return;
    }

    const extraColors = parseExtraToolColors(this.properties.extraToolColors);
    const renderCard = (it: IPromptItem): string => {
      const isFav = this._favMap.has(it.Id);
      const promptText = stripHtml(it.Prompt);
      const preview = promptText.length > PROMPT_PREVIEW_LEN ? promptText.substring(0, PROMPT_PREVIEW_LEN) + '…' : promptText;
      const action = it.A_x00e7__x00e3_o || '';
      const actionIcon = ACTION_ICON[action] || ACTION_ICON_DEFAULT;
      const tool = it.Funcionacom || '';
      const toolColor = extraColors[tool] || TOOL_COLOR[tool] || TOOL_COLOR_DEFAULT;
      const toolHtml = tool
        ? `<span class="bp-tool-name" title="${esc(tool)}" style="background:${toolColor}">${esc(tool)}</span>`
        : '';
      const borderStyle = tool ? `border-color:${toolColor};` : '';
      const cat = it.Categoria || '';
      return `
        <div class="bp-card" data-id="${it.Id}" tabindex="0" role="button" title="Clique para ver detalhes" style="${borderStyle}">
          <div class="bp-card-head">
            <span class="bp-action-icon" title="${esc(action)}">${actionIcon}</span>
            ${toolHtml}
          </div>
          <div class="bp-card-title" title="${esc(it.Title)}">${esc(it.Title || '–')}</div>
          <div class="bp-card-body" title="${esc(promptText)}">${esc(preview || '–')}</div>
          <div class="bp-card-foot">
            <div class="bp-card-meta">
              <span aria-hidden="true">📁</span>
              <span class="bp-cat">${esc(cat || '–')}</span>
              <span class="bp-sep">|</span>
              <span class="bp-act">${esc(action || '–')}</span>
            </div>
            <div class="bp-card-actions">
              <button type="button" class="bp-icon-btn" data-copy="${it.Id}" title="Copiar prompt">📋</button>
              <button type="button" class="bp-icon-btn" data-fav="${it.Id}" style="color:${isFav ? FAV_FILLED : FAV_EMPTY}" title="${isFav ? 'Remover dos favoritos' : 'Favoritar'}">${isFav ? '★' : '☆'}</button>
            </div>
          </div>
        </div>
      `;
    };

    const groupField = this._viewMode === 'app' ? 'Funcionacom'
      : this._viewMode === 'area' ? 'Categoria0'
      : this._viewMode === 'funcao' ? 'Categoria'
      : null;

    if (groupField) {
      const groups: { [k: string]: IPromptItem[] } = {};
      for (const it of filtered) {
        const key = ((it as unknown as Record<string, unknown>)[groupField] as string) || 'Sem valor';
        if (!groups[key]) groups[key] = [];
        groups[key].push(it);
      }
      const sortedKeys = Object.keys(groups).sort((a, b) => {
        if (a === 'Sem valor') return 1;
        if (b === 'Sem valor') return -1;
        return a.localeCompare(b, 'pt-BR');
      });
      const sections = sortedKeys.map((k) => `
        <h3 class="bp-section-title">${esc(k)} <span class="bp-section-count">(${groups[k].length})</span></h3>
        <div class="bp-grid">${groups[k].map(renderCard).join('')}</div>
      `).join('');
      container.innerHTML = sections;
    } else {
      container.innerHTML = `<div class="bp-grid">${filtered.map(renderCard).join('')}</div>`;
    }

    container.querySelectorAll<HTMLElement>('.bp-card').forEach((card) => {
      card.addEventListener('click', (ev) => {
        const target = ev.target as HTMLElement;
        if (target.closest('.bp-card-actions')) return;
        const id = Number(card.getAttribute('data-id'));
        const item = this._items.find((x) => x.Id === id);
        if (item) this._openDetails(item);
      });
      card.addEventListener('keydown', (ev: KeyboardEvent) => {
        if (ev.key === 'Enter' || ev.key === ' ') {
          ev.preventDefault();
          const id = Number(card.getAttribute('data-id'));
          const item = this._items.find((x) => x.Id === id);
          if (item) this._openDetails(item);
        }
      });
    });
    container.querySelectorAll<HTMLButtonElement>('[data-fav]').forEach((btn) => {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const id = Number(btn.getAttribute('data-fav'));
        if (!isNaN(id)) this._toggleFavorite(id, btn).catch((err) => this._showError(err));
      });
    });
    container.querySelectorAll<HTMLButtonElement>('[data-copy]').forEach((btn) => {
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const id = Number(btn.getAttribute('data-copy'));
        const item = this._items.find((x) => x.Id === id);
        if (item) this._copyPrompt(item, btn).catch((err) => this._showError(err));
      });
    });
  }

  private _copyPrompt(item: IPromptItem, btn: HTMLButtonElement): Promise<void> {
    const text = stripHtml(item.Prompt);
    const done = (): void => {
      btn.textContent = '✓';
      btn.classList.add('copied');
      setTimeout(() => {
        btn.textContent = '📋';
        btn.classList.remove('copied');
      }, 1500);
    };
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text).then(done).catch((err) => this._showError(err));
    }
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); done(); } catch (e) { this._showError(e); }
    document.body.removeChild(ta);
    return Promise.resolve();
  }

  private _toggleFavorite(promptId: number, btn: HTMLButtonElement): Promise<void> {
    if (this._pending.has(promptId)) return Promise.resolve();
    this._pending.add(promptId);
    btn.disabled = true;

    const wasFav = this._favMap.has(promptId);
    const op = wasFav ? this._removeFavorite(promptId) : this._addFavorite(promptId);

    return op
      .then(() => {
        btn.textContent = this._favMap.has(promptId) ? '★' : '☆';
        btn.style.color = this._favMap.has(promptId) ? FAV_FILLED : FAV_EMPTY;
        btn.title = this._favMap.has(promptId) ? 'Remover dos favoritos' : 'Favoritar';
        if (this._viewMode === 'favorites') this._renderRows();
      })
      .catch((err) => this._showError(err))
      .then(() => {
        this._pending.delete(promptId);
        btn.disabled = false;
      });
  }

  private _addFavorite(promptId: number): Promise<void> {
    const item = this._items.find((x) => x.Id === promptId);
    if (!item) return Promise.reject(new Error(`Prompt ${promptId} não encontrado localmente`));

    const webUrl = this.context.pageContext.web.absoluteUrl;
    const favTitle = encodeURIComponent(this.properties.favoritesListTitle);
    const promptIdField = this.properties.promptIdField || 'PromptID';
    const copyFields = (this.properties.copyFields || '').split(',').map((s) => s.trim()).filter(Boolean);

    const body: Record<string, unknown> = { Title: item.Title };
    for (const f of copyFields) body[f] = (item as unknown as Record<string, unknown>)[f] ?? null;
    body[promptIdField] = promptId;

    const url = `${webUrl}/_api/web/lists/getByTitle('${favTitle}')/items`;
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    }).then((r: SPHttpClientResponse) => {
      if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Adicionar favorito falhou (HTTP ${r.status}): ${t}`)));
      return r.json();
    }).then((data: { Id: number }) => {
      this._favMap.set(promptId, Number(data.Id));
    });
  }

  private _removeFavorite(promptId: number): Promise<void> {
    const favEntryId = this._favMap.get(promptId);
    if (favEntryId === undefined) return Promise.resolve();
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const favTitle = encodeURIComponent(this.properties.favoritesListTitle);
    const url = `${webUrl}/_api/web/lists/getByTitle('${favTitle}')/items(${favEntryId})`;
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
        'odata-version': ''
      }
    }).then((r: SPHttpClientResponse) => {
      if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Remover favorito falhou (HTTP ${r.status}): ${t}`)));
      this._favMap.delete(promptId);
    });
  }

  private _openNewPromptDialog(): void {
    const dialog = new NewPromptDialog(this._choices, this._theme);
    dialog.show()
      .then(() => {
        const data = dialog.result;
        if (!data) return undefined;
        return this._createPrompt(data);
      })
      .catch((err) => this._showError(err));
  }

  private _openDetails(item: IPromptItem): void {
    const initial: IPromptDetailsData = {
      titulo: item.Title || '',
      acao: item.A_x00e7__x00e3_o || '',
      prompt: stripHtml(item.Prompt),
      segmento: item.Categoria || '',
      categoria: item.Categoria0 || '',
      funcionaCom: item.Funcionacom || ''
    };
    const dialog = new PromptDetailsDialog(initial, this._choices, this._theme);
    dialog.show()
      .then(() => {
        const data = dialog.result;
        if (!data) return undefined;
        return this._updatePrompt(item.Id, data);
      })
      .catch((err) => this._showError(err));
  }

  private _updatePrompt(id: number, data: IPromptDetailsData): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.targetListTitle);
    const body: Record<string, unknown> = {
      Title: data.titulo,
      A_x00e7__x00e3_o: data.acao,
      Prompt: data.prompt,
      Categoria: data.segmento,
      Categoria0: data.categoria,
      Funcionacom: data.funcionaCom
    };
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/items(${id})`;
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    }).then((r: SPHttpClientResponse) => {
      if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Atualizar prompt falhou (HTTP ${r.status}): ${t}`)));
      this._showToast('Prompt atualizado com sucesso!');
      return undefined;
    }).then(() => this._loadAll());
  }

  private _createPrompt(data: INewPromptData): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.targetListTitle);
    const body: Record<string, unknown> = {
      Title: data.titulo,
      A_x00e7__x00e3_o: data.acao,
      Prompt: data.prompt,
      Categoria: data.segmento,
      Categoria0: data.categoria,
      Funcionacom: data.funcionaCom
    };
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/items`;
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    }).then((r: SPHttpClientResponse) => {
      if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Criar prompt falhou (HTTP ${r.status}): ${t}`)));
      this._showToast('Prompt criado com sucesso!');
      return undefined;
    }).then(() => {
      return this._loadAll();
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Configuração da biblioteca. Preencha os títulos exatos das listas do site.' },
        groups: [
          {
            groupName: 'Listas',
            groupFields: [
              PropertyPaneTextField('targetListTitle', {
                label: 'Título da lista de prompts',
                description: 'Ex.: Biblioteca de Prompts'
              }),
              PropertyPaneTextField('favoritesListTitle', {
                label: 'Título da lista de favoritos',
                description: 'Ex.: ⭐ Meus Favoritos'
              }),
              PropertyPaneTextField('promptIdField', {
                label: 'Coluna numérica de PromptID (na lista de favoritos)',
                description: 'Nome interno da coluna'
              }),
              PropertyPaneTextField('copyFields', {
                label: 'Campos copiados ao favoritar',
                description: 'Nomes internos separados por vírgula'
              })
            ]
          },
          {
            groupName: 'Cores adicionais',
            groupFields: [
              PropertyPaneTextField('extraToolColors', {
                label: 'Cores por ferramenta',
                description: 'Uma por linha, no formato "Nome=#RRGGBB". Ex.: Google Gemini=#4285F4. Sobrescreve a cor padrão se o nome coincidir com uma já existente.',
                multiline: true,
                rows: 6
              })
            ]
          }
        ]
      }]
    };
  }
}
