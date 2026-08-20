import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import PromptDetailsDialog, { IPromptDetailsData } from '../bibliotecaPrompt/PromptDetailsDialog';
import { IPromptChoices } from '../bibliotecaPrompt/promptChoices';

export interface IAnalyticsWebPartProps {
  targetListTitle: string;
  favoritesListTitle: string;
  promptIdField: string;
  topN: number;
}

interface IPromptItem {
  Id: number;
  Title: string;
  Created: string;
  Prompt?: string;
  A_x00e7__x00e3_o?: string;
  Categoria?: string;
  Categoria0?: string;
  Funcionacom?: string;
  AuthorId?: number;
  Author?: { Title: string; EMail: string };
  Ativo?: boolean;
}

interface IFavItem {
  Id: number;
  Created: string;
  [k: string]: unknown;
}

type Period = 'all' | '30d' | '90d';

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

function daysAgoIso(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString();
}

export default class AnalyticsWebPart extends BaseClientSideWebPart<IAnalyticsWebPartProps> {
  private _items: IPromptItem[] = [];
  private _favs: IFavItem[] = [];
  private _period: Period = 'all';
  private _loaded: boolean = false;
  private _theme: 'light' | 'dark' = 'light';

  public render(): void {
    this.domElement.innerHTML = this._shellHtml();
    this._wire();
    if (!this._loaded) {
      this._loadAll().catch((err) => this._showError(err));
    } else {
      this._renderAll();
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

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  private _shellHtml(): string {
    return `
      <style>
        .an-wrap {
          --an-bg: transparent;
          --an-surface: #ffffff;
          --an-surface-alt: #faf9f8;
          --an-hover: #f3f2f1;
          --an-text: #323130;
          --an-text-2: #605e5c;
          --an-text-3: #a19f9d;
          --an-border: #edebe9;
          --an-border-2: #f3f2f1;
          --an-primary: #0078d4;
          --an-danger: #a4262c;
          --an-shadow: rgba(0, 0, 0, 0.08);

          font-family: 'Segoe UI', sans-serif;
          color-scheme: light;
          color: var(--an-text);
          background: var(--an-bg);
          padding: 12px;
          border-radius: 8px;
          min-height: 100vh;
          box-sizing: border-box;
          transition: background 0.2s ease, color 0.2s ease;
        }
        .an-wrap.an-dark {
          color-scheme: dark;
          --an-bg: #1f1f1f;
          --an-surface: #2b2b2b;
          --an-surface-alt: #262626;
          --an-hover: #3b3a39;
          --an-text: #f3f2f1;
          --an-text-2: #c8c6c4;
          --an-text-3: #8a8886;
          --an-border: #3b3a39;
          --an-border-2: #333333;
          --an-primary: #4cb2ff;
          --an-danger: #f1707b;
          --an-shadow: rgba(0, 0, 0, 0.4);
        }
        .an-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 8px; }
        .an-title { font-size: 20px; font-weight: 600; margin: 0; color: var(--an-text); }
        .an-head-right { display: flex; align-items: center; gap: 8px; }
        .an-theme-toggle {
          background: transparent; border: 1px solid var(--an-border);
          border-radius: 999px; width: 36px; height: 36px;
          cursor: pointer; font-size: 16px; line-height: 1;
          display: inline-flex; align-items: center; justify-content: center;
          color: var(--an-text);
        }
        .an-theme-toggle:hover { background: var(--an-hover); }
        .an-period { display: flex; gap: 4px; background: var(--an-hover); border-radius: 999px; padding: 4px; }
        .an-period button { border: none; background: transparent; padding: 6px 14px; border-radius: 999px; cursor: pointer; font-family: inherit; font-size: 13px; color: var(--an-text-2); }
        .an-period button.active { background: var(--an-surface); color: var(--an-primary); font-weight: 600; box-shadow: 0 1px 3px var(--an-shadow); }
        .an-status { padding: 8px 12px; font-size: 13px; color: var(--an-text-2); }
        .an-status.error { color: var(--an-danger); }

        .an-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .an-kpi { background: var(--an-surface); border: 1px solid var(--an-border); border-radius: 12px; padding: 20px; }
        .an-kpi-value { font-size: 32px; font-weight: 700; color: var(--an-text); line-height: 1; margin-bottom: 4px; }
        .an-kpi-label { font-size: 13px; color: var(--an-text-2); font-weight: 500; }

        .an-cols { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 20px; }
        .an-panel { background: var(--an-surface); border: 1px solid var(--an-border); border-radius: 12px; padding: 20px; }
        .an-panel-title { font-size: 15px; font-weight: 600; margin: 0 0 16px; color: var(--an-text); }

        .an-rank { list-style: none; padding: 0; margin: 0; }
        .an-rank li { display: flex; align-items: center; gap: 12px; padding: 8px 0; border-bottom: 1px solid var(--an-border-2); }
        .an-rank li:last-child { border-bottom: none; }
        .an-rank li.an-rank-clickable { cursor: pointer; padding: 8px; border-radius: 6px; margin: 0 -8px; border-bottom: none; }
        .an-rank li.an-rank-clickable:hover { background: var(--an-hover); }
        .an-rank li.an-rank-clickable:focus { outline: 2px solid var(--an-primary); outline-offset: 2px; }
        .an-rank-pos { flex-shrink: 0; width: 28px; height: 28px; border-radius: 50%; background: var(--an-hover); color: var(--an-text-2); display: inline-flex; align-items: center; justify-content: center; font-size: 12px; font-weight: 700; }
        .an-rank li:nth-child(1) .an-rank-pos { background: #FFB900; color: #323130; }
        .an-rank li:nth-child(2) .an-rank-pos { background: #c8c8c8; color: #323130; }
        .an-rank li:nth-child(3) .an-rank-pos { background: #cd7f32; color: #fff; }
        .an-rank-body { flex: 1; min-width: 0; }
        .an-rank-name { font-size: 14px; font-weight: 500; color: var(--an-text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .an-rank-bar { position: relative; height: 6px; background: var(--an-hover); border-radius: 999px; margin-top: 6px; overflow: hidden; }
        .an-rank-bar span { position: absolute; left: 0; top: 0; bottom: 0; background: var(--an-primary); border-radius: 999px; }
        .an-rank-count { flex-shrink: 0; font-size: 13px; color: var(--an-text-2); font-weight: 600; min-width: 36px; text-align: right; }

        .an-empty { text-align: center; color: var(--an-text-2); padding: 20px; font-size: 13px; }
      </style>
      <div class="an-wrap${this._theme === 'dark' ? ' an-dark' : ''}">
        <div class="an-head">
          <h2 class="an-title">📊 Dashboard — Biblioteca de Prompts</h2>
          <div class="an-head-right">
            <button type="button" class="an-theme-toggle" data-role="theme-toggle" title="Alternar tema">${this._theme === 'dark' ? '☀️' : '🌙'}</button>
            <div class="an-period" data-role="period">
              <button type="button" data-value="all" class="active">Todos</button>
              <button type="button" data-value="30d">Últimos 30 dias</button>
              <button type="button" data-value="90d">Últimos 90 dias</button>
            </div>
          </div>
        </div>
        <div class="an-status" data-role="status">Carregando...</div>
        <div data-role="kpis"></div>
        <div data-role="rankings"></div>
      </div>
    `;
  }

  private _wire(): void {
    const buttons = this.domElement.querySelectorAll<HTMLButtonElement>('[data-role="period"] button');
    buttons.forEach((b) => {
      b.addEventListener('click', () => {
        this._period = (b.getAttribute('data-value') as Period) || 'all';
        buttons.forEach((x) => x.classList.toggle('active', x === b));
        this._renderAll();
      });
    });
    const themeBtn = this.domElement.querySelector('[data-role="theme-toggle"]') as HTMLButtonElement | null;
    if (themeBtn) themeBtn.addEventListener('click', () => this._toggleTheme());
  }

  private _setStatus(text: string, isError: boolean = false): void {
    const el = this.domElement.querySelector('[data-role="status"]') as HTMLElement | null;
    if (!el) return;
    el.textContent = text;
    el.classList.toggle('error', isError);
    el.style.display = text ? 'block' : 'none';
  }

  private _showError(err: unknown): void {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[Analytics]', err);
    this._setStatus(msg, true);
  }

  private _loadAll(): Promise<void> {
    if (!this.properties.targetListTitle || !this.properties.favoritesListTitle) {
      this._setStatus('Configure os títulos das listas no painel de propriedades da web part.', true);
      return Promise.resolve();
    }
    this._setStatus('Carregando...');
    return Promise.all([this._loadItems(), this._loadFavs()])
      .then(() => {
        this._loaded = true;
        this._setStatus('');
        this._renderAll();
      })
      .catch((err) => this._showError(err));
  }

  private _loadItems(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.targetListTitle);
    const select = ['Id', 'Title', 'Created', 'AuthorId', 'Prompt', 'A_x00e7__x00e3_o', 'Categoria', 'Categoria0', 'Funcionacom', 'Ativo', 'Author/Title', 'Author/EMail'].join(',');
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/items?$select=${select}&$expand=Author&$top=5000`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Falha ao carregar prompts (HTTP ${r.status}): ${t}`)));
        return r.json();
      })
      .then((data: { value: IPromptItem[] }) => { this._items = data.value || []; });
  }

  private _loadFavs(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const list = encodeURIComponent(this.properties.favoritesListTitle);
    const promptIdField = this.properties.promptIdField || 'PromptID';
    const url = `${webUrl}/_api/web/lists/getByTitle('${list}')/items?$select=Id,Created,${promptIdField}&$top=5000`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) return r.text().then((t) => Promise.reject(new Error(`Falha ao carregar favoritos (HTTP ${r.status}): ${t}`)));
        return r.json();
      })
      .then((data: { value: IFavItem[] }) => { this._favs = data.value || []; });
  }

  private _filterByPeriod<T extends { Created: string }>(arr: T[]): T[] {
    if (this._period === 'all') return arr;
    const days = this._period === '30d' ? 30 : 90;
    const cutoff = daysAgoIso(days);
    return arr.filter((x) => x.Created >= cutoff);
  }

  private _renderAll(): void {
    const items = this._filterByPeriod(this._items);
    const favs = this._filterByPeriod(this._favs);
    const topN = this.properties.topN || 10;

    this._renderKpis(items, favs);
    this._renderRankings(items, favs, topN);
  }

  private _renderKpis(items: IPromptItem[], favs: IFavItem[]): void {
    const container = this.domElement.querySelector('[data-role="kpis"]') as HTMLElement | null;
    if (!container) return;
    const authors = new Set<string>();
    for (const it of items) {
      const name = it.Author?.Title || (it.AuthorId ? `#${it.AuthorId}` : '');
      if (name) authors.add(name);
    }
    container.innerHTML = `
      <div class="an-kpis">
        <div class="an-kpi"><div class="an-kpi-value">${items.length}</div><div class="an-kpi-label">Prompts publicados</div></div>
        <div class="an-kpi"><div class="an-kpi-value">${favs.length}</div><div class="an-kpi-label">Favoritamentos</div></div>
        <div class="an-kpi"><div class="an-kpi-value">${authors.size}</div><div class="an-kpi-label">Autores distintos</div></div>
      </div>
    `;
  }

  private _renderRankings(items: IPromptItem[], favs: IFavItem[], topN: number): void {
    const container = this.domElement.querySelector('[data-role="rankings"]') as HTMLElement | null;
    if (!container) return;

    // Top contribuidores
    const byAuthor: { [name: string]: number } = {};
    for (const it of items) {
      const name = it.Author?.Title || (it.AuthorId ? `Usuário #${it.AuthorId}` : 'Anônimo');
      byAuthor[name] = (byAuthor[name] || 0) + 1;
    }
    const contribs = Object.keys(byAuthor).map((name) => ({ name, count: byAuthor[name] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, topN);
    const maxContrib = contribs.length > 0 ? contribs[0].count : 0;

    // Prompts mais favoritados
    const promptIdField = this.properties.promptIdField || 'PromptID';
    const byPrompt: { [id: number]: number } = {};
    for (const f of favs) {
      const pid = Number(f[promptIdField]);
      if (!isNaN(pid)) byPrompt[pid] = (byPrompt[pid] || 0) + 1;
    }
    const promptById: { [id: number]: IPromptItem } = {};
    for (const it of this._items) promptById[it.Id] = it;
    const popular = Object.keys(byPrompt).map((k) => {
      const id = Number(k);
      const it = promptById[id];
      return { id, title: it ? it.Title : `Prompt #${id} (excluído)`, count: byPrompt[id], ativo: !it || it.Ativo !== false };
    }).filter((p) => p.ativo).sort((a, b) => b.count - a.count).slice(0, topN);
    const maxPop = popular.length > 0 ? popular[0].count : 0;

    container.innerHTML = `
      <div class="an-cols">
        <div class="an-panel">
          <h3 class="an-panel-title">👥 Top contribuidores</h3>
          ${contribs.length === 0 ? '<div class="an-empty">Sem dados no período.</div>' : `
            <ol class="an-rank">
              ${contribs.map((c, i) => `
                <li>
                  <span class="an-rank-pos">${i + 1}</span>
                  <div class="an-rank-body">
                    <div class="an-rank-name" title="${esc(c.name)}">${esc(c.name)}</div>
                    <div class="an-rank-bar"><span style="width:${maxContrib ? (c.count / maxContrib * 100) : 0}%"></span></div>
                  </div>
                  <span class="an-rank-count">${c.count}</span>
                </li>
              `).join('')}
            </ol>
          `}
        </div>
        <div class="an-panel">
          <h3 class="an-panel-title">⭐ Prompts mais favoritados</h3>
          ${popular.length === 0 ? '<div class="an-empty">Ainda ninguém favoritou prompts no período.</div>' : `
            <ol class="an-rank">
              ${popular.map((p, i) => `
                <li class="${promptById[p.id] ? 'an-rank-clickable' : ''}" ${promptById[p.id] ? `data-prompt-id="${p.id}" tabindex="0" role="button" title="Clique para ver detalhes"` : ''}>
                  <span class="an-rank-pos">${i + 1}</span>
                  <div class="an-rank-body">
                    <div class="an-rank-name" title="${esc(p.title)}">${esc(p.title)}</div>
                    <div class="an-rank-bar"><span style="width:${maxPop ? (p.count / maxPop * 100) : 0}%"></span></div>
                  </div>
                  <span class="an-rank-count">${p.count}</span>
                </li>
              `).join('')}
            </ol>
          `}
        </div>
      </div>
    `;

    container.querySelectorAll<HTMLElement>('[data-prompt-id]').forEach((el) => {
      const openIt = (): void => {
        const id = Number(el.getAttribute('data-prompt-id'));
        const item = this._items.find((x) => x.Id === id);
        if (item) this._openDetails(item);
      };
      el.addEventListener('click', openIt);
      el.addEventListener('keydown', (ev: KeyboardEvent) => {
        if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); openIt(); }
      });
    });
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
    const choices: IPromptChoices = {
      acoes: initial.acao ? [initial.acao] : [],
      segmentos: initial.segmento ? [initial.segmento] : [],
      categorias: initial.categoria ? [initial.categoria] : [],
      funcionaCom: initial.funcionaCom ? [initial.funcionaCom] : []
    };
    const dialog = new PromptDetailsDialog(initial, choices, this._theme, true);
    dialog.show().catch((err) => this._showError(err));
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'Configuração do dashboard. Deve apontar para as mesmas listas da web part principal.' },
        groups: [{
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
              label: 'Coluna PromptID (favoritos)',
              description: 'Nome interno da coluna'
            }),
            PropertyPaneSlider('topN', {
              label: 'Top N por ranking',
              min: 3,
              max: 25,
              step: 1
            })
          ]
        }]
      }]
    };
  }
}
