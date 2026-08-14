import { Log } from '@microsoft/sp-core-library';
import { HttpClient, IHttpClientOptions, HttpClientResponse, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import NewPromptDialog, { INewPromptData } from './NewPromptDialog';

export interface IListButtonCommandSetProperties {
  // URL do gatilho HTTP do Power Automate (novo prompt)
  flowUrl: string;
  // Título exato da lista principal (onde ficam os prompts)
  targetListTitle: string;
  // Título exato da lista de favoritos
  favoritesListTitle: string;
  // Nome interno da coluna Numeric na lista de favoritos que armazena o ID do prompt original
  promptIdField: string;
  // Nomes internos separados por vírgula das colunas a copiar do prompt para o favorito
  copyFields: string;
  // (obsoleto — mantidos por retro-compat com config antiga)
  sampleTextOne?: string;
  sampleTextTwo?: string;
}

const LOG_SOURCE: string = 'ListButtonCommandSet';

const FAV_BTN_CLASS = 'spfx-fav-btn';
const FAV_ID_TITLE_PREFIX = 'spfx-fav-id-';
const FAV_INIT_ATTR = 'data-fav-init';
const FAV_ACTIVE_COLOR = '#FFB900';
const FAV_INACTIVE_COLOR = '#605E5C';

type FavMode = 'off' | 'main' | 'favorites';

export default class ListButtonCommandSet extends BaseListViewCommandSet<IListButtonCommandSetProperties> {
  private _favMap: Map<number, number> = new Map();
  private _pending: Set<number> = new Set();
  private _observer: MutationObserver | undefined;
  private _favMode: FavMode = 'off';

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ListButtonCommandSet');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      compareOneCommand.visible = true;
    }

    this._initFavoritesFeature().catch((err: Error) => {
      console.error('[Favorites] init failed:', err);
      Log.error(LOG_SOURCE, err);
    });

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this._openNewPromptDialog();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _openNewPromptDialog(): void {
    const dialog = new NewPromptDialog();
    dialog.show()
      .then(() => {
        const data = dialog.result;
        if (!data) return undefined;
        const flowUrl: string = this.properties.flowUrl;
        if (!flowUrl) {
          return Dialog.alert('URL do Power Automate não configurada (propriedade "flowUrl").').then(() => undefined);
        }
        return this._sendToFlow(flowUrl, data);
      })
      .catch((err: Error) => {
        Log.error(LOG_SOURCE, err);
        Dialog.alert(`Erro: ${err.message}`).catch(() => { /* noop */ });
      });
  }

  private _sendToFlow(flowUrl: string, data: INewPromptData): Promise<void> {
    const payload = {
      ...data,
      listId: this.context.pageContext.list?.id.toString(),
      listTitle: this.context.pageContext.list?.title,
      webUrl: this.context.pageContext.web.absoluteUrl,
      userEmail: this.context.pageContext.user.email
    };

    const options: IHttpClientOptions = {
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    };

    return this.context.httpClient
      .post(flowUrl, HttpClient.configurations.v1, options)
      .then((response: HttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`HTTP ${response.status} ${response.statusText}`);
        }
        return Dialog.alert('Prompt enviado com sucesso!').then(() => undefined);
      })
      .catch((err: Error) => {
        Log.error(LOG_SOURCE, err);
        return Dialog.alert(`Erro ao enviar ao fluxo: ${err.message}`).then(() => undefined);
      });
  }

  private _initFavoritesFeature(): Promise<void> {
    const currentTitle = this.context.pageContext.list?.title;
    const targetTitle = this.properties.targetListTitle;
    const favoritesTitle = this.properties.favoritesListTitle;
    console.log('[Favorites] currentTitle:', currentTitle, '| target:', targetTitle, '| favorites:', favoritesTitle);

    if (!targetTitle || !favoritesTitle) {
      console.warn('[Favorites] Skipped — targetListTitle or favoritesListTitle not configured');
      return Promise.resolve();
    }

    if (currentTitle === targetTitle) {
      this._favMode = 'main';
    } else if (currentTitle === favoritesTitle) {
      this._favMode = 'favorites';
    } else {
      console.log('[Favorites] Skipped — list not supported');
      return Promise.resolve();
    }

    console.log('[Favorites] mode:', this._favMode);

    const loadPromise = this._favMode === 'main' ? this._loadFavorites() : Promise.resolve();
    return loadPromise.then(() => {
      this._startObserver();
      this._scanAndDecorate(document.body);
    });
  }

  private _loadFavorites(): Promise<void> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const favTitle = encodeURIComponent(this.properties.favoritesListTitle);
    const promptIdField = this.properties.promptIdField;
    const userId = (this.context.pageContext as unknown as { legacyPageContext?: { userId?: number } }).legacyPageContext?.userId;
    const filter = userId ? `AuthorId eq ${userId}` : `Author/EMail eq '${this.context.pageContext.user.email}'`;
    const url = `${webUrl}/_api/web/lists/getByTitle('${favTitle}')/items?$filter=${encodeURIComponent(filter)}&$select=Id,${promptIdField}&$top=5000`;
    console.log('[Favorites] query URL:', url);

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((txt) => {
            console.error('[Favorites] error body:', txt);
            throw new Error(`Load favorites HTTP ${response.status}`);
          });
        }
        return response.json();
      })
      .then((data: { value: Array<Record<string, unknown>> } | undefined) => {
        if (!data) return;
        this._favMap.clear();
        for (const item of data.value) {
          const promptId = Number(item[promptIdField]);
          const favId = Number(item.Id);
          if (!isNaN(promptId) && !isNaN(favId)) {
            this._favMap.set(promptId, favId);
          }
        }
        console.log(`[Favorites] Loaded ${this._favMap.size} favorites`);
      });
  }

  private _startObserver(): void {
    this._observer = new MutationObserver((mutations) => {
      for (const m of mutations) {
        m.addedNodes.forEach((node) => {
          if (node.nodeType === Node.ELEMENT_NODE) {
            this._scanAndDecorate(node as HTMLElement);
          }
        });
      }
    });
    this._observer.observe(document.body, { childList: true, subtree: true });
  }

  private _scanAndDecorate(root: HTMLElement): void {
    if (this._favMode === 'off') return;
    if (root.querySelectorAll) {
      root.querySelectorAll(`.${FAV_BTN_CLASS}:not([${FAV_INIT_ATTR}])`).forEach((btn) => this._decorate(btn as HTMLElement));
    }
    if (root.classList?.contains(FAV_BTN_CLASS) && !root.hasAttribute(FAV_INIT_ATTR)) {
      this._decorate(root);
    }
  }

  private _decorate(btn: HTMLElement): void {
    btn.setAttribute(FAV_INIT_ATTR, '1');
    const id = this._extractPromptId(btn);
    if (id === undefined) return;

    const isActive = this._favMode === 'favorites' ? true : this._favMap.has(id);
    this._applyState(btn, isActive);

    btn.addEventListener('click', (ev: MouseEvent) => {
      ev.preventDefault();
      ev.stopPropagation();
      if (this._favMode === 'favorites') {
        this._deleteFromFavoritesList(btn, id).catch((err: Error) => console.error('[Favorites]', err));
      } else {
        this._toggle(btn, id).catch((err: Error) => console.error('[Favorites]', err));
      }
    }, true);
  }

  private _extractPromptId(btn: HTMLElement): number | undefined {
    const title = btn.getAttribute('title') || '';
    const idx = title.indexOf(FAV_ID_TITLE_PREFIX);
    if (idx < 0) return undefined;
    const rest = title.substring(idx + FAV_ID_TITLE_PREFIX.length);
    const n = parseInt(rest, 10);
    return isNaN(n) ? undefined : n;
  }

  private _applyState(btn: HTMLElement, active: boolean): void {
    btn.style.color = active ? FAV_ACTIVE_COLOR : FAV_INACTIVE_COLOR;
    btn.setAttribute('title', active ? 'Remover dos favoritos' : 'Favoritar este prompt');
  }

  private _toggle(btn: HTMLElement, promptId: number): Promise<void> {
    if (this._pending.has(promptId)) return Promise.resolve();
    this._pending.add(promptId);

    const wasActive = this._favMap.has(promptId);
    this._applyState(btn, !wasActive);

    const op = wasActive
      ? this._removeFavorite(promptId).then(() => { this._favMap.delete(promptId); })
      : this._addFavorite(promptId).then((newId) => { this._favMap.set(promptId, newId); });

    return op
      .catch((err: Error) => {
        this._applyState(btn, wasActive);
        Log.error(LOG_SOURCE, err);
      })
      .then(() => { this._pending.delete(promptId); });
  }

  private _addFavorite(promptId: number): Promise<number> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const sourceTitle = encodeURIComponent(this.properties.targetListTitle);
    const favTitle = encodeURIComponent(this.properties.favoritesListTitle);
    const promptIdField = this.properties.promptIdField;
    const copyFields = (this.properties.copyFields || '').split(',').map(s => s.trim()).filter(Boolean);
    const selectFields = ['Title', ...copyFields].join(',');
    const sourceUrl = `${webUrl}/_api/web/lists/getByTitle('${sourceTitle}')/items(${promptId})?$select=${selectFields}`;

    return this.context.spHttpClient.get(sourceUrl, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) {
          return r.text().then((txt) => {
            console.error('[Favorites] fetch source error body:', txt);
            throw new Error(`Fetch source HTTP ${r.status}`);
          });
        }
        return r.json();
      })
      .then((src: Record<string, unknown> | undefined) => {
        const body: Record<string, unknown> = {
          Title: (src?.Title as string) || `Favorito ${promptId}`
        };
        for (const f of copyFields) {
          body[f] = src?.[f] ?? null;
        }
        body[promptIdField] = promptId;
        console.log('[Favorites] creating favorite with body:', body);
        const url = `${webUrl}/_api/web/lists/getByTitle('${favTitle}')/items`;
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': 'application/json;odata=nometadata',
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify(body)
        });
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          return response.text().then((txt) => {
            console.error('[Favorites] add favorite error body:', txt);
            throw new Error(`Add favorite HTTP ${response.status}`);
          }) as unknown as Promise<{ Id: number }>;
        }
        return response.json();
      })
      .then((data: { Id: number }) => Number(data.Id));
  }

  private _removeFavorite(promptId: number): Promise<void> {
    const favEntryId = this._favMap.get(promptId);
    if (favEntryId === undefined) return Promise.resolve();
    return this._deleteFavoriteEntry(favEntryId);
  }

  private _deleteFavoriteEntry(favEntryId: number): Promise<void> {
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
    }).then((response: SPHttpClientResponse) => {
      if (!response.ok) throw new Error(`Remove favorite HTTP ${response.status}`);
    });
  }

  private _deleteFromFavoritesList(btn: HTMLElement, favEntryId: number): Promise<void> {
    if (this._pending.has(favEntryId)) return Promise.resolve();
    this._pending.add(favEntryId);

    const card = btn.closest<HTMLElement>('[role="listitem"], [data-list-index], [data-selection-index]');
    if (card) card.style.opacity = '0.4';
    this._applyState(btn, false);

    return this._deleteFavoriteEntry(favEntryId)
      .then(() => {
        if (card) {
          card.style.transition = 'opacity 0.2s ease';
          card.style.opacity = '0';
          setTimeout(() => { card.style.display = 'none'; }, 200);
        }
      })
      .catch((err: Error) => {
        if (card) card.style.opacity = '1';
        this._applyState(btn, true);
        console.error('[Favorites] delete failed:', err);
      })
      .then(() => { this._pending.delete(favEntryId); });
  }
}


