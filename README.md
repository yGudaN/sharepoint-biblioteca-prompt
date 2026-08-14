# SharePoint Biblioteca de Prompt

Extensão SPFx (`ListViewCommandSet`) que adiciona uma biblioteca de prompts com favoritos por usuário em listas do SharePoint Online, integrada a um fluxo do Power Automate para criação de novos prompts.

- **Tipo:** SPFx List View Command Set
- **SPFx:** 1.23.2
- **Node:** >=22.14.0 <23.0.0
- **Component ID:** `76f8f770-0a2d-4498-ba55-ba1898198894`

## Recursos

- Botão **"Enviar prompt"** na command bar da lista, que abre um diálogo customizado e dispara um fluxo HTTP do Power Automate com o payload do novo prompt.
- Botão **favorito** (⭐) por linha, com sincronização entre a lista principal e a lista de favoritos do usuário.
- Cópia automática dos campos configurados do prompt original para o item de favorito.

## Estrutura

```
src/extensions/listButton/
  ListButtonCommandSet.ts        # Lógica do command set + favoritos
  ListButtonCommandSet.manifest.json
  NewPromptDialog.ts             # Diálogo de novo prompt
sharepoint/assets/
  elements.xml                   # CustomAction (deploy tenant-wide)
  ClientSideInstance.xml         # Instância pré-configurada da extensão
config/
  package-solution.json          # Metadados do .sppkg
  serve.json                     # Config do gulp serve (debug local)
```

## Pré-requisitos

- Node.js 22.14+ (< 23)
- npm 10+
- Conta no Microsoft 365 com App Catalog do SharePoint disponível
- Fluxo do Power Automate publicado com gatilho **When an HTTP request is received** (retorna a URL com `sig=`)
- Duas listas no site alvo:
  - **Lista principal** (biblioteca de prompts)
  - **Lista de favoritos** com uma coluna numérica que guarda o ID do prompt original

## Instalação

```powershell
git clone <URL_DO_REPO>
cd SharepointBibliotecaPrompt
npm install
```

## Configuração multi-cliente

Todos os valores específicos de cliente estão como **placeholders** (`<...>`) para você substituir antes de compilar. **Nada de segredo é versionado.**

### 1. `sharepoint/assets/elements.xml` e `sharepoint/assets/ClientSideInstance.xml`

Substitua os placeholders dentro do JSON de `ClientSideComponentProperties` / `Properties`:

| Placeholder | Descrição | Exemplo |
|---|---|---|
| `<FLOW_URL>` | URL HTTP do gatilho do Power Automate (inclui `?sig=...`). **Segredo** — não commite. | `https://prod-XX.brazilsouth.logic.azure.com:443/workflows/.../triggers/manual/paths/invoke?...&sig=...` |
| `<TARGET_LIST_TITLE>` | Título exato da lista principal | `Biblioteca de Prompts` |
| `<FAVORITES_LIST_TITLE>` | Título exato da lista de favoritos | `⭐ Meus Favoritos` |
| `<PROMPT_ID_FIELD>` | Nome interno da coluna numérica que guarda o ID do prompt na lista de favoritos | `PromptID` |
| `<COPY_FIELDS>` | Nomes internos das colunas a copiar do prompt para o favorito, separados por vírgula | `Titulo,Prompt,Categoria` |

> **Atenção aos caracteres especiais:** dentro de XML, o `&` na URL do flow precisa estar como `&amp;`.

### 2. `config/serve.json` (apenas debug local com `gulp serve` / `npm start`)

Substitua também `<TENANT>`, `<SITE>`, `<LIST_URL_ENCODED>`, `<FAVORITES_LIST_URL_ENCODED>` e os mesmos placeholders acima.

### 3. `config/package-solution.json`

Se quiser publicar como solução distinta por cliente, altere:
- `solution.name`
- `solution.id` (gere um novo GUID)
- `solution.version`
- `paths.zippedPackage` (nome do `.sppkg`)

## Debug local

```powershell
npm start
```

Isso abre a workbench do SharePoint com a lista definida em `serve.json` e a extensão carregada do `https://localhost:4321`.

## Build de produção (gerar o `.sppkg`)

```powershell
npm run clean
npm run build
```

O pacote sai em `sharepoint/solution/sharepoint-biblioteca-prompt.sppkg` (esse diretório é ignorado pelo Git).

## Deploy no cliente

1. Abrir o **App Catalog** do tenant (`https://<TENANT>-admin.sharepoint.com` → SharePoint → Mais recursos → Catálogo de Aplicativos).
2. Fazer upload do `.sppkg` em **Apps for SharePoint**.
3. Marcar **Make this solution available to all sites in the organization** (para deploy tenant-wide) e clicar em **Deploy**.
4. Aguardar propagação (alguns minutos).
5. Como as propriedades já vêm no `elements.xml` / `ClientSideInstance.xml`, o botão aparece automaticamente nas listas com `RegistrationId="100"` (listas genéricas).

Se preferir deploy por site, remova o `ClientSideInstance.xml` de `features.assets.elementManifests` em `package-solution.json` e adicione o app manualmente pelo **Site Contents → Add an app**.

## Fluxo típico para replicar em outro cliente

1. `git clone <URL_DO_REPO> cliente-X`
2. Preencher os placeholders com os valores do cliente X (flow URL, títulos de lista, etc.)
3. Gerar novo `solution.id` em `package-solution.json` (para não colidir com outras soluções)
4. `npm install` → `npm run build`
5. Publicar o `.sppkg` no App Catalog do cliente X

> Dica: mantenha uma branch por cliente com os XMLs preenchidos, ou script de substituição a partir de um `.env` local.

## Segurança

- **Nunca** commite a URL do flow com `sig=`. O `.gitignore` cobre artefatos de build (`lib/`, `release/`, `temp/`, `*.sppkg`), mas os XMLs em `sharepoint/assets/` **são versionados**, então precisam ficar com placeholders no repositório.
- Se por engano subir um `sig=` real, **regenere o gatilho HTTP** no Power Automate (isso invalida a URL antiga).

## Licença

MIT (ou defina a licença apropriada).
