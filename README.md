# 📚 SharePoint Biblioteca de Prompts

Solução completa para gerenciar uma biblioteca de prompts de IA dentro do SharePoint e Teams, com favoritos por usuário, dashboard de contribuições e dark mode automático.

Feita como duas Web Parts SPFx que rodam em qualquer site SharePoint moderno.

---

## 📋 Índice

1. [O que essa solução faz](#-o-que-essa-solução-faz)
2. [Pré-requisitos (o que você precisa ter instalado)](#-pré-requisitos)
3. [Como baixar o projeto](#-como-baixar-o-projeto)
4. [Como abrir no VS Code](#-como-abrir-no-vs-code)
5. [Compilar o pacote (.sppkg)](#-compilar-o-pacote-sppkg)
6. [Implantar em um cliente novo (passo a passo)](#-implantar-em-um-cliente-novo-passo-a-passo)
7. [Personalização e manutenção](#-personalização-e-manutenção)
8. [Adicionar como aba no Teams](#-adicionar-como-aba-no-teams)
9. [Estrutura do projeto](#-estrutura-do-projeto)
10. [Atualizar a solução (depois de mudar código)](#-atualizar-a-solução)
11. [Solução de problemas comuns](#-solução-de-problemas-comuns)
12. [Segurança e Git](#-segurança-e-git)

---

## 🎯 O que essa solução faz

Duas Web Parts, cada uma numa página:

**1. Biblioteca de Prompts** — a tela principal
- Grid de cards com todos os prompts publicados
- Busca por texto (título, prompt, categoria)
- Filtros por **App**, **Área**, **Função** (agrupa em seções) e **⭐ Favoritos**
- Botão ⭐ para favoritar (cada usuário tem seus próprios favoritos)
- Botão 📋 para copiar o texto do prompt
- Botão **+ Novo prompt** com diálogo pra criar
- Clique no card abre modal com edição
- 🌙/☀️ toggle de dark mode (detecta tema do Teams automaticamente)

**2. Dashboard — Biblioteca de Prompts** — página opcional de métricas
- KPIs: total de prompts, favoritamentos, autores distintos
- Ranking dos **contribuidores** (quem publicou mais prompts)
- Ranking dos **prompts mais favoritados**
- Filtro por período: todos / 30 dias / 90 dias

Ambas rodam em **SharePoint (browser)**, **abas do Teams** e são **totalmente customizáveis** pelo admin do cliente sem mexer em código.

---

## ✅ Pré-requisitos

Você precisa ter instalado na sua máquina:

### 1. **Git** (para baixar o projeto)
Download: https://git-scm.com/downloads

Depois de instalar, abra o **PowerShell** e digite:
```powershell
git --version
```
Se aparecer algo como `git version 2.x.x`, tá certo.

### 2. **Node.js 22.14 ou superior** (para compilar)
Download: https://nodejs.org/en/download

⚠️ **Atenção:** precisa ser a versão **22.x**. Não use versões mais antigas (18, 20) nem versões maiores (23, 24).

Verifique:
```powershell
node --version
```
Deve mostrar `v22.14.0` ou superior (mas menor que 23).

### 3. **PowerShell 7 ou superior** (para os scripts de configuração do SharePoint)
Download: https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows

⚠️ **Importante:** o Windows já vem com PowerShell 5.1, mas os scripts precisam da versão 7+. **Não use o "Windows PowerShell" azul** — use o **"PowerShell"** preto que você instala pelo link acima.

Verifique:
```powershell
$PSVersionTable.PSVersion
```
Deve mostrar `Major: 7` (ou superior).

### 4. **Módulo PnP.PowerShell** (para os scripts)
Abra o **PowerShell 7** e rode uma vez só:
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser -Force
```
Vai levar 1-2 minutos. Depois de instalado, nunca mais precisa.

### 5. **VS Code** (para editar o código, opcional se você só quer instalar)
Download: https://code.visualstudio.com/

### 6. **Conta Microsoft 365** com permissões apropriadas:
- **Admin do tenant** → precisa pra rodar o registro inicial do PnP no Entra ID
- **App Catalog admin** → precisa pra subir o `.sppkg`
- **Owner do site** → precisa pra criar as listas e páginas

---

## 📥 Como baixar o projeto

1. Abra o **PowerShell 7**.
2. Navegue até a pasta onde você quer salvar o projeto:
   ```powershell
   cd C:\Users\SEU_USUARIO\Documents
   ```
3. Baixe o projeto:
   ```powershell
   git clone https://github.com/yGudaN/sharepoint-biblioteca-prompt.git
   ```
4. Entre na pasta baixada:
   ```powershell
   cd sharepoint-biblioteca-prompt
   ```

Pronto — você tem o projeto na sua máquina.

---

## 🎨 Como abrir no VS Code

**Opção 1 — pela linha de comando (mais rápido):**

Dentro da pasta do projeto no PowerShell:
```powershell
code .
```
(o ponto no final é importante — significa "abre a pasta atual")

**Opção 2 — pela interface:**

Abre o VS Code → **File → Open Folder** → navega até a pasta `sharepoint-biblioteca-prompt` → **Selecionar pasta**.

Se aparecer um aviso "Do you trust the authors of the files in this folder?" → clica em **Yes, I trust the authors**.

---

## 📦 Compilar o pacote (.sppkg)

**Só precisa fazer isso 1 vez** (ou toda vez que mudar código).

No terminal do VS Code (menu **Terminal → New Terminal**) ou PowerShell dentro da pasta do projeto:

### 1. Instalar dependências (só na primeira vez)
```powershell
npm install
```
Vai baixar todas as bibliotecas necessárias. Pode levar 5–10 minutos. Ao final você verá vários avisos de "deprecated" — pode ignorar.

### 2. Compilar
```powershell
npm run build
```
Ao final aparece `-------------------- Finished --------------------` e o arquivo compilado fica em:

```
sharepoint/solution/sharepoint-biblioteca-prompt.sppkg
```

Esse é o **pacote** que você vai subir no SharePoint do cliente.

---

## 🚀 Implantar em um cliente novo (passo a passo)

Faça na ordem. Cada passo é independente do outro.

### Passo 1 — Registrar o App do PnP (só 1 vez por tenant)

O PnP PowerShell precisa de um "App Registration" no Entra ID do tenant do cliente pra funcionar. Faz uma vez só por tenant — depois nunca mais.

No PowerShell 7, rode (troca o nome do tenant):

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell - Biblioteca de Prompts" -Tenant "SEU_TENANT.onmicrosoft.com" -SharePointDelegatePermissions "AllSites.FullControl" -GraphDelegatePermissions "User.Read"
```

> 💡 **Por que passar permissões explícitas?** Sem elas, o cmdlet registra várias permissões do Graph que a gente não precisa (Directory.Read, Group.Read etc.). Passando `AllSites.FullControl` (SharePoint) e `User.Read` (Graph), o admin do cliente aprova só o **mínimo necessário** — mais rápido de aprovar e mais seguro.

> 💡 **Onde acho o "SEU_TENANT"?** É o prefixo do endereço do SharePoint. Se o site é `https://empresax.sharepoint.com/...`, o tenant é `empresax`.

Vai:
1. Abrir uma janela de login → faça login com conta que tenha permissão de criar app registrations no tenant (normalmente conta admin)
2. Mostrar uma tela pedindo permissões (SharePoint + Graph) → **Aceite**
3. Voltar pro terminal com um `ClientId` no output. Tipo:
   ```
   ClientId: 12345678-1234-1234-1234-123456789012
   ```

**COPIE ESSE `ClientId`.** Você vai usar em todos os próximos scripts.

Guarde ele numa variável de ambiente pra não precisar digitar toda hora:
```powershell
[Environment]::SetEnvironmentVariable("PNP_CLIENT_ID", "12345678-1234-1234-1234-123456789012", "User")
```
(depois feche e reabra o PowerShell)

#### ⚠️ Se você NÃO é admin do tenant

Se você é apenas usuário do tenant (com acesso a criar sites/listas, mas sem ser Global Admin), duas coisas podem acontecer:

**1. O comando cria a app mas não consegue aprovar as permissões**

Você vai receber o ClientId, mas quando tentar rodar os scripts vai dar erro de permissão. Nesse caso:
- Vá no Portal Entra do tenant: https://entra.microsoft.com → **Applications** → **App registrations** → encontre sua app.
- Aba **API permissions** → clique em **"Request admin consent"** (ou algum botão similar). Isso envia uma solicitação pro admin.
- Ou peça diretamente ao admin: "Preciso que você aprove essa app: `PnP PowerShell - Biblioteca de Prompts`. As permissões são `AllSites.FullControl` (SharePoint) e `User.Read` (Graph)."
- Assim que ele clicar em **"Grant admin consent"**, a app tá liberada. Você não precisa fazer nada do seu lado — mesmo ClientId, mesma app, funciona.

**2. Ainda de olho: `Set-PnPSite -SocialBarOnSitePagesDisabled $true`**

Esse comando específico precisa que **sua conta** seja **Site Collection Administrator** do site (não Global Admin do tenant, só do site em si). Peça pro admin do site te adicionar em:
- Site → Engrenagem ⚙️ → **Site permissions** → **Site collection administrators** → adicionar seu usuário.

Com esses dois liberados (admin consent + site collection admin), você roda todos os scripts do repositório autonomamente.

### Passo 2 — Criar as listas no site do cliente

O site já precisa existir no SharePoint. Se não existir, crie primeiro em https://SEU_TENANT.sharepoint.com → **+ Criar site**.

Com o site criado, rode:

```powershell
.\scripts\Setup-BibliotecaPrompts.ps1 -SiteUrl "https://SEU_TENANT.sharepoint.com/sites/NOME_DO_SITE" -ClientId "SEU_CLIENT_ID"
```

Exemplo real:
```powershell
.\scripts\Setup-BibliotecaPrompts.ps1 -SiteUrl "https://empresax.sharepoint.com/sites/BibliotecaPrompts" -ClientId "12345678-1234-1234-1234-123456789012"
```

O script cria:
- 📋 Lista **Biblioteca de Prompts** — onde ficam os prompts
- ⭐ Lista **Meus Favoritos** — favoritos por usuário
- Todas as colunas com os nomes internos que a web part espera
- Todas as opções padrão nos campos Choice (Segmento, Ação, Categoria, Funciona com)

Você pode rodar 2x sem problema — ele não duplica nada (é **idempotente**).

### Passo 3 — Fazer upload do pacote no App Catalog

1. Acesse `https://SEU_TENANT-admin.sharepoint.com` → menu esquerdo **More features** → **Apps** → clica em **Open** → **App Catalog**.
   
   > Se aparecer uma tela pedindo pra criar o App Catalog (primeira vez do tenant), aceite. Leva ~2min pra propagar.
2. Dentro do App Catalog, no menu esquerdo clica **Apps for SharePoint**.
3. Clica em **Upload** (ou arrasta o arquivo).
4. Seleciona o arquivo:
   ```
   sharepoint/solution/sharepoint-biblioteca-prompt.sppkg
   ```
5. Vai abrir um diálogo **"Do you trust..."**:
   - ❌ **NÃO marque** "Make this solution available to all sites" — queremos deploy só nesse site
   - Clica em **Deploy**
6. Confira que aparece **Deployed = Yes**, **Enabled = Yes**, **App Package Error Message = (vazio)**

### Passo 4 — Instalar o app no site do cliente

1. Vá em `https://SEU_TENANT.sharepoint.com/sites/NOME_DO_SITE`
2. Clica na engrenagem ⚙️ (canto superior direito) → **Add an app**
3. Procure por **sharepoint-biblioteca-prompt-webpart** → clica em **Add**
4. Aguarde alguns segundos. O app aparece em Site Contents.

> ⚠️ Se der erro "Couldn't add this app", espera 10–15min (propagação) e tenta de novo.

### Passo 5 — Desligar a barra social do site (uma vez por site)

Isso esconde os botões de "curtir", contador de visualizações e comentários em todas as páginas do site.

```powershell
Connect-PnPOnline -Url "https://SEU_TENANT.sharepoint.com/sites/NOME_DO_SITE" -Interactive -ClientId "SEU_CLIENT_ID"

Set-PnPSite -SocialBarOnSitePagesDisabled $true

Disconnect-PnPOnline
```

### Passo 6 — Criar a página da Biblioteca

1. No site, canto superior esquerdo: **+ Novo → Página**
2. Escolhe layout **Em branco** → **Criar página**
3. Nome: `Biblioteca de Prompts` (ou o que quiser)
4. Clica no **+** no meio da página → busca por **Biblioteca**
5. Clica em **Biblioteca de Prompts** (a nossa web part)
6. No painel de propriedades à direita, confira os valores padrão:
   - **Título da lista de prompts**: `Biblioteca de Prompts`
   - **Título da lista de favoritos**: `⭐ Meus Favoritos`
   - **Coluna PromptID**: `PromptID`
   - **Campos copiados**: `A_x00e7__x00e3_o,Prompt,Categoria,Categoria0,Funcionacom`
7. Feche o painel.
8. **Publicar** (botão azul canto superior direito).

### Passo 7 — Aplicar tela cheia na página (sem header, sem command bar)

Rode:

```powershell
.\scripts\Configure-Page.ps1 -SiteUrl "https://SEU_TENANT.sharepoint.com/sites/NOME_DO_SITE" -PageName "Biblioteca-de-Prompts.aspx" -ClientId "SEU_CLIENT_ID"
```

> 💡 **Como sei o nome do arquivo `.aspx`?** Abra a página no browser e veja o final da URL: `.../SitePages/NOME-AQUI.aspx`. Copia o `NOME-AQUI.aspx`.

Depois disso a página fica full-screen, sem cabeçalho e sem barra de comandos — só a web part.

### Passo 8 (opcional) — Criar página do Dashboard

Repita o Passo 6 e 7, mas usando a web part **Dashboard — Biblioteca de Prompts** ao invés da Biblioteca. Configure com os mesmos títulos de lista.

### Passo 9 (opcional) — Adicionar como aba no Teams

Ver [seção específica](#-adicionar-como-aba-no-teams) abaixo.

---

## 🎨 Personalização e manutenção

**Toda a personalização é feita pelo admin sem tocar em código.** Não precisa recompilar nada nesses casos.

### Editar opções dos campos (Segmento, Ação, Categoria, Funciona com)

Ex.: adicionar `Copilot` como opção em "Ação", ou remover `Fabric` de "Funciona com".

1. Vá na lista `Biblioteca de Prompts` no SharePoint
2. Engrenagem ⚙️ → **Configurações da lista**
3. Clica no nome da coluna (ex.: `Segmento`)
4. Na seção **Opções**, adicione/remova valores
5. Clica **OK**
6. Na próxima abertura da página, os dropdowns do "Novo prompt" e "Editar prompt" refletem as opções novas automaticamente.

Vale pra qualquer campo Choice: `A_x00e7__x00e3_o`, `Categoria`, `Categoria0`, `Funcionacom`.

### Adicionar cor customizada para uma nova ferramenta

Se você adicionar `Google Gemini` em "Funciona com" e quiser que o card apareça com a cor azul do Gemini (em vez do cinza padrão):

1. Vá na página que tem a web part → clique em **Editar**
2. Clica na web part → aparece um ícone de lápis à esquerda (Editar web part)
3. No painel à direita, seção **Cores adicionais**
4. Cole no campo (uma por linha, formato `Nome=#HEXCOLOR`):
   ```
   Google Gemini=#4285F4
   Claude=#D97757
   Perplexity=#20B2AA
   ```
5. Fecha o painel
6. **Republica** a página

Os cards vão ficar com badge e borda coloridos com a cor definida.

> 💡 Ferramentas já embutidas com cor: Outlook, Teams, OneNote, Word, Excel, PowerPoint, Power BI, M365 Copilot, Copilot Studio, D365 CCaaS, D365 Customer Insights, D365 Sales, Fabric, Power Apps, Power Automate, Power Pages, Whiteboard. Essas você não precisa mexer.

### Trocar títulos das listas

Se preferir chamar de "Prompts" ao invés de "Biblioteca de Prompts":

1. Renomeia a lista no SharePoint (Configurações da lista → Título)
2. Vai na página → Editar → property pane da web part → altera **Título da lista de prompts** pra bater
3. Republica

### Modo escuro (dark mode)

- **Automático**: web part detecta o tema do Teams (`dark`/`default`) e aplica.
- **Manual**: botão 🌙/☀️ no header de cada web part alterna e salva no browser.
- Configuração fica salva em `localStorage` — vale só naquele browser.

---

## 💬 Adicionar como aba no Teams

Sua página SharePoint pode virar uma aba dentro de um canal do Teams.

1. Abra o Teams → canal desejado → **+ Adicionar aba**
2. Busca **SharePoint** → clica no app oficial (ícone azul)
3. No painel que abre, procura por **"Colar link do SharePoint"** (fica no rodapé)
4. Cola a URL da página publicada:
   ```
   https://SEU_TENANT.sharepoint.com/sites/NOME_DO_SITE/SitePages/Biblioteca-de-Prompts.aspx
   ```
5. **Salvar**

A página abre dentro do Teams. Como a página tá em layout `SingleWebPartAppPage`, ela já vem sem chrome — perfeita pra Teams.

Repita para a página do Dashboard, se tiver.

---

## 📁 Estrutura do projeto

```
sharepoint-biblioteca-prompt/
├── config/                          # Configuração do SPFx (não mexa)
├── scripts/                         # Scripts PowerShell para provisionar
│   ├── Setup-BibliotecaPrompts.ps1  # Cria as listas no site
│   └── Configure-Page.ps1           # Configura página em tela cheia
├── src/webparts/
│   ├── bibliotecaPrompt/            # Web part principal
│   │   ├── BibliotecaPromptWebPart.ts
│   │   ├── NewPromptDialog.ts       # Diálogo "Novo prompt"
│   │   ├── PromptDetailsDialog.ts   # Diálogo detalhes/editar
│   │   └── promptChoices.ts         # Defaults dos campos Choice
│   └── analytics/                   # Dashboard
│       └── AnalyticsWebPart.ts
├── sharepoint/solution/             # Onde o .sppkg é gerado (ignorado pelo Git)
├── package.json                     # Dependências npm
└── README.md                        # Este arquivo
```

---

## 🔄 Atualizar a solução

Se você mudou algo no código e quer aplicar nos clientes:

1. **Rebuild**:
   ```powershell
   npm run build
   ```
2. **Upload no App Catalog** de cada cliente — arraste o `.sppkg` novo por cima. Confirme "Replace it".
3. Cada cliente vai ter que **clicar em "Update" no Site Contents** (o SharePoint mostra o link "Update available").
4. Ctrl+F5 nas páginas.

> ⚠️ Mudar o código NÃO afeta as opções personalizadas pelo admin (choices, cores extras, título das listas). Property pane e listas ficam intactos.

---

## 🩹 Solução de problemas comuns

### `Please specify a valid client id for an Entra ID App Registration`
Você não passou o `-ClientId` no comando. Ou não rodou o `Register-PnPEntraIDAppForInteractiveLogin` no tenant ainda. Veja Passo 1.

### `Couldn't add this app` no site
Espere 10-15 min (propagação do App Catalog) e tenta de novo. Se persistir, delete o app do App Catalog (inclusive da lixeira) e reinstale.

### Web part aparece mas dá erro "Falha ao carregar prompts (HTTP 404)"
Os títulos das listas no property pane não batem com os títulos reais. Confere. Ou você ainda não rodou o script de setup das listas.

### Dropdown do "Novo prompt" vazio
Coluna Choice não tem opções cadastradas na lista, OU o campo está com nome interno diferente. Confira em Configurações da Lista.

### Cards sem cor de marca em ferramenta X
Cor não está mapeada. Adicione no property pane em "Cores adicionais" (ver [Personalização](#-personalização-e-manutenção)).

### `Set-PnPPage: You cannot host text controls inside a one column full width section`
Você colocou a web part em uma seção de largura total pelo browser. Use **`SingleWebPartAppPage`** via `Configure-Page.ps1` ao invés disso — ele resolve automaticamente.

### PowerShell 5.1 vs 7
Verifique com `$PSVersionTable.PSVersion`. Se `Major` for `5`, você está no PowerShell antigo do Windows. Use o PowerShell 7 preto (baixe pelo link em pré-requisitos).

### `A parameter cannot be found that matches parameter name 'Interactive'`
Você tá tentando usar o parâmetro `-Interactive` no `Register-PnPEntraIDAppForInteractiveLogin`. Ele foi removido em versões recentes. Rode sem o `-Interactive` (ele é padrão).

### Backtick (`` ` ``) somindo em comandos copiados
Cola tudo em uma linha só, sem quebra. O backtick é frágil no copy-paste.

### Erro `Insufficient privileges` ao rodar os scripts com ClientId
A app foi criada mas as permissões ainda não foram aprovadas pelo admin. Você tem duas coisas pra checar:

1. **Admin consent das permissões da app** — vai no Portal Entra → App registrations → sua app → **API permissions**. Cada permissão deve estar com status 🟢 **Granted for `<tenant>`**. Se estiver 🟡 **Not granted**, peça pro admin clicar em **Grant admin consent**.
2. **Sua conta como Site Collection Admin** (só necessário se for rodar `Set-PnPSite -SocialBarOnSitePagesDisabled`). Vá no site → engrenagem ⚙️ → **Site permissions** → **Site collection administrators** → adicione seu usuário. Se você não puder fazer isso, peça pro admin do site.

---

## 🔒 Segurança e Git

- **Nenhum segredo é versionado** neste projeto. Toda config (URLs, IDs, tokens) vive no property pane da web part ou nas listas do SharePoint.
- O repositório pode ficar **público** sem risco.
- `.gitignore` cobre corretamente `node_modules`, `lib`, `sharepoint/solution/`, `*.sppkg` — nada de artefato compilado é commitado.
- Os `ClientId`s do PnP são específicos por tenant — cada cliente registra o seu.

---

## 📝 Licença

MIT — use, copie, modifique à vontade. Sem garantias.

---

## 🙋 FAQ rápida

**P: Preciso ser desenvolvedor pra instalar isso?**
R: Não. Precisa saber copiar comandos no PowerShell e entender a UI do SharePoint. O README cobre passo a passo.

**P: O mesmo `.sppkg` serve pra todos os clientes?**
R: Sim. Toda config específica de cliente é feita no property pane e nas listas SharePoint. Zero rebuild por cliente.

**P: Se eu adicionar uma coluna nova na lista, ela aparece na web part?**
R: Não automaticamente. A web part só sabe dos campos hard-coded (Title, Ação, Prompt, Segmento, Categoria, Funciona com). Nova coluna precisa de mudança de código.

**P: Se o admin remover uma opção do Choice que já foi usada num item, o que acontece?**
R: O item mantém o valor antigo, mas a opção não aparece mais no dropdown. Se editar o item, precisa escolher outra opção.

**P: Funciona no Teams mobile?**
R: Sim, mas layout otimizado pra desktop. Cards de 260px ficam apertados em telas pequenas.

**P: Posso ter uma biblioteca por departamento (várias no mesmo site)?**
R: Sim — crie listas com nomes diferentes (ex.: `Prompts RH`, `Prompts Comercial`) e configure duas páginas, cada uma com um property pane apontando pra uma lista.

---

## 🚀 Roadmap / Ideias futuras

Coisas que podem entrar em versões futuras:

- [ ] Filtros combinados (App + Área ao mesmo tempo)
- [ ] Editar em lote
- [ ] Import/export CSV
- [ ] Analytics por período customizado
- [ ] Notificação por email quando alguém publica prompt (via Power Automate)
- [ ] Botão de "curtir" separado do favoritar

Contribuições / issues são bem-vindas.
