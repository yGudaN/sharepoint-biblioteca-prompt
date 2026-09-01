# 📚 SharePoint Biblioteca de Prompts

Solução completa para gerenciar uma biblioteca de prompts de IA dentro do SharePoint e Teams, com favoritos por usuário, dashboard de contribuições e dark mode automático.

Duas Web Parts SPFx que rodam em qualquer site SharePoint moderno.

---

## 📋 Índice

1. [O que essa solução faz](#-o-que-essa-solução-faz)
2. [Pré-requisitos](#-pré-requisitos)
3. [Como instalar (assistente gráfico)](#-como-instalar-assistente-gráfico)
4. [Personalização e manutenção](#-personalização-e-manutenção)
5. [Adicionar como aba no Teams](#-adicionar-como-aba-no-teams)
6. [Solução de problemas](#-solução-de-problemas)
7. [Para desenvolvedores](#-para-desenvolvedores)

---

## 🎯 O que essa solução faz

Duas Web Parts, cada uma numa página:

**1. Biblioteca de Prompts** — a tela principal
- Grid de cards com todos os prompts publicados
- Busca por texto (título, prompt, categoria)
- Filtros por **App**, **Área**, **Função** (agrupa em seções) e **⭐ Favoritos**
- Botão ⭐ para favoritar (cada usuário tem seus próprios favoritos)
- Botão 📋 para copiar o texto do prompt
- Botão **+ Novo prompt** com diálogo para criar
- Clique no card abre modal com edição (só o autor pode editar/excluir)
- 🌙/☀️ toggle de dark mode (detecta tema do Teams automaticamente)

**2. Dashboard — Biblioteca de Prompts** — página opcional de métricas
- KPIs: total de prompts, favoritamentos, autores distintos
- Ranking dos **contribuidores** (quem publicou mais prompts)
- Ranking dos **prompts mais favoritados** (clicáveis para abrir detalhes)
- Filtro por período: todos / 30 dias / 90 dias

Ambas rodam em **SharePoint (browser)** e **abas do Teams**, totalmente customizáveis pelo admin sem mexer em código.

---

## ✅ Pré-requisitos

Você precisa de:

### 1. Windows 10 ou 11

### 2. PowerShell 7 ou superior
O Windows já vem com PowerShell 5.1, mas o assistente precisa da versão **7+**.

Download: https://aka.ms/powershell-release?tag=stable

Como saber se você tem: abra o menu Iniciar, digite "PowerShell 7". Se aparecer, tá instalado.

### 3. Conta Microsoft 365 com permissões:
- **Admin do tenant** (para registrar o App do PnP no Entra ID — apenas na primeira vez do tenant)
- **App Catalog admin** (para subir o `.sppkg` no App Catalog)
- **Owner do site** (para criar as listas e páginas)

Se você não é Global Admin mas o admin do tenant já registrou o App do PnP em uma instalação anterior, ele pode te passar o `ClientId` existente e você pula o passo 3.

---

## 🧙 Como instalar (assistente gráfico)

### 1. Baixe a pasta de implantação

Vá em [github.com/yGudaN/sharepoint-biblioteca-prompt](https://github.com/yGudaN/sharepoint-biblioteca-prompt) → clique em **Code → Download ZIP** → extraia o zip.

Dentro do zip, você só precisa da pasta **`implantacao/`**. O resto (código-fonte) pode ignorar.

> 💡 Pode salvar a pasta `implantacao/` em qualquer lugar do PC (Área de Trabalho, OneDrive, pen drive). Copie/renomeie à vontade — o assistente funciona a partir de qualquer caminho.

### 2. Abra o assistente

Dentro da pasta `implantacao/`, dê **duplo-clique** em **`Iniciar-Assistente.bat`**.

Se aparecer aviso do Windows SmartScreen:
1. Clique em **Mais informações**
2. Clique em **Executar assim mesmo**

O assistente abre em uma janela gráfica.

### 3. Siga os 7 passos

O assistente te guia por tudo. Você só preenche formulários — nenhum comando manual.

| Passo | O que faz | O que você precisa saber |
|---|---|---|
| **1. Boas-vindas** | Resume o que vem pela frente | — |
| **2. Pré-requisitos** | Verifica PowerShell 7 e módulo PnP.PowerShell (instala se faltar) | Só clicar em **Verificar** |
| **3. Registro do App no Entra ID** | Cria o App Registration no tenant (só uma vez por tenant) | Nome do tenant (ex.: `empresax.onmicrosoft.com`). Valida o formato do `ClientId` antes de avançar |
| **4. Provisionar listas** | Cria (ou atualiza) as listas `Biblioteca de Prompts` e `⭐ Meus Favoritos` no site com campos obrigatórios e view padrão | URL do site e o `ClientId` do passo 3. Idempotente — pode rodar de novo em cima |
| **5. Upload do `.sppkg` + instalar app** | Instruções para o passo manual (subir no App Catalog + Add an app no site) | Detecta o `.sppkg` na pasta, tem botão **📂 Abrir pasta**. Marque a caixa de confirmação para avançar |
| **6. Criar páginas + configurar** | **Automático**: cria as páginas, insere as web parts (Biblioteca + Dashboard), aplica layout tela cheia (`SingleWebPartAppPage`) e desliga a barra social do site | Nome das páginas a criar (ex.: `Biblioteca-de-Prompts`, `Dashboard-Biblioteca-de-Prompts`) |
| **7. Concluído** | Resumo da instalação com botão para abrir a Biblioteca no browser | — |

### 4. Feito!

Depois do passo 7, a Biblioteca de Prompts tá funcionando. Acesse a página no SharePoint ou adicione como aba no Teams.

### ✨ Novidades desta versão

- **Criação automática de páginas** — o assistente cria as páginas, adiciona as web parts com as propriedades corretas, aplica o layout tela cheia e desliga a barra social. Zero configuração manual pós-upload do `.sppkg`.
- **Listas com campos obrigatórios** — os campos `Ação`, `Prompt`, `Segmento`, `Categoria` e `Funciona com` viram obrigatórios. View padrão já vem com todas as colunas relevantes.
- **Verificação idempotente** — no passo 4 o botão vira "Verificar as listas". Roda em cima de instalações existentes sem duplicar nada e atualiza a config quando precisa.
- **Validação clara** — botão Avançar sempre habilitado; se faltar validação, aparece mensagem explicando o que falta.
- **Diagnóstico** — em `avancado/` tem `Diagnosticar-WebParts.ps1` e `Criar-Pagina-Webpart.ps1` para debug quando algo dá errado.
- **Identidade visual Bizapp** — cores roxas com gradiente no header, fonte Inter embutida (sem depender de instalação), logo Bizapp opcional em `assets/logo-bizapp.png`.

---

## 🎨 Personalização e manutenção

**Toda a personalização é feita pelo admin sem tocar em código.** Não precisa recompilar nada.

### Editar opções dos campos (Segmento, Ação, Categoria, Funciona com)

Ex.: adicionar `Copilot` como opção em "Ação", ou remover `Fabric` de "Funciona com".

1. Vá na lista `Biblioteca de Prompts` no SharePoint
2. Engrenagem ⚙️ → **Configurações da lista**
3. Clica no nome da coluna (ex.: `Segmento`)
4. Na seção **Opções**, adicione/remova valores
5. Clica **OK**
6. Na próxima abertura da página, os dropdowns já refletem as opções novas.

Vale para qualquer campo Choice: `acao`, `Segmento`, `Categoria`, `Funcionacom`.

### Adicionar cor customizada para uma nova ferramenta

Se você adicionar `Google Gemini` em "Funciona com" e quiser que o card apareça com a cor azul do Gemini:

1. Vá na página que tem a web part → clique em **Editar**
2. Clique na web part → aparece um ícone de lápis à esquerda (Editar web part)
3. No painel à direita, seção **Cores adicionais**
4. Cole no campo (uma por linha, formato `Nome=#HEXCOLOR`):
   ```
   Google Gemini=#4285F4
   Claude=#D97757
   Perplexity=#20B2AA
   ```
5. Feche o painel
6. **Republique** a página

Os cards vão ficar com badge e borda coloridos.

> 💡 Ferramentas já embutidas com cor: Outlook, Teams, OneNote, Word, Excel, PowerPoint, Power BI, M365 Copilot, Copilot Studio, D365 CCaaS, D365 Customer Insights, D365 Sales, Fabric, Power Apps, Power Automate, Power Pages, Whiteboard.

### Trocar títulos das listas

Se preferir "Prompts" em vez de "Biblioteca de Prompts":

1. Renomeie a lista no SharePoint (Configurações da lista → Título)
2. Vá na página → Editar → property pane da web part → altere **Título da lista de prompts** para bater

---

## 💬 Adicionar como aba no Teams

1. No Teams, entre no canal onde quer a aba
2. Clica no **+** no topo do canal
3. Procure por **SharePoint pages**
4. Selecione a página **Biblioteca de Prompts** do site
5. Salvar

A web part já detecta que está dentro do Teams e ajusta o tema automaticamente (claro/escuro).

---

## 🔧 Solução de problemas

**Assistente não abre com duplo-clique**  
Instale o PowerShell 7 (link nos pré-requisitos) e tente de novo.

**Windows SmartScreen bloqueia o `.bat`**  
Clique em **Mais informações → Executar assim mesmo**.

**"Couldn't add this app" no passo 5 do SharePoint**  
Aguarde 10–15 minutos (propagação do App Catalog) e tente de novo.

**Passo 3 (Registro do App) — não sou Global Admin**  
Peça pro admin do tenant executar o registro. Ele te passa o `ClientId` e você usa no passo 4.

**Passo 4 (Provisionar listas) — falha de permissão**  
Confirme que você é **Owner do site** onde vai instalar. Se não for, peça pro admin do site te adicionar.

**Passo 6 (Configurar páginas) — falha ao desligar barra social**  
Esse passo específico exige que sua conta seja **Site Collection Administrator**. Vá em Site → Engrenagem ⚙️ → **Permissões do site** → **Administradores da coleção de sites** → adicione seu usuário. Depois rode o passo 6 de novo.

**Web part instalada mas os prompts não aparecem**  
Confirme no painel de propriedades da web part:
- **Título da lista de prompts**: precisa bater exatamente com o nome da lista criada pelo assistente (padrão: `Biblioteca de Prompts`).
- **Título da lista de favoritos**: idem (padrão: `⭐ Meus Favoritos`).

---

## 👨‍💻 Para desenvolvedores

Se você quer **modificar o código** (não só instalar), precisa de:

- **Node.js 22.14+** (não use 20, 23 ou 24) — https://nodejs.org
- **VS Code** — https://code.visualstudio.com
- **Git** — https://git-scm.com

### Clonar e compilar

```powershell
git clone https://github.com/yGudaN/sharepoint-biblioteca-prompt.git
cd sharepoint-biblioteca-prompt
npm install
npm run build
```

O `.sppkg` gerado fica em `sharepoint/solution/sharepoint-biblioteca-prompt.sppkg`.

### Sincronizar a pasta `implantacao/` depois do build

```powershell
.\Update-ImplantacaoFolder.ps1
```

Isso copia o `.sppkg` recém-compilado, o `Assistente.ps1` e os scripts avulsos para a pasta `implantacao/`, deixando ela pronta para distribuir.

### Estrutura do projeto

```
├── src/webparts/
│   ├── bibliotecaPrompt/    ← web part principal
│   └── analytics/            ← dashboard
├── scripts/                  ← scripts PowerShell avulsos (uso via CLI)
├── assistente/               ← wizard WinForms (fonte)
├── implantacao/              ← pacote pronto para distribuição
│   ├── Iniciar-Assistente.bat
│   ├── Assistente.ps1
│   ├── sharepoint-biblioteca-prompt.sppkg
│   └── avancado/             ← scripts CLI (opcionais para instalação manual)
├── sharepoint/solution/      ← saída do build (.sppkg original)
└── Update-ImplantacaoFolder.ps1  ← sincroniza implantacao/ após build
```

### Regra de ouro para dev

Sempre que mudar código ou script, **rode o build + sync antes de commitar**:

```powershell
npm run build
.\Update-ImplantacaoFolder.ps1
git add -A
git commit -m "..."
git push
```

Isso garante que a pasta `implantacao/` no repositório sempre reflete a versão mais recente do código.

---

## 🔗 Links úteis

- Repositório: https://github.com/yGudaN/sharepoint-biblioteca-prompt
- SPFx docs: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
- PnP.PowerShell: https://pnp.github.io/powershell/
