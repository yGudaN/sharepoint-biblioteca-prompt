# 🧙 Assistente de Implantação

Wizard gráfico (WinForms) que guia o usuário pela implantação completa da Biblioteca de Prompts em um cliente, sem precisar abrir PowerShell manualmente.

## Como rodar (modo desenvolvedor)

Pré-requisito: **PowerShell 7+** (baixar em https://learn.microsoft.com/pt-br/powershell/scripting/install/installing-powershell-on-windows).

Duplo-clique não abre — precisa rodar via terminal:

```powershell
# Se der erro de execução policy, libera pra sessão atual:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Rodar
pwsh -File .\assistente\Assistente.ps1
```

Ou dentro de um terminal PowerShell 7 já aberto:

```powershell
cd .\assistente
.\Assistente.ps1
```

## Como gerar um `.exe` (distribuir pro usuário final)

Assim o usuário só clica no `.exe` e usa. Não precisa saber PowerShell.

```powershell
# Instalar o compilador (só uma vez por máquina)
Install-Module ps2exe -Scope CurrentUser -Force

# Compilar
Invoke-PS2EXE .\assistente\Assistente.ps1 .\dist\BibliotecaPrompt-Assistente.exe -noConsole -title "Biblioteca de Prompts - Assistente" -description "Wizard de implantação da Biblioteca de Prompts" -company "Sua Empresa" -product "BibliotecaPrompt" -copyright "MIT"
```

Vai gerar `BibliotecaPrompt-Assistente.exe` (100-200 KB). Distribua esse arquivo.

> ⚠️ **Requisito na máquina do usuário final:** PowerShell 7+ instalado. O `.exe` gerado pelo PS2EXE ainda depende do PowerShell — ele apenas evita que o usuário precise abrir o terminal. Se você quiser fazer 100% standalone (sem PowerShell instalado), precisaria migrar pra Electron/WPF (fora do escopo dessa versão).

## O que o assistente faz (etapa por etapa)

### 1. Boas-vindas
Explica as etapas que vem pela frente.

### 2. Verificação de pré-requisitos
Checa:
- PowerShell 7+
- Módulo `PnP.PowerShell` instalado (se faltar, botão instala automaticamente)

### 3. Registro do App no Entra ID
Formulário:
- Tenant (só o prefixo, ex.: `empresax`)
- Nome do App (default: "PnP PowerShell - Biblioteca de Prompts")

Ao clicar em **Registrar**, roda `Register-PnPEntraIDAppForInteractiveLogin` com as permissões mínimas (`AllSites.FullControl` + `User.Read`). Retorna o ClientId no campo. Mostra aviso sobre admin consent se necessário.

Alternativa: se você **já tem o ClientId** (de outro cliente ou já registrado), cole direto no campo e pule o botão Registrar.

### 4. Provisionar listas
Formulário:
- URL do site (pré-preenchido com `https://<tenant>.sharepoint.com/sites/`)
- Nome da lista de prompts (default: `Biblioteca de Prompts`)
- Nome da lista de favoritos (default: `⭐ Meus Favoritos`)

Ao clicar em **Criar listas**, o assistente:
- Conecta ao SharePoint com o ClientId
- Cria (ou reutiliza) as duas listas
- Adiciona todas as colunas necessárias (Ação, Prompt, Segmento, Categoria, Funciona com, PromptID, Ativo)
- Popula os choices padrão

### 5. Passo manual (upload + página)
Mostra instruções pra você fazer no browser:
1. Subir o `.sppkg` no App Catalog
2. Instalar o app no site
3. Criar página em branco + adicionar web part + publicar
4. (Opcional) Segunda página com o Dashboard

Depois, você preenche no assistente:
- Nome do arquivo `.aspx` da página da Biblioteca
- (Opcional) Nome do arquivo `.aspx` da página do Dashboard

### 6. Configurar páginas
Ao clicar em **Configurar**, o assistente:
- Conecta ao SharePoint
- Aplica `LayoutType = SingleWebPartAppPage` em cada página (tela cheia sem chrome)
- Publica cada página
- Desativa a barra social do site (`SocialBarOnSitePagesDisabled = $true`)

### 7. Conclusão
Mostra resumo com todos os dados usados, incluindo o ClientId (pra você guardar). Botão pra abrir a página no browser.

## Fluxo de dados entre etapas

O ClientId gerado na etapa 3 é usado nas etapas 4 e 6. Se você fechar o assistente no meio, precisa começar do zero (não há salvamento de estado entre execuções).

## Limitações conhecidas

- **Upload do `.sppkg` não é automatizado.** É passo manual da etapa 5. Poderia ser automatizado via REST API do SharePoint, mas envolve complexidade adicional (multipart upload, permissões).
- **Sem retry automático.** Se algum passo falhar, você precisa voltar e tentar de novo manualmente.
- **UI é WinForms clássico.** Funcional mas com visual "anos 2000". Trocar por WPF/Electron aumentaria a complexidade.

## Roadmap futuro

- [ ] Automatizar upload do `.sppkg` no App Catalog via REST
- [ ] Salvar estado entre execuções (JSON local)
- [ ] Suporte a instalar em vários sites em batch
- [ ] Loga tudo em arquivo pra debug
- [ ] Executável Windows sem depender de PowerShell instalado (Electron/WPF)
