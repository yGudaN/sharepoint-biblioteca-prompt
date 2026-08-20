<#
.SYNOPSIS
    Provisiona as listas SharePoint necessárias para a web part Biblioteca de Prompts.

.DESCRIPTION
    Cria (ou reutiliza) duas listas:
      1. Lista principal (biblioteca de prompts)
      2. Lista de favoritos por usuário
    Com todas as colunas + choices padrão nos nomes internos que a web part espera.
    Idempotente: rodar mais de uma vez não duplica campos.

.PARAMETER SiteUrl
    URL completa do site SharePoint alvo. Ex.: https://<tenant>.sharepoint.com/sites/<site>

.PARAMETER PromptsListTitle
    Título da lista principal. Padrão: "Biblioteca de Prompts".

.PARAMETER FavoritesListTitle
    Título da lista de favoritos. Padrão: "⭐ Meus Favoritos".

.PARAMETER ClientId
    Client ID (GUID) do App Registration no Entra ID. Obrigatório desde PnP.PowerShell 2.x.
    Se omitido, tenta ler de $env:PNP_CLIENT_ID.

    Para gerar (uma vez por tenant):
      Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell - Biblioteca de Prompts" -Tenant "<tenant>.onmicrosoft.com"

.EXAMPLE
    .\Setup-BibliotecaPrompts.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -ClientId "<guid>"

.NOTES
    Pré-requisito: Install-Module PnP.PowerShell -Scope CurrentUser -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [string]$PromptsListTitle = "Biblioteca de Prompts",

    [string]$FavoritesListTitle = "⭐ Meus Favoritos",

    [string]$ClientId = $env:PNP_CLIENT_ID
)

$ErrorActionPreference = 'Stop'

# ---- opções padrão dos campos Choice ----
$ACOES        = @('Analisar','Perguntar','Resumir','Criar','Encontrar','Aprender','Otimizar','Se preparar','Entender')
$SEGMENTOS    = @('Comercial','DP','Financeiro','Infra','Projetos','RH','Analista Funcional','Desenvolvedor(a)','Gerente de Projetos')
$CATEGORIAS   = @('Área','Função')
$FUNCIONA_COM = @('Outlook','Teams','OneNote','Word','Excel','PowerPoint','Power BI','M365 Copilot','Copilot Studio','D365 CCaaS / Customer Service','D365 Customer Insights - Journeys','D365 Sales','Fabric','Power Apps','Power Automate','Power Pages','Whiteboard')

function New-ChoiceFieldXml {
    param([string]$Name, [string]$DisplayName, [string[]]$Choices)
    $escaped = $Choices | ForEach-Object { [System.Security.SecurityElement]::Escape($_) }
    $choicesXml = ($escaped | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
    return "<Field Type='Choice' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Format='Dropdown'><CHOICES>$choicesXml</CHOICES></Field>"
}

function New-NoteFieldXml {
    param([string]$Name, [string]$DisplayName)
    return "<Field Type='Note' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' RichText='FALSE' NumLines='6' />"
}

function New-NumberFieldXml {
    param([string]$Name, [string]$DisplayName, [bool]$Required = $false)
    $req = if ($Required) { 'TRUE' } else { 'FALSE' }
    return "<Field Type='Number' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Required='$req' />"
}

function New-BooleanFieldXml {
    param([string]$Name, [string]$DisplayName, [bool]$Default = $true)
    $def = if ($Default) { '1' } else { '0' }
    return "<Field Type='Boolean' Name='$Name' StaticName='$Name' DisplayName='$DisplayName'><Default>$def</Default></Field>"
}

function Ensure-List {
    param([string]$Title)
    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "  [+] Criando lista '$Title'..." -ForegroundColor Yellow
        $list = New-PnPList -Title $Title -Template GenericList -EnableVersioning
    } else {
        Write-Host "  [=] Lista '$Title' já existe" -ForegroundColor Gray
    }
    return $list
}

function Ensure-Field {
    param([string]$ListTitle, [string]$InternalName, [string]$FieldXml)
    $existing = Get-PnPField -List $ListTitle -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host "    [+] Campo '$InternalName'" -ForegroundColor Yellow
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $FieldXml | Out-Null
    } else {
        Write-Host "    [=] Campo '$InternalName' já existe" -ForegroundColor Gray
    }
}

# ---- verifica ClientId ----
if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host ""
    Write-Host "ERRO: Nenhum ClientId informado." -ForegroundColor Red
    Write-Host ""
    Write-Host "A partir do PnP.PowerShell 2.x, você precisa de um App Registration próprio." -ForegroundColor Yellow
    Write-Host "Rode UMA VEZ por tenant (admin do tenant):" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP PowerShell - Biblioteca de Prompts' -Tenant '<tenant>.onmicrosoft.com'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Depois passe o ClientId retornado:" -ForegroundColor Yellow
    Write-Host "  .\Setup-BibliotecaPrompts.ps1 -SiteUrl '...' -ClientId '<guid>'" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

# ---- verifica módulo PnP ----
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "Módulo PnP.PowerShell não instalado. Rode: Install-Module PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}
Import-Module PnP.PowerShell

Write-Host ""
Write-Host "== Biblioteca de Prompts - Provisionamento ==" -ForegroundColor Cyan
Write-Host "Site:      $SiteUrl"
Write-Host "Prompts:   $PromptsListTitle"
Write-Host "Favoritos: $FavoritesListTitle"
Write-Host ""

Write-Host "Conectando..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# =====================================================
# Lista principal
# =====================================================
Write-Host ""
Write-Host "[1/2] Lista principal ('$PromptsListTitle')" -ForegroundColor Cyan
Ensure-List -Title $PromptsListTitle | Out-Null

$xmlAcao       = New-ChoiceFieldXml -Name 'A_x00e7__x00e3_o' -DisplayName 'Ação' -Choices $ACOES
$xmlPrompt     = New-NoteFieldXml -Name 'Prompt' -DisplayName 'Prompt'
$xmlSegmento   = New-ChoiceFieldXml -Name 'Categoria' -DisplayName 'Segmento' -Choices $SEGMENTOS
$xmlCategoria  = New-ChoiceFieldXml -Name 'Categoria0' -DisplayName 'Categoria' -Choices $CATEGORIAS
$xmlFuncCom    = New-ChoiceFieldXml -Name 'Funcionacom' -DisplayName 'Funciona com' -Choices $FUNCIONA_COM
$xmlPromptId   = New-NumberFieldXml -Name 'PromptID' -DisplayName 'PromptID' -Required $true
$xmlAtivo      = New-BooleanFieldXml -Name 'Ativo' -DisplayName 'Ativo' -Default $true

Ensure-Field -ListTitle $PromptsListTitle -InternalName 'A_x00e7__x00e3_o' -FieldXml $xmlAcao
Ensure-Field -ListTitle $PromptsListTitle -InternalName 'Prompt' -FieldXml $xmlPrompt
Ensure-Field -ListTitle $PromptsListTitle -InternalName 'Categoria' -FieldXml $xmlSegmento
Ensure-Field -ListTitle $PromptsListTitle -InternalName 'Categoria0' -FieldXml $xmlCategoria
Ensure-Field -ListTitle $PromptsListTitle -InternalName 'Funcionacom' -FieldXml $xmlFuncCom
Ensure-Field -ListTitle $PromptsListTitle -InternalName 'Ativo' -FieldXml $xmlAtivo

# =====================================================
# Lista de favoritos
# =====================================================
Write-Host ""
Write-Host "[2/2] Lista de favoritos ('$FavoritesListTitle')" -ForegroundColor Cyan
Ensure-List -Title $FavoritesListTitle | Out-Null

Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'PromptID' -FieldXml $xmlPromptId
Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'A_x00e7__x00e3_o' -FieldXml $xmlAcao
Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'Prompt' -FieldXml $xmlPrompt
Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'Categoria' -FieldXml $xmlSegmento
Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'Categoria0' -FieldXml $xmlCategoria
Ensure-Field -ListTitle $FavoritesListTitle -InternalName 'Funcionacom' -FieldXml $xmlFuncCom

Write-Host ""
Write-Host "== Concluído com sucesso! ==" -ForegroundColor Green
Write-Host "Próximos passos:"
Write-Host "  1. Adicione a web part 'Biblioteca de Prompts' numa página do site"
Write-Host "  2. No property pane, confirme:"
Write-Host "       Título da lista de prompts    = $PromptsListTitle"
Write-Host "       Título da lista de favoritos  = $FavoritesListTitle"
Write-Host "  3. Publique a página"
Write-Host ""

Disconnect-PnPOnline
