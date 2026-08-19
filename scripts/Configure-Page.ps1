<#
.SYNOPSIS
    Configura uma página SharePoint para exibição em tela cheia usando o layout
    SingleWebPartAppPage: sem cabeçalho, sem command bar, sem barra social,
    sem comentários — só a web part ocupando 100% da tela.

.DESCRIPTION
    Requisitos:
      - A página deve conter APENAS 1 web part. Se tiver mais, as extras serão removidas.
      - O script publica a página automaticamente ao final.
      - Se você já tem a web part configurada e funcionando, é só rodar este script.

.PARAMETER SiteUrl
    URL do site SharePoint.

.PARAMETER PageName
    Nome do arquivo da página, ex.: "Biblioteca-de-Prompts.aspx".

.PARAMETER ClientId
    Client ID do App Registration. Se omitido, usa $env:PNP_CLIENT_ID.

.EXAMPLE
    .\Configure-Page.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/<site>" -PageName "Biblioteca-de-Prompts.aspx" -ClientId "<guid>"

.NOTES
    Pré-requisito: Install-Module PnP.PowerShell -Scope CurrentUser -Force
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$PageName,
    [string]$ClientId = $env:PNP_CLIENT_ID
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Error "Passe -ClientId ou defina `$env:PNP_CLIENT_ID"
    exit 1
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "Módulo PnP.PowerShell não instalado. Rode: Install-Module PnP.PowerShell -Scope CurrentUser -Force"
    exit 1
}
Import-Module PnP.PowerShell

Write-Host ""
Write-Host "== Configurando página '$PageName' em tela cheia ==" -ForegroundColor Cyan
Write-Host "Site: $SiteUrl"
Write-Host ""

Write-Host "Conectando..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$page = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue
if (-not $page) {
    Write-Error "Página '$PageName' não encontrada. Verifique o nome do arquivo (com .aspx)."
    Disconnect-PnPOnline
    exit 1
}

Write-Host "[1/2] Aplicando layout SingleWebPartAppPage (tela cheia, sem chrome)..." -ForegroundColor Yellow
Set-PnPPage -Identity $PageName -LayoutType SingleWebPartAppPage | Out-Null

Write-Host "[2/2] Publicar..." -ForegroundColor Yellow
Set-PnPPage -Identity $PageName -Publish | Out-Null

Write-Host ""
Write-Host "== Pronto! ==" -ForegroundColor Green
Write-Host ""
Write-Host "A página agora ocupa 100% da tela, sem cabeçalho, sem barra de comandos, sem rodapé."
Write-Host "Se ainda ver a barra social (likes/views/comments) em outras páginas do site, rode 1 vez:"
Write-Host "  Set-PnPSite -SocialBarOnSitePagesDisabled `$true" -ForegroundColor Gray
Write-Host ""

Disconnect-PnPOnline
