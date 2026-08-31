# Requires -Version 7.0
<#
.SYNOPSIS
    Diagnostica quais web parts / apps estao instaladas em um site SharePoint,
    e mostra os IDs exatos que o PnP retorna.

.DESCRIPTION
    Rode este script para descobrir o GUID EXATO da web part
    "Biblioteca de Prompts" no seu site (que pode diferir do manifest).

    Uso:
      .\Diagnosticar-WebParts.ps1 -SiteUrl "https://empresa.sharepoint.com/sites/xyz" -ClientId "xxxx-..."
#>

param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " DIAGNOSTICO - Web Parts & Apps do site" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Site:     $SiteUrl"
Write-Host "ClientId: $ClientId"
Write-Host ""

Write-Host "Conectando..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

# ===================================================================
# 1) Apps instalados no site
# ===================================================================
Write-Host ""
Write-Host "=== 1) APPS INSTALADOS NESTE SITE (Get-PnPApp -Scope Site) ===" -ForegroundColor Green
try {
    $apps = Get-PnPApp -Scope Site -ErrorAction Stop
    if ($apps) {
        $apps | ForEach-Object {
            Write-Host "----"
            Write-Host "Id:                 $($_.Id)"
            Write-Host "Title:              $($_.Title)"
            Write-Host "ProductId:          $($_.ProductId)"
            Write-Host "AppCatalogVersion:  $($_.AppCatalogVersion)"
            Write-Host "InstalledVersion:   $($_.InstalledVersion)"
            Write-Host "Deployed:           $($_.Deployed)"
        }
    } else {
        Write-Host "Nenhum app instalado neste site." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Falha ao listar apps: $($_.Exception.Message)" -ForegroundColor Red
}

# ===================================================================
# 2) Cria uma pagina temporaria para consultar componentes
# ===================================================================
$tempPage = "Diag-Temp-$(Get-Random -Maximum 9999)"
Write-Host ""
Write-Host "=== 2) COMPONENTES DISPONIVEIS (pagina temp: $tempPage) ===" -ForegroundColor Green
try {
    Add-PnPPage -Name $tempPage | Out-Null
    $comps = Get-PnPAvailablePageComponents -Page "$tempPage.aspx"
    Write-Host "Total de componentes retornados: $($comps.Count)" -ForegroundColor Cyan
    Write-Host ""

    # Procura por qualquer coisa que pareca a nossa web part
    Write-Host "--- Componentes que citam 'biblioteca', 'prompt' ou 'analytics' ---" -ForegroundColor Yellow
    $ours = $comps | Where-Object {
        $_.Name -match 'iblioteca|prompt|nalytics' -or
        $_.Title -match 'iblioteca|Prompt|nalytics' -or
        $_.Manifest -match 'biblioteca|prompt'
    }
    if ($ours) {
        $ours | ForEach-Object {
            Write-Host "===="
            Write-Host "Id:            $($_.Id)"
            Write-Host "Name:          $($_.Name)"
            Write-Host "Title:         $($_.Title)"
            Write-Host "ComponentType: $($_.ComponentType)"
            Write-Host "Status:        $($_.Status)"
        }
    } else {
        Write-Host "Nada encontrado com nomes 'biblioteca/prompt/analytics'." -ForegroundColor Red
        Write-Host ""
        Write-Host "--- Todos os componentes disponiveis (Id | Name) ---" -ForegroundColor Cyan
        $comps | Sort-Object Name | ForEach-Object {
            Write-Host "  $($_.Id)  |  $($_.Name)"
        }
    }
} catch {
    Write-Host "Falha ao listar componentes: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Remove a pagina temp
    try {
        Remove-PnPPage -Identity "$tempPage.aspx" -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host ""
        Write-Host "Pagina temporaria '$tempPage' removida." -ForegroundColor DarkGray
    } catch {}
}

Disconnect-PnPOnline

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " FIM DO DIAGNOSTICO" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Copie a saida acima e me envie para eu ajustar o assistente." -ForegroundColor Yellow
Write-Host ""
