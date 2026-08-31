# Requires -Version 7.0
<#
.SYNOPSIS
    Cria uma pagina modern + adiciona a web part da Biblioteca (ou Dashboard),
    com log verbose em cada passo para debug.

.DESCRIPTION
    Roda fora do ThreadJob do assistente para exibir erros que o wizard pode
    estar engolindo. Cole o output aqui no chat pra debug.

.PARAMETER SiteUrl
    URL completa do site SharePoint.

.PARAMETER ClientId
    ClientId do App PnP registrado no Entra.

.PARAMETER PageName
    Nome da pagina (sem .aspx).

.PARAMETER Kind
    'biblioteca' (padrao) ou 'dashboard'.

.EXAMPLE
    .\Criar-Pagina-Webpart.ps1 -SiteUrl "https://crm202149.sharepoint.com/sites/TestePaginaAutomatica1" `
        -ClientId "e378f672-792d-4b55-8d3d-cc94d2dd3b83" `
        -PageName "Biblioteca-Teste-Auto" -Kind biblioteca
#>

param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$ClientId,
    [Parameter(Mandatory)][string]$PageName,
    [ValidateSet('biblioteca','dashboard')]
    [string]$Kind = 'biblioteca'
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$file = if ($PageName -match '\.aspx$') { $PageName } else { "$PageName.aspx" }
$pageBase = $PageName -replace '\.aspx$',''

# Manifest IDs / nomes esperados
$WEBPARTS = @{
    biblioteca = @{
        Id    = 'c8e4a1f2-7b3d-4e9a-8f5c-6d2b1a9e3f4c'
        Name  = 'Biblioteca de Prompts'
        Props = @{
            targetListTitle    = 'Biblioteca de Prompts'
            favoritesListTitle = '⭐ Meus Favoritos'
            promptIdField      = 'PromptID'
            copyFields         = 'acao,Prompt,Segmento,Categoria,Funcionacom'
            extraToolColors    = ''
        }
    }
    dashboard = @{
        Id    = 'd3f7a2e4-6b91-4c8f-a5d2-1c9e4b7f3a8b'
        Name  = 'Dashboard — Biblioteca de Prompts'
        Props = @{
            targetListTitle    = 'Biblioteca de Prompts'
            favoritesListTitle = '⭐ Meus Favoritos'
            promptIdField      = 'PromptID'
            topN               = 10
        }
    }
}

$wp = $WEBPARTS[$Kind]

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " Criar pagina + web part + configurar (debug)" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "Site:      $SiteUrl"
Write-Host "Page:      $file"
Write-Host "Kind:      $Kind"
Write-Host "WP Id:     $($wp.Id)"
Write-Host "WP Name:   $($wp.Name)"
Write-Host "Props:     $($wp.Props | ConvertTo-Json -Compress)"
Write-Host ""

Import-Module PnP.PowerShell -ErrorAction Stop

# ------------------------------------------------------------
# 0) Conectar
# ------------------------------------------------------------
Write-Host "[0/9] Conectando ao SharePoint..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    Write-Host "     OK conectado." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# ------------------------------------------------------------
# 1) Verificar se a pagina ja existe
# ------------------------------------------------------------
Write-Host "[1/9] Verificando se a pagina ja existe..." -ForegroundColor Cyan
try {
    $existing = Get-PnPPage -Identity $file -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "     ATENCAO: pagina '$file' ja existe. Delete manualmente para recriar." -ForegroundColor Yellow
        Write-Host "     Encerrando." -ForegroundColor Yellow
        Disconnect-PnPOnline
        exit 0
    } else {
        Write-Host "     OK nao existe, prosseguir." -ForegroundColor Green
    }
} catch {
    Write-Host "     Erro na consulta: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# ------------------------------------------------------------
# 2) Criar a pagina
# ------------------------------------------------------------
Write-Host "[2/9] Criando pagina '$pageBase'..." -ForegroundColor Cyan
try {
    $newPage = Add-PnPPage -Name $pageBase -ErrorAction Stop
    Write-Host "     OK pagina criada." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}
Write-Host ""

# ------------------------------------------------------------
# 3) Buscar componente
# ------------------------------------------------------------
Write-Host "[3/9] Buscando componente '$($wp.Name)'..." -ForegroundColor Cyan
$comps = @()
try {
    $comps = Get-PnPAvailablePageComponents -Page $file -ErrorAction Stop
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}

$wpIdNorm = $wp.Id.ToLower()
$found = @($comps | Where-Object {
    $cid = ($_.Id.ToString() -replace '[{}]','').ToLower()
    $cid -eq $wpIdNorm -or $_.Name -eq $wp.Name
})

if ($found.Count -eq 0) {
    Write-Host "     NAO ENCONTRADO." -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}
$comp = $found[0]
Write-Host "     ENCONTRADO. Id=$($comp.Id)  Name='$($comp.Name)'" -ForegroundColor Green
Write-Host ""

# ------------------------------------------------------------
# 4) Adicionar a web part (SEM properties primeiro, para nao falhar)
# ------------------------------------------------------------
Write-Host "[4/9] Adicionando web part (Add-PnPPageWebPart -Component ...)..." -ForegroundColor Cyan
try {
    Add-PnPPageWebPart -Page $file -Component $comp -ErrorAction Stop | Out-Null
    Write-Host "     OK web part adicionada." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline
    exit 1
}
Write-Host ""

# ------------------------------------------------------------
# 5) Configurar propriedades via Set-PnPPageWebPart
# ------------------------------------------------------------
Write-Host "[5/9] Configurando propriedades da web part (Set-PnPPageWebPart -PropertiesJson)..." -ForegroundColor Cyan
$propsJson = $wp.Props | ConvertTo-Json -Compress
Write-Host "     JSON: $propsJson"
try {
    $onPage = Get-PnPPageComponent -Page $file -ErrorAction Stop
    $ourInstance = $onPage | Where-Object {
        $wid = ($_.WebPartId.ToString() -replace '[{}]','').ToLower()
        $wid -eq $wpIdNorm
    } | Select-Object -First 1

    if (-not $ourInstance) {
        Write-Host "     ATENCAO: nao achei instancia da web part na pagina para configurar." -ForegroundColor Yellow
    } else {
        Write-Host "     Instance id (Identity): $($ourInstance.InstanceId)"
        Set-PnPPageWebPart -Page $file -Identity $ourInstance.InstanceId -PropertiesJson $propsJson -ErrorAction Stop | Out-Null
        Write-Host "     OK propriedades atualizadas." -ForegroundColor Green
    }
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# ------------------------------------------------------------
# 6) Aplicar layout SingleWebPartAppPage (tela cheia sem chrome)
# ------------------------------------------------------------
Write-Host "[6/9] Aplicando layout SingleWebPartAppPage..." -ForegroundColor Cyan
try {
    Set-PnPPage -Identity $file -LayoutType SingleWebPartAppPage -ErrorAction Stop | Out-Null
    Write-Host "     OK layout aplicado." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# ------------------------------------------------------------
# 7) Publicar
# ------------------------------------------------------------
Write-Host "[7/9] Publicando pagina..." -ForegroundColor Cyan
try {
    Set-PnPPage -Identity $file -Publish -ErrorAction Stop | Out-Null
    Write-Host "     OK publicada." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# ------------------------------------------------------------
# 8) Desligar barra social do site
# ------------------------------------------------------------
Write-Host "[8/9] Desligando barra social do site (Set-PnPSite -SocialBarOnSitePagesDisabled)..." -ForegroundColor Cyan
try {
    Set-PnPSite -SocialBarOnSitePagesDisabled $true -ErrorAction Stop
    Write-Host "     OK barra social desligada." -ForegroundColor Green
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}
Write-Host ""

# ------------------------------------------------------------
# 9) Verificar componentes na pagina
# ------------------------------------------------------------
Write-Host "[9/9] Verificando componentes ATUAIS na pagina..." -ForegroundColor Cyan
try {
    $onPage = Get-PnPPageComponent -Page $file -ErrorAction Stop
    if ($onPage) {
        Write-Host "     Encontrei $($onPage.Count) componente(s):" -ForegroundColor Green
        foreach ($c in $onPage) {
            Write-Host "       - Tipo: $($c.GetType().Name)"
            Write-Host "         Title: $($c.Title)"
            Write-Host "         WebPartId: $($c.WebPartId)"
            Write-Host "         InstanceId: $($c.InstanceId)"
        }
    } else {
        Write-Host "     ATENCAO: pagina esta VAZIA - web part nao ficou salva." -ForegroundColor Red
    }
} catch {
    Write-Host "     FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "URL da pagina: $SiteUrl/SitePages/$file" -ForegroundColor Cyan
Disconnect-PnPOnline
Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " FIM - copie a saida acima e mande no chat" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
