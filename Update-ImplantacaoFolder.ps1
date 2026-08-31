# Requires -Version 7.0
<#
.SYNOPSIS
    Sincroniza a pasta implantacao/ com o Assistente, o .sppkg e os scripts
    mais recentes do repositorio.

.DESCRIPTION
    Copia:
      assistente\Assistente.ps1                          -> implantacao\
      sharepoint\solution\sharepoint-biblioteca-prompt.sppkg -> implantacao\
      scripts\Setup-BibliotecaPrompts.ps1                -> implantacao\avancado\
      scripts\Configure-Page.ps1                         -> implantacao\avancado\

    Rode este script sempre que:
      - alterar o Assistente.ps1
      - recompilar o .sppkg (gulp bundle --ship && gulp package-solution --ship)
      - alterar algum script avulso

.EXAMPLE
    .\Update-ImplantacaoFolder.ps1
#>

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$dest = Join-Path $root 'implantacao'
$destAvancado = Join-Path $dest 'avancado'

if (-not (Test-Path $dest))         { New-Item -ItemType Directory -Path $dest | Out-Null }
if (-not (Test-Path $destAvancado)) { New-Item -ItemType Directory -Path $destAvancado | Out-Null }

$mapping = @(
    @{ From = 'assistente\Assistente.ps1';                        To = $dest }
    @{ From = 'sharepoint\solution\sharepoint-biblioteca-prompt.sppkg'; To = $dest }
    @{ From = 'scripts\Setup-BibliotecaPrompts.ps1';              To = $destAvancado }
    @{ From = 'scripts\Configure-Page.ps1';                       To = $destAvancado }
    @{ From = 'scripts\Diagnosticar-WebParts.ps1';                To = $destAvancado }
    @{ From = 'scripts\Criar-Pagina-Webpart.ps1';                 To = $destAvancado }
)

Write-Host ''
Write-Host '=== Sincronizando pasta implantacao/ ===' -ForegroundColor Cyan
Write-Host ''

$copiados = 0
$faltando = @()
foreach ($m in $mapping) {
    $src = Join-Path $root $m.From
    if (-not (Test-Path $src)) {
        $faltando += $m.From
        Write-Host "  [FALTA] $($m.From)" -ForegroundColor Yellow
        continue
    }
    Copy-Item -Force $src $m.To
    Write-Host "  [OK]    $($m.From)" -ForegroundColor Green
    $copiados++
}

# Sincroniza pasta assets/ (logo, icones)
$assetsSrc = Join-Path $root 'assistente\assets'
$assetsDst = Join-Path $dest 'assets'
if (Test-Path $assetsSrc) {
    if (-not (Test-Path $assetsDst)) { New-Item -ItemType Directory -Force -Path $assetsDst | Out-Null }
    Copy-Item -Force -Recurse "$assetsSrc\*" $assetsDst
    Write-Host "  [OK]    assistente\assets\* (recursivo)" -ForegroundColor Green
    $copiados++
}

Write-Host ''
Write-Host "Arquivos copiados: $copiados" -ForegroundColor Cyan

if ($faltando.Count -gt 0) {
    Write-Host ''
    Write-Host 'AVISOS:' -ForegroundColor Yellow
    if ($faltando -contains 'sharepoint\solution\sharepoint-biblioteca-prompt.sppkg') {
        Write-Host '  .sppkg nao encontrado. Rode antes:' -ForegroundColor Yellow
        Write-Host '     gulp bundle --ship' -ForegroundColor Yellow
        Write-Host '     gulp package-solution --ship' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host 'Pasta pronta em:' -ForegroundColor Cyan
Write-Host "  $dest"
Write-Host ''
Write-Host 'Para distribuir: zipa a pasta implantacao/ e envia.' -ForegroundColor Cyan
