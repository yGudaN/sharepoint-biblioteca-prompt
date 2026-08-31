# Requires -Version 7.0
<#
.SYNOPSIS
    Assistente grafico de implantacao da solucao Biblioteca de Prompts.

.DESCRIPTION
    Wizard passo-a-passo que guia o usuario pela instalacao completa:
      1. Boas-vindas
      2. Verificacao de pre-requisitos (PowerShell 7+, modulo PnP.PowerShell)
      3. Registro do App no Entra ID (gera ClientId)
      4. Provisionamento das listas SharePoint
      5. Passo manual: subir .sppkg no App Catalog + instalar app no site
      6. Criacao automatica das paginas + web parts + layout tela cheia
      7. Conclusao

    Todos os comandos PowerShell rodam no fundo. O usuario so preenche formularios.

.NOTES
    Para compilar como .exe:
      Install-Module ps2exe -Scope CurrentUser -Force
      Invoke-PS2EXE .\Assistente.ps1 .\BibliotecaPrompt-Assistente.exe -noConsole -title "Biblioteca de Prompts - Assistente"
#>

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================================================================
# FONTES CUSTOMIZADAS (Inter)
# =========================================================================
$global:privateFonts = New-Object System.Drawing.Text.PrivateFontCollection
$global:fontUiName = 'Segoe UI'  # fallback se Inter nao carregar
$fontRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if ($fontRoot) {
    $fontsDir = @(
        (Join-Path $fontRoot 'assets\fonts'),
        (Join-Path $fontRoot 'fonts')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($fontsDir) {
        Get-ChildItem -Path $fontsDir -Filter 'Inter-*.ttf' -File -ErrorAction SilentlyContinue | ForEach-Object {
            try { $global:privateFonts.AddFontFile($_.FullName) } catch { }
        }
        if ($global:privateFonts.Families.Count -gt 0) {
            $interFam = $global:privateFonts.Families | Where-Object { $_.Name -match 'Inter' } | Select-Object -First 1
            if ($interFam) { $global:fontUiName = $interFam.Name }
        }
    }
}

function global:New-UiFont {
    param([single]$Size = 10, [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular, [string]$FamilyOverride)
    $fam = if ($FamilyOverride) { $FamilyOverride } else { $global:fontUiName }
    if ($fam -eq $global:fontUiName -and $global:privateFonts.Families.Count -gt 0) {
        $family = $global:privateFonts.Families | Where-Object { $_.Name -eq $fam } | Select-Object -First 1
        if ($family) { return New-Object System.Drawing.Font($family, $Size, $Style) }
    }
    return New-Object System.Drawing.Font($fam, $Size, $Style)
}

# =========================================================================
# ESTADO GLOBAL
# =========================================================================
$global:state = @{
    Tenant             = ''
    AppName            = 'PnP PowerShell - Biblioteca de Prompts'
    ClientId           = ''
    SiteUrl            = ''
    PromptsListTitle   = 'Biblioteca de Prompts'
    FavoritesListTitle = '⭐ Meus Favoritos'
    BibliotecaPageName = ''
    DashboardPageName  = ''
}

$global:currentStep = 1
$global:totalSteps  = 7
$global:onNext = $null   # callback opcional executado antes de avancar

# Paleta Bizapp
$colorPrimary        = [System.Drawing.Color]::FromArgb(0x95, 0x3C, 0xCC)  # #953CCC roxo principal
$colorPrimaryH       = [System.Drawing.Color]::FromArgb(0x7B, 0x2E, 0xB0)  # roxo mais escuro (hover)
$colorGradientStart  = [System.Drawing.Color]::FromArgb(0x80, 0x34, 0xAE)  # #8034AE
$colorGradientEnd    = [System.Drawing.Color]::FromArgb(0xCB, 0x61, 0xE8)  # #CB61E8
$colorBg             = [System.Drawing.Color]::FromArgb(0xF3, 0xF2, 0xF5)  # #F3F2F5 off-white
$colorText           = [System.Drawing.Color]::FromArgb(0x06, 0x01, 0x0A)  # #06010A preto
$colorMuted          = [System.Drawing.Color]::FromArgb(96, 94, 92)
$colorSuccess        = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorDanger         = [System.Drawing.Color]::FromArgb(164, 38, 44)
$colorHeaderText     = [System.Drawing.Color]::White                        # #FFFFFF
$colorDisabledBg     = [System.Drawing.Color]::FromArgb(0xCF, 0xCE, 0xD1)  # cinza claro
$colorDisabledText   = [System.Drawing.Color]::FromArgb(0x06, 0x01, 0x0A)  # texto preto

# =========================================================================
# HELPERS
# =========================================================================
function global:Test-HasEmoji {
    param([string]$Text)
    if (-not $Text) { return $false }
    foreach ($c in $Text.ToCharArray()) {
        $code = [int]$c
        if ($code -ge 0x2600  -and $code -le 0x27BF) { return $true }
        if ($code -ge 0x2B00  -and $code -le 0x2BFF) { return $true }
        if ($code -ge 0xD800  -and $code -le 0xDBFF) { return $true }
        if ($code -ge 0xFE00  -and $code -le 0xFE0F) { return $true }
    }
    return $false
}

function global:Test-GuidFormat {
    param([string]$Value)
    if (-not $Value) { return $false }
    return ($Value -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

function global:Get-UiFontName {
    param([string]$Text)
    if (Test-HasEmoji $Text) { return 'Segoe UI Emoji' } else { return $global:fontUiName }
}

function global:New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 680, [int]$FontSize = 10, [bool]$Bold = $false, [System.Drawing.Color]$Color = $colorText)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.AutoSize = ($Width -eq 0)
    $h = [int]($FontSize * 1.8) + 8
    if ($Width -gt 0) { $l.Size = New-Object System.Drawing.Size($Width, $h) }
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $l.Font = New-UiFont -FamilyOverride (Get-UiFontName $Text) -Size $FontSize -Style $style
    $l.ForeColor = $Color
    return $l
}

function global:New-TextInput {
    param([int]$X, [int]$Y, [int]$Width = 400, [string]$Text = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($Width, 24)
    $t.Font = New-UiFont -Size 10
    $t.Text = $Text
    return $t
}

function global:New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 140, [bool]$Primary = $false)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($Width, 32)
    $b.Font = New-UiFont -FamilyOverride (Get-UiFontName $Text) -Size 10
    $b.FlatStyle = 'Flat'
    if ($Primary) {
        $b.BackColor = $colorPrimary
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderSize = 0
        $b.Tag = 'primary'
    } else {
        $b.BackColor = [System.Drawing.Color]::White
        $b.ForeColor = $colorText
        $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
        $b.Tag = 'secondary'
    }
    # Estado desabilitado: fundo cinza + texto preto
    $applyEnabledState = {
        if ($b.Enabled) {
            if ($b.Tag -eq 'primary') {
                $b.BackColor = $colorPrimary
                $b.ForeColor = [System.Drawing.Color]::White
            } else {
                $b.BackColor = [System.Drawing.Color]::White
                $b.ForeColor = $colorText
            }
        } else {
            $b.BackColor = $colorDisabledBg
            $b.ForeColor = $colorDisabledText
        }
    }.GetNewClosure()
    $b.Add_EnabledChanged($applyEnabledState)
    return $b
}

function global:New-LogBox {
    param([int]$X, [int]$Y, [int]$Width, [int]$Height)
    $l = New-Object System.Windows.Forms.TextBox
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($Width, $Height)
    $l.Multiline = $true
    $l.ScrollBars = 'Vertical'
    $l.Font = New-Object System.Drawing.Font('Consolas', 9)
    $l.ReadOnly = $true
    $l.BackColor = [System.Drawing.Color]::White
    $l.BorderStyle = 'FixedSingle'
    return $l
}

function global:Write-Log {
    param($Box, [string]$Message, $Color = $null)
    if ($Color -eq $colorDanger) {
        Write-Host $Message -ForegroundColor Red
    } elseif ($Color -eq $colorSuccess) {
        Write-Host $Message -ForegroundColor Green
    } elseif ($Color -eq $colorMuted) {
        Write-Host $Message -ForegroundColor DarkGray
    } else {
        Write-Host $Message
    }
    if ($Box) {
        try {
            $Box.AppendText("$Message`r`n")
            [System.Windows.Forms.Application]::DoEvents()
        } catch {}
    }
}

function global:Show-InfoBox {
    param([string]$Message, [string]$Title = 'Informação')
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', 'Information') | Out-Null
}

function global:Show-ErrorBox {
    param([string]$Message, [string]$Title = 'Erro')
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
}

function global:Show-Confirm {
    param([string]$Message, [string]$Title = 'Confirmar')
    return ([System.Windows.Forms.MessageBox]::Show($Message, $Title, 'YesNo', 'Question') -eq 'Yes')
}

# =========================================================================
# TELAS
# =========================================================================
function Show-StepWelcome {
    $global:currentStep = 1
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '👋 Bem-vindo!' -X 20 -Y 20 -FontSize 20 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Este assistente vai te guiar pela implantação da Biblioteca de Prompts.' -X 20 -Y 65 -Width 700))
    $content.Controls.Add((New-Label -Text 'Etapas:' -X 20 -Y 110 -Bold $true -FontSize 11))

    $steps = @(
        '1. Verificação de pré-requisitos (PowerShell + módulo PnP.PowerShell)',
        '2. Registro do App no Entra ID (gera o ClientId)',
        '3. Provisionamento das listas SharePoint',
        '4. Passo manual: subir .sppkg no App Catalog + instalar app no site',
        '5. Criação automática das páginas + web parts + layout tela cheia',
        '6. Conclusão'
    )
    $y = 140
    foreach ($s in $steps) {
        $content.Controls.Add((New-Label -Text $s -X 40 -Y $y -Width 660 -Color $colorMuted))
        $y += 26
    }

    $content.Controls.Add((New-Label -Text 'Clique em "Avançar >" para começar.' -X 20 -Y 320 -Color $colorMuted))

    $btnBack.Visible = $false
    $btnNext.Text = 'Avançar >'
    $btnNext.Tag = 'welcome'
}

function Show-StepPrerequisites {
    $global:currentStep = 2
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '✅ Verificação de pré-requisitos' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Vamos checar se sua máquina tem tudo o que precisa.' -X 20 -Y 50 -Width 700 -Color $colorMuted))

    $log = New-LogBox -X 20 -Y 90 -Width 700 -Height 280
    $content.Controls.Add($log)

    $btnCheck = New-Button -Text 'Verificar agora' -X 20 -Y 380 -Width 160 -Primary $true
    $btnInstallPnP = New-Button -Text 'Instalar PnP.PowerShell' -X 190 -Y 380 -Width 200
    $btnInstallPnP.Visible = $false
    $content.Controls.Add($btnCheck)
    $content.Controls.Add($btnInstallPnP)

    $global:prereqOk = $false

    $btnCheck.Add_Click({
        $log.Clear()
        $allOk = $true

        # PowerShell version
        $psv = $PSVersionTable.PSVersion
        if ($psv.Major -ge 7) {
            Write-Log $log ("✅ PowerShell $($psv.ToString())") $colorSuccess
        } else {
            Write-Log $log ("❌ PowerShell $($psv.ToString()) - precisa ser 7.0+") $colorDanger
            Write-Log $log 'Instale em: https://learn.microsoft.com/pt-br/powershell/scripting/install/installing-powershell-on-windows' $colorMuted
            $allOk = $false
        }

        # PnP.PowerShell module
        $pnp = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction SilentlyContinue
        if ($pnp) {
            Write-Log $log ("✅ Módulo PnP.PowerShell versão $($pnp[0].Version.ToString())") $colorSuccess
        } else {
            Write-Log $log '⚠️  Módulo PnP.PowerShell não instalado' $colorDanger
            Write-Log $log 'Clique em "Instalar PnP.PowerShell" para instalar automaticamente.' $colorMuted
            $btnInstallPnP.Visible = $true
            $allOk = $false
        }

        if ($allOk) {
            Write-Log $log ''
            Write-Log $log '🎉 Tudo pronto! Clique em Avançar.' $colorSuccess
            $global:prereqOk = $true
        }
    }.GetNewClosure())

    $btnInstallPnP.Add_Click({
        Write-Log $log ''
        Write-Log $log '📦 Instalando PnP.PowerShell (pode demorar 1-2 min)...' $colorMuted
        try {
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log $log '✅ Instalado com sucesso.' $colorSuccess
            $btnInstallPnP.Visible = $false
            # re-check
            $btnCheck.PerformClick()
        } catch {
            Write-Log $log "❌ Falha: $($_.Exception.Message)" $colorDanger
        }
    }.GetNewClosure())

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Tag = 'prereq'

    # Roda verificação depois que a UI acabar de renderizar (via Timer 100ms).
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    $timer.Add_Tick({
        $timer.Stop()
        $timer.Dispose()
        $btnCheck.PerformClick()
    }.GetNewClosure())
    $timer.Start()
}

function Show-StepAppRegistration {
    $global:currentStep = 3
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '🔐 Registrar App no Entra ID' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Preencha tenant e nome. Se ja tem o ClientId, cole no campo embaixo e clique Avancar.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $content.Controls.Add((New-Label -Text 'Tenant (so o prefixo, ex.: empresax):' -X 20 -Y 85))
    $txtTenant = New-TextInput -X 20 -Y 110 -Width 400 -Text $global:state.Tenant
    $content.Controls.Add($txtTenant)

    $content.Controls.Add((New-Label -Text 'Nome do App:' -X 20 -Y 145))
    $txtAppName = New-TextInput -X 20 -Y 170 -Width 400 -Text $global:state.AppName
    $content.Controls.Add($txtAppName)

    $content.Controls.Add((New-Label -Text 'Comando pronto (atualiza automaticamente):' -X 20 -Y 205 -Bold $true))

    $txtCommand = New-Object System.Windows.Forms.TextBox
    $txtCommand.Location = New-Object System.Drawing.Point(20, 230)
    $txtCommand.Size = New-Object System.Drawing.Size(700, 60)
    $txtCommand.Multiline = $true
    $txtCommand.ReadOnly = $true
    $txtCommand.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtCommand.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
    $txtCommand.ScrollBars = 'Vertical'
    $content.Controls.Add($txtCommand)

    $updateCmd = {
        $t = $txtTenant.Text.Trim()
        $n = $txtAppName.Text.Trim()
        if (-not $t) { $t = '<seu-tenant>' }
        if (-not $n) { $n = 'PnP PowerShell - Biblioteca de Prompts' }
        $txtCommand.Text = "Register-PnPEntraIDAppForInteractiveLogin -ApplicationName '$n' -Tenant '$t.onmicrosoft.com' -SharePointDelegatePermissions 'AllSites.FullControl' -GraphDelegatePermissions 'User.Read'"
    }.GetNewClosure()

    $txtTenant.Add_TextChanged($updateCmd)
    $txtAppName.Add_TextChanged($updateCmd)
    & $updateCmd

    $btnCopy = New-Button -Text '📋 Copiar comando' -X 20 -Y 300 -Width 180
    $btnOpenPs = New-Button -Text '🚀 Abrir PowerShell' -X 210 -Y 300 -Width 200 -Primary $true
    $content.Controls.Add($btnCopy)
    $content.Controls.Add($btnOpenPs)

    $lblInstr = New-Label -Text 'Fluxo: 1) Copiar comando  2) Abrir PowerShell  3) Colar e Enter  4) Volte aqui e cole o ClientId embaixo' -X 20 -Y 345 -Width 720 -Color $colorMuted
    $content.Controls.Add($lblInstr)

    $content.Controls.Add((New-Label -Text 'ClientId (cole aqui apos rodar o comando):' -X 20 -Y 375 -Bold $true))
    $txtClientId = New-TextInput -X 20 -Y 400 -Width 400 -Text $global:state.ClientId
    $content.Controls.Add($txtClientId)

    $content.Controls.Add((New-Label -Text 'Formato esperado: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (36 caracteres com hifens)' -X 20 -Y 430 -Width 600 -Color $colorMuted -FontSize 9))

    $lblClientIdStatus = New-Label -Text '' -X 430 -Y 405 -Width 290 -FontSize 9
    $content.Controls.Add($lblClientIdStatus)

    $updateClientIdStatus = {
        $v = $txtClientId.Text.Trim()
        if (-not $v) {
            $lblClientIdStatus.Text = ''
        } elseif (Test-GuidFormat $v) {
            $lblClientIdStatus.Text = "✓ Formato ok"
            $lblClientIdStatus.ForeColor = $colorSuccess
        } else {
            $lblClientIdStatus.Text = "✗ Formato invalido ($($v.Length) chars)"
            $lblClientIdStatus.ForeColor = $colorDanger
        }
    }.GetNewClosure()
    $txtClientId.Add_TextChanged($updateClientIdStatus)
    & $updateClientIdStatus

    $btnCopy.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtCommand.Text)
            Write-Host '✅ Comando copiado para a area de transferencia!' -ForegroundColor Green
        } catch {
            Write-Host "❌ Falha ao copiar: $($_.Exception.Message)" -ForegroundColor Red
        }
    }.GetNewClosure())

    $btnOpenPs.Add_Click({
        try {
            $wd = Split-Path -Parent $PSCommandPath
            if (-not $wd) { $wd = (Get-Location).Path }
            # Abre nova janela pwsh ja no diretorio do assistente
            Start-Process pwsh -ArgumentList '-NoExit', '-NoProfile', '-Command', "Set-Location '$wd'; Write-Host '📋 Cole o comando (Ctrl+V) e pressione Enter.' -ForegroundColor Cyan"
            Write-Host '🚀 Nova janela do PowerShell aberta.' -ForegroundColor Green
        } catch {
            Write-Host "❌ Falha ao abrir PowerShell: $($_.Exception.Message)" -ForegroundColor Red
        }
    }.GetNewClosure())

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Tag = 'appreg'
    $global:onNext = {
        $cid = $txtClientId.Text.Trim()
        if (-not (Test-GuidFormat $cid)) {
            Show-ErrorBox "ClientId invalido. O formato esperado e:`r`nxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (36 caracteres com hifens).`r`n`r`nValor atual: '$cid' ($($cid.Length) caracteres)."
            return $false
        }
        $global:state.Tenant   = $txtTenant.Text.Trim()
        $global:state.AppName  = $txtAppName.Text.Trim()
        $global:state.ClientId = $cid
        return $true
    }.GetNewClosure()
}

function Show-StepLists {
    $global:currentStep = 4
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '📋 Provisionar listas' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Cria as listas que faltam e atualiza colunas/config nas existentes.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $content.Controls.Add((New-Label -Text 'URL do site SharePoint:' -X 20 -Y 85))
    $suggested = if ($global:state.SiteUrl) { $global:state.SiteUrl } elseif ($global:state.Tenant) { "https://$($global:state.Tenant).sharepoint.com/sites/" } else { '' }
    $txtSiteUrl = New-TextInput -X 20 -Y 110 -Width 600 -Text $suggested
    $content.Controls.Add($txtSiteUrl)

    $content.Controls.Add((New-Label -Text 'Nome da lista de prompts:' -X 20 -Y 143))
    $txtPromptsList = New-TextInput -X 20 -Y 168 -Width 400 -Text $global:state.PromptsListTitle
    $content.Controls.Add($txtPromptsList)

    $content.Controls.Add((New-Label -Text 'Nome da lista de favoritos:' -X 20 -Y 201))
    $txtFavList = New-TextInput -X 20 -Y 226 -Width 400 -Text $global:state.FavoritesListTitle
    $content.Controls.Add($txtFavList)

    $btnProvision = New-Button -Text 'Verificar as listas' -X 20 -Y 260 -Width 200 -Primary $true
    $content.Controls.Add($btnProvision)

    $log = New-LogBox -X 20 -Y 300 -Width 700 -Height 170
    $content.Controls.Add($log)

    $lblStatus = New-Label -Text 'Aguardando verificação das listas...' -X 20 -Y 485 -Width 700 -Color $colorMuted -Bold $true
    $content.Controls.Add($lblStatus)

    $global:listsOk = $false

    $btnProvision.Add_Click({
        $global:state.SiteUrl = $txtSiteUrl.Text.Trim()
        $global:state.PromptsListTitle = $txtPromptsList.Text.Trim()
        $global:state.FavoritesListTitle = $txtFavList.Text.Trim()
        if (-not $global:state.SiteUrl -or -not $global:state.ClientId) {
            Show-ErrorBox 'Preencha URL do site e ClientId (etapa anterior).'
            return
        }
        $log.Clear()
        Write-Log $log '🔌 Conectando ao SharePoint...' $colorMuted
        try {
            Import-Module PnP.PowerShell -ErrorAction Stop
            Connect-PnPOnline -Url $global:state.SiteUrl -Interactive -ClientId $global:state.ClientId -ErrorAction Stop
            Write-Log $log '✅ Conectado.' $colorSuccess
            Invoke-ListProvisioning -Log $log
            Write-Log $log ''
            Write-Log $log '🎉 Listas verificadas! Pode avançar.' $colorSuccess
            $global:listsOk = $true
            $lblStatus.Text = '✅ Listas verificadas — pode avançar'
            $lblStatus.ForeColor = $colorSuccess
            $btnNext.Enabled = $true
        } catch {
            Write-Log $log "❌ Falha: $($_.Exception.Message)" $colorDanger
            $lblStatus.Text = '❌ Falha na verificação - veja o log acima'
            $lblStatus.ForeColor = $colorDanger
        } finally {
            try { Disconnect-PnPOnline } catch {}
        }
    }.GetNewClosure())

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Enabled = $false
    $btnNext.Tag = 'lists'
    $global:onNext = $null
}

function Show-StepManualUpload {
    $global:currentStep = 5
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '📦 Passo manual: subir .sppkg e instalar no site' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Faça os 5 passos abaixo. As páginas serão criadas automaticamente na próxima etapa.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $instr = @'
1. Localize o arquivo sharepoint-biblioteca-prompt.sppkg (nesta pasta, se veio zipado, ou baixe do GitHub)
2. Vá em: https://<seu-tenant>-admin.sharepoint.com → SharePoint → Apps → App Catalog
3. Apps for SharePoint → Upload → escolha o .sppkg → Deploy
4. Vá no site alvo → engrenagem → Add an app → procure "sharepoint-biblioteca-prompt-webpart" → Add
5. Aguarde alguns segundos até o app aparecer em Site Contents

Dica: se o app não aparecer para adicionar no site (passo 4), volte no passo 3 e marque
"Make this solution available to all sites in the organization" antes de clicar em Deploy.
'@
    $txtInstr = New-Object System.Windows.Forms.TextBox
    $txtInstr.Multiline = $true
    $txtInstr.ScrollBars = 'Vertical'
    $txtInstr.ReadOnly = $true
    $txtInstr.Text = $instr
    $txtInstr.Location = New-Object System.Drawing.Point(20, 90)
    $txtInstr.Size = New-Object System.Drawing.Size(700, 155)
    $txtInstr.Font = New-UiFont -Size 9
    $content.Controls.Add($txtInstr)

    # Detecta .sppkg ao lado do assistente (pasta implantacao/ empacotada)
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $localSppkg = $null
    if ($scriptDir) {
        $candidato = Get-ChildItem -Path $scriptDir -Filter '*.sppkg' -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($candidato) { $localSppkg = $candidato.FullName }
    }

    if ($localSppkg) {
        $lblLocal = New-Label -Text "📦 .sppkg detectado nesta pasta: $(Split-Path $localSppkg -Leaf)" -X 20 -Y 255 -Width 500 -Color $colorSuccess
        $content.Controls.Add($lblLocal)
        $btnOpenFolder = New-Button -Text '📂 Abrir pasta' -X 530 -Y 252 -Width 130
        $btnOpenFolder.Add_Click({
            Start-Process 'explorer.exe' -ArgumentList "/select,`"$localSppkg`""
        }.GetNewClosure())
        $content.Controls.Add($btnOpenFolder)
    }

    $lblWarn = New-Label -Text '⚠ Não avance sem completar TODOS os 5 passos acima. Depois clique em "Já fiz o upload".' -X 20 -Y 300 -Width 700 -Color $colorDanger -Bold $true -FontSize 10
    $content.Controls.Add($lblWarn)

    $btnConfirm = New-Button -Text '✓ Já fiz o upload' -X 20 -Y 340 -Width 200 -Primary $true
    $content.Controls.Add($btnConfirm)

    $lblStatus = New-Label -Text 'Aguardando confirmação...' -X 240 -Y 345 -Width 480 -Color $colorMuted
    $content.Controls.Add($lblStatus)

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Enabled = $false
    $global:onNext = $null

    $btnConfirm.Add_Click({
        $lblStatus.Text = '✅ Confirmado — pode avançar'
        $lblStatus.ForeColor = $colorSuccess
        $btnNext.Enabled = $true
        $btnConfirm.Enabled = $false
    }.GetNewClosure())
}

function Show-StepConfigure {
    $global:currentStep = 6
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '🎨 Criar páginas e configurar' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Vou criar as páginas, inserir as web parts, aplicar layout tela cheia e desligar a barra social.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $content.Controls.Add((New-Label -Text 'Nome da página da Biblioteca a ser criada (Ex.: Biblioteca-de-Prompts):' -X 20 -Y 88 -Width 700))
    $bibDefault = if ($global:state.BibliotecaPageName) { ($global:state.BibliotecaPageName -replace '\.aspx$','') } else { 'Biblioteca-de-Prompts' }
    $txtBib = New-TextInput -X 20 -Y 113 -Width 400 -Text $bibDefault
    $content.Controls.Add($txtBib)

    $content.Controls.Add((New-Label -Text 'Nome da página do Dashboard a ser criada (Ex.: Dashboard-Biblioteca-de-Prompts) - opcional:' -X 20 -Y 148 -Width 700))
    $dashDefault = if ($global:state.DashboardPageName) { ($global:state.DashboardPageName -replace '\.aspx$','') } else { 'Dashboard-Biblioteca-de-Prompts' }
    $txtDash = New-TextInput -X 20 -Y 173 -Width 400 -Text $dashDefault
    $content.Controls.Add($txtDash)

    $content.Controls.Add((New-Label -Text 'Recomendado usar hífen entre as palavras. Pode ter espaço, mas fica melhor sem.' -X 20 -Y 203 -Width 700 -Color $colorMuted -FontSize 9))

    $btnRun = New-Button -Text 'Criar e configurar agora' -X 20 -Y 230 -Width 220 -Primary $true
    $content.Controls.Add($btnRun)

    $log = New-LogBox -X 20 -Y 275 -Width 700 -Height 260
    $content.Controls.Add($log)

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Enabled = $false
    $global:onNext = $null

    $btnRun.Add_Click({
        $bibName = ($txtBib.Text.Trim() -replace '\.aspx$','')
        $dashName = ($txtDash.Text.Trim() -replace '\.aspx$','')
        if (-not $bibName) {
            Show-ErrorBox 'Informe o nome da página da Biblioteca.'
            return
        }
        $global:state.BibliotecaPageName = "$bibName.aspx"
        $global:state.DashboardPageName = if ($dashName) { "$dashName.aspx" } else { '' }

        $log.Clear()
        Write-Log $log '🔌 Conectando (rodando em thread separada para não travar a tela)...' $colorMuted
        Write-Log $log '   Uma janela do browser pode abrir para login - complete e volte aqui.' $colorMuted
        Write-Log $log ''

        try {
            $siteUrl  = $global:state.SiteUrl
            $clientId = $global:state.ClientId
            $bibPage  = $bibName
            $dashPage = $dashName
            $bibWpId  = 'c8e4a1f2-7b3d-4e9a-8f5c-6d2b1a9e3f4c'
            $dashWpId = 'd3f7a2e4-6b91-4c8f-a5d2-1c9e4b7f3a8b'

            $bibProps = @{
                targetListTitle    = $global:state.PromptsListTitle
                favoritesListTitle = $global:state.FavoritesListTitle
                promptIdField      = 'PromptID'
                copyFields         = 'acao,Prompt,Segmento,Categoria,Funcionacom'
                extraToolColors    = ''
            } | ConvertTo-Json -Compress

            $dashProps = @{
                targetListTitle    = $global:state.PromptsListTitle
                favoritesListTitle = $global:state.FavoritesListTitle
                promptIdField      = 'PromptID'
                topN               = 10
            } | ConvertTo-Json -Compress

            $job = Start-ThreadJob -ScriptBlock {
                param($SiteUrl, $ClientId, $BibPage, $DashPage, $BibWpId, $DashWpId, $BibProps, $DashProps)
                Import-Module PnP.PowerShell
                Write-Output ">> Conectando ao SharePoint..."
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
                Write-Output ">> Conectado."
                Write-Output ""

                function Setup-Page {
                    param([string]$PageName, [string]$WebPartId, [string]$Label, [string]$PropsJson)
                    $file = "$PageName.aspx"
                    $existing = Get-PnPPage -Identity $file -ErrorAction SilentlyContinue
                    if ($existing) {
                        Write-Output "  Pagina '$PageName' ja existe - pulando (delete manualmente para recriar)."
                        Write-Output ""
                        return $false
                    }
                    Write-Output "  [1/7] Criando pagina '$PageName' em branco..."
                    Add-PnPPage -Name $PageName -ErrorAction Stop | Out-Null

                    Write-Output "  [2/7] Buscando web part '$Label' entre os componentes do site..."
                    $comps = Get-PnPAvailablePageComponents -Page $file -ErrorAction SilentlyContinue
                    $wpIdNorm = $WebPartId.ToLower()
                    $comp = $comps | Where-Object {
                        $cid = ($_.Id.ToString() -replace '[{}]','').ToLower()
                        $cid -eq $wpIdNorm -or $_.Name -eq $Label
                    } | Select-Object -First 1

                    if (-not $comp) {
                        Write-Output "  AVISO: web part '$Label' NAO encontrada."
                        Write-Output "  Verifique se o app foi instalado no site (passo 5)."
                        Write-Output ""

                        Write-Output "  -- Apps instalados neste site --"
                        try {
                            $apps = Get-PnPApp -Scope Site -ErrorAction SilentlyContinue
                            if ($apps) {
                                foreach ($a in $apps) {
                                    Write-Output "    Id=$($a.Id)  Title='$($a.Title)'  Deployed=$($a.Deployed)"
                                }
                            } else {
                                Write-Output "    (nenhum)"
                            }
                        } catch {
                            Write-Output "    Falha ao listar apps: $($_.Exception.Message)"
                        }
                        Write-Output ""
                        Write-Output "  >> Adicione a web part manualmente na pagina '$file'."
                        Write-Output ""
                        return $false
                    }

                    Write-Output "  [3/7] Componente encontrado: $($comp.Name) [$($comp.Id)]"

                    Write-Output "  [4/7] Adicionando web part na pagina..."
                    Add-PnPPageWebPart -Page $file -Component $comp -ErrorAction Stop | Out-Null

                    Write-Output "  [5/7] Configurando propriedades da web part..."
                    try {
                        $onPage = Get-PnPPageComponent -Page $file -ErrorAction Stop
                        $ourInstance = $onPage | Where-Object {
                            $wid = ($_.WebPartId.ToString() -replace '[{}]','').ToLower()
                            $wid -eq $wpIdNorm
                        } | Select-Object -First 1
                        if ($ourInstance) {
                            Set-PnPPageWebPart -Page $file -Identity $ourInstance.InstanceId -PropertiesJson $PropsJson -ErrorAction Stop | Out-Null
                            Write-Output "     Propriedades salvas (instance $($ourInstance.InstanceId))."
                        } else {
                            Write-Output "     ATENCAO: nao achei instancia da web part para configurar propriedades."
                        }
                    } catch {
                        Write-Output "     ATENCAO ao configurar propriedades: $($_.Exception.Message)"
                    }

                    Write-Output "  [6/7] Aplicando layout SingleWebPartAppPage (tela cheia)..."
                    try {
                        Set-PnPPage -Identity $file -LayoutType SingleWebPartAppPage -ErrorAction Stop | Out-Null
                        Write-Output "     Layout aplicado."
                    } catch {
                        Write-Output "     ATENCAO ao aplicar layout: $($_.Exception.Message)"
                    }

                    Write-Output "  [7/7] Publicando pagina..."
                    Set-PnPPage -Identity $file -Publish -ErrorAction Stop | Out-Null

                    Write-Output "  Pagina '$PageName' pronta e configurada."
                    Write-Output ""
                    return $true
                }

                Write-Output ">> [1] Pagina da Biblioteca: '$BibPage'"
                Setup-Page -PageName $BibPage -WebPartId $BibWpId -Label 'Biblioteca de Prompts' -PropsJson $BibProps

                if ($DashPage) {
                    Write-Output ">> [2] Pagina do Dashboard: '$DashPage'"
                    Setup-Page -PageName $DashPage -WebPartId $DashWpId -Label 'Dashboard — Biblioteca de Prompts' -PropsJson $DashProps
                } else {
                    Write-Output ">> [2] Dashboard pulado (nao informado)."
                    Write-Output ""
                }

                Write-Output ">> [3] Desligando barra social do site..."
                Set-PnPSite -SocialBarOnSitePagesDisabled $true
                Write-Output "  Barra social desligada."
                Write-Output ""
                Write-Output ">> Tudo pronto."

                try { Disconnect-PnPOnline } catch {}
            } -ArgumentList $siteUrl, $clientId, $bibPage, $dashPage, $bibWpId, $dashWpId, $bibProps, $dashProps

            # Poll com streaming em tempo real
            while ($job.State -in 'Running','NotStarted') {
                $partial = Receive-Job -Job $job -ErrorAction SilentlyContinue
                if ($partial) {
                    foreach ($line in $partial) {
                        Write-Log $log ($line.ToString())
                    }
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 300
            }
            # Drena o que sobrou
            $final = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($final) {
                foreach ($line in $final) {
                    Write-Log $log ($line.ToString())
                }
            }
            Remove-Job -Job $job -Force

            if ($job.State -eq 'Completed') {
                Write-Log $log ''
                Write-Log $log '🎉 Tudo configurado! Clique em Avançar.' $colorSuccess
                $btnNext.Enabled = $true
            } else {
                Write-Log $log "❌ Job terminou com estado: $($job.State)" $colorDanger
            }
        } catch {
            Write-Log $log "❌ Falha: $($_.Exception.Message)" $colorDanger
        }
    }.GetNewClosure())
}

function Show-StepDone {
    $global:currentStep = 7
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '🎉 Concluído!' -X 20 -Y 20 -FontSize 22 -Bold $true -Color $colorSuccess))
    $content.Controls.Add((New-Label -Text 'A Biblioteca de Prompts está pronta pra uso.' -X 20 -Y 70 -Width 720))

    $sum = "Resumo:`r`n"
    $sum += "  Site: $($global:state.SiteUrl)`r`n"
    $sum += "  Lista principal: $($global:state.PromptsListTitle)`r`n"
    $sum += "  Lista favoritos: $($global:state.FavoritesListTitle)`r`n"
    $sum += "  Página Biblioteca: $($global:state.BibliotecaPageName)`r`n"
    if ($global:state.DashboardPageName) { $sum += "  Página Dashboard: $($global:state.DashboardPageName)`r`n" }
    $sum += "  ClientId (guardar!): $($global:state.ClientId)"

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.Location = New-Object System.Drawing.Point(20, 110)
    $txt.Size = New-Object System.Drawing.Size(700, 200)
    $txt.Font = New-Object System.Drawing.Font('Consolas', 10)
    $txt.Text = $sum
    $content.Controls.Add($txt)

    $btnOpen = New-Button -Text 'Abrir Biblioteca no browser' -X 20 -Y 330 -Width 220 -Primary $true
    $btnOpen.Add_Click({
        $url = "$($global:state.SiteUrl)/SitePages/$($global:state.BibliotecaPageName)"
        Start-Process $url
    }.GetNewClosure())
    $content.Controls.Add($btnOpen)

    $btnBack.Visible = $false
    $btnNext.Text = 'Fechar'
    $btnNext.Tag = 'done'
    $btnNext.Enabled = $true
}

# =========================================================================
# LÓGICA DE PROVISIONAMENTO (usada na etapa Listas)
# =========================================================================
function global:Invoke-ListProvisioning {
    param($Log)

    $ACOES        = @('Analisar','Perguntar','Resumir','Criar','Encontrar','Aprender','Otimizar','Se preparar','Entender')
    $SEGMENTOS    = @('Comercial','DP','Financeiro','Infra','Projetos','RH','Analista Funcional','Desenvolvedor(a)','Gerente de Projetos')
    $CATEGORIAS   = @('Área','Função')
    $FUNCIONA_COM = @('Outlook','Teams','OneNote','Word','Excel','PowerPoint','Power BI','M365 Copilot','Copilot Studio','D365 CCaaS / Customer Service','D365 Customer Insights - Journeys','D365 Sales','Fabric','Power Apps','Power Automate','Power Pages','Whiteboard')

    function ChoiceXml { param([string]$Name, [string]$DisplayName, [string[]]$Choices, [bool]$Required = $false)
        $esc = $Choices | ForEach-Object { [System.Security.SecurityElement]::Escape($_) }
        $cxml = ($esc | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
        $req = if ($Required) { "Required='TRUE'" } else { '' }
        return "<Field Type='Choice' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Format='Dropdown' $req><CHOICES>$cxml</CHOICES></Field>"
    }

    function EnsureList { param([string]$Title, $L)
        $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
        if (-not $list) {
            Write-Log $L "  ➕ Criando lista '$Title'..." $colorMuted
            $list = New-PnPList -Title $Title -Template GenericList -EnableVersioning
        } else {
            Write-Log $L "  ✓ Lista '$Title' já existe" $colorMuted
        }
    }
    function EnsureField { param([string]$List, [string]$Name, [string]$Xml, $L)
        $ex = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
        if (-not $ex) {
            Write-Log $L "    ➕ Campo '$Name'" $colorMuted
            Add-PnPFieldFromXml -List $List -FieldXml $Xml | Out-Null
        } else {
            Write-Log $L "    ✓ Campo '$Name' já existe" $colorMuted
        }
    }
    function SetDefaultViewFields { param([string]$List, [string[]]$Fields, $L)
        try {
            $view = Get-PnPView -List $List -ErrorAction SilentlyContinue | Where-Object { $_.DefaultView } | Select-Object -First 1
            if ($view) {
                Set-PnPView -List $List -Identity $view.Id -Fields $Fields | Out-Null
                Write-Log $L "    ✓ View padrão atualizada: $($Fields -join ', ')" $colorMuted
            }
        } catch {
            Write-Log $L "    ⚠ Não foi possível atualizar view padrão: $($_.Exception.Message)" $colorMuted
        }
    }

    $xAcao      = ChoiceXml 'acao' 'Ação' $ACOES $true
    $xPrompt    = "<Field Type='Note' Name='Prompt' StaticName='Prompt' DisplayName='Prompt' RichText='FALSE' NumLines='6' Required='TRUE' />"
    $xSegmento  = ChoiceXml 'Segmento' 'Segmento' $SEGMENTOS $true
    $xCategoria = ChoiceXml 'Categoria' 'Categoria' $CATEGORIAS $true
    $xFuncCom   = ChoiceXml 'Funcionacom' 'Funciona com' $FUNCIONA_COM $true
    $xAtivo     = "<Field Type='Boolean' Name='Ativo' StaticName='Ativo' DisplayName='Ativo'><Default>1</Default></Field>"
    $xPromptId  = "<Field Type='Number' Name='PromptID' StaticName='PromptID' DisplayName='PromptID' Required='TRUE' />"

    Write-Log $Log ('[1/2] Lista principal: ' + $global:state.PromptsListTitle) $colorMuted
    EnsureList $global:state.PromptsListTitle $Log
    EnsureField $global:state.PromptsListTitle 'acao' $xAcao $Log
    EnsureField $global:state.PromptsListTitle 'Prompt' $xPrompt $Log
    EnsureField $global:state.PromptsListTitle 'Segmento' $xSegmento $Log
    EnsureField $global:state.PromptsListTitle 'Categoria' $xCategoria $Log
    EnsureField $global:state.PromptsListTitle 'Funcionacom' $xFuncCom $Log
    EnsureField $global:state.PromptsListTitle 'Ativo' $xAtivo $Log
    SetDefaultViewFields $global:state.PromptsListTitle @('Title','acao','Segmento','Categoria','Funcionacom','Ativo','Editor','Modified') $Log

    Write-Log $Log ''
    Write-Log $Log ('[2/2] Lista de favoritos: ' + $global:state.FavoritesListTitle) $colorMuted
    EnsureList $global:state.FavoritesListTitle $Log
    EnsureField $global:state.FavoritesListTitle 'PromptID' $xPromptId $Log
    EnsureField $global:state.FavoritesListTitle 'acao' $xAcao $Log
    EnsureField $global:state.FavoritesListTitle 'Prompt' $xPrompt $Log
    EnsureField $global:state.FavoritesListTitle 'Segmento' $xSegmento $Log
    EnsureField $global:state.FavoritesListTitle 'Categoria' $xCategoria $Log
    EnsureField $global:state.FavoritesListTitle 'Funcionacom' $xFuncCom $Log
    SetDefaultViewFields $global:state.FavoritesListTitle @('Title','PromptID','acao','Segmento','Categoria','Funcionacom','Author','Created') $Log
}

# =========================================================================
# NAVEGAÇÃO
# =========================================================================
function Show-Current {
    $global:onNext = $null  # reset callback do passo
    switch ($global:currentStep) {
        1 { Show-StepWelcome }
        2 { Show-StepPrerequisites }
        3 { Show-StepAppRegistration }
        4 { Show-StepLists }
        5 { Show-StepManualUpload }
        6 { Show-StepConfigure }
        7 { Show-StepDone }
    }
}

function Update-Header {
    $lblStep.Text = "Etapa $($global:currentStep) de $global:totalSteps"
    $progressBar.Value = [int](($global:currentStep / $global:totalSteps) * 100)
    $btnNext.Enabled = $true
}

# =========================================================================
# FORMULÁRIO
# =========================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Biblioteca de Prompts — Assistente de Implantação'
$form.Size = New-Object System.Drawing.Size(770, 740)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-UiFont -Size 9

# Header (com gradiente roxo)
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 70
$header.BackColor = $colorPrimary
$header.Add_Paint({
    param($sender, $e)
    $rect = $sender.ClientRectangle
    if ($rect.Width -le 0 -or $rect.Height -le 0) { return }
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $colorGradientStart, $colorGradientEnd, 0)
    $e.Graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
})

$title = New-Object System.Windows.Forms.Label
$title.Text = '📚 Biblioteca de Prompts'
$title.Font = New-UiFont -FamilyOverride (Get-UiFontName $title.Text) -Size 15 -Style ([System.Drawing.FontStyle]::Bold)
$title.ForeColor = $colorHeaderText
$title.BackColor = [System.Drawing.Color]::Transparent
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20, 12)
$header.Controls.Add($title)

$lblStep = New-Object System.Windows.Forms.Label
$lblStep.Text = 'Etapa 1 de 7'
$lblStep.Font = New-UiFont -Size 9
$lblStep.ForeColor = $colorHeaderText
$lblStep.BackColor = [System.Drawing.Color]::Transparent
$lblStep.AutoSize = $true
$lblStep.Location = New-Object System.Drawing.Point(20, 42)
$header.Controls.Add($lblStep)

# Logo Bizapp (canto superior direito, se o arquivo existir)
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$logoPath = $null
if ($scriptRoot) {
    $candidates = @(
        (Join-Path $scriptRoot 'assets\logo-bizapp.png'),
        (Join-Path $scriptRoot 'logo-bizapp.png')
    )
    foreach ($p in $candidates) { if (Test-Path $p) { $logoPath = $p; break } }
}
if ($logoPath) {
    try {
        $logoImg = [System.Drawing.Image]::FromFile($logoPath)
        $picLogo = New-Object System.Windows.Forms.PictureBox
        $picLogo.Image = $logoImg
        $picLogo.SizeMode = 'Zoom'
        $picLogo.BackColor = [System.Drawing.Color]::Transparent
        # Encaixa 130x50 no canto direito, com margem
        $picLogo.Size = New-Object System.Drawing.Size(130, 50)
        $picLogo.Location = New-Object System.Drawing.Point(($header.Width - 150), 10)
        $picLogo.Anchor = 'Top,Right'
        $header.Controls.Add($picLogo)
    } catch {
        Write-Host "Nao consegui carregar o logo: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

$form.Controls.Add($header)

# Progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = 'Top'
$progressBar.Height = 6
$progressBar.Style = 'Continuous'
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 14
$form.Controls.Add($progressBar)

# Footer
$footer = New-Object System.Windows.Forms.Panel
$footer.Dock = 'Bottom'
$footer.Height = 75
$footer.BackColor = $colorBg

$btnNext = New-Button -Text 'Avançar >' -X 620 -Y 15 -Width 120 -Primary $true
$btnBack = New-Button -Text '< Voltar' -X 490 -Y 15 -Width 120

$lblCopyright = New-Object System.Windows.Forms.Label
$lblCopyright.Text = '© 2026 Bizapp (bizappcrm.com). Todos os direitos reservados.'
$lblCopyright.Font = New-UiFont -Size 8
$lblCopyright.ForeColor = $colorMuted
$lblCopyright.AutoSize = $true
$lblCopyright.Location = New-Object System.Drawing.Point(20, 55)
$footer.Controls.Add($lblCopyright)

$footer.Controls.Add($btnNext)
$footer.Controls.Add($btnBack)
$form.Controls.Add($footer)

# Content
$content = New-Object System.Windows.Forms.Panel
$content.Dock = 'Fill'
$content.BackColor = $colorBg
$form.Controls.Add($content)

# --- ordem correta pra Dock ---
$form.Controls.SetChildIndex($footer, 0)
$form.Controls.SetChildIndex($content, 1)
$form.Controls.SetChildIndex($progressBar, 2)
$form.Controls.SetChildIndex($header, 3)

# Navegação
$btnNext.Add_Click({
    # Executa callback do passo atual. Se retornar $false, cancela a navegação.
    if ($global:onNext) {
        $goAhead = $true
        try {
            $r = & $global:onNext
            if ($r -is [bool] -and -not $r) { $goAhead = $false }
        } catch { }
        if (-not $goAhead) { return }
        $global:onNext = $null
    }
    if ($btnNext.Tag -eq 'done') { $form.Close(); return }
    if ($global:currentStep -lt $global:totalSteps) {
        $global:currentStep++
        Show-Current
    }
})
$btnBack.Add_Click({
    if ($global:currentStep -gt 1) {
        $global:currentStep--
        Show-Current
    }
})

# Kick off
Show-Current
[void]$form.ShowDialog()

