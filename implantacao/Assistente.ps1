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
      5. Passo manual: subir .sppkg no App Catalog + criar paginas
      6. Configuracao das paginas (layout tela cheia + social bar desligada)
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

# Paleta
$colorPrimary   = [System.Drawing.Color]::FromArgb(0, 120, 212)
$colorPrimaryH  = [System.Drawing.Color]::FromArgb(16, 110, 190)
$colorBg        = [System.Drawing.Color]::FromArgb(250, 249, 248)
$colorText      = [System.Drawing.Color]::FromArgb(50, 49, 48)
$colorMuted     = [System.Drawing.Color]::FromArgb(96, 94, 92)
$colorSuccess   = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorDanger    = [System.Drawing.Color]::FromArgb(164, 38, 44)

# =========================================================================
# HELPERS
# =========================================================================
function global:New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 680, [int]$FontSize = 10, [bool]$Bold = $false, [System.Drawing.Color]$Color = $colorText)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.AutoSize = ($Width -eq 0)
    $h = [int]($FontSize * 1.8) + 8
    if ($Width -gt 0) { $l.Size = New-Object System.Drawing.Size($Width, $h) }
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $l.Font = New-Object System.Drawing.Font('Segoe UI', $FontSize, $style)
    $l.ForeColor = $Color
    return $l
}

function global:New-TextInput {
    param([int]$X, [int]$Y, [int]$Width = 400, [string]$Text = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($Width, 24)
    $t.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $t.Text = $Text
    return $t
}

function global:New-Button {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 140, [bool]$Primary = $false)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($Width, 32)
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $b.FlatStyle = 'Flat'
    if ($Primary) {
        $b.BackColor = $colorPrimary
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderSize = 0
    } else {
        $b.BackColor = [System.Drawing.Color]::White
        $b.ForeColor = $colorText
        $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 198, 196)
    }
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
        '4. Passo manual: você sobe o .sppkg + cria as páginas',
        '5. Configuração das páginas (tela cheia + barra social desligada)',
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
        $global:state.Tenant   = $txtTenant.Text.Trim()
        $global:state.AppName  = $txtAppName.Text.Trim()
        $global:state.ClientId = $txtClientId.Text.Trim()
    }.GetNewClosure()
}

function Show-StepLists {
    $global:currentStep = 4
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '📋 Provisionar listas' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Preencha os dados do site alvo. As listas serão criadas com as colunas corretas.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $content.Controls.Add((New-Label -Text 'URL do site SharePoint:' -X 20 -Y 90))
    $suggested = if ($global:state.SiteUrl) { $global:state.SiteUrl } elseif ($global:state.Tenant) { "https://$($global:state.Tenant).sharepoint.com/sites/" } else { '' }
    $txtSiteUrl = New-TextInput -X 20 -Y 115 -Width 600 -Text $suggested
    $content.Controls.Add($txtSiteUrl)

    $content.Controls.Add((New-Label -Text 'Nome da lista de prompts:' -X 20 -Y 155))
    $txtPromptsList = New-TextInput -X 20 -Y 180 -Width 400 -Text $global:state.PromptsListTitle
    $content.Controls.Add($txtPromptsList)

    $content.Controls.Add((New-Label -Text 'Nome da lista de favoritos:' -X 20 -Y 220))
    $txtFavList = New-TextInput -X 20 -Y 245 -Width 400 -Text $global:state.FavoritesListTitle
    $content.Controls.Add($txtFavList)

    $btnProvision = New-Button -Text 'Criar listas agora' -X 20 -Y 285 -Width 200 -Primary $true
    $content.Controls.Add($btnProvision)

    $log = New-LogBox -X 20 -Y 330 -Width 700 -Height 100
    $content.Controls.Add($log)

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
            Write-Log $log '🎉 Listas prontas! Clique em Avançar.' $colorSuccess
            $global:listsOk = $true
        } catch {
            Write-Log $log "❌ Falha: $($_.Exception.Message)" $colorDanger
        } finally {
            try { Disconnect-PnPOnline } catch {}
        }
    }.GetNewClosure())

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Tag = 'lists'
}

function Show-StepManualUpload {
    $global:currentStep = 5
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '📦 Passo manual: subir .sppkg e criar página' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Faça os passos abaixo no browser. Quando terminar, preencha os nomes das páginas que você criou.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $instr = @'
1. Localize o arquivo sharepoint-biblioteca-prompt.sppkg (nesta pasta, se veio zipado, ou baixe do GitHub / compile localmente)
2. Vá em: https://<seu-tenant>-admin.sharepoint.com → SharePoint → Apps → App Catalog
3. Apps for SharePoint → Upload → escolha o .sppkg → NÃO marque "make available to all sites" → Deploy
4. Vá no site alvo → engrenagem ⚙️ → Add an app → procure "sharepoint-biblioteca-prompt-webpart" → Add
5. + Novo → Página em branco → Nome: "Biblioteca de Prompts" (ou o que preferir)
6. + no meio → busca "Biblioteca" → adicione a web part → confirme os títulos das listas no painel de propriedades → Publicar
7. (Opcional) Repita passos 5-6 para uma segunda página com a web part "Dashboard - Biblioteca de Prompts"
'@
    $txtInstr = New-Object System.Windows.Forms.TextBox
    $txtInstr.Multiline = $true
    $txtInstr.ScrollBars = 'Vertical'
    $txtInstr.ReadOnly = $true
    $txtInstr.Text = $instr
    $txtInstr.Location = New-Object System.Drawing.Point(20, 90)
    $txtInstr.Size = New-Object System.Drawing.Size(700, 180)
    $txtInstr.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $content.Controls.Add($txtInstr)

    # Detecta .sppkg ao lado do assistente (pasta implantacao/ empacotada)
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $localSppkg = $null
    if ($scriptDir) {
        $candidato = Get-ChildItem -Path $scriptDir -Filter '*.sppkg' -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($candidato) { $localSppkg = $candidato.FullName }
    }

    if ($localSppkg) {
        $lblLocal = New-Label -Text "📦 .sppkg detectado nesta pasta: $(Split-Path $localSppkg -Leaf)" -X 20 -Y 275 -Width 500 -Color $colorSuccess
        $content.Controls.Add($lblLocal)
        $btnOpenFolder = New-Button -Text '📂 Abrir pasta' -X 530 -Y 272 -Width 130
        $btnOpenFolder.Add_Click({
            Start-Process 'explorer.exe' -ArgumentList "/select,`"$localSppkg`""
        }.GetNewClosure())
        $content.Controls.Add($btnOpenFolder)
        $yLbl1 = 305; $yInp1 = 330; $yLbl2 = 370; $yInp2 = 395
    } else {
        $yLbl1 = 285; $yInp1 = 310; $yLbl2 = 350; $yInp2 = 375
    }

    $content.Controls.Add((New-Label -Text 'Nome do arquivo .aspx da página da Biblioteca (ex.: Biblioteca-de-Prompts.aspx):' -X 20 -Y $yLbl1 -Width 700))
    $txtBibliotecaPage = New-TextInput -X 20 -Y $yInp1 -Width 400 -Text $global:state.BibliotecaPageName
    $content.Controls.Add($txtBibliotecaPage)

    $content.Controls.Add((New-Label -Text 'Nome do arquivo .aspx da página do Dashboard (opcional):' -X 20 -Y $yLbl2 -Width 700))
    $txtDashboardPage = New-TextInput -X 20 -Y $yInp2 -Width 400 -Text $global:state.DashboardPageName
    $content.Controls.Add($txtDashboardPage)

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
    $btnNext.Enabled = $true
    $btnNext.Tag = 'manual'
    $global:onNext = {
        $global:state.BibliotecaPageName = $txtBibliotecaPage.Text.Trim()
        $global:state.DashboardPageName = $txtDashboardPage.Text.Trim()
    }.GetNewClosure()
}

function Show-StepConfigure {
    $global:currentStep = 6
    Update-Header
    $content.Controls.Clear()

    $content.Controls.Add((New-Label -Text '🎨 Configurar páginas' -X 20 -Y 15 -FontSize 16 -Bold $true))
    $content.Controls.Add((New-Label -Text 'Aplicando layout de tela cheia e desligando a barra social do site.' -X 20 -Y 50 -Width 720 -Color $colorMuted))

    $summary = "Site: $($global:state.SiteUrl)`r`nBiblioteca: $($global:state.BibliotecaPageName)"
    if ($global:state.DashboardPageName) { $summary += "`r`nDashboard: $($global:state.DashboardPageName)" }
    $lblSummary = New-Label -Text $summary -X 20 -Y 90 -Width 0
    $lblSummary.MaximumSize = New-Object System.Drawing.Size(700, 0)
    $content.Controls.Add($lblSummary)

    $btnRun = New-Button -Text 'Configurar agora' -X 20 -Y 180 -Width 180 -Primary $true
    $content.Controls.Add($btnRun)

    $log = New-LogBox -X 20 -Y 220 -Width 700 -Height 210
    $content.Controls.Add($log)

    $btnRun.Add_Click({
        if (-not $global:state.BibliotecaPageName) {
            Show-ErrorBox 'Informe o nome da página da Biblioteca (etapa anterior).'
            return
        }
        $log.Clear()
        Write-Log $log '🔌 Conectando em thread separada. Uma janela do browser pode abrir para login - complete e volte aqui.' $colorMuted
        try {
            $siteUrl  = $global:state.SiteUrl
            $clientId = $global:state.ClientId
            $bibPage  = $global:state.BibliotecaPageName
            $dashPage = $global:state.DashboardPageName

            $job = Start-ThreadJob -ScriptBlock {
                param($SiteUrl, $ClientId, $BibPage, $DashPage)
                Import-Module PnP.PowerShell
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop

                Write-Output "Configurando pagina '$BibPage'..."
                Set-PnPPage -Identity $BibPage -LayoutType SingleWebPartAppPage | Out-Null
                Set-PnPPage -Identity $BibPage -Publish | Out-Null
                Write-Output "Pagina da Biblioteca OK."

                if ($DashPage) {
                    Write-Output "Configurando pagina '$DashPage'..."
                    Set-PnPPage -Identity $DashPage -LayoutType SingleWebPartAppPage | Out-Null
                    Set-PnPPage -Identity $DashPage -Publish | Out-Null
                    Write-Output "Pagina do Dashboard OK."
                }

                Write-Output "Desligando barra social do site..."
                Set-PnPSite -SocialBarOnSitePagesDisabled $true
                Write-Output "Barra social desligada."

                try { Disconnect-PnPOnline } catch {}
            } -ArgumentList $siteUrl, $clientId, $bibPage, $dashPage

            # Poll mantendo a UI viva
            while ($job.State -in 'Running','NotStarted') {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 200
            }

            $out = Receive-Job -Job $job -ErrorAction Continue 2>&1 | Out-String
            Remove-Job -Job $job -Force
            Write-Log $log $out

            if ($job.State -eq 'Completed') {
                Write-Log $log ''
                Write-Log $log '🎉 Tudo configurado! Clique em Avançar.' $colorSuccess
            } else {
                Write-Log $log "❌ Job terminou com estado: $($job.State)" $colorDanger
            }
        } catch {
            Write-Log $log "❌ Falha: $($_.Exception.Message)" $colorDanger
        }
    }.GetNewClosure())

    $btnBack.Visible = $true
    $btnNext.Text = 'Avançar >'
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

    function ChoiceXml { param([string]$Name, [string]$DisplayName, [string[]]$Choices)
        $esc = $Choices | ForEach-Object { [System.Security.SecurityElement]::Escape($_) }
        $cxml = ($esc | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ''
        return "<Field Type='Choice' Name='$Name' StaticName='$Name' DisplayName='$DisplayName' Format='Dropdown'><CHOICES>$cxml</CHOICES></Field>"
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

    $xAcao      = ChoiceXml 'acao' 'Ação' $ACOES
    $xPrompt    = "<Field Type='Note' Name='Prompt' StaticName='Prompt' DisplayName='Prompt' RichText='FALSE' NumLines='6' />"
    $xSegmento  = ChoiceXml 'Segmento' 'Segmento' $SEGMENTOS
    $xCategoria = ChoiceXml 'Categoria' 'Categoria' $CATEGORIAS
    $xFuncCom   = ChoiceXml 'Funcionacom' 'Funciona com' $FUNCIONA_COM
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

    Write-Log $Log ''
    Write-Log $Log ('[2/2] Lista de favoritos: ' + $global:state.FavoritesListTitle) $colorMuted
    EnsureList $global:state.FavoritesListTitle $Log
    EnsureField $global:state.FavoritesListTitle 'PromptID' $xPromptId $Log
    EnsureField $global:state.FavoritesListTitle 'acao' $xAcao $Log
    EnsureField $global:state.FavoritesListTitle 'Prompt' $xPrompt $Log
    EnsureField $global:state.FavoritesListTitle 'Segmento' $xSegmento $Log
    EnsureField $global:state.FavoritesListTitle 'Categoria' $xCategoria $Log
    EnsureField $global:state.FavoritesListTitle 'Funcionacom' $xFuncCom $Log
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
$form.Size = New-Object System.Drawing.Size(770, 620)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 70
$header.BackColor = $colorPrimary

$title = New-Object System.Windows.Forms.Label
$title.Text = '📚 Biblioteca de Prompts'
$title.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::White
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20, 12)
$header.Controls.Add($title)

$lblStep = New-Object System.Windows.Forms.Label
$lblStep.Text = 'Etapa 1 de 7'
$lblStep.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$lblStep.ForeColor = [System.Drawing.Color]::White
$lblStep.AutoSize = $true
$lblStep.Location = New-Object System.Drawing.Point(20, 42)
$header.Controls.Add($lblStep)

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
$footer.Height = 60
$footer.BackColor = $colorBg

$btnNext = New-Button -Text 'Avançar >' -X 620 -Y 15 -Width 120 -Primary $true
$btnBack = New-Button -Text '< Voltar' -X 490 -Y 15 -Width 120

$footer.Controls.Add($btnNext)
$footer.Controls.Add($btnBack)
$form.Controls.Add($footer)

# Content
$content = New-Object System.Windows.Forms.Panel
$content.Dock = 'Fill'
$content.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($content)

# --- ordem correta pra Dock ---
$form.Controls.SetChildIndex($footer, 0)
$form.Controls.SetChildIndex($content, 1)
$form.Controls.SetChildIndex($progressBar, 2)
$form.Controls.SetChildIndex($header, 3)

# Navegação
$btnNext.Add_Click({
    # Executa callback do passo atual (salvar campos digitados no state) antes de avançar.
    if ($global:onNext) {
        try { & $global:onNext } catch { }
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

