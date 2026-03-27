Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$AppTitle = "ZIP Multiplas Pasta - PRO v6.6.2"
$MailSubject = "ZIP Multiplas Pasta - PRO v6.6.2"
$StateFile = Join-Path -Path $env:APPDATA -ChildPath "VHugoDevIA\zip_multiplas_pastas_pro_v6_6_2_state.json"
$script:CurrentLogFile = $null
$script:ComputerSuffix = if ($env:COMPUTERNAME -and $env:COMPUTERNAME.Length -ge 4) { $env:COMPUTERNAME.Substring($env:COMPUTERNAME.Length - 4).ToUpper() } elseif ($env:COMPUTERNAME) { $env:COMPUTERNAME.ToUpper() } else { "PC" }

# Theme variables
$script:CurrentTheme = "Light"
$script:ThemeColors = @{
    Light = @{
        FormBackColor = [System.Drawing.Color]::White
        FormForeColor = [System.Drawing.Color]::Black
        PanelBackColor = [System.Drawing.SystemColors]::Control
        ButtonBackColor = [System.Drawing.SystemColors]::Control
        ButtonForeColor = [System.Drawing.Color]::Black
        TextBoxBackColor = [System.Drawing.Color]::White
        TextBoxForeColor = [System.Drawing.Color]::Black
        LabelForeColor = [System.Drawing.Color]::Black
        GroupBoxForeColor = [System.Drawing.Color]::Black
        ProgressBarBackColor = [System.Drawing.SystemColors]::Control
        ProgressBarForeColor = [System.Drawing.Color]::Green
        ListBoxBackColor = [System.Drawing.Color]::White
        ListBoxForeColor = [System.Drawing.Color]::Black
    }
    Dark = @{
        FormBackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        FormForeColor = [System.Drawing.Color]::White
        PanelBackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
        ButtonBackColor = [System.Drawing.Color]::FromArgb(63, 63, 70)
        ButtonForeColor = [System.Drawing.Color]::White
        TextBoxBackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        TextBoxForeColor = [System.Drawing.Color]::White
        LabelForeColor = [System.Drawing.Color]::White
        GroupBoxForeColor = [System.Drawing.Color]::White
        ProgressBarBackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
        ProgressBarForeColor = [System.Drawing.Color]::LimeGreen
        ListBoxBackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        ListBoxForeColor = [System.Drawing.Color]::White
    }
}

function ApplyTheme {
    param([string]$ThemeName)

    $colors = $script:ThemeColors[$ThemeName]
    $form.BackColor = $colors.FormBackColor
    $form.ForeColor = $colors.FormForeColor

    foreach ($control in $form.Controls) {
        ApplyThemeToControl -Control $control -Colors $colors
    }
}

function ApplyThemeToControl {
    param([System.Windows.Forms.Control]$Control, [hashtable]$Colors)

    switch ($Control.GetType().Name) {
        "Button" {
            $Control.BackColor = $Colors.ButtonBackColor
            $Control.ForeColor = $Colors.ButtonForeColor
            $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        }
        "TextBox" {
            $Control.BackColor = $Colors.TextBoxBackColor
            $Control.ForeColor = $Colors.TextBoxForeColor
        }
        "Label" {
            $Control.ForeColor = $Colors.LabelForeColor
        }
        "GroupBox" {
            $Control.ForeColor = $Colors.GroupBoxForeColor
            foreach ($subControl in $Control.Controls) {
                ApplyThemeToControl -Control $subControl -Colors $Colors
            }
        }
        "ListBox" {
            $Control.BackColor = $Colors.ListBoxBackColor
            $Control.ForeColor = $Colors.ListBoxForeColor
        }
        "ProgressBar" {
            $Control.BackColor = $Colors.ProgressBarBackColor
            $Control.ForeColor = $Colors.ProgressBarForeColor
        }
        "Panel" {
            $Control.BackColor = $Colors.PanelBackColor
            foreach ($subControl in $Control.Controls) {
                ApplyThemeToControl -Control $subControl -Colors $Colors
            }
        }
        "MenuStrip" {
            $Control.BackColor = $Colors.PanelBackColor
            $Control.ForeColor = $Colors.FormForeColor
        }
        default {
            if ($Control.HasChildren) {
                foreach ($subControl in $Control.Controls) {
                    ApplyThemeToControl -Control $subControl -Colors $Colors
                }
            }
        }
    }
}

if (-not (Test-Path -LiteralPath $SevenZip)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Nao encontrei o 7z.exe em:`n$SevenZip",
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

function Ensure-StateFolder {
    $dir = Split-Path -Path $StateFile -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Load-AppState {
    if (Test-Path -LiteralPath $StateFile) {
        try {
            return Get-Content -LiteralPath $StateFile -Raw | ConvertFrom-Json
        }
        catch {
            return $null
        }
    }
    return $null
}

function Save-AppState {
    param([hashtable]$State)

    try {
        Ensure-StateFolder
        $State | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $StateFile -Encoding UTF8
    }
    catch {
    }
}

function Get-UniqueArchivePath {
    param(
        [string]$Parent,
        [string]$BaseName,
        [string]$Extension
    )

    $dest = Join-Path -Path $Parent -ChildPath ($BaseName + "." + $Extension)
    if (-not (Test-Path -LiteralPath $dest)) {
        return $dest
    }

    $i = 2
    do {
        $dest = Join-Path -Path $Parent -ChildPath ("{0}_{1}.{2}" -f $BaseName, $i, $Extension)
        $i++
    } while (Test-Path -LiteralPath $dest)

    return $dest
}

function Remove-EmptyFolderIfPossible {
    param([string]$FolderPath)

    try {
        if (Test-Path -LiteralPath $FolderPath -PathType Container) {
            $items = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop
            if ($items.Count -eq 0) {
                Remove-Item -LiteralPath $FolderPath -Force -ErrorAction Stop
            }
        }
    }
    catch {
    }
}

function Remove-OriginalFolderAfterSuccess {
    param([string]$FolderPath)

    for ($i = 0; $i -lt 8; $i++) {
        try {
            if (Test-Path -LiteralPath $FolderPath) {
                Remove-Item -LiteralPath $FolderPath -Recurse -Force -ErrorAction Stop
            }
            return $true
        }
        catch {
            Start-Sleep -Milliseconds 250
        }
    }
    return $false
}

function Remove-CreatedArchives {
    param([System.Collections.Generic.List[string]]$ArchivePaths)

    foreach ($archive in $ArchivePaths) {
        for ($i = 0; $i -lt 8; $i++) {
            try {
                if (Test-Path -LiteralPath $archive) {
                    Remove-Item -LiteralPath $archive -Force -ErrorAction Stop
                }
                break
            }
            catch {
                Start-Sleep -Milliseconds 250
            }
        }
    }
}

function Add-FolderToList {
    param(
        [string]$Path,
        [System.Windows.Forms.ListBox]$ListBox,
        [System.Collections.Generic.List[string]]$Store
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return }

    $full = [System.IO.Path]::GetFullPath($Path)

    if (-not $Store.Contains($full)) {
        $Store.Add($full)
        [void]$ListBox.Items.Add($full)
    }
}

function Add-LogLine {
    param(
        [System.Windows.Forms.TextBox]$TextBox,
        [string]$Text
    )

    $TextBox.AppendText($Text + [Environment]::NewLine)
    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()

    if ($script:CurrentLogFile) {
        try {
            Add-Content -LiteralPath $script:CurrentLogFile -Value $Text -Encoding Unicode
        }
        catch {
        }
    }
}

function Start-NewProcessingLog {
    $script:CurrentLogFile = $null

    try {
        $documentsPath = [Environment]::GetFolderPath("MyDocuments")
        $logFolder = Join-Path -Path $documentsPath -ChildPath ".ZIP_Múltiplas_Pastas-PRO LOG"

        if (-not (Test-Path -LiteralPath $logFolder)) {
            New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
        }

        $logName = "ZIP_Multiplas_pastas_{0}.log" -f (Get-Date -Format "dd-MM-yyyy_HH-mm")
        $script:CurrentLogFile = Join-Path -Path $logFolder -ChildPath $logName

        $mailTo = "mailto:VHugoDevIA@gmail.com?subject={0}" -f [uri]::EscapeDataString($MailSubject)
        $headerLine = "".PadLeft(57, "=")

        @(
            "V.Hugo - Dev.IA"
            "`tVHugoDevIA@gmail.com"
            "`t$mailTo"
            ""
            "V.Hugo - Dev.IA"
            "`t$AppTitle"
            ""
            "$headerLine LOG DE PROCESSAMENTO $headerLine"
            ""
            "$AppTitle"
            ""
            "Criado em: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')"
            ""
            "----------------------------------------"
            "Log iniciado em: $script:CurrentLogFile"
            ""
        ) | Set-Content -LiteralPath $script:CurrentLogFile -Encoding Unicode
    }
    catch {
        $script:CurrentLogFile = $null
    }
}

function Build-SevenZipArguments {
    param(
        [string]$ArchivePath,
        [string]$FolderName,
        [string]$Format,
        [string]$Level,
        [decimal]$Threads,
        [bool]$DeleteAfter,
        [bool]$CompressShared,
        [bool]$CreateSfx,
        [string]$VolumeSize,
        [string]$Password,
        [string]$EncryptionMethod,
        [bool]$EncryptHeaders,
        [string]$Method,
        [string]$DictionarySize,
        [string]$WordSize,
        [string]$SolidBlock,
        [string]$ExtraParams
    )

    $args = New-Object System.Collections.Generic.List[string]

    $args.Add("a")
    $args.Add("-t$Format")
    $args.Add("`"$ArchivePath`"")
    $args.Add("`"$FolderName`"")
    $args.Add("-mx=$Level")
    $args.Add("-mmt=$([int]$Threads)")
    $args.Add("-bsp1")
    $args.Add("-bso1")
    $args.Add("-bse1")
    $args.Add("-y")

    if ($CompressShared) { $args.Add("-ssw") }
    if (-not [string]::IsNullOrWhiteSpace($VolumeSize)) { $args.Add("-v$($VolumeSize.Trim())") }
    if (($Format -eq "7z") -and $CreateSfx) { $args.Add("-sfx") }

    if (-not [string]::IsNullOrEmpty($Password)) {
        $args.Add("-p`"$Password`"")

        if ($Format -eq "zip") {
            switch ($EncryptionMethod) {
                "ZipCrypto" { $args.Add("-mem=ZipCrypto") }
                "AES128"    { $args.Add("-mem=AES128") }
                "AES192"    { $args.Add("-mem=AES192") }
                "AES256"    { $args.Add("-mem=AES256") }
            }
        }
        elseif ($Format -eq "7z") {
            if ($EncryptHeaders) { $args.Add("-mhe=on") }
        }
    }

    if ($Format -eq "zip") {
        if ($Method -and ($Method -ne "Automatico")) {
            $args.Add("-mm=$Method")
        }

        switch ($Method) {
            "Deflate" {
                if ($WordSize -and ($WordSize -ne "Automatico")) {
                    $args.Add("-mfb=$WordSize")
                }
            }
            "Deflate64" {
                if ($WordSize -and ($WordSize -ne "Automatico")) {
                    $args.Add("-mfb=$WordSize")
                }
            }
            "BZip2" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) {
                    $args.Add("-mfb=$WordSize")
                }
            }
            "LZMA" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) {
                    $args.Add("-mfb=$WordSize")
                }
            }
            "PPMd" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) {
                    $args.Add("-mfb=$WordSize")
                }
            }
        }
    }
    elseif ($Format -eq "7z") {
        if ($Method -and ($Method -ne "Automatico")) {
            $args.Add("-mm=$Method")
        }

        if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
            $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
            $args.Add("-md=$dict")
        }

        if ($WordSize -and ($WordSize -ne "Automatico")) {
            $args.Add("-mfb=$WordSize")
        }

        if ($SolidBlock -and ($SolidBlock -ne "Automatico")) {
            switch ($SolidBlock) {
                "Off" { $args.Add("-ms=off") }
                "On"  { $args.Add("-ms=on") }
                default { $args.Add("-ms=$SolidBlock") }
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExtraParams)) {
        foreach ($piece in ($ExtraParams.Trim() -split '\s+')) {
            if ($piece.Trim().Length -gt 0) {
                $args.Add($piece.Trim())
            }
        }
    }

    return ,$args
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = New-Object System.Drawing.Size(1100, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 950)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.AllowDrop = $true

# ToolTip
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500
$toolTip.ShowAlways = $true

# Menu Strip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "&Arquivo"
$menuStrip.Items.Add($fileMenu)

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "&Sair"
$exitItem.Add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($exitItem)

# View Menu
$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "&Visualizar"
$menuStrip.Items.Add($viewMenu)

$themeMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$themeMenu.Text = "&Tema"
$viewMenu.DropDownItems.Add($themeMenu)

$lightThemeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$lightThemeItem.Text = "&Claro"
$lightThemeItem.Add_Click({
    $script:CurrentTheme = "Light"
    ApplyTheme -ThemeName $script:CurrentTheme
})
$themeMenu.DropDownItems.Add($lightThemeItem)

$darkThemeItem = New-Object System.Windows.Forms.ToolStripMenuItem
$darkThemeItem.Text = "&Escuro"
$darkThemeItem.Add_Click({
    $script:CurrentTheme = "Dark"
    ApplyTheme -ThemeName $script:CurrentTheme
})
$themeMenu.DropDownItems.Add($darkThemeItem)

# Help Menu
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "&Ajuda"
$menuStrip.Items.Add($helpMenu)

$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutItem.Text = "&Sobre"
$aboutItem.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("$AppTitle`nCriado por V.Hugo - Dev.IA", "Sobre")
})
$helpMenu.DropDownItems.Add($aboutItem)

$linkEmail = New-Object System.Windows.Forms.LinkLabel
$linkEmail.Text = "VHugoDevIA@gmail.com"
$linkEmail.Location = New-Object System.Drawing.Point(450, 35)
$linkEmail.Size = New-Object System.Drawing.Size(220, 20)
$linkEmail.TextAlign = "MiddleCenter"
$form.Controls.Add($linkEmail)

$lblBrand = New-Object System.Windows.Forms.Label
$lblBrand.Text = "V.Hugo-AppDev.+IA"
$lblBrand.Location = New-Object System.Drawing.Point(840, 35)
$lblBrand.Size = New-Object System.Drawing.Size(220, 20)
$lblBrand.TextAlign = "MiddleRight"
$form.Controls.Add($lblBrand)

# Logo/Icon area
$lblLogo = New-Object System.Windows.Forms.Label
$lblLogo.Text = "📦 ZIP PRO"
$lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblLogo.Location = New-Object System.Drawing.Point(20, 30)
$lblLogo.Size = New-Object System.Drawing.Size(200, 30)
$lblLogo.TextAlign = "MiddleLeft"
$form.Controls.Add($lblLogo)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = "Arraste várias pastas do Explorador para a lista abaixo, ou use 'Adicionar pasta'."
$lblTop.Location = New-Object System.Drawing.Point(20, 60)
$lblTop.Size = New-Object System.Drawing.Size(1060, 20)
$form.Controls.Add($lblTop)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Adicionar pasta"
$btnAdd.Location = New-Object System.Drawing.Point(20, 90)
$btnAdd.Size = New-Object System.Drawing.Size(140, 35)
$form.Controls.Add($btnAdd)
$toolTip.SetToolTip($btnAdd, "Adicionar uma pasta para compressão usando o explorador de arquivos")

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remover selecionada"
$btnRemove.Location = New-Object System.Drawing.Point(170, 90)
$btnRemove.Size = New-Object System.Drawing.Size(160, 35)
$form.Controls.Add($btnRemove)
$toolTip.SetToolTip($btnRemove, "Remover a pasta selecionada da lista")

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Limpar lista"
$btnClear.Location = New-Object System.Drawing.Point(340, 90)
$btnClear.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($btnClear)
$toolTip.SetToolTip($btnClear, "Remover todas as pastas da lista")

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text = "Limpar log"
$btnClearLog.Location = New-Object System.Drawing.Point(470, 90)
$btnClearLog.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($btnClearLog)
$toolTip.SetToolTip($btnClearLog, "Limpar o conteúdo da área de log")

$btnOptions = New-Object System.Windows.Forms.Button
$btnOptions.Text = "Opções..."
$btnOptions.Location = New-Object System.Drawing.Point(600, 90)
$btnOptions.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($btnOptions)
$toolTip.SetToolTip($btnOptions, "Configurar opções adicionais do aplicativo")

$btnCancelMain = New-Object System.Windows.Forms.Button
$btnCancelMain.Text = "Cancelar"
$btnCancelMain.Location = New-Object System.Drawing.Point(820, 90)
$btnCancelMain.Size = New-Object System.Drawing.Size(120, 35)
$btnCancelMain.Enabled = $false
$form.Controls.Add($btnCancelMain)
$toolTip.SetToolTip($btnCancelMain, "Cancelar a operação de compressão em andamento")

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Iniciar"
$btnRun.Location = New-Object System.Drawing.Point(950, 90)
$btnRun.Size = New-Object System.Drawing.Size(130, 35)
$btnRun.Enabled = $false
$form.Controls.Add($btnRun)
$toolTip.SetToolTip($btnRun, "Iniciar a compressão das pastas selecionadas")

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20, 140)
$listBox.Size = New-Object System.Drawing.Size(1060, 130)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$form.Controls.Add($listBox)
$toolTip.SetToolTip($listBox, "Lista de pastas a comprimir. Arraste pastas aqui ou use 'Adicionar pasta'")

$grpMain = New-Object System.Windows.Forms.GroupBox
$grpMain.Text = "Compressão"
$grpMain.Location = New-Object System.Drawing.Point(20, 280)
$grpMain.Size = New-Object System.Drawing.Size(520, 270)
$form.Controls.Add($grpMain)

$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = "Formato do arquivo:"
$lblFormat.Location = New-Object System.Drawing.Point(15, 30)
$lblFormat.Size = New-Object System.Drawing.Size(130, 20)
$grpMain.Controls.Add($lblFormat)

$cmbFormat = New-Object System.Windows.Forms.ComboBox
$cmbFormat.Location = New-Object System.Drawing.Point(160, 27)
$cmbFormat.Size = New-Object System.Drawing.Size(120, 24)
$cmbFormat.DropDownStyle = "DropDownList"
[void]$cmbFormat.Items.AddRange(@("zip", "7z"))
$cmbFormat.SelectedItem = "zip"
$grpMain.Controls.Add($cmbFormat)

$lblLevel = New-Object System.Windows.Forms.Label
$lblLevel.Text = "Nível de compressão:"
$lblLevel.Location = New-Object System.Drawing.Point(15, 60)
$lblLevel.Size = New-Object System.Drawing.Size(130, 20)
$grpMain.Controls.Add($lblLevel)

$cmbLevel = New-Object System.Windows.Forms.ComboBox
$cmbLevel.Location = New-Object System.Drawing.Point(160, 57)
$cmbLevel.Size = New-Object System.Drawing.Size(120, 24)
$cmbLevel.DropDownStyle = "DropDownList"
[void]$cmbLevel.Items.AddRange(@("0","1","3","5","7","9"))
$cmbLevel.SelectedItem = "9"
$grpMain.Controls.Add($cmbLevel)

$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Método de compressão:"
$lblMethod.Location = New-Object System.Drawing.Point(15, 90)
$lblMethod.Size = New-Object System.Drawing.Size(140, 20)
$grpMain.Controls.Add($lblMethod)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Location = New-Object System.Drawing.Point(160, 87)
$cmbMethod.Size = New-Object System.Drawing.Size(120, 24)
$cmbMethod.DropDownStyle = "DropDownList"
$grpMain.Controls.Add($cmbMethod)

$lblDict = New-Object System.Windows.Forms.Label
$lblDict.Text = "Tamanho do dicionário:"
$lblDict.Location = New-Object System.Drawing.Point(15, 120)
$lblDict.Size = New-Object System.Drawing.Size(140, 20)
$grpMain.Controls.Add($lblDict)

$cmbDict = New-Object System.Windows.Forms.ComboBox
$cmbDict.Location = New-Object System.Drawing.Point(160, 117)
$cmbDict.Size = New-Object System.Drawing.Size(120, 24)
$cmbDict.DropDownStyle = "DropDownList"
$grpMain.Controls.Add($cmbDict)

$lblWord = New-Object System.Windows.Forms.Label
$lblWord.Text = "Tamanho da palavra:"
$lblWord.Location = New-Object System.Drawing.Point(15, 150)
$lblWord.Size = New-Object System.Drawing.Size(140, 20)
$grpMain.Controls.Add($lblWord)

$cmbWord = New-Object System.Windows.Forms.ComboBox
$cmbWord.Location = New-Object System.Drawing.Point(160, 147)
$cmbWord.Size = New-Object System.Drawing.Size(120, 24)
$cmbWord.DropDownStyle = "DropDownList"
$grpMain.Controls.Add($cmbWord)

$lblSolid = New-Object System.Windows.Forms.Label
$lblSolid.Text = "Blocos sólidos:"
$lblSolid.Location = New-Object System.Drawing.Point(15, 180)
$lblSolid.Size = New-Object System.Drawing.Size(140, 20)
$grpMain.Controls.Add($lblSolid)

$cmbSolid = New-Object System.Windows.Forms.ComboBox
$cmbSolid.Location = New-Object System.Drawing.Point(160, 177)
$cmbSolid.Size = New-Object System.Drawing.Size(120, 24)
$cmbSolid.DropDownStyle = "DropDownList"
$grpMain.Controls.Add($cmbSolid)

$lblThreads = New-Object System.Windows.Forms.Label
$lblThreads.Text = "Nº processos CPU:"
$lblThreads.Location = New-Object System.Drawing.Point(320, 30)
$lblThreads.Size = New-Object System.Drawing.Size(120, 20)
$grpMain.Controls.Add($lblThreads)

$numThreads = New-Object System.Windows.Forms.NumericUpDown
$numThreads.Location = New-Object System.Drawing.Point(445, 27)
$numThreads.Size = New-Object System.Drawing.Size(55, 24)
$numThreads.Minimum = 1
$numThreads.Maximum = [Math]::Max(1, [Environment]::ProcessorCount)
$numThreads.Value = [Math]::Max(1, [Environment]::ProcessorCount)
$grpMain.Controls.Add($numThreads)

$lblVolume = New-Object System.Windows.Forms.Label
$lblVolume.Text = "Dividir por volumes:"
$lblVolume.Location = New-Object System.Drawing.Point(320, 60)
$lblVolume.Size = New-Object System.Drawing.Size(120, 20)
$grpMain.Controls.Add($lblVolume)

$txtVolume = New-Object System.Windows.Forms.TextBox
$txtVolume.Location = New-Object System.Drawing.Point(445, 57)
$txtVolume.Size = New-Object System.Drawing.Size(55, 24)
$grpMain.Controls.Add($txtVolume)

$lblUpdateMode = New-Object System.Windows.Forms.Label
$lblUpdateMode.Text = "Se o ZIP já existir:"
$lblUpdateMode.Location = New-Object System.Drawing.Point(320, 90)
$lblUpdateMode.Size = New-Object System.Drawing.Size(120, 20)
$grpMain.Controls.Add($lblUpdateMode)

$cmbUpdateMode = New-Object System.Windows.Forms.ComboBox
$cmbUpdateMode.Location = New-Object System.Drawing.Point(445, 87)
$cmbUpdateMode.Size = New-Object System.Drawing.Size(70, 24)
$cmbUpdateMode.DropDownStyle = "DropDownList"
[void]$cmbUpdateMode.Items.AddRange(@("Novo","Ignorar","Nome unico"))
$cmbUpdateMode.SelectedItem = "Nome unico"
$grpMain.Controls.Add($cmbUpdateMode)

$lblPathMode = New-Object System.Windows.Forms.Label
$lblPathMode.Text = "Passar pasta ao 7-Zip:"
$lblPathMode.Location = New-Object System.Drawing.Point(320, 120)
$lblPathMode.Size = New-Object System.Drawing.Size(120, 30)
$grpMain.Controls.Add($lblPathMode)

$cmbPathMode = New-Object System.Windows.Forms.ComboBox
$cmbPathMode.Location = New-Object System.Drawing.Point(445, 122)
$cmbPathMode.Size = New-Object System.Drawing.Size(70, 24)
$cmbPathMode.DropDownStyle = "DropDownList"
[void]$cmbPathMode.Items.AddRange(@("Relativo","Absoluto"))
$cmbPathMode.SelectedItem = "Relativo"
$grpMain.Controls.Add($cmbPathMode)

$chkSfx = New-Object System.Windows.Forms.CheckBox
$chkSfx.Text = "Criar arquivo SFX"
$chkSfx.Location = New-Object System.Drawing.Point(320, 155)
$chkSfx.Size = New-Object System.Drawing.Size(160, 22)
$grpMain.Controls.Add($chkSfx)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = "Comprimir ficheiros partilhados"
$chkShared.Location = New-Object System.Drawing.Point(320, 180)
$chkShared.Size = New-Object System.Drawing.Size(190, 22)
$grpMain.Controls.Add($chkShared)

$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Eliminar ficheiros após compressão"
$chkDelete.Location = New-Object System.Drawing.Point(320, 205)
$chkDelete.Size = New-Object System.Drawing.Size(190, 22)
$chkDelete.Checked = $true
$grpMain.Controls.Add($chkDelete)

$grpEnc = New-Object System.Windows.Forms.GroupBox
$grpEnc.Text = "Encriptação"
$grpEnc.Location = New-Object System.Drawing.Point(560, 280)
$grpEnc.Size = New-Object System.Drawing.Size(520, 270)
$form.Controls.Add($grpEnc)

$lblPass1 = New-Object System.Windows.Forms.Label
$lblPass1.Text = "Introduza a palavra-passe:"
$lblPass1.Location = New-Object System.Drawing.Point(15, 30)
$lblPass1.Size = New-Object System.Drawing.Size(180, 20)
$grpEnc.Controls.Add($lblPass1)

$txtPass1 = New-Object System.Windows.Forms.TextBox
$txtPass1.Location = New-Object System.Drawing.Point(200, 27)
$txtPass1.Size = New-Object System.Drawing.Size(300, 24)
$txtPass1.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass1)

$lblPass2 = New-Object System.Windows.Forms.Label
$lblPass2.Text = "Reintroduza a palavra-passe:"
$lblPass2.Location = New-Object System.Drawing.Point(15, 60)
$lblPass2.Size = New-Object System.Drawing.Size(180, 20)
$grpEnc.Controls.Add($lblPass2)

$txtPass2 = New-Object System.Windows.Forms.TextBox
$txtPass2.Location = New-Object System.Drawing.Point(200, 57)
$txtPass2.Size = New-Object System.Drawing.Size(300, 24)
$txtPass2.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass2)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Mostrar palavra-passe"
$chkShowPass.Location = New-Object System.Drawing.Point(15, 90)
$chkShowPass.Size = New-Object System.Drawing.Size(180, 22)
$grpEnc.Controls.Add($chkShowPass)

$lblEncMethod = New-Object System.Windows.Forms.Label
$lblEncMethod.Text = "Método de encriptação:"
$lblEncMethod.Location = New-Object System.Drawing.Point(15, 125)
$lblEncMethod.Size = New-Object System.Drawing.Size(150, 20)
$grpEnc.Controls.Add($lblEncMethod)

$cmbEncMethod = New-Object System.Windows.Forms.ComboBox
$cmbEncMethod.Location = New-Object System.Drawing.Point(200, 122)
$cmbEncMethod.Size = New-Object System.Drawing.Size(150, 24)
$cmbEncMethod.DropDownStyle = "DropDownList"
$grpEnc.Controls.Add($cmbEncMethod)

$chkEncryptHeaders = New-Object System.Windows.Forms.CheckBox
$chkEncryptHeaders.Text = "Encriptar nomes dos ficheiros (7z)"
$chkEncryptHeaders.Location = New-Object System.Drawing.Point(15, 160)
$chkEncryptHeaders.Size = New-Object System.Drawing.Size(250, 22)
$grpEnc.Controls.Add($chkEncryptHeaders)

$grpExtra = New-Object System.Windows.Forms.GroupBox
$grpExtra.Text = "Opções extra"
$grpExtra.Location = New-Object System.Drawing.Point(20, 570)
$grpExtra.Size = New-Object System.Drawing.Size(1060, 80)
$form.Controls.Add($grpExtra)

$lblExtraParams = New-Object System.Windows.Forms.Label
$lblExtraParams.Text = "Parametros:"
$lblExtraParams.Location = New-Object System.Drawing.Point(15, 30)
$lblExtraParams.Size = New-Object System.Drawing.Size(100, 20)
$grpExtra.Controls.Add($lblExtraParams)

$txtExtraParams = New-Object System.Windows.Forms.TextBox
$txtExtraParams.Location = New-Object System.Drawing.Point(120, 27)
$txtExtraParams.Size = New-Object System.Drawing.Size(920, 24)
$grpExtra.Controls.Add($txtExtraParams)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Estado: parado"
$lblStatus.Location = New-Object System.Drawing.Point(20, 665)
$lblStatus.Size = New-Object System.Drawing.Size(1060, 20)
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20, 690)
$progressFolder.Size = New-Object System.Drawing.Size(1060, 24)
$progressFolder.Minimum = 0
$progressFolder.Maximum = 100
$form.Controls.Add($progressFolder)

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.Text = "Progresso total:"
$lblTotal.Location = New-Object System.Drawing.Point(20, 720)
$lblTotal.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblTotal)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20, 745)
$progressTotal.Size = New-Object System.Drawing.Size(1060, 24)
$progressTotal.Minimum = 0
$progressTotal.Maximum = 100
$form.Controls.Add($progressTotal)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 780)
$txtLog.Size = New-Object System.Drawing.Size(1060, 160)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)
$toolTip.SetToolTip($txtLog, "Área de log mostrando o progresso e mensagens da compressão")

$selectedFolders = New-Object 'System.Collections.Generic.List[string]'

$optionFlags = @{
    openLogFolderAfterFinish = $true
    confirmBeforeClearList   = $false
}

function Refresh-UiState {
    $btnRun.Enabled = ($selectedFolders.Count -gt 0)
}

function Update-DynamicOptions {
    $format = [string]$cmbFormat.SelectedItem
    $method = [string]$cmbMethod.SelectedItem

    $cmbSolid.Enabled = ($format -eq "7z")
    $chkEncryptHeaders.Enabled = ($format -eq "7z")
    $chkSfx.Enabled = ($format -eq "7z")

    if ($format -eq "zip") {
        switch ($method) {
            "Deflate" {
                $cmbDict.SelectedItem = "32KB"
                $cmbDict.Enabled = $false
                $cmbWord.Enabled = $true
            }
            "Deflate64" {
                $cmbDict.Enabled = $false
                $cmbWord.Enabled = $true
            }
            default {
                $cmbDict.Enabled = $true
                $cmbWord.Enabled = $true
            }
        }
    }
    else {
        $cmbDict.Enabled = $true
        $cmbWord.Enabled = $true
    }
}

function Update-FormatOptions {
    $format = [string]$cmbFormat.SelectedItem
    $prevMethod = [string]$cmbMethod.SelectedItem
    $prevDict = [string]$cmbDict.SelectedItem
    $prevWord = [string]$cmbWord.SelectedItem
    $prevSolid = [string]$cmbSolid.SelectedItem
    $prevEnc = [string]$cmbEncMethod.SelectedItem

    $cmbMethod.Items.Clear()
    $cmbDict.Items.Clear()
    $cmbWord.Items.Clear()
    $cmbSolid.Items.Clear()
    $cmbEncMethod.Items.Clear()

    if ($format -eq "zip") {
        [void]$cmbMethod.Items.AddRange(@("Automatico","Deflate","Deflate64","BZip2","LZMA","PPMd"))
        [void]$cmbDict.Items.AddRange(@("Automatico","32KB","64KB","128KB","256KB","512KB","1MB","2MB","4MB","8MB","16MB","32MB"))
        [void]$cmbWord.Items.AddRange(@("Automatico","32","64","128","192","256","273"))
        [void]$cmbSolid.Items.AddRange(@("Automatico"))
        [void]$cmbEncMethod.Items.AddRange(@("ZipCrypto","AES128","AES192","AES256"))

        if ($prevMethod -and $cmbMethod.Items.Contains($prevMethod)) { $cmbMethod.SelectedItem = $prevMethod } else { $cmbMethod.SelectedItem = "Automatico" }
        if ($prevDict -and $cmbDict.Items.Contains($prevDict)) { $cmbDict.SelectedItem = $prevDict } else { $cmbDict.SelectedItem = "Automatico" }
        if ($prevWord -and $cmbWord.Items.Contains($prevWord)) { $cmbWord.SelectedItem = $prevWord } else { $cmbWord.SelectedItem = "Automatico" }
        $cmbSolid.SelectedItem = "Automatico"
        if ($prevEnc -and $cmbEncMethod.Items.Contains($prevEnc)) { $cmbEncMethod.SelectedItem = $prevEnc } else { $cmbEncMethod.SelectedItem = "ZipCrypto" }
    }
    else {
        [void]$cmbMethod.Items.AddRange(@("Automatico","LZMA2","LZMA","PPMd","BZip2"))
        [void]$cmbDict.Items.AddRange(@("Automatico","64KB","1MB","2MB","4MB","8MB","16MB","32MB","64MB","128MB","256MB"))
        [void]$cmbWord.Items.AddRange(@("Automatico","8","12","16","24","32","48","64","96","128","192","273"))
        [void]$cmbSolid.Items.AddRange(@("Automatico","Off","On","1m","2m","4m","8m","16m","32m","64m","128m","256m","512m"))
        [void]$cmbEncMethod.Items.AddRange(@("AES256"))

        if ($prevMethod -and $cmbMethod.Items.Contains($prevMethod)) { $cmbMethod.SelectedItem = $prevMethod } else { $cmbMethod.SelectedItem = "Automatico" }
        if ($prevDict -and $cmbDict.Items.Contains($prevDict)) { $cmbDict.SelectedItem = $prevDict } else { $cmbDict.SelectedItem = "Automatico" }
        if ($prevWord -and $cmbWord.Items.Contains($prevWord)) { $cmbWord.SelectedItem = $prevWord } else { $cmbWord.SelectedItem = "Automatico" }
        if ($prevSolid -and $cmbSolid.Items.Contains($prevSolid)) { $cmbSolid.SelectedItem = $prevSolid } else { $cmbSolid.SelectedItem = "Automatico" }
        $cmbEncMethod.SelectedItem = "AES256"
    }

    Update-DynamicOptions
}

$cmbFormat.Add_SelectedIndexChanged({
    Update-FormatOptions
})

$cmbMethod.Add_SelectedIndexChanged({
    Update-DynamicOptions
})

$chkShowPass.Add_CheckedChanged({
    $visible = $chkShowPass.Checked
    $txtPass1.UseSystemPasswordChar = -not $visible
    $txtPass2.UseSystemPasswordChar = -not $visible
})

$linkEmail.Add_LinkClicked({
    $mailTo = "mailto:VHugoDevIA@gmail.com?subject={0}" -f [uri]::EscapeDataString($MailSubject)
    Start-Process $mailTo
})

$btnOptions.Add_Click({
    $optForm = New-Object System.Windows.Forms.Form
    $optForm.Text = "Opções"
    $optForm.Size = New-Object System.Drawing.Size(450, 270)
    $optForm.StartPosition = "CenterParent"
    $optForm.FormBorderStyle = "FixedDialog"
    $optForm.MaximizeBox = $false
    $optForm.MinimizeBox = $false

    $chkOpenLog = New-Object System.Windows.Forms.CheckBox
    $chkOpenLog.Text = "Abrir pasta do log no final das operações"
    $chkOpenLog.Location = New-Object System.Drawing.Point(20, 20)
    $chkOpenLog.Size = New-Object System.Drawing.Size(350, 24)
    $chkOpenLog.Checked = [bool]$optionFlags.openLogFolderAfterFinish
    $optForm.Controls.Add($chkOpenLog)

    $chkConfirmClear = New-Object System.Windows.Forms.CheckBox
    $chkConfirmClear.Text = "Pedir confirmação ao limpar lista"
    $chkConfirmClear.Location = New-Object System.Drawing.Point(20, 55)
    $chkConfirmClear.Size = New-Object System.Drawing.Size(320, 24)
    $chkConfirmClear.Checked = [bool]$optionFlags.confirmBeforeClearList
    $optForm.Controls.Add($chkConfirmClear)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Estas opções também ficam guardadas na última sessão."
    $lblInfo.Location = New-Object System.Drawing.Point(20, 95)
    $lblInfo.Size = New-Object System.Drawing.Size(380, 40)
    $optForm.Controls.Add($lblInfo)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(230, 165)
    $btnOk.Size = New-Object System.Drawing.Size(75, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $optForm.Controls.Add($btnOk)

    $btnOptCancel = New-Object System.Windows.Forms.Button
    $btnOptCancel.Text = "Cancelar"
    $btnOptCancel.Location = New-Object System.Drawing.Point(315, 165)
    $btnOptCancel.Size = New-Object System.Drawing.Size(75, 30)
    $btnOptCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $optForm.Controls.Add($btnOptCancel)

    $optForm.AcceptButton = $btnOk
    $optForm.CancelButton = $btnOptCancel

    if ($optForm.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $optionFlags.openLogFolderAfterFinish = $chkOpenLog.Checked
        $optionFlags.confirmBeforeClearList = $chkConfirmClear.Checked
    }
})

$btnCancelMain.Add_Click({
    if (-not $script:CurrentProcess -and -not $btnCancelMain.Enabled) { return }

    $r = [System.Windows.Forms.MessageBox]::Show(
        "Queres cancelar a compactação em curso?",
        "Cancelar",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:CancelRequested = $true

    try {
        if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
            $script:CurrentProcess.Kill()
            $script:CurrentProcess.WaitForExit()
        }
    }
    catch {
    }

    Add-LogLine -TextBox $txtLog -Text "[CANCELADO] Pedido de cancelamento enviado pelo utilizador."
    $lblStatus.Text = "Estado: a cancelar..."
})

$btnAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecione uma pasta para adicionar"
    $dlg.ShowNewFolderButton = $false

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FolderToList -Path $dlg.SelectedPath -ListBox $listBox -Store $selectedFolders
        Add-LogLine -TextBox $txtLog -Text "Pasta adicionada: $($dlg.SelectedPath)"
        Refresh-UiState
    }
})

$btnRemove.Add_Click({
    if ($listBox.SelectedItems.Count -eq 0) { return }

    $toRemove = @()
    foreach ($item in $listBox.SelectedItems) {
        $toRemove += [string]$item
    }

    foreach ($item in $toRemove) {
        [void]$selectedFolders.Remove($item)
        [void]$listBox.Items.Remove($item)
        Add-LogLine -TextBox $txtLog -Text "Removida: $item"
    }

    Refresh-UiState
})

$btnClear.Add_Click({
    if ($optionFlags.confirmBeforeClearList) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Queres mesmo limpar a lista?",
            "Confirmar",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $selectedFolders.Clear()
    $listBox.Items.Clear()
    Add-LogLine -TextBox $txtLog -Text "Lista limpa."
    Refresh-UiState
})

$btnClearLog.Add_Click({
    $txtLog.Clear()
    if ($script:CurrentLogFile) {
        try {
            "=== LOG CONTINUADO ===" | Add-Content -LiteralPath $script:CurrentLogFile -Encoding Unicode
        }
        catch {
        }
    }
})

$listBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$listBox.Add_DragDrop({
    $items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)

    foreach ($item in $items) {
        if (Test-Path -LiteralPath $item -PathType Container) {
            Add-FolderToList -Path $item -ListBox $listBox -Store $selectedFolders
            Add-LogLine -TextBox $txtLog -Text "Pasta adicionada por arrastar: $item"
        }
    }

    Refresh-UiState
})

$btnRun.Add_Click({
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Adiciona pelo menos uma pasta.")
        return
    }

    $script:CancelRequested = $false
    $script:CurrentProcess = $null
    $createdArchivesThisRun = New-Object 'System.Collections.Generic.List[string]'
    $foldersToDeleteAfterSuccess = New-Object 'System.Collections.Generic.List[string]'

    Start-NewProcessingLog

    if ($script:CurrentLogFile) {
        Add-LogLine -TextBox $txtLog -Text "Log iniciado em: $script:CurrentLogFile"
    }

    $pass1 = $txtPass1.Text
    $pass2 = $txtPass2.Text

    if (($pass1.Length -gt 0) -or ($pass2.Length -gt 0)) {
        if ($pass1 -ne $pass2) {
            [System.Windows.Forms.MessageBox]::Show("As palavras-passe não coincidem.", "Erro")
            $btnRun.Enabled = $true
            $btnAdd.Enabled = $true
            $btnRemove.Enabled = $true
            $btnClear.Enabled = $true
            $btnClearLog.Enabled = $true
            $btnOptions.Enabled = $true
            $btnCancelMain.Enabled = $false
            Refresh-UiState
            return
        }
    }

    $btnRun.Enabled = $false
    $btnAdd.Enabled = $false
    $btnRemove.Enabled = $false
    $btnClear.Enabled = $false
    $btnClearLog.Enabled = $false
    $btnOptions.Enabled = $false
    $btnCancelMain.Enabled = $true

    $progressFolder.Value = 0
    $progressTotal.Value = 0

    $totalCount = $selectedFolders.Count
    $currentIndex = 0

    foreach ($folder in @($selectedFolders)) {
        if ($script:CancelRequested) { break }

        $currentIndex++

        try {
            $folderPath = [string]$folder
            $folderObj = Get-Item -LiteralPath $folderPath -ErrorAction Stop
            $folderName = $folderObj.Name
            $parent = $folderObj.Parent.FullName
            $format = [string]$cmbFormat.SelectedItem
            $archiveBaseName = "{0}_{1}" -f $folderName, $script:ComputerSuffix

            $existingDefault = Join-Path -Path $parent -ChildPath ($archiveBaseName + "." + $format)
            $destExistedBefore = $false

            switch ([string]$cmbUpdateMode.SelectedItem) {
                "Ignorar" {
                    if (Test-Path -LiteralPath $existingDefault) {
                        Add-LogLine -TextBox $txtLog -Text ""
                        Add-LogLine -TextBox $txtLog -Text "[IGNORADO] Já existe: $existingDefault"
                        $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
                        continue
                    }
                    $dest = $existingDefault
                }
                "Novo" {
                    $dest = $existingDefault
                }
                default {
                    $dest = Get-UniqueArchivePath -Parent $parent -BaseName $archiveBaseName -Extension $format
                }
            }

            $destExistedBefore = Test-Path -LiteralPath $dest

            $lblStatus.Text = "Estado: a comprimir $folderName ($currentIndex de $totalCount)"
            $progressFolder.Value = 0
            Add-LogLine -TextBox $txtLog -Text ""
            Add-LogLine -TextBox $txtLog -Text "A processar: $folderPath"
            Add-LogLine -TextBox $txtLog -Text "Destino: $dest"

            if (-not $destExistedBefore -and -not $createdArchivesThisRun.Contains($dest)) {
                $createdArchivesThisRun.Add($dest)
            }

            $itemToArchive = if ($cmbPathMode.SelectedItem -eq "Absoluto") { $folderPath } else { $folderName }

            $argList = Build-SevenZipArguments `
                -ArchivePath $dest `
                -FolderName $itemToArchive `
                -Format ([string]$cmbFormat.SelectedItem) `
                -Level ([string]$cmbLevel.SelectedItem) `
                -Threads $numThreads.Value `
                -DeleteAfter $chkDelete.Checked `
                -CompressShared $chkShared.Checked `
                -CreateSfx $chkSfx.Checked `
                -VolumeSize $txtVolume.Text `
                -Password $pass1 `
                -EncryptionMethod ([string]$cmbEncMethod.SelectedItem) `
                -EncryptHeaders $chkEncryptHeaders.Checked `
                -Method ([string]$cmbMethod.SelectedItem) `
                -DictionarySize ([string]$cmbDict.SelectedItem) `
                -WordSize ([string]$cmbWord.SelectedItem) `
                -SolidBlock ([string]$cmbSolid.SelectedItem) `
                -ExtraParams $txtExtraParams.Text

            $argString = ($argList -join " ")
            Add-LogLine -TextBox $txtLog -Text "Comando: $argString"

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $SevenZip
            $psi.WorkingDirectory = $parent
            $psi.Arguments = $argString
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $true

            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $psi
            $null = $p.Start()
            $script:CurrentProcess = $p

            while (-not $p.HasExited) {
                while (($line = $p.StandardOutput.ReadLine()) -ne $null) {
                    if ($line -match '(\d+)%') {
                        $pct = [int]$matches[1]
                        if ($pct -ge 0 -and $pct -le 100) {
                            $progressFolder.Value = $pct
                        }
                    }

                    if (-not [string]::IsNullOrWhiteSpace($line)) {
                        Add-LogLine -TextBox $txtLog -Text $line
                    }

                    [System.Windows.Forms.Application]::DoEvents()
                }

                while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                    if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                        Add-LogLine -TextBox $txtLog -Text "[ERRO] $errLine"
                    }

                    [System.Windows.Forms.Application]::DoEvents()
                }

                Start-Sleep -Milliseconds 120
                [System.Windows.Forms.Application]::DoEvents()

                if ($script:CancelRequested) {
                    break
                }
            }

            while (($line = $p.StandardOutput.ReadLine()) -ne $null) {
                if ($line -match '(\d+)%') {
                    $pct = [int]$matches[1]
                    if ($pct -ge 0 -and $pct -le 100) {
                        $progressFolder.Value = $pct
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Add-LogLine -TextBox $txtLog -Text $line
                }
            }

            while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                    Add-LogLine -TextBox $txtLog -Text "[ERRO] $errLine"
                }
            }

            $p.WaitForExit()
            $exitCode = $p.ExitCode
            $script:CurrentProcess = $null

            if ($script:CancelRequested) {
                Add-LogLine -TextBox $txtLog -Text "[CANCELADO] Processo interrompido: $folderName"
                break
            }
            elseif ($exitCode -eq 0) {
                $progressFolder.Value = 100
                Add-LogLine -TextBox $txtLog -Text "[OK] Concluído: $folderName"

                if ($chkDelete.Checked -and -not $foldersToDeleteAfterSuccess.Contains($folderPath)) {
                    $foldersToDeleteAfterSuccess.Add($folderPath)
                    Add-LogLine -TextBox $txtLog -Text "[PENDENTE] Originais serão eliminados apenas no fim, se toda a operação terminar com sucesso: $folderName"
                }
            }
            else {
                Add-LogLine -TextBox $txtLog -Text "[ERRO] Falhou: $folderName (código $exitCode)"
            }
        }
        catch {
            $script:CurrentProcess = $null
            Add-LogLine -TextBox $txtLog -Text "[ERRO] Exceção ao processar '$folder': $($_.Exception.Message)"
        }

        $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($script:CancelRequested) {
        Remove-CreatedArchives -ArchivePaths $createdArchivesThisRun
        $progressFolder.Value = 0
        $lblStatus.Text = "Estado: cancelado"
        Add-LogLine -TextBox $txtLog -Text ""
        Add-LogLine -TextBox $txtLog -Text "Processamento cancelado pelo utilizador."
        Add-LogLine -TextBox $txtLog -Text "Arquivos criados nesta operação foram removidos."
    }
    else {
        if ($chkDelete.Checked) {
            foreach ($folderToDelete in @($foldersToDeleteAfterSuccess)) {
                $removedOk = Remove-OriginalFolderAfterSuccess -FolderPath $folderToDelete
                if ($removedOk) {
                    Add-LogLine -TextBox $txtLog -Text "[OK] Originais eliminados após sucesso global: $folderToDelete"
                }
                else {
                    Add-LogLine -TextBox $txtLog -Text "[AVISO] Não foi possível eliminar a pasta original: $folderToDelete"
                }
            }
        }

        $lblStatus.Text = "Estado: terminado"
        Add-LogLine -TextBox $txtLog -Text ""
        Add-LogLine -TextBox $txtLog -Text "Processamento concluído."
    }

    if ($script:CurrentLogFile) {
        Add-LogLine -TextBox $txtLog -Text "Log gravado em: $script:CurrentLogFile"
    }

    $script:CurrentProcess = $null
    $script:CancelRequested = $false
    $btnRun.Enabled = $true
    $btnAdd.Enabled = $true
    $btnRemove.Enabled = $true
    $btnClear.Enabled = $true
    $btnClearLog.Enabled = $true
    $btnOptions.Enabled = $true
    $btnCancelMain.Enabled = $false
    Refresh-UiState

    if ($script:CancelRequested) {
        [System.Windows.Forms.MessageBox]::Show("Processamento cancelado.", $AppTitle)
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Concluído.", $AppTitle)
    }

    if ($script:CurrentLogFile -and $optionFlags.openLogFolderAfterFinish) {
        try {
            Start-Process explorer.exe "/select,`"$script:CurrentLogFile`""
        }
        catch {
        }
    }

})

$appState = Load-AppState
Update-FormatOptions

if ($appState) {
    try {
        if ($appState.Window) {
            if ($appState.Window.Width -and $appState.Window.Height) {
                $form.Size = New-Object System.Drawing.Size([int]$appState.Window.Width, [int]$appState.Window.Height)
            }
            if (($null -ne $appState.Window.Left) -and ($null -ne $appState.Window.Top)) {
                $form.StartPosition = "Manual"
                $form.Location = New-Object System.Drawing.Point([int]$appState.Window.Left, [int]$appState.Window.Top)
            }
        }

        if ($appState.Theme) { $script:CurrentTheme = $appState.Theme }
        if ($appState.Format) { $cmbFormat.SelectedItem = $appState.Format }
        Update-FormatOptions

        if ($appState.Level) { $cmbLevel.SelectedItem = $appState.Level }
        if ($appState.Method) { $cmbMethod.SelectedItem = $appState.Method }
        if ($appState.DictionarySize) { $cmbDict.SelectedItem = $appState.DictionarySize }
        if ($appState.WordSize) { $cmbWord.SelectedItem = $appState.WordSize }
        if ($appState.SolidBlock) { $cmbSolid.SelectedItem = $appState.SolidBlock }
        if ($appState.Threads) { $numThreads.Value = [decimal]$appState.Threads }
        if ($appState.VolumeSize) { $txtVolume.Text = $appState.VolumeSize }
        if ($appState.UpdateMode) { $cmbUpdateMode.SelectedItem = $appState.UpdateMode }
        if ($appState.PathMode) { $cmbPathMode.SelectedItem = $appState.PathMode }
        if ($null -ne $appState.CreateSfx) { $chkSfx.Checked = [bool]$appState.CreateSfx }
        if ($null -ne $appState.CompressShared) { $chkShared.Checked = [bool]$appState.CompressShared }
        if ($null -ne $appState.DeleteAfter) { $chkDelete.Checked = [bool]$appState.DeleteAfter }
        if ($appState.EncryptionMethod) { $cmbEncMethod.SelectedItem = $appState.EncryptionMethod }
        if ($null -ne $appState.EncryptHeaders) { $chkEncryptHeaders.Checked = [bool]$appState.EncryptHeaders }
        if ($appState.ExtraParams) { $txtExtraParams.Text = $appState.ExtraParams }

        if ($appState.Options) {
            if ($null -ne $appState.Options.openLogFolderAfterFinish) {
                $optionFlags.openLogFolderAfterFinish = [bool]$appState.Options.openLogFolderAfterFinish
            }
            if ($null -ne $appState.Options.confirmBeforeClearList) {
                $optionFlags.confirmBeforeClearList = [bool]$appState.Options.confirmBeforeClearList
            }
        }
    }
    catch {
    }
}

ApplyTheme -ThemeName $script:CurrentTheme
Update-DynamicOptions
Refresh-UiState

$form.Add_FormClosing({
    $state = @{
        Window = @{
            Left = $form.Left
            Top = $form.Top
            Width = $form.Width
            Height = $form.Height
        }
        Theme = $script:CurrentTheme
        Format = [string]$cmbFormat.SelectedItem
        Level = [string]$cmbLevel.SelectedItem
        Method = [string]$cmbMethod.SelectedItem
        DictionarySize = [string]$cmbDict.SelectedItem
        WordSize = [string]$cmbWord.SelectedItem
        SolidBlock = [string]$cmbSolid.SelectedItem
        Threads = [int]$numThreads.Value
        VolumeSize = $txtVolume.Text
        UpdateMode = [string]$cmbUpdateMode.SelectedItem
        PathMode = [string]$cmbPathMode.SelectedItem
        CreateSfx = $chkSfx.Checked
        CompressShared = $chkShared.Checked
        DeleteAfter = $chkDelete.Checked
        EncryptionMethod = [string]$cmbEncMethod.SelectedItem
        EncryptHeaders = $chkEncryptHeaders.Checked
        ExtraParams = $txtExtraParams.Text
        Options = @{
            openLogFolderAfterFinish = $optionFlags.openLogFolderAfterFinish
            confirmBeforeClearList   = $optionFlags.confirmBeforeClearList
        }
    }
    Save-AppState -State $state
})

[void]$form.ShowDialog()