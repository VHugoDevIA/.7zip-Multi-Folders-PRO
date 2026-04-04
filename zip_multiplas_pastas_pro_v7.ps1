
# ZIP Múltiplas Pastas/Ficheiros - PRO v7.0
# Permite inserir múltiplos ficheiros para compactar
# Cria um ZIP por cada pasta onde estão os ficheiros
# Acrescenta opção de eliminar ficheiros após compressão
# Se "comprimir ficheiros partilhados" estiver ativo, não elimina ficheiros no fim

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$AppTitle = "ZIP Multiplas Pasta/Ficheiros - PRO v7.0"
$MailSubject = "ZIP Multiplas Pasta/Ficheiros - PRO v7.0"
$StateFile = Join-Path -Path $env:APPDATA -ChildPath "VHugoDevIA\zip_multiplas_pastas_pro_v7_state.json"
$script:CurrentLogFile = $null
$script:ComputerSuffix = if ($env:COMPUTERNAME -and $env:COMPUTERNAME.Length -ge 4) { $env:COMPUTERNAME.Substring($env:COMPUTERNAME.Length - 4).ToUpper() } elseif ($env:COMPUTERNAME) { $env:COMPUTERNAME.ToUpper() } else { "PC" }

# --- Listas separadas para pastas e ficheiros ---
$selectedFolders = New-Object 'System.Collections.Generic.List[string]'
$selectedFiles = New-Object 'System.Collections.Generic.List[string]'

$optionFlags = @{
    openLogFolderAfterFinish = $true
    confirmBeforeClearList   = $false
}

# --- Funções auxiliares ---
function Add-LogLine {
    param([System.Windows.Forms.TextBox]$TextBox, [string]$Text)
    $TextBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Text`r`n")
    $TextBox.SelectionStart = $TextBox.Text.Length
    $TextBox.ScrollToCaret()
}

function Refresh-UiState {
    $btnRun.Enabled = ($selectedFolders.Count -gt 0 -or $selectedFiles.Count -gt 0)
    $btnRemove.Enabled = ($listBox.SelectedItems.Count -gt 0)
    $btnRemoveFile.Enabled = ($listBox.SelectedItems.Count -gt 0)
    $btnClear.Enabled = ($selectedFolders.Count -gt 0 -or $selectedFiles.Count -gt 0)
    $chkDeleteFiles.Enabled = -not $chkShared.Checked
    if ($chkShared.Checked) { $chkDeleteFiles.Checked = $false }
}

function Update-DynamicOptions {
    $format = [string]$cmbFormat.SelectedItem
    $method = [string]$cmbMethod.SelectedItem

    $cmbSolid.Enabled = ($format -eq "7z")
    $chkEncryptHeaders.Enabled = ($format -eq "7z")
    $chkSFX.Enabled = ($format -eq "7z")
    $cmbEncMethod.Enabled = $true

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
    } else {
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

function Build-SevenZipArguments {
    param(
        [string]$ArchivePath,
        [string]$FolderName,
        [string]$Format,
        [string]$Level,
        [int]$Threads,
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
    $args.Add("-mmt=$Threads")
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
                "AES128" { $args.Add("-mem=AES128") }
                "AES192" { $args.Add("-mem=AES192") }
                "AES256" { $args.Add("-mem=AES256") }
            }
        } elseif ($Format -eq "7z") {
            if ($EncryptHeaders) { $args.Add("-mhe=on") }
        }
    }

    if ($Format -eq "zip") {
        if ($Method -and ($Method -ne "Automatico")) { $args.Add("-mm=$Method") }

        switch ($Method) {
            "Deflate" {
                if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
            }
            "Deflate64" {
                if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
            }
            "BZip2" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
            }
            "LZMA" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
            }
            "PPMd" {
                if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
                    $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
                    $args.Add("-md=$dict")
                }
                if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
            }
        }
    } elseif ($Format -eq "7z") {
        if ($Method -and ($Method -ne "Automatico")) { $args.Add("-mm=$Method") }
        if ($DictionarySize -and ($DictionarySize -ne "Automatico")) {
            $dict = $DictionarySize.ToLower() -replace "kb","k" -replace "mb","m"
            $args.Add("-md=$dict")
        }
        if ($WordSize -and ($WordSize -ne "Automatico")) { $args.Add("-mfb=$WordSize") }
        if ($SolidBlock -and ($SolidBlock -ne "Automatico")) {
            switch ($SolidBlock) {
                "Off" { $args.Add("-ms=off") }
                "On" { $args.Add("-ms=on") }
                default { $args.Add("-ms=$SolidBlock") }
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ExtraParams)) {
        foreach ($piece in ($ExtraParams.Trim() -split '\s+')) {
            if ($piece.Trim().Length -gt 0) { $args.Add($piece.Trim()) }
        }
    }

    return ,$args
}

function Save-State {
    $state = @{
        windowSize = @{ Width = $form.Width; Height = $form.Height }
        windowLocation = @{ X = $form.Location.X; Y = $form.Location.Y }
        format = $cmbFormat.SelectedItem
        compression = $cmbCompression.SelectedItem
        method = $cmbMethod.SelectedItem
        dictSize = $cmbDict.SelectedItem
        wordSize = $cmbWord.SelectedItem
        solidBlocks = $cmbSolid.SelectedItem
        threads = $numThreads.Value
        volumeSize = $txtVolume.Text
        updateMode = $cmbUpdateMode.SelectedItem
        pathMode = $cmbPathMode.SelectedItem
        createSfx = $chkSfx.Checked
        password = $txtPass1.Text
        password2 = $txtPass2.Text
        encMethod = $cmbEncMethod.SelectedItem
        encryptHeaders = $chkEncryptHeaders.Checked
        deleteAfterCompress = $chkDeleteFiles.Checked
        sharedFiles = $chkShared.Checked
        relativePath = $chkRelative.Checked
        extraParams = $txtExtra.Text
        outputFolder = $txtOutput.Text
        openLogFolder = $optionFlags.openLogFolderAfterFinish
        confirmClear = $optionFlags.confirmBeforeClearList
    }
    $stateDir = Split-Path -Parent $StateFile
    if (-not (Test-Path $stateDir)) { New-Item -ItemType Directory -Path $stateDir -Force | Out-Null }
    $state | ConvertTo-Json | Set-Content -Path $StateFile -Encoding UTF8
}

function Load-State {
    if (Test-Path $StateFile) {
        try {
            $state = Get-Content -Path $StateFile -Encoding UTF8 | ConvertFrom-Json
            if ($state.windowSize) { $form.Size = New-Object System.Drawing.Size($state.windowSize.Width, $state.windowSize.Height) }
            if ($state.windowLocation) { $form.Location = New-Object System.Drawing.Point($state.windowLocation.X, $state.windowLocation.Y) }
            if ($state.format) { $cmbFormat.SelectedItem = $state.format }
            if ($state.compression) { $cmbCompression.SelectedItem = $state.compression }
            if ($state.method) { $cmbMethod.SelectedItem = $state.method }
            if ($state.dictSize) { $cmbDict.SelectedItem = $state.dictSize }
            if ($state.wordSize) { $cmbWord.SelectedItem = $state.wordSize }
            if ($state.solidBlocks) { $cmbSolid.SelectedItem = $state.solidBlocks }
            if ($state.threads) { $numThreads.Value = [int]$state.threads }
            if ($state.volumeSize) { $txtVolume.Text = $state.volumeSize }
            if ($state.updateMode) { $cmbUpdateMode.SelectedItem = $state.updateMode }
            if ($state.pathMode) { $cmbPathMode.SelectedItem = $state.pathMode }
            if ($state.createSfx) { $chkSFX.Checked = $state.createSfx }
            if ($state.password) { $txtPass1.Text = $state.password }
            if ($state.password2) { $txtPass2.Text = $state.password2 }
            if ($state.encMethod) { $cmbEncMethod.SelectedItem = $state.encMethod }
            if ($state.encryptHeaders) { $chkEncryptHeaders.Checked = $state.encryptHeaders }
            if ($state.deleteAfterCompress) { $chkDeleteFiles.Checked = $state.deleteAfterCompress }
            if ($state.sharedFiles) { $chkShared.Checked = $state.sharedFiles }
            if ($state.relativePath) { $chkRelative.Checked = $state.relativePath }
            if ($state.extraParams) { $txtExtra.Text = $state.extraParams }
            if ($state.outputFolder) { $txtOutput.Text = $state.outputFolder }
            if ($state.openLogFolder) { $optionFlags.openLogFolderAfterFinish = $state.openLogFolder }
            if ($state.confirmClear) { $optionFlags.confirmBeforeClearList = $state.confirmClear }
        } catch { }
    }
}

# --- Criar form e controles ---
$form = New-Object System.Windows.Forms.Form
$form.Text = $AppTitle
$form.Size = New-Object System.Drawing.Size(1080, 950)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.AllowDrop = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.Add_FormClosing({ Save-State })

# --- GroupBox para Configuração de compressão ---
$grpCompression = New-Object System.Windows.Forms.GroupBox
$grpCompression.Text = "Configuração de compressão"
$grpCompression.Location = New-Object System.Drawing.Point(20, 20)
$grpCompression.Size = New-Object System.Drawing.Size(700, 230)
$form.Controls.Add($grpCompression)

# --- Labels e ComboBoxes dentro do grpCompression ---
$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = "Formato:"
$lblFormat.Location = New-Object System.Drawing.Point(15, 25)
$lblFormat.Size = New-Object System.Drawing.Size(60, 20)
$grpCompression.Controls.Add($lblFormat)

$cmbFormat = New-Object System.Windows.Forms.ComboBox
$cmbFormat.Location = New-Object System.Drawing.Point(15, 45)
$cmbFormat.Size = New-Object System.Drawing.Size(120, 23)
$cmbFormat.DropDownStyle = "DropDownList"
$cmbFormat.Items.AddRange(@("zip", "7z"))
$cmbFormat.SelectedIndex = 0
$grpCompression.Controls.Add($cmbFormat)

$lblCompression = New-Object System.Windows.Forms.Label
$lblCompression.Text = "Compressão:"
$lblCompression.Location = New-Object System.Drawing.Point(160, 25)
$lblCompression.Size = New-Object System.Drawing.Size(80, 20)
$grpCompression.Controls.Add($lblCompression)

$cmbCompression = New-Object System.Windows.Forms.ComboBox
$cmbCompression.Location = New-Object System.Drawing.Point(160, 45)
$cmbCompression.Size = New-Object System.Drawing.Size(120, 23)
$cmbCompression.DropDownStyle = "DropDownList"
$cmbCompression.Items.AddRange(@("0", "1", "3", "5", "7", "9"))
$cmbCompression.SelectedIndex = 3
$grpCompression.Controls.Add($cmbCompression)

$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Método:"
$lblMethod.Location = New-Object System.Drawing.Point(310, 25)
$lblMethod.Size = New-Object System.Drawing.Size(60, 20)
$grpCompression.Controls.Add($lblMethod)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Location = New-Object System.Drawing.Point(310, 45)
$cmbMethod.Size = New-Object System.Drawing.Size(120, 23)
$cmbMethod.DropDownStyle = "DropDownList"
$grpCompression.Controls.Add($cmbMethod)

$lblDict = New-Object System.Windows.Forms.Label
$lblDict.Text = "Dicionário:"
$lblDict.Location = New-Object System.Drawing.Point(15, 78)
$lblDict.Size = New-Object System.Drawing.Size(70, 20)
$grpCompression.Controls.Add($lblDict)

$cmbDict = New-Object System.Windows.Forms.ComboBox
$cmbDict.Location = New-Object System.Drawing.Point(15, 98)
$cmbDict.Size = New-Object System.Drawing.Size(120, 23)
$cmbDict.DropDownStyle = "DropDownList"
$grpCompression.Controls.Add($cmbDict)

$lblWord = New-Object System.Windows.Forms.Label
$lblWord.Text = "Tamanho palavra:"
$lblWord.Location = New-Object System.Drawing.Point(160, 78)
$lblWord.Size = New-Object System.Drawing.Size(100, 20)
$grpCompression.Controls.Add($lblWord)

$cmbWord = New-Object System.Windows.Forms.ComboBox
$cmbWord.Location = New-Object System.Drawing.Point(160, 98)
$cmbWord.Size = New-Object System.Drawing.Size(120, 23)
$cmbWord.DropDownStyle = "DropDownList"
$grpCompression.Controls.Add($cmbWord)

$lblSolid = New-Object System.Windows.Forms.Label
$lblSolid.Text = "Blocos sólidos:"
$lblSolid.Location = New-Object System.Drawing.Point(310, 78)
$lblSolid.Size = New-Object System.Drawing.Size(90, 20)
$grpCompression.Controls.Add($lblSolid)

$cmbSolid = New-Object System.Windows.Forms.ComboBox
$cmbSolid.Location = New-Object System.Drawing.Point(310, 98)
$cmbSolid.Size = New-Object System.Drawing.Size(120, 23)
$cmbSolid.DropDownStyle = "DropDownList"
$grpCompression.Controls.Add($cmbSolid)

$lblUpdateMode = New-Object System.Windows.Forms.Label
$lblUpdateMode.Text = "Se o ZIP já existir:"
$lblUpdateMode.Location = New-Object System.Drawing.Point(15, 128)
$lblUpdateMode.Size = New-Object System.Drawing.Size(120, 20)
$grpCompression.Controls.Add($lblUpdateMode)

$cmbUpdateMode = New-Object System.Windows.Forms.ComboBox
$cmbUpdateMode.Location = New-Object System.Drawing.Point(15, 148)
$cmbUpdateMode.Size = New-Object System.Drawing.Size(120, 23)
$cmbUpdateMode.DropDownStyle = "DropDownList"
$cmbUpdateMode.Items.AddRange(@("Adicionar", "Atualizar", "Sincronizar", "Substituir"))
$cmbUpdateMode.SelectedIndex = 0
$grpCompression.Controls.Add($cmbUpdateMode)

$lblPathMode = New-Object System.Windows.Forms.Label
$lblPathMode.Text = "Passar pasta ao 7-Zip:"
$lblPathMode.Location = New-Object System.Drawing.Point(160, 128)
$lblPathMode.Size = New-Object System.Drawing.Size(135, 20)
$grpCompression.Controls.Add($lblPathMode)

$cmbPathMode = New-Object System.Windows.Forms.ComboBox
$cmbPathMode.Location = New-Object System.Drawing.Point(160, 148)
$cmbPathMode.Size = New-Object System.Drawing.Size(120, 23)
$cmbPathMode.DropDownStyle = "DropDownList"
$cmbPathMode.Items.AddRange(@("Mantém caminho", "Relativo", "Sem caminho"))
$cmbPathMode.SelectedIndex = 0
$grpCompression.Controls.Add($cmbPathMode)

$lblThreads = New-Object System.Windows.Forms.Label
$lblThreads.Text = "Threads:"
$lblThreads.Location = New-Object System.Drawing.Point(310, 128)
$lblThreads.Size = New-Object System.Drawing.Size(60, 20)
$grpCompression.Controls.Add($lblThreads)

$numThreads = New-Object System.Windows.Forms.NumericUpDown
$numThreads.Location = New-Object System.Drawing.Point(310, 148)
$numThreads.Size = New-Object System.Drawing.Size(120, 23)
$numThreads.Minimum = 0
$numThreads.Maximum = 32
$numThreads.Value = 0
$grpCompression.Controls.Add($numThreads)

# --- GroupBox para Destino / extras ---
$grpOutput = New-Object System.Windows.Forms.GroupBox
$grpOutput.Text = "Destino / extras"
$grpOutput.Location = New-Object System.Drawing.Point(20, 260)
$grpOutput.Size = New-Object System.Drawing.Size(700, 110)
$form.Controls.Add($grpOutput)

# --- Labels e TextBoxes dentro do grpOutput ---
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Pasta de saída:"
$lblOutput.Location = New-Object System.Drawing.Point(15, 25)
$lblOutput.Size = New-Object System.Drawing.Size(90, 20)
$grpOutput.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(15, 45)
$txtOutput.Size = New-Object System.Drawing.Size(470, 23)
$grpOutput.Controls.Add($txtOutput)

$lblVolume = New-Object System.Windows.Forms.Label
$lblVolume.Text = "Volume (ex: 100m):"
$lblVolume.Location = New-Object System.Drawing.Point(500, 25)
$lblVolume.Size = New-Object System.Drawing.Size(110, 20)
$grpOutput.Controls.Add($lblVolume)

$txtVolume = New-Object System.Windows.Forms.TextBox
$txtVolume.Location = New-Object System.Drawing.Point(500, 45)
$txtVolume.Size = New-Object System.Drawing.Size(170, 23)
$grpOutput.Controls.Add($txtVolume)

$lblExtra = New-Object System.Windows.Forms.Label
$lblExtra.Text = "Parâmetros extra:"
$lblExtra.Location = New-Object System.Drawing.Point(15, 75)
$lblExtra.Size = New-Object System.Drawing.Size(100, 20)
$grpOutput.Controls.Add($lblExtra)

$txtExtra = New-Object System.Windows.Forms.TextBox
$txtExtra.Location = New-Object System.Drawing.Point(15, 95)
$txtExtra.Size = New-Object System.Drawing.Size(655, 23)
$grpOutput.Controls.Add($txtExtra)

# --- GroupBox de Encriptação ---
$grpEnc = New-Object System.Windows.Forms.GroupBox
$grpEnc.Text = "Encriptação"
$grpEnc.Location = New-Object System.Drawing.Point(740, 20)
$grpEnc.Size = New-Object System.Drawing.Size(320, 240)
$form.Controls.Add($grpEnc)

$lblPass1 = New-Object System.Windows.Forms.Label
$lblPass1.Text = "Password:"
$lblPass1.Location = New-Object System.Drawing.Point(15, 25)
$lblPass1.Size = New-Object System.Drawing.Size(70, 20)
$grpEnc.Controls.Add($lblPass1)

$txtPass1 = New-Object System.Windows.Forms.TextBox
$txtPass1.Location = New-Object System.Drawing.Point(15, 45)
$txtPass1.Size = New-Object System.Drawing.Size(290, 23)
$txtPass1.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass1)

$lblPass2 = New-Object System.Windows.Forms.Label
$lblPass2.Text = "Repetir password:"
$lblPass2.Location = New-Object System.Drawing.Point(15, 75)
$lblPass2.Size = New-Object System.Drawing.Size(110, 20)
$grpEnc.Controls.Add($lblPass2)

$txtPass2 = New-Object System.Windows.Forms.TextBox
$txtPass2.Location = New-Object System.Drawing.Point(15, 95)
$txtPass2.Size = New-Object System.Drawing.Size(290, 23)
$txtPass2.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass2)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Mostrar password"
$chkShowPass.Location = New-Object System.Drawing.Point(15, 125)
$chkShowPass.Size = New-Object System.Drawing.Size(130, 20)
$grpEnc.Controls.Add($chkShowPass)

$lblEncMethod = New-Object System.Windows.Forms.Label
$lblEncMethod.Text = "Método encriptação:"
$lblEncMethod.Location = New-Object System.Drawing.Point(15, 155)
$lblEncMethod.Size = New-Object System.Drawing.Size(120, 20)
$grpEnc.Controls.Add($lblEncMethod)

$cmbEncMethod = New-Object System.Windows.Forms.ComboBox
$cmbEncMethod.Location = New-Object System.Drawing.Point(15, 175)
$cmbEncMethod.Size = New-Object System.Drawing.Size(120, 23)
$cmbEncMethod.DropDownStyle = "DropDownList"
$grpEnc.Controls.Add($cmbEncMethod)

$chkEncryptHeaders = New-Object System.Windows.Forms.CheckBox
$chkEncryptHeaders.Text = "Encriptar nomes ficheiros"
$chkEncryptHeaders.Location = New-Object System.Drawing.Point(15, 205)
$chkEncryptHeaders.Size = New-Object System.Drawing.Size(190, 20)
$grpEnc.Controls.Add($chkEncryptHeaders)

# --- GroupBox para Opções ---
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = "Opções"
$grpOptions.Location = New-Object System.Drawing.Point(740, 270)
$grpOptions.Size = New-Object System.Drawing.Size(320, 110)
$form.Controls.Add($grpOptions)

$chkSFX = New-Object System.Windows.Forms.CheckBox
$chkSFX.Text = "SFX (7z)"
$chkSFX.Location = New-Object System.Drawing.Point(15, 25)
$chkSFX.Size = New-Object System.Drawing.Size(90, 20)
$grpOptions.Controls.Add($chkSFX)

$chkDeleteFiles = New-Object System.Windows.Forms.CheckBox
$chkDeleteFiles.Text = "Eliminar ficheiros após compressão"
$chkDeleteFiles.Location = New-Object System.Drawing.Point(15, 50)
$chkDeleteFiles.Size = New-Object System.Drawing.Size(220, 20)
$grpOptions.Controls.Add($chkDeleteFiles)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = "Comprimir ficheiros partilhados"
$chkShared.Location = New-Object System.Drawing.Point(15, 75)
$chkShared.Size = New-Object System.Drawing.Size(220, 20)
$chkShared.Add_CheckedChanged({ Refresh-UiState })
$grpOptions.Controls.Add($chkShared)

$chkRelative = New-Object System.Windows.Forms.CheckBox
$chkRelative.Text = "Caminho relativo"
$chkRelative.Location = New-Object System.Drawing.Point(15, 95)
$chkRelative.Size = New-Object System.Drawing.Size(220, 20)
$grpOptions.Controls.Add($chkRelative)

$cmbFormat.Add_SelectedIndexChanged({ Update-FormatOptions })
$cmbMethod.Add_SelectedIndexChanged({ Update-DynamicOptions })
$chkShowPass.Add_CheckedChanged({
    $txtPass1.UseSystemPasswordChar = -not $chkShowPass.Checked
    $txtPass2.UseSystemPasswordChar = -not $chkShowPass.Checked
})

# --- Painel para seleção ---
$panelSelect = New-Object System.Windows.Forms.Panel
$panelSelect.Location = New-Object System.Drawing.Point(20, 390)
$panelSelect.Size = New-Object System.Drawing.Size(700, 240)
$form.Controls.Add($panelSelect)

# ListBox único para ficheiros e pastas
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(0, 0)
$listBox.Size = New-Object System.Drawing.Size(700, 240)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$listBox.Anchor = "Top, Left, Right"
$panelSelect.Controls.Add($listBox)

# --- Painel lateral para botões ---
$panelBtns = New-Object System.Windows.Forms.Panel
$panelBtns.Location = New-Object System.Drawing.Point(740, 390)
$panelBtns.Size = New-Object System.Drawing.Size(320, 420)
$panelBtns.Anchor = "Top, Right"
$form.Controls.Add($panelBtns)

# --- Botões ---
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Adicionar pasta(s)"
$btnAdd.Size = New-Object System.Drawing.Size(130, 32)
$btnAdd.FlatStyle = "Flat"
$btnAdd.BackColor = [System.Drawing.Color]::FromArgb(186, 225, 255)
$btnAdd.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "Adicionar ficheiro(s)"
$btnAddFile.Size = New-Object System.Drawing.Size(130, 32)
$btnAddFile.FlatStyle = "Flat"
$btnAddFile.BackColor = [System.Drawing.Color]::FromArgb(186, 225, 255)
$btnAddFile.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remover pasta"
$btnRemove.Size = New-Object System.Drawing.Size(130, 32)
$btnRemove.FlatStyle = "Flat"
$btnRemove.BackColor = [System.Drawing.Color]::FromArgb(255, 179, 186)
$btnRemove.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnRemoveFile = New-Object System.Windows.Forms.Button
$btnRemoveFile.Text = "Remover ficheiro"
$btnRemoveFile.Size = New-Object System.Drawing.Size(130, 32)
$btnRemoveFile.FlatStyle = "Flat"
$btnRemoveFile.BackColor = [System.Drawing.Color]::FromArgb(255, 179, 186)
$btnRemoveFile.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Limpar lista"
$btnClear.Size = New-Object System.Drawing.Size(130, 32)
$btnClear.FlatStyle = "Flat"
$btnClear.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 59)
$btnClear.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text = "Limpar log"
$btnClearLog.Size = New-Object System.Drawing.Size(130, 32)
$btnClearLog.FlatStyle = "Flat"
$btnClearLog.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 59)
$btnClearLog.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnOptions = New-Object System.Windows.Forms.Button
$btnOptions.Text = "Opções"
$btnOptions.Size = New-Object System.Drawing.Size(130, 32)
$btnOptions.FlatStyle = "Flat"
$btnOptions.BackColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnOptions.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnCancelMain = New-Object System.Windows.Forms.Button
$btnCancelMain.Text = "Cancelar"
$btnCancelMain.Size = New-Object System.Drawing.Size(130, 32)
$btnCancelMain.FlatStyle = "Flat"
$btnCancelMain.BackColor = [System.Drawing.Color]::FromArgb(255, 87, 87)
$btnCancelMain.ForeColor = [System.Drawing.Color]::White

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Iniciar"
$btnRun.Size = New-Object System.Drawing.Size(130, 32)
$btnRun.FlatStyle = "Flat"
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
$btnRun.ForeColor = [System.Drawing.Color]::White

# Adicionar botões ao painel
$panelBtns.Controls.AddRange(@($btnAdd, $btnAddFile, $btnRemove, $btnRemoveFile, $btnClear, $btnClearLog, $btnOptions, $btnCancelMain, $btnRun))

# Posicionar botões verticalmente com espaçamento uniforme
$btnAdd.Location = New-Object System.Drawing.Point(25, 10)
$btnAddFile.Location = New-Object System.Drawing.Point(25, 55)
$btnRemove.Location = New-Object System.Drawing.Point(25, 100)
$btnRemoveFile.Location = New-Object System.Drawing.Point(25, 145)
$btnClear.Location = New-Object System.Drawing.Point(25, 190)
$btnClearLog.Location = New-Object System.Drawing.Point(25, 235)
$btnOptions.Location = New-Object System.Drawing.Point(25, 280)
$btnCancelMain.Location = New-Object System.Drawing.Point(25, 325)
$btnRun.Location = New-Object System.Drawing.Point(25, 370)

# --- Progress bars ---
$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20, 520)
$progressFolder.Size = New-Object System.Drawing.Size(700, 20)
$form.Controls.Add($progressFolder)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20, 550)
$progressTotal.Size = New-Object System.Drawing.Size(700, 20)
$form.Controls.Add($progressTotal)

# --- TextBox para log ---
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 580)
$txtLog.Size = New-Object System.Drawing.Size(900, 320)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Anchor = "Top, Bottom, Left, Right"
$form.Controls.Add($txtLog)

# --- Event handlers ---

# Botão para adicionar pastas
$btnAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecione uma ou mais pastas para adicionar"
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $folder = $dlg.SelectedPath
        if (-not $selectedFolders.Contains($folder)) {
            $selectedFolders.Add($folder)
            [void]$listBox.Items.Add("[Pasta] $folder")
            Add-LogLine -TextBox $txtLog -Text "Pasta adicionada: $folder"
        }
    }
    Refresh-UiState
})

# Botão para adicionar ficheiros
$btnAddFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Selecione um ou mais ficheiros para adicionar"
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $dlg.FileNames) {
            if (-not $selectedFiles.Contains($file)) {
                $selectedFiles.Add($file)
                [void]$listBox.Items.Add("[Ficheiro] $file")
                Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado: $file"
            }
        }
    }
    Refresh-UiState
})

# Botão para remover pastas
$btnRemove.Add_Click({
    if ($listBox.SelectedItems.Count -eq 0) { return }
    $toRemove = @()
    foreach ($item in $listBox.SelectedItems) {
        if ($item -like '[Pasta]*') {
            $folderPath = $item -replace '^\[Pasta\] ', ''
            $toRemove += $folderPath
        }
    }
    foreach ($folderPath in $toRemove) {
        [void]$selectedFolders.Remove($folderPath)
        [void]$listBox.Items.Remove("[Pasta] $folderPath")
        Add-LogLine -TextBox $txtLog -Text "Pasta removida: $folderPath"
    }
    Refresh-UiState
})

# Botão para remover ficheiros
$btnRemoveFile.Add_Click({
    if ($listBox.SelectedItems.Count -eq 0) { return }
    $toRemove = @()
    foreach ($item in $listBox.SelectedItems) {
        if ($item -like '[Ficheiro]*') {
            $filePath = $item -replace '^\[Ficheiro\] ', ''
            $toRemove += $filePath
        }
    }
    foreach ($filePath in $toRemove) {
        [void]$selectedFiles.Remove($filePath)
        [void]$listBox.Items.Remove("[Ficheiro] $filePath")
        Add-LogLine -TextBox $txtLog -Text "Ficheiro removido: $filePath"
    }
    Refresh-UiState
})

# Botão para limpar lista
$btnClear.Add_Click({
    if ($optionFlags.confirmBeforeClearList) {
        $result = [System.Windows.Forms.MessageBox]::Show("Tem certeza que deseja limpar a lista?", "Confirmação", "YesNo", "Question")
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $selectedFolders.Clear()
    $selectedFiles.Clear()
    $listBox.Items.Clear()
    Add-LogLine -TextBox $txtLog -Text "Lista limpa"
    Refresh-UiState
})

# Botão para limpar log
$btnClearLog.Add_Click({
    $txtLog.Clear()
})

# Botão de opções
$btnOptions.Add_Click({
    $formOptions = New-Object System.Windows.Forms.Form
    $formOptions.Text = "Opções"
    $formOptions.Size = New-Object System.Drawing.Size(300, 200)
    $formOptions.StartPosition = "CenterParent"
    $formOptions.FormBorderStyle = "FixedDialog"
    $formOptions.MaximizeBox = $false

    $chkOpenLog = New-Object System.Windows.Forms.CheckBox
    $chkOpenLog.Text = "Abrir pasta do log após terminar"
    $chkOpenLog.Location = New-Object System.Drawing.Point(20, 20)
    $chkOpenLog.Checked = $optionFlags.openLogFolderAfterFinish
    $formOptions.Controls.Add($chkOpenLog)

    $chkConfirmClear = New-Object System.Windows.Forms.CheckBox
    $chkConfirmClear.Text = "Confirmar antes de limpar lista"
    $chkConfirmClear.Location = New-Object System.Drawing.Point(20, 50)
    $chkConfirmClear.Checked = $optionFlags.confirmBeforeClearList
    $formOptions.Controls.Add($chkConfirmClear)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(100, 120)
    $btnOK.Add_Click({
        $optionFlags.openLogFolderAfterFinish = $chkOpenLog.Checked
        $optionFlags.confirmBeforeClearList = $chkConfirmClear.Checked
        $formOptions.Close()
    })
    $formOptions.Controls.Add($btnOK)

    [void]$formOptions.ShowDialog()
})

# Botão cancelar
$btnCancelMain.Add_Click({
    $script:CancelRequested = $true
    Add-LogLine -TextBox $txtLog -Text "[CANCELADO] Cancelamento solicitado pelo utilizador"
})

# Drag&Drop para ficheiros e pastas no listBox único
$listBox.Add_DragDrop({
    $items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($item in $items) {
        if (Test-Path -LiteralPath $item -PathType Container) {
            if (-not $selectedFolders.Contains($item)) {
                $selectedFolders.Add($item)
                [void]$listBox.Items.Add("[Pasta] $item")
                Add-LogLine -TextBox $txtLog -Text "Pasta adicionada por arrastar: $item"
            }
        } elseif (Test-Path -LiteralPath $item -PathType Leaf) {
            if (-not $selectedFiles.Contains($item)) {
                $selectedFiles.Add($item)
                [void]$listBox.Items.Add("[Ficheiro] $item")
                Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado por arrastar: $item"
            }
        }
    }
    Refresh-UiState
})

$listBox.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    } else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

# --- Processamento ---
$btnRun.Add_Click({
    if ($selectedFolders.Count -eq 0 -and $selectedFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Adiciona pelo menos uma pasta ou ficheiro.")
        return
    }

    if ($txtPass1.Text -or $txtPass2.Text) {
        if ($txtPass1.Text -ne $txtPass2.Text) {
            [System.Windows.Forms.MessageBox]::Show("As passwords não coincidem.")
            return
        }
    }

    $script:CancelRequested = $false
    $btnRun.Enabled = $false
    $btnCancelMain.Enabled = $true

    # Criar pasta de log
    $logDir = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "ZIP_Múltiplas_Pastas-PRO LOG"
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:CurrentLogFile = Join-Path -Path $logDir -ChildPath ("Log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")

    # Preparar argumentos base
    $format = [string]$cmbFormat.SelectedItem
    $compression = [string]$cmbCompression.SelectedItem
    $method = [string]$cmbMethod.SelectedItem
    $dictSize = [string]$cmbDict.SelectedItem
    $wordSize = [string]$cmbWord.SelectedItem
    $solidBlocks = [string]$cmbSolid.SelectedItem
    $threads = [int]$numThreads.Value
    $updateMode = [string]$cmbUpdateMode.SelectedItem
    $pathMode = [string]$cmbPathMode.SelectedItem
    $password = $txtPass1.Text
    $encMethod = [string]$cmbEncMethod.SelectedItem
    $encryptHeaders = $chkEncryptHeaders.Checked
    $sfx = $chkSFX.Checked
    $relative = $chkRelative.Checked
    $volume = $txtVolume.Text.Trim()
    $extra = $txtExtra.Text.Trim()
    $outputFolder = $txtOutput.Text.Trim()

    if (-not $outputFolder) {
        if ($selectedFolders.Count -gt 0) {
            $outputFolder = Split-Path -Parent $selectedFolders[0]
        } elseif ($selectedFiles.Count -gt 0) {
            $outputFolder = Split-Path -Parent $selectedFiles[0]
        }
    }

    # Processamento de pastas
    if ($selectedFolders.Count -gt 0) {
        $totalCount = $selectedFolders.Count
        $currentIndex = 0
        foreach ($folder in $selectedFolders) {
            $currentIndex++
            $folderName = [System.IO.Path]::GetFileName($folder)
            $archiveName = "{0}_{1}.{2}" -f $folderName, $script:ComputerSuffix, $format
            $archivePath = Join-Path -Path $outputFolder -ChildPath $archiveName

            # Verificar se arquivo já existe
            if (Test-Path $archivePath) {
                $archivePath = Join-Path -Path $outputFolder -ChildPath ("{0}_{1}_2.{2}" -f $folderName, $script:ComputerSuffix, $format)
                $counter = 3
                while (Test-Path $archivePath) {
                    $archivePath = Join-Path -Path $outputFolder -ChildPath ("{0}_{1}_{2}.{3}" -f $folderName, $script:ComputerSuffix, $counter, $format)
                    $counter++
                }
            }

            $folderArg = if ($pathMode -eq "Mantém caminho") { $folder } else { $folderName }
            $argList = Build-SevenZipArguments `
                -ArchivePath $archivePath `
                -FolderName $folderArg `
                -Format $format `
                -Level $compression `
                -Threads $threads `
                -DeleteAfter $chkDeleteFiles.Checked `
                -CompressShared $chkShared.Checked `
                -CreateSfx $sfx `
                -VolumeSize $volume `
                -Password $password `
                -EncryptionMethod $encMethod `
                -EncryptHeaders $encryptHeaders `
                -Method $method `
                -DictionarySize $dictSize `
                -WordSize $wordSize `
                -SolidBlock $solidBlocks `
                -ExtraParams $extra
            $argString = ($argList -join " ")
            Add-LogLine -TextBox $txtLog -Text "Comprimindo pasta: $folderName"
            Add-LogLine -TextBox $txtLog -Text "Comando: $SevenZip $argString"

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $SevenZip
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
                        if ($pct -ge 0 -and $pct -le 100) { $progressFolder.Value = $pct }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($line)) { Add-LogLine -TextBox $txtLog -Text $line }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                    if (-not [string]::IsNullOrWhiteSpace($errLine)) { Add-LogLine -TextBox $txtLog -Text "[ERRO] $errLine" }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                Start-Sleep -Milliseconds 120
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:CancelRequested) { break }
            }
            $p.WaitForExit()
            $exitCode = $p.ExitCode
            $script:CurrentProcess = $null

            if ($script:CancelRequested) {
                Add-LogLine -TextBox $txtLog -Text "[CANCELADO] Processo interrompido: $folderName"
                break
            } elseif ($exitCode -eq 0) {
                $progressFolder.Value = 100
                Add-LogLine -TextBox $txtLog -Text "[OK] Concluído: $folderName"
                # Eliminar pasta se opção ativa
                if ($chkDeleteFiles.Checked) {
                    try {
                        Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                        Add-LogLine -TextBox $txtLog -Text "[OK] Pasta eliminada: $folder"
                    } catch {
                        Add-LogLine -TextBox $txtLog -Text "[AVISO] Não foi possível eliminar a pasta: $folder"
                    }
                }
            } else {
                Add-LogLine -TextBox $txtLog -Text "[ERRO] Falhou: $folderName (código $exitCode)"
            }
            $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    # Processamento de ficheiros
    if ($selectedFiles.Count -gt 0) {
        $filesByFolder = @{}
        foreach ($file in $selectedFiles) {
            $dir = [System.IO.Path]::GetDirectoryName($file)
            if (-not $filesByFolder.ContainsKey($dir)) { $filesByFolder[$dir] = @() }
            $filesByFolder[$dir] += $file
        }
        $totalCount = $filesByFolder.Keys.Count
        $currentIndex = 0
        foreach ($dir in $filesByFolder.Keys) {
            $currentIndex++
            $files = $filesByFolder[$dir]
            $folderName = [System.IO.Path]::GetFileName($dir)
            $archiveBaseName = "{0}_{1}" -f $folderName, $script:ComputerSuffix
            $archiveName = "{0}.{1}" -f $archiveBaseName, $format
            $archivePath = Join-Path -Path $outputFolder -ChildPath $archiveName

            # Verificar se arquivo já existe
            if (Test-Path $archivePath) {
                $archivePath = Join-Path -Path $outputFolder -ChildPath ("{0}_2.{1}" -f $archiveBaseName, $format)
                $counter = 3
                while (Test-Path $archivePath) {
                    $archivePath = Join-Path -Path $outputFolder -ChildPath ("{0}_{1}.{2}" -f $archiveBaseName, $counter, $format)
                    $counter++
                }
            }

            $fileArgs = @()
            foreach ($file in $files) {
                $fileArg = if ($pathMode -eq "Mantém caminho") { $file } else { [System.IO.Path]::GetFileName($file) }
                $fileArgs += "`"$fileArg`""
            }

            $argList = Build-SevenZipArguments `
                -ArchivePath $archivePath `
                -FolderName $fileArgs[0] `
                -Format $format `
                -Level $compression `
                -Threads $threads `
                -DeleteAfter $chkDeleteFiles.Checked `
                -CompressShared $chkShared.Checked `
                -CreateSfx $sfx `
                -VolumeSize $volume `
                -Password $password `
                -EncryptionMethod $encMethod `
                -EncryptHeaders $encryptHeaders `
                -Method $method `
                -DictionarySize $dictSize `
                -WordSize $wordSize `
                -SolidBlock $solidBlocks `
                -ExtraParams $extra
            if ($fileArgs.Count -gt 1) { $argList.AddRange($fileArgs[1..($fileArgs.Count - 1)]) }
            $argString = ($argList -join " ")
            Add-LogLine -TextBox $txtLog -Text "Comprimindo ficheiros da pasta: $folderName"
            Add-LogLine -TextBox $txtLog -Text "Comando: $SevenZip $argString"

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $SevenZip
            $psi.WorkingDirectory = $dir
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
                        if ($pct -ge 0 -and $pct -le 100) { $progressFolder.Value = $pct }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($line)) { Add-LogLine -TextBox $txtLog -Text $line }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                    if (-not [string]::IsNullOrWhiteSpace($errLine)) { Add-LogLine -TextBox $txtLog -Text "[ERRO] $errLine" }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                Start-Sleep -Milliseconds 120
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:CancelRequested) { break }
            }
            $p.WaitForExit()
            $exitCode = $p.ExitCode
            $script:CurrentProcess = $null

            if ($script:CancelRequested) {
                Add-LogLine -TextBox $txtLog -Text "[CANCELADO] Processo interrompido: $folderName"
                break
            } elseif ($exitCode -eq 0) {
                $progressFolder.Value = 100
                Add-LogLine -TextBox $txtLog -Text "[OK] Concluído: $folderName"
                # Eliminar ficheiros se opção ativa e não for partilhado
                if ($chkDeleteFiles.Checked -and -not $chkShared.Checked) {
                    foreach ($f in $files) {
                        try {
                            Remove-Item -LiteralPath $f -Force -ErrorAction Stop
                            Add-LogLine -TextBox $txtLog -Text "[OK] Ficheiro eliminado: $f"
                        } catch {
                            Add-LogLine -TextBox $txtLog -Text "[AVISO] Não foi possível eliminar o ficheiro: $f"
                        }
                    }
                }
            } else {
                Add-LogLine -TextBox $txtLog -Text "[ERRO] Falhou: $folderName (código $exitCode)"
            }
            $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    $btnRun.Enabled = $true
    $btnCancelMain.Enabled = $false
    $progressFolder.Value = 0
    $progressTotal.Value = 0

    if ($optionFlags.openLogFolderAfterFinish -and $script:CurrentLogFile) {
        Start-Process "explorer.exe" -ArgumentList "/select,""$($script:CurrentLogFile)"""
    }

    Add-LogLine -TextBox $txtLog -Text "Processo terminado"
})

# --- Inicialização ---
Update-FormatOptions
Load-State
Update-DynamicOptions
Refresh-UiState

# --- Mostrar form ---
[void]$form.ShowDialog()
