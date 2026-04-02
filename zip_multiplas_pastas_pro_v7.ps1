
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

function Save-State {
    $state = @{
        windowSize = @{ Width = $form.Width; Height = $form.Height }
        windowLocation = @{ X = $form.Location.X; Y = $form.Location.Y }
        format = $cmbFormat.SelectedItem
        compression = $cmbCompression.SelectedItem
        method = $cmbMethod.SelectedItem
        threads = $txtThreads.Text
        password = $txtPassword.Text
        encryptHeaders = $chkEncryptHeaders.Checked
        deleteAfterCompress = $chkDeleteFiles.Checked
        sharedFiles = $chkShared.Checked
        sfx = $chkSFX.Checked
        relativePath = $chkRelative.Checked
        volumeSize = $txtVolume.Text
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
            if ($state.threads) { $txtThreads.Text = $state.threads }
            if ($state.password) { $txtPassword.Text = $state.password }
            if ($state.encryptHeaders) { $chkEncryptHeaders.Checked = $state.encryptHeaders }
            if ($state.deleteAfterCompress) { $chkDeleteFiles.Checked = $state.deleteAfterCompress }
            if ($state.sharedFiles) { $chkShared.Checked = $state.sharedFiles }
            if ($state.sfx) { $chkSFX.Checked = $state.sfx }
            if ($state.relativePath) { $chkRelative.Checked = $state.relativePath }
            if ($state.volumeSize) { $txtVolume.Text = $state.volumeSize }
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
$form.Size = New-Object System.Drawing.Size(1080, 930)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.AllowDrop = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.Add_FormClosing({ Save-State })

# --- Labels ---
$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = "Formato:"
$lblFormat.Location = New-Object System.Drawing.Point(20, 20)
$lblFormat.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($lblFormat)

$lblCompression = New-Object System.Windows.Forms.Label
$lblCompression.Text = "Compressão:"
$lblCompression.Location = New-Object System.Drawing.Point(140, 20)
$lblCompression.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($lblCompression)

$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Método:"
$lblMethod.Location = New-Object System.Drawing.Point(260, 20)
$lblMethod.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($lblMethod)

$lblThreads = New-Object System.Windows.Forms.Label
$lblThreads.Text = "Threads:"
$lblThreads.Location = New-Object System.Drawing.Point(380, 20)
$lblThreads.Size = New-Object System.Drawing.Size(55, 20)
$form.Controls.Add($lblThreads)

$lblPassword = New-Object System.Windows.Forms.Label
$lblPassword.Text = "Password:"
$lblPassword.Location = New-Object System.Drawing.Point(500, 20)
$lblPassword.Size = New-Object System.Drawing.Size(65, 20)
$form.Controls.Add($lblPassword)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Pasta de saída:"
$lblOutput.Location = New-Object System.Drawing.Point(20, 50)
$lblOutput.Size = New-Object System.Drawing.Size(90, 20)
$form.Controls.Add($lblOutput)

$lblVolume = New-Object System.Windows.Forms.Label
$lblVolume.Text = "Volume (ex: 100m):"
$lblVolume.Location = New-Object System.Drawing.Point(500, 50)
$lblVolume.Size = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($lblVolume)

$lblExtra = New-Object System.Windows.Forms.Label
$lblExtra.Text = "Parâmetros extra:"
$lblExtra.Location = New-Object System.Drawing.Point(20, 80)
$lblExtra.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblExtra)

# --- ComboBoxes ---
$cmbFormat = New-Object System.Windows.Forms.ComboBox
$cmbFormat.Location = New-Object System.Drawing.Point(20, 35)
$cmbFormat.Size = New-Object System.Drawing.Size(100, 20)
$cmbFormat.Items.AddRange(@("zip", "7z"))
$cmbFormat.SelectedIndex = 0
$form.Controls.Add($cmbFormat)

$cmbCompression = New-Object System.Windows.Forms.ComboBox
$cmbCompression.Location = New-Object System.Drawing.Point(140, 35)
$cmbCompression.Size = New-Object System.Drawing.Size(100, 20)
$cmbCompression.Items.AddRange(@("0", "1", "3", "5", "7", "9"))
$cmbCompression.SelectedIndex = 3
$form.Controls.Add($cmbCompression)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Location = New-Object System.Drawing.Point(260, 35)
$cmbMethod.Size = New-Object System.Drawing.Size(100, 20)
$cmbMethod.Items.AddRange(@("LZMA2", "LZMA", "PPMd", "BZip2", "Deflate", "Deflate64"))
$cmbMethod.SelectedIndex = 0
$form.Controls.Add($cmbMethod)

# --- TextBoxes ---
$txtThreads = New-Object System.Windows.Forms.TextBox
$txtThreads.Location = New-Object System.Drawing.Point(380, 35)
$txtThreads.Size = New-Object System.Drawing.Size(50, 20)
$txtThreads.Text = "0"
$form.Controls.Add($txtThreads)

$txtPassword = New-Object System.Windows.Forms.TextBox
$txtPassword.Location = New-Object System.Drawing.Point(500, 35)
$txtPassword.Size = New-Object System.Drawing.Size(150, 20)
$txtPassword.UseSystemPasswordChar = $true
$form.Controls.Add($txtPassword)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20, 65)
$txtOutput.Size = New-Object System.Drawing.Size(450, 20)
$form.Controls.Add($txtOutput)

$txtVolume = New-Object System.Windows.Forms.TextBox
$txtVolume.Location = New-Object System.Drawing.Point(500, 65)
$txtVolume.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($txtVolume)

$txtExtra = New-Object System.Windows.Forms.TextBox
$txtExtra.Location = New-Object System.Drawing.Point(20, 95)
$txtExtra.Size = New-Object System.Drawing.Size(630, 20)
$form.Controls.Add($txtExtra)

# --- CheckBoxes ---
$chkEncryptHeaders = New-Object System.Windows.Forms.CheckBox
$chkEncryptHeaders.Text = "Encriptar nomes ficheiros (7z)"
$chkEncryptHeaders.Location = New-Object System.Drawing.Point(680, 35)
$chkEncryptHeaders.Size = New-Object System.Drawing.Size(180, 20)
$form.Controls.Add($chkEncryptHeaders)

$chkDeleteFiles = New-Object System.Windows.Forms.CheckBox
$chkDeleteFiles.Text = "Eliminar ficheiros após compressão"
$chkDeleteFiles.Location = New-Object System.Drawing.Point(680, 55)
$chkDeleteFiles.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($chkDeleteFiles)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = "Comprimir ficheiros partilhados"
$chkShared.Location = New-Object System.Drawing.Point(680, 75)
$chkShared.Size = New-Object System.Drawing.Size(180, 20)
$chkShared.Add_CheckedChanged({ Refresh-UiState })
$form.Controls.Add($chkShared)

$chkSFX = New-Object System.Windows.Forms.CheckBox
$chkSFX.Text = "SFX (7z)"
$chkSFX.Location = New-Object System.Drawing.Point(680, 95)
$chkSFX.Size = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($chkSFX)

$chkRelative = New-Object System.Windows.Forms.CheckBox
$chkRelative.Text = "Caminho relativo"
$chkRelative.Location = New-Object System.Drawing.Point(760, 95)
$chkRelative.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($chkRelative)

# --- Painel para seleção ---
$panelSelect = New-Object System.Windows.Forms.Panel
$panelSelect.Location = New-Object System.Drawing.Point(20, 113)
$panelSelect.Size = New-Object System.Drawing.Size(710, 300)
$form.Controls.Add($panelSelect)

# ListBox único para ficheiros e pastas
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(0, 0)
$listBox.Size = New-Object System.Drawing.Size(500, 300)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$panelSelect.Controls.Add($listBox)

# --- Painel lateral para botões ---
$panelBtns = New-Object System.Windows.Forms.Panel
$panelBtns.Location = New-Object System.Drawing.Point(740, 113)
$panelBtns.Size = New-Object System.Drawing.Size(180, 380)
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

# Posicionar botões verticalmente
$btnAdd.Location = New-Object System.Drawing.Point(30, 10)
$btnAddFile.Location = New-Object System.Drawing.Point(30, 50)
$btnRemove.Location = New-Object System.Drawing.Point(30, 90)
$btnRemoveFile.Location = New-Object System.Drawing.Point(30, 130)
$btnClear.Location = New-Object System.Drawing.Point(30, 170)
$btnClearLog.Location = New-Object System.Drawing.Point(30, 210)
$btnOptions.Location = New-Object System.Drawing.Point(30, 250)
$btnCancelMain.Location = New-Object System.Drawing.Point(30, 290)
$btnRun.Location = New-Object System.Drawing.Point(30, 330)

# --- Progress bars ---
$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20, 430)
$progressFolder.Size = New-Object System.Drawing.Size(700, 20)
$form.Controls.Add($progressFolder)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20, 460)
$progressTotal.Size = New-Object System.Drawing.Size(700, 20)
$form.Controls.Add($progressTotal)

# --- TextBox para log ---
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 490)
$txtLog.Size = New-Object System.Drawing.Size(900, 350)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
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
    if ($selectedFolders.Count -eq 0 -and $selectedFiles.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Adiciona pelo menos uma pasta ou ficheiro.")
        return
    }

    $script:CancelRequested = $false
    $btnRun.Enabled = $false
    $btnCancelMain.Enabled = $true

    # Criar pasta de log
    $logDir = Join-Path -Path ([Environment]::GetFolderPath("MyDocuments")) -ChildPath "ZIP_Múltiplas_Pastas-PRO LOG"
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:CurrentLogFile = Join-Path -Path $logDir -ChildPath ("Log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")

    # Preparar argumentos base
    $format = $cmbFormat.SelectedItem
    $compression = $cmbCompression.SelectedItem
    $method = $cmbMethod.SelectedItem
    $threads = [int]$txtThreads.Text
    $password = $txtPassword.Text
    $encryptHeaders = $chkEncryptHeaders.Checked
    $sfx = $chkSFX.Checked
    $relative = $chkRelative.Checked
    $volume = $txtVolume.Text.Trim()
    $extra = $txtExtra.Text.Trim()
    $outputFolder = $txtOutput.Text.Trim()

    if (-not $outputFolder) { $outputFolder = Split-Path -Parent $selectedFolders[0] }

    $argBase = "a -t$format -$compression -$method"
    if ($threads -gt 0) { $argBase += " -mmt$threads" }
    if ($password) {
        $argBase += " -p$password"
        if ($format -eq "7z" -and $encryptHeaders) { $argBase += " -mhe=on" }
    }
    if ($sfx -and $format -eq "7z") { $argBase += " -sfx" }
    if ($relative) { $argBase += " -spf" }
    if ($volume) { $argBase += " -v$volume" }
    if ($extra) { $argBase += " $extra" }

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

            $argString = "$argBase ""$archivePath"" ""$folder"""
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

            $fileList = $files -join '" "'
            $argString = "$argBase ""$archivePath"" ""$fileList"""
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
Load-State
Refresh-UiState

# --- Mostrar form ---
[void]$form.ShowDialog()
