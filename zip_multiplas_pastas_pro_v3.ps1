Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$SevenZip = "C:\Program Files\7-Zip\7z.exe"

if (-not (Test-Path -LiteralPath $SevenZip)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Nao encontrei o 7z.exe em:`n$SevenZip",
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

function Get-UniqueArchivePath {
    param(
        [string]$Parent,
        [string]$BaseName,
        [string]$Extension
    )

    if ([string]::IsNullOrWhiteSpace($Parent)) {
        throw "O caminho da pasta-mae esta vazio."
    }

    if ([string]::IsNullOrWhiteSpace($BaseName)) {
        throw "O nome base da pasta esta vazio."
    }

    if ([string]::IsNullOrWhiteSpace($Extension)) {
        throw "A extensao do arquivo esta vazia."
    }

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
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "ZIP multiplas pastas - Pro v3"
$form.Size = New-Object System.Drawing.Size(1080, 860)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.AllowDrop = $true

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = "Arraste varias pastas do Explorador para a lista abaixo, ou use 'Adicionar pasta'."
$lblTop.Location = New-Object System.Drawing.Point(20, 15)
$lblTop.Size = New-Object System.Drawing.Size(1020, 20)
$form.Controls.Add($lblTop)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Adicionar pasta"
$btnAdd.Location = New-Object System.Drawing.Point(20, 45)
$btnAdd.Size = New-Object System.Drawing.Size(130, 32)
$form.Controls.Add($btnAdd)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remover selecionada"
$btnRemove.Location = New-Object System.Drawing.Point(160, 45)
$btnRemove.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($btnRemove)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Limpar lista"
$btnClear.Location = New-Object System.Drawing.Point(320, 45)
$btnClear.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($btnClear)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Iniciar"
$btnRun.Location = New-Object System.Drawing.Point(920, 45)
$btnRun.Size = New-Object System.Drawing.Size(120, 32)
$btnRun.Enabled = $false
$form.Controls.Add($btnRun)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20, 90)
$listBox.Size = New-Object System.Drawing.Size(1020, 200)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$form.Controls.Add($listBox)

# ===== Grupo: formato e compressao =====
$grpMain = New-Object System.Windows.Forms.GroupBox
$grpMain.Text = "Compressao"
$grpMain.Location = New-Object System.Drawing.Point(20, 310)
$grpMain.Size = New-Object System.Drawing.Size(500, 230)
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
$lblLevel.Text = "Nivel de compressao:"
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
$lblMethod.Text = "Metodo de compressao:"
$lblMethod.Location = New-Object System.Drawing.Point(15, 90)
$lblMethod.Size = New-Object System.Drawing.Size(140, 20)
$grpMain.Controls.Add($lblMethod)

$cmbMethod = New-Object System.Windows.Forms.ComboBox
$cmbMethod.Location = New-Object System.Drawing.Point(160, 87)
$cmbMethod.Size = New-Object System.Drawing.Size(120, 24)
$cmbMethod.DropDownStyle = "DropDownList"
$grpMain.Controls.Add($cmbMethod)

$lblDict = New-Object System.Windows.Forms.Label
$lblDict.Text = "Tamanho do dicionario:"
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
$lblSolid.Text = "Blocos solidos:"
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
$lblThreads.Location = New-Object System.Drawing.Point(300, 30)
$lblThreads.Size = New-Object System.Drawing.Size(120, 20)
$grpMain.Controls.Add($lblThreads)

$numThreads = New-Object System.Windows.Forms.NumericUpDown
$numThreads.Location = New-Object System.Drawing.Point(425, 27)
$numThreads.Size = New-Object System.Drawing.Size(55, 24)
$numThreads.Minimum = 1
$numThreads.Maximum = [Math]::Max(1, [Environment]::ProcessorCount)
$numThreads.Value = [Math]::Max(1, [Environment]::ProcessorCount)
$grpMain.Controls.Add($numThreads)

$lblVolume = New-Object System.Windows.Forms.Label
$lblVolume.Text = "Dividir por volumes:"
$lblVolume.Location = New-Object System.Drawing.Point(300, 60)
$lblVolume.Size = New-Object System.Drawing.Size(120, 20)
$grpMain.Controls.Add($lblVolume)

$txtVolume = New-Object System.Windows.Forms.TextBox
$txtVolume.Location = New-Object System.Drawing.Point(425, 57)
$txtVolume.Size = New-Object System.Drawing.Size(55, 24)
$grpMain.Controls.Add($txtVolume)

$lblVolumeHint = New-Object System.Windows.Forms.Label
$lblVolumeHint.Text = "ex: 700m, 2g"
$lblVolumeHint.Location = New-Object System.Drawing.Point(300, 85)
$lblVolumeHint.Size = New-Object System.Drawing.Size(180, 20)
$grpMain.Controls.Add($lblVolumeHint)

$chkSfx = New-Object System.Windows.Forms.CheckBox
$chkSfx.Text = "Criar arquivo SFX"
$chkSfx.Location = New-Object System.Drawing.Point(300, 120)
$chkSfx.Size = New-Object System.Drawing.Size(160, 22)
$grpMain.Controls.Add($chkSfx)

$chkShared = New-Object System.Windows.Forms.CheckBox
$chkShared.Text = "Comprimir ficheiros partilhados"
$chkShared.Location = New-Object System.Drawing.Point(300, 145)
$chkShared.Size = New-Object System.Drawing.Size(190, 22)
$grpMain.Controls.Add($chkShared)

$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Eliminar ficheiros apos compressao"
$chkDelete.Location = New-Object System.Drawing.Point(300, 170)
$chkDelete.Size = New-Object System.Drawing.Size(190, 22)
$chkDelete.Checked = $true
$grpMain.Controls.Add($chkDelete)

# ===== Grupo: encriptacao =====
$grpEnc = New-Object System.Windows.Forms.GroupBox
$grpEnc.Text = "Encriptacao"
$grpEnc.Location = New-Object System.Drawing.Point(540, 310)
$grpEnc.Size = New-Object System.Drawing.Size(500, 230)
$form.Controls.Add($grpEnc)

$lblPass1 = New-Object System.Windows.Forms.Label
$lblPass1.Text = "Introduza a palavra-passe:"
$lblPass1.Location = New-Object System.Drawing.Point(15, 30)
$lblPass1.Size = New-Object System.Drawing.Size(180, 20)
$grpEnc.Controls.Add($lblPass1)

$txtPass1 = New-Object System.Windows.Forms.TextBox
$txtPass1.Location = New-Object System.Drawing.Point(200, 27)
$txtPass1.Size = New-Object System.Drawing.Size(280, 24)
$txtPass1.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass1)

$lblPass2 = New-Object System.Windows.Forms.Label
$lblPass2.Text = "Reintroduza a palavra-passe:"
$lblPass2.Location = New-Object System.Drawing.Point(15, 60)
$lblPass2.Size = New-Object System.Drawing.Size(180, 20)
$grpEnc.Controls.Add($lblPass2)

$txtPass2 = New-Object System.Windows.Forms.TextBox
$txtPass2.Location = New-Object System.Drawing.Point(200, 57)
$txtPass2.Size = New-Object System.Drawing.Size(280, 24)
$txtPass2.UseSystemPasswordChar = $true
$grpEnc.Controls.Add($txtPass2)

$chkShowPass = New-Object System.Windows.Forms.CheckBox
$chkShowPass.Text = "Mostrar palavra-passe"
$chkShowPass.Location = New-Object System.Drawing.Point(15, 90)
$chkShowPass.Size = New-Object System.Drawing.Size(180, 22)
$grpEnc.Controls.Add($chkShowPass)

$lblEncMethod = New-Object System.Windows.Forms.Label
$lblEncMethod.Text = "Metodo de encriptacao:"
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

# ===== Grupo: extra =====
$grpExtra = New-Object System.Windows.Forms.GroupBox
$grpExtra.Text = "Opcoes extra"
$grpExtra.Location = New-Object System.Drawing.Point(20, 555)
$grpExtra.Size = New-Object System.Drawing.Size(1020, 100)
$form.Controls.Add($grpExtra)

$lblExtraParams = New-Object System.Windows.Forms.Label
$lblExtraParams.Text = "Parametros:"
$lblExtraParams.Location = New-Object System.Drawing.Point(15, 30)
$lblExtraParams.Size = New-Object System.Drawing.Size(100, 20)
$grpExtra.Controls.Add($lblExtraParams)

$txtExtraParams = New-Object System.Windows.Forms.TextBox
$txtExtraParams.Location = New-Object System.Drawing.Point(110, 27)
$txtExtraParams.Size = New-Object System.Drawing.Size(890, 24)
$grpExtra.Controls.Add($txtExtraParams)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Estado: parado"
$lblStatus.Location = New-Object System.Drawing.Point(20, 670)
$lblStatus.Size = New-Object System.Drawing.Size(1020, 20)
$form.Controls.Add($lblStatus)

$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20, 700)
$progressFolder.Size = New-Object System.Drawing.Size(1020, 24)
$progressFolder.Minimum = 0
$progressFolder.Maximum = 100
$form.Controls.Add($progressFolder)

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.Text = "Progresso total:"
$lblTotal.Location = New-Object System.Drawing.Point(20, 730)
$lblTotal.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblTotal)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20, 755)
$progressTotal.Size = New-Object System.Drawing.Size(1020, 24)
$progressTotal.Minimum = 0
$progressTotal.Maximum = 100
$form.Controls.Add($progressTotal)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 785)
$txtLog.Size = New-Object System.Drawing.Size(1020, 45)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$selectedFolders = New-Object 'System.Collections.Generic.List[string]'

function Refresh-UiState {
    $btnRun.Enabled = ($selectedFolders.Count -gt 0)
}

function Update-FormatOptions {
    $format = [string]$cmbFormat.SelectedItem

    $cmbMethod.Items.Clear()
    $cmbDict.Items.Clear()
    $cmbWord.Items.Clear()
    $cmbSolid.Items.Clear()
    $cmbEncMethod.Items.Clear()

    if ($format -eq "zip") {
        [void]$cmbMethod.Items.AddRange(@("Deflate","Deflate64","BZip2","LZMA","PPMd"))
        [void]$cmbDict.Items.AddRange(@("32KB","64KB","128KB","256KB","512KB","1MB","2MB","4MB","8MB","16MB","32MB"))
        [void]$cmbWord.Items.AddRange(@("32","64","128","192","256","273"))
        [void]$cmbSolid.Items.AddRange(@("Nao aplicavel"))
        [void]$cmbEncMethod.Items.AddRange(@("ZipCrypto","AES128","AES192","AES256"))

        $cmbMethod.SelectedItem = "Deflate"
        $cmbDict.SelectedItem = "32KB"
        $cmbWord.SelectedItem = "128"
        $cmbSolid.SelectedItem = "Nao aplicavel"
        $cmbEncMethod.SelectedItem = "ZipCrypto"

        $cmbSolid.Enabled = $false
        $chkEncryptHeaders.Enabled = $false
        $chkSfx.Enabled = $false
    }
    else {
        [void]$cmbMethod.Items.AddRange(@("LZMA2","LZMA","PPMd","BZip2"))
        [void]$cmbDict.Items.AddRange(@("64KB","1MB","2MB","4MB","8MB","16MB","32MB","64MB","128MB","256MB"))
        [void]$cmbWord.Items.AddRange(@("8","12","16","24","32","48","64","96","128","192","273"))
        [void]$cmbSolid.Items.AddRange(@("Off","On","1m","2m","4m","8m","16m","32m","64m","128m","256m","512m"))
        [void]$cmbEncMethod.Items.AddRange(@("AES256"))

        $cmbMethod.SelectedItem = "LZMA2"
        $cmbDict.SelectedItem = "16MB"
        $cmbWord.SelectedItem = "32"
        $cmbSolid.SelectedItem = "On"
        $cmbEncMethod.SelectedItem = "AES256"

        $cmbSolid.Enabled = $true
        $chkEncryptHeaders.Enabled = $true
        $chkSfx.Enabled = $true
    }
}

$cmbFormat.Add_SelectedIndexChanged({ Update-FormatOptions })

$chkShowPass.Add_CheckedChanged({
    $visible = $chkShowPass.Checked
    $txtPass1.UseSystemPasswordChar = -not $visible
    $txtPass2.UseSystemPasswordChar = -not $visible
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
    $selectedFolders.Clear()
    $listBox.Items.Clear()
    Add-LogLine -TextBox $txtLog -Text "Lista limpa."
    Refresh-UiState
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

    $pass1 = $txtPass1.Text
    $pass2 = $txtPass2.Text

    if (($pass1.Length -gt 0) -or ($pass2.Length -gt 0)) {
        if ($pass1 -ne $pass2) {
            [System.Windows.Forms.MessageBox]::Show("As palavras-passe nao coincidem.", "Erro")
            return
        }
    }

    $btnRun.Enabled = $false
    $btnAdd.Enabled = $false
    $btnRemove.Enabled = $false
    $btnClear.Enabled = $false

    $progressFolder.Value = 0
    $progressTotal.Value = 0

    $totalCount = $selectedFolders.Count
    $currentIndex = 0

    foreach ($folder in @($selectedFolders)) {
        $currentIndex++

        try {
            $folderPath = [string]$folder
            $folderObj = Get-Item -LiteralPath $folderPath -ErrorAction Stop
            $folderName = $folderObj.Name
            $parent = $folderObj.Parent.FullName

            $format = [string]$cmbFormat.SelectedItem
            $dest = Get-UniqueArchivePath -Parent $parent -BaseName $folderName -Extension $format

            $lblStatus.Text = "Estado: a comprimir $folderName ($currentIndex de $totalCount)"
            $progressFolder.Value = 0
            Add-LogLine -TextBox $txtLog -Text ""
            Add-LogLine -TextBox $txtLog -Text "A processar: $folderPath"
            Add-LogLine -TextBox $txtLog -Text "Destino: $dest"

            $args = New-Object System.Collections.Generic.List[string]
            $args.Add("a")
            $args.Add("-t$format")
            $args.Add("`"$dest`"")
            $args.Add("`"$folderName`"")
            $args.Add("-mx=$($cmbLevel.SelectedItem)")
            $args.Add("-mmt=$([int]$numThreads.Value)")
            $args.Add("-bsp1")
            $args.Add("-bso1")
            $args.Add("-bse1")
            $args.Add("-y")

            if ($cmbMethod.SelectedItem -and ($cmbMethod.SelectedItem -ne "")) {
                $method = [string]$cmbMethod.SelectedItem
                $methodMap = @{
                    "Deflate"   = "Deflate"
                    "Deflate64" = "Deflate64"
                    "BZip2"     = "BZip2"
                    "LZMA"      = "LZMA"
                    "LZMA2"     = "LZMA2"
                    "PPMd"      = "PPMd"
                }
                if ($methodMap.ContainsKey($method)) {
                    $args.Add("-mm=$($methodMap[$method])")
                }
            }

            if ($cmbDict.SelectedItem -and ($cmbDict.SelectedItem -ne "")) {
                $dict = ([string]$cmbDict.SelectedItem).ToLower()
                $dict = $dict -replace "kb","k"
                $dict = $dict -replace "mb","m"
                $args.Add("-md=$dict")
            }

            if ($cmbWord.SelectedItem -and ($cmbWord.SelectedItem -ne "")) {
                $args.Add("-mfb=$($cmbWord.SelectedItem)")
            }

            if (($format -eq "7z") -and $cmbSolid.SelectedItem -and ($cmbSolid.SelectedItem -ne "Nao aplicavel")) {
                $solid = [string]$cmbSolid.SelectedItem
                switch ($solid) {
                    "Off" { $args.Add("-ms=off") }
                    "On"  { $args.Add("-ms=on") }
                    default { $args.Add("-ms=$solid") }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($txtVolume.Text)) {
                $args.Add("-v$($txtVolume.Text.Trim())")
            }

            if ($chkSfx.Checked -and $format -eq "7z") {
                $args.Add("-sfx")
            }

            if ($chkShared.Checked) {
                $args.Add("-ssw")
            }

            if ($chkDelete.Checked) {
                $args.Add("-sdel")
            }

            if (-not [string]::IsNullOrEmpty($pass1)) {
                $args.Add("-p`"$pass1`"")

                if ($format -eq "zip") {
                    $enc = [string]$cmbEncMethod.SelectedItem
                    switch ($enc) {
                        "ZipCrypto" { $args.Add("-mem=ZipCrypto") }
                        "AES128"    { $args.Add("-mem=AES128") }
                        "AES192"    { $args.Add("-mem=AES192") }
                        "AES256"    { $args.Add("-mem=AES256") }
                    }
                }
                elseif ($format -eq "7z") {
                    if ($chkEncryptHeaders.Checked) {
                        $args.Add("-mhe=on")
                    }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($txtExtraParams.Text)) {
                $extra = $txtExtraParams.Text.Trim()
                if ($extra.Length -gt 0) {
                    foreach ($piece in ($extra -split '\s+')) {
                        if ($piece.Trim().Length -gt 0) {
                            $args.Add($piece.Trim())
                        }
                    }
                }
            }

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $SevenZip
            $psi.WorkingDirectory = $parent
            $psi.Arguments = ($args -join " ")
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $true

            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $psi
            $null = $p.Start()

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

            if ($exitCode -eq 0) {
                $progressFolder.Value = 100
                Add-LogLine -TextBox $txtLog -Text "[OK] Concluido: $folderName"

                if ($chkDelete.Checked) {
                    Remove-EmptyFolderIfPossible -FolderPath $folderPath
                }
            }
            else {
                Add-LogLine -TextBox $txtLog -Text "[ERRO] Falhou: $folderName (codigo $exitCode)"
            }
        }
        catch {
            Add-LogLine -TextBox $txtLog -Text "[ERRO] Excecao ao processar '$folder': $($_.Exception.Message)"
        }

        $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblStatus.Text = "Estado: terminado"
    Add-LogLine -TextBox $txtLog -Text ""
    Add-LogLine -TextBox $txtLog -Text "Processamento concluido."

    $btnRun.Enabled = $true
    $btnAdd.Enabled = $true
    $btnRemove.Enabled = $true
    $btnClear.Enabled = $true
    Refresh-UiState

    [System.Windows.Forms.MessageBox]::Show("Concluido.", "ZIP multiplas pastas")
})

Update-FormatOptions
Refresh-UiState
[void]$form.ShowDialog()