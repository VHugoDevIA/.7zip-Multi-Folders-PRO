Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$DefaultArchiveType = "zip"
$DefaultCompressionLevel = "9"

if (-not (Test-Path -LiteralPath $SevenZip)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Nao encontrei o 7z.exe em:`n$SevenZip",
        "Erro",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

function Get-UniqueZipPath {
    param(
        [string]$Parent,
        [string]$BaseName
    )

    if ([string]::IsNullOrWhiteSpace($Parent)) {
        throw "O caminho da pasta-mae esta vazio."
    }

    if ([string]::IsNullOrWhiteSpace($BaseName)) {
        throw "O nome base da pasta esta vazio."
    }

    $dest = Join-Path -Path $Parent -ChildPath ($BaseName + ".zip")
    if (-not (Test-Path -LiteralPath $dest)) {
        return $dest
    }

    $i = 2
    do {
        $dest = Join-Path -Path $Parent -ChildPath ("{0}_{1}.zip" -f $BaseName, $i)
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

$form = New-Object System.Windows.Forms.Form
$form.Text = "ZIP multiplas pastas - Pro v2"
$form.Size = New-Object System.Drawing.Size(900, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.AllowDrop = $true

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = "Arraste varias pastas do Explorador para a lista abaixo, ou use 'Adicionar pasta'."
$lblTop.Location = New-Object System.Drawing.Point(20, 15)
$lblTop.Size = New-Object System.Drawing.Size(820, 20)
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
$btnRun.Location = New-Object System.Drawing.Point(740, 45)
$btnRun.Size = New-Object System.Drawing.Size(120, 32)
$btnRun.Enabled = $false
$form.Controls.Add($btnRun)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20, 90)
$listBox.Size = New-Object System.Drawing.Size(840, 230)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$form.Controls.Add($listBox)

$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Eliminar ficheiros apos compressao"
$chkDelete.Location = New-Object System.Drawing.Point(20, 340)
$chkDelete.Size = New-Object System.Drawing.Size(280, 24)
$chkDelete.Checked = $true
$form.Controls.Add($chkDelete)

$lblLevel = New-Object System.Windows.Forms.Label
$lblLevel.Text = "Nivel de compressao:"
$lblLevel.Location = New-Object System.Drawing.Point(20, 375)
$lblLevel.Size = New-Object System.Drawing.Size(140, 24)
$form.Controls.Add($lblLevel)

$cmbLevel = New-Object System.Windows.Forms.ComboBox
$cmbLevel.Location = New-Object System.Drawing.Point(165, 372)
$cmbLevel.Size = New-Object System.Drawing.Size(70, 24)
$cmbLevel.DropDownStyle = "DropDownList"
[void]$cmbLevel.Items.AddRange(@("0", "1", "3", "5", "7", "9"))
$cmbLevel.SelectedItem = $DefaultCompressionLevel
$form.Controls.Add($cmbLevel)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Estado: parado"
$lblStatus.Location = New-Object System.Drawing.Point(20, 410)
$lblStatus.Size = New-Object System.Drawing.Size(840, 20)
$form.Controls.Add($lblStatus)

$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20, 440)
$progressFolder.Size = New-Object System.Drawing.Size(840, 24)
$progressFolder.Minimum = 0
$progressFolder.Maximum = 100
$form.Controls.Add($progressFolder)

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.Text = "Progresso total:"
$lblTotal.Location = New-Object System.Drawing.Point(20, 475)
$lblTotal.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblTotal)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20, 500)
$progressTotal.Size = New-Object System.Drawing.Size(840, 24)
$progressTotal.Minimum = 0
$progressTotal.Maximum = 100
$form.Controls.Add($progressTotal)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 545)
$txtLog.Size = New-Object System.Drawing.Size(840, 130)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$selectedFolders = New-Object 'System.Collections.Generic.List[string]'

function Refresh-UiState {
    $btnRun.Enabled = ($selectedFolders.Count -gt 0)
}

function Add-Log {
    param([string]$Text)

    $txtLog.AppendText($Text + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
}

$btnAdd.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Selecione uma pasta para adicionar"
    $dlg.ShowNewFolderButton = $false

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FolderToList -Path $dlg.SelectedPath -ListBox $listBox -Store $selectedFolders
        Add-Log "Pasta adicionada: $($dlg.SelectedPath)"
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
        Add-Log "Removida: $item"
    }

    Refresh-UiState
})

$btnClear.Add_Click({
    $selectedFolders.Clear()
    $listBox.Items.Clear()
    Add-Log "Lista limpa."
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
            Add-Log "Pasta adicionada por arrastar: $item"
        }
    }

    Refresh-UiState
})

$btnRun.Add_Click({
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Adiciona pelo menos uma pasta.")
        return
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
            $dest = Get-UniqueZipPath -Parent $parent -BaseName $folderName

            $lblStatus.Text = "Estado: a comprimir $folderName ($currentIndex de $totalCount)"
            $progressFolder.Value = 0
            Add-Log ""
            Add-Log "A processar: $folderPath"
            Add-Log "Destino: $dest"

            $args = @(
                "a"
                "-t$DefaultArchiveType"
                "`"$dest`""
                "`"$folderName`""
                "-mx=$($cmbLevel.SelectedItem)"
                "-bsp1"
                "-bso1"
                "-bse1"
                "-y"
            )

            if ($chkDelete.Checked) {
                $args += "-sdel"
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
                        Add-Log $line
                    }

                    [System.Windows.Forms.Application]::DoEvents()
                }

                while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                    if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                        Add-Log "[ERRO] $errLine"
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
                    Add-Log $line
                }
            }

            while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                    Add-Log "[ERRO] $errLine"
                }
            }

            $p.WaitForExit()
            $exitCode = $p.ExitCode

            if ($exitCode -eq 0) {
                $progressFolder.Value = 100
                Add-Log "[OK] Concluido: $folderName"

                if ($chkDelete.Checked) {
                    Remove-EmptyFolderIfPossible -FolderPath $folderPath
                }
            }
            else {
                Add-Log "[ERRO] Falhou: $folderName (codigo $exitCode)"
            }
        }
        catch {
            Add-Log "[ERRO] Excecao ao processar '$folder': $($_.Exception.Message)"
        }

        $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblStatus.Text = "Estado: terminado"
    Add-Log ""
    Add-Log "Processamento concluido."

    $btnRun.Enabled = $true
    $btnAdd.Enabled = $true
    $btnRemove.Enabled = $true
    $btnClear.Enabled = $true
    Refresh-UiState

    [System.Windows.Forms.MessageBox]::Show("Concluido.", "ZIP multiplas pastas")
})

Refresh-UiState
[void]$form.ShowDialog()