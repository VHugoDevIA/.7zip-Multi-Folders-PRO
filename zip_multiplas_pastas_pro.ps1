Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========= CONFIG =========
$SevenZip = "C:\Program Files\7-Zip\7z.exe"
$DefaultCompressionLevel = "9"
$DefaultArchiveType = "zip"
# =========================

if (-not (Test-Path $SevenZip)) {
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

    $dest = Join-Path $Parent ($BaseName + ".zip")
    if (-not (Test-Path $dest)) { return $dest }

    $i = 2
    do {
        $candidate = Join-Path $Parent ("{0}_{1}.zip" -f $BaseName, $i)
        $i++
    } while (Test-Path $candidate)

    return $candidate
}

function Remove-EmptyFolderRecursive {
    param([string]$FolderPath)

    if (-not (Test-Path $FolderPath)) { return }

    try {
        $items = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop
        if ($items.Count -eq 0) {
            Remove-Item -LiteralPath $FolderPath -Force -ErrorAction Stop
        }
    } catch {
    }
}

function Select-FoldersDialog {
    $selected = New-Object System.Collections.Generic.List[string]

    while ($true) {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Selecione uma pasta para adicionar. Carregue Cancelar quando terminar."
        $dlg.ShowNewFolderButton = $false

        $result = $dlg.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { break }

        if (-not $selected.Contains($dlg.SelectedPath)) {
            $selected.Add($dlg.SelectedPath)
        }
    }

    return $selected
}

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "ZIP multiplas pastas - Pro"
$form.Size = New-Object System.Drawing.Size(760, 620)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$btnSelect = New-Object System.Windows.Forms.Button
$btnSelect.Text = "Selecionar pastas"
$btnSelect.Location = New-Object System.Drawing.Point(20,20)
$btnSelect.Size = New-Object System.Drawing.Size(150,32)
$form.Controls.Add($btnSelect)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Iniciar"
$btnRun.Location = New-Object System.Drawing.Point(590,20)
$btnRun.Size = New-Object System.Drawing.Size(130,32)
$btnRun.Enabled = $false
$form.Controls.Add($btnRun)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20,70)
$listBox.Size = New-Object System.Drawing.Size(700,180)
$form.Controls.Add($listBox)

$chkDelete = New-Object System.Windows.Forms.CheckBox
$chkDelete.Text = "Eliminar ficheiros apos compressao"
$chkDelete.Location = New-Object System.Drawing.Point(20,270)
$chkDelete.Size = New-Object System.Drawing.Size(280,24)
$chkDelete.Checked = $true
$form.Controls.Add($chkDelete)

$chkMinimize7z = New-Object System.Windows.Forms.CheckBox
$chkMinimize7z.Text = "Executar discretamente"
$chkMinimize7z.Location = New-Object System.Drawing.Point(320,270)
$chkMinimize7z.Size = New-Object System.Drawing.Size(180,24)
$chkMinimize7z.Checked = $true
$form.Controls.Add($chkMinimize7z)

$lblLevel = New-Object System.Windows.Forms.Label
$lblLevel.Text = "Nivel de compressao:"
$lblLevel.Location = New-Object System.Drawing.Point(20,305)
$lblLevel.Size = New-Object System.Drawing.Size(150,24)
$form.Controls.Add($lblLevel)

$cmbLevel = New-Object System.Windows.Forms.ComboBox
$cmbLevel.Location = New-Object System.Drawing.Point(170,302)
$cmbLevel.Size = New-Object System.Drawing.Size(80,24)
$cmbLevel.DropDownStyle = "DropDownList"
[void]$cmbLevel.Items.AddRange(@("0","1","3","5","7","9"))
$cmbLevel.SelectedItem = $DefaultCompressionLevel
$form.Controls.Add($cmbLevel)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Estado: parado"
$lblStatus.Location = New-Object System.Drawing.Point(20,345)
$lblStatus.Size = New-Object System.Drawing.Size(700,24)
$form.Controls.Add($lblStatus)

$progressFolder = New-Object System.Windows.Forms.ProgressBar
$progressFolder.Location = New-Object System.Drawing.Point(20,375)
$progressFolder.Size = New-Object System.Drawing.Size(700,24)
$progressFolder.Minimum = 0
$progressFolder.Maximum = 100
$form.Controls.Add($progressFolder)

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.Text = "Progresso total:"
$lblTotal.Location = New-Object System.Drawing.Point(20,410)
$lblTotal.Size = New-Object System.Drawing.Size(200,24)
$form.Controls.Add($lblTotal)

$progressTotal = New-Object System.Windows.Forms.ProgressBar
$progressTotal.Location = New-Object System.Drawing.Point(20,440)
$progressTotal.Size = New-Object System.Drawing.Size(700,24)
$progressTotal.Minimum = 0
$progressTotal.Maximum = 100
$form.Controls.Add($progressTotal)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20,475)
$txtLog.Size = New-Object System.Drawing.Size(700,95)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$selectedFolders = New-Object System.Collections.Generic.List[string]

function Add-Log {
    param([string]$Text)
    $txtLog.AppendText($Text + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
}

$btnSelect.Add_Click({
    $folders = Select-FoldersDialog
    if ($folders.Count -gt 0) {
        foreach ($f in $folders) {
            if (-not $selectedFolders.Contains($f)) {
                $selectedFolders.Add($f)
                [void]$listBox.Items.Add($f)
            }
        }
        $btnRun.Enabled = ($selectedFolders.Count -gt 0)
        Add-Log "Pastas carregadas: $($selectedFolders.Count)"
    }
})

$btnRun.Add_Click({
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Selecione pelo menos uma pasta.")
        return
    }

    $btnRun.Enabled = $false
    $btnSelect.Enabled = $false
    $progressFolder.Value = 0
    $progressTotal.Value = 0

    $totalCount = $selectedFolders.Count
    $currentIndex = 0

    foreach ($folder in $selectedFolders) {
        $currentIndex++
        $folderName = [System.IO.Path]::GetFileName($folder.TrimEnd('\'))
        $parent = Split-Path -LiteralPath $folder -Parent
        $dest = Get-UniqueZipPath -Parent $parent -BaseName $folderName

        $lblStatus.Text = "Estado: a comprimir $folderName ($currentIndex de $totalCount)"
        $progressFolder.Value = 0
        Add-Log "A processar: $folder"
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
        $p.Start() | Out-Null

        while (-not $p.HasExited) {
            while (($line = $p.StandardOutput.ReadLine()) -ne $null) {
                if ($line -match '(\d+)%') {
                    $pct = [int]$matches[1]
                    if ($pct -ge 0 -and $pct -le 100) {
                        $progressFolder.Value = $pct
                    }
                }
                if ($line.Trim() -ne "") {
                    Add-Log $line
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
                if ($errLine.Trim() -ne "") {
                    Add-Log "[ERRO] $errLine"
                }
                [System.Windows.Forms.Application]::DoEvents()
            }

            Start-Sleep -Milliseconds 150
            [System.Windows.Forms.Application]::DoEvents()
        }

        while (($line = $p.StandardOutput.ReadLine()) -ne $null) {
            if ($line -match '(\d+)%') {
                $pct = [int]$matches[1]
                if ($pct -ge 0 -and $pct -le 100) {
                    $progressFolder.Value = $pct
                }
            }
            if ($line.Trim() -ne "") {
                Add-Log $line
            }
        }

        while (($errLine = $p.StandardError.ReadLine()) -ne $null) {
            if ($errLine.Trim() -ne "") {
                Add-Log "[ERRO] $errLine"
            }
        }

        $p.WaitForExit()
        $exitCode = $p.ExitCode

        if ($exitCode -eq 0) {
            $progressFolder.Value = 100
            Add-Log "[OK] Concluido: $folderName"

            if ($chkDelete.Checked) {
                Remove-EmptyFolderRecursive -FolderPath $folder
            }
        }
        else {
            Add-Log "[ERRO] Falhou: $folderName (codigo $exitCode)"
        }

        $progressTotal.Value = [math]::Min([math]::Round(($currentIndex / $totalCount) * 100), 100)
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblStatus.Text = "Estado: terminado"
    Add-Log "Processamento concluido."
    [System.Windows.Forms.MessageBox]::Show("Concluido.", "ZIP multiplas pastas")
    $btnRun.Enabled = $true
    $btnSelect.Enabled = $true
})

[void]$form.ShowDialog()