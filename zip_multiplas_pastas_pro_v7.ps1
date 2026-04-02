
# 1. Criar o form e painéis principais primeiro
$form = New-Object System.Windows.Forms.Form
$form.Text = "ZIP Multiplas Pasta/Ficheiros - PRO v7.0"
$form.Size = New-Object System.Drawing.Size(1080, 930)
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.AllowDrop = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Painel para seleção de ficheiros e botões lado a lado
$panelSelect = New-Object System.Windows.Forms.Panel
$panelSelect.Location = New-Object System.Drawing.Point(20, 113)
$panelSelect.Size = New-Object System.Drawing.Size(710, 300)
$form.Controls.Add($panelSelect)

# Painel lateral para botões
$panelBtns = New-Object System.Windows.Forms.Panel
$panelBtns.Location = New-Object System.Drawing.Point(740, 113)
$panelBtns.Size = New-Object System.Drawing.Size(180, 380)
$form.Controls.Add($panelBtns)

# 2. Só depois criar botões, listas, etc.
$btnAddFile = New-Object System.Windows.Forms.Button
$btnAddFile.Text = "Adicionar ficheiro(s)"
$btnAddFile.Size = New-Object System.Drawing.Size(130, 32)
$btnAddFile.FlatStyle = "Flat"
$btnAddFile.BackColor = [System.Drawing.Color]::FromArgb(186, 225, 255)
$btnAddFile.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

$btnRemoveFile = New-Object System.Windows.Forms.Button
$btnRemoveFile.Text = "Remover ficheiro"
$btnRemoveFile.Size = New-Object System.Drawing.Size(130, 32)
$btnRemoveFile.FlatStyle = "Flat"
$btnRemoveFile.BackColor = [System.Drawing.Color]::FromArgb(255, 179, 186)
$btnRemoveFile.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)

# ListBox único para ficheiros e pastas
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(0, 0)
$listBox.Size = New-Object System.Drawing.Size(500, 300)
$listBox.SelectionMode = "MultiExtended"
$listBox.HorizontalScrollbar = $true
$listBox.AllowDrop = $true
$panelSelect.Controls.Add($listBox)

# 3. Só agora adicionar botões ao painel
$panelBtns.Controls.Clear()
$btnAdd.Location = New-Object System.Drawing.Point(30, 10)
$btnAddFile.Location = New-Object System.Drawing.Point(30, 50)
$btnRemove.Location = New-Object System.Drawing.Point(30, 90)
$btnRemoveFile.Location = New-Object System.Drawing.Point(30, 130)
$btnClear.Location = New-Object System.Drawing.Point(30, 170)
$btnClearLog.Location = New-Object System.Drawing.Point(30, 210)
$btnOptions.Location = New-Object System.Drawing.Point(30, 250)
$btnCancelMain.Location = New-Object System.Drawing.Point(30, 290)
$btnRun.Location = New-Object System.Drawing.Point(30, 330)
$panelBtns.Controls.AddRange(@($btnAdd, $btnAddFile, $btnRemove, $btnRemoveFile, $btnClear, $btnClearLog, $btnOptions, $btnCancelMain, $btnRun))
			if (-not $selectedFiles.Contains($item)) {
				$selectedFiles.Add($item)
				[void]$listBoxFiles.Items.Add($item)
				Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado por arrastar: $item"
			}
		}

# Alinhar todos os botões verticalmente
$panelBtns.Controls.Clear()
$btnAdd.Location = New-Object System.Drawing.Point(30, 10)
$btnAddFile.Location = New-Object System.Drawing.Point(30, 50)
$btnRemove.Location = New-Object System.Drawing.Point(30, 90)
$btnRemoveFile.Location = New-Object System.Drawing.Point(30, 130)
$btnClear.Location = New-Object System.Drawing.Point(30, 170)
$btnClearLog.Location = New-Object System.Drawing.Point(30, 210)
$btnOptions.Location = New-Object System.Drawing.Point(30, 250)
$btnCancelMain.Location = New-Object System.Drawing.Point(30, 290)
$btnRun.Location = New-Object System.Drawing.Point(30, 330)
$panelBtns.Size = New-Object System.Drawing.Size(180, 380)
$panelBtns.Controls.AddRange(@($btnAdd, $btnAddFile, $btnRemove, $btnRemoveFile, $btnClear, $btnClearLog, $btnOptions, $btnCancelMain, $btnRun))
	foreach ($item in $items) {
		if (Test-Path -LiteralPath $item -PathType Leaf) {
			if (-not $selectedFiles.Contains($item)) {
				$selectedFiles.Add($item)
				[void]$listBoxFiles.Items.Add($item)
				Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado por arrastar: $item"
			}
		}
	}
	Refresh-UiState
})
$listBoxFiles.Add_DragEnter({
	if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
		$_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
	} else {
		$_.Effect = [System.Windows.Forms.DragDropEffects]::None
	}
})

# --- Botões de adicionar/remover ficheiros ---
$btnAddFile.Add_Click({
	$dlg = New-Object System.Windows.Forms.OpenFileDialog
	$dlg.Title = "Selecione um ou mais ficheiros para adicionar"
	$dlg.Multiselect = $true
	if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
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
	$chkDeleteFiles.Enabled = -not $chkShared.Checked
	if ($chkShared.Checked) { $chkDeleteFiles.Checked = $false }
})

# --- Processamento de ficheiros ---
$btnRun.Add_Click({
	if ($selectedFolders.Count -eq 0 -and $selectedFiles.Count -eq 0) {
		[System.Windows.Forms.MessageBox]::Show("Adiciona pelo menos uma pasta ou ficheiro.")
		return
	}

	# ...código de processamento de pastas igual ao anterior...

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
		Refresh-UiState
	}
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
})
# Versão 7 do script baseada em v6_7
# - Permite inserir múltiplos ficheiros para compactar
# - Cria um ZIP por cada pasta onde estão os ficheiros
# - Acrescenta opção de eliminar ficheiros após compressão (como já existe para pastas)
# - Se "comprimir ficheiros partilhados" estiver ativo, não elimina ficheiros no fim

# O código será copiado e adaptado da versão v6_7

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

# ...restante código copiado da v6_7, com as seguintes alterações principais...
# 1. Adicionar botões para adicionar/remover ficheiros
# 2. Permitir drag&drop de ficheiros e pastas
# 3. No processamento, se houver ficheiros, agrupar por pasta e criar um ZIP por pasta
# 4. Adicionar opção de eliminar ficheiros após compressão (desativada se 'comprimir ficheiros partilhados')
# 5. Atualizar Refresh-UiState para ativar o botão Iniciar se houver ficheiros ou pastas

#
# O código completo será inserido nas próximas etapas...
