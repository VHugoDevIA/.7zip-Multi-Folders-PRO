# --- Botões para ficheiros ---
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

# Adicionar botões ao painel de botões
$panelBtns.Controls.Add($btnAddFile)
$panelBtns.Controls.Add($btnRemoveFile)
$btnAddFile.Location = New-Object System.Drawing.Point(30, 290)
$btnRemoveFile.Location = New-Object System.Drawing.Point(30, 330)

# --- ListBox para ficheiros ---
$listBoxFiles = New-Object System.Windows.Forms.ListBox
$listBoxFiles.Location = New-Object System.Drawing.Point(510, 0)
$listBoxFiles.Size = New-Object System.Drawing.Size(200, 300)
$listBoxFiles.SelectionMode = "MultiExtended"
$listBoxFiles.HorizontalScrollbar = $true
$listBoxFiles.AllowDrop = $true
$panelSelect.Controls.Add($listBoxFiles)

# --- Drag&Drop para ficheiros e pastas ---
$listBox.Add_DragDrop({
	$items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
	foreach ($item in $items) {
		if (Test-Path -LiteralPath $item -PathType Container) {
			Add-FolderToList -Path $item -ListBox $listBox -Store $selectedFolders
			Add-LogLine -TextBox $txtLog -Text "Pasta adicionada por arrastar: $item"
		} elseif (Test-Path -LiteralPath $item -PathType Leaf) {
			if (-not $selectedFiles.Contains($item)) {
				$selectedFiles.Add($item)
				[void]$listBoxFiles.Items.Add($item)
				Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado por arrastar: $item"
			}
		}
	}
	Refresh-UiState
})
$listBoxFiles.Add_DragDrop({
	$items = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
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
		foreach ($file in $dlg.FileNames) {
			if (-not $selectedFiles.Contains($file)) {
				$selectedFiles.Add($file)
				[void]$listBoxFiles.Items.Add($file)
				Add-LogLine -TextBox $txtLog -Text "Ficheiro adicionado: $file"
			}
		}
		Refresh-UiState
	}
})
$btnRemoveFile.Add_Click({
	if ($listBoxFiles.SelectedItems.Count -eq 0) { return }
	$toRemove = @()
	foreach ($item in $listBoxFiles.SelectedItems) { $toRemove += [string]$item }
	foreach ($item in $toRemove) {
		[void]$selectedFiles.Remove($item)
		[void]$listBoxFiles.Items.Remove($item)
		Add-LogLine -TextBox $txtLog -Text "Ficheiro removido: $item"
	}
	Refresh-UiState
})

# --- Atualizar Refresh-UiState ---
function Refresh-UiState {
	$btnRun.Enabled = ($selectedFolders.Count -gt 0 -or $selectedFiles.Count -gt 0)
}

# --- Opção de eliminar ficheiros após compressão ---
$chkDeleteFiles = New-Object System.Windows.Forms.CheckBox
$chkDeleteFiles.Text = "Eliminar ficheiros após compressão"
$chkDeleteFiles.Location = New-Object System.Drawing.Point(300, 230)
$chkDeleteFiles.Size = New-Object System.Drawing.Size(220, 22)
$chkDeleteFiles.Checked = $true
$grpMain.Controls.Add($chkDeleteFiles)

# Desativar se compressão partilhada
$chkShared.Add_CheckedChanged({
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
			$format = [string]$cmbFormat.SelectedItem
			$dest = Get-UniqueArchivePath -Parent $dir -BaseName $archiveBaseName -Extension $format
			$lblStatus.Text = "Estado: a comprimir ficheiros em $folderName ($currentIndex de $totalCount)"
			$progressFolder.Value = 0
			Add-LogLine -TextBox $txtLog -Text "A processar ficheiros em: $dir"
			Add-LogLine -TextBox $txtLog -Text "Destino: $dest"
			$argList = Build-SevenZipArguments `
				-ArchivePath $dest `
				-FolderName ($files -join ' ') `
				-Format $format `
				-Level ([string]$cmbLevel.SelectedItem) `
				-Threads $numThreads.Value `
				-DeleteAfter $false `
				-CompressShared $chkShared.Checked `
				-CreateSfx $chkSfx.Checked `
				-VolumeSize $txtVolume.Text `
				-Password $txtPass1.Text `
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
