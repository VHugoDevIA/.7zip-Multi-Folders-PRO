$Form = New-Object System.Windows.Forms.Form
$Form.Text = '7zip - Multi Folders PRO'
$Form.WindowState = 'Maximized'

$Label = New-Object System.Windows.Forms.Label
$Label.Text = 'Select the folders to compress:'
$Label.Size = New-Object System.Drawing.Size(200,20)
$Label.Location = New-Object System.Drawing.Point(10,10)

$ButtonBrowse = New-Object System.Windows.Forms.Button
$ButtonBrowse.Text = 'Browse'
$ButtonBrowse.Location = New-Object System.Drawing.Point(10,40)

$ButtonCompress = New-Object System.Windows.Forms.Button
$ButtonCompress.Text = 'Compress'
$ButtonCompress.Location = New-Object System.Drawing.Point(10,70)

# Dark Mode
$Form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$Label.ForeColor = [System.Drawing.Color]::White
$ButtonBrowse.BackColor = [System.Drawing.Color]::DarkSlateGray
$ButtonCompress.BackColor = [System.Drawing.Color]::DarkSlateGray

# Define event for theme switcher
$ThemeSwitcherButton = New-Object System.Windows.Forms.Button
$ThemeSwitcherButton.Text = 'Switch Theme'
$ThemeSwitcherButton.Location = New-Object System.Drawing.Point(10,100)
$Form.Controls.Add($ThemeSwitcherButton)

$Form.Controls.Add($Label)
$Form.Controls.Add($ButtonBrowse)
$Form.Controls.Add($ButtonCompress)

$ThemeSwitcherButton.Add_Click({
    if ($Form.BackColor -eq [System.Drawing.Color]::FromArgb(30, 30, 30)) {
        $Form.BackColor = [System.Drawing.Color]::White
        $Label.ForeColor = [System.Drawing.Color]::Black
        $ButtonBrowse.BackColor = [System.Drawing.Color]::LightGray
        $ButtonCompress.BackColor = [System.Drawing.Color]::LightGray
    } else {
        $Form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $Label.ForeColor = [System.Drawing.Color]::White
        $ButtonBrowse.BackColor = [System.Drawing.Color]::DarkSlateGray
        $ButtonCompress.BackColor = [System.Drawing.Color]::DarkSlateGray
    }
})

$Form.Controls.Add($ThemeSwitcherButton)
$Form.Add_Shown({$Form.Activate()})
[void]$Form.ShowDialog()