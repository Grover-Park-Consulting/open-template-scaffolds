# Start-up form for an Open Template Scaffolds run.
#
# Shows a window with the standing statements from start-form-text.md (beside this script) and one
# button. The script does not end until the button is pressed, the window is closed, or the time
# runs out. It never changes anything on the machine and writes nothing to disk.
#
# Exit codes:  0 = Continue pressed     1 = window closed without continuing
#              2 = timed out            3 = skipped (no desktop, forms unavailable, or text unreadable)
# The one line it prints says which, so a caller can read the outcome without the exit code.
# The default wait (100 seconds) is shorter than a two-minute command limit, so it ends itself first.
param(
    [int]$TimeoutSeconds = 100,
    [string]$TextPath = (Join-Path $PSScriptRoot 'start-form-text.md')
)

function Read-FormText([string]$path) {
    # Returns a hashtable of heading -> text, or $null when the file is missing or has no cards.
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $parts = @{}
    $name = $null
    $buffer = New-Object System.Collections.Generic.List[string]
    foreach ($line in (Get-Content -LiteralPath $path -Encoding UTF8)) {
        if ($line -match '^##\s+(.+?)\s*$') {
            if ($name) { $parts[$name] = (($buffer -join ' ') -replace '\*\*', '' -replace '\s+', ' ').Trim() }
            $name = $Matches[1]
            $buffer.Clear()
        }
        elseif ($name -and $line.Trim().Length -gt 0) {
            $buffer.Add($line.Trim())
        }
    }
    if ($name) { $parts[$name] = (($buffer -join ' ') -replace '\*\*', '' -replace '\s+', ' ').Trim() }
    if (-not ($parts.Keys | Where-Object { $_ -match '^Card \d+: text$' })) { return $null }
    return $parts
}

try {
    if (-not [Environment]::UserInteractive) {
        Write-Output 'start-form: no desktop session; skipped'
        exit 3
    }
    $text = Read-FormText $TextPath
    if ($null -eq $text) {
        Write-Output 'start-form: text file missing or unreadable; skipped'
        exit 3
    }
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $accent = [System.Drawing.Color]::FromArgb(31, 95, 139)
    $ink    = [System.Drawing.Color]::FromArgb(29, 39, 51)
    $paper  = [System.Drawing.Color]::FromArgb(246, 244, 239)

    # Large type, sized to fit the screen's usable area.
    $area   = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $width  = [Math]::Min(1000, $area.Width - 40)
    $height = [Math]::Min(900, $area.Height - 40)
    $margin = 28

    function Center([string]$s, [int]$cols) {
        $pad = [Math]::Max(0, $cols - $s.Length)
        $left = [int][Math]::Floor($pad / 2)
        return (' ' * $left) + $s + (' ' * ($pad - $left))
    }
    $line1 = $text['Banner, line 1']
    $line2 = $text['Banner, line 2']
    $cols  = [Math]::Max(50, ([Math]::Max($line1.Length, $line2.Length) + 4))
    $bar   = '+' + ('-' * $cols) + '+'
    $banner = @($bar, ('|' + (Center $line1 $cols) + '|'), ('|' + (Center $line2 $cols) + '|'), $bar) -join "`r`n"

    $cards = @()
    for ($n = 1; $text.ContainsKey("Card ${n}: text"); $n++) {
        $cards += , @($text["Card ${n}: heading"], $text["Card ${n}: text"])
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Open Template Scaffolds'
    $form.ClientSize = New-Object System.Drawing.Size($width, $height)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    $form.BackColor = $paper
    $form.ForeColor = $ink
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 14)

    $bannerLabel = New-Object System.Windows.Forms.Label
    $bannerLabel.Text = $banner
    $bannerLabel.Font = New-Object System.Drawing.Font('Consolas', 15, [System.Drawing.FontStyle]::Bold)
    $bannerLabel.ForeColor = $accent
    $bannerLabel.AutoSize = $false
    $bannerLabel.TextAlign = 'MiddleCenter'
    $bannerLabel.SetBounds(0, 10, $width, 110)
    $form.Controls.Add($bannerLabel)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = $text['Window title (the large line above the cards)']
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
    $title.AutoSize = $false
    $title.SetBounds($margin, 126, ($width - 2 * $margin), 44)
    $form.Controls.Add($title)

    $buttonHeight = 64
    $goTop     = $height - $buttonHeight - 28 - 44
    $buttonTop = $height - $buttonHeight - 18
    $panelTop  = 176
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false
    $panel.AutoScroll = $true
    $panel.SetBounds($margin, $panelTop, ($width - 2 * $margin), ($goTop - $panelTop - 6))
    $form.Controls.Add($panel)
    $inner = $panel.Width - 36
    foreach ($c in $cards) {
        $h = New-Object System.Windows.Forms.Label
        $h.Text = $c[0]
        $h.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
        $h.ForeColor = $accent
        $h.AutoSize = $true
        $h.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
        $panel.Controls.Add($h)
        $b = New-Object System.Windows.Forms.Label
        $b.Text = $c[1]
        $b.Font = New-Object System.Drawing.Font('Segoe UI', 14)
        $b.AutoSize = $true
        $b.MaximumSize = New-Object System.Drawing.Size($inner, 0)
        $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        $panel.Controls.Add($b)
    }

    $go = New-Object System.Windows.Forms.Label
    $go.Text = $text['Line above the button']
    $go.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $go.AutoSize = $false
    $go.SetBounds($margin, $goTop, ($width - 2 * $margin), 40)
    $form.Controls.Add($go)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text['Button label']
    $button.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
    $button.SetBounds(([int](($width - 440) / 2)), $buttonTop, 440, $buttonHeight)
    $button.BackColor = $accent
    $button.ForeColor = [System.Drawing.Color]::White
    $button.FlatStyle = 'Flat'
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($button)
    $form.AcceptButton = $button

    $script:timedOut = $false
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [Math]::Max(1, $TimeoutSeconds) * 1000
    $timer.Add_Tick({ $script:timedOut = $true; $timer.Stop(); $form.Close() })
    $timer.Start()

    $result = $form.ShowDialog()
    $timer.Stop()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Output 'start-form: continued'
        exit 0
    }
    if ($script:timedOut) {
        Write-Output 'start-form: timed out'
        exit 2
    }
    Write-Output 'start-form: closed without continuing'
    exit 1
}
catch {
    Write-Output ('start-form: could not show the form (' + $_.Exception.Message + ')')
    exit 3
}
