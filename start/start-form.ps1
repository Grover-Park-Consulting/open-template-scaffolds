# Start-up form for an Open Template Scaffolds run.
#
# Shows a window with the standing statements from start-form-text.md (beside this script) and one
# button. The first showing does not end until the button is pressed, the window is closed, or the
# countdown runs out. Pressing the button or running out of time then parks a copy of the window on
# the taskbar, where the developer can reopen it for the rest of the run. Closing the window with its
# X parks nothing. It never changes anything on the machine and writes nothing to disk.
#
#   start-form.ps1 [-TimeoutSeconds n]   first showing; if a parked copy already exists, it is brought
#                                        forward instead and the script returns at once
#   start-form.ps1 -Close                closes the parked copy (end of the run)
#   start-form.ps1 -Parked               internal: the parked copy itself; started by the first showing
#
# Exit codes:  0 = continued, brought forward, or closed   1 = window closed without continuing
#              2 = timed out                                3 = skipped (no desktop, forms unavailable,
#                                                               or text unreadable)
# The one line it prints says which, so a caller can read the outcome without the exit code.
# The caller sets -TimeoutSeconds below its own command limit, so the script ends itself first.
param(
    [int]$TimeoutSeconds = 100,
    [switch]$Parked,
    [switch]$Close,
    [string]$TextPath = (Join-Path $PSScriptRoot 'start-form-text.md')
)

$MutexName      = 'Local\OpenTemplateScaffolds.StartForm'
$ShowEventName  = 'Local\OpenTemplateScaffolds.StartForm.Show'
$CloseEventName = 'Local\OpenTemplateScaffolds.StartForm.Close'

function Open-Event([string]$name) {
    # Returns the named event if a parked copy has created it, otherwise $null.
    try { return [System.Threading.EventWaitHandle]::OpenExisting($name) } catch { return $null }
}

function Read-FormText([string]$path) {
    # Returns a hashtable of heading -> text, or $null when the file is missing or has no cards.
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $parts = @{}
    $name = $null
    $buffer = New-Object System.Collections.Generic.List[string]
    # Banner lines keep their spacing (the letter-spaced title relies on it); other text is reflowed.
    $finish = {
        $joined = ($buffer -join ' ') -replace '\*\*', ''
        if ($name -notlike 'Banner*') { $joined = $joined -replace '\s+', ' ' }
        $parts[$name] = $joined.Trim()
    }
    foreach ($line in (Get-Content -LiteralPath $path -Encoding UTF8)) {
        if ($line -match '^##\s+(.+?)\s*$') {
            if ($name) { . $finish }
            $name = $Matches[1]
            $buffer.Clear()
        }
        elseif ($name -and $line.Trim().Length -gt 0) {
            $buffer.Add($line.Trim())
        }
    }
    if ($name) { . $finish }
    if (-not ($parts.Keys | Where-Object { $_ -match '^Card \d+: text$' })) { return $null }
    return $parts
}

try {
    if ($Close) {
        $ev = Open-Event $CloseEventName
        if ($null -eq $ev) { Write-Output 'start-form: no parked window to close'; exit 0 }
        [void]$ev.Set()
        Write-Output 'start-form: parked window closed'
        exit 0
    }

    if (-not $Parked) {
        $ev = Open-Event $ShowEventName
        if ($null -ne $ev) {
            [void]$ev.Set()
            Write-Output 'start-form: parked window brought forward'
            exit 0
        }
    }

    if (-not [Environment]::UserInteractive) {
        Write-Output 'start-form: no desktop session; skipped'
        exit 3
    }
    $text = Read-FormText $TextPath
    if ($null -eq $text) {
        Write-Output 'start-form: text file missing or unreadable; skipped'
        exit 3
    }

    if ($Parked) {
        # One parked copy only: a second one exits at once.
        $createdNew = $false
        $script:mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$createdNew)
        if (-not $createdNew) { exit 0 }
        $showEvent  = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', $ShowEventName)
        $closeEvent = New-Object System.Threading.EventWaitHandle($false, 'AutoReset', $CloseEventName)
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $accent = [System.Drawing.Color]::FromArgb(31, 95, 139)
    $ink    = [System.Drawing.Color]::FromArgb(29, 39, 51)
    $paper  = [System.Drawing.Color]::FromArgb(246, 244, 239)

    # Large type. The window is as tall as the cards need, so they scroll only on a screen too small
    # to hold them.
    $area   = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $width  = [Math]::Min(1150, $area.Width - 40)
    $margin = 28
    $headFont = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
    $bodyFont = New-Object System.Drawing.Font('Segoe UI', 14)

    function Center([string]$s, [int]$cols) {
        $pad = [Math]::Max(0, $cols - $s.Length)
        $left = [int][Math]::Floor($pad / 2)
        return (' ' * $left) + $s + (' ' * ($pad - $left))
    }
    function Pick([string]$preferred, [string]$fallback) {
        if ($text.ContainsKey($preferred) -and $text[$preferred]) { return $text[$preferred] }
        return $text[$fallback]
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

    $panelTop = 176
    $inner    = $width - 2 * $margin - 36
    $labels   = @()
    $content  = 24
    foreach ($c in $cards) {
        $h = New-Object System.Windows.Forms.Label
        $h.Text = $c[0]
        $h.Font = $headFont
        $h.ForeColor = $accent
        $h.AutoSize = $true
        $h.MaximumSize = New-Object System.Drawing.Size($inner, 0)
        $h.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
        $b = New-Object System.Windows.Forms.Label
        $b.Text = $c[1]
        $b.Font = $bodyFont
        $b.AutoSize = $true
        $b.MaximumSize = New-Object System.Drawing.Size($inner, 0)
        $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
        $labels += $h, $b
        $content += 16 + $h.PreferredSize.Height + $b.PreferredSize.Height
    }
    $height = [Math]::Min(($area.Height - 40), ($panelTop + $content + 6 + 184))

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Open Template Scaffolds'
    $form.ClientSize = New-Object System.Drawing.Size($width, $height)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.MinimizeBox = [bool]$Parked
    $form.ShowInTaskbar = $true
    $form.Icon = [System.Drawing.SystemIcons]::Information
    $form.TopMost = -not $Parked
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
    if ($Parked) { $title.Text = Pick 'Window title, when reopened' 'Window title (the large line above the cards)' }
    else         { $title.Text = $text['Window title (the large line above the cards)'] }
    $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
    $title.AutoSize = $false
    $title.SetBounds($margin, 126, ($width - 2 * $margin), 44)
    $form.Controls.Add($title)

    $buttonHeight = 64
    $buttonTop    = $height - $buttonHeight - 18
    $countTop     = $buttonTop - 36
    $goTop        = $countTop - 66
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = 'TopDown'
    $panel.WrapContents = $false
    $panel.AutoScroll = $true
    $panel.SetBounds($margin, $panelTop, ($width - 2 * $margin), ($goTop - $panelTop - 6))
    $form.Controls.Add($panel)
    foreach ($l in $labels) { $panel.Controls.Add($l) }

    $go = New-Object System.Windows.Forms.Label
    $go.Text = $text['Line above the button']
    $go.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $go.AutoSize = $false
    $go.SetBounds($margin, $goTop, ($width - 2 * $margin), 62)
    $form.Controls.Add($go)

    $count = New-Object System.Windows.Forms.Label
    $count.Font = New-Object System.Drawing.Font('Segoe UI', 12)
    $count.ForeColor = $accent
    $count.AutoSize = $false
    $count.TextAlign = 'MiddleCenter'
    $count.SetBounds($margin, $countTop, ($width - 2 * $margin), 30)
    $form.Controls.Add($count)

    $button = New-Object System.Windows.Forms.Button
    if ($Parked) { $button.Text = Pick 'Button label, when reopened' 'Button label' }
    else         { $button.Text = $text['Button label'] }
    $button.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
    $button.SetBounds(([int](($width - 440) / 2)), $buttonTop, 440, $buttonHeight)
    $button.BackColor = $accent
    $button.ForeColor = [System.Drawing.Color]::White
    $button.FlatStyle = 'Flat'
    $form.Controls.Add($button)
    $form.AcceptButton = $button

    if ($Parked) {
        # Lives on the taskbar until -Close or the X. The button sets it aside again; a second
        # launch of the script signals it to come forward.
        $form.WindowState = 'Minimized'
        $button.Add_Click({ $form.WindowState = 'Minimized' })
        $poll = New-Object System.Windows.Forms.Timer
        $poll.Interval = 300
        $poll.Add_Tick({
            if ($closeEvent.WaitOne(0)) { $poll.Stop(); $form.Close(); return }
            if ($showEvent.WaitOne(0)) {
                $form.WindowState = 'Normal'
                $form.TopMost = $true
                $form.Activate()
                $form.TopMost = $false
            }
        })
        $poll.Start()
        [System.Windows.Forms.Application]::Run($form)
        exit 0
    }

    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $script:remaining = [Math]::Max(10, $TimeoutSeconds)
    $script:timedOut = $false
    $countText = $text['Countdown line']
    function Show-Count {
        $span = '{0}:{1:00}' -f [int][Math]::Floor($script:remaining / 60), ($script:remaining % 60)
        if ($countText) { $count.Text = $countText -replace '\{time\}', $span }
    }
    Show-Count
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        $script:remaining--
        Show-Count
        if ($script:remaining -le 0) { $script:timedOut = $true; $timer.Stop(); $form.Close() }
    })
    $timer.Start()

    $result = $form.ShowDialog()
    $timer.Stop()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK -and -not $script:timedOut) {
        Write-Output 'start-form: closed without continuing'
        exit 1
    }
    $child = Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-WindowStyle', 'Hidden', '-File', ('"' + $PSCommandPath + '"'),
        '-Parked', '-TextPath', ('"' + $TextPath + '"'))
    $how = if ($script:timedOut) { 'timed out' } else { 'continued' }
    Write-Output ('start-form: ' + $how + '; parked on the taskbar (process ' + $child.Id + ')')
    if ($script:timedOut) { exit 2 }
    exit 0
}
catch {
    Write-Output ('start-form: could not show the form (' + $_.Exception.Message + ')')
    exit 3
}
