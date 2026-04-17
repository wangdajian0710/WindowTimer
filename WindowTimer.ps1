# Window Timer v8 - Minimize + Sort + Time Fix + Remove
try {
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class W {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
}
"@

$script:startTimes  = @{}  # key: "pid|name", value: actual DateTime
$script:ghostKeys   = @()  # keys to skip this cycle (PID reuse ghost)
$script:fgHwnd = [IntPtr]::Zero
$script:lastKey = ""
$script:formRef = $null
$script:tracked = @{}  # whitelist: empty = all, pid = only this process
$script:blocked  = @{}  # blocklist: title prefixes to ignore
# Initialize blocklist with titles to ignore
$script:blocked['Zebra Sample-offers'] = $true
$script:blocked['Zebra Sample Questionnaire Assistant'] = $true
$script:blocked['Microsoft Text Input Application'] = $true
$script:blocked['微信'] = $true
$script:blocked['Octo Browser'] = $true
$script:blocked['计算器'] = $true
$script:blocked['指纹打开'] = $true
$script:blocked['DeepL'] = $true
$script:blocked['网易有道翻译'] = $true
$script:blocked['QClaw'] = $true
$script:blocked['181.xlsx'] = $true
$script:blocked['模拟输入-界首本发'] = $true
$script:blocked['Window Timer'] = $true
$script:blocked['IPWEB和另外'] = $true
$script:dragSrcIdx = -1
$script:dragSrcTag = $null
$script:dragSrcText = $null
$script:dragSrcBack = $null
$script:dragSrcFore = $null

function UpdateTitle {
    if ($script:tracked.Count -eq 0) {
        $lblHeader.Text = '  Window Timer  [All Windows]'
    } else {
        $lblHeader.Text = '  Window Timer  [Tracking ' + $script:tracked.Count + ' windows]'
    }
}

function FmtTime($ts) {
    $h = [int]($ts.TotalHours)
    $m = [int]($ts.TotalMinutes % 60)
    $s = [int]($ts.TotalSeconds % 60)
    if ($ts.TotalHours -ge 24) {
        $r = $h.ToString() + 'd' + $ts.Hours.ToString('D2') + 'h'; return $r
    }
    if ($ts.TotalHours -ge 1) {
        $r = $ts.Hours.ToString() + 'h' + $m.ToString('D2') + 'm'; return $r
    }
    if ($ts.TotalMinutes -ge 1) {
        $r = $m.ToString() + 'm' + $s.ToString('D2') + 's'; return $r
    }
    $r = $s.ToString('D2') + 's'; return $r
}

$fontFamily = 'Segoe UI'
$fonts = [System.Drawing.FontFamily]::Families
foreach ($f in $fonts) { if ($f.Name -eq 'Segoe UI') { $fontFamily = 'Segoe UI'; break } }

$form = New-Object System.Windows.Forms.Form
$script:formRef = $form
$form.Text = 'Window Timer'
$form.Size = New-Object System.Drawing.Size(620, 520)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(14, 16, 22)
$form.FormBorderStyle = 'FixedToolWindow'

# Title bar
$titleBar = New-Object System.Windows.Forms.Panel
$titleBar.Dock = 'Top'; $titleBar.Height = 36
$titleBar.BackColor = [System.Drawing.Color]::FromArgb(22, 26, 38)
$titleBar.Cursor = 'SizeAll'

$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = '  Window Timer'
$lblHeader.Dock = 'Fill'
$lblHeader.Font = New-Object System.Drawing.Font($fontFamily, 11)
$lblHeader.ForeColor = [System.Drawing.Color]::White
$lblHeader.TextAlign = 'MiddleLeft'; $lblHeader.Cursor = 'SizeAll'
[void]$titleBar.Controls.Add($lblHeader)

# Minimize + Close buttons (right side)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Size = New-Object System.Drawing.Size(36, 36)
$btnClose.Dock = 'Right'
$btnClose.Text = 'X'
$btnClose.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnClose.ForeColor = [System.Drawing.Color]::White
$btnClose.BackColor = [System.Drawing.Color]::Transparent
$btnClose.FlatStyle = 'Flat'
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Cursor = 'Hand'
$btnClose.Add_Click({ $master.Stop(); $ni.Dispose(); $form.Close(); [System.Windows.Forms.Application]::Exit() })

$btnMin = New-Object System.Windows.Forms.Button
$btnMin.Size = New-Object System.Drawing.Size(36, 36)
$btnMin.Dock = 'Right'
$btnMin.Text = '_'
$btnMin.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnMin.ForeColor = [System.Drawing.Color]::White
$btnMin.BackColor = [System.Drawing.Color]::Transparent
$btnMin.FlatStyle = 'Flat'
$btnMin.FlatAppearance.BorderSize = 0
$btnMin.Cursor = 'Hand'
$btnMin.Add_Click({ $script:formRef.Hide() })

[void]$titleBar.Controls.Add($btnClose)
[void]$titleBar.Controls.Add($btnMin)

# Status bar
$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = 'Bottom'; $statusBar.Height = 24
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(18, 22, 30)
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Dock = 'Fill'; $lblStatus.Text = '  ...'
$lblStatus.Font = New-Object System.Drawing.Font('Consolas', 9)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(120, 130, 160)
$lblStatus.TextAlign = 'MiddleLeft'
[void]$statusBar.Controls.Add($lblStatus)

# ListView
$script:lv = New-Object System.Windows.Forms.ListView
$lv = $script:lv
$lv.Dock = 'Fill'
$lv.View = 'Details'
$lv.BackColor = [System.Drawing.Color]::FromArgb(14, 16, 22)
$lv.ForeColor = [System.Drawing.Color]::White
$lv.Font = New-Object System.Drawing.Font('Consolas', 11)
$lv.BorderStyle = 'None'
$lv.HeaderStyle = 'None'
$lv.FullRowSelect = $true
$lv.MultiSelect = $false
$lv.HideSelection = $false
$lv.Cursor = 'Hand'
$lv.AllowDrop = $true
$prop = $lv.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic)
$prop.SetValue($lv, $true, $null)
[void]$lv.Columns.Add('Entry', 570)

# Right-click menu — whitelist mode
$lv.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip

$miAdd = New-Object System.Windows.Forms.ToolStripMenuItem('Add to tracking')
$miAdd.Add_Click({
    $sel = $script:lv.SelectedItems
    if ($sel.Count -eq 0) { return }
    $tag = $sel[0].Tag
    if ($null -eq $tag) { return }
    $pid = [int]($tag.Split('|')[0])
    $script:tracked[$pid] = $true
    $script:lastKey = ''
    UpdateTitle
})

$miRemove = New-Object System.Windows.Forms.ToolStripMenuItem('Remove from tracking')
$miRemove.Add_Click({
    $sel = $script:lv.SelectedItems
    if ($sel.Count -eq 0) { return }
    $tag = $sel[0].Tag
    if ($null -ne $tag) { $script:tracked.Remove([int]($tag.Split('|')[0])) }
    $script:lv.Items.Remove($sel[0])
    $script:lastKey = ''
    UpdateTitle
})

$miClear = New-Object System.Windows.Forms.ToolStripMenuItem('Reset to all windows')
$miClear.Add_Click({
    $script:tracked.Clear()
    $script:lastKey = ''
    UpdateTitle
})

$miIgnore = New-Object System.Windows.Forms.ToolStripMenuItem('Ignore this window')
$miIgnore.Add_Click({
    $sel = $script:lv.SelectedItems
    if ($sel.Count -eq 0) { return }
    $tag = $sel[0].Tag
    if ($null -eq $tag) { return }
    $pid = [int]($tag.Split('|')[0])
    # Find the window title
    $p = Get-Process -Id $pid -ErrorAction SilentlyContinue
    if (-not $p) { return }
    $title = $p.MainWindowTitle
    if ($title.Length -eq 0) { return }
    # Add first 30 chars as blocklist key (avoid overfitting to long titles)
    $key = $title.Substring(0, [Math]::Min(30, $title.Length))
    $script:blocked[$key] = $true
    $script:lv.Items.Remove($sel[0])
    $script:lastKey = ''
})

$lv.ContextMenuStrip.Items.Add($miAdd)
$lv.ContextMenuStrip.Items.Add($miRemove)
$lv.ContextMenuStrip.Items.Add($miClear)
$lv.ContextMenuStrip.Items.Add($miIgnore)

[void]$form.Controls.AddRange(@($lv, $statusBar, $titleBar))

# Drag window
$script:dragOff = $null
$titleBar.Add_MouseDown({ param($s,$e)
    if ($e.Button -eq 'Left') { $script:dragOff = $e.Location }
})
$titleBar.Add_MouseMove({ param($s,$e)
    if ($script:dragOff) {
        $ptX = $form.Location.X + $e.X - $script:dragOff.X
        $ptY = $form.Location.Y + $e.Y - $script:dragOff.Y
        $form.Location = [System.Drawing.Point]::new($ptX, $ptY)
    }
})
$titleBar.Add_MouseUp({ $script:dragOff = $null })

# Click to flash
$script:lv.Add_Click({
    $sel = $script:lv.SelectedItems
    if ($sel.Count -eq 0) { return }
    $rowId = $sel[0].Tag
    if ($null -eq $rowId) { return }
    $pid = [int]([string]$rowId.Split('|')[0])
    $p = Get-Process -Id $pid -ErrorAction SilentlyContinue
    if (-not $p) { return }
    $h = $p.MainWindowHandle
    if ($h -eq [IntPtr]::Zero) { return }
    $flash = New-Object System.Windows.Forms.Timer
    $flash.Interval = 150
    $script:flashN = 0
    $script:flashH = $h
    $flash.Add_Tick({
        $script:flashN++
        if ($script:flashN -eq 1) {
            [W]::ShowWindow($script:flashH, 9) | Out-Null
            [W]::SetForegroundWindow($script:flashH) | Out-Null
        }
        if ($script:flashN -ge 3) { $flash.Stop(); $flash.Dispose() }
    })
    $flash.Start()
})

# Drag to reorder list
$script:lv.Add_MouseDown({ param($s,$e)
    if ($e.Button -ne 'Left') { return }
    $hit = $script:lv.HitTest($e.Location)
    if ($null -eq $hit -or $null -eq $hit.Item) { return }
    $srcItem = $hit.Item
    $script:dragSrcIdx = $srcItem.Index
    $script:dragSrcTag = $srcItem.Tag
    $script:dragSrcText = $srcItem.Text
    $script:dragSrcBack = $srcItem.BackColor
    $script:dragSrcFore = $srcItem.ForeColor
})

$script:lv.Add_MouseMove({ param($s,$e)
    if ($script:dragSrcIdx -lt 0 -or $e.Button -ne 'Left') { return }
    $script:lv.DoDragDrop($null, [System.Windows.Forms.DragDropEffects]::Link)
})

$script:lv.Add_DragOver({ param($s,$e)
    $e.Effect = [System.Windows.Forms.DragDropEffects]::Link
})

$script:lv.Add_DragDrop({ param($s,$e)
    if ($script:dragSrcIdx -lt 0) { return }
    $pt = $script:lv.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
    $hit = $script:lv.HitTest($pt)
    $dstIdx = 0
    if ($null -ne $hit -and $null -ne $hit.Item) { $dstIdx = $hit.Item.Index }
    $srcIdx = $script:dragSrcIdx
    $script:dragSrcIdx = -1
    if ($dstIdx -eq $srcIdx) { return }
    $script:lv.BeginUpdate()
    $srcItem = $script:lv.Items[$srcIdx]
    $tmpTag = $script:dragSrcTag
    $tmpText = $script:dragSrcText
    $tmpBack = $script:dragSrcBack
    $tmpFore = $script:dragSrcFore
    $script:lv.Items.Remove($srcItem)
    $newItem = New-Object System.Windows.Forms.ListViewItem($tmpText)
    $newItem.Tag = $tmpTag
    $newItem.BackColor = $tmpBack
    $newItem.ForeColor = $tmpFore
    if ($dstIdx -lt $script:lv.Items.Count) {
        $script:lv.Items.Insert($dstIdx, $newItem)
    } else {
        [void]$script:lv.Items.Add($newItem)
    }
    $newItem.Selected = $true
    $script:lv.EndUpdate()
})

$script:lv.Add_MouseUp({ param($s,$e)
    $script:dragSrcIdx = -1
})

# Refresh
function DoRefresh {
    try {
        $now = Get-Date
        $fgHwnd = [W]::GetForegroundWindow()
        if ($fgHwnd -ne $form.Handle) { $script:fgHwnd = $fgHwnd }

        $allProcs = @(Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle.Length -gt 0 } | Sort-Object Id)

        # Blocklist filter: skip windows whose title starts with any blocked prefix
        if ($script:blocked.Count -gt 0) {
            $allProcs = @($allProcs | Where-Object {
                $t = $_.MainWindowTitle
                $skip = $false
                foreach ($b in $script:blocked.Keys) {
                    if ($t.StartsWith($b)) { $skip = $true; break }
                }
                -not $skip
            })
        }

        # Whitelist filter: if non-empty, only show tracked pids
        if ($script:tracked.Count -gt 0) {
            $procs = @($allProcs | Where-Object { $script:tracked.ContainsKey($_.Id) })
        } else {
            $procs = $allProcs
        }

        $newKey = ($procs | ForEach-Object { $_.Id }) -join ','
        $needRebuild = ($newKey -ne $script:lastKey)
        $script:lastKey = $newKey

        # Build composite key: "pid|name" to detect PID reuse
        $curKeys = @{}
        foreach ($p in $procs) {
            $key = $p.Id.ToString() + '|' + $p.ProcessName
            $curKeys[$key] = $true
            if (-not $script:startTimes.ContainsKey($key)) {
                # Use real process StartTime if available; otherwise fall back to now
                try {
                    $script:startTimes[$key] = $p.StartTime
                } catch {
                    $script:startTimes[$key] = $now
                }
            }
        }
        # Cleanup: remove keys for PIDs that no longer exist (use allProcs to avoid
        # removing blocklisted entries, which would lose their elapsed time)
        $rmKeys = @()
        foreach ($k in $script:startTimes.Keys) {
            $pidPart = [int]($k.Split('|')[0])
            $found = $false
            foreach ($p in $allProcs) {
                if ($p.Id -eq $pidPart) { $found = $true; break }
            }
            if (-not $found) { $rmKeys += $k }
        }
        foreach ($k in $rmKeys) { $script:startTimes.Remove($k) }

        if ($needRebuild) {
            $script:lv.BeginUpdate()
            $script:lv.Items.Clear()
            $idx = 1
            foreach ($p in $procs) {
                $key = $p.Id.ToString() + '|' + $p.ProcessName
                $el = $now - $script:startTimes[$key]
                $ts = FmtTime $el
                $title = $p.MainWindowTitle
                $isFg = ($p.MainWindowHandle -eq $script:fgHwnd)
                $numStr = ' #' + $idx.ToString()
                $tsStr = '  ' + $ts
                $rowText = $numStr + $tsStr + '  ' + $title
                $li = New-Object System.Windows.Forms.ListViewItem($rowText)
                $li.Tag = $key
                if ($isFg) {
                    $li.BackColor = [System.Drawing.Color]::FromArgb(36, 56, 84)
                    $li.ForeColor = [System.Drawing.Color]::FromArgb(140, 255, 200)
                }
                [void]$script:lv.Items.Add($li)
                $idx++
            }
            $script:lv.EndUpdate()
        } else {
            for ($i = 0; $i -lt $script:lv.Items.Count; $i++) {
                $li = $script:lv.Items[$i]
                $key = $li.Tag
                if ($null -eq $key) { continue }
                if (-not $script:startTimes.ContainsKey($key)) { continue }
                $el = $now - $script:startTimes[$key]
                $ts = FmtTime $el
                $rowPid = [int]($key.Split('|')[0])
                $foundP = $null
                foreach ($pp in $procs) { if ($pp.Id -eq $rowPid) { $foundP = $pp; break } }
                if ($null -eq $foundP) { continue }
                $title = $foundP.MainWindowTitle
                $isFg = ($foundP.MainWindowHandle -eq $script:fgHwnd)
                $numStr = ' #' + ($i + 1).ToString()
                $tsStr = '  ' + $ts
                $li.Text = $numStr + $tsStr + '  ' + $title
                if ($isFg) {
                    $li.BackColor = [System.Drawing.Color]::FromArgb(36, 56, 84)
                    $li.ForeColor = [System.Drawing.Color]::FromArgb(140, 255, 200)
                } else {
                    $li.BackColor = [System.Drawing.Color]::FromArgb(14, 16, 22)
                    $li.ForeColor = [System.Drawing.Color]::White
                }
            }
        }

        $fgP = $null
        foreach ($pp in $procs) { if ($pp.MainWindowHandle -eq $script:fgHwnd) { $fgP = $pp; break } }
        if ($fgP) {
            $modeLabel = if ($script:tracked.Count -gt 0) { '[' + $script:tracked.Count + '] ' } else { '' }
            $lblStatus.Text = '  ' + $modeLabel + 'Active: ' + $fgP.MainWindowTitle
        } else {
            $modeLabel = if ($script:tracked.Count -gt 0) { '[' + $script:tracked.Count + '] ' } else { '' }
            $lblStatus.Text = '  ' + $modeLabel + 'Windows: ' + $procs.Count.ToString()
        }
    } catch {
        # skip
    }
}

$master = New-Object System.Windows.Forms.Timer
$master.Interval = 1000
$master.Add_Tick({ DoRefresh })

$ni = New-Object System.Windows.Forms.NotifyIcon
$ni.Icon = [System.Drawing.SystemIcons]::Clock
$ni.Visible = $true; $ni.Text = 'Window Timer'
$ni.Add_DoubleClick({ $form.Show(); $form.Activate() })
$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miShow = New-Object System.Windows.Forms.ToolStripMenuItem('Show')
$miShow.Add_Click({ $form.Show(); $form.Activate() })
$miExit = New-Object System.Windows.Forms.ToolStripMenuItem('Exit')
$miExit.Add_Click({ $master.Stop(); $ni.Dispose(); $form.Close(); [System.Windows.Forms.Application]::Exit() })
$menu.Items.Add($miShow)
$menu.Items.Add($miExit)
$ni.ContextMenuStrip = $menu

$form.Add_Closing({ $master.Stop(); $ni.Dispose() })
UpdateTitle
$master.Start()
DoRefresh
$form.Show()
[System.Windows.Forms.Application]::Run()

} catch {
    $msg = $_.Exception.Message
    $st = $_.ScriptStackTrace
    [System.Windows.Forms.MessageBox]::Show('Error: ' + $msg + [Environment]::NewLine + [Environment]::NewLine + $st, 'Window Timer Error')
}
