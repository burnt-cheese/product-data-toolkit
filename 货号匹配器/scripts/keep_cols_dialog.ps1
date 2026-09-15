# keep_cols_dialog.ps1 — WinForms 列选择弹窗（generate_matched.py 调用）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File keep_cols_dialog.ps1 <in.json> <out.json>
#   in.json  : [[列号, 列名], ...]
#   out.json : [保留的列号, ...]（空数组 = 全部不保留，全部按 A006 匹配）
# 任何关闭路径（点确定 / 点全部不保留 / 关窗 / ESC）都会写出 out.json 防止调用方卡死
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$OutputEncoding = [Text.Encoding]::UTF8

if ($args.Count -ne 2) {
    Write-Error "Usage: powershell -File keep_cols_dialog.ps1 <in.json> <out.json>"
    exit 1
}
$inPath  = $args[0]
$outPath = $args[1]

if (-not (Test-Path $inPath)) {
    Write-Error "输入文件不存在: $inPath"
    exit 1
}

try {
    $raw = Get-Content -Raw -Path $inPath -Encoding UTF8
    $cols = $raw | ConvertFrom-Json
} catch {
    Write-Error "解析输入 JSON 失败: $_"
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = '选择要保留的输入列（不取 A006）'
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true
$form.ClientSize = New-Object System.Drawing.Size(640, 560)
$form.MinimumSize = New-Object System.Drawing.Size(440, 360)
$form.KeyPreview = $true

# 顶部说明
$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "勾选后，这些列将直接取自你的输入表、不再去 A006 匹配："
$lbl.Location = New-Object System.Drawing.Point(12, 10)
$lbl.AutoSize = $true
$lbl.MaximumSize = New-Object System.Drawing.Size(610, 0)
$form.Controls.Add($lbl)

# 任何非"确定"路径的兜底：写出 [] = 全部不保留
$form.Add_FormClosed({
    if (-not (Test-Path $outPath)) {
        '[]' | Set-Content -Path $outPath -Encoding UTF8
    }
})

# 可滚动 CheckList 区域
$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(8, 50)
$panel.Size = New-Object System.Drawing.Size(622, 420)
$panel.AutoScroll = $true
$panel.BorderStyle = 'FixedSingle'
$form.Controls.Add($panel)

$checkboxes = New-Object System.Collections.Generic.List[object]
$y = 4
foreach ($c in $cols) {
    $idx  = [int]$c[0]
    $name = [string]$c[1]
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = "$name   (第 $idx 列)"
    $cb.Location = New-Object System.Drawing.Point(14, $y)
    $cb.Size = New-Object System.Drawing.Size(580, 24)
    $cb.AutoSize = $false
    $cb.AutoEllipsis = $true
    $cb.Tag = $idx
    $cb.Checked = $false
    $panel.Controls.Add($cb)
    $checkboxes.Add($cb)
    $y += 26
}

# 底部固定按钮栏
$btnBar = New-Object System.Windows.Forms.Panel
$btnBar.Dock = 'Bottom'
$btnBar.Height = 56
$form.Controls.Add($btnBar)

$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = '确定'
$btnOk.Size = New-Object System.Drawing.Size(120, 36)
$btnOk.Location = New-Object System.Drawing.Point(160, 10)
$btnOk.Add_Click({
    $selected = New-Object System.Collections.Generic.List[int]
    foreach ($cb in $checkboxes) {
        if ($cb.Checked) { $selected.Add([int]$cb.Tag) }
    }
    $json = ConvertTo-Json -InputObject @($selected) -Compress
    Set-Content -Path $outPath -Value $json -Encoding UTF8
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$btnBar.Controls.Add($btnOk)

$btnNo = New-Object System.Windows.Forms.Button
$btnNo.Text = '全部不保留'
$btnNo.Size = New-Object System.Drawing.Size(120, 36)
$btnNo.Location = New-Object System.Drawing.Point(340, 10)
$btnNo.Add_Click({
    Set-Content -Path $outPath -Value '[]' -Encoding UTF8
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})
$btnBar.Controls.Add($btnNo)

# Enter = 确定；ESC = 全部不保留
$form.Add_KeyDown({
    if ($_.KeyCode -eq 'Escape') {
        Set-Content -Path $outPath -Value '[]' -Encoding UTF8
        $form.Close()
    } elseif ($_.KeyCode -eq 'Return') {
        $btnOk.PerformClick()
    }
})

# 默认焦点给"全部不保留"，避免误回车确认默认勾选（虽然默认都没勾）
$form.Shown.Add({
    $btnNo.Focus()
})

[void]$form.ShowDialog()
