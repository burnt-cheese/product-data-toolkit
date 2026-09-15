# station_pick.ps1 — 拖入文件后立即弹出的「目标站点」选择窗
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File station_pick.ps1 <out.txt>
#   out.txt : 单行站点 token（cn / es / gr），默认 cn
# 任何关闭路径都会写出 out.txt（默认 cn），防止调用方卡死
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$OutputEncoding = [Text.Encoding]::UTF8

if ($args.Count -ne 1) {
    Write-Error "Usage: powershell -File station_pick.ps1 <out.txt>"
    exit 1
}
$outPath = $args[0]
$default = "cn"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = '选择目标站点'
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true
$form.ClientSize = New-Object System.Drawing.Size(470, 300)
$form.MinimumSize = New-Object System.Drawing.Size(440, 260)
$form.KeyPreview = $true

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "拖入文件后，请选择本次目标站点："
$lbl.Location = New-Object System.Drawing.Point(16, 14)
$lbl.Size = New-Object System.Drawing.Size(440, 24)
$lbl.AutoSize = $false
$form.Controls.Add($lbl)

$lbl2 = New-Object System.Windows.Forms.Label
$lbl2.Text = "中国站走完整流程（分类 + 自动上传）；西班牙 / 希腊站仅匹配信息生成表格，不上传。"
$lbl2.Location = New-Object System.Drawing.Point(16, 38)
$lbl2.Size = New-Object System.Drawing.Size(440, 34)
$lbl2.AutoSize = $false
$form.Controls.Add($lbl2)

# 兜底：任何关闭路径都写出默认 cn，避免 bat 卡死
$form.Add_FormClosed({
    if (-not (Test-Path $outPath)) {
        [System.IO.File]::WriteAllText($outPath, "cn", [System.Text.Encoding]::ASCII)
    }
})

# 单选组
$group = New-Object System.Windows.Forms.GroupBox
$group.Text = "站点"
$group.Location = New-Object System.Drawing.Point(16, 78)
$group.Size = New-Object System.Drawing.Size(438, 150)
$form.Controls.Add($group)

$stations = @(
    @{key="cn"; name="中国站（完整流程：分类 + 自动上传）"},
    @{key="es"; name="西班牙站（仅匹配信息，生成上传表格）"},
    @{key="gr"; name="希腊站（仅匹配信息，生成上传表格）"}
)
$radios = New-Object System.Collections.Generic.List[object]
$y = 24
foreach ($s in $stations) {
    $rb = New-Object System.Windows.Forms.RadioButton
    $rb.Text = $s.name
    $rb.Location = New-Object System.Drawing.Point(16, $y)
    $rb.Size = New-Object System.Drawing.Size(410, 30)
    $rb.Tag = $s.key
    if ($s.key -eq $default) { $rb.Checked = $true }
    $group.Controls.Add($rb)
    $radios.Add($rb)
    $y += 38
}

# 底部按钮
$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = '确定'
$btnOk.Size = New-Object System.Drawing.Size(120, 36)
$btnOk.Location = New-Object System.Drawing.Point(110, 248)
$btnOk.Add_Click({
    $sel = "cn"
    foreach ($rb in $radios) { if ($rb.Checked) { $sel = $rb.Tag } }
    [System.IO.File]::WriteAllText($outPath, $sel, [System.Text.Encoding]::ASCII)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})
$form.Controls.Add($btnOk)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = '取消'
$btnCancel.Size = New-Object System.Drawing.Size(120, 36)
$btnCancel.Location = New-Object System.Drawing.Point(250, 248)
$btnCancel.Add_Click({
    # 取消 = 沿用默认中国站
    [System.IO.File]::WriteAllText($outPath, "cn", [System.Text.Encoding]::ASCII)
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})
$form.Controls.Add($btnCancel)

# ESC = 取消（默认 cn）
$form.Add_KeyDown({
    if ($_.KeyCode -eq 'Escape') {
        [System.IO.File]::WriteAllText($outPath, "cn", [System.Text.Encoding]::ASCII)
        $form.Close()
    } elseif ($_.KeyCode -eq 'Return') {
        $btnOk.PerformClick()
    }
})

[void]$form.ShowDialog()
