# review_confirm.ps1 — 待审核商品分类填回「确认按钮」窗口（Windows Forms）
# =====================================================================
# 由 一键全流程.bat 在用户关掉 Excel 后调用：
#   powershell -NoProfile -ExecutionPolicy Bypass -File review_confirm.ps1 <候选JSON>
#
# 行为：
#   - 读取 fillback.py prepare 输出的 JSON（{resolved:[...], unresolved:[...]}）
#   - 弹出窗口：标题 + 说明 + 滚动列表(已解决待填回项)
#              + 仍为待确认的警告 + 「确认回填并上传」/「取消」按钮
#   - 点「确认回填并上传」 -> exit 0
#   - 点「取消」(或右上角X) -> exit 1
#
# 设计：纯 UI，不碰任何 xlsx。逻辑(填回/清理)由 fillback.py apply 在 exit 0 后执行。
param(
    [Parameter(Mandatory=$true)][string]$JsonPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 读取候选 JSON（fillback 用 ensure_ascii=False 写出，UTF-8 无 BOM）
if (-not (Test-Path $JsonPath)) {
    [System.Windows.Forms.MessageBox]::Show("找不到候选清单文件：`n$JsonPath", "错误",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit 2
}
$raw = Get-Content -Raw -Encoding UTF8 $JsonPath
try {
    $data = $raw | ConvertFrom-Json
} catch {
    [System.Windows.Forms.MessageBox]::Show("候选清单 JSON 解析失败：`n$_", "错误",
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    exit 2
}

$resolved   = @($data.resolved)
$unresolved = @($data.unresolved)

$form = New-Object System.Windows.Forms.Form
$form.Text = "待审核商品分类填回确认"
$form.Size = New-Object System.Drawing.Size(660, 540)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.MinimizeBox = $false
$form.MaximizeBox = $false
$form.FormBorderStyle = "FixedDialog"

# 说明
$lbl = New-Object System.Windows.Forms.Label
$lbl.Location = New-Object System.Drawing.Point(16, 14)
$lbl.Size = New-Object System.Drawing.Size(624, 22)
$lbl.Text = "以下待审核商品你已填好有效分类，确认将回填主数据库，并清理上传辅助列后上传："
$form.Controls.Add($lbl)

# 已自动采纳区（程序推荐分类已自动写入主表）
$autoApplied = @($data.auto_applied)
if ($autoApplied.Count -gt 0) {
    $lblAuto = New-Object System.Windows.Forms.Label
    $lblAuto.Location = New-Object System.Drawing.Point(16, 40)
    $lblAuto.Size = New-Object System.Drawing.Size(624, 18)
    $lblAuto.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)
    $lblAuto.Text = "已自动采纳 $($autoApplied.Count) 条程序推荐分类到主表（无需你手动拷贝）："
    $form.Controls.Add($lblAuto)

    $lstAuto = New-Object System.Windows.Forms.ListBox
    $lstAuto.Location = New-Object System.Drawing.Point(16, 62)
    $lstAuto.Size = New-Object System.Drawing.Size(624, 100)
    $lstAuto.Anchor = "Left,Top,Right"
    $lstAuto.IntegralHeight = $false
    $lstAuto.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)
    foreach ($a in $autoApplied) {
        $lstAuto.Items.Add("$($a.name) [$($a.code)]  ->  $($a.chain)")
    }
    $form.Controls.Add($lstAuto)

    # "将填回主数据库"小标题
    $lblBack = New-Object System.Windows.Forms.Label
    $lblBack.Location = New-Object System.Drawing.Point(16, 172)
    $lblBack.Size = New-Object System.Drawing.Size(624, 18)
    $lblBack.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $lblBack.Text = "将填回主数据库："
    $form.Controls.Add($lblBack)

    $listTop = 192
    $listHeight = 200
    $warnTop = 400
    $btnTop = 446
} else {
    $listTop = 62
    $listHeight = 320
    $warnTop = 392
    $btnTop = 446
}

# 滚动列表（将填回主数据库的项）
$list = New-Object System.Windows.Forms.ListBox
$list.Location = New-Object System.Drawing.Point(16, $listTop)
$list.Size = New-Object System.Drawing.Size(624, $listHeight)
$list.Anchor = "Left,Top,Right"
$list.IntegralHeight = $false
$list.Font = New-Object System.Drawing.Font("Microsoft YaHei", 9)
if ($resolved.Count -eq 0) {
    $list.Items.Add("（无待审核项需回填）")
} else {
    foreach ($r in $resolved) {
        $list.Items.Add("$($r.name) [$($r.code)]  ->  $($r.chain)")
    }
}
$form.Controls.Add($list)

# 仍为待确认的警告
if ($unresolved.Count -gt 0) {
    $warn = New-Object System.Windows.Forms.Label
    $warn.Location = New-Object System.Drawing.Point(16, $warnTop)
    $warn.Size = New-Object System.Drawing.Size(624, 40)
    $warn.ForeColor = [System.Drawing.Color]::FromArgb(200, 90, 0)
    $warn.Text = "警告：另有 $($unresolved.Count) 条仍为『待确认』未填好，将保留原样上传（可能报错误）。"
    $form.Controls.Add($warn)
}

# 确认按钮
$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Location = New-Object System.Drawing.Point(360, $btnTop)
$btnOk.Size = New-Object System.Drawing.Size(150, 38)
$btnOk.Text = "确认回填并上传"
$btnOk.Font = New-Object System.Drawing.Font("Microsoft YaHei", 10, [System.Drawing.FontStyle]::Bold)
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($btnOk)

# 取消按钮
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Location = New-Object System.Drawing.Point(520, $btnTop)
$btnCancel.Size = New-Object System.Drawing.Size(120, 38)
$btnCancel.Text = "取消"
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($btnCancel)

$form.AcceptButton = $btnOk
$form.CancelButton = $btnCancel

$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    exit 0
} else {
    exit 1
}
