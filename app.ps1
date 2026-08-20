param([switch]$SelfTest)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:AppDir = if ($env:RENAMER_APP_DIR) {
    $env:RENAMER_APP_DIR.TrimEnd('\')
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}
$script:Bridge = Join-Path $script:AppDir 'bridge.py'
$pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    [System.Windows.Forms.MessageBox]::Show(
        '未找到 Python。请先安装 Python 3，并勾选 Add Python to PATH。',
        '无法启动',
        'OK',
        'Error'
    ) | Out-Null
    exit 1
}
$script:PythonExe = $pythonCommand.Source

function Invoke-RenamerBridge {
    param([string]$Action)

    $resultFile = [System.IO.Path]::GetTempFileName()
    try {
        $arguments = @(
            $script:Bridge,
            $Action,
            '--output', $resultFile,
            '--folder', $folderText.Text,
            '--roster', $rosterText.Text,
            '--project', $projectText.Text,
            '--template', $templateText.Text
        )
        & $script:PythonExe @arguments 2>$null
        if (-not (Test-Path -LiteralPath $resultFile)) {
            throw '后台程序未生成运行结果。'
        }
        $content = [System.IO.File]::ReadAllText($resultFile, [System.Text.Encoding]::UTF8)
        return $content | ConvertFrom-Json
    }
    finally {
        if (Test-Path -LiteralPath $resultFile) {
            Remove-Item -LiteralPath $resultFile -Force
        }
    }
}

function Show-ErrorMessage {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, '操作失败', 'OK', 'Error') | Out-Null
}

function Update-PreviewGrid {
    param($Result)
    $grid.Rows.Clear()
    $labels = @{ ready = '待改名'; skipped = '跳过'; conflict = '冲突' }
    foreach ($item in $Result.items) {
        $sourceName = [System.IO.Path]::GetFileName([string]$item.source)
        $targetName = [System.IO.Path]::GetFileName([string]$item.target)
        $rowIndex = $grid.Rows.Add($labels[[string]$item.status], $sourceName, $targetName, [string]$item.reason)
        $row = $grid.Rows[$rowIndex]
        if ($item.status -eq 'conflict') {
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
        }
        elseif ($item.status -eq 'ready') {
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::ForestGreen
        }
        else {
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
        }
    }
    $statusLabel.Text = "待改名 $($Result.summary.ready) 个，跳过 $($Result.summary.skipped) 个，冲突 $($Result.summary.conflict) 个。"
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '花名册批量文件重命名助手'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1040, 700)
$form.MinimumSize = New-Object System.Drawing.Size(850, 560)
$form.Font = New-Object System.Drawing.Font('微软雅黑', 9)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.Padding = New-Object System.Windows.Forms.Padding(14)
$layout.ColumnCount = 3
$layout.RowCount = 7
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 125)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 155)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 55)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 35)))
$form.Controls.Add($layout)

function Add-FieldLabel([string]$text, [int]$row) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleLeft'
    $layout.Controls.Add($label, 0, $row)
}

Add-FieldLabel '待改名文件夹：' 0
$folderText = New-Object System.Windows.Forms.TextBox
$folderText.Dock = 'Fill'
$folderText.Margin = New-Object System.Windows.Forms.Padding(4, 7, 8, 5)
$layout.Controls.Add($folderText, 1, 0)
$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = '选择文件夹'
$folderButton.Dock = 'Fill'
$layout.Controls.Add($folderButton, 2, 0)
$folderButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '选择待重命名文件所在的文件夹'
    if ($dialog.ShowDialog() -eq 'OK') { $folderText.Text = $dialog.SelectedPath }
    $dialog.Dispose()
})

Add-FieldLabel '花名册：' 1
$rosterText = New-Object System.Windows.Forms.TextBox
$rosterText.Dock = 'Fill'
$rosterText.Margin = New-Object System.Windows.Forms.Padding(4, 7, 8, 5)
$layout.Controls.Add($rosterText, 1, 1)
$rosterButton = New-Object System.Windows.Forms.Button
$rosterButton.Text = '选择花名册'
$rosterButton.Dock = 'Fill'
$layout.Controls.Add($rosterButton, 2, 1)
$rosterButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = '选择花名册'
    $dialog.Filter = 'Excel/CSV 花名册|*.xlsx;*.csv;*.tsv;*.txt|所有文件|*.*'
    if ($dialog.ShowDialog() -eq 'OK') { $rosterText.Text = $dialog.FileName }
    $dialog.Dispose()
})

Add-FieldLabel '统一项目名：' 2
$projectText = New-Object System.Windows.Forms.TextBox
$projectText.Dock = 'Fill'
$projectText.Margin = New-Object System.Windows.Forms.Padding(4, 7, 8, 5)
$layout.Controls.Add($projectText, 1, 2)
$projectHint = New-Object System.Windows.Forms.Label
$projectHint.Text = '留空则读花名册项目列'
$projectHint.Dock = 'Fill'
$projectHint.TextAlign = 'MiddleLeft'
$layout.Controls.Add($projectHint, 2, 2)

Add-FieldLabel '命名模板：' 3
$templateText = New-Object System.Windows.Forms.TextBox
$templateText.Text = '{姓名}_{学号}_{项目}'
$templateText.Dock = 'Fill'
$templateText.Margin = New-Object System.Windows.Forms.Padding(4, 7, 8, 5)
$layout.Controls.Add($templateText, 1, 3)
$templateHint = New-Object System.Windows.Forms.Label
$templateHint.Text = '{姓名}  {学号}  {项目}'
$templateHint.Dock = 'Fill'
$templateHint.TextAlign = 'MiddleLeft'
$layout.Controls.Add($templateHint, 2, 3)

$actions = New-Object System.Windows.Forms.FlowLayoutPanel
$actions.Dock = 'Fill'
$actions.FlowDirection = 'LeftToRight'
$layout.SetColumnSpan($actions, 3)
$layout.Controls.Add($actions, 0, 4)
$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Text = '1. 预览'
$previewButton.Size = New-Object System.Drawing.Size(110, 35)
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Text = '2. 确认执行'
$executeButton.Size = New-Object System.Drawing.Size(125, 35)
$undoButton = New-Object System.Windows.Forms.Button
$undoButton.Text = '撤销最近一次'
$undoButton.Size = New-Object System.Drawing.Size(140, 35)
$actions.Controls.AddRange(@($previewButton, $executeButton, $undoButton))

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
[void]$grid.Columns.Add('状态', '状态')
[void]$grid.Columns.Add('原文件名', '原文件名')
[void]$grid.Columns.Add('新文件名', '新文件名')
[void]$grid.Columns.Add('说明', '说明')
$grid.Columns[0].FillWeight = 15
$grid.Columns[1].FillWeight = 35
$grid.Columns[2].FillWeight = 45
$grid.Columns[3].FillWeight = 35
$layout.SetColumnSpan($grid, 3)
$layout.Controls.Add($grid, 0, 5)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = '请选择文件夹和花名册，然后点击预览。'
$statusLabel.Dock = 'Fill'
$statusLabel.TextAlign = 'MiddleLeft'
$layout.SetColumnSpan($statusLabel, 3)
$layout.Controls.Add($statusLabel, 0, 6)

$previewButton.Add_Click({
    try {
        $result = Invoke-RenamerBridge 'preview'
        if (-not $result.ok) { Show-ErrorMessage ([string]$result.error); return }
        Update-PreviewGrid $result
    }
    catch { Show-ErrorMessage $_.Exception.Message }
})

$executeButton.Add_Click({
    try {
        $preview = Invoke-RenamerBridge 'preview'
        if (-not $preview.ok) { Show-ErrorMessage ([string]$preview.error); return }
        Update-PreviewGrid $preview
        if ([int]$preview.summary.conflict -gt 0) {
            Show-ErrorMessage '存在冲突，请根据表格说明修正后重新预览。'
            return
        }
        if ([int]$preview.summary.ready -eq 0) {
            Show-ErrorMessage '没有可以重命名的文件。'
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "将重命名 $($preview.summary.ready) 个文件。`r`n`r`n是否继续？",
            '确认重命名',
            'YesNo',
            'Question'
        )
        if ($answer -ne 'Yes') { return }
        $result = Invoke-RenamerBridge 'execute'
        if (-not $result.ok) { Show-ErrorMessage ([string]$result.error); return }
        [System.Windows.Forms.MessageBox]::Show(
            "已成功重命名 $($result.summary.ready) 个文件。",
            '完成', 'OK', 'Information'
        ) | Out-Null
        $refresh = Invoke-RenamerBridge 'preview'
        if ($refresh.ok) { Update-PreviewGrid $refresh }
    }
    catch { Show-ErrorMessage $_.Exception.Message }
})

$undoButton.Add_Click({
    try {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            '将撤销该文件夹最近一次重命名。是否继续？',
            '确认撤销', 'YesNo', 'Question'
        )
        if ($answer -ne 'Yes') { return }
        $result = Invoke-RenamerBridge 'undo'
        if (-not $result.ok) { Show-ErrorMessage ([string]$result.error); return }
        [System.Windows.Forms.MessageBox]::Show(
            "已恢复 $($result.count) 个文件的原名。",
            '撤销完成', 'OK', 'Information'
        ) | Out-Null
        if ($rosterText.Text) {
            $refresh = Invoke-RenamerBridge 'preview'
            if ($refresh.ok) { Update-PreviewGrid $refresh }
        }
    }
    catch { Show-ErrorMessage $_.Exception.Message }
})

if ($SelfTest) {
    $form.Dispose()
    Write-Output 'PowerShell GUI create OK'
    exit 0
}

[void]$form.ShowDialog()
$form.Dispose()
