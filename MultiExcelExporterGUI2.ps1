### ✅ 반드시 STA 모드에서 실행 ###
if (-not ([System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA')) {
    Write-Host "⚠ The script is re-run in STA mode for GUI execution..."
    powershell.exe -STA -File $PSCommandPath
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

### ✅ -------------------------
### ✅ 폴더 선택 GUI
### ✅ -------------------------
$sourceFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$sourceFolderDialog.Description = " Select the folder containing the original Excel files"

if ($sourceFolderDialog.ShowDialog() -ne "OK") {
    Write-Host "Exiting because no folder was selected."
    exit
}
$SourceFolder = $sourceFolderDialog.SelectedPath

$excelFiles = Get-ChildItem -Path $SourceFolder -Filter *.xlsx

if ($excelFiles.Count -eq 0) {
    Write-Host "There is no Excel file in the folder."
    exit
}

$targetFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$targetFolderDialog.Description = "☞ Select folder to save copy"

if ($targetFolderDialog.ShowDialog() -ne "OK") {
    Write-Host "Exiting because no save folder was selected."
    exit
}
$TargetFolder = $targetFolderDialog.SelectedPath

if (-not (Test-Path $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder | Out-Null
}

### ✅ 총 파일 개수
$total = $excelFiles.Count
$count = 0

### ✅ -------------------------
### ✅ ProgressBar UI 생성
### ✅ -------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SingleFileExporter - Progress"
$form.Size = New-Object System.Drawing.Size(500,180)
$form.StartPosition = "CenterScreen"

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,60)
$progressBar.Size = New-Object System.Drawing.Size(440,30)
$progressBar.Minimum = 0
$progressBar.Maximum = $total
$form.Controls.Add($progressBar)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20,20)
$label.Size = New-Object System.Drawing.Size(440,30)
$label.Text = "Preparing for processing..."
$form.Controls.Add($label)

$percentLabel = New-Object System.Windows.Forms.Label
$percentLabel.Location = New-Object System.Drawing.Point(20,100)
$percentLabel.Size = New-Object System.Drawing.Size(440,30)
$percentLabel.Text = "0% complete"
$form.Controls.Add($percentLabel)

$form.Show()
[System.Windows.Forms.Application]::DoEvents()

### ✅ 엑셀 COM 준비
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

### ✅ -------------------------
### ✅ 파일 반복 처리
### ✅ -------------------------
foreach ($file in $excelFiles) {

    $count++
    $label.Text = "Processing: $($file.Name)"
    $progressBar.Value = $count
    $percent = [math]::Round(($count / $total) * 100)
    $percentLabel.Text = "$percent% complete"
    [System.Windows.Forms.Application]::DoEvents()

    $SourcePath = $file.FullName
    $TargetPath = Join-Path $TargetFolder ("PSC_" + $file.Name)

    $sourceWorkbook = $null
    $newWorkbook = $null

    try {
        $sourceWorkbook = $excel.Workbooks.Open($SourcePath)
        $newWorkbook = $excel.Workbooks.Add()

        while ($newWorkbook.Sheets.Count -gt 1) {
            $newWorkbook.Sheets.Item(1).Delete()
        }

        $sourceWorkbook.Sheets.Item(1).Copy($newWorkbook.Sheets.Item(1))

        for ($i = 2; $i -le $sourceWorkbook.Sheets.Count; $i++) {
            $sh = $sourceWorkbook.Sheets.Item($i)
            $sh.Copy([System.Reflection.Missing]::Value,
                     $newWorkbook.Sheets.Item($newWorkbook.Sheets.Count))
        }

        $newWorkbook.Sheets.Item(1).Delete()
        $newWorkbook.SaveAs($TargetPath)
    }
    catch {
        Write-Host "error: $($_.Exception.Message)"
    }
    finally {
        if ($sourceWorkbook) { $sourceWorkbook.Close($false) }
        if ($newWorkbook)    { $newWorkbook.Close($true) }
    }
}

$excel.Quit()

$label.Text = "All files have been processed!"
$percentLabel.Text = "100%"
$progressBar.Value = $total
[System.Windows.Forms.Application]::DoEvents()

Start-Sleep -Seconds 2
$form.Close()

Write-Host "job done!"
Pause
