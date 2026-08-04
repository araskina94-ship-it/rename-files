# Rename Tool - Batch Rename Folders and Files
# Works without AI and tokens

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide PowerShell window
$ShowWindow = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "ShowWindow" -Namespace Win32 -PassThru

if ($Host.Name -ne "Windows PowerShell ISE") {
    $null = $ShowWindow::ShowWindowAsync((Get-Process -Id $PID).MainWindowHandle, 0)
}

$script:csvPath = ""
$script:folderPath = ""
$script:renameMode = "barcode"  # "barcode", "sequential", or "gather"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Rename Tool"
$form.Size = New-Object System.Drawing.Size(500, 475)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Batch Rename Folders and Files"
$titleLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(50, 20)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
$titleLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($titleLabel)

# Mode Selection
$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = "Rename Mode:"
$modeLabel.Location = New-Object System.Drawing.Point(50, 60)
$modeLabel.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($modeLabel)

$modeComboBox = New-Object System.Windows.Forms.ComboBox
$modeComboBox.Location = New-Object System.Drawing.Point(170, 58)
$modeComboBox.Size = New-Object System.Drawing.Size(280, 25)
$modeComboBox.DropDownStyle = "DropDownList"
$modeComboBox.Items.Add("Barcode (Stokmann)") | Out-Null
$modeComboBox.Items.Add("Sequential Numbers") | Out-Null
$modeComboBox.Items.Add("Gather Folders") | Out-Null
$modeComboBox.SelectedIndex = 0
$modeComboBox.Add_SelectedIndexChanged({
    $script:renameMode = switch ($modeComboBox.SelectedIndex) {
        0 { "barcode" }
        1 { "sequential" }
        2 { "gather" }
    }
})
$form.Controls.Add($modeComboBox)

# CSV Selection
$csvLabel = New-Object System.Windows.Forms.Label
$csvLabel.Text = "Select Excel/CSV file (Articul | Barcode)"
$csvLabel.Location = New-Object System.Drawing.Point(50, 95)
$csvLabel.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($csvLabel)

$csvButton = New-Object System.Windows.Forms.Button
$csvButton.Text = "Select File"
$csvButton.Location = New-Object System.Drawing.Point(50, 120)
$csvButton.Size = New-Object System.Drawing.Size(120, 30)
$csvButton.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($ofd.ShowDialog() -eq "OK") {
        $script:csvPath = $ofd.FileName
        $csvStatus.Text = "+ " + [System.IO.Path]::GetFileName($script:csvPath)
        $csvStatus.ForeColor = [System.Drawing.Color]::Green
    }
})
$form.Controls.Add($csvButton)

$csvStatus = New-Object System.Windows.Forms.Label
$csvStatus.Text = "(not selected)"
$csvStatus.Location = New-Object System.Drawing.Point(180, 125)
$csvStatus.Size = New-Object System.Drawing.Size(270, 20)
$csvStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($csvStatus)

# Template Download
$templateLabel = New-Object System.Windows.Forms.Label
$templateLabel.Text = "Download CSV Template"
$templateLabel.Location = New-Object System.Drawing.Point(50, 165)
$templateLabel.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($templateLabel)

$templateButton = New-Object System.Windows.Forms.Button
$templateButton.Text = "Download Template"
$templateButton.Location = New-Object System.Drawing.Point(50, 190)
$templateButton.Size = New-Object System.Drawing.Size(120, 30)
$templateButton.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Excel files (*.xlsx)|*.xlsx"
    $sfd.FileName = "template-articul-barcode.xlsx"
    if ($sfd.ShowDialog() -eq "OK") {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add(1)
            $sheet = $workbook.Worksheets.Item(1)

            # Headers
            $sheet.Cells.Item(1, 1).Value = "Articul"
            $sheet.Cells.Item(1, 2).Value = "Barcode"

            # Sample data
            $sheet.Cells.Item(2, 1).Value = "ART001"
            $sheet.Cells.Item(2, 2).Value = "4600000000001"
            $sheet.Cells.Item(3, 1).Value = "ART002"
            $sheet.Cells.Item(3, 2).Value = "4600000000002"
            $sheet.Cells.Item(4, 1).Value = "ART003"
            $sheet.Cells.Item(4, 2).Value = "0123456789012"

            # Set column B as TEXT format (preserves leading zeros)
            $columnB = $sheet.Columns.Item(2)
            $columnB.NumberFormat = "@"

            $workbook.SaveAs($sfd.FileName, 51)  # 51 = xlsx format
            $workbook.Close()
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            [System.Windows.Forms.MessageBox]::Show("Template saved!`n`n" + $sfd.FileName + "`n`nIMPORTANT: Column B is TEXT format - leading zeros are preserved!", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating Excel file!`n`n" + $_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})
$form.Controls.Add($templateButton)

# Folder Selection
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Select Folder with Items"
$folderLabel.Location = New-Object System.Drawing.Point(50, 240)
$folderLabel.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($folderLabel)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = "Select Folder"
$folderButton.Location = New-Object System.Drawing.Point(50, 265)
$folderButton.Size = New-Object System.Drawing.Size(120, 30)
$folderButton.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select folder with items"
    if ($fbd.ShowDialog() -eq "OK") {
        $script:folderPath = $fbd.SelectedPath
        $folderStatus.Text = "+ " + [System.IO.Path]::GetFileName($script:folderPath)
        $folderStatus.ForeColor = [System.Drawing.Color]::Green
    }
})
$form.Controls.Add($folderButton)

$folderStatus = New-Object System.Windows.Forms.Label
$folderStatus.Text = "(not selected)"
$folderStatus.Location = New-Object System.Drawing.Point(180, 270)
$folderStatus.Size = New-Object System.Drawing.Size(270, 20)
$folderStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($folderStatus)

# Backup Checkbox
$backupCheckbox = New-Object System.Windows.Forms.CheckBox
$backupCheckbox.Text = "Create Backup Copy"
$backupCheckbox.Location = New-Object System.Drawing.Point(50, 315)
$backupCheckbox.Size = New-Object System.Drawing.Size(200, 25)
$backupCheckbox.Checked = $true
$form.Controls.Add($backupCheckbox)

# Run Button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "START RENAMING"
$runButton.Location = New-Object System.Drawing.Point(50, 355)
$runButton.Size = New-Object System.Drawing.Size(400, 45)
$runButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$runButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$runButton.ForeColor = [System.Drawing.Color]::White
$runButton.FlatStyle = "Flat"
$runButton.Add_Click({
    # Check folder is always required
    if ([string]::IsNullOrEmpty($script:folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Select folder first!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    # CSV only required for barcode mode
    if (($script:renameMode -eq "barcode") -and [string]::IsNullOrEmpty($script:csvPath)) {
        [System.Windows.Forms.MessageBox]::Show("Select CSV file first!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $csvMsg = if ($script:csvPath) { "CSV: $($script:csvPath)`n`n" } else { "" }
    $res = [System.Windows.Forms.MessageBox]::Show("Start renaming?`n`nMode: $($script:renameMode)`n${csvMsg}Folder: $($script:folderPath)", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($res -ne "Yes") { return }

    if ($backupCheckbox.Checked) {
        $backupPath = $script:folderPath + "_backup_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
        Copy-Item -Path $script:folderPath -Destination $backupPath -Recurse
    }

    # SEQUENTIAL MODE - Rename files inside folders as 1, 2, 3...
    if ($script:renameMode -eq "sequential") {
        # Suffix priority function
        function Get-SuffixPriority {
            param($suffix)
            switch ($suffix) {
                ""      { return 1 }  # No suffix = highest priority
                "_E"    { return 2 }
                "_Q"    { return 3 }
                default { return 999 }  # Unknown suffixes = lowest
            }
        }

        $renamedFiles = 0
        $errors = 0
        $unknownSuffixes = @{}
        $totalFiles = 0

        # Get all folders
        $folders = Get-ChildItem -Path $script:folderPath -Directory

        foreach ($folder in $folders) {
            # Get files in this folder
            $files = Get-ChildItem -Path $folder.FullName -File
            if ($files.Count -eq 0) { continue }

            # Group and sort by suffix
            $fileGroups = @{}
            foreach ($file in $files) {
                $totalFiles++
                $baseName = $file.BaseName

                # Detect suffix
                $suffix = ""
                if ($baseName -match "_E$") { $suffix = "_E" }
                elseif ($baseName -match "_Q$") { $suffix = "_Q" }

                # Track unknown suffixes
                $priority = Get-SuffixPriority -suffix $suffix
                if ($priority -eq 999) {
                    if (-not $unknownSuffixes.ContainsKey($suffix)) {
                        $unknownSuffixes[$suffix] = 0
                    }
                    $unknownSuffixes[$suffix]++
                }

                if (-not $fileGroups.ContainsKey($suffix)) {
                    $fileGroups[$suffix] = @()
                }
                $fileGroups[$suffix] += $file
            }

            # Sort groups by priority, then rename sequentially
            $sortedSuffixes = $fileGroups.Keys | Sort-Object { Get-SuffixPriority -suffix $_ }

            $fileNumber = 1
            foreach ($suffix in $sortedSuffixes) {
                $groupFiles = $fileGroups[$suffix]
                foreach ($file in $groupFiles) {
                    $newName = "$fileNumber$($file.Extension)"
                    try {
                        Rename-Item -Path $file.FullName -NewName $newName -ErrorAction Stop
                        $renamedFiles++
                        $fileNumber++
                    } catch {
                        $errors++
                    }
                }
            }
        }

        # Show warning for unknown suffixes
        $warningMsg = ""
        if ($unknownSuffixes.Count -gt 0) {
            $warningMsg = "`n`n⚠ Unknown suffixes found:`n"
            foreach ($suffix in $unknownSuffixes.Keys) {
                $warningMsg += "  $suffix : $($unknownSuffixes[$suffix]) file(s)`n"
            }
            $warningMsg += "These files were renamed with lowest priority."
        }

        $backupMsg = if ($backupCheckbox.Checked) { "`nBackup: $backupPath" } else { "" }
        [System.Windows.Forms.MessageBox]::Show("Done!`n`nFiles renamed: $renamedFiles`nErrors: $errors$backupMsg$warningMsg", "Sequential Rename Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # GATHER MODE - Group files into folders based on base name
    if ($script:renameMode -eq "gather") {
        # Get base name (everything before first underscore)
        function Get-BaseName {
            param($fileName)
            $underscorePos = $fileName.IndexOf("_")
            if ($underscorePos -gt 0) {
                return $fileName.Substring(0, $underscorePos)
            } else {
                return $fileName
            }
        }

        $movedFiles = 0
        $errors = 0
        $createdFolders = 0

        # Get all files in the selected folder (not recursive)
        $files = Get-ChildItem -Path $script:folderPath -File

        # Group files by base name
        $fileGroups = @{}
        foreach ($file in $files) {
            $baseName = Get-BaseName -fileName $file.BaseName
            if (-not $fileGroups.ContainsKey($baseName)) {
                $fileGroups[$baseName] = @()
            }
            $fileGroups[$baseName] += $file
        }

        # Create folders and move files
        foreach ($baseName in $fileGroups.Keys) {
            $folderPath = Join-Path -Path $script:folderPath -ChildPath $baseName

            # Skip if folder already exists (avoid moving files into existing folder)
            if (Test-Path -Path $folderPath) {
                [System.Windows.Forms.MessageBox]::Show("Folder '$baseName' already exists!`n`nSkipping this group.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                continue
            }

            # Create folder
            try {
                New-Item -Path $folderPath -ItemType Directory -ErrorAction Stop | Out-Null
                $createdFolders++
            } catch {
                $errors++
                continue
            }

            # Move files to folder
            foreach ($file in $fileGroups[$baseName]) {
                $destPath = Join-Path -Path $folderPath -ChildPath $file.Name
                try {
                    Move-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
                    $movedFiles++
                } catch {
                    $errors++
                }
            }
        }

        $backupMsg = if ($backupCheckbox.Checked) { "`nBackup: $backupPath" } else { "" }
        [System.Windows.Forms.MessageBox]::Show("Done!`n`nFolders created: $createdFolders`nFiles moved: $movedFiles`nErrors: $errors$backupMsg", "Gather Folders Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    # BARCODE MODE - Original logic

    # Read data from file (Excel or CSV)
    try {
        $ext = [System.IO.Path]::GetExtension($script:csvPath).ToLower()
        $csvData = @()

        if ($ext -eq ".xlsx" -or $ext -eq ".xls") {
            # Read Excel file
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($script:csvPath)
            $sheet = $workbook.Worksheets.Item(1)

            $rowCount = 1
            $lastRow = $sheet.UsedRange.Rows.Count

            for ($i = 2; $i -le $lastRow; $i++) {
                $articul = $sheet.Cells.Item($i, 1).Text
                $barcode = $sheet.Cells.Item($i, 2).Text
                if (-not [string]::IsNullOrWhiteSpace($articul) -and -not [string]::IsNullOrWhiteSpace($barcode)) {
                    $csvData += [PSCustomObject]@{Articul = $articul; Barcode = $barcode}
                }
            }

            $workbook.Close($false)
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        } else {
            # Read CSV file
            $csvData = Import-Csv -Path $script:csvPath -Encoding UTF8 -Delimiter ';'
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("File read error!`n`n$_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $renamedFolders = 0
    $renamedFiles = 0
    $errors = 0

    foreach ($row in $csvData) {
        $articul = $row.Articul
        $barcode = $row.Barcode
        if ([string]::IsNullOrEmpty($articul) -or [string]::IsNullOrEmpty($barcode)) { continue }

        $oldPath = Join-Path -Path $script:folderPath -ChildPath $articul
        if (Test-Path -Path $oldPath) {
            $newPath = Join-Path -Path $script:folderPath -ChildPath $barcode
            try {
                Rename-Item -Path $oldPath -NewName $barcode -ErrorAction Stop
                $renamedFolders++
                $files = Get-ChildItem -Path $newPath -File
                $counter = 1
                foreach ($file in $files) {
                    $newFileName = "$barcode-$counter$($file.Extension)"
                    try {
                        Rename-Item -Path $file.FullName -NewName $newFileName -ErrorAction Stop
                        $renamedFiles++
                        $counter++
                    } catch { $errors++ }
                }
            } catch { $errors++ }
        }
    }

    $backupMsg = if ($backupCheckbox.Checked) { "`nBackup: $backupPath" } else { "" }
    [System.Windows.Forms.MessageBox]::Show("Done!`n`nFolders: $renamedFolders`nFiles: $renamedFiles`nErrors: $errors$backupMsg", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})
$form.Controls.Add($runButton)

$form.ShowDialog() | Out-Null
