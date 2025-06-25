
$filePath = "\\ACT-tHGhawNuPBe\C$\SalesTimesheet\DO_NOT_OPEN\SalesTimesheets.xlsx"
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($filePath)

Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Sales Timesheet"
$form.Size = New-Object System.Drawing.Size(985, 300)
$form.StartPosition = "CenterScreen"

# Create the Explaination label
$explanationLabel = New-Object System.Windows.Forms.Label
$explanationLabel.Location = New-Object System.Drawing.Point(390, 20)
$explanationLabel.Size = New-Object System.Drawing.Size(480, 40)
$user = $env:USERNAME
if ($user -eq "aluna") {
    $name = "Alberto Luna"
}
$explanationLabel.Text = "Welcome "+ $name
$explanationLabel.Font = New-Object Drawing.Font('Arial', 14)
$form.Controls.Add($explanationLabel)

# Create buttons
$PunchInButton = New-Object System.Windows.Forms.Button
$PunchInButton.Location = New-Object System.Drawing.Point(250, 70)
$PunchInButton.Size = New-Object System.Drawing.Size(480, 30)
$PunchInButton.Text = "Punch In"
$form.Controls.Add($PunchInButton)

$PunchOutButton = New-Object System.Windows.Forms.Button
$PunchOutButton.Location = New-Object System.Drawing.Point(250, 120)
$PunchOutButton.Size = New-Object System.Drawing.Size(480, 30)
$PunchOutButton.Text = "Punch Out"
$form.Controls.Add($PunchOutButton)

$user = $env:USERNAME
if ($user -eq "aluna") {
    $userSheet = 1
}

$worksheet = $workbook.Sheets.Item($userSheet) #Change number to change sheet

$row = 4
$column = 2

$cellValue1 = $worksheet.Cells.Item($row,1).Value2 #date

    while ($cellValue1 -ne $currentDate) {
        $row++
        $cellValue1 = $worksheet.Cells.Item($row,1).Value2
    }
    if ($cellValue1 -eq $null) {
        $row--
    }

$PunchInButton.Add_Click({


    $currentDate = (Get-Date).ToString('ddd MM/dd/yyyy')
    $currentTime = (Get-Date).ToString('hh:mm:ss tt')

    $empty = ""

    $cellValue1 = $worksheet.Cells.Item($row,1).Value2 #date

    while ($cellValue1 -ne $currentDate) {
        $row++
        $cellValue1 = $worksheet.Cells.Item($row,1).Value2
        if ($cellValue1 -eq $null) {
            Write-Host("loop in if")
            $worksheet.Cells.Item($row,1).Value2 = $currentDate
            $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
  
            while ($cellValue2 -ne $null) {
                $column++
                $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
            }
            $worksheet.Cells.Item($row,$column).Value2 = $currentTime
            break
        }
    }

    if ($cellValue1 -eq $currentDate) {
        Write-Host("loop2")
        $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
  
        while ($cellValue2 -ne $null) {
            $column++
            $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
        }
        $worksheet.Cells.Item($row,$column).Value2 = $currentTime
    }

    $workbook.Save()

    $workbook.Close()
    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

    $response = [System.Windows.Forms.MessageBox]::Show("Sucessfully Punched in.", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK)

    if ($response -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "You clicked Ok."
        get-process powershell | stop-process -force
    }

})

$PunchOutButton.Add_Click({


    $currentDate = (Get-Date).ToString('ddd MM/dd/yyyy')
    $currentTime = (Get-Date).ToString('hh:mm:ss tt')

    $empty = ""

    $cellValue1 = $worksheet.Cells.Item($row,1).Value2 #date

    while ($cellValue1 -ne $currentDate) {
        $row++
        $cellValue1 = $worksheet.Cells.Item($row,1).Value2
        if ($cellValue1 -eq $null) {
            Write-Host("loop in if")
            $worksheet.Cells.Item($row,1).Value2 = $currentDate
            $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
  
            while ($cellValue2 -ne $null) {
                $column++
                $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
            }
            $worksheet.Cells.Item($row,$column).Value2 = $currentTime
            break
        }
    }

    if ($cellValue1 -eq $currentDate) {
        Write-Host("loop2")
        $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
  
        while ($cellValue2 -ne $null) {
            $column++
            $cellValue2 = $worksheet.Cells.Item($row,$column).Value2
        }
        $worksheet.Cells.Item($row,$column).Value2 = $currentTime
    }

    $workbook.Save()

    $workbook.Close()
    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)

    $response = [System.Windows.Forms.MessageBox]::Show("Sucessfully Punched Out.", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::OK)

    if ($response -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "You clicked Ok."
        get-process powershell | stop-process -force
    }

})

# Create a DataGridView to display a table
$dataGridView = New-Object Windows.Forms.DataGridView
$dataGridView.Size = New-Object Drawing.Size(945, 70)
$dataGridView.Location = New-Object Drawing.Point(10, 170)

$dataTable = New-Object System.Data.DataTable
$dataTable.Columns.Add("Date")
$dataTable.Columns.Add("In")
$dataTable.Columns.Add("Break 1 Out")
$dataTable.Columns.Add("Break 1 In")
$dataTable.Columns.Add("Lunch Out")
$dataTable.Columns.Add("Lunch In")
$dataTable.Columns.Add("Break 2 Out")
$dataTable.Columns.Add("Break 2 In")
$dataTable.Columns.Add("Out")

$cell2 = $worksheet.Cells.Item($row,2).Value2
$cell3 = $worksheet.Cells.Item($row,3).Value2
$cell4 = $worksheet.Cells.Item($row,4).Value2
$cell5 = $worksheet.Cells.Item($row,5).Value2
$cell6 = $worksheet.Cells.Item($row,6).Value2
$cell7 = $worksheet.Cells.Item($row,7).Value2
$cell8 = $worksheet.Cells.Item($row,8).Value2
$cell9 = $worksheet.Cells.Item($row,9).Value2

for ($i = 2; $i -le 9; $i++) {

    $decimalTime = $worksheet.Cells.Item($row,$i).Value2

    $totalHours = $decimalTime * 24

    $hours24 = [math]::Floor($totalHours)
    $minutes = [math]::Round(($totalHours - $hours24) * 60)

    if ($minutes -eq 1) {
        $minutes = "01"
    } elseif ($minutes -eq 2) {
        $minutes = "02"
    } elseif ($minutes -eq 3) {
        $minutes = "03"
    } elseif ($minutes -eq 4) {
        $minutes = "04"
    } elseif ($minutes -eq 5) {
        $minutes = "05"
    } elseif ($minutes -eq 6) {
        $minutes = "06"
    } elseif ($minutes -eq 7) {
        $minutes = "07"
    } elseif ($minutes -eq 8) {
        $minutes = "08"
    } elseif ($minutes -eq 9) {
        $minutes = "09"
    }

    $amPm = if ($hours24 -ge 12) { "PM" } else { "AM" }

    $hours12 = $hours24 % 12

    if ($hours12 -eq 0) {
        $hours12 = 12
    }

    $formattedTime = "$hours12 : $minutes $amPm"
    if ($minutes -eq 0) {
        $formattedTime = " "
    }
    if ($i -eq 2) {
        $cellTime2 = $formattedTime
    } elseif ($i -eq 3) {
        $cellTime3 = $formattedTime
    } elseif ($i -eq 4) {
        $cellTime4 = $formattedTime
    } elseif ($i -eq 5) {
        $cellTime5 = $formattedTime
    } elseif ($i -eq 6) {
        $cellTime6 = $formattedTime
    } elseif ($i -eq 7) {
        $cellTime7 = $formattedTime
    } elseif ($i -eq 8) {
        $cellTime8 = $formattedTime
    } elseif ($i -eq 9) {
        $cellTime9 = $formattedTime
    }
}

$row1 = $dataTable.NewRow()
$row1["Date"] = $worksheet.Cells.Item($row,1).Value2
$row1["In"] = $cellTime2
$row1["Break 1 Out"] = $cellTime3
$row1["Break 1 In"] = $cellTime4
$row1["Lunch Out"] = $cellTime5
$row1["Lunch In"] = $cellTime6
$row1["Break 2 Out"] = $cellTime7
$row1["Break 2 In"] = $cellTime8
$row1["Out"] = $cellTime9
$dataTable.Rows.Add($row1)

# Set the DataGridView's data source to the DataTable
$dataGridView.DataSource = $dataTable

# Add the DataGridView to the form
$form.Controls.Add($dataGridView)

# Define what to do when the form is closed
$Form_Closed = {
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
}
$form.Add_Closed($Form_Closed)

# Show the form
$form.ShowDialog() | Out-Null