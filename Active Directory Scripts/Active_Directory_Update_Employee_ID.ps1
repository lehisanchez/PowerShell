# Update the employee's employeeID attribute in Active Directory
# Using the 'Employee Username' column

$excel = New-Object -ComObject excel.application
$excel.Visible = $true
$excel.DisplayAlerts = $true

$wb = $excel.Workbooks.Open('C:\employees.xlsx')

foreach ($worksheet in $wb.Sheets){
    
    $intRow = 2

    Do {
        $employee_name = $worksheet.Cells.Item($intRow, 8).Value()
        $employee_id   = $worksheet.Cells.Item($intRow, 4).Value()

        if ($employee_name -ne $null){
            Set-ADUser -Identity $employee_name -Replace @{employeeID=$employee_id}
            # Write-Host "$employee_name, $employee_id"
        }

        $intRow++
    } 

    While ($worksheet.Cells.Item($intRow,1).Value() -ne $null)

}

$excel.Workbooks.Close()
$excel.Quit()