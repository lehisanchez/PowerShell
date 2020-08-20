# Poulate Spreadsheet with employee USERNAME and EMAIL from Active Directory

$excel = New-Object -ComObject excel.application
$excel.Visible = $true
$excel.DisplayAlerts = $true

$wb = $excel.Workbooks.Open('C:\users.xls')

foreach ($worksheet in $wb.Sheets){
    
    $intRow = 3

    Do {
        if ($worksheet.Cells.Item($intRow, 3).Value() -ne $null) {

            $employee_id = $worksheet.Cells.Item($intRow, 3).Value()

            $user = Get-ADUser -Filter {employeeID -eq $employee_id} -Properties mail

            if ($user -ne $null){
                $worksheet.Cells.Item($intRow, 5) = $user.SamAccountName
                $worksheet.Cells.Item($intRow, 6) = $user.Mail
            }
        }

        $intRow++
    } 

    While ($worksheet.Cells.Item($intRow,1).Value() -ne $null)

}

$excel.Workbooks.Close()
$excel.Quit()