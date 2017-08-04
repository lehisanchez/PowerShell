# Populate the 'Supervisor Username' and 'Supervisor Email' columns in a spreadsheet
# Loop over the list of supervisors and using the 'Supervisor ID'
# Query Active Directory using the User.employeeID attribute

$excel = New-Object -ComObject excel.application
$excel.Visible = $true
$excel.DisplayAlerts = $true

$wb = $excel.Workbooks.Open('C:\supervisors.xls')

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