# Populate the 'Employee Username' column in a spreadsheet
# Loop over the list of employees and using the 'Display Name' column
# Query Active Directory to grab the user's SamAccountName

$excel = New-Object -ComObject excel.application
$excel.Visible = $true
$excel.DisplayAlerts = $true

$wb = $excel.Workbooks.Open('C:\employees.xlsx')

foreach ($worksheet in $wb.Sheets){
    
    $intRow = 2

    Do {
        $employee_name = $worksheet.Cells.Item($intRow, 1).Value()
        $user          = Get-ADUser -Filter {displayName -like $employee_name}
        
        if ($user) {
            $username = $user.SamAccountName
            $worksheet.Cells.Item($intRow, 7) = $username
        }

        $intRow++
    } 

    While ($worksheet.Cells.Item($intRow,1).Value() -ne $null)

}

$excel.Workbooks.Close($true)
$excel.Quit()