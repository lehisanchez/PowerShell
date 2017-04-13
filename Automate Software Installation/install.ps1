# CHECK SYSTEM ARCHITECTURE (32-BIT / 64-BIT)
# ===========================================
# ===========================================
$os_type = (Get-WmiObject -Class Win32_ComputerSystem).SystemType -match '(x64)'



# COMMAND LINE MENU
# ===========================================
# ===========================================
function Show-Menu {
    
    param (
        [string]$Title = 'Laptop Update Options'
    )

    cls

    Write-Host "===================================================="
    Write-Host "================ $Title ================"
    Write-Host "===================================================="
    Write-Host "Press '1' to register Acme Business Files."
    Write-Host "Press '2' to install Web Browser."
    Write-Host "Press '3' to install Printer Files."

    # CHECK COMPUTER NAME. ANY COMPUTER STARTING WITH 'LAP9*' OR 'LAP10*' CAN BE SHOWN THE FOLLOWING OPTIONS
    # ==============================================================================================
    if ($env:computerName.contains("LAP10") -or $env:computerName.contains("LAP9")) {
        Write-Host "Press '5' to uninstall Acme Anti-Virus."
        Write-Host "Press '4' to install Acme Anti-Virus."
    }
    Write-Host "Press 'Q' to quit."
}

do {
    
    Show-Menu

    $input = Read-Host "`nPlease make a selection"
     
    switch ($input) {
        

        # REGISTER ACME BUSINESS FILES
        # ==============================================================================================
        # ==============================================================================================
        '1' {

            cls

            Write-Host "Registering Acme Business files..."
            
            # CHECK ARCHITECTURE
            # ==================
            if ($os_type -eq "True") {
                C:\Windows\sysWOW64\regsvr32.exe "C:\Program Files (x86)\Acme Business\Bin\file_001.ocx" | Out-Null
                C:\Windows\sysWOW64\regsvr32.exe "C:\Program Files (x86)\Acme Business\Bin\file_002.ocx" | Out-Null
                C:\Windows\sysWOW64\regsvr32.exe "C:\Program Files (x86)\Acme Business\Bin\file_003.dll" | Out-Null
                C:\Windows\sysWOW64\regsvr32.exe "C:\Program Files (x86)\Acme Business\Bin\file_004.ocx" | Out-Null
            } else {
                regsvr32 "C:\Program Files\Acme Business\Bin\file_001.ocx" | Out-Null
                regsvr32 "C:\Program Files\Acme Business\Bin\file_002.ocx" | Out-Null
                regsvr32 "C:\Program Files\Acme Business\Bin\file_003.dll" | Out-Null
                regsvr32 "C:\Program Files\Acme Business\Bin\file_004.ocx" | Out-Null
            }

            Write-Host "Done!`n"

        } 
        

        # INSTALL WEB BROWSER
        # ==============================================================================================
        # ==============================================================================================
        '2' {
            
            cls

            Write-Host "Installing Web Browser..."
            
            # CHECK ARCHITECTURE
            # ==================
            if ($os_type -eq "True") {
                Start-Process msiexec.exe -Wait -ArgumentList '/I \\network_folder\Web_Browser_x64.msi /quiet'
            } else {
                Start-Process msiexec.exe -Wait -ArgumentList '/I \\network_folder\Web_Browser_x86.msi /quiet'
            }

            Write-Host "Done!`n"

        } 


        # INSTALL PRINTER FILES
        # ==============================================================================================
        # ==============================================================================================
        '3' {
            
            cls
            
            Write-Host "Installing Printer Files..."

            # CHECK ARCHITECTURE
            # ==================
            if ($os_type -eq "True") {
                Start-Process -Wait \\network_folder\Laptop_x64_04102017.printerExport
            } else {
                Start-Process -Wait \\network_folder\Laptop_x32_04102017.printerExport
            }

            Write-Host "Done!`n"

        } 


        # UNINSTALL ACME ANTI-VIRUS & RESTART COMPUTER
        # ==============================================================================================
        # ==============================================================================================
        '4' {
            
            cls
            
            Write-Host "Looking for Acme Anti-Virus...`n"

            $app = Get-WmiObject -Class Win32_Product -Filter "Name = 'Acme Anti-Virus Agent'"

            # IF ACME ANTI-VIRUS IS INSTALLED ON THE MACHINE
            # ==============================================
            if ($app) {
                
                cls
                
                Write-Host "Uninstalling Acme Anti-Virus..."

                $app.Uninstall()

                Write-Host "Done!`n"

                $restart_computer = Read-Host -Prompt "Would you like to restart the computer?(Y/N)"

                if ($restart_computer = "Y") {

                    Restart-Computer -Force

                    {break}

                } else {
                    Write-Host "Acme Anti-Virus was just uninstalled. Please restart the computer before reinstalling Acme Anti-Virus.`n"
                }

            } else {
                Write-Host "Acme Anti-Virus was not found on this computer. Please install Acme Anti-Virus."
            }
            
        } 


        # INSTALL ACME ANTI-VIRUS & RESTART COMPUTER
        # ==============================================================================================
        # ==============================================================================================
        '5' {
            
            cls
            
            Write-Host "Installing Acme Anti-Virus..."

            Start-Process msiexec.exe -Wait -ArgumentList '/I \\network_folder\Acme_Laptop.msi /quiet'
            
            Write-Host "Done!`n"

        } 
        
        'q' {
            return
        }
    }
    
    pause

}

until ($input -eq 'q')