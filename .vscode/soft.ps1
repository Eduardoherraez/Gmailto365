
function ShowMainMenu {
# main menu 
    $choice = $null
        do {
        Clear-Host
        checkeol #check if connected to exchange online
        showgam #check if gam is working
        Write-Host "Main Menu" 
        Write-Host "1. User tools"
        Write-Host "2. Group tools"
        Write-Host "3. Connect to Exchange online"
        Write-Host "4. Tool Settings"
        Write-Host "Q. Quit"
        $choice = Read-Host "Please select an option"
        switch ($choice) {
            '1' { ShowUserMenu }
            '2' { ShowGroupMenu}
            '3' { connectEOL  }
            '4' { settingsmenu  }
            'Q' { return }
            default { Write-Host "Invalid choice, please try again." }
        }
        #        PauseForUser
    } while ($choice -ne 'Q')
}

function settingsmenu {
    $choice = $null
        do {
        Clear-Host
        checkeol #check if connected to exchange online
        showgam #check if gam is working
        Write-Host "Setting Menu"
        Write-Host "1. Set Google to 365 routing group"
        Write-Host "2. Set 365 to google Group tools"
        Write-Host "Q. Quit"
        $choice = Read-Host "Please select an option"
        switch ($choice) {
            '1' { }
            '2' { }
            'Q' { return }
            default { Write-Host "Invalid choice, please try again." }}
     } while ($choice -ne 'Q') }


    Function connectEOL {
        #function to connect to exchange online
        Clear-Host
        $connect =Read-Host "Do you want to connect to Exchange Online? Y/(n)"
        switch ($connect) {
            'Y' { Connect-ExchangeOnline
                 Return }
            default { write-host "not connected" }
        }
    }
function ShowUserMenu {
    $choice = $null
    do {
        Clear-Host
        checkeol
        showgam
        Write-Host "User menu"
        Write-Host "1. Single user tools"
        Write-Host "2. Bulk user tools"
        Write-Host "B. Back to Main Menu"
        Write-Host "Q. Quit"

        $choice = Read-Host "Please select an option"

        switch ($choice) {

            '1' {clear-host
                 $username=askforuser 
                 ShowUserSingleMenu -userInput $Username
                 } #show single user menu
            '2' { Write-Host ([Environment]::UserName) }
            'B' { return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

        PauseForUser

    } while ($choice -ne 'B')
}
function askforuser{
    $Username = read-host "enter username to process"
    $check=checkExistoffice365 -userinput $Username #check if user exists in office 365
    if ($check -eq $null) {
    CheckexistsGoogle -userInput $Username
    Write-Host "The variable is null."
    Write-Host "username" $username
    Write-Host "check" $check
    PauseForUser
} else {
    Write-Host "The variable is not null."
    Write-Host "username" $username
    Write-Host "check" $check
    pauseforuser
}

    # if checkExistoffice365 -is null call checkexistsGoogle

    

    
  
    #CheckexistsGoogle -userinput $Username #check if user exists in google workspace
return $Username}


function ShowUserSingleMenu {
    param (
        [string]$userInput
    )
    $choice = $null
    do {
        Clear-Host
        checkeol #check if connected to exchange online
        showgam #check if gam is working
        
        Write-Host "current User : "-NoNewline
        Write-Host $userinput -ForegroundColor Green
        UserMigrationStatus -userinput $userInput         
        Write-Host ""
        Write-Host "User menu"
        Write-Host "1. Mailbox Refresh"
        Write-Host "2. Start routing to Outlook"
        Write-Host "3. Archive management"
        Write-Host "4. Test User"
        Write-Host "5. routing User"
        Write-host "9. Change Username"
        Write-Host "B. Back to Main Menu"
        Write-Host "Q. Quit"

               $choice = Read-Host "Please select an option"

        switch ($choice) {
            '1' { Write-Host "Refreshing mailbox"
                Refreshmailbox -userinput  $userinput
                 }

            '2' { Write-Host "Routing to outlook "
             Routingtooutlook -userinput  $userinput
                 }

            '3' { Write-Host "Archive management"
            enableDisableArchive -userinput  $userinput
                 }
            '4' { Write-Host "test user"
            write-host "checking if $userinput exists"
            pause 5
            checkExistoffice365  -userinput  $userinput
                 }
                 '5' { Write-Host "Status management"
                 cheeckgooglerouting -userInput $userInput
                      }

            '9' {Clear-Host
                 $userInput=  read-host "enter username to process"
                 }
            'B' { return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

        PauseForUser

    } while ($choice -ne 'B')
}

Function checkeol{
    if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
        Write-Host "Status: " -NoNewline
        Write-Host "You are connected to Exchange Online." -ForegroundColor Green
    } else {
        Write-Host "Status:" -NoNewline
        Write-Host "You are NOT connected to Exchange Online." -ForegroundColor Red
        } 
    
 }
function showGam {
    param (
        [string]$show
    )
    if ($output -notmatch "ERROR") {
        Write-Host "Status: " -NoNewline
        Write-Host "GAM seems to be set up correctly." -ForegroundColor Green
        } else {
            Write-Host "Status:" -NoNewline
            Write-Host "GAM is not set up or there's an issue with connectivity."-ForegroundColor Red
        
    }

}
 Function checkGAM {
    $output = & gam info domain 2>&1
 showgam -show $output
  }

  function CheckexistsGoogle{
    param (
        [string] $userInput
    )
    $userToCheck = $userInput
    # Check if the user exists using GAM
try {
    $output = & gam info user $userToCheck 2>&1
    if ($output -like "*Error:*") {
        Write-Host "$userToCheck does not exist." -ForegroundColor Red
        #call again a askforuser function and try again
        $userToCheck = askforuser
      #  CheckexistsGoogle -userInput $userToCheck

        
    } else {
        Write-Host "$userToCheck exists. "
        #call again a askforuser function and try again sugessting try with simiral address
       

      
    }
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
}
  #function to check whit gam if a user exists in google workspace
}


function checkExistoffice365 {
    param (
        [string] $userInput
    )
            try {
        # We'll try to retrieve just the alias
        $userUPN = Get-Mailbox -Filter "alias -eq '$userInput'" -ErrorAction Stop | Select-Object -ExpandProperty alias
        if ($userUPN) {
            Write-Host "$userInput exists in Exchange Online." -foreground Green
               
        } else {
            Write-Host "$userInput does not exist in office 365."
            #if does not exist the function return a null value
            
            return $null
       }
       return $userInput
    } catch {
        Write-Host "$userInput does not exist."
    }
}
       
function PauseForUser {
    Write-Host "Press any key to continue..."
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
function Refreshmailbox {
    param (
        [string]$userInput
    )
    Start-ManagedFolderAssistant $userInput
Get-MailboxStatistics $userInput  | format-table DisplayName, TotalItemSize
Get-MailboxStatistics $userInput -Archive | Format-Table DisplayName, TotalItemSize

}
function check365routing
{
    param (
        [string]$userInput
    )
    
    $groupMembers = Get-UnifiedGroupLinks -Identity "migradosa365" -LinkType Members | Select-Object -ExpandProperty PrimarySmtpAddress


    return $groupMembers -contains $UserEmail
}



function Routingtooutlook {
    param (
        [string]$userInput
    )
    gam update group reenviooffice add member $userInput
    Add-DistributionGroupMember -Identity migradosa365 -Member $userInput 
}
function UserMigrationStatus{
    param (
        [string] $userInput
    )
    # chequear tamaño buzon
   # $userUPN = Get-Mailbox -Filter "alias -eq '$userInput'" -ErrorAction Stop | Select-Object -ExpandProperty alias
 #Get-MailboxStatistics $userInput  | Select-Object  TotalItemSize
#   Get-MailboxStatistics $userInput  | format-table  TotalItemSize
$mailboxtotalSize = (Get-MailboxStatistics $userInput | Select-Object -ExpandProperty TotalItemSize)
$archivetotalSize = (Get-MailboxStatistics $userInput -Archive| Select-Object -ExpandProperty TotalItemSize)
$retentionPolicy = (Get-Mailbox -Identity $userInput).RetentionPolicy

if ($null -eq $retentionPolicy) {
    Write-Host "No retention policy is set for $userInput." -foreground green
} else {
    Write-Host "Retention policy for $userInput is: $retentionPolicy" -foreground red
}

write-host "Tamaño del buzon principal:" $mailboxtotalSize
write-host "Tamaño del buzon de Archivo:" $archivetotalSize

 #   Get-MailboxStatistics $userInput -Archive | Format-Table TotalItemSize
    # chequear archivo
    showarchive -userInput $userInput 

    # chequear grupo Reenvio 
    cheeckgooglerouting -userInput $userInput

    processreport userInput $userInput
    function gmailusedSpace {
        param (
            [Parameter(Mandatory=$true)]
            [string]$CsvPath,
    
            [Parameter(Mandatory=$true)]
            [string]$UserNameOrEmail
        )
    
        # Importar el CSV
        $data = Import-Csv -Path $CsvPath
    
        # Buscar al usuario basado en la parte del nombre de usuario/email antes del '@'
        $userData = $data | Where-Object { ($_.UserName -split "@")[0] -eq $UserNameOrEmail }
    
        # Si encontramos al usuario, devolver el espacio utilizado
        if ($userData) {
            return $userData.UsageMB
        } else {
            throw "El usuario $UserNameOrEmail no se encontró en el CSV."
        }
        try {
            $userSpace = GetUsedSpace -CsvPath "path\to\data.csv" -UserNameOrEmail "user1"
            Write-Output "El espacio utilizado por el usuario es: $userSpace MB"
        } catch {
            Write-Error $_.Exception.Message
        }
          
      
    }
    
    # Ejemplo de uso
    
}


function cheeckgooglerouting {
    param (
        [string] $userInput
    )
# Define group and user
$groupEmail = "reenviooffice"
$userEmail = $userInput

# Call GAM to retrieve members of the group and capture the output
$members = & gam print group-members group reenviooffice 2>$null

# Check if the user is in the output list
if ($members -like "*$userEmail*") {
    Write-Host "The user $userEmail is a member of the group reenviooffice." -foreground green
} else {
    Write-Host "The user $userEmail is not a member of the group reenviooffice." -foreground red
}


}

function showarchive {
    param (
        [string]$userInput
    )

    switch ($Archivestatus = get-mailbox $userInput | Select-Object -ExpandProperty ArchiveStatus)
        
{
    'Active' { write-host "archivo Habilitado"-ForegroundColor Green}
    'none'   { write-host "archivo Deshabilitado"-ForegroundColor Red}
   } 

}

function processreport {
    param (
        [string]$userInput
    )
    $userdata= $userInput
    $data = Import-Csv c:\2\gmail-used-1609.csv -delimiter "," # "Source Email" "Destination Email"

    # Convert MB to GB with decimals
    $convertedData = $data | ForEach-Object {
        # Ensure the MB value is treated as a double and perform division
        $gbValue = [math]::Round(([double]$_."Gmail storage used (MB) [2023-09-13 GMT]" / 1024), 2)  # the "2" here indicates to round to 2 decimal places
        $_ | Add-Member -Type NoteProperty -Name "UsageGB" -Value $gbValue
        $_
    }
      

    # Ask for the username
   # $userName = Read-Host "Please enter the username"
    #"Gmail storage used (MB) [2023-09-13 GMT]"
    # Import the CSV
    # Find the user and display the used space
    $userData = $data | Where-Object { ($_.User -split "@")[0] -eq $userName }
    
    if ($userData) {
        Write-Output "User $($userData.User) has used $($userData.UsageGB) GB in Google Workspace"
    } else {
        Write-Output "Username $userName not found in the CSV."
    }
    
    $usedGbuser = $userData.UsageGB
       
}
function enableDisableArchive {
    param (
        [string]$userInput
    )


    $choice = $null
    do {
        UserMigrationStatus
        write-host "The current status of the current mailbox archive are : " -NoNewline 
        Write-Host $Archivestatus -ForegroundColor Green
        write-host "Choose an option"
        write-host "S. Show Archive status and size"
        write-host "E. Enable Archive"
        write-host "D. Disable Archive"
        Write-Host "B. Back"
        Write-Host "Q. Quit"

        $choice = Read-Host "Please select an option"

        switch ($choice) {

            'B' { return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

        PauseForUser

    } while ($choice -ne 'B')   
    
    


}

# Call the main menu

#processreportq
ShowMainMenu
$userinput= $null
$Username = $null
$Global:grupoenrutamientoGmail = "reenviooffice"
$Global:grupoenrutamiento365= "migradoosa365"