
function ShowMainMenu {
# main menu (1) user tools (2) group tools (3) connect to exchange online (4) connect to azure ad (5) tool settings (q) quit
    $choice = $null
        do {
        Clear-Host
        check_eol #check if connected to exchange online
        show_gam #check if gam is working
        check_AzureAD #check if connected to azure ad
        Write-Host "Main Menu" 
        Write-Host "1. User tools"
        Write-Host "2. Group tools"
        Write-Host "3. Connect to Exchange online"
        write-host "4. Connect to Azure AD"
        Write-Host "5. Tool Settings"
        Write-Host "Q. Quit"
        $choice = Read-Host "Please select an option"
        switch ($choice) {
            '1' { ShowUserMenu }
            '2' { ShowGroupMenu}
            '3' { Connect_EOL}
            '4' { Connect_AzureAD}
            '5' { settingsmenu  }
            'Q' { return }
            default { Write-Host "Invalid choice, please try again." }
        }
        #        PauseForUser
    } while ($choice -ne 'Q')
}


function ShowUserMenu {
    $choice = $null
    do {
        Clear-Host
        check_eol
        show_gam
        Write-Host "User menu"
        Write-Host "1. Single user tools"
        Write-Host "2. Bulk user tools"
        Write-Host "B. Back to Main Menu"
        Write-Host "Q. Quit"

        $choice = Read-Host "Please select an option"

        switch ($choice) {

            '1' {clear-host
                 do {$username=askforuser }
                 while ($null -eq $username) {} 
                 ShowUserSingleMenu -UserObject $username #access third variable of the array and show single user menu
                } 
            '2' {}
            'B' { return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

        PauseForUser

    } while ($choice -ne 'B')
}
function ShowUserSingleMenu {
        param (
        [PSCustomObject]$UserObject
    )
         # Accessing the properties of the object
     $officeUser = $UserObject.Office365User
     $googleUser = $UserObject.GoogleWorkspaceUser
  
     $choice = $null
    do {
       #check_eol #check if connected to exchange online
       #show_gam #check if gam is working
        Write-Host "current User : "-NoNewline
        Write-Host $googleUser -ForegroundColor Green
        UserMigrationStatus -UserObject $UserObject      
        Write-Host ""
        Write-Host "User menu"
        Write-Host "1. Mailbox Refresh"
        Write-Host "2. Start routing to Outlook"
        Write-Host "3. Stop routing to Outlook"
        Write-Host "4. Test User"
        Write-Host "5. routing User"
        Write-Host "6. remove retention policy"
        Write-Host "7. Archive management"
        Write-host "9. Change Username"
        write-host "P. Purge menu"
        Write-Host "B. Back to Main Menu"
        Write-Host "Q. Quit"

               $choice = Read-Host "Please select an option"

        switch ($choice) {
            '1' { Write-Host "Refreshing mailbox"
                Refreshmailbox -userinput  $officeUser
                 }

            '2' { Write-Host "Routing to outlook "
             Routingtooutlook -userinput  $officeUser
                 }

            '3' { Write-Host "removing routing to oulook"
            RemoveRoutingtooutlook -userinput  $officeUser
                 }
            '4' { 
                 }
                 '5' { Write-Host "Status management"
                 cheeckgooglerouting -userInput $googleUser
                      }
                 '6' { Write-Host "remove retention policy"
                      remove-retentionpolicy -userInput $officeUser
                             }
                '7' { Write-Host "Archive management"
                   enableDisableArchive -UserObject  $UserObject
                                  }
            '9' {Clear-Host
                do
                {
                   $username=askforuser 
                }
                   while ($null -eq $username) {
                      
                } 
                ShowUserSingleMenu -userObject $username #access third variable of the array and show single user menu
                 }
            'P' { Menu_purge_mailbox -UserObject $UserObject }
            'B' { return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

     #   PauseForUser

    } while ($choice -ne 'B')
}

function remove-retentionpolicy {
    param (
        [string]$userInput
    )
    Set-Mailbox -Identity $userInput -RetentionPolicy $null
} 

function Menu_Purge_mailbox {
    param (
        [string]$UserObject
    )
    $officeUser = $UserObject.Office365User
    $choice = $null
    Write-Host "current User : "-NoNewline
    Write-Host $userInput -ForegroundColor Green
    UserMigrationStatus -UserObject $UserObject      
    Write-Host ""
    Write-Host "PurgueUser menu"
    Write-Host "1. Mailbox Refresh"
    Write-Host "2. Show retention policy"
    Write-Host "3. change retention policy to purge mailbox"
    Write-Host "4. create compliance purge job"
    Write-Host "5. start purge Job"
    write-Host "B. Back to Main Menu"
    Write-Host "Q. Quit"
    $choice = Read-Host "Please select an option"

    switch ($choice) {
        '1' { Write-Host "Refreshing mailbox"
            Refreshmailbox -userinput  $UserObject}
        '2' { Write-Host "Show retention policy"
            showretentionpolicy -userinput  $UserObject}
        '3' { Write-Host "change retention policy to purge mailbox"
        Set-Mailbox $officeUser -RetentionPolicy "Purgado buzon" -RetainDeletedItemsFor 0 -SingleItemRecoveryEnabled $false
                    }
        '4' { Write-Host "create compliance purge job"
                }
        '5' { Write-Host "start purge Job"
                }
        'B' { return }
        'Q' { Exit }
        default { Write-Host "Invalid choice, please try again." }
            
}
}

Function check_eol{
    if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
        Write-Host "Status: " -NoNewline
        Write-Host "You are connected to Exchange Online." -ForegroundColor Green
    } else {
        Write-Host "Status:" -NoNewline
        Write-Host "You are NOT connected to Exchange Online." -ForegroundColor Red
        } 
    
 }
function show_gam {
    param (
        [string]$show
    )
    if ($output -notmatch "ERROR" -or $output -notmatch "*OAuth*")  {
        Write-Host "Status: " -NoNewline
        Write-Host "GAM seems to be set up correctly." -ForegroundColor Green
        } else {
            Write-Host "Status:" -NoNewline
            Write-Host "GAM is not set up or there's an issue with connectivity."-ForegroundColor Red
        
    }

}
 Function checkGAM {
    $output = & gam info domain 2>&1
 show_gam -show $output
  }

  function CheckexistsGoogle{
    param (
        [string] $userInput
    )
        # Check if the user exists using GAM
try {
    $output = & gam info user  $userInput 2>&1
    if ($output -like "*Error:*") {
        Write-Host " $userInput does not exist in Google Workspace." -ForegroundColor Red
        #call again a askforuser function and try again
        return $null
      #  CheckexistsGoogle -userInput  $userInput
      
    } else {
        Write-Host "$userInput exists in Google Workspace " -ForegroundColor Green
        #call again a askforuser function and try again sugessting try with simiral address
         return $userInput    
    }
    return $userInput    
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
        $userAlias = Get-Mailbox -Filter "alias -eq '$userInput'" -ErrorAction Stop | Select-Object -ExpandProperty alias
        if ($userAlias) {
            Write-Host "$userInput exists in in office 365." -foreground Green
               
        } else {
            Write-Host "$userInput does not exist in office 365." -foreground Red
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
Write-Host "Mailbox Refreshed" -ForegroundColor Green
#3 seconds pause
Start-Sleep -s 3

Clear-Host

}
function check365routing
{

    param (
        [string]$userInput
    )
  $distList = "migradosa365"

# Check if the user is a member of the distribution list
$members = Get-DistributionGroupMember -Identity $distList  | select-object alias

if ($members | Where-Object { $_.alias -eq $userInput }) {
    Write-Host "$userInput is a member of $distList." -foreground green 
    return $true
} else {
    Write-Host "$userInput is NOT a member of $distList." -foreground red
    return $false
}
}

function RemoveRoutingtooutlook 
{
    param (
        [string]$userInput
    )
    gam update group reenviooffice remove member $userInput
    Remove-DistributionGroupMember -Identity migradosa365 -Member $userInput 
    #gam update group reenviooffice remove member $userInput
   # $check2=cheeckgooglerouting -userInput $userInput
    #$check1=check365routing -userInput $userInput 
    #if ($check1 -eq $true) {
        
    #}
    #if ($check2 -eq $true) {
        
        #& gam update group reenviooffice remove member $userInput
    #}
    #Add-DistributionGroupMember -Identity migradosa365 -Member $userInput
}


function Routingtooutlook {
    param (
        [string]$userInput
    )
    gam update group reenviooffice add member $userInput
    Add-DistributionGroupMember -Identity migradosa365 -Member $userInput 
    $check2=cheeckgooglerouting -userInput $userInput
    $check1=check365routing -userInput $userInput 

    }
    
function UserMigrationStatus{

<#
.SYNOPSIS
    Checks the migration status of a user.
.DESCRIPTION
    shows all info about a user in both platforms

.EXAMPLE
    UserMigrationStatus -userInput $userInput

#>
    param (
        [PSCustomObject]$UserObject
    )
   
    $officeUser = $UserObject.Office365User
    $googleUser = $UserObject.GoogleWorkspaceUser

    # show if archive are enabled or not (active or none)
    $Archivestatus = get-mailbox $officeUser | Select-Object -ExpandProperty ArchiveStatus 

    # shows exchange online total size
$mailboxtotalSize = (Get-MailboxStatistics $officeUser | Select-Object -ExpandProperty TotalItemSize)
write-host "Exchange online principal mailbox:" $mailboxtotalSize

# shows exchange online archive size if exist, show Archive not enabled if not exist
if ($Archivestatus -eq "Active") {
    $archivetotalSize = (Get-MailboxStatistics $officeUser -Archive| Select-Object -ExpandProperty TotalItemSize)
    write-host "Exchange online Archive mailbox::" $archivetotalSize
} elseif ($Archivestatus -eq "none") {
    write-host "Archive not enabled" -Foreground Yellow
}

# shows google workspace size
$usedGbuser=processreport -userInput $googleUser
write-host "Google workspace mailbox:"$usedGbuser "GB"

# shows retention policy
$retentionPolicy = (Get-Mailbox -Identity $officeUser).RetentionPolicy
if ($null -eq $retentionPolicy) {
    Write-Host "No retention policy is set for $officeUser." -foreground green
} else {
    Write-Host "Retention policy for $officeUser is: $retentionPolicy" -foreground red
}
#$prueba = showarchive -userInput $userInput
    # chequear archivo
  #  showarchive -userInput $userInput 

    # chequear grupo Reenvio 
    cheeckgooglerouting -userInput $googleUser
    check365routing -userInput $officeUser
   
}
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
           } catch {
            Write-Error $_.Exception.Message
        }
          
      
    }
    
    # Ejemplo de uso
    



function cheeckgooglerouting {
    param (
        [string] $userInput
    )
# Define group and user
#$groupEmail = "reenviooffice"
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

function checkarchive {
    param (
        [string]$userInput
    )
    $Archivestatus = get-mailbox $userInput | Select-Object -ExpandProperty ArchiveStatus
    #switch ($Archivestatus = get-mailbox $userInput | Select-Object -ExpandProperty ArchiveStatus)
     #   
#{
 #   'Active' { write-host "archivo Habilitado"-ForegroundColor Green}
  #  'none'   { write-host "archivo Deshabilitado"-ForegroundColor Red}
  # } 
 return $Archivestatus
} 

function check_EOL_PLAN{
    param (
        [string]$userInput
     )

    $plan = (Get-MsolAccountSku).AccountSkuId
    if ($plan -like "*ENTERPRISEPACK*") {
        Write-Host "Status: " -NoNewline
        Write-Host "You have an E3 license." -ForegroundColor Green
    } else {
        Write-Host "Status:" -NoNewline
        Write-Host "You do not have an E3 license." -ForegroundColor Red
        } q
}

function Connect_AzureAD {
    Clear-Host
    $connect =Read-Host "Do you want to connect to Azure AD? Y/(n)"
    switch ($connect) {
        'Y' { Connect-AzureAD
             Return }
        default { write-host "not connected" }
    }
}

function check_AzureAD {
    if (Get-Command Get-AzureADUser -ErrorAction SilentlyContinue) {
        Write-Host "Status: " -NoNewline
        Write-Host "You are connected to Azure AD." -ForegroundColor Green
    } else {
        Write-Host "Status:" -NoNewline
        Write-Host "You are NOT connected to Azure AD." -ForegroundColor Red
        } 
}
function settingsmenu {
    $choice = $null
        do {
        Clear-Host
        check_eol #check if connected to exchange online
        show_gam #check if gam is working
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


    Function Connect_EOL {
        #function to connect to exchange online
        Clear-Host
        $connect =Read-Host "Do you want to connect to Exchange Online? Y/(n)"
        switch ($connect) {
            'Y' { Connect-ExchangeOnline
                 Return }
            default { write-host "not connected" }
        }
    }
function processreport {
    param (
        [string]$userInput
    )
   # $userdata= $userInput
    $data = Import-Csv c:\2\gmail-used-1609.csv -delimiter "," # "Source Email" "Destination Email"
# Convert MB to GB with decimals
$convertedData = $data | ForEach-Object {
    # Extract the storage value and ensure it's treated as a double
    $storageMB = $_."Gmail storage used (MB) [2023-09-13 GMT]" -replace '[^0-9.]', ''  # Removing non-numeric characters

    # Ensure the MB value is treated as a double and perform division
    $gbValue = [math]::Round(([double]$storageMB / 1024), 2)  # the "2" here indicates to round to 2 decimal places
    $_ | Add-Member -Type NoteProperty -Name "UsageGB" -Value $gbValue
    $_
}

    $userData = $data | Where-Object { ($_.User -split "@")[0] -eq $userInput }
    
    if ($userData) {
      #  Write-Output "User $($userData.User) has used $($userData.UsageGB) GB in Google Workspace"
            } else {
        Write-Output "Username $userName not found in the CSV."
    }
    
    $usedGbuser = $userData.UsageGB
     return $usedGbuser  
}
function enableDisableArchive {
      
        param (
            [PSCustomObject]$UserObject
        )
       
        $officeUser = $UserObject.Office365User
        $googleUser = $UserObject.GoogleWorkspaceUser     


    $choice = $null
    do {
        UserMigrationStatus -UserObject $userObject
        $status=checkarchive -userInput $officeUser
        #UserMigrationStatus
        write-host "The current status of the current mailbox archive are : " -NoNewline 
        Write-Host $status -ForegroundColor Green
        write-host "Choose an option"
        write-host "E. Enable Archive"
        write-host "D. Disable Archive"
        Write-Host "B. Back"
        Write-Host "Q. Quit"

        $choice = Read-Host "Please select an option"

        switch ($choice) {
            'E' { Write-Host "Enabling Archive"
                enable-mailbox $officeUser -archive 
                 }	
            'D' { Write-Host "Disabling Archive"
                disable-mailbox $officeUser -archive 
        }

            'B' { clear-host
                return }
            'Q' { Exit }
            default { Write-Host "Invalid choice, please try again." }
        }

#        PauseForUser

    } while ($choice -ne 'B')   
    
    


}
function askforuser {
    do {
        # Prompt for Office 365 username
        $O365Username = Read-Host "Enter Office 365 username to process"
        $check_o365 = checkExistoffice365 -userinput $O365Username

        # If the user doesn't exist in Office 365, inform and ask again
        if ($null -eq $check_o365) {
            Write-Host "The user $O365Username does not exist in Office 365." -ForegroundColor red
            continue
        } else {
            # Since Office 365 username is valid, ask if you want to change the Google Workspace username
            $decision = Read-Host "Office 365 username is valid. Do you want to provide a different Google Workspace username? (Y/N)"
            if ($decision -eq 'Y' -or $decision -eq 'y') {
                while ($true) {
                    $GWorkspaceUsername = Read-Host "Enter Google Workspace username to process"
                    $check_gam = CheckexistsGoogle -userInput $GWorkspaceUsername

                    # If the user exists in Google Workspace, break out of the loop
                    if ($null -ne $check_gam) {
                        Write-Host "The user $GWorkspaceUsername exists in Google Workspace." -ForegroundColor green
                        break
                    } else {
                        Write-Host "The user $GWorkspaceUsername does not exist in Google Workspace." -ForegroundColor red
                    }
                }
            } else {
                $GWorkspaceUsername = $O365Username
            }
        }

    } while ($null -eq $check_o365)

    # Return an object with both valid usernames
    return @{
        Office365User = $O365Username
        GoogleWorkspaceUser = $GWorkspaceUsername
    }
}
#processreportq
Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll'
ShowMainMenu
$userinput= $null
$Username = $null
$Global:grupoenrutamientoGmail = "reenviooffice"
$Global:grupoenrutamiento365= "migradoosa365"