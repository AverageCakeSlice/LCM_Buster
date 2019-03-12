#LCM Buster - Moves LCM machines to their correct location in AD, and sets AD object properties for the new machine.

#~~~~~~~~~VARIABLE INITIALIZATION~~~~~~~~~#
[string]$newLCMComputerName  =  Read-Host -prompt 'Input the New Computer Name'
[string]$oldLCMComputerName = Read-Host -prompt 'Input the Old Computer Name'
[string]$username = Read-Host -prompt 'Input assigned user of the new asset'
[string]$budgetCode = Read-Host -prompt 'Input the budget code for the assigned user'
[string]$shortOUname = Read-Host -prompt 'Input the desired OU for the new asset'

#~~~~~~~~~FUNCTION DEFINITIONS~~~~~~~~~#

#The Get-VPNStatus function checks if Cisco Anyconnect is actually connected. If it is not, it will launch the anyconnect window to prompt the user to login.
function Get-VPNStatus()
{
    $netAdapter = Get-NetAdapter -InterfaceDescription *AnyConnect* | Select-Object -ExpandProperty status
    while ($netAdapter -ne "Up") {
            Write-Host -ForegroundColor Yellow "* Connection failed *"
            Start-Process -FilePath "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
            Read-Host "AnyConnect VPN will now open to verify login. Press 'ENTER' here to try again"
            $netAdapter  =  Get-NetAdapter -InterfaceDescription *AnyConnect* | Select-Object -ExpandProperty status
        }  
    Write-Host -ForegroundColor Green "* Connection to VPN successful *" 
}

#The Set-NewMachine function calls the Test functions to validate input, then moves the selected AD object to the user-specified OU. 
function Set-NewMachine ([string]$budgetCode, [string]$newLCMComputerName, [string]$shortOUName, [string]$username)
{
    Test-NewMachine $newLCMComputerName
    Test-Staff $username
    Test-OU $shortOUName

    Get-ADComputer $newLCMComputerName | Set-ADComputer -replace @{businessCategory = "$budgetCode"} -Description "Assigned to: $username" -Verbose
        
    $fullOUName = Get-ADOrganizationalUnit -LDAPFilter "(Name=*$shortOUname)" -Searchbase 'OU=University Computers,DC=University,DC=liberty,DC=edu'

    Get-ADComputer $newLCMComputerName | Move-ADObject -TargetPath $fullOUName -Verbose   
    if($? -eq 'True') 
        {
            Write-Host -ForegroundColor Yellow 'New asset information has been updated with your input. Please note, it can take up to a minute for these changes to reflect in AD.'
        }
    else
        {
            Get-ADComputer $newLCMComputerName | Move-ADObject -targetpath "OU=Workstations,OU=University Computers,DC=University,DC=liberty,DC=edu"
            Write-Host -ForegroundColor Yellow 'The new asset was not moved - This is likely due to typing the name incorrectly. It will default to the Workstations OU.'
        }
}#End of Set-NewMachine function

#The Set-OldMachine function moves the old device to the 'Retiring Computers' OU, and changes the description to reflect the retirmenet status.
#If an error is encountered, the user will be reprompted to enter in the old computer's serial number.
function Set-OldMachine([string]$oldLCMComputerName)
{
    do 
    {
        try
        {
            $wasComputerMovedStatus = $true
            Get-ADComputer $oldLCMComputerName  | Move-ADObject -TargetPath "OU=Retiring Computers,OU=Customer Support,OU=LUIS,OU=Workstations,OU=University Computers,DC=University,DC=liberty,DC=edu"
        } 
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
        {
            $wasComputerMovedStatus = $false
            Write-Host -ForegroundColor Yellow "The old asset was not moved properly. The old asset may not exist in Active Directory, or it was bound using a different name. Please re-enter the serial number."
            $oldLCMComputerName = Read-Host -prompt 'Input the Old Computer Name'
        }
    } while ($wasComputerMovedStatus -ne 'True')
    Write-Host -ForegroundColor Yellow "The old asset has been moved to the Retiring Computers OU and a description was set. The script is now finished."

    Set-ADComputer $oldLCMComputerName -Description "Retired - LCM"
}

#The Test-NewMachine function checks to see if the new device that the user is trying to move actually exists as an object in Active Directory.
#Performs input validation and reprompts if the name does not exist in AD.
function Test-NewMachine ([string]$newLCMComputerName)
{
    do
    {
        try
        {
            $machineCheck = Get-ADComputer $newLCMComputerName
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
        {
            Write-Host -ForegroundColor Red "The asset that you entered does not exist in ActiveDirectory. Please try again."
            $newLCMComputerName  =  Read-Host -prompt 'Input the New Computer Name'
        }
    }while ($machineCheck -notlike "*$newLCMComputerName*")
    Write-Host -foregroundcolor Yellow "New asset exists in ActiveDirectory."
}#End of Test-Machine function

#The Test-OU function examines the short OU name and verifies that it exists.
#If the shortname does not match an existing OU name, it will reject the input and reprompt the user for input.
function Test-OU ([string]$shortOUName)
{
    do
    {
        $OUNameList = Get-ADObject -Filter {ObjectClass -eq 'OrganizationalUnit'} | Select-Object Name
        for ($i = 0; $i -lt $OUNameList.Length; $i++)
        {
            if ($OUNameList.Item($i).Name -eq "$shortOUName")
            {
                $isValidOU = $true
                break
            }
        }
        if ($isValidOU -ne $true)
        {
            Write-Host -ForegroundColor Red "The OU you entered does not exist within University.liberty.edu. Please enter a valid OU." 
            $shortOUname = Read-Host -prompt 'Input the desired OU for the new asset'
        }
    } while ($isValidOU -ne $true)
    Write-Host -ForegroundColor Yellow "Verified that this OU exists in our domain."
}#End of Test-OU function

#The Test-Staff function will examine the username that you enter and see what AD Groups that user is in.
#If the username entered is not in the Staff group, an error will be caught, and user will be repromted to input the username.
function Test-Staff ([string]$username)
{
    do
    {
        try
        {
            $staffCheck = Get-ADUser $username -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup -Properties name | Select-Object name | Where-Object -Property name -eq "Staff"
        } 
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
        {
            Write-Host -ForegroundColor Red "Username is not in the Staff group. Please re-enter the username."
            $username = Read-Host -prompt 'Input assigned user of the new asset'
        }
    } while ($staffCheck -notlike "*Staff*")
    Write-Host -foregroundcolor yellow "User is in the Staff group."
}#End of Test-Staff function

#~~~~~~~~~MAIN PROCESS~~~~~~~~~#
Get-VPNStatus

Set-NewMachine $budgetCode $newLCMComputerName $shortOUname $username 

Pause

Set-OldMachine $oldLCMComputerName

#End of script LCMBuster
