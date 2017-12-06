###########################################################
# AUTHOR  : Emilien Devos - Marius / Hican 
# DATE    : 26-04-2012 
# EDIT    : 04-12-2017
# COMMENT : This script creates, updates and removes Active Directory users,
#           including different kind of properties, based
#           on an input_create_ad_users.csv.
# VERSION : 1.4
###########################################################

# CHANGELOG
# Version 1.2: 15-04-2014 - Changed the code for better
# - Added better Error Handling and Reporting.
# - Changed input file with more logical headers.
# - Added functionality for account Enabled,
#   PasswordNeverExpires, ProfilePath, ScriptPath,
#   HomeDirectory and HomeDrive
# - Added the option to move every user to a different OU.
# Version 1.3: 08-07-2014
# - Added functionality for ProxyAddresses
# Version 1.4: 04-12-2017
# - Removing useless fields
# - Add ability to update and delete users automatically

# TODO
# - Ability to add multiple groups

# ERROR REPORTING ALL
Set-StrictMode -Version latest

#----------------------------------------------------------
# LOAD ASSEMBLIES AND MODULES
#----------------------------------------------------------
Try
{
  Import-Module ActiveDirectory -ErrorAction Stop
}
Catch
{
  Write-Host "[ERROR]`t ActiveDirectory Module couldn't be loaded. Script will stop!"
  Exit 1
}

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
$path     = Split-Path -parent $MyInvocation.MyCommand.Definition
$newpath  = $path + "\import_create_ad_users.csv"
$log      = $path + "\create_ad_users.log"
$date     = Get-Date
$addn     = (Get-ADDomain).DistinguishedName
$dnsroot  = (Get-ADDomain).DNSRoot
$i        = 1

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------
Function Start-Commands
{
  Create-Users
  Delete-Users
}

#----------------------------------------------------------
#FUNCTION Create and update users
#----------------------------------------------------------

Function Create-Users
{
  "Processing started (on " + $date + "): " | Out-File $log -append
  "--------------------------------------------" | Out-File $log -append
  Import-CSV $newpath -Encoding UTF8 | ForEach-Object {
    If (($_.Implement.ToLower()) -eq "yes")
    {
      If (($_.GivenName -eq "") -Or ($_.LastName -eq "") -Or ($_.Group -eq "") -Or ($_.ProfilePath -eq "") -Or ($_.PasswordNeverExpires -eq "") -Or ($_.Enabled -eq ""))
      {
        Write-Host "[ERROR]`t Please provide valid GivenName, LastName and Group. Processing skipped for line $($i)`r`n"
        "[ERROR]`t Please provide valid GivenName, LastName and Group. Processing skipped for line $($i)`r`n" | Out-File $log -append
      }
      Else
      {

        # Set the target OU
        $location = "OU=DB,$($addn)"

        # Set the Enabled and PasswordNeverExpires properties
        If (($_.Enabled.ToLower()) -eq "true") { $enabled = $True } Else { $enabled = $False }
        If (($_.PasswordNeverExpires.ToLower()) -eq "true") { $expires = $True } Else { $expires = $False }

        # Replace dots / points (.) in names, because AD will error when a 
        # name ends with a dot (and it looks cleaner as well)
        $replaceName = $_.Lastname.Replace(".","")
        If($replaceName.length -lt 4)
        {
          $lastname = $replaceName
        }
        Else
        {
          $lastname = $replaceName.substring(0,4)
        }
        $replaceFirstName = $_.GivenName.Replace(".","")
        If($replaceFirstName.length -lt 2)
        {
          $givenname = $replaceFirstName
        }
        Else
        {
          $givenname = $replaceFirstName.substring(0,2)
        }

        $lastname = $lastname.substring(0,1).toupper()+$lastname.substring(1).tolower()   
        $sam = $lastname + $givenname.ToUpper()


        Try   { $exists = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" }
        Catch { }
        If(!$exists)
        {
          # Set all variables according to the table names in the Excel 
          # sheet / import CSV. The names can differ in every project, but 
          # if the names change, make sure to change it below as well.
          $setpass = ConvertTo-SecureString -AsPlainText $_.Password -force

          Try
          {
            Write-Host "[INFO]`t Creating user : $($sam)"
            "[INFO]`t Creating user : $($sam)" | Out-File $log -append
            New-ADUser $sam -GivenName $_.GivenName `
            -Surname $_.LastName -DisplayName ($_.LastName + "," + $_.GivenName) `
            -UserPrincipalName ($sam + "@" + $dnsroot) `
            -AccountPassword $setpass `
            -profilePath $_.ProfilePath `
            -Enabled $enabled -PasswordNeverExpires $expires
            Write-Host "[INFO]`t Created new user : $($sam)"
            "[INFO]`t Created new user : $($sam)" | Out-File $log -append
     
            $dn = (Get-ADUser $sam).DistinguishedName
       
            # Move the user to the OU ($location) you set above. If you don't
            # want to move the user(s) and just create them in the global Users
            # OU, comment the string below
            If ([adsi]::Exists("LDAP://$($location)"))
            {
              Move-ADObject -Identity $dn -TargetPath $location
              Write-Host "[INFO]`t User $sam moved to target OU : $($location)"
              "[INFO]`t User $sam moved to target OU : $($location)" | Out-File $log -append
            }
            Else
            {
              Write-Host "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!"
              "[ERROR]`t Targeted OU couldn't be found. Newly created user wasn't moved!" | Out-File $log -append
            }
       
            # Rename the object to a good looking name (otherwise you see
            # the 'ugly' shortened sAMAccountNames as a name in AD. This
            # can't be set right away (as sAMAccountName) due to the 20
            # character restriction
            $newdn = (Get-ADUser $sam).DistinguishedName
            Rename-ADObject -Identity $newdn -NewName ($_.GivenName + " " + $_.LastName)
            Write-Host "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)"
            "[INFO]`t Renamed $($sam) to $($_.GivenName) $($_.LastName)" | Out-File $log -append
            
            Write-Host "[INFO]`t Added $($sam) to group $($_.Group)"
            "[INFO]`t Added $($sam) to group $($_.Group)`r`n" | Out-File $log -append
            Add-ADGroupMember -Identity $_.Group -Members $sam
          }
          Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
          }
        }
        elseif($exists)
        {
          $setpass = ConvertTo-SecureString -AsPlainText $_.Password -force

          Try
          {

            # Updating user using the values from the CSV file
            Write-Host "[INFO]`t User already exist, updating user : $($sam)"
            "[INFO]`t User already exist, updating : $($sam)" | Out-File $log -append
            Set-ADUser $sam -GivenName $_.GivenName `
            -Surname $_.LastName -DisplayName ($_.LastName + "," + $_.GivenName) `
            -UserPrincipalName ($sam + "@" + $dnsroot) `
            -profilePath $_.ProfilePath `
            -Enabled $enabled -PasswordNeverExpires $expires

            # Changing the password of the user according to the CSV file
            Set-ADAccountPassword $sam -NewPassword $setpass
            Write-Host "[INFO]`t Updated user : $($sam)"
            "[INFO]`t Updated user : $($sam)" | Out-File $log -append
            
            # If his group has changed on the CSV, deleting his old group and adding to the new group.

            $secondgroup = Get-ADPrincipalGroupMembership $sam | select name | select -Index 1 | Select -ExpandProperty "name"
            if ($secondgroup -ne $_.Group)
            {
              Remove-ADGroupMember -Identity $secondgroup -Members $sam -Confirm:$false
              Write-Host "[INFO]`t Updated group of $($sam) to group $($_.Group)"
              "[INFO]`t Added $($sam) to group $($_.Group)`r`n" | Out-File $log -append
              Add-ADGroupMember -Identity $_.Group -Members $sam
            }
          }
          Catch
          {
            Write-Host "[ERROR]`t Oops, something went wrong: $($_.Exception.Message)`r`n"
          }
        }
        Else
        {
          Write-Host "[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) returned an error!`r`n"
          "[SKIP]`t User $($sam) ($($_.GivenName) $($_.LastName)) returned an error!" | Out-File $log -append
        }
      }
    }
    Else
    {
      Write-Host "[SKIP]`t User ($($_.GivenName) $($_.LastName)) will be skipped for processing!`r`n"
      "[SKIP]`t User ($($_.GivenName) $($_.LastName)) will be skipped for processing!" | Out-File $log -append
    }
    $i++
  }
  "--------------------------------------------" + "`r`n" | Out-File $log -append
}

#-----------------------------------------------------------
#FUNCTION Delete users that doesn't exist anymore in the CSV
#-----------------------------------------------------------

Function Delete-Users
{
  $ADUserParams=@{ 
   'Searchbase' = 'OU=DB,DC=DMDevo04,DC=local' 
   'Searchscope'= 'Subtree'
   'Filter' = '*' 
   'Properties' = 'SAMAccountname' 
  }
 
  $SelectParams=@{ 
   'Property' = 'SAMAccountname' 
  } 
 
  # Getting a list of users from the specified OU (Searchbase)
  $listUsers = get-aduser @ADUserParams | select-object @SelectParams | ForEach{
    $userexist = $false
    $user = $_.SAMAccountname
    # Importing the CSV
    Import-CSV $newpath -Encoding UTF8 | ForEach-Object{
      $replaceName = $_.Lastname.Replace(".","")
      If($replaceName.length -lt 4)
      {
        $lastname = $replaceName
      }
      Else
      {
        $lastname = $replaceName.substring(0,4)
      }
      $replaceFirstName = $_.GivenName.Replace(".","")
      If($replaceFirstName.length -lt 2)
      {
        $givenname = $replaceFirstName
      }
      Else
      {
        $givenname = $replaceFirstName.substring(0,2)
      }

      $lastname = $lastname.substring(0,1).toupper()+$lastname.substring(1).tolower()   
      $sam = $lastname + $givenname.ToUpper()
      if ($sam -eq $user)
      {
        # Setting userexist var if the user exist.
        $userexist = $true
      }
    }
    # If the var userexist is set to false, removing the user.
    if ($userexist -eq $false)
    {
      Remove-ADUser $user -Confirm:$false
      Write-Host "[INFO]`t User $user doesn't exist, deleting!`r`n"
      "[INFO]`t User $user doesn't exist, deleting!" | Out-File $log -append
    }
  }
}

Write-Host "STARTED SCRIPT`r`n"
Start-Commands
Write-Host "STOPPED SCRIPT"

Write-Host "Press any key to continue ..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
