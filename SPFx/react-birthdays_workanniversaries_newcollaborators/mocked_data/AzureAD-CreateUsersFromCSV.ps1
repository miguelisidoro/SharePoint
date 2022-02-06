<#
//-----------------------------------------------------------------------

//     Copyright (c) {charbelnemnom.com}. All rights reserved.

//-----------------------------------------------------------------------

.SYNOPSIS
Create Azure AD User Account.

.DESCRIPTION
Azure AD Bulk user creation and assign the new users to an Azure AD group.

.NOTES
File Name : Invoke-AzureADBulkUserCreation.ps1
Author    : Charbel Nemnom
Version   : 1.7
Date      : 27-February-2018
Update    : 30-July-2019
Requires  : PowerShell Version 3.0 or above
Module    : AzureAD Version 2.0.0.155 or above
Product   : Azure Active Directory

.LINK
To provide feedback or for further assistance please visit:
https://charbelnemnom.com

.EXAMPLE-1
./Invoke-AzureADBulkUserCreation -FilePath <FilePath> -Credential <Username\Password> -Verbose
This example will import all users from a CSV File and then create the corresponding account in Azure Active Directory.
The user will be asked to change his password at first log on.

.EXAMPLE-2
./Invoke-AzureADBulkUserCreation -FilePath <FilePath> -Credential <Username\Password> -AadGroupName <AzureAD-GroupName> -Verbose
This example will import all users from a CSV File and then create the corresponding account in Azure Active Directory.
The user will be a member of the specified Azure AD Group Name.
The user will be asked to change his password at first log on.
#>

[CmdletBinding()]
Param(
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = 'Specify the path of the CSV file')]
    [Alias('CSVFile')]
    [string]$FilePath="Mocked Data Azure AD4.csv",
    [Parameter(Position = 1, Mandatory = $false, HelpMessage = 'Specify Credentials')]
    [Alias('Cred')]
    [PSCredential]$Credential,
    #MFA Account for Azure AD Account
    [Parameter(Position = 2, Mandatory = $false, HelpMessage = 'Specify if account is MFA enabled')]
    [Alias('2FA')]
    [Switch]$MFA
)

function Test-DomainExistsInAad {
      param(
             [Parameter(mandatory=$true)]
             [string]$DomainName
       )

       return $true
}

Try {
    $CSVData = @(Import-CSV -Path $FilePath -ErrorAction Stop)
    Write-Verbose "Successfully imported entries from $FilePath"
    Write-Verbose "Total no. of entries in CSV are : $($CSVData.count)"
} 
Catch {
    Write-Verbose "Failed to read from the CSV file $FilePath Exiting!"
    Break
}

Import-Module -Name AzureAD -ErrorAction Stop -Verbose:$false | Out-Null

Try {
    Write-Verbose "Connecting to Azure AD..."
    if ($MFA) {
        Connect-AzureAD -ErrorAction Stop | Out-Null
    }
    Else {
        Connect-AzureAD -Credential $Credential -ErrorAction Stop | Out-Null
    }
}
Catch {
    Write-Verbose "Cannot connect to Azure AD. Please check your credentials. Exiting!"
    Break
}

$CheckedDomains = @{}

Foreach ($Entry in $CSVData) {
    # Verify that mandatory properties are defined for each object
    $DisplayName = $Entry.FirstName + " " + $Entry.LastName
    $MailNickName = $Entry.EmailNickName
    $UserPrincipalName = $Entry.Email
    $Password = $Entry.Password
    
    If (!$DisplayName) {
        Write-Warning '$DisplayName is not provided. Continuing to the next record'
        Continue
    }

    If (!$MailNickName) {
        Write-Warning '$MailNickName is not provided. Continuing to the next record'
        Continue
    }

    If (!$UserPrincipalName) {
        Write-Warning '$UserPrincipalName is not provided. Continuing to the next record'
        Continue
    }

    If (!$Password) {
        Write-Warning "Password is not provided for $DisplayName in the CSV file!"
        $Password = Read-Host -Prompt "Enter desired Password" -AsSecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $Password
        $PasswordProfile.EnforceChangePasswordPolicy = 1
        $PasswordProfile.ForceChangePasswordNextLogin = 1
    }
    Else {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $Password
        $PasswordProfile.EnforceChangePasswordPolicy = 1
        $PasswordProfile.ForceChangePasswordNextLogin = 1
    }   

    #Verify that the domain is registered in AAD
    $domain = $UserPrincipalName.SubString($UserPrincipalName.IndexOf("@") + 1)
    $domainExists = $False
    if (!$CheckedDomains.ContainsKey($domain))
    {
        $CheckedDomains.Add($domain, (Test-DomainExistsInAad($domain)))
    }
    $domainExists = $CheckedDomains[$domain];

    if(!$domainExists)
    {
        Write-Warning "Domain for user $UserPrincipalName is not registered in Azure AD. Continuing to next user."
        Continue
    }
    
    #See if the user exists.
    Try{
        $ADuser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'"
        }
    Catch{}

    #If so then movea along, otherwise create the user.
    If ($ADuser)
    {
        Write-Verbose "$UserPrincipalName already exists. User will be added to group if specified."
    }
    Else
    {

        Try {    
            New-AzureADUser -DisplayName $DisplayName `
                -GivenName $Entry.FirstName `
                -Surname $Entry.LastName `
                -AccountEnabled $true `
                -MailNickName $MailNickName `
                -UserPrincipalName $UserPrincipalName `
                -PasswordProfile $PasswordProfile `
                } 
        Catch {
            Write-Error "$DisplayName : Error occurred while creating Azure AD Account. $_"
            Continue
        }

        #Make sure the user exists now.
        Try{
            $ADuser = Get-AzureADUser -Filter "userPrincipalName eq '$UserPrincipalName'"
        }
        Catch{
            Write-Warning "$DisplayName : Newly created account could not be found.  Continuing to next user. $_"
            Continue
        }

        Write-Verbose "$DisplayName : AAD Account is created successfully!"     
    }    
}
