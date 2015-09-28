

<#
.Synopsis
   Export Mailbox Permissions for use in importing permissions to other mailbox
.DESCRIPTION
   Export Mailbox Permissions for use in importing permissions to other mailbox
.EXAMPLE
   Export-D2MTMailboxPermissions -Identity HRMail@contoso.com -verbose
.EXAMPLE
	$Maillist = "HRMail@contoso.com","FINMail@contoso.com";Export-D2MTMailboxPermission -Identity $MailList -verbose
	
	Exporting multiple Mailboxes
   
.EXAMPLE
	$Maillist = "HRMail@contoso.com","FINMail@contoso.com";Export-D2MTMailboxPermission -Identity $MailList -verbose -ExportCSV
	Exporting multiple Mailboxes and outputting to CSV

.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Export-D2MTMailboxPermission
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Identity of Mailbox
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("Mailbox")]
		[string[]] 
        $Identity,

        # Export to CSV File
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [switch]
        $ExportCSV

       
    )

    Begin
    {
		$exportedPerms = $null
    }
    Process
    {
        foreach ($i in $identity)
		{
			Write-Verbose -Message "Retrieving non-inherited Mailbox permissions for $i"

			$ADperms = $null
			$MBperms = $null
			try
			{
				$mb = get-mailbox $i
				$MailboxEmail = $mb.WindowsEmailAddress
				$MBperms = $mb|get-mailboxpermission
				$ADperms = $mb|get-ADpermission
			}
			catch
			{
				$ErrorMessage = $_.Exception.Message
				$FailedItem = $_.Exception.ItemName
				Write-Error -Message "Error retrieving Permissions for Mailbox $i - $ErrorMessage - $FailedItem"
			}

			if ($null -ne $MBperms)
			{
				
                $ExportedPerms += Format-D2MTExportedMailboxPermission -MailboxEmail $MailboxEmail -Permissions $MBperms -PermissionType Mailbox
                
			}
			

			if ($null -ne $ADperms)
			{
				
                $ExportedPerms += Format-D2MTExportedMailboxPermission -MailboxEmail $MailboxEmail -Permissions $ADperms -PermissionType Recipient
                
			}

		}
    }
    End
    {
        Write-Output -InputObject $exportedPerms

        if ($ExportCSV)
        {
			$ResultsCSV = ("ExportedMailboxPermissions_"+(((Get-date -Format u).Replace(':','-').Replace(' ','_'))+".csv"))
			write-verbose -Message "Saving Results to CSV File $ResultsCSV"
			$exportedPerms|Export-Csv $resultsCSV -NoTypeInformation
			Write-Verbose -Message "Complete!"
        }
    }
}

<#
.Synopsis
   Format Exported Mailbox Permissions to prepare for Import
.DESCRIPTION
	Format Exported Mailbox Permissions to prepare for Import
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Format-D2MTExportedMailboxPermission
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Mailbox Email Address
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]
        $MailboxEmail,
        # Permissions Collection
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Permissions,
		# Permission Type
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
		[ValidateSet("Mailbox", "Recipient")]
        $PermissionType




    )

    Begin
    {
    }
    Process
    {
		switch ($PermissionType)

		{
			"Mailbox" {

				$filteredPerms = $Permissions|Where-Object {($_.IsInherited -eq $false) -and ($_.User -notlike “NT AUTHORITY\SELF”)}
                
			}

			"Recipient" {

				$filteredPerms = $Permissions|Where-Object {($_.ExtendedRights -like “*Send-As*”) -and ($_.IsInherited -eq $false) -and ($_.User -notlike “NT AUTHORITY\SELF”)}
			}
		}
        
        

				$filteredPerms|ForEach-Object{

                    $AccessRights = $null

                    if ($PermissionType -eq 'Mailbox')
                    {
                        $AccessRights = $_.AccessRights -join ";"
                    }
                    else
                    {
                        $AccessRights = ($_.ExtendedRights -join ";").Replace("Send-As","SendAs")
                    }
					$UserResolved= $null
                    $UserEmail=$null
                    $UserDomain=$null
                    $UserResolved=$null
                    $UserName = $null
					$UserDisplayName = $null
					
                   

                   if ($_.user -like "*\*")
                   {
                       $UserDomain = $_.user.split('\')[0]
                       $UserName = $_.user.split('\')[1]

                       $Addomain = get-addomain -server $UserDomain|Select-Object -expand dnsRoot

                       $Adobject = $null
                       Write-verbose -Message "Querying $ADdomain for $UserName"
                       $AdObject = Get-ADObject -filter {Samaccountname -eq $UserName} -Server $ADDomain -Properties mail,displayname

                      

                       if ($null -notlike $AdObject)
                       {
                        $UserResolved = $true

                        $userEmail = $Adobject.mail
                       $userType = $adObject.objectclass
					   $userDisplayName = $AdObject.displayname

                        if ($null -notlike $UserEmail)
                        {
                            $userEligible = $true
                        }
                        else
                        {
                            $userEligible = $false
                        }
                       }
                       else
                       {
                        $UserResolved= $false
                        $userEligible = $false
                       }
                   }
                   else
                   {
                         $UserResolved= $false
                        $userEligible = $false
                   }

					[pscustomobject][ordered]@{Identity=$_.Identity;MailBoxEmail=$MailboxEmail;PermissionType=$PermissionType;User=$_.User;UserDisplayName=$UserDisplayName;UserEmail=$UserEmail;UserName=$UserName;UserDomain=$UserDomain;UserType=$userType;UserResolved=$UserResolved;UserEligible=$UserEligible;AccessRights=$AccessRights}
					
				}
    }
    End
    {
    }
}

<#
.Synopsis
   Grant Mailbox and Recipient permissions on Exchange Online Mailbox
.DESCRIPTION
   Grant Mailbox and Recipient permissions on Exchange Online Mailbox
.EXAMPLE
   Grant-D2MTMailboxPermission -MailboxEmail test@contoso.com -UserEmail bob.smith@contoso.com -PermissionType Mailbox -AccessRights FullAccess
.EXAMPLE
   Grant-D2MTMailboxPermission -MailboxEmail test@contoso.com -UserEmail bob.smith@contoso.com -PermissionType Recipient -AccessRights SendAs
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Grant-D2MTMailboxPermission
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # SMTP Email address of Mailbox
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("Email")]
		[string] 
        $MailboxEmail,

        # Access Rights Collection to be assigned to Trustee
        [Parameter(ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]
        $AccessRights,

		# SMTP Emaill Address of user/group being added access for
		[Parameter(ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
		[Alias("Trustee")]
        [string]
        $UserEmail,

        # Permission Type being added
        [Parameter(ParameterSetName='Parameter Set 1')]
        [String]
		[ValidateSet("Mailbox", "Recipient")]
        $PermissionType
    )

    Begin
    {
    }
    Process
    {
        
		
		
			switch ($PermissionType)
			{

				"Mailbox" {

					Write-Verbose -Message "Adding Recpient $AccessRights for $UserEmail to $MailboxEmail Mailbox"
					if ($pscmdlet.ShouldProcess($MailboxEmail, "Add $AccessRights for $userEmail"))
					{
						Add-MailboxPermission -Identity $MailboxEmail -AccessRights $AccessRights -User $UserEmail
					}
				}

				"Recipient" {

					Write-Verbose -Message "Adding Mailbox $AccessRights for $UserEmail to $MailboxEmail Mailbox"
					if ($pscmdlet.ShouldProcess($MailboxEmail, "Add $AccessRights for $userEmail"))
					{
						Add-RecipientPermission -Identity $MailboxEmail -AccessRights $AccessRights -Trustee $UserEmail -Confirm:$false
					}
				}
			}
			
     
    }
    End
    {
    }
}

<#
.Synopsis
    Import Mailbox Permissions from Exported Permissions into MT Mailbox
.DESCRIPTION
    Import Mailbox Permissions from Exported Permissions into MT Mailbox
.EXAMPLE
   Import-D2MTMailboxPermission -File .\ExportedPermissions.csv -Domain contoso
.NOTES
   Specify short domain names of ACLs to import from exported data
#>
function Import-D2MTMailboxPermission
{
    
	  [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # File name of the CSV of Exported Mailbox Permissions
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $File,

        # Domain Names to include when filtering permissions to apply
		[Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string[]]
        $Domain
    )

    Begin
    {
		if (Test-Path -Path $File)
		{
			$csv = Import-Csv -Path $File
		}
		else
		{
			Write-Error -Message "$file Not Found!  Please check path for file exists and is not locked."
		}
    }
    Process
    {
		foreach ($dom in $domain)
		{
			$filteredCsv = $csv|Where-Object {$_.UserEligible -eq $true -and $_.UserDomain -EQ $dom}

			$filteredCsv|ForEach-Object{

				Grant-D2MTMailboxPermission -MailboxEmail $_.MailboxEmail -AccessRights $_.AccessRights -UserEmail $_.UserEmail -PermissionType $_.PermissionType
			}
		}
    }
    End
    {
    }
}

<#
.Synopsis
   Resolve SIDS on Migrated Mailbox to Recipients
.DESCRIPTION
   Enumerates unresolved SIDS on the mailbox, and attempts to translate them to existing security principal
.EXAMPLE
   Resolve-D2MTMigratedMailboxPermissions -Identity "test@contoso.com"
.EXAMPLE
      Resolve-D2MTMigratedMailboxPermissions -Identity "test@contoso.com","test2@contoso.com"
.EXAMPLE
	  $MailBoxes = get-content mailboxes.txt
      Resolve-D2MTMigratedMailboxPermissions -Identity $MailBoxes
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   Load the Exchange Cmdlets in your PSSession before attempting to utilize this module so that the cmdlets are availiable
   Not all SIDS will be resolved since they can be from defunct or disconnected domains
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
  Resolve SIDS on Migrated Mailbox to Recipients
#>
function Resolve-D2MTMigratedMailboxPermission
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param
    (
        # Email Address of the Mailbox to Resolve Permissions
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [Alias("Mailbox")]
		[string[]]
        $Identity
    )

    Begin
    {
    }
    Process
    {
		foreach($mbEmail in $Identity)
		{
			
			$mbMailbox = get-mailbox $mbEmail
			$mbPermissions = $mbMailbox|get-mailboxpermission

			$needsResolve = $mbPermissions|Where-Object user -like 's-1-5-21*'

			$needsResolve|ForEach-Object{
	
				$resolvedUser = $null
				$userToResolve = $_.user
				$UserIsResolved = $false
				Write-Verbose -Message "Attempting to resolve $userToResolve"
				$checkSid = New-Object System.Security.Principal.SecurityIdentifier($userToResolve)
				try {
					$resolvedUser = $checkSID.Translate([System.Security.Principal.NTAccount])
					}
				catch{}

				 if ($resolvedUser -like "*\*")
                   {

					   Write-Verbose -Message "Resolved $userToResolve to $resolvedUser"
                       $UserDomain = ($resolvedUser.value.split('\')[0]).ToString()
                       $UserName = $resolvedUser.value.split('\')[1]

					  

                       $Adobject = $null
                       Write-Verbose -Message "Querying $UserDomain for $UserName"
                       $AdObject = Get-ADObject -filter {Samaccountname -eq $UserName} -Server $UserDomain -Properties mail,displayname

                      

                       if ($null -notlike $AdObject)
                       {
							$UserIsResolved = $true

							$userEmail = $Adobject.mail
							$userType = $adObject.objectclass
							$userDisplayName = $AdObject.displayname
						}

					   if (($userIsResolved -eq $true) -and ($null -notlike $userEMail))
					   {
						   $Accessrights = $_.AccessRights

						    if ($pscmdlet.ShouldProcess($mbEmail, "Add $AccessRights for $userEmail"))
							{
        
								Import-D2MTMailboxPermissions -MailboxEmail $mbEmail -AccessRights $AccessRights -UserEmail $userEmail -PermissionType Mailbox
							}
					   }


			}
				else
				{
					Write-Verbose -Message "Unable to Resolve $userToResolve.  Access will not be added."
				}

		}


       
    }
	}
    End
    {
    }
}