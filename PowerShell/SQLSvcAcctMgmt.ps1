<# 
.Synopsis
This script allows for DBA to manage the service account for SQL Services.

.Description
This solution provides the ability to manage service accounts for SQL Server Services.  The solution
allows you to update the password or change the service account.

This script can be called with following parameter combinations:
ComputerName, ServiceAccountName, ServiceAccountOldPassword, ServiceAccountNewPassword
ComputerName, ServiceAccountName, NewServiceAccountName, ServiceAccountNewPassword

If NewServiceAccountName and ServiceAccountOldPassword both are supplied then it is assumed service
account change is required and ServiceAccountOldPassword is ignored.

.Parameter ComputerName
Server name where SQL Server services are installed for which you wish to change the password.

.Parameter ServiceAccountName
Current service account assigned to the services.  The script finds the services based on this service
account.

.Parameter NewServiceAccountName
If you wish to change existing service account, you must supply ServiceAccountName and the
NewServiceAccountName name with ServiceAccountNewPassword.

.Parameter ServiceType
Specify the SQL Service you wish to update the password or service account for; current supported are:
All, SqlServer, SqlAgent, and AnalysisServer.

.Parameter ServiceAccountOldPassword
Required to change the existing password from current to new password.

.Parameter ServiceAccountNewPassword
Required to change the existing password to new password or to change existing service account to
new service account.

.Example 
.\SQLSvcAcctMgmt.ps1 -ComputerName Contoso -ServiceAccountName Contoso\SQLSvc -ServiceAccountOldPassword Password123 -ServiceAccountNewPassword P@ssword123
Change the service account for all SQL Services that have account Contoso\SQLSvc.

.Example
.\SQLSvcAcctMgmt.ps1 -ComputerName Contoso -ServiceAccountName Contoso\SQLSvc -ServiceAccountOldPassword Password123 -ServiceAccountNewPassword P@ssword123 -WhatIf
What services will be affected if we change the password for service acocunt.

.Example
.\SQLSvcAcctMgmt.ps1 -ComputerName Contoso -ServiceAccountName Contoso\SQLSvc -ServiceAccountOldPassword Password123 -ServiceAccountNewPassword P@ssword123 -Verbose
Change the service account for all SQL Services that have account Contoso\SQLSvc providing verbose output.

.Example 
.\SQLSvcAcctMgmt.ps1 -ComputerName Contoso -ServiceAccountName Contoso\SQLSvc -NewServiceAccountName Contoso\SQLSvcNew -ServiceAccountNewPassword P@ssword123
Change the SQL Service account from current to new account.

.Link 
https://github.com/sqlcan/SQLSvcAcctMgmt/
http://sqlcan.com/

.Notes
Date         Version     Author      Comments
------------ ----------- ----------- ----------------------------------------------------------------
2017.08.28   3.10.0000   mogupta     Adding functionality to handle service account changes.
2017.08.30   3.20.0000   mogupta     Added ability to handle service type.
#> 

param
(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $ComputerName,
    [Parameter(Mandatory=$true)] [String] $ServiceAccountName,
    [Parameter(Mandatory=$false)] [String] $ServiceType = 'All',
    [Parameter(Mandatory=$false)] [String] $NewServiceAccountName,
    [Parameter(Mandatory=$false)] [String] $ServiceAccountOldPassword,
    [Parameter(Mandatory=$false)] [String] $ServiceAccountNewPassword,
    [Parameter(Mandatory=$false)] [Switch] $WhatIf
)

BEGIN
{
    #Load the SqlWmiManagement assembly off of the DLL
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SqlWmiManagement") | Out-null

    $ServerAccessible = $true
    $ErrorActionPreference = “Stop”
    $SQLServices = @()

    function Get-SQLServiceObject
    {
        $SvcObj = New-Object -TypeName PSObject
        $SvcObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceName -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceType -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceMode -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceState -Value "Unknown"        
        $SvcObj | Add-Member -MemberType NoteProperty -Name OperationStatus -Value "Unknown"

        return $SvcObj
    }    
    
    if ($NewServiceAccountName -eq '')
    {
        Write-Verbose "$(Get-Date) Updating SQL Server Services' Service Accounts' Passwords"
    }
    else
    {
        Write-Verbose "$(Get-Date) Updating SQL Server Services' Service Account"
    }

    if (!($WhatIf))
    {
        # If NewServiceAccountName is not supplied assumed that it is password change request.
        # Therefore, user must supply both old and new service account password.
        if (($NewServiceAccountName -eq '') -and (($ServiceAccountOldPassword -eq '') -or ($ServiceAccountNewPassword -eq '')))
        {
            Write-Error "Must supply both ServiceAccountOldPassword and ServiceAccountNewPassword."
        }                
        elseif (($NewServiceAccountName -ne '') -and ($ServiceAccountNewPassword -eq ''))
        {
            Write-Error "Must supply both ServiceAccountNewPassword when NewServiceAccountName is supplied."
        }
    }

}

PROCESS
{

    try
    {

        $FilteredServices = $null
        $SQLService = Get-SQLServiceObject
        $SQLService.ComputerName = $ComputerName
        $OSName  = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName).Caption

        #Ideally I should be able to get the operating system name from Caption attribute.
        #However, some cases I have seen that return null value.  Therefore adding additional
        #logic to handle the issue.

        if ($OSName -eq '')
        {
            $OSVersion = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName).Version
            $OSVersion = $OSVersion.Version
            $OSVersionsSplits = $OSVersion.Split('.')
            [int]$OSMajorVersion = $OSVersionsSplits[0]
            [int]$OSMinorVersion = $OSVersionsSplits[1]
            $OSVersionNumber = ($OSMajorVersion * 10) + $OSMinorVersion 

            switch ($OSVersionNumber)
            {
                40 {$OSName = "Windows NT 4.0"}
                50 {$OSName = "Windows Server 2000"}
                52 {$OSName = "Windows Server 2003/R2"}
                60 {$OSName = "Windows Server 2008"}
                61 {$OSName = "Windows Server 2008 R2"}
                62 {$OSName = "Windows Server 2012"}
                63 {$OSName = "Windows Server 2012 R2"}
                100 {$OSName = "Windows Server 2016"}
                else {$OSName = "Unknown"}
            }
        }

        $SQLService.OperatingSystem = $OSName
        $SMOWmiserver = New-Object ('Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer') $ComputerName

        #These just act as some queries about the SQL Services on the machine you specified.
        #$SMOWmiserver.Services | Select name, type, ServiceAccount, DisplayName, Properties, StartMode, StartupParameters | Format-Table
        #($SMOWmiserver.Services) | GM

        if ($ServiceType -eq 'All')
        {
            #Only target services that have service account name set to passed in parameters.        
            $FilteredServices = $SMOWmiserver.Services | Where {$_.ServiceAccount -eq $ServiceAccountName}
        }
        else
        {
            #Only target services that have service account name set to passed in parameters.        
            $FilteredServices = $SMOWmiserver.Services | Where {$_.ServiceAccount -eq $ServiceAccountName -and $_.Type -eq $ServiceType }
        }
    }
    catch
    {
        $SQLService.OperationStatus = "WMI Error"
        Write-Verbose "$(Get-Date) WMI Call Failed on Target \\$($SQLService.ComputerName)"
        Write-Verbose "$(Get-Date) Exception Type: $($_.Exception.GetType().FullName)”
        Write-Verbose "$(Get-Date) Exception Message: $($_.Exception.Message)”
        $SQLServices += $SQLService
        return
    }

    if (($FilteredServices | Measure-Object).Count -eq 0)
    {
        Write-Verbose "$(Get-Date) No services found for service account [$ServiceAccountName] on Target \\$($SQLService.ComputerName)"
        $SQLService.OperationStatus = "No Service Found"
        $SQLServices += $SQLService
    }

    ForEach ($Service IN $FilteredServices)
    {
        # Services which are disabled cannot have their service account or password information updated.

        $SQLService.ServiceMode = $Service.StartMode
        $SQLService.ServiceName = $Service.Name
        $SQLService.ServiceType = $Service.Type
        $SQLService.ServiceState = $Service.ServiceState

        if (($Service.StartMode -eq 'Auto') -or (($Service.StartMode -eq 'Manual') -and ($Service.ServiceState -eq 'Running')))
        {
            try
            {
                if (!($WhatIf))
                {
                    if ($NewServiceAccountName -ne '')
                    {
                        Write-Verbose "$(Get-Date) Changing Service Account on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method ChangePassword (SQLSMO) *Service Restarted*"
                        $Service.SetServiceAccount($NewServiceAccountName, $ServiceAccountNewPassword)
                        $SQLService.OperationStatus = "Service Account Changed"
                    }
                    else
                    {
                        Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method ChangePassword (SQLSMO)"
                        $Service.ChangePassword($ServiceAccountOldPassword, $ServiceAccountNewPassword)
                        $SQLService.OperationStatus = "Password Change Completed"
                    }
                }
                else
                {
                    if ($NewServiceAccountName -ne '')
                    {
                        Write-Verbose "$(Get-Date) What if: Changing Service Account on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method ChangePassword (SQLSMO) *Restart Required*"                        
                        $SQLService.OperationStatus = "Dry Run - Service Account Change"
                    }
                    else
                    {
                        Write-Verbose "$(Get-Date) What if: Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method ChangePassword (SQLSMO)"                        
                        $SQLService.OperationStatus = "Dry Run - Password Change"
                    }                    
                }
            }
            catch
            {
                if ($NewServiceAccountName -ne '')
                {
                    $SQLService.OperationStatus = "Service Account Change Failed"
                }
                else
                {
                    $SQLService.OperationStatus = "Password Change Failed"
                }
                Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) Failed!"
                Write-Verbose "$(Get-Date) Exception Type: $($_.Exception.GetType().FullName)”
                Write-Verbose "$(Get-Date) Exception Message: $($_.Exception.Message)”
            }
        }
        elseif (($Service.StartMode -eq 'Manual') -and ($Service.ServiceState -ne 'Running'))
        {
            # If a service is found to be configured in Manual mode and not running; assumption is being made
            # it is clustered service.  Alternatively we can check the AdvancedPrperties "CLUSTERED". However,
            # I did not have chane to test this against older versions of SQL Server.
            #
            # Method being used below to update SQL Server Services information via the Win32_Services
            # is only supported for Password Change.
            #
            # THIS METHOD IS NOT RECOMMENDED TO CHANGE SERVICE ACCOUNT.
            #
            # Below method is not required however, due to KB972387 I am using the below method.
            # I do not like this method as the Change method of Win32_Service does not validate the password
            # therefore, it will change the password to even an incorrect value and will not cause any errors.

            try
            {
                if (!($WhatIf))
                {
                    if ($NewServiceAccountName -ne '')
                    {
                        Write-Verbose "$(Get-Date) Changing Service Account on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method Change (Win32_Service)"
                        $WMIServices = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName
                        $SQLSvc = $WMIServices | ? {$_.Name -Like "$($SQLService.ServiceName)"}
                        $Results = $SQLSvc.change($null,$null,$null,$null,$null,$null,$NewServiceAccountName,$ServiceAccountNewPassword,$null,$null)
                        $SQLService.OperationStatus = "Service Account Completed"
                    }
                    else
                    {
                        Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method Change (Win32_Service)"
                        $WMIServices = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName
                        $SQLSvc = $WMIServices | ? {$_.Name -Like "$($SQLService.ServiceName)"}
                        $Results = $SQLSvc.change($null,$null,$null,$null,$null,$null,$null,$ServiceAccountNewPassword,$null,$null)
                        $SQLService.OperationStatus = "Password Change Completed"
                    }
                }
                else
                {
                    if ($NewServiceAccountName -ne '')
                    {
                        Write-Verbose "$(Get-Date) What if: Changing Service Account on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method Change (Win32_Service)"                        
                    }
                    else
                    {
                        Write-Verbose "$(Get-Date) What if: Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method Change (Win32_Service)"                        
                    }
                    $SQLService.OperationStatus = "Dry Run"
                }
            }
            catch
            {
                if ($NewServiceAccountName -ne '')
                {
                    $SQLService.OperationStatus = "Service Account Change Failed"
                }
                else
                {
                    $SQLService.OperationStatus = "Password Change Failed"
                }                
                Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) Failed!"
                Write-Verbose "$(Get-Date) Exception Type: $($_.Exception.GetType().FullName)”
                Write-Verbose "$(Get-Date) Exception Message: $($_.Exception.Message)”
            }
        }
        else
        {
            Write-Verbose "$(Get-Date) Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) Skipped, Service Disabled"
            $SQLService.OperationStatus = "Skipped Service Disabled"
        }


        $SQLServices += $SQLService        

        $SQLService = Get-SQLServiceObject
        $SQLService.ComputerName = $ComputerName
        $SQLService.OperatingSystem = $OSName

    }

}
END
{
    Write-Verbose "$(Get-Date) Script Completed."
    return $SQLServices
}

