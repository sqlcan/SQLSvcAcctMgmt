#This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
#THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
#INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
#We grant you a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute
#the object code form of the Sample Code, provided that you agree:
#(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded;
#(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and
#(iii) to indemnify, hold harmless, and defend Us and our suppliers from and against any claims or lawsuits, 
#      including attorneys' fees, that arise or result from the use or distribution of the Sample Code.
#
# Developed By: Mohit K. Gupta (mogupta@microsoft.com)
# Last Updated: May 4, 2017
#      Version: 3.0

param
(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $ComputerName,
    [Parameter(Mandatory=$true)] [String] $ServiceAccountName,
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
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceMode -Value "Unknown"
        $SvcObj | Add-Member -MemberType NoteProperty -Name ServiceState -Value "Unknown"        
        $SvcObj | Add-Member -MemberType NoteProperty -Name OperationStatus -Value "Unknown"

        return $SvcObj
    }

    Write-Verbose "$(Get-Date) Updating SQL Server Services' Service Accounts' Passwords"
    
    if (!($WhatIf) -and  (($ServiceAccountOldPassword -eq '') -or ($ServiceAccountNewPassword -eq '')))
    {
        Write-Error "Must supply both ServiceAccountOldPassword and ServiceAccountNewPassword."
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

        #Only target services that have service account name set to passed in parameters.        
        $FilteredServices = $SMOWmiserver.Services | Where {$_.ServiceAccount -eq $ServiceAccountName}
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
        $SQLService.ServiceState = $Service.ServiceState

        if (($Service.StartMode -eq 'Auto') -or (($Service.StartMode -eq 'Manual') -and ($Service.ServiceState -eq 'Running')))
        {
            try
            {
                if (!($WhatIf))
                {
                    Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method ChangePassword (SQLSMO)"
                    $Service.ChangePassword($ServiceAccountOldPassword, $ServiceAccountNewPassword)
                    $SQLService.OperationStatus = "Password Change Completed"
                }
                else
                {
                    Write-Verbose "$(Get-Date) What if: Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName)"
                    $SQLService.OperationStatus = "Dry Run"
                }
            }
            catch
            {
                $SQLService.OperationStatus = "Password Change Failed"
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
                    Write-Verbose "$(Get-Date) Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName) via Method Change (Win32_Service)"
                    $WMIServices = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName
                    $SQLSvc = $WMIServices | ? {$_.Name -Like "$($SQLService.ServiceName)"}
                    $Results = $SQLSvc.change($null,$null,$null,$null,$null,$null,$null,$ServiceAccountNewPassword,$null,$null)
                    $SQLService.OperationStatus = "Password Change Completed"

                }
                else
                {
                    Write-Verbose "$(Get-Date) What if: Changing Password on Target \\$($SQLService.ComputerName)\$($SQLService.ServiceName)"
                    $SQLService.OperationStatus = "Dry Run"
                }
            }
            catch
            {
                $SQLService.OperationStatus = "Password Change Failed"
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

