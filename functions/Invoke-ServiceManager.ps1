function Invoke-ServiceManager()
<#
.SYNOPSIS
    Manages Windows services by starting, stopping, or restarting them, and optionally sets them to start automatically.

.DESCRIPTION
    The Invoke-ServiceManager function processes one or more Windows services specified by name. It performs the requested operation (Start, Stop, or Restart) on each service and can optionally set the service to start automatically. The function returns an array of objects containing the status and details of each processed service.

.PARAMETER serviceNames
    An array of service names to process.

.PARAMETER Operation
    The operation to perform on the services. Valid values are 'Start', 'Stop', or 'Restart'.

.PARAMETER AutoStart
    If specified, sets the service to start automatically.

.EXAMPLE
    Invoke-ServiceManager -serviceNames 'wuauserv','bits' -Operation Start -AutoStart

    Starts the 'wuauserv' and 'bits' services and sets them to start automatically.

.NOTES
    Author: MahmoudZ
    Requires: PowerShell 5.1 or later
#>
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$serviceNames,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Start', 'Stop', 'Restart', 'Status')]
        [string]$Operation,
        [switch]$AutoStart
    )    

    #region Print verbose log of input parameters
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Processing $($serviceNames.Count) service(s)."
    Write-Verbose "[$functionName] Service Names: $($serviceNames -join ', ')"
    Write-Verbose "[$functionName] Operation: $Operation"
    Write-Verbose "[$functionName] AutoStart: $AutoStart"
    #endregion
    
    $services = @()
    $serviceObject = [PSCustomObject]@{}
    foreach ($serviceName in $serviceNames)
    {
        Write-Verbose "[$functionName] Processing service: $serviceName"
        $serviceObject = [ordered] @{
            Name      = $serviceName
            Operation = $Operation
            autoStart = $AutoStart
        }
        try
        {
            $service = Get-Service -Name $serviceName
            $serviceObject += @{
                DisplayName        = $service.DisplayName         
                ServiceType        = $service.ServiceType
                StartType          = $service.StartType
                Status             = $service.Status
                DependentServices  = $service.DependentServices | ForEach-Object { $_.Name }
                ServicesDependedOn = $service.ServicesDependedOn | ForEach-Object { $_.Name }
                NotFound           = $false
            }
        }
        catch
        {
            Write-Warning "[$functionName] Service '$serviceName' not found: $_"
            $serviceObject += @{
                NotFound = $true
            }
            continue
        }   
        
        #region Perform the desired operation on each service
        switch ($Operation)
        {
            'Start'
            {
                #Check if the services is already running
                if ($service.Status -eq 'Running')
                {
                    Write-Verbose "[$functionName] Service '$serviceName' is already running."
                    $newStatus = 'No Change'
                }
                else
                {
                    try
                    {
                        Start-Service -Name $serviceName -ErrorAction Stop
                        Write-Verbose "[$functionName] Started service: $serviceName"
                        $newStatus = (Get-Service -Name $serviceName).Status
                    }
                    catch
                    {
                        Write-Error "[$functionName] Failed to start service '$serviceName': $_"
                        $newStatus = 'Failed'
                    }
                }
            }
            'Stop'
            {
                #Check if the service is already stopped
                if ($service.Status -eq 'Stopped')
                {
                    Write-Verbose "[$functionName] Service '$serviceName' is already stopped."
                    $newStatus = 'No Change'
                }
                else
                {
                    try
                    {
                        Stop-Service -Name $serviceName -ErrorAction Stop
                        Write-Verbose "[$functionName] Stopped service: $serviceName"
                        $newStatus = (Get-Service -Name $serviceName).Status
                    }
                    catch
                    {
                        Write-Error "[$functionName] Failed to stop service '$serviceName': $_"
                        $newStatus = 'Failed'
                    }
                }
            }
            'Restart'
            {
                try
                {
                    Restart-Service -Name $serviceName -ErrorAction Stop
                    Write-Verbose "[$functionName] Restarted service: $serviceName"
                    $newStatus = (Get-Service -Name $serviceName).Status
                }
                catch
                {
                    Write-Error "[$functionName] Failed to restart service '$serviceName': $_"
                    $newStatus = 'Failed'
                }
            }
            'Status'
            {
                # Report current status without changing service state
                $newStatus = $service.Status
                Write-Verbose "[$functionName] Service '$serviceName' current status: $newStatus"
            }
            default
            {
                Write-Verbose "[$functionName] Unknown operation: $Operation"
                $newStatus = 'Unknown'
            }
        }
        
        # If AutoStart is specified, set the service to start automatically
        if ($AutoStart)
        {
            if ($service.StartType -eq 'Automatic')
            {
                Write-Verbose "[$functionName] Service '$serviceName' is already set to start automatically."
                $serviceObject += @{
                    AutoStartSet = $true
                    newAutoStart = 'No Change'
                }
            }
            else
            {
                try
                {
                    Set-Service -Name $serviceName -StartupType Automatic -ErrorAction Stop
                    $updated = Get-Service -Name $serviceName
                    Write-Verbose "[$functionName] Set service to start automatically: $serviceName (Now: $($updated.StartType))"
                    $serviceObject += @{
                        AutoStartSet = $true
                        newAutoStart = $updated.StartType
                    }
                }
                catch
                {
                    Write-Error "[$functionName] Failed to set service '$serviceName' to start automatically: $_"
                    $serviceObject += @{
                        AutoStartSet = $false
                        newAutoStart = $service.StartType
                    }
                }
            }
        }
        if (-not $serviceObject.notFound)
        {
            # Add the processed service status to the service object
            $serviceObject += @{
                newStatus = $newStatus
            }
        }
        $services += $serviceObject
    }
    #endregion
    Write-Verbose "[$functionName] Processed $($services.Count) service(s)."
    write-log -LogFile $LogFile -Module $functionName -Message "Processed $($services.Count) service(s)." -LogLevel "Information"
    return @($services)
}
