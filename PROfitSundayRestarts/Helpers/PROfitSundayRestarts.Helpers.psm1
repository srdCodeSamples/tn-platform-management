Import-Module "$PSScriptRoot\PSFunctions_v3.1.psm1"
#region - Restart web service functions
function Restart-TNWebService {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$AppPoolName,
        [Parameter(Mandatory = $True)]
        [array]$Servers,
        [Parameter(Mandatory = $True)]
        [int]$Pause,
        [Parameter(Mandatory = $False)]
        [string]$HMUri = $null,
        [Parameter(Mandatory = $False)]
        [string]$HMProtocol = 'http',
        [Parameter(Mandatory = $False)]
        [switch]$AbortOnHMfailure,
        [Parameter(Mandatory = $False)]
        [string]$LogPath = $null,
        [Parameter(Mandatory = $False)]
        [string]$Thread = [System.String]::Empty
    )

    begin {
        $isLoggingEnabled = ![String]::IsNullOrEmpty($LogPath)
    }

    process {
        foreach ($srv in $Servers) {
            if ($isLoggingEnabled) {
                Write-CSTMLog -FilePath $LogPath -Type "INFO" -Message "Recycling $AppPoolName on server $srv" -Thread $Thread
            }
            try {
                $err = @()
                Invoke-Command -ComputerName $srv -ScriptBlock {Restart-WebAppPool -Name $Args[0]} -ArgumentList $AppPoolName -ErrorVariable +err
            }
            catch {
                if ($isLoggingEnabled) {
                    Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message "$($Error[0])" -Thread $Thread
                }
                $Error.Clear()
            }
            finally {
                if ($err.count -gt 0) {
                    if ($isLoggingEnabled) {
                        Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message "$err" -Thread $Thread
                    }
                    $err = @()
                }
            }
            if ($isLoggingEnabled) {
                Write-CSTMLog -FilePath $LogPath -Type "DEBUG" -Message "Waiting for $Pause seconds after recycle." -Thread $Thread
            }
            Start-Sleep -Seconds $Pause
            if (![String]::IsNullOrEmpty($HMUri)) {
                if ($isLoggingEnabled) {
                    Write-CSTMLog -FilePath $LogPath -Type "DEBUG" -Message "Making healthcheck with URL $HMProtocol`://$srv/$HMUri" -Thread $Thread
                }
                try {
                    $hmResult = 0
                    $err = @()
                    $hmResult = (Invoke-WebRequest -Uri "$HMProtocol`://$srv/$HMUri" -MaximumRedirection 0 -ErrorVariable +err).StatusCode
                }
                catch {
                    if ($Error.count -gt 0 -and $isLoggingEnabled) {
                        Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message $Error[0].Exception.Message.ToString() -Thread $Thread
                    }
                    $Error.Clear()
                }
                finally {
                    if($err.Count -gt 0 -and $isLoggingEnabled) {
                        Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message $err.InnerException.Message.ToString() -Thread $Thread
                    }
                    $err = @()
                }
                if ($hmResult -ne 200) {
                    if ($isLoggingEnabled) {
                        Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message "Health Check for server $srv failed with status code: $hmResult" -Thread $Thread
                    }
                    if ($AbortOnHMfailure.IsPresent) {
                        if ($isLoggingEnabled) {
                            Write-CSTMLog -FilePath $LogPath -Type "ERROR" -Message "Aborting $AppPoolName recycles due to HM failure on $srv" -Thread $Thread
                        }
                        throw "Aborting $AppPoolName recycles due to HM failure on $srv"
                    }
                }
                else {
                    if ($isLoggingEnabled) {
                        Write-CSTMLog -FilePath $LogPath -Type "DEBUG" -Message "Health Check for server $srv passed with status code: $hmResult" -Thread $Thread
                    }
                }
            }
            else {
                if ($isLoggingEnabled) {
                    Write-CSTMLog -FilePath $LogPath -Type "DEBUG" -Message "No health check after recycle is configured." -Thread $Thread
                }
            }
        }
    }
    end {
    }
}
#endregion

#region- Build TN components
function Get-TNDscComponents {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True)]
        [string]$DscEnvConfigPath,
        [Parameter(Mandatory = $True)]
        [string]$DscAppPath,
        [Parameter(Mandatory = $False)]
        [string]$LogPath
    )

    begin {
        $isLoggingEnabled = ![String]::IsNullOrEmpty($LogPath)
    }

    process {
        $baseDir = [System.IO.Path]::GetDirectoryName($DscEnvConfigPath).ToString()
        $file = [System.IO.Path]::GetFileName($DscEnvConfigPath).ToString()
        $dscNodes = Import-LocalizedData -BaseDirectory $baseDir -FileName $file
        $dscApps = Import-Csv -Path "$DscAppPath"
        $tnSvcs = @()
        foreach ($node in $dscNodes.Allnodes.where({($PSItem.Role -notin @('RootSite')) -and ($PSitem.NodeName -ne '*')})) {
            foreach ($cmp in $node.Component) {
                $currentSvc = $null
                $currentSvc = $tnSvcs.where({$PSitem.Component -eq $cmp -and $PSitem.EnvType -eq $node.Type})[0]
                if ($currentSvc -eq $null) {
                    $currentSvc = New-Object -TypeName PSObject -Property @{
                        'Component' = $cmp
                        'EnvType' = $node.Type
                        'Servers' = @($node.NodeName)
                        'Type' = $node.Role
                        'SvcName' = [System.String]::Empty
                        'HealthCheck' = [System.String]::Empty
                        'HMProtocol' = 'http'
                        'Priority' = 0
                    }
                    if ($currentSvc.Type -eq 'App') {
                        $currentSvc.SvcName = $dscApps.Where({$PSItem.Component -eq $currentSvc.Component -and $PSItem.Type -eq $currentSvc.EnvType})[0].ServiceName
                    }
                    else {
                        $currentSvc.SvcName = "$cmp.AppPool"
                    }
                    $tnSvcs += $currentSvc
                }
                else {
                    $currentSvc.Servers += $node.NodeName
                }

            }
        }
        return $tnSvcs
    }

    end {
    }
}

#endregion