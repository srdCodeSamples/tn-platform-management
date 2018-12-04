# Get the ENV as param passed when staring the file.
Param([Parameter(Mandatory = $true, Position = 1)][string]$env)
$env = $env.ToUpper()
$scriptStartTime = Get-Date

#Import Finctions
Import-Module "$PSScriptRoot\Helpers\PROfitSundayRestarts.Helpers.psm1"
Import-Module "$PSScriptRoot\Helpers\PSFunctions_v3.1.psm1"

#Log File
$logFilePath = "$PSScriptRoot\Logs\Log_$(Get-Date -Format "yyyyMMdd").log"

#status File
$statusFilePath = "$PSScriptRoot\Logs\Status_$(Get-Date -Format "yyyyMMdd").txt"

#region - Select Data based on environment
Write-CSTMLog -FilePath $logFilePath -Type "INFO" -Message "[START] Initiating sunday restart script for: $env" -Thread "Main" -Append $false
$envFilePath = "$PSScriptRoot\Data\Environments.json"
$err = @()
try {
    $environments = ConvertFrom-Json -InputObject "$(Get-Content -Path $envFilePath)" -ErrorVariable +err
}
catch {
    Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $Error[0] -Thread "Main"
    Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message "There was error parsing environments data file - $envFilePath. Aborting execution." -Thread "Main"
    $err = @()
    exit
}
finally {
    if($err.count -gt 0) {
        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $Error[0] -Thread "Main"
        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message "There was error parsing environemts data file - $envFilePath. Aborting execution." -Thread "Main"
        exit
    }
}
$selectedEnv = $environments.Where( {$PSItem.Name -eq $env})[0]
if ($selectedEnv -eq $null) {
    Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message "$env is not a valid environement. Aborting execution." -Thread "Main"
    exit
}
else {
    $DSCsrv = $selectedEnv.DSCsrv
    $MailSMTP = $selectedEnv.MailSMTP
    $Mailform = $selectedEnv.Mailform
    $Mailto = @($selectedEnv.Mailto)
    $Mailcc = @($selectedEnv.Mailcc)
}
Remove-Variable -Name 'envFilePath','environments','selectedEnv'

#endregion

#region - Build lists of services to restrt

$svcsToRestartCsv = @(Import-Csv -Path "$PSScriptRoot\Data\RestartedServices.csv")
$allTnSvcs = @(Get-TNDscComponents -DscEnvConfigPath "\\$DSCsrv\d$\DSC\Config_$env.psd1" -DscAppPath "\\$DSCsrv\d$\DSC\$env\Applications.csv")
$tnSvcsToRestart = $allTnSvcs.Where( {$PSItem.Component -in $svcsToRestartCsv.Component})
Remove-Variable -Name 'allTnSvcs'
foreach ($svc in $tnSvcsToRestart) {
    $svc.HealthCheck = $svcsToRestartCsv.Where( {$PSItem.Component -eq $svc.Component})[0].HealthCheck
    $svc.Priority = [int]$svcsToRestartCsv.Where( {$PSItem.Component -eq $svc.Component})[0].Priority
    $svc.HMProtocol = $svcsToRestartCsv.Where( {$PSItem.Component -eq $svc.Component})[0].HMProtocol
}
Remove-Variable -Name 'svcsToRestartCsv'
$orderedSvcs = $tnSvcsToRestart.Where( {$PSitem.Priority -gt 0}) | Sort-Object -Property Priority -Descending
$unorderedSvcs = $tnSvcsToRestart.Where( {$PSitem.Priority -eq 0})

#endregion

#region - Send Start e-Mail

Write-CSTMLog -FilePath $logFilePath -Message "Sending [START] e-mail via SMPT server: `"$MailSMTP`"" -Type "INFO" -Thread "Main"
$SMailsubj = "[$env][START] Sunday Maintenance"
$SMailbody = "Dear Tech Support Team<br><br>"
$SMailbody += "The Sunday maintenance tasks are now in progress.<br>"
$SMailbody += "Please ignore the upcoming SCOM alarms until further notice.<br><br>"
$SMailbody += "Kind Regards<br>"
$SMailbody += "System and Application Administration Team"
try {
    Send-MailMessage -From $Mailform -To $Mailto -Cc $Mailcc -SmtpServer $MailSMTP -Subject $SMailsubj -BodyAsHtml $SMailbody -ErrorVariable genErr
}
catch {
    Write-CSTMLog -FilePath $logFilePath -Message $Error[0] -Type "ERROR" -Thread "Main"
    $Error.Clear()
}
finally {
    if ($genErr.count -ne 0) {
        Write-CSTMLog -FilePath $logFilePath -Message $genErr -Type "ERROR" -Thread "Main"
    }
    $genErr = @()
}
Remove-Variable -Name 'SMailsubj', 'SMailbody'
Start-Sleep -Seconds 30

#endregion

#region - Restart ordered services

$orderedSvcsPause = 90
if ($orderedSvcs.count -gt 0) {
    foreach ($osvc in $orderedSvcs) {
        switch ($osvc.Type) {
            "App" {
                Write-CSTMLog -FilePath $logFilePath -Message "Restarting `"$($osvc.SvcName)`" on $($osvc.Servers)" -Type "INFO" -Thread "Main"
                try {
                    $err = @()
                    Invoke-Command -ComputerName $osvc.Servers -ScriptBlock {Restart-Service -Name $Args[0]} -ArgumentList $osvc.SvcName -ErrorVariable +err
                }
                catch {
                    Write-CSTMLog -FilePath $logFilePath -Message $Error[0] -Type "ERROR" -Thread "Main"
                    $Error.Clear()
                }
                finally {
                    if ($err.Count -gt 0) {
                        Write-CSTMLog -FilePath $logFilePath -Message $err -Type "ERROR" -Thread "Main"
                        $err = @()
                    }
                }
                Write-CSTMLog -FilePath $logFilePath -Message "Waiting for $orderedSvcsPause seconds after restart." -Type "DEBUG" -Thread "Main"
                Start-Sleep -Seconds $orderedSvcsPause
                $appSvcCheck = $null
                $appSvcCheck = Invoke-Command -ComputerName $osvc.Servers -ScriptBlock {Get-Service -Name $Args[0]} -ArgumentList $osvc.SvcName
                Write-CSTMLog -FilePath $logFilePath -Message "Status of `"$($osvc.SvcName)`" on $($osvc.Servers): $($appSvcCheck.Status)" -Type "DEBUG" -Thread "Main"
            }
            default {
                try {
                    Restart-TNWebService -Servers $osvc.Servers -AppPoolName $osvc.SvcName -LogPath $logFilePath -Pause $orderedSvcsPause -HMUri $osvc.HealthCheck -HMProtocol $osvc.HMProtocol -Thread "Main"
                }
                catch {
                    Write-CSTMLog -FilePath $logFilePath -Message $Error[0] -Type "ERROR" -Thread "Main"
                    $Error.Clear()
                }
            }
        }
    }
}

#endregion

#region - Restart unordered Web services in parallel

$unOrderedWebSvcsPause = 90
$unOrderedWebSvcs = $unorderedSvcs.Where( {$PSItem.Type -ne "App"})
$webSvcRestartJobs = @()
foreach ($webSvc in $unOrderedWebSvcs) {
    $threadName = "$($webSvc.Component)-$($webSvc.EnvType)-Recycle"
    $webSvcRestartJobs += start-job -ScriptBlock {
        Param ($Servers, $SvcName, $unOrderedWebSvcsPause, $logFilePath, $HealthCheck, $HMProtocol, $threadName, $PSScriptRoot)
        Import-Module "$PSScriptRoot\Helpers\PROfitSundayRestarts.Helpers.psm1"
        Import-Module "$PSScriptRoot\Helpers\PSFunctions_v3.1.psm1"
        try {
            Restart-TNWebService -Servers $Servers -AppPoolName $SvcName -Pause $unOrderedWebSvcsPause -LogPath $logFilePath -HMUri $HealthCheck -HMProtocol $HMProtocol -Thread $threadName
        }
        catch {
            Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Thread $threadName -Message $Error[0]
        }
    } -ArgumentList $webSvc.Servers, $webSvc.SvcName, $unOrderedWebSvcsPause, $logFilePath, $webSvc.HealthCheck, $webSvc.HMProtocol, $threadName, $PSScriptRoot
}
#endregion

#region - Restart unordered App services

$unorderedAppSvcsPause = 90
$unOrderedAppSvcs = $unorderedSvcs.Where( {$PSItem.Type -eq "App"})
foreach ($appSvc in $unOrderedAppSvcs) {
    Write-CSTMLog -FilePath $logFilePath -Message "Restarting `"$($appSvc.SvcName)`" on $($appSvc.Servers)" -Type "INFO" -Thread "Main"
    try {
        $err = @()
        Invoke-Command -ComputerName $appSvc.Servers -ScriptBlock {Restart-Service -Name $Args[0]} -ArgumentList $appSvc.SvcName -ErrorVariable +err
    }
    catch {
        Write-CSTMLog -FilePath $logFilePath -Message $Error[0] -Type "ERROR" -Thread "Main"
        $Error.Clear()
    }
    finally {
        if ($err.Count -gt 0) {
            Write-CSTMLog -FilePath $logFilePath -Message $err -Type "ERROR" -Thread "Main"
            $err = @()
        }
    }
    Write-CSTMLog -FilePath $logFilePath -Message "Waiting for $unorderedAppSvcsPause seconds after restart." -Type "DEBUG" -Thread "Main"
    Start-Sleep -Seconds $unorderedAppSvcsPause
    $appSvcCheck = $null
    $appSvcCheck = Invoke-Command -ComputerName $appSvc.Servers -ScriptBlock {Get-Service -Name $Args[0]} -ArgumentList $appSvc.SvcName
    Write-CSTMLog -FilePath $logFilePath -Message "Status of `"$($appSvc.SvcName)`" on $($appSvc.Servers): $($appSvcCheck.Status)" -Type "DEBUG" -Thread "Main"
}

#endregion

$timeout = 1800
Write-CSTMLog -FilePath $logFilePath -Message "Finished restarting windows services. Waiting for the web services restart jobs to finish. Time out is set to $timeout seconds." -Type "DEBUG" -Thread "Main"
$null = $webSvcRestartJobs | Wait-Job -Timeout $timeout
Write-CSTMLog -FilePath $logFilePath -Message "Restarts of web services finished." -Type "DEBUG" -Thread "Main"

#region - Generate Status report for all restarted services

#region windows services

Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Generating Windows services status" -Thread "Main"
$tnAppSvcs = $tnSvcsToRestart.Where( {$PSitem.Type -eq 'App'})
$tnAppSrvs = @()
foreach ($svc in $tnAppSvcs) { $tnAppSrvs += $svc.Servers }
$tnAppSrvs = $tnAppSrvs | Select-Object -Unique
$err = @()
try {
    $tnAppSvcsState = Invoke-Command -ComputerName $tnAppSrvs -ScriptBlock {
        Param($filter)
        $tnSvcs = @(Get-service | Where-Object -FilterScript {$PSitem.Name -in $filter})
        foreach ( $svc in $tnSvcs.where{$PSItem.Status -eq "Running"}) {
            $stime = @()
            $stime = (Get-Process -id (Get-CimInstance -ClassName Win32_service -Filter "Name = '$($svc.Name)'").ProcessId).StartTime
            $svc | Add-Member -Name "StartTime" -Value $Stime -MemberType NoteProperty
        }
        return $tnSvcs
    } -ArgumentList (, $tnAppSvcs.SvcName) -ErrorVariable +err
}
catch {
    Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $Error[0] -Thread "Main"
    $Error.Clear()
}
finally {
    if ($err.count -gt 0) {
        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $err -Thread "Main"
        $err = @()
    }
}

#endregion

#region web services - get AppPools state

Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Generating Web services status" -Thread "Main"
$tnWebSvcs = $tnSvcsToRestart.Where( {$PSitem.Type -ne 'App'})
$tnWebSrvs = @()
foreach ($svc in $tnWebSvcs) { $tnWebSrvs += $svc.Servers }
$tnWebSrvs = $tnWebSrvs | Select-Object -Unique
$err = @()
try {
    $tnWebSvcsState = Invoke-Command -ComputerName $tnWebSrvs -ScriptBlock {
        Param($filter)
        Import-Module WebAdministration
        $tnSvcs = @(Get-ChildItem -path "IIS:\AppPools" | Where-Object -FilterScript {$PSitem.Name -in $filter})
        foreach ( $svc in $tnSvcs.where{$PSItem.State -eq "Started"}) {
            $stime = @()
            $stime = (Get-ChildItem -path IIS:\AppPools\$($svc.Name)\workerprocesses).StartTime
            $svc | Add-Member -Name "StartTime" -Value $Stime -MemberType NoteProperty
        }
        return $tnSvcs
    } -ArgumentList (, $tnWebSvcs.SvcName) -ErrorVariable +err
}
catch {
    Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $Error[0] -Thread "Main"
    $Error.Clear()
}
finally {
    if ($err.count -gt 0) {
        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $err -Thread "Main"
    }
    $err = @()
}

#endregion

#region web services - make health checks
Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Performing Web Services HealthChecks" -Thread "Main"
foreach ($webSvc in $tnWebSvcs) {
    foreach ($srv in $webSvc.Servers) {
        if (![System.String]::IsNullOrEmpty($webSvc.HMProtocol)) {
            $hmResult = "failed"
            $hmUri = $null
            $hmUri = "$($webSvc.HMProtocol)`://$srv/$($webSvc.HealthCheck)"
            $hmAttempt = 1
            $hmMaxAttempts = 3
            $hmAttemptPause = 3
            $hmRetry = $false
            DO {
                Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Making healthcheck, attempt: $hmAttempt, for $($webSvc.Component) on $srv with URL $hmUri" -Thread "Main"
                try {
                    $err = @()
                    $hmResult = (Invoke-WebRequest -Uri "$hmUri" -MaximumRedirection 0 -ErrorVariable +err).StatusDescription
                }
                catch {
                    if ($Error.count -gt 0 -and $isLoggingEnabled) {
                        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $Error[0].Exception.Message.ToString() -Thread "Main"
                    }
                    $Error.Clear()
                }
                finally {
                    if ($err.Count -gt 0 -and $isLoggingEnabled) {
                        Write-CSTMLog -FilePath $logFilePath -Type "ERROR" -Message $err.InnerException.Message.ToString() -Thread "Main"
                    }
                    $err = @()
                }
                if (($hmResult -eq 'failed') -and ($hmAttempt -le $hmMaxAttempts)) {
                    $hmRetry = $true
                    $hmAttempt++
                    Start-Sleep -Seconds $hmAttemptPause
                }
                else {
                    $hmRetry = $false
                }
            } while ($hmRetry)
        }
        else {
            $hmResult = "NotConfigured"
        }
        Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Health check for $($webSvc.Component) on $srv with URL $hmUri : $hmResult" -Thread "Main"
        $tnWebSvcsState.where{($PSitem.name -eq $webSvc.SvcName) -and ($PSitem.PSComputerName -eq $srv)}[0] | Add-Member -MemberType NoteProperty -Value $hmResult -Name HMState
    }
}

#endregion

#region commit status resulst to file

Write-CSTMLog -FilePath $logFilePath -Type "DEBUG" -Message "Commiting status results to file $statusFilePath" -Thread "Main"
Out-File -FilePath $statusFilePath -InputObject "Services' State after restarts on $scriptStartTime", ""
Out-File -FilePath $statusFilePath -InputObject "Windows services' status:" -Append
Out-File -FilePath $statusFilePath -InputObject ($tnAppSvcsState | Sort-Object -Property PSComputerName, DisplayName | Format-Table -AutoSize -Property DisplayName, Status, StartTime, PSComputerName) -Append
Out-File -FilePath $statusFilePath -InputObject "Web services' status:" -Append
Out-File -FilePath $statusFilePath -InputObject ($tnWebSvcsState | Sort-Object -Property Name, PSComputerName | Format-Table -AutoSize -Property Name, State, StartTime, HMState, PSComputerName) -Append

#endregion

#endregion

#region - Send e-mail with log and status

Write-CSTMLog -FilePath $logFilePath -Message "Sending [END] e-mail via SMPT server: `"$MailSMTP`"" -Type "INFO" -Thread "Main"
$endMailsubj = "[$env][END] Sunday Maintenance"
$endMailbody = "Dear Tech Support Team,<br><br>"
$endMailbody += "The Sunday maintenance tasks are finished and you can resume normal SCOM monitoring.<br>"
$endMailbody += "Please check the attached status file.<br>"
$endMailbody += "<font color=`"FF0000`">In case in the status report there is:</font>"
$endMailbody += "<blockquote>1. Status different from Runnig/Started.<br>"
$endMailbody += "2. HMState that is failed.<br>"
$endMailbody += "3. StartTime that is from before the time of the restart's begining (indicated in the begining of the status file).</blockquote>"
$endMailbody += "<font color=`"FF0000`">Call the current on call member of the TechSysAndApp team.</font><br><br>"
$endMailbody += "Kind Regards<br>"
$endMailbody += "System and Application Administration Team"
$endMailAttachment = @($logFilePath, $statusFilePath)
try {
    Send-MailMessage -From $Mailform -To $Mailto -Cc $Mailcc -SmtpServer $MailSMTP -Subject $endMailsubj -BodyAsHtml $endMailbody -Attachments $endMailAttachment -ErrorVariable genErr
}
catch {
    Write-CSTMLog -FilePath $logFilePath -Message $Error[0] -Type "ERROR" -Thread "Main"
    $Error.Clear()
}
finally {
    if ($genErr.count -ne 0) {
        Write-CSTMLog -FilePath $logFilePath -Message $genErr -Type "ERROR" -Thread "Main"
    }
    $genErr = @()
}

#endregion

Write-CSTMLog -FilePath $logFilePath -Message "[END] Finished Sunday restarts script for: $env" -Thread "Main"