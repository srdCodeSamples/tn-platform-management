#region - Inital Configurations
#Function for choosing action variable - forces user to type any of the words given and returns it.
Function Set-CSTMChoice {
    [cmdletbinding()]
    Param (
    [Parameter(Mandatory=$True,Position=1)]
    [string]$Prompt,
    [Parameter(Mandatory=$True,Position=2)]
    [array]$Choices,
    [Parameter(ParameterSetName='Override',Mandatory=$false)]
    [array]$overrides = @($false),
    [Parameter(ParameterSetName='Override',Mandatory=$false)]
    $choiceOnOveride = [string]::Empty
    )
    Process {
        [bool]$override = $true
        foreach ($cnd in $overrides)
        {
            $override = $override -and $cnd
        }
        if(!$override)
        {
            DO {
                $choice = Read-Host -Prompt "$Prompt ? ($([System.String]::Join(',',$Choices)))"
            } Until ($choice -in $Choices)
            return $choice
        } else {
            return $choiceOnOveride
        }
    }
}


## Pprompt user for environment and script execution mode
Write-Host -Object "Make sure that:" -ForegroundColor Yellow
Write-Host -Object "1. The .xml used for the update is $PSScriptRoot\update.xml","" -ForegroundColor Yellow

$DC = Set-CSTMChoice -Prompt "Choose environment" -Choices "PROD","NYPROD","AWSPROD",'SARANYPROD','SARAPROD'
Switch ($DC) {
"PROD"    { $publishsrv = "prodpublish.tradenetworks.ams"
            $rdshost = "amsrdshost1.tradenetworks.ams"
          }
"NYPROD"  { $publishsrv = "nypublish.tradenetworks.ams"
            $rdshost = "nyrdshost1.tradenetworks.ams"
          }
"AWSPROD"  { $publishsrv = "awspublish.aws.local"
            $rdshost = "awsrdshost1.aws.local"
          }
"SARANYPROD"  { $publishsrv = "nypublish.tradenetworks.ams"
                $rdshost = "nyrdshost1.tradenetworks.ams"
              }
 "SARAPROD"    { $publishsrv = "prodpublish.tradenetworks.ams"
             $rdshost = "amsrdshost1.tradenetworks.ams"
           }
}

#promt the user to set the execution mode
$executionModeChoice = Set-CSTMChoice -Prompt "Choose environment" -Choices "AutoFullVersion","AutoUpdateOnly","Manual"
switch($executionModeChoice)
{
    "AutoFullVersion" {
        $AutoMode = $true
        $copyMode = "FullVersion"
    }
    "AutoUpdateOnly" {
        $AutoMode = $true
        $copyMode = "UpdateOnly"
    }
    "Manual" {
        $AutoMode = $false
        $copyMode = @()
    }
}

#Source publish server
$sourcePublishSrv = 'sofentpublish.cloudad.local'

#Local Path for storing files
$localPath = "$env:USERPROFILE\Desktop\TempVersion"

#get TN credentials when debugging with PSIE as it does not load the PS Profile
$tnuser = "tradenetworks\slavd"
$tnpass =  (Get-Content -Path 'C:\Users\slav.donchev\Documents\WindowsPowerShell\tnpass.txt' | ConvertTo-SecureString)
$tncred = New-Object System.Management.Automation.PSCredential -ArgumentList $tnuser, $tnpass

#Get TN credentials - set in the powershell profile to load at PS startup or uncomment the line below.
#$tncred = Get-Credential -UserName "Tradenetworks\slavd" -Message "TN Credentials"
$tnusr = $tncred.UserName

#Session variable
$session = $null

#session options
$so = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck

#TN user's temp dir on the RDS host
$tempdir = $null

#Name for the temp PSDrives that will need to be created
$tmpDriveName = 'TMP'

#endregion

#region - Get the Lastest versions.csv and applicaitons.csv

#prompt the user to wheather to get them
$choice = @()
$choice = Set-CSTMChoice -Prompt "Get the latest versions.csv and applications.csv from $publishsrv" -Choices "Y","N" -overrides $AutoMode,$(!$Error) -choiceOnOveride 'Y'
$Error.Clear()
#get the latest versions.csv
if ($choice  -eq 'Y') {
    Write-Host "Copying latest versions.csv and applications.csv from $publishsrv to $PSScriptRoot\$DC..."
    if($session -eq $null)
    {
        Write-Host -Object "Creating session to $rdshost"
        $session = New-PSSession -ComputerName $rdshost -Credential $tncred -UseSSL -SessionOption $so
    }
    if($tempdir -eq $null)
    {
        Write-Host -Object "Getting $tnuser's temp dir on $rdshost"
        $tempdir = Invoke-Command -Session $session -ScriptBlock { return $env:temp }
    }
    Write-Host -Object "Copying files from $publishsrv to $rdshost..."
    Invoke-Command -Session $session -ScriptBlock {
                                                    Param($publishsrv,$tncred,$DC,$tmpDriveName,$tempdir)
                                                    $null = New-PSDrive -Name $tmpDriveName -Credential $tncred -Root "\\$publishsrv\d$" -PSProvider FileSystem
                                                    Copy-Item -Path "$tmpDriveName`:\dsc\$DC\versions.csv" -Destination "$tempdir\versions.csv" -Force
                                                    Copy-Item -Path "$tmpDriveName`:\dsc\$DC\applications.csv" -Destination "$tempdir\applications.csv" -Force
                                                  } -ArgumentList $publishsrv,$tncred,$DC,$tmpDriveName,$tempdir
    Write-Host -Object "Copying files from $rdshost to $PSScriptRoot\$DC..."
    $null = New-PSDrive -Name $tmpDriveName -Credential $tncred -Root "\\$rdshost\$($tempdir.Replace(':','$'))" -PSProvider FileSystem
    Copy-Item -Path "$tmpDriveName`:\versions.csv" -Destination "$PSScriptRoot\$DC\Versions.csv" -Force
    Copy-Item -Path "$tmpDriveName`:\applications.csv" -Destination "$PSScriptRoot\$DC\applications.csv" -Force
    Write-Host -Object "Done."
}
#endregion

#region - Create the new versions.csv with the updated versions
#load the xml,versions.csv and applications.csv
$skipcmp = 'EndOfDayCompliter','Backoffice','Launcher','EndClient','ManagedAccountsGateway','GeoLocationRestrictions'
[xml]$xml = Get-Content -Path "$PSScriptRoot\update.xml"
$vers = Import-csv -Path "$PSScriptRoot\$DC\Versions.csv"
$apps = Import-csv -Path "$PSScriptRoot\$DC\applications.csv"
$tfsToDSCMap = Import-csv -Path "$PSScriptRoot\TFStoDSCmap.csv"

#Fix differences in the component names between the TFS and the DSC
foreach($dif in $xml.root.file | Where-Object -Property "Service" -In -Value $tfsToDSCMap.TFS)
    {
        $dif.service = $tfsToDSCMap.where{$PSitem.TFS -eq $dif.service}.DSC
    }

#Update the Versions.csv
foreach($component in $xml.root.file) {
    if ($component.service -notin $skipcmp) {
        Write-Host -Object "Setting version - " -NoNewline
        Write-Host -Object $("{0,-33}{1,-10}{2}" -f $component.service,$component.Branch,$component.version) -ForegroundColor Yellow
        ($vers | Where-Object -Property Component -EQ -Value  $component.service).Version = $component.version
        ($vers | Where-Object -Property Component -EQ -Value  $component.service).Branch = $component.Branch
    }
}


## Generate the .csv
Write-Host -Object "Generating $localPath\versions.csv ..."
Out-File -InputObject "Component,Branch,Version" -FilePath "$localPath\versions.csv" -Force
foreach ($cmp in $vers) {
Out-File -InputObject $($cmp.component+','+$cmp.Branch+','+$cmp.Version) -FilePath "$localPath\versions.csv" -Append
}

#endregion

#region - Copy the versions.csv back to the publish server
$choice = @()
$choice = Set-CSTMChoice -Prompt "Copy versions.csv back to $publishsrv" -Choices 'Y','N' -overrides $AutoMode,$(!$Error) -choiceOnOveride 'Y'
$Error.Clear()
if ($choice -eq 'Y') {
    if($session -eq $null)
    {
        Write-Host -Object "Creating session to $rdshost"
        $session = New-PSSession -ComputerName $rdshost -Credential $tncred -UseSSL -SessionOption $so
    }
    if($tempdir -eq $null)
    {
        Write-Host -Object "Getting $tnuser's temp dir on $rdshost"
        $tempdir = Invoke-Command -Session $session -ScriptBlock { return $env:temp }
    }
    if(!(Test-Path -Path "$tmpDriveName`:\"))
    {
        $null = New-PSDrive -Name $tmpDriveName -Credential $tncred -Root "\\$rdshost\$($tempdir.Replace(':','$'))" -PSProvider FileSystem
    }
    Write-Host -Object "Copying files to $rdshost..."
    Copy-Item -Path "$localPath\versions.csv" -Destination "$tmpDriveName`:\versions.csv" -Force
    Invoke-Command -Session $session {
                                       Param($publishsrv,$tmpDriveName,$tncred,$DC,$tempdir)
                                       if(!(Test-Path -Path "$tmpDriveName`:\"))
                                       {
                                           $null = New-PSDrive -Name $tmpDriveName -Credential $tncred -Root "\\$publishsrv\d$" -PSProvider FileSystem
                                       }
                                       Copy-Item -Path "$tempdir\versions.csv" -Destination "$tmpDriveName`:\dsc\$DC\versions.csv" -Force
                                       Remove-Item -Path "$tempdir\versions.csv" -Force
                                     } -ArgumentList $publishsrv,$tmpDriveName,$tncred,$DC,$tempdir
}

#endregion

#region - Copy the version files from sofent publish
$choice = @()
$choice = Set-CSTMChoice -Prompt "Get components from $sourcePublishSrv" -Choices "Y","N" -overrides $AutoMode,$(!$Error) -choiceOnOveride 'Y'
$Error.Clear()
if ($choice -eq "N")
{
    if($session -ne $null)
    {
        Write-Host -Object "Closing the session to $rdshost"
        Remove-PSSession -Session $session
    }
    exit
}

$skipcmpcopy = 'Launcher','EndClient','ManagedAccountsGateway','GeoLocationRestrictions'
$svsNotInApp = 'BackOffice','EndOfDayCompliter'
$verfolder = "$DC"+"_$(Get-Date -Format MMddHHmm)"
$verRootPath = "D:\Tradenetworks"
$cmpPathMap = @()

Write-Host -Object "Creating session to $sourcePublishSrv"
$sofentpublishSession = New-PSSession -ComputerName $sourcePublishSrv

$copyMode = Set-CSTMChoice -Prompt "Choose copy mode:" -Choices "FullVersion","UpdateOnly" -overrides $AutoMode,$(!$Error) -choiceOnOveride $copyMode
$verBranch = @($xml.Root.File.Branch | Select-Object -Unique)
if($verBranch.Count -ne 1 -and $copyMode -eq "FullVersion")
{
    Write-Host -Object "Componenets are in more than one branch. Switching to `"UpdateOnly`" mode." -ForegroundColor Red
    $copyMode = "UpdateOnly"
}
$archive = Invoke-Command -Session $sofentpublishSession {return "$env:TEMP\$($Args[0]).zip"} -ArgumentList $verfolder
Switch($copyMode)
{
    "FullVersion" {
        if($xml.root.File.service[0].Contains('BackOffice'))
        {
            $boVersion = $xml.root.File.where{$PSItem.service -eq 'BackOffice'}.Version
            $boSrcPath = "$verRootPath\Clients\BackOffice\PROfit_Setup_BackOffice_PROD_$boVersion.exe"
            $boDstPath = "$verRootPath\$DC\$verBranch\Clients\BackOffice\PROfit_Setup_BackOffice_PROD_$boVersion.exe"
            if(Test-Path -Path "\\$sourcePublishSrv\$($boSrcPath.Replace(":","$"))")
            {
                Write-Host -Object "Copying the BackOffice executable to the version directory on $sourcePublishSrv..."
                Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                    Param($boSrcPath,$boDstPath)
                    $null = [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($boDstPath))
                    Copy-Item -Path $boSrcPath -Destination $boDstPath -Force
                } -ArgumentList $boSrcPath,$boDstPath
            } else {
                Write-Host -Object "The BO executable was not found at the expected path: $boSrcPath. Skipping it" -ForegroundColor Red
            }
        }
        Write-Host -Object "Creating archive on $sourcePublishSrv"
        $archiveJob = Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                        Param($archive,$verRootPath,$DC,$verBranch)
                        $source = "$verRootPath\$DC\$verBranch"
                        Add-Type -AssemblyName 'system.io.compression.filesystem'
                        [io.compression.zipfile]::CreateFromDirectory($source,$archive)
                   } -ArgumentList $archive,$verRootPath,$DC,$verBranch -AsJob
        $timer = $(New-Object -TypeName System.Diagnostics.Stopwatch)
        $timer.Start()
        $iteration = 0
        do {
            $iteration++
            if($iteration -eq 32) {
                $archiveCurrentSize = (Get-ItemProperty -Path "\\$sourcePublishSrv\$($archive.Replace(':','$'))" -Name Length).Length
                $iteration = 0
            }
            Write-Progress -Activity "Creating archive..." -Status $("Elapsed time: {0:D2}:{1:D2}" -f $timer.Elapsed.Minutes,$timer.Elapsed.Seconds) -CurrentOperation $("Archive size: {0:N0} KB" -f $($archiveCurrentSize / 1024))
            Start-Sleep -Milliseconds 500
        } while ($archiveJob.State -eq 'Running')
        $timer.Stop()
        $archiveCurrentSize = (Get-ItemProperty -Path "\\$sourcePublishSrv\$($archive.Replace(':','$'))" -Name Length).Length
        Write-Progress -Activity "Creating archive..." -Status $("Done in {0:D2}:{1:D2}" -f $timer.Elapsed.Minutes,$timer.Elapsed.Seconds) -CurrentOperation $("Archive size: {0:N0} KB" -f $($archiveCurrentSize / 1024))
        $cmpPathMap += New-Object -TypeName PSObject -Property @{
            'Name' = 'FullVersion'
            'Environment' = 'All'
            'RelativePath' = $verBranch
            'ExistInDest' = [bool]$true
        }
        break
    }
    "UpdateOnly" {
        #region - Copy only the components mentioned in the update.xml
        foreach ($cmp in $xml.root.file) {
            if(($cmp.service -notin $skipcmpcopy) -and ($cmp.service -notin $svsNotInApp)) {
                $envs = ($apps | Where-Object -Property 'Component' -EQ -Value $cmp.Service).Type
                foreach ($env in $envs){
                    $spath = @()
                    Write-Host "Start copying $($cmp.Service) - $env to Temp folder on $sourcePublishSrv" -ForegroundColor Yellow
                    $spath = "$verRootPath\$DC\$($cmp.Branch)\$env\$($cmp.Service)\$($cmp.Version)"
                    if (Invoke-Command -Session $sofentpublishSession -ScriptBlock { Test-Path -Path $args[0]} -ArgumentList $spath) {
                        Write-Host "Source path: $spath confirmed. Proceeding with copy..." -ForegroundColor Green
                        $cmpRelPath = "$($cmp.Branch)\$env\$($cmp.Service)\$($cmp.Version)"
                        $dpath = "$verfolder\$cmpRelPath"
                        Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                            param($spath,$dpath)
                            Copy-Item -Path $spath -Destination "$env:TEMP\$dpath" -Recurse
                        } -ArgumentList $spath,$dpath
                        $cmpPathMap += New-Object -TypeName PSObject -Property @{
                                           'Name' = $cmp.Service
                                           'Environment' = $env
                                           'RelativePath' = $cmpRelPath
                                           'ExistInDest' = [bool]$true
                                       }
                    } else {
                        write-host "Source path: $spath does not exist on $sourcePublishSrv. Skipping it." -ForegroundColor Red
                    }
                }
            } elseif ($cmp.service -in $svsNotInApp) {
                switch($cmp.service) {
                    "BackOffice" {
                        $spath = @()
                        $env = 'Central'
                        Write-Host "Start copying $($cmp.Service) to Temp folder on $sourcePublishSrv" -ForegroundColor Yellow
                        $spath = "$verRootPath\Clients\BackOffice\PROfit_Setup_BackOffice_PROD_$($cmp.version).exe"
                        if (Invoke-Command -Session $sofentpublishSession -ScriptBlock { Test-Path -Path $args[0]} -ArgumentList $spath) {
                            Write-Host "Source path: $spath confirmed. Proceeding with copy..." -ForegroundColor Green
                            $cmpRelPath = "$($cmp.Branch)\Clients\BackOffice\$([System.IO.Path]::GetFileName($spath))"
                            $dpath = "$verfolder\$cmpRelPath"
                            Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                                param($spath,$dpath)
                                $dfolder = "$env:TEMP\$([System.IO.Path]::GetDirectoryName($dpath))"
                                if(!(Test-Path -Path $dfolder)) {
                                    $null = New-Item -Type Directory -Path $dfolder
                                }
                                Copy-Item -Path $spath -Destination "$env:TEMP\$dpath"
                            } -ArgumentList $spath,$dpath
                            $cmpPathMap += New-Object -TypeName PSObject -Property @{
                                               'Name' = $cmp.Service
                                               'Environment' = $env
                                               'RelativePath' = $cmpRelPath
                                               'ExistInDest' = [bool]$true
                                           }
                        } else {
                            write-host "Source path: $spath does not exist on $sourcePublishSrv. Skipping it." -ForegroundColor Red
                        }
                        break
                    }
                    "EndOfDayCompliter" {
                        $spath = @()
                        $env = 'Live'
                        Write-Host "Start copying $($cmp.Service) - $env to Temp folder on $sourcePublishSrv" -ForegroundColor Yellow
                        $spath = "$verRootPath\$DC\$($cmp.Branch)\$env\$($cmp.Service)\$($cmp.Version)"
                        if (Invoke-Command -Session $sofentpublishSession -ScriptBlock { Test-Path -Path $args[0]} -ArgumentList $spath) {
                            Write-Host "Source path: $spath confirmed. Proceeding with copy..." -ForegroundColor Green
                            $cmpRelPath = "$($cmp.Branch)\$env\$($cmp.Service)\$($cmp.Version)"
                            $dpath = "$verfolder\$cmpRelPath"
                            Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                                param($spath,$dpath)
                                Copy-Item -Path $spath -Destination "$env:TEMP\$dpath" -Recurse
                            } -ArgumentList $spath,$dpath
                            $cmpPathMap += New-Object -TypeName PSObject -Property @{
                                               'Name' = $cmp.Service
                                               'Environment' = $env
                                               'RelativePath' = $cmpRelPath
                                               'ExistInDest' = [bool]$true
                                           }
                        } else {
                            write-host "Source path: $spath does not exist on $sourcePublishSrv. Skipping it." -ForegroundColor Red
                        }
                        break
                    }
                    "Default" {
                        Write-Host "No copy algorithm set for $($cmp.Service). Skipping it." -ForegroundColor Red
                    }
                }
            }
        }

        Write-Host -Object "Creating archive on $sourcePublishSrv"
        $archiveJob = Invoke-Command -Session $sofentpublishSession -ScriptBlock {
                       Param($verfolder,$archive)
                       $source = "$env:TEMP\$verfolder"
                       Add-Type -AssemblyName 'system.io.compression.filesystem'
                       [io.compression.zipfile]::CreateFromDirectory($source,$archive)
                   } -ArgumentList $verfolder,$archive -AsJob
        $timer = $(New-Object -TypeName System.Diagnostics.Stopwatch)
        $timer.Start()
        $iteration = 0
        do {
            $iteration++
            if($iteration -eq 10) {
                $archiveCurrentSize = (Get-ItemProperty -Path "\\$sourcePublishSrv\$($archive.Replace(':','$'))" -Name Length).Length
                $iteration = 0
            }
            Write-Progress -Activity "Creating archive..." -Status $("Elapsed time: {0:D2}:{1:D2}" -f $timer.Elapsed.Minutes,$timer.Elapsed.Seconds) -CurrentOperation $("Archive size: {0:N0} KB" -f $($archiveCurrentSize / 1024))
            Start-Sleep -Milliseconds 500
        } while ($archiveJob.State -eq 'Running')
        $timer.Stop()
        $archiveCurrentSize = (Get-ItemProperty -Path "\\$sourcePublishSrv\$($archive.Replace(':','$'))" -Name Length).Length
        Write-Progress -Activity "Creating archive..." -Status $("Done in {0:D2}:{1:D2}" -f $timer.Elapsed.Minutes,$timer.Elapsed.Seconds) -CurrentOperation $("Archive size: {0:N0} KB" -f $($archiveCurrentSize / 1024))
        break
    }
}


Write-Host "Copying archive from $sourcePublishSrv to $localPath\$verfolder\ ..."
$zipfile = $([System.IO.Path]::GetFileName($archive))
$ziparchive = "$localPath\$zipfile"
Copy-Item -Path "\\$sourcePublishSrv\$($archive.Replace(':','$'))" -Destination $ziparchive
Write-Host "Cleaning Temp folder on $sourcePublishSrv..."
Invoke-Command -Session $sofentpublishSession -ScriptBlock {
    Remove-Item -Path "$env:TEMP\$($Args[0])" -Force -Recurse -ErrorAction SilentlyContinue
    Remove-Item -Path $Args[1] -Force
} -ArgumentList $verfolder,$archive

Write-Host "Closing PSSession to $sourcePublishSrv"
Remove-PSSession -Session $sofentpublishSession

#endregion

#region - Copy the version archive to the desktop's publish server
$choice = @()
$choice = Set-CSTMChoice -Prompt "Copy the version archive to your desktop on $publishsrv" -Choices 'Y','N' -overrides $AutoMode,$(!$Error) -choiceOnOveride 'Y'
$Error.Clear()
if ($choice -eq 'Y')
{
    if($session -eq $null)
    {
        Write-Host -Object "Creating session to $rdshost"
        $session = New-PSSession -ComputerName $rdshost -Credential $tncred -UseSSL -SessionOption $so
    }
    if($tempdir -eq $null)
    {
        Write-Host -Object "Getting $tnuser's temp dir on $rdshost"
        $tempdir = Invoke-Command -Session $session -ScriptBlock { return $env:temp }
    }
    if(!(Test-Path -Path "$tmpDriveName`:\"))
    {
        $null = New-PSDrive -Name $tmpDriveName -Credential $tncred -Root "\\$rdshost\$($tempdir.Replace(':','$'))" -PSProvider FileSystem
    }
    Write-Host -Object "Copying $ziparchive to your desktop on $publishsrv ..."
    Copy-Item -Path $ziparchive -Destination "$tmpDriveName`:\$zipfile" -Force
    Invoke-Command -Session $session {
                                        Param($publishsrv,$tnusr,$zipfile,$tncred,$tempdir)
                                        $null = New-PSDrive -Name 'DSKTP' -Credential $tncred -Root "\\$publishsrv\c$\users\$($tnusr.split("\")[1])\desktop" -PSProvider FileSystem
                                        Copy-Item -Path "$tempdir\$zipfile" -Destination "DSKTP:\$zipfile" -Force
                                        Remove-Item -Path "$tempdir\$zipfile" -Force
                                     } -ArgumentList $publishsrv,$tnusr,$zipfile,$tncred,$tempdir

##Extract the archive on the publish server
    $DSCdeploymentFolder = "D:\Tradenetworks\$DC"
    $choice = @()
    $choice = Set-CSTMChoice -Prompt "Extract $zipfile to $DSCdeploymentFolder on $publishsrv ? If any of the components already exist on $publishsrv extraction will be aborted." -Choices 'Y','N' -overrides $AutoMode,$(!$Error) -choiceOnOveride 'Y'
    $Error.Clear()
    if ($choice -eq 'Y')
    {
        $publishsrvZipFile = "C:\Users\$($tnusr.split("\")[1])\Desktop\$zipfile"
        Write-Host "Extracting files on $publishsrv to $DSCdeploymentFolder. Please be patient..."
        $existPathMap = Invoke-Command -Session $session  -ScriptBlock {
                                            param($publishsrv,$DSCdeploymentFolder,$publishsrvZipFile,$tncred,$cmpPathMap,$copyMode,$verBranch)
                                            foreach ($mapCmp in $cmpPathMap)
                                            {
                                                $mapCmp.ExistInDest = Test-Path -Path "\\$publishsrv\$($DSCdeploymentFolder.Replace(':','$'))\$($mapCmp.RelativePath)"
                                            }
                                            $existPathMap = @($cmpPathMap.where{$PSitem.ExistInDest -eq $true})
                                            if($existPathMap.count -eq 0)
                                            {
                                                if($copyMode -eq 'FullVersion')
                                                {
                                                    $pubVersionPath = "$DSCdeploymentFolder\$verBranch"
                                                } else {
                                                    $pubVersionPath = $DSCdeploymentFolder
                                                }
                                                Invoke-Command -ComputerName $publishsrv -ScriptBlock {
                                                                                                        if($args[2] -eq 'FullVersion')
                                                                                                        {
                                                                                                            $null = [System.IO.Directory]::CreateDirectory($args[1])
                                                                                                        }
                                                                                                        Add-Type -AssemblyName 'system.io.compression.filesystem'
                                                                                                        [io.compression.zipfile]::ExtractToDirectory($args[0],$args[1])
                                                } -Credential $tncred -ArgumentList $publishsrvZipFile,$pubVersionPath,$copyMode
                                            }
                                            return $existPathMap
                                          } -ArgumentList $publishsrv,$DSCdeploymentFolder,$publishsrvZipFile,$tncred,$cmpPathMap,$copyMode,$verBranch
        if($existPathMap.count -eq 0)
        {
            Write-Host "Finished."
        } else {
            Write-Host "Extraction was aborted because the following components exist on $publishsrv `n" -ForegroundColor Red
            $existPathMap | Format-Table -AutoSize -Wrap -Property Name,Environment,RelativePath
        }
    }
}

#endregion

if($session -ne $null)
{
    Write-Host -Object "Closing the session to $rdshost"
    Remove-PSSession -Session $session
}

Read-Host -Prompt "Finished. Press enter to quit"