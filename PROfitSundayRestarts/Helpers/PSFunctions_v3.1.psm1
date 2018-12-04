 <#
.SYNOPSIS
Set-CSTMChoice prompts the user to input one of several specified strings and returns it.
.DESCRIPTION
.PARAMETER Prompt
The prompt displayed to the user. Mandatory.
.PARAMETER Choices
The choices given to the user to input
.EXAMPLE
Set-CSTMChoice -prompt "Do you with to continue" -choices "Yes","No"
#>
Function Set-CSTMChoice {
    [cmdletbinding()]
    Param (
    [Parameter(Mandatory=$True,Position=1)]
    [string]$Prompt,
    [Parameter(Mandatory=$True,Position=2)]
    [array]$Choices
    )
    Process {
        DO {
            $choice = Read-Host -Prompt "$Prompt ? ($([System.String]::Join(',',$Choices)))"
        } Until ($choice -in $Choices)
        return $choice
    }
}
#End

<#
.SYNOPSIS
Start-CSTMSleep pauses for a sertain amaount of time while displaying progress indicator with timer in "Status"
.DESCRIPTION
Start-CSTMSleep creates or uses a user provided "Stopwatch" object of type [System.Diagnostics.Stopwatch]. It then restarts or continues to run the timer
for a specified ammount of time while displaying the elapsed time in a progres indicatior (using write-progress) as status. The Activity and Current operation are
user defined via prameters. After the specified time period the function can stop or leave on the timer as well as return it as an object.
.PARAMETER Seconds
Specify the time length for the pause in seconds. Mandatory.
.PARAMETER Activity
String to be displayed in the Activity section of the progress indicator. Mandatory
.PARAMETER Currentop
String to be displayed in the Current operation section of the progress indicator.
.PARAMETER CSTMStopWatch
An object of type System.Diagnostics.Stopwatch for the funcion to use instead of creating a new one (default)
.PARAMETER Reset
Specifies wheather to reset the stopwatch or to continue counting from it's current progress. Default is False.
In case no object was specified via the CSTMStopWatch parameter the timer will start from zero anyway.
.PARAMETER Stop
Specifies wheather to stop the stopwatch after the pause time elapses. Default is True.
.PARAMETER ReturnTimer
Specifies wheather to return the timer as a result of the function. Default is False.
.EXAMPLE
Simply pause for 30s while displaying progress
Start-CSTMSleep -activity "Waiting for some time" -seconds 30
.EXAMPLE
Pause for 30 seconds while displaying progress, current operation and continue whth an already started timer which is not stopped and is returned after the pause finishes.
Start-CSTMSleep -activity "Waiting for some time" -seconds 30 -currentop "step 1" -$CSTMStopWatch $timer -stop $false -returntimer $true
#>
Function Start-CSTMSleep {
    [cmdletbinding()]
    Param (
    [Parameter(Mandatory=$True,Position=1)]
    [int]$Seconds,
    [Parameter(Mandatory=$True,Position=2)]
    [String]$Activity,
    [Parameter(Mandatory=$False,Position=3)]
    [String]$Currentop = '',
    [Parameter(Mandatory=$False)]
    [System.Diagnostics.Stopwatch]$CSTMStopWatch = $(New-Object -TypeName System.Diagnostics.Stopwatch),
    [Parameter(Mandatory=$False)]
    [bool]$Reset = $False,
    [Parameter(Mandatory=$False)]
    [bool]$Stop = $True,
    [Parameter(Mandatory=$False)]
    [bool]$ReturnTimer = $False
    )
    Process {
        if ($Reset) {
        $CSTMStopWatch.Restart()
        $offset = 0
        } else {
        $offset = $CSTMStopWatch.Elapsed.TotalMilliseconds
        $CSTMStopWatch.start()
        }
        DO {
        Write-progress -Activity $Activity -CurrentOperation $Currentop -Status "Elaspesd Time: $("{0:D2}:{1:D2}" -f $CSTMStopWatch.Elapsed.Minutes,$CSTMStopWatch.Elapsed.seconds)"
        Start-Sleep -Milliseconds 100
        } Until ($CSTMStopWatch.Elapsed.TotalMilliseconds -ge (($Seconds*1000)+$offset))
        if ($Stop) {
        $CSTMStopWatch.stop()
        }
        if($ReturnTimer) {return $CSTMStopWatch}
    }
}
#End

<#
.SYNOPSIS
Write a "Log" message file
.DESCRIPTION
A simple function to write logs to a file

.PARAMETER FilePath
String containing the path of the log file. Mandatory
.PARAMETER Type
String denoting the type of the message e.g. INFO - default, ERROR etc.
.PARAMETER $Message
String containing the message to be logged. Mandatory
.PARAMETER $Append
Boolean denoting wheather the message will be appended to a file or overwrite it essentially creating a new file. Default $true

.EXAMPLE
Write-CSTMLog -FilePath C:\logs\log.txt -message "some message"
writes a message to C:\logs\log.txt in the format [yyyy/MM/dd HH:mm:ss.fffZ][INFO] Some mesage
.EXAMPLE
Write-CSTMLog -FilePath C:\logs\log.txt -message "some message" -Type "ERROR"
writes a message to C:\logs\log.txt in the format [yyyy/MM/dd HH:mm:ss.fffZ][ERROR] Some mesage
#>
Function Write-CSTMLog
{
    [cmdletbinding()]
    Param (
    [Parameter(Mandatory=$True,Position=1)]
    [string]$FilePath,
    [Parameter(Mandatory=$False,Position=2)]
    [String]$Type = "INFO",
    [Parameter(Mandatory=$False,Position=3)]
    $Message = @(),
    [Parameter(Mandatory=$False,Position=4)]
    [bool]$Append = $true,
    [Parameter(Mandatory = $False)]
    [string]$Thread = [System.String]::Empty
    )
    Begin {
        $MaxRetryCount = 10
    }
    Process {
        foreach ($item in $Message) {
            $retryCount = 0
            do {
                if([System.String]::IsNullOrEmpty($Thread)) {
                    $outputText = "[$(Get-Date (Get-Date).ToUniversalTime() -Format "yyyy/MM/dd HH:mm:ss.fffZ")][$Type] $item"
                }
                else {
                    $outputText = "[$(Get-Date (Get-Date).ToUniversalTime() -Format "yyyy/MM/dd HH:mm:ss.fffZ")][$Thread][$Type] $item"
                }
                try {
                    $retry = $false
                    Out-File -FilePath $FilePath -InputObject $outputText -Append:$Append
                }
                catch [System.IO.IOException] {
                    if($retryCount -le $MaxRetryCount) {
                        $retry = $true
                    }
                    $retryCount++
                }
            } while ($retry)
        }
    }
}