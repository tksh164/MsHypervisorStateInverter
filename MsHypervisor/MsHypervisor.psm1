#Requires -RunAsAdministrator

$BCDEDIT_EXE_PATH = 'C:\Windows\System32\bcdedit.exe'

function getHypervisorPresent
{
    [OutputType([bool])]
    param()

    return (Get-WmiObject -Class 'Win32_ComputerSystem').HypervisorPresent
}


function executeCmdline
{
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $CommandFileName,

        [Parameter(Mandatory = $true, Position = 1)][AllowEmptyString()]
        [string] $CommandArguments,

        [switch] $NoNewWindow
    )

    try
    {
        # Prepare the process execution.
        $process = New-Object -TypeName 'System.Diagnostics.Process'
        $process.StartInfo.FileName = $CommandFileName
        $process.StartInfo.Arguments = $CommandArguments
        $process.StartInfo.CreateNoWindow = $NoNewWindow
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true

        # Register the events.
        $stringBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        $eventHandler = {
            if ($EventArgs.Data -ne $null)
            {
                [void] $Event.MessageData.AppendLine($EventArgs.Data)
            }
        }
        $outputEventJob = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action $eventHandler -MessageData $stringBuilder
        $errorEventJob = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action $eventHandler -MessageData $stringBuilder

        # Start the process execution.
        [void] $process.Start()
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        Write-Verbose ('[Start Process] Pid:{0}, "{1}" {2}' -f $process.Id, $FileName, $Arguments)

        # Wait the process execution end.
        $process.WaitForExit()
        Write-Verbose ('[Exit Process] Pid:{0}, ExitCode:{1}' -f $process.Id, $process.ExitCode)

        # Un-register the events.
        Unregister-Event -SourceIdentifier $outputEventJob.Name
        Unregister-Event -SourceIdentifier $errorEventJob.Name

        # Return process execution results.
        return New-Object -TypeName 'System.Management.Automation.PSObject' -Property (@{
            'FileName'  = $FileName
            'Arguments' = $Arguments
            'Pid'       = $process.Id
            'ExitCode'  = $process.ExitCode
            'StartTime' = $process.StartTime
            'ExitTime'  = $process.ExitTime
            'Output'    = $stringBuilder.ToString().Trim()
        })
    }
    catch
    {
        throw $_
    }
}


function getHypervisorLaunchType
{
    [OutputType([string])]
    param()

    $result = executeCmdline -CommandFileName $BCDEDIT_EXE_PATH -CommandArguments '/enum {current}' -NoNewWindow
    
    if ($result.ExitCode -ne 0) { return 'Unknown' }

    return $result.Output -split "`n" | ForEach-Object -Process {
        if ($_ -like 'hypervisorlaunchtype*')
        {
            ($name, $value) = $_.Split(' ', 2, [System.StringSplitOptions]::RemoveEmptyEntries)
            return $value
        }
    }
}


function setHypervisorLaunchType
{
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Auto', 'Off')]
        [string] $Value
    )

    $cmdArgs = ('/set {{current}} hypervisorlaunchtype {0}' -f $Value)
    $result = executeCmdline -CommandFileName $BCDEDIT_EXE_PATH -CommandArguments $cmdArgs -NoNewWindow
    return ($result.ExitCode -eq 0)
}


function restartComputer
{
    [OutputType([void])]
    param()

    $title = 'Restart Computer'
    $message = 'You need to restart the computer for apply the changes.'

    $yes = New-Object -TypeName 'System.Management.Automation.Host.ChoiceDescription' -ArgumentList '&Yes', 'Restart the computer.'
    $no = New-Object -TypeName 'System.Management.Automation.Host.ChoiceDescription' -ArgumentList '&No', 'Do not restart the computer.'
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)

    $defaultChoice = 0
    $result = $host.UI.PromptForChoice($title, $message, $options, $defaultChoice)
    if ($result -eq 0)
    {
        $delaySeconds = 10
        while ($true)
        {
            if ($delaySeconds -le 0)
            {
                Write-Host 'restart'
                Restart-Computer
                break
            }
            Write-Warning -Message ('Restart after {0} seconds.' -f $delaySeconds)
            Start-Sleep -Seconds 1
            $delaySeconds--
        }
    }
}


function Test-MsHypervisor
{
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $present = getHypervisorPresent
    if ($present)
    {
        Write-Verbose -Message 'Hypervisor is currently running.'
    }
    return $present
}


function Enable-MsHypervisor
{
    [CmdletBinding()]
    [OutputType([void])]
    param()

    $currentLaunchType = getHypervisorLaunchType
    if ($currentLaunchType -eq 'Auto')
    {
        Write-Warning -Message 'Hypervisor is already enabled.'
        return
    }

    if ($currentLaunchType -eq 'Off')
    {
        if ((setHypervisorLaunchType -Value 'Auto'))
        {
            restartComputer
        }
        else
        {
            Write-Error -Message 'Unable to enable the hypervisor.'
        }
    }
    else
    {
        Write-Error -Message ('Unknown HypervisorLaunchType: {0}.' -f $currentLaunchType)
    }
}


function Disable-MsHypervisor
{
    [CmdletBinding()]
    [OutputType([void])]
    param()

    $currentLaunchType = getHypervisorLaunchType
    if ($currentLaunchType -eq 'Off')
    {
        Write-Warning -Message 'Hypervisor is already disabled.'
        return
    }

    if ($currentLaunchType -eq 'Auto')
    {
        if ((setHypervisorLaunchType -Value 'Off'))
        {
            restartComputer
        }
        else
        {
            Write-Error -Message 'Unable to disablethe hypervisor.'
        }
    }
    else
    {
        Write-Error -Message ('Unknown HypervisorLaunchType: {0}.' -f $currentLaunchType)
    }
}


Export-ModuleMember -Function @(
    'Test-MsHypervisor',
    'Enable-MsHypervisor',
    'Disable-MsHypervisor'
)
