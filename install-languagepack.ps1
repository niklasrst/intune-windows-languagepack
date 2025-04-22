<#
    .SYNOPSIS
    Windows language pack installer

    .DESCRIPTION
    Install:   %WINDIR%\SysNative\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "install-languagepack.ps1" -install -LanguageCode "en-us"
    Uninstall: %WINDIR%\SysNative\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "install-languagepack.ps1" -uninstall -LanguageCode "en-us"

    .RUNSAS
    SYSTEM

    Language code information:
    https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/available-language-packs-for-windows

    .ENVIRONMENT
    PowerShell 5.0

    .AUTHOR
    Niklas Rast
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $true, ParameterSetName = 'install')]
	[switch]$install,
	[Parameter(Mandatory = $true, ParameterSetName = 'uninstall')]
	[switch]$uninstall,
    [Parameter(Mandatory = $true)]
    [string]$LanguageCode
)

# Script settings
$companyName = ""
$logFile = ('{0}\{1}.log' -f "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs", [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name))

if ($install) {
    Start-Transcript -Path $logFile -Append
    try {
        # Install Language Pack
        try {
            # Install .cab files
            $languagePackPath = Join-Path -Path $PsScriptRoot -ChildPath "*.cab"
            $cabFiles = Get-ChildItem -Path $languagePackPath -File

            if ($cabFiles.Count -eq 0) {
                throw "No .cab files found in $PsScriptRoot."
            }

            foreach ($cabFile in $cabFiles) {
                Write-Verbose "Installing language pack from $($cabFile.FullName)"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList @(
                    "-Command",
                    "Add-WindowsPackage -PackagePath $($cabFile.FullName) -Online -NoRestart"
                ) -NoNewWindow -Wait -PassThru
                
                if ($process.ExitCode -ne 0) {
                    throw "Failed to install $($cabFile.FullName) with exit code $($process.ExitCode)."
                }
            }

            # Create registry key and set value
            $regPath = "HKLM:\Software\$companyName\Windows-LanguagePack"
            if (-Not (Test-Path -Path $regPath)) {
                New-Item -Path "HKLM:\Software\$companyName" -Name "Windows-LanguagePack" -Force | Out-Null
            }
            Set-ItemProperty -Path $regPath -Name "LanguageCode" -Value $LanguageCode -Force
            Write-Verbose "Registry key created at $regPath with LanguagePack value set to $languagePackPath"
        }
        catch {
            $tempErrorMessage = "`r`nCould not add language. Possible issue:`r`n{0}`r`nCancel installation.`r`n" -f $_.Exception.Message
            Write-Error -Message $tempErrorMessage -Category WriteError -ErrorId 1
            return 1
        }
        
        # Set Display Language
        try {
            Write-Verbose "Setting System Language to $LanguageCode"

            # Create task script
            $scriptContent = @"
param([string]`$LanguageCode, [string]`$TaskName)
Start-Transcript -Path C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\SetLanguageAndReboot.log -Append
Set-SystemPreferredUILanguage -Language `$LanguageCode -Verbose
Disable-ScheduledTask -TaskName `$TaskName
Start-Sleep -Seconds 5
Stop-Transcript
Restart-Computer -Force
"@

            $scriptPath = "C:\Windows\Temp\SetLanguageAndReboot.ps1"
            if (-Not (Test-Path -Path $scriptPath)) {
                New-Item -Path $scriptPath -ItemType File -Force | Out-Null
            }
            Set-Content -Path $scriptPath -Value $scriptContent -Force

            # Create a scheduled task to set the system preferred UI language and reboot
            $taskName = "SetSystemLanguageAndReboot"
            $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -LanguageCode `"$LanguageCode`" -TaskName `"$taskName`""
            $taskTrigger = New-ScheduledTaskTrigger -AtStartup
            $taskPrincipal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
            $taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -Compatibility Win8

            Register-ScheduledTask -TaskName $taskName -Action $taskAction -Trigger $taskTrigger -Principal $taskPrincipal -Settings $taskSettings
            Write-Verbose "Scheduled task '$taskName' created to set system language and reboot."
        }
        catch {
            $tempErrorMessage = "`r`nCould not configure language. Possible issue:`r`n{0}`r`nCancel installation.`r`n" -f $_.Exception.Message
            Write-Error -Message $tempErrorMessage -Category WriteError -ErrorId 1
            return 1
        } 
        
        Write-Warning "Reboot required."
        exit 1641
    } catch {
        $PSCmdlet.WriteError($_)
        exit 1
    }
    Stop-Transcript
}

if ($uninstall) {
    Start-Transcript -Path $logFile -Append
    try {
        # Uninstall Language Pack
        try {
            Write-Verbose "Uninstalling Language pack"
            Uninstall-Language -Language $LanguageCode -Verbose
        }
        catch {
            $tempErrorMessage = "`r`nCould not remove language. Possible issue:`r`n{0}`r`nCancel installation.`r`n" -f $_.Exception.Message
            Write-Error -Message $tempErrorMessage -Category WriteError -ErrorId 1
            return 1
        }

        Write-Warning "Reboot required."
        exit 1641
    } catch {
        $PSCmdlet.WriteError($_)
        exit 1
    }
    Stop-Transcript
}