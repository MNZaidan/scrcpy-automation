#region Script Metadata
<#
.SYNOPSIS
    A PowerShell script to automate scrcpy session with an interactive 
    preset manager, device selection, and recording capabilities.

.DESCRIPTION
    This script provides a menu-driven interface to manage scrcpy presets 
    from a JSON configuration file. It offers a powerful set of features 
    to streamline your scrcpy usage, including quick-launch options, 
    device management, and configurable recording settings.

.NOTES
    Version: 2.25
    Requirements:
        - PowerShell 7 or later
        - scrcpy installed and available in PATH
        - ADB installed (typically included with scrcpy)
        - FFmpeg in PATH (optional, for MP4 remuxing)
        - Android device with USB debugging enabled

.PARAMETER DeviceSerial
    Optional. The ADB serial of the device to connect to. If not provided, 
    the script will prompt for device selection.
.PARAMETER Preset
    Optional. Launch scrcpy directly with the specified preset name.
    Supports fuzzy matching - will prompt for confirmation if an exact match isn't found.

.PARAMETER Log
    Enable logging to file.

.PARAMETER NoClear
    Disable screen clearing between menus.

.PARAMETER RealTimeOutput
    Enable real-time output reading (resource intensive). By default, scrcpy output is displayed
    but not captured for logging until the process exits.

.PARAMETER ConfigPath
    Specify a custom path for the configuration file. Defaults to "scrcpy-config.json" in the script directory.

.PARAMETER LogPath
    Specify a custom path for the log file. Defaults to "scrcpy-automation.log" in the script directory.

.EXAMPLE
    .\scrcpy-automation.ps1
    Launches the script and displays the main menu with all options.

.EXAMPLE
    .\scrcpy-automation.ps1 -Preset "Low Latency"
    Searches for a preset matching "Low Latency" (will find "Low Latency / Gaming")
    and asks for confirmation before launching.

.EXAMPLE
    .\scrcpy-automation.ps1 -RealTimeCapture
    Launches the script with real-time scrcpy output capture enabled (more resource intensive).

.LINK
    [scrcpy Automation](https://github.com/MNZaidan/scrcpy-automation)
    [scrcpy Documentation](https://github.com/Genymobile/scrcpy)
    [FFmpeg Download](https://ffmpeg.org/download.html)
#>
#endregion

#region Parameters
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "The ADB serial of the device to connect to.")]
    [string]$DeviceSerial,
    [Parameter(Mandatory = $false, HelpMessage = "Launch scrcpy directly with the specified preset name.")]
    [string]$Preset,
    [Parameter(Mandatory = $false, HelpMessage = "Enable logging to file.")]
    [switch]$Log,
    [Parameter(Mandatory = $false, HelpMessage = "Disable screen clearing between menus.")]
    [switch]$NoClear,
    [Parameter(Mandatory = $false, HelpMessage = "Enable real-time scrcpy output reading (resource intensive).")]
    [switch]$RealTimeCapture,
    [Parameter(Mandatory = $false, HelpMessage = "Specify a custom path for the configuration file.")]
    [string]$ConfigPath,
    [Parameter(Mandatory = $false, HelpMessage = "Specify a custom path for the log file.")]
    [string]$LogPath,
    [Parameter(Mandatory = $false, HelpMessage = "Maximum log file size in KB before rotation.")]
    [int]$LogMaxSize = 50,
    [Parameter(Mandatory = $false, HelpMessage = "Percentage of lines to remove when rotating log (0-100).")]
    [ValidateRange(0,100)]
    [int]$LogTrimPercentage = 15
)
#endregion

#region Global Variables and Defaults
$global:LastAdbOperation = $null
$OutputEncoding = [System.Text.Encoding]::UTF8
$ScriptVersion = "2.25"
$MaxMenuItems = 19 # The maximum number of items to display in menus before scrolling
$DisableClearHost = $NoClear
$PresetProperties = @(
    'name',
    'description',
    'tags',
    'favorite',
    'resolution',
    'videoCodec',
    'videoBitrate',
    'videoBuffer',
    'audioCodec',
    'audioBitrate',
    'audioBuffer',
    'otherOptions'
)
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot "scrcpy-config.json"
}
if (-not $LogPath) {
    $LogPath = Join-Path $PSScriptRoot "scrcpy-automation.log"
}
#endregion

#region Logging and Error Handling Functions

function Write-DetailedLog {
    param (
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$Exception = $null,
        [string]$Level = "INFO",
        [string]$ForegroundColor = $null,
        [switch]$NoConsoleOutput
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    $consoleOutput = "[$Level] $Message"

    if ($Exception) {
        $logEntry += " - Exception: $($Exception.Exception.Message)"
        $logEntry += " (Type: $($Exception.Exception.GetType().FullName))"
    }
    
    if ($Log) {
        try {
            $logEntry | Add-Content -Path $LogPath -Force
        }
        catch {
            if (-not $NoConsoleOutput) {
                Write-Host "LOG ERROR: Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    if (-not $NoConsoleOutput) {
        switch ($Level) {
            "ERROR" { Write-Host $consoleOutput -ForegroundColor Red }
            "WARN"  { Write-Host $consoleOutput -ForegroundColor Yellow }
            "DEBUG" { Write-Host $consoleOutput -ForegroundColor Magenta }
            default {             
                if ($ForegroundColor) {
                    Write-Host $consoleOutput -ForegroundColor $ForegroundColor
                }
                else {
                    Write-Host $consoleOutput
                }
            }
        }
    }
}

function Write-LogOnly {
    param ([string]$Message)
    Write-DetailedLog -Message $Message -Level "DEBUG" -NoConsoleOutput
}

function Write-DebugLog {
    param ([string]$Message)
    if ($DebugPreference -ne "SilentlyContinue") {
        Write-DetailedLog -Message $Message -Level "DEBUG"
    }
}

function Write-WarnLog {
    param (
        [string]$Message,
        [string]$ForegroundColor = "Yellow"
    )
    Write-DetailedLog -Message $Message -Level "WARN"
}

function Write-InfoLog {
    param(
        [string]$Message,
        [string]$ForegroundColor = $null
    )
    Write-DetailedLog -Message $Message -Level "INFO" -ForegroundColor $ForegroundColor
}

function Write-ErrorLog {
    param([string]$Message, [Exception]$Exception = $null)
    if ($Exception) {
        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
            $Exception,
            "ScriptError",
            [System.Management.Automation.ErrorCategory]::NotSpecified,
            $null
        )
        Write-DetailedLog -Message $Message -Exception $errorRecord -Level "ERROR"
    }
    else {
        Write-DetailedLog -Message $Message -Level "ERROR"
    }
}
function Invoke-LogRotation {
    param(
        [string]$LogPath,
        [int]$MaxSizeKB,
        [int]$TrimPercentage
    )

    if (-not (Test-Path $LogPath)) { return }
    try {
        $logFile = Get-Item $LogPath -ErrorAction Stop
        $maxSizeBytes = $MaxSizeKB * 1KB
        
        if ($logFile.Length -gt $maxSizeBytes) {
            $lines = Get-Content $LogPath -ErrorAction Stop
            $totalLines = $lines.Count
            $linesToRemove = [math]::Ceiling($totalLines * ($TrimPercentage / 100))
            
            if ($linesToRemove -gt 0) {
                $remainingLines = $lines | Select-Object -Skip $linesToRemove
                $remainingLines | Set-Content $LogPath -Force -ErrorAction Stop
                Write-DebugLog "Trimmed $linesToRemove lines from log file (exceeded ${MaxSizeKB}KB)"
            }
        }
    }
    catch {
        Write-ErrorLog "Failed to rotate log file: $($_.Exception.Message)"
    }
}
if ($Log) {
    Invoke-LogRotation -LogPath $LogPath -MaxSizeKB $LogMaxSize -TrimPercentage $LogTrimPercentage
}

function Invoke-SafeCommand {
    param (
        [ScriptBlock]$Command,
        [string]$ErrorMessage,
        [switch]$ContinueOnError
    )
    
    try {
        $output = & $Command 2>&1
        if ($LASTEXITCODE -ne 0 -and -not $ContinueOnError) {
            throw "$ErrorMessage (Exit code: $LASTEXITCODE)"
        }
        return $output
    }
    catch {
        Write-ErrorLog -Message $ErrorMessage -Exception $_.Exception
        if (-not $ContinueOnError) {
            throw
        }
        return $null
    }
}
function Compare-ConfigChanges {
    param($oldConfig, $newConfig)
    $changes = @()
    
    $properties = @('recordingPath', 'recordingFormat', 'lastUsedPreset', 'quickLaunchPreset', 'selectedDevice')
    foreach ($prop in $properties) {
        if ($oldConfig.$prop -ne $newConfig.$prop) {
            $changes += "$prop changed from '$($oldConfig.$prop)' to '$($newConfig.$prop)'"
        }
    }
    
    if ($oldConfig.presets.Count -ne $newConfig.presets.Count) {
        $changes += "Preset count changed from $($oldConfig.presets.Count) to $($newConfig.presets.Count)"
    } else {
        for ($i = 0; $i -lt $oldConfig.presets.Count; $i++) {
            $oldPreset = $oldConfig.presets[$i]
            $newPreset = $newConfig.presets[$i]
            
            foreach ($prop in $PresetProperties) {
                if ($oldPreset.$prop -ne $newPreset.$prop) {
                    $changes += "Preset $($i+1) ($($oldPreset.name)) $prop changed from '$($oldPreset.$prop)' to '$($newPreset.$prop)'"
                }
            }
        }
    }
    return $changes
}
#endregion

#region Menu and Input Functions
function Show-Menu {
    param (
        [string]$Title,
        [array]$Options,
        [int]$SelectedIndex = 0,
        [int[]]$HighlightIndices = @(),
        [string]$HighlightColor = "Yellow",
        [int[]]$CategoryIndices = @(),
        [switch]$SkipCategoriesOnNavigate,
        [string[]]$Footer = @(),
        [int[]]$AdditionalReturnKeyCodes = @()
    )
    $isScrollable = $Options.Count -gt $MaxMenuItems

    while ($true) {
        if (-not $DisableClearHost) { Clear-Host }
        Write-Host "$Title`n" -ForegroundColor Cyan

        $displayStart = 0
        $displayCount = $Options.Count

        if ($isScrollable) {
            $displayCount = $MaxMenuItems
            $halfView = [math]::Floor($MaxMenuItems / 2)
            $displayStart = $SelectedIndex - $halfView
            $displayStart = [math]::Max(0, $displayStart)
            $displayStart = [math]::Min($displayStart, $Options.Count - $MaxMenuItems)
        }
        
        $hasTopScroll = $isScrollable -and $displayStart -gt 0
        if ($hasTopScroll) {
            $remainingTop = $displayStart
            Write-Host "   ... ($remainingTop more above) ..." -ForegroundColor Gray
        }
        else {
            Write-Host ""
        }

        for ($i = $displayStart; $i -lt ($displayStart + $displayCount); $i++) {
            if ($i -eq $SelectedIndex) {
                Write-Host " > $($Options[$i])" -ForegroundColor Black -BackgroundColor White
            }
            elseif ($CategoryIndices -contains $i) {
                Write-Host "   $($Options[$i])" -ForegroundColor Magenta
            }
            elseif ($HighlightIndices -contains $i) {
                Write-Host "   $($Options[$i])" -ForegroundColor $HighlightColor
            }
            else {
                Write-Host "   $($Options[$i])"
            }
        }

        if (-not $isScrollable) {
            $blankLinesNeeded = $MaxMenuItems - $Options.Count
            for ($i = 0; $i -lt $blankLinesNeeded; $i++) {
                Write-Host ""
            }
        }

        $hasBottomScroll = $isScrollable -and ($displayStart + $MaxMenuItems) -lt $Options.Count
        if ($hasBottomScroll) {
            $remainingBottom = $Options.Count - ($displayStart + $MaxMenuItems)
            Write-Host "   ... ($remainingBottom more below) ..." -ForegroundColor Gray
        }
        else {
            Write-Host ""
        }

        foreach ($line in $Footer) {
            Write-Host $line -ForegroundColor Blue
        }

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        $returnKeys = @(27, 88, 13) + $AdditionalReturnKeyCodes  # Escape, X, Enter + other
        switch ($key.VirtualKeyCode) {
            27 { return @{ Key = 'Escape'; Index = -1; KeyInfo = $key } } # Escape
            88 { return @{ Key = 'x'; Index = -1; KeyInfo = $key } } # 'x' key
            13 { return @{ Key = 'Enter'; Index = $SelectedIndex; KeyInfo = $key } } # Enter
            38 {
                # Up Arrow
                if ($SkipCategoriesOnNavigate) {
                    $prevIndex = -1
                    for ($i = $SelectedIndex - 1; $i -ge 0; $i--) {
                        if (-not ($CategoryIndices -contains $i)) {
                            $prevIndex = $i
                            break
                        }
                    }
                    if ($prevIndex -ne -1) { $SelectedIndex = $prevIndex }
                }
                else {
                    if ($SelectedIndex -gt 0) { $SelectedIndex-- }
                }
                continue
            }
            40 {
                # Down Arrow
                if ($SkipCategoriesOnNavigate) {
                    $nextIndex = -1
                    for ($i = $SelectedIndex + 1; $i -lt $Options.Count; $i++) {
                        if (-not ($CategoryIndices -contains $i)) {
                            $nextIndex = $i
                            break
                        }
                    }
                    if ($nextIndex -ne -1) { $SelectedIndex = $nextIndex }
                }
                else {
                    if ($SelectedIndex -lt ($Options.Count - 1)) { $SelectedIndex++ }
                }
                continue
            }
            default { # Other keys
                if ($returnKeys -contains $key.VirtualKeyCode) {
                    return @{ Key = 'Default'; Index = $SelectedIndex; KeyInfo = $key }
                }
            }
        }
    }
}

function Read-Input {
    param (
        [string]$Prompt,
        [string]$DefaultValue = "",
        [switch]$HideDefaultValue
    )

    $userInput = [System.Text.StringBuilder]::new($DefaultValue)
    $cursorIndex = $DefaultValue.Length
    
    Write-Host "`n$Prompt" -NoNewline -ForegroundColor Yellow
    if (-not [string]::IsNullOrEmpty($DefaultValue) -and -not $HideDefaultValue) {
        Write-Host " (Default: '$DefaultValue'): " -NoNewline -ForegroundColor Green
    }
    else {
        Write-Host ": " -NoNewline -ForegroundColor Green
    }

    $promptPosition = $Host.UI.RawUI.CursorPosition

    function Update-InputBufferDisplay {
        param (
            [string]$CurrentInput,
            [int]$CurrentCursorIndex
        )
        
        $Host.UI.RawUI.CursorPosition = $promptPosition
        
        $lineLength = $Host.UI.RawUI.BufferSize.Width - $promptPosition.X
        Write-Host (" " * $lineLength) -NoNewline
        
        $Host.UI.RawUI.CursorPosition = $promptPosition
        Write-Host $CurrentInput -NoNewline
        
        $totalChars = $promptPosition.X + $CurrentCursorIndex
        $newX       = $totalChars % $Host.UI.RawUI.BufferSize.Width
        $newY       = $promptPosition.Y + [math]::Floor($totalChars / $Host.UI.RawUI.BufferSize.Width)
        
        $Host.UI.RawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new($newX, $newY)
    }

    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex

    while ($true) {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        switch ($key.VirtualKeyCode) {
            13 {
                # Enter key
                Write-Host ""
                Write-DebugLog "User input completed: $($userInput.ToString())"
                return $userInput.ToString() 
            }
            27 {
                # Escape key
                Write-Host ""
                Write-InfoLog "User canceled input"
                return $null
            }
            8 {
                # Backspace
                if ($cursorIndex -gt 0) {
                    $userInput.Remove($cursorIndex - 1, 1) | Out-Null
                    $cursorIndex--
                    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex
                }
            }
            46 {
                # Delete
                if ($cursorIndex -lt $userInput.Length) {
                    $userInput.Remove($cursorIndex, 1) | Out-Null
                    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex
                }
            }
            37 {
                # Left Arrow
                if ($cursorIndex -gt 0) {
                    $cursorIndex--
                    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex
                }
            }
            39 {
                # Right Arrow
                if ($cursorIndex -lt $userInput.Length) {
                    $cursorIndex++
                    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex
                }
            }
            default {
                # Handle character input
                if ($key.Character -ne "`0") {
                    $userInput.Insert($cursorIndex, $key.Character) | Out-Null
                    $cursorIndex++
                    Update-InputBufferDisplay -CurrentInput $userInput.ToString() -CurrentCursorIndex $cursorIndex
                }
            }
        }
    }
}

function Wait-Enter {
    Write-Host "Press Enter to continue..." -NoNewline
    while ($true) {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        if ($key.VirtualKeyCode -eq 13) {  # Enter key
            Write-Host ""
            break
        }
    }
}
#endregion

#region ADB Device Management Functions
function Get-AdbDeviceList {
    param ([string]$adbPath)
    
    $deviceList = @()
    try {
        $adbOutput = Invoke-SafeCommand -Command { & $adbPath devices -l } -ErrorMessage "Failed to run adb devices command"
        if ($null -ne $adbOutput) {
            $adbOutput | ForEach-Object {
                if ($_ -is [string] -and $_.Trim() -ne "") {
                    Write-DebugLog "  $($_.Trim())"
                }
            }
        }
        if ($null -eq $adbOutput) { return $null }
    }
    catch {
        Write-ErrorLog "ADB command failed completely" $_.Exception
        return $null
    }
    
    $adbOutput | Where-Object { $_ -is [string] } | Select-Object -Skip 1 | ForEach-Object {
        $line = $_.Trim()
        if ($line -and $line -match '^(\S+)\s+(\S+)\b') {
            $serial = $matches[1]
            $state = $matches[2]
            
            if ($state -in @('device', 'offline', 'unauthorized')) {
                $modelMatch = $line | Select-String -Pattern 'model:([^\s]+)'
                
                if ($modelMatch) {
                    $model = $modelMatch.Matches.Groups[1].Value
                    $deviceList += [pscustomobject]@{ Serial = $serial; Model = $model; State = $state }
                }
                else {
                    $deviceList += [pscustomobject]@{ Serial = $serial; Model = ""; State = $state }
                }
            }
        }
    }
    return $deviceList
}
function Get-DeviceDisplayName {
    param ([string]$adbPath, [string]$deviceSerial)
    
    $deviceSerial = $deviceSerial.Trim()
    
    if ([string]::IsNullOrWhiteSpace($deviceSerial)) { 
        Write-DebugLog "No device selected"
        return "No device selected" 
    }
    
    Write-DebugLog "Checking status for selected device: '$deviceSerial'"
    
    $device = $null
    $retryCount = 0
    $maxRetries = 3
    
    while ($retryCount -lt $maxRetries -and $null -eq $device) {
        $deviceList = Get-AdbDeviceList -adbPath $adbPath
        $device = $deviceList | Where-Object { 
            $_.Serial.Trim() -eq $deviceSerial -and $_.State -eq 'device'
        } | Select-Object -First 1
        
        if ($null -eq $device) {
            $retryCount++
            Write-DebugLog "Selected device not found in 'device' state (attempt $retryCount/$maxRetries)"
            Start-Sleep -Milliseconds 500
        }
    }
    
    if ($device) {
        Write-DebugLog "Selected device found with state: $($device.State)"
        if ($device.Model) {
            return "$($device.Model) ($($device.Serial))"
        }
        else {
            return "$($device.Serial)"
        }
    }
    else { 
        $anyStateDevice = (Get-AdbDeviceList -adbPath $adbPath) | Where-Object { $_.Serial.Trim() -eq $deviceSerial } | Select-Object -First 1
        if ($anyStateDevice) {
            Write-DebugLog "Selected device found but in state: $($anyStateDevice.State)"
            return "$deviceSerial [$($anyStateDevice.State)]"
        }
        Write-DebugLog "Selected device not found"
        return "$deviceSerial (disconnected)" 
    }
}

function Invoke-AdbTcpip {
    param ([string]$adbPath)
    Write-Host "`nThis requires your device to be currently connected via USB." -ForegroundColor Yellow
    $port = Read-Input -Prompt "Enter port number" -DefaultValue "5555"
    if ([string]::IsNullOrWhiteSpace($port)) { return }
    if (-not $DisableClearHost) { Clear-Host }
    Write-InfoLog "Running: adb tcpip $port"
    try {
        $result = Invoke-SafeCommand -Command { & $adbPath tcpip $port } -ErrorMessage "Failed to run adb tcpip command" -ContinueOnError
        Write-InfoLog $result
        $global:LastAdbOperation = Get-Date
        Write-InfoLog "`nADB tcpip command finished. " -ForegroundColor Green
        Write-Host "You can now connect your device wirelessly using `adb connect <device-ip>:$port`."
    }
    catch {
        Write-ErrorLog "An error occurred while running adb tcpip." $_.Exception
    }
    Wait-Enter
}

function Invoke-AdbPair {
    param ([string]$adbPath)
    Write-Host "`nWarning: This is for pairing a device wirelessly for the first time." -ForegroundColor Yellow
    Write-Host "You will need to enter the device's IP address and pairing code." -ForegroundColor Yellow
    $ip = Read-Input -Prompt "Enter device IP address" -DefaultValue "192.168.1.100" -HideDefaultValue
    if ([string]::IsNullOrWhiteSpace($ip)) { return }
    $port = Read-Input -Prompt "Enter pairing port" -DefaultValue "45389" -HideDefaultValue
    if ([string]::IsNullOrWhiteSpace($port)) { return }
    $code = Read-Input -Prompt "Enter pairing code"
    if ([string]::IsNullOrWhiteSpace($code)) { return }
    if (-not $DisableClearHost) { Clear-Host }
    Write-InfoLog "Running: adb pair $ip`:$port"
    try {
        $result = Invoke-SafeCommand -Command { & $adbPath pair "$ip`:$port" $code } -ErrorMessage "Failed to run adb pair command" -ContinueOnError
        Write-InfoLog $result
        Write-InfoLog "`nADB pair command finished." -ForegroundColor 
        $global:LastAdbOperation = Get-Date
    }
    catch {
        Write-ErrorLog "An error occurred while running adb pair." $_.Exception
    }
    Wait-Enter
}

function Invoke-AdbConnect {
    param ([string]$adbPath)
    Write-InfoLog "Initiating ADB connect"
    Write-Host "`nConnect to a wireless device." -ForegroundColor Yellow
    $ip = Read-Input -Prompt "Enter device IP address" -DefaultValue "192.168.1.100" -HideDefaultValue
    if ([string]::IsNullOrWhiteSpace($ip)) { 
        Write-DebugLog "User canceled IP input"
        return 
    }
    $port = Read-Input -Prompt "Enter ADB port" -DefaultValue "5555"
    if ([string]::IsNullOrWhiteSpace($port)) { 
        Write-DebugLog "User canceled port input"
        return 
    }
    if (-not $DisableClearHost) { Clear-Host }
    Write-InfoLog "Running: adb connect $ip`:$port"
    try {
        $result = Invoke-SafeCommand -Command { & $adbPath connect "$ip`:$port" } -ErrorMessage "Failed to run adb connect command" -ContinueOnError
        Write-InfoLog $result
        $global:LastAdbOperation = Get-Date
    }
    catch {
        Write-ErrorLog "An error occurred while running adb connect." $_.Exception
    }
    Wait-Enter
}

function Invoke-AdbAutoConnect {
    param ([string]$adbPath)
    Write-Host "`nAttempting to find and connect to an ADB device on your local network..." -ForegroundColor Yellow
    
    $port = 5555
    $found = $false

    Write-InfoLog "Scanning network neighbors using Get-NetNeighbor..."
    try {
        $neighbors = Get-NetNeighbor -State Reachable | Where-Object { $_.IPAddress -like "192.168.*" }
        
        if ($neighbors) {
            foreach ($neighbor in $neighbors) {
                $ip = $neighbor.IPAddress
                Write-InfoLog "Trying: adb connect $ip`:$port"
                $result = Invoke-SafeCommand -Command { & $adbPath connect "$ip`:$port" } -ErrorMessage "Failed to connect to $ip" -ContinueOnError
                Write-InfoLog $result
                $global:LastAdbOperation = Get-Date
                
                if ($result -match "connected to") {
                    Write-InfoLog "Successfully connected to $ip`:$port" -ForegroundColor Green
                    $found = $true
                    break
                }
            }
        } else {
            Write-InfoLog "No devices found with Get-NetNeighbor, trying ARP fallback..."
            throw "No devices found"
        }
    }
    catch {
        Write-InfoLog "Using ARP table scan as fallback..."
        $arpOutput = Invoke-SafeCommand -Command { & arp -a } -ErrorMessage "Failed to run arp command" -ContinueOnError
        
        foreach ($line in $arpOutput) {
            if ($line.Trim() -ne "" -and $line -match '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b') {
                $parts = $line.Trim() -split '\s+'
                $ip = $parts[0]
                
                Write-InfoLog "Trying: adb connect $ip`:$port"
                $result = Invoke-SafeCommand -Command { & $adbPath connect "$ip`:$port" } -ErrorMessage "Failed to connect to $ip" -ContinueOnError
                Write-InfoLog $result
                
                if ($result -match "connected to") {
                    Write-InfoLog "Successfully connected to $ip`:$port" -ForegroundColor Green
                    $found = $true
                    break
                }
            }
        }
    }

    if (-not $found) {
        Write-WarnLog "No device could be connected using network discovery."
        Write-Host "Please ensure your device is connected to the same network and ADB over WiFi is enabled." -ForegroundColor Yellow
    }
    Wait-Enter
}

function Invoke-AdbKillServer {
    param ([string]$adbPath)
    if (-not $DisableClearHost) { Clear-Host }
    Write-InfoLog "Running: adb kill-server"
    try {
        $result = Invoke-SafeCommand -Command { & $adbPath kill-server } -ErrorMessage "Failed to run adb kill-server command" -ContinueOnError
        Write-InfoLog $result
        Write-InfoLog "adb kill-server executed" 
        Write-Host "`nThe script might be a bit slow for a while." -ForegroundColor Red
        $global:LastAdbOperation = Get-Date
    }
    catch {
        Write-ErrorLog "An error occurred while running adb kill-server." $_.Exception
    }
    Wait-Enter
}

function Show-AdbOptionsMenu {
    param([string]$adbPath)
    while ($true) {
        $options = @(
            "adb connect (auto)",
            "adb connect (manual)"
            "adb pair",
            "adb tcpip"
            "adb kill-server",
            "Back"
        )
        $menuResult = Show-Menu -Title "ADB Connection Options" -Options $options -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Back")
        $choiceIndex = $menuResult.Index

        if ($choiceIndex -eq -1 -or $choiceIndex -eq ($options.Count - 1)) {
            return
        }

        switch ($choiceIndex) {
            0 { Invoke-AdbAutoConnect -adbPath $adbPath }
            1 { Invoke-AdbConnect -adbPath $adbPath }
            2 { Invoke-AdbPair -adbPath $adbPath }
            3 { Invoke-AdbTcpip -adbPath $adbPath }
            4 { Invoke-AdbKillServer -adbPath $adbPath }
        }
    }
}

function Show-DeviceSelection {
    param ([string]$adbPath, [string]$currentDevice = "")
    
    Write-DebugLog "Showing device selection menu"
    
    while ($true) {
        $deviceList = Get-AdbDeviceList -adbPath $adbPath
        
        $options = @("Refresh Device List", "ADB Options")
        
        if ($null -eq $deviceList -or $deviceList.Count -eq 0) {
            Write-WarnLog "No ADB devices found. Please connect a device and ensure USB debugging is enabled."
            Start-Sleep -Seconds 1
            $options += "Back"
        }
        else {
            $deviceList | ForEach-Object {
                $displayText = if ($_.Model) { "$($_.Model) ($($_.Serial))" } else { $_.Serial }
                if ($_.State -ne 'device') {
                    $displayText = "$displayText [$($_.State)]"
                }
                if ($_.Serial -eq $currentDevice) {
                    $displayText = "[SELECTED] $displayText"
                }
                $options += $displayText
            }
            $options += "Back"
        }
        
        $menuResult  = Show-Menu -Title "Select a Device" -Options $options -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Back")
        $choiceIndex = $menuResult.Index

        if ($choiceIndex -eq -1) {
            Write-DebugLog "User canceled device selection via ESC/X key"
            return $null
        }
        elseif ($choiceIndex -eq ($options.Count - 1)) {
            Write-DebugLog "User selected 'Back' option in device selection menu"
            return $null
        }
        elseif ($choiceIndex -eq 0) {
            Write-DebugLog "User refreshed device list"
            continue
        }        
        elseif ($choiceIndex -eq 1) {
            Write-DebugLog "User selected ADB options"
            Show-AdbOptionsMenu -adbPath $adbPath
            continue
        }
        else {
            $selectedDeviceIndex = $choiceIndex - 2
            $selectedDevice = $deviceList[$selectedDeviceIndex].Serial.Trim()
            Write-DebugLog "User selected device: $selectedDevice (Index: $selectedDeviceIndex)"
            
            if ($deviceList[$selectedDeviceIndex].State -ne 'device') {
                Write-WarnLog "Device is in $($deviceList[$selectedDeviceIndex].State) state."
                Write-Host "It may take a moment to become ready..." -ForegroundColor Yellow
                
                $attempts = 0
                $maxAttempts = 5
                $deviceReady = $false
                
                while ($attempts -lt $maxAttempts -and -not $deviceReady) {
                    Start-Sleep -Seconds 1
                    $refreshedList = Get-AdbDeviceList -adbPath $adbPath
                    $refreshedDevice = $refreshedList | Where-Object { $_.Serial -eq $selectedDevice }
                    
                    if ($refreshedDevice -and $refreshedDevice.State -eq 'device') {
                        $deviceReady = $true
                        Write-InfoLog "Device is now ready!" -ForegroundColor Green
                    }
                    $attempts++
                }
                
                if (-not $deviceReady) {
                    Write-ErrorLog "Device did not become ready. Please check connection."
                    Start-Sleep -Seconds 2
                    continue
                }
            }
            
            Write-InfoLog "User selected device: $selectedDevice"
            return $selectedDevice
        }
    }
}
#endregion

#region Configuration Management Functions
function Get-Config {
    if (-not (Test-Path $ConfigPath)) {
        Write-DebugLog "Config file not found. Creating a default one."
        $defaultConfig = [ordered]@{
            recordingPath     = Join-Path $PSScriptRoot "recordings"
            recordingFormat   = "AlwaysMKV"
            lastUsedPreset    = ""
            quickLaunchPreset = ""
            selectedDevice    = ""
            presets           = @(
                [ordered]@{
                    name         = "Default"
                    description  = "Standard mirroring settings."
                    favorite     = $false
                    resolution   = "720"
                    videoCodec   = "h264"
                    videoBitrate = "8M"
                    videoBuffer  = ""
                    audioCodec   = "opus"
                    audioBitrate = "128K"
                    audioBuffer  = "50"
                    otherOptions = ""
                }
            )
        }
        try {
            $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath -ErrorAction Stop
        }
        catch {
            Write-ErrorLog "Failed to create default config file" $_.Exception
            return $null
        }
    }
    
    try {
        $jsonContent = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
        $loadedConfig = $jsonContent | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        
        $config = [pscustomobject]@{
            recordingPath     = if ($loadedConfig.ContainsKey('recordingPath')) { $loadedConfig.recordingPath } else { Join-Path $PSScriptRoot "recordings" }
            recordingFormat   = if ($loadedConfig.ContainsKey('recordingFormat')) { $loadedConfig.recordingFormat } else { "AlwaysMKV" }
            lastUsedPreset    = if ($loadedConfig.ContainsKey('lastUsedPreset')) { $loadedConfig.lastUsedPreset } else { "" }
            quickLaunchPreset = if ($loadedConfig.ContainsKey('quickLaunchPreset')) { $loadedConfig.quickLaunchPreset } else { "" }
            selectedDevice    = if ($loadedConfig.ContainsKey('selectedDevice')) { $loadedConfig.selectedDevice } else { "" }
            presets           = if ($loadedConfig.ContainsKey('presets')) {
                $loadedConfig.presets | ForEach-Object {
                    $newPreset = [ordered]@{}
                    foreach ($prop in $PresetProperties) {
                        $defaultValue    = if ($prop -eq 'favorite') { $false } else { "" }
                        $newPreset.$prop = if ($_.ContainsKey($prop)) { $_.$prop } else { $defaultValue }
                    }
                    [pscustomobject]$newPreset
                }
            }
            else { @() }
        }
        
        return $config
    }
    catch {
        Write-ErrorLog "Error reading or parsing config file: $ConfigPath" $_.Exception
        
        try {
            $backupPath = $ConfigPath + ".corrupt." + (Get-Date -Format "yyyyMMdd_HHmmss")
            Copy-Item -Path $ConfigPath -Destination $backupPath -ErrorAction Stop
            Write-InfoLog "Backup of corrupt config created at: $backupPath"
        }
        catch {
            Write-ErrorLog "Failed to create backup of corrupt config" $_.Exception
        }
        
        try {
            $defaultConfig = [ordered]@{
                recordingPath     = Join-Path $PSScriptRoot "recordings"
                recordingFormat   = "AlwaysMKV"
                lastUsedPreset    = ""
                quickLaunchPreset = ""
                selectedDevice    = ""
                presets           = @(
                    [ordered]@{
                        name         = "Default"
                        description  = "Standard mirroring settings."
                        favorite     = $false
                        resolution   = "720"
                        videoCodec   = "h264"
                        videoBitrate = "8M"
                        videoBuffer  = ""
                        audioCodec   = "opus"
                        audioBitrate = "128K"
                        audioBuffer  = "50"
                        otherOptions = ""
                    }
                )
            }
            $defaultConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath -ErrorAction Stop
            Write-InfoLog "Created new default config file due to corruption"
            return Get-Config
        }
        catch {
            Write-ErrorLog "Failed to create new default config file" $_.Exception
            return $null
        }
    }
}

function Save-Config {
    param ($config)
    
    $oldConfig = $null
    if (Test-Path $ConfigPath) {
        try {
            $jsonContent = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
            $oldConfig   = $jsonContent | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        }
        catch {
            Write-DebugLog "Could not read old config for comparison: $($_.Exception.Message)"
        }
    }
    
    try {
        $deviceValue = if ($null -eq $config.selectedDevice) { 
            "" 
        } elseif ($config.selectedDevice -is [array]) {
            $firstValid = $config.selectedDevice | Where-Object { -not [string]::IsNullOrEmpty($_) } | Select-Object -First 1
            if ($null -eq $firstValid) { "" } else { $firstValid.ToString().Trim() }
        } else {
            $config.selectedDevice.ToString().Trim()
        }
        
        $sanitizedConfig = [ordered]@{
            recordingPath     = $config.recordingPath
            recordingFormat   = $config.recordingFormat
            lastUsedPreset    = $config.lastUsedPreset
            quickLaunchPreset = $config.quickLaunchPreset
            selectedDevice    = $deviceValue
            presets           = @()
        }
        
        foreach ($preset in $config.presets) {
            $sanitizedPreset = [ordered]@{}
            
            foreach ($prop in $PresetProperties) {
                if ($preset.PSObject.Properties.Name -contains $prop) {
                    $value = $preset.$prop

                    if ($prop -eq 'favorite') {
                        if ($value -eq $true) {
                            $sanitizedPreset.$prop = $true
                        }
                    }
                    else {
                        if (-not ([string]::IsNullOrWhiteSpace($value))) {
                            $sanitizedPreset.$prop = $value
                        }
                    }
                }
            }
            $sanitizedConfig.presets += [pscustomobject]$sanitizedPreset
        }
        
        $hasChanges = $false
        if ($oldConfig) {
            $changes = Compare-ConfigChanges -oldConfig $oldConfig -newConfig $sanitizedConfig
            if ($changes.Count -gt 0) {
                $hasChanges = $true
                if ($DebugPreference -ne "SilentlyContinue") {
                    Write-DebugLog "Configuration changed:"
                    foreach ($change in $changes) {
                        Write-DebugLog "  $change"
                    }
                }
            } else {
                Write-DebugLog "No configuration changed"
                return
            }
        } else {
            $hasChanges = $true
        }

        if ($hasChanges) {
            $sanitizedConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigPath -ErrorAction Stop
            Write-InfoLog "Configuration saved successfully." -ForegroundColor Green
        }
    }
    catch {
        Write-ErrorLog "Error saving configuration to '$ConfigPath'." $_.Exception
    }
}

function Show-RecordingOptions {
    param ($config)
    
    $selectedIndex = 0
    while ($true) {
        $currentFormat = if ($config.recordingFormat) { 
            if ($config.recordingFormat -eq "RemuxToMP4") { "Record in MKV then remux to MP4" } else { "Always MKV" }
        }
        else { "Always MKV" }
        $currentPath = $config.recordingPath
        
        $options = @(
            "Change Recording Path: $currentPath",
            "Recording Format: $currentFormat",
            "Back"
        )
        
        $menuResult = Show-Menu -Title "Recording Options" -Options $options -SelectedIndex $selectedIndex -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Back")
        $choiceIndex = $menuResult.Index
        
        if ($choiceIndex -eq -1 -or $choiceIndex -eq ($options.Count - 1)) {
            return $config
        }
        
        $selectedIndex = $choiceIndex
        
        switch ($choiceIndex) {
            0 {
                $newPath = Read-Input -Prompt "Enter new recording save path" -DefaultValue $currentPath -HideDefaultValue
                
                if (-not [string]::IsNullOrWhiteSpace($newPath)) {
                    if (-not (Test-Path $newPath -PathType Container)) {
                        $confirm = Read-Input -Prompt "Path '$newPath' does not exist. Create it? (y/n)" -DefaultValue "y" -HideDefaultValue
                        if ($confirm -eq 'y') {
                            try {
                                New-Item -Path $newPath -ItemType Directory -ErrorAction Stop | Out-Null
                                Write-InfoLog "Directory created." -ForegroundColor Green
                            }
                            catch {
                                Write-ErrorLog "Error creating directory: $($_.Exception.Message)"
                                Start-Sleep -Seconds 2
                                continue
                            }
                        }
                        else { continue }
                    }
                    $config.recordingPath = $newPath
                    Save-Config $config
                    Write-InfoLog "Recording path updated to: $newPath" -ForegroundColor Green
                    Start-Sleep -Seconds 2
                }
            }
            1 {
                $formatOptions = @("Always MKV", "Record in MKV then remux to MP4")
                $currentFormatIndex = if ($config.recordingFormat -eq "RemuxToMP4") { 1 } else { 0 }
                $formatMenuResult = Show-Menu -Title "Select Recording Format" -Options $formatOptions -SelectedIndex $currentFormatIndex -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Back")
                $formatChoice = $formatMenuResult.Index

                if ($formatChoice -eq -1) { continue }
                
                if ($formatChoice -eq 1) {
                    $ffmpegPath = Get-Command ffmpeg -ErrorAction SilentlyContinue
                    if (-not $DisableClearHost) { Clear-Host }
                    Write-Host "=== IMPORTANT NOTICE ===" -ForegroundColor Red
                    Write-Host "The 'Remux to MP4' option requires FFmpeg to be installed." -ForegroundColor Yellow
                    if ($ffmpegPath) {
                        Write-Host "FFmpeg Status: FOUND at $($ffmpegPath.Path)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "FFmpeg Status: NOT FOUND in your system's PATH." -ForegroundColor Red
                    }
                    Write-Host "This process re-encodes the audio to AAC to ensure compatibility." -ForegroundColor Cyan
                    Write-Host "If FFmpeg is not available, remuxing will fail and the file will remain MKV." -ForegroundColor Cyan
                    Write-Host ""
                    $confirm = Read-Input -Prompt "Do you want to continue with this option? (y/n)" -DefaultValue "y" -HideDefaultValue
                    
                    if ($confirm -ne 'y') { continue }
                }
                
                $config.recordingFormat = if ($formatChoice -eq 1) { "RemuxToMP4" } else { "AlwaysMKV" }
                Save-Config $config
                Write-InfoLog "Recording format set to: $($formatOptions[$formatChoice])" -ForegroundColor Green
                Start-Sleep -Seconds 2
            }
        }
    }
}

function Find-IsCategory {
    param ([psobject]$Preset)
    if ($null -eq $Preset -or -not $Preset.PSObject.Properties.Name.Contains('name')) { return $false }
    return $Preset.name -like "-*-"
}

function Get-FormattedPresetList {
    param (
        [Parameter(Mandatory = $true)]
        $Config
    )
    
    $quickLaunchPresetName = $Config.quickLaunchPreset
    $maxNameLength = ($Config.presets.name | Measure-Object -Maximum -Property Length).Maximum
    if ($null -eq $maxNameLength) { $maxNameLength = 0 }

    return $Config.presets | ForEach-Object {
        $description = $_.description
        if ($null -ne $description -and $description.Length -gt 80) {
            $description = $description.Substring(0, 80) + "..."
        }
        $name = $_.name.PadRight($maxNameLength)
        $FavoriteStar = if ($_.PSObject.Properties.Name -contains 'favorite' -and $_.favorite) { '★ ' } else { '  ' }
        $quickLaunchStar = if ($_.name -eq $quickLaunchPresetName) { '► ' } else { '  ' }
        
        "$quickLaunchStar$FavoriteStar$name - $description"
    }
    if ($null -eq $Config.presets -or $Config.presets.Count -eq 0) {
        return @()
    }
}

function Build-ScrcpyArguments {
    param (
        [Parameter(Mandatory = $true)]
        $SelectedPreset,
        [Parameter(Mandatory = $true)]
        [string]$SelectedDevice
    )
    
    if ([string]::IsNullOrEmpty($SelectedDevice)) {
        throw "SelectedDevice parameter cannot be empty"
    }
    
    $finalArgs = @("--serial", $SelectedDevice)
    
    
    if (-not [string]::IsNullOrEmpty($SelectedPreset.resolution)) { $finalArgs += "-m", $SelectedPreset.resolution }
    if (-not [string]::IsNullOrEmpty($SelectedPreset.videoCodec)) { $finalArgs += "--video-codec", $SelectedPreset.videoCodec }
    if (-not [string]::IsNullOrEmpty($SelectedPreset.videoBitrate)) { $finalArgs += "--video-bit-rate", $SelectedPreset.videoBitrate }
    if ($SelectedPreset.videoBuffer -gt 0) { $finalArgs += "--video-buffer", $SelectedPreset.videoBuffer.ToString() }
    if (-not [string]::IsNullOrEmpty($SelectedPreset.audioCodec)) { $finalArgs += "--audio-codec", $SelectedPreset.audioCodec }
    if (-not [string]::IsNullOrEmpty($SelectedPreset.audioBitrate)) { $finalArgs += "--audio-bit-rate", $SelectedPreset.audioBitrate }
    if ($SelectedPreset.audioBuffer -gt 0) { $finalArgs += "--audio-buffer", $SelectedPreset.audioBuffer.ToString() }
    if (-not [string]::IsNullOrEmpty($SelectedPreset.otherOptions)) { $finalArgs += $SelectedPreset.otherOptions.Split(' ') }

    return $finalArgs
}
#endregion

#region Preset Management Functions
function Show-PresetEditor {
    param (
        [pscustomobject]$Preset = $null,
        [array]$ExistingPresets
    )
    $isNew = ($null -eq $Preset)
    if ($isNew) {
        $title = "Add New Preset"
        $originalName = $null
        $orderedPreset = [ordered]@{}
        foreach ($prop in $PresetProperties) {
            $orderedPreset.$prop = if ($prop -eq 'favorite') { $false } else { "" }
        }
        $preset = [pscustomobject]$orderedPreset
    }
    else {
        $title = "Edit Preset: $($Preset.name)"
        $originalName = $Preset.name
        $orderedPreset = [ordered]@{}
        foreach ($prop in $PresetProperties) {
            $orderedPreset.$prop = if ($Preset.PSObject.Properties.Name -contains $prop) {
                $Preset.$prop
            }
            else {
                if ($prop -eq 'favorite') { $false } else { "" }
            }
        }
        $preset = $orderedPreset
    }

    $fields = @(
        @{ Name = 'name';           Prompt = "Preset or -Category- name" },
        @{ Name = 'description';    Prompt = 'Description' },
        @{ Name = 'tags';           Prompt = 'Tags (comma-separated)' },
        @{ Name = 'favorite';       Prompt = 'Favorite (true/false)' },
        @{ Name = 'resolution';     Prompt = 'Max Resolution [-m] (e.g., 1080)' },
        @{ Name = 'videoCodec';     Prompt = '[--video-codec] (h264, h265, av1)' },
        @{ Name = 'videoBitrate';   Prompt = '[--video-bit-rate] (e.g., 8M)' },
        @{ Name = 'videoBuffer';    Prompt = '[--video-buffer] (ms)' },
        @{ Name = 'audioCodec';     Prompt = '[--audio-codec] (opus, aac, flac)' },
        @{ Name = 'audioBitrate';   Prompt = '[--audio-bit-rate] (e.g., 128K)' },
        @{ Name = 'audioBuffer';    Prompt = '[--audio-buffer] (ms)' },
        @{ Name = 'otherOptions';   Prompt = 'Other scrcpy arguments' }
    )

    $maxPromptLength = ($fields.Prompt | Measure-Object -Maximum -Property Length).Maximum
    $currentField = 0
    while ($true) {
        if (-not $DisableClearHost) { Clear-Host }
        Write-Host "$title`n" -ForegroundColor Cyan
        
        $terminalWidth = $Host.UI.RawUI.BufferSize.Width
        $valueStartColumn = $maxPromptLength + 6  # 3 spaces + prompt + ": "
        $maxValueWidth = $terminalWidth - $valueStartColumn - 1
        
        for ($i = 0; $i -lt $fields.Count; $i++) {
            $field = $fields[$i]
            $value = if ($null -ne $preset.($field.Name)) { $preset.($field.Name) } else { "[Empty]" }
            
            $wrappedValues = @()
            if ($value.Length -gt $maxValueWidth) {
                $words = $value -split ' '
                $currentLine = ""
                
                foreach ($word in $words) {
                    if (($currentLine.Length + $word.Length + 1) -le $maxValueWidth) {
                        $currentLine += if ($currentLine -eq "") { $word } else { " $word" }
                    } else {
                        if ($currentLine -ne "") { $wrappedValues += $currentLine }
                        $currentLine = $word
                    }
                }
                if ($currentLine -ne "") { $wrappedValues += $currentLine }
            } else {
                $wrappedValues = @($value)
            }
            
            for ($lineIndex = 0; $lineIndex -lt $wrappedValues.Count; $lineIndex++) {
                $lineText = if ($lineIndex -eq 0) {
                    "$($field.Prompt): ".PadRight($maxPromptLength + 2) + $wrappedValues[$lineIndex]
                } else {
                    "".PadRight($maxPromptLength + 2) + $wrappedValues[$lineIndex]
                }
                
                if ($i -eq $currentField) {
                    Write-Host "   $lineText" -ForegroundColor Black -BackgroundColor White
                } else {
                    Write-Host "   $lineText"
                }
            }
            
        }
        
        Write-Host ""
        Write-Host "[↑/↓]  Navigate  |  [Enter]  Edit" -ForegroundColor Blue
        Write-Host "[ S ]  Save      |  [ESC/X]  Cancel" -ForegroundColor Blue

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        switch ($key.VirtualKeyCode) {
            27 { return $null } # Escape
            88 { return $null } # 'x' key
            13 {
                # Enter
                $fieldToEdit = $fields[$currentField]
                $newValue = Read-Input -Prompt "Enter new value for $($fieldToEdit.Prompt) (Enter to keep, 'e' to leave empty)`n" -DefaultValue $preset.($fieldToEdit.Name) -HideDefaultValue
                
                if ($null -eq $newValue) { continue }
                elseif ($newValue -eq 'e') { $preset.($fieldToEdit.Name) = "" }
                else { $preset.($fieldToEdit.Name) = $newValue }
            }
            38 { if ($currentField -gt 0) { $currentField-- } } # Up
            40 { if ($currentField -lt ($fields.Count - 1)) { $currentField++ } } # Down
            83 {
                # 'S' key
                $newName = $preset.name.Trim()
                if ([string]::IsNullOrWhiteSpace($newName)) {
                    Write-ErrorLog "`nPreset Name cannot be empty." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                $isNameDuplicate = $ExistingPresets | Where-Object { $_.name -ne $originalName -and $_.name -eq $newName }
                if ($isNameDuplicate) {
                    Write-ErrorLog "`nError: A preset with the name '$newName' already exists."
                    Start-Sleep -Seconds 3
                    continue
                }
                $preset.name = $newName
                return $preset
            }
        }
    }
}

function Find-Presets {
    param(
        [Parameter(Mandatory = $false)]
        [string]$SearchTerm,

        [Parameter(Mandatory = $true)]
        [array]$AllPresets
    )

    if ([string]::IsNullOrWhiteSpace($SearchTerm)) {
        return $AllPresets | Where-Object { -not (Find-IsCategory -Preset $_) }
    }

    $searchTermLower = $SearchTerm.ToLower()
    Write-DebugLog "Searching for: '$searchTermLower'"
    
    $results = @()
    
    foreach ($preset in $AllPresets) {
        if (Find-IsCategory -Preset $preset) { continue }
        
        $nameLower = if ($preset.name) { $preset.name.ToLower() } else { "" }
        $descLower = if ($preset.description) { $preset.description.ToLower() } else { "" }
        $tagsLower = if ($preset.tags) { $preset.tags.ToLower() } else { "" }
        
        $nameMatch = $nameLower -like "*$searchTermLower*"
        $descMatch = $descLower -like "*$searchTermLower*"
        $tagsMatch = $tagsLower -like "*$searchTermLower*"
        
        if (-not ($nameMatch -or $descMatch -or $tagsMatch)) {
            $otherContent = @(
                $preset.otherOptions, $preset.videoCodec, $preset.audioCodec,
                $preset.resolution, $preset.videoBitrate, $preset.audioBitrate
            ) | Where-Object { $_ } | ForEach-Object { $_.ToLower() }
            $otherContentString = $otherContent -join " "
            $otherMatch = $otherContentString -like "*$searchTermLower*"
        }
        else {
            $otherMatch = $false
        }
        
        if ($nameMatch -or $descMatch -or $tagsMatch -or $otherMatch) {
            $score = 0
            if ($nameMatch) { $score += 100 }
            if ($descMatch) { $score += 50 }
            if ($tagsMatch) { $score += 40 }
            if ($otherMatch) { $score += 10 }
            if ($preset.favorite) { $score += 5 }
            
            Write-DebugLog "  Preset '$nameLower' matched with score: $score"
            $results += [pscustomobject]@{ Preset = $preset; Score = $score }
        }
    }
    
    if ($results.Count -gt 0) {
        Write-DebugLog "Found $($results.Count) matching presets"
        return $results | Sort-Object -Property Score -Descending | Select-Object -ExpandProperty Preset
    }
    else {
        Write-DebugLog "No presets found matching '$searchTermLower'"
        return @()
    }
}

function Invoke-PresetSearch {
    param (
        $config,
        [switch]$ReturnSelection
    )

    $searchQuery = ""
    $selectedIndex = 0
    $results = @()
    $enterAction = if ($ReturnSelection) { "Select" } else { "Edit" }

    while ($true) {
        $results = Find-Presets -SearchTerm $searchQuery -AllPresets $config.presets

        if (-not $DisableClearHost) { Clear-Host }
        if ($ReturnSelection) {
            Write-Host "Search and Select a Preset" -ForegroundColor Cyan
        }
        else {
            Write-Host "Search Presets" -ForegroundColor Cyan
        }
        Write-Host "Type to search`n" -ForegroundColor Gray
        Write-Host "Search: $searchQuery" -ForegroundColor Yellow
        Write-Host "Found: $($results.Count) results" -ForegroundColor Green
        Write-Host "--------------------"

        if ($results.Count -gt 0) {
            $maxNameLength = (@($results.name) | Measure-Object -Maximum -Property Length).Maximum
            if ($null -eq $maxNameLength) { $maxNameLength = 0 }

            $displayOptions = @()
            foreach ($preset in $results) {
                $description = if ($preset.description) { $preset.description } else { "" }
                if ($description.Length -gt 80) {
                    $description = $description.Substring(0, 80) + "..."
                }
                $FavoriteStar = if ($preset.favorite) { '★ ' } else { '  ' }
                $paddedName = $preset.name.PadRight($maxNameLength)
                $displayOptions += "$FavoriteStar$paddedName - $description"
            }

            for ($i = 0; $i -lt $displayOptions.Count; $i++) {
                if ($i -eq $selectedIndex) {
                    Write-Host " > $($displayOptions[$i])" -ForegroundColor Black -BackgroundColor White
                }
                else {
                    Write-Host "   $($displayOptions[$i])"
                }
            }
        }
        else {
            Write-Host "   No presets found matching '$searchQuery'" -ForegroundColor Gray
        }

        Write-Host "`n[ ↑/↓ ] Navigate" -ForegroundColor Blue
        Write-Host "[Enter] $enterAction" -ForegroundColor Blue
        Write-Host "[ESC] Back" -ForegroundColor Blue

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        switch ($key.VirtualKeyCode) {
            27 { return $null } # Escape
            13 { # Enter
                if ($results.Count -gt 0) {
                    $selectedPreset = $results[$selectedIndex]
                    if ($ReturnSelection) {
                        return $selectedPreset
                    }
                    else {
                        $originalIndex = [array]::FindIndex($config.presets, [Predicate[object]] { param($p) $p.name -eq $selectedPreset.name })
                        if ($originalIndex -ge 0) {
                            $editedPreset = Show-PresetEditor -Preset $selectedPreset.psobject.Copy() -ExistingPresets $config.presets
                            if ($editedPreset) {
                                $config.presets[$originalIndex] = [pscustomobject]$editedPreset
                                Save-Config $config
                                return
                            }
                        }
                    }
                }
            }
            8 { # Backspace
                if ($searchQuery.Length -gt 0) {
                    $searchQuery = $searchQuery.Substring(0, $searchQuery.Length - 1)
                    $selectedIndex = 0
                }
            }
            38 { if ($selectedIndex -gt 0) { $selectedIndex-- } } # Up Arrow
            40 { if ($selectedIndex -lt ($results.Count - 1)) { $selectedIndex++ } } # Down Arrow
            default {
                if ($key.Character -ne "`0") {
                    $searchQuery += $key.Character
                    $selectedIndex = 0
                }
            }
        }
    }
}

function Invoke-PresetManager {
    param ($config)
    $selectedIndex = 0
    Write-DebugLog "Opening preset manager"
    
    while ($true) {
        $presetOptions = Get-FormattedPresetList -Config $config
        $menuOptions   = @("Add New Preset or Category", "Search Presets...") + $presetOptions + @("Back")
        
        $categoryIndices = for ($i = 0; $i -lt $config.presets.Count; $i++) {
            if (Find-IsCategory -Preset $config.presets[$i]) {
                $i + 2
            }
        }
        
        $footer = @(
            "[  ↑/↓   ] Navigate      | [Enter] Edit",
            "[ PageUp ] Move Up       | [ DEL ] Delete",
            "[PageDown] Move Down     | [  D  ] Duplicate",
            "[   Q    ] Quick Launch  | [  F  ] Favorite",
            "[ ESC/X  ] Back"
        )

        $menuResult    = Show-Menu -Title "Preset Manager" -Options $menuOptions -SelectedIndex $selectedIndex -CategoryIndices $categoryIndices -Footer $footer -AdditionalReturnKeyCodes @(46, 33, 34, 68, 70, 81)
        $selectedIndex = $menuResult.Index
        $key           = $menuResult.KeyInfo

        switch ($key.VirtualKeyCode) {
            27 { 
                Write-DebugLog "User exited preset manager via ESC"
                return $config 
            }
            88 { 
                Write-DebugLog "User exited preset manager via X"
                return $config 
            }
            13 {
                if ($selectedIndex -eq 0) {
                    Write-DebugLog "User selected 'Add New Preset or Category'"
                    $newPreset = Show-PresetEditor -ExistingPresets $config.presets
                    if ($newPreset) {
                        Write-InfoLog "Added new preset: $($newPreset.name)"
                        $tempPresets = [System.Collections.Generic.List[object]]::new($config.presets)
                        $tempPresets.Add([pscustomobject]$newPreset)
                        $config.presets = $tempPresets.ToArray()
                        Save-Config $config
                    }
                }
                elseif ($selectedIndex -eq 1) { # Search Presets
                    Write-DebugLog "User selected 'Search Presets...'"  
                    Invoke-PresetSearch -config $config
                }
                elseif ($selectedIndex -eq ($menuOptions.Count - 1)) { # Back
                    Write-DebugLog "User selected 'Back'"
                    return $config
                }
                else { # Edit Preset
                    $presetIndex = $selectedIndex - 2
                    $selectedPreset = $config.presets[$presetIndex]
                    Write-DebugLog "User is editing $($selectedPreset.name)"
                    $editedPreset = Show-PresetEditor -Preset $selectedPreset.psobject.Copy() -ExistingPresets $config.presets
                    if ($editedPreset) {
                        Write-InfoLog "Edited preset: $($editedPreset.name)"
                        $config.presets[$presetIndex] = [pscustomobject]$editedPreset
                        Save-Config $config
                    }
                }
            }
            46 { # Delete
                if ($selectedIndex -gt 1 -and $selectedIndex -lt ($menuOptions.Count - 1)) {
                    $presetIndex = $selectedIndex - 2
                    $selectedPreset = $config.presets[$presetIndex]
                    Write-DebugLog "User is deleting $($selectedPreset.name)"
                    $confirm = Read-Input -Prompt "Are you sure you want to remove '$($selectedPreset.name)'? (y/n)" -DefaultValue "n" -HideDefaultValue
                    if ($confirm -eq 'y') {
                        Write-InfoLog "Removed preset: $($selectedPreset.name)"
                        $config.presets = $config.presets | Where-Object { $_.name -ne $selectedPreset.name }
                        if ($config.lastUsedPreset -eq $selectedPreset.name) { $config.lastUsedPreset = "" }
                        if ($selectedIndex -ge ($menuOptions.Count - 2)) { $selectedIndex = $menuOptions.Count - 3 }
                        Save-Config $config
                    }
                }
            }
            33 {  # PageUp
                if ($selectedIndex -gt 2) {
                    $currentIndex = $selectedIndex - 2
                    $presetToMove = $config.presets[$currentIndex]
                    $newIndex = $currentIndex - 1
                    Write-DebugLog "User is moving preset '$($presetToMove.name)' from position $($currentIndex + 1) to position $($newIndex + 1)"
                    $tempPresets  = [System.Collections.Generic.List[object]]::new($config.presets)
                    $tempPresets.RemoveAt($selectedIndex - 2)
                    $tempPresets.Insert($newIndex, $presetToMove)
                    $config.presets = $tempPresets.ToArray()
                    Save-Config $config
                    $selectedIndex--
                }
            }
            34 {  # PageDown
                if ($selectedIndex -gt 1 -and $selectedIndex -lt ($menuOptions.Count - 2)) {
                    $currentIndex = $selectedIndex - 2
                    $presetToMove = $config.presets[$currentIndex]
                    $newIndex = $currentIndex + 1
                    Write-DebugLog "User is moving preset '$($presetToMove.name)' from position $($currentIndex + 1) to position $($newIndex + 1)"
                    $tempPresets  = [System.Collections.Generic.List[object]]::new($config.presets)
                    $tempPresets.RemoveAt($selectedIndex - 2)
                    $tempPresets.Insert($newIndex, $presetToMove)
                    $config.presets = $tempPresets.ToArray()
                    Save-Config $config
                    $selectedIndex++
                }
            }
            68 { # 'D' key - Duplicate
                if ($selectedIndex -gt 1 -and $selectedIndex -lt ($menuOptions.Count - 1)) {
                    $presetIndex    = $selectedIndex - 2
                    $selectedPreset = $config.presets[$presetIndex]
                    Write-DebugLog "User is duplicating $($selectedPreset.name)"
                    $newPreset      = $selectedPreset.psobject.Copy()
                    $newPreset.name = "$($newPreset.name) (copy)"
                    
                    $editedPreset = Show-PresetEditor -Preset $newPreset -ExistingPresets $config.presets
                    if ($editedPreset) {
                        $tempPresets = [System.Collections.Generic.List[object]]::new($config.presets)
                        $tempPresets.Insert($presetIndex + 1, [pscustomobject]$editedPreset)
                        $config.presets = $tempPresets.ToArray()
                        Save-Config $config
                        $selectedIndex = $presetIndex + 3
                    }
                }
            }
            70 { # 'F' key - Favorite
                if ($selectedIndex -gt 1 -and $selectedIndex -lt ($menuOptions.Count - 1)) {
                    $presetIndex = $selectedIndex - 2
                    $current     = $config.presets[$presetIndex]
                    Write-DebugLog "User is toggling favorite for $($current.name)"

                    if ($current.PSObject.Properties.Name -contains 'favorite' -and $current.favorite -eq $true) {
                        Write-InfoLog "Unfavorited preset: $($current.name)"
                        $current.psobject.Properties.Remove('favorite')
                    }
                    else {
                        Write-InfoLog "Favorited preset: $($current.name)"
                        if ($current.PSObject.Properties.Name -contains 'favorite') {
                            $current.favorite = $true
                        }
                        else {
                            $current | Add-Member -NotePropertyName favorite -NotePropertyValue $true -PassThru
                        }
                    }
                    Save-Config $config
                }
            }
            81 { # 'Q' key - Quick Launch
                if ($selectedIndex -gt 1 -and $selectedIndex -lt ($menuOptions.Count - 1)) {
                    $presetIndex    = $selectedIndex - 2
                    $selectedPreset = $config.presets[$presetIndex]
                    Write-DebugLog "User is setting $($selectedPreset.name) as the Quick Launch preset."
                    if (Find-IsCategory -Preset $selectedPreset) {
                        Write-Host "Categories cannot be set as the Quick Launch preset." -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        continue
                    }
                    if ($config.quickLaunchPreset -eq $selectedPreset.name) {
                        $config.quickLaunchPreset = ""
                    }
                    else { $config.quickLaunchPreset = $selectedPreset.name }
                    Save-Config $config
                }
            }
        }
    }
}
#endregion

#region Miscellaneous functions
function Find-Executables {
    Write-DebugLog "Attempting to find scrcpy and adb..."
    try {
        $scrcpyPath = (Get-Command scrcpy -ErrorAction SilentlyContinue).Path
        $adbPath = (Get-Command adb -ErrorAction SilentlyContinue).Path
    }
    catch {
        Write-ErrorLog "Failed to query commands from PATH." $_.Exception
        return [pscustomobject]@{ Success = $false }
    }

    if ($scrcpyPath -and $adbPath) {
        Write-DebugLog "Found scrcpy: $scrcpyPath"
        Write-DebugLog "Found adb: $adbPath"
        return [pscustomobject]@{ Success = $true; ScrcpyPath = $scrcpyPath; AdbPath = $adbPath }
    }

    Write-ErrorLog "scrcpy or adb not found in your system's PATH."
    return [pscustomobject]@{ Success = $false }
}

function Invoke-Remux {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $mp4Path = $SourcePath -replace '\.mkv$', '.mp4'
    Write-InfoLog "`nStarting remuxing process..." -ForegroundColor Yellow

    try {
        $ffmpegPath = (Get-Command ffmpeg -ErrorAction SilentlyContinue).Path
        if (-not $ffmpegPath) {
            throw "FFmpeg not found. Cannot remux. Please install FFmpeg and add it to your system's PATH."
        }
        
        Write-InfoLog "Using FFmpeg's native AAC encoder for audio (VBR Quality)." -ForegroundColor Yellow
        $audioCodec = "aac"
        $audioQuality = "1.8" # VBR quality setting, adjust as needed
        
        Write-Host "Muxing final MP4 file..."
        Invoke-SafeCommand -Command { & $ffmpegPath -i $SourcePath -map 0:v:0 -map 0:a:0 -c:v copy -c:a $audioCodec -q:a $audioQuality -movflags +faststart $mp4Path -y -hide_banner -loglevel error } -ErrorMessage "FFmpeg remuxing failed" -ContinueOnError

        if (-not (Test-Path $mp4Path)) {
            throw "FFmpeg failed to mux the final video file."
        }
        Write-InfoLog "Remuxing complete. Recording saved to: $mp4Path" -ForegroundColor Cyan
        
        if (Test-Path $mp4Path) {
            Remove-Item $SourcePath -Force
        }
    }
    catch {
        Write-ErrorLog "An error occurred during remuxing: $($_.Exception.Message)"
        Write-ErrorLog "The original MKV file has been preserved at: $SourcePath"
    }
}
#endregion

#region scrcpy Session Functions
function Start-Scrcpy {
    param (
        $executables,
        $config,
        [switch]$IsRecording,
        [string]$InitialPresetName = $null,
        [string]$DeviceSerial = $null
    )
    
    $relaunch          = $false
    $currentPresetName = $InitialPresetName

    if (-not [string]::IsNullOrEmpty($DeviceSerial)) {
        $config.selectedDevice = $DeviceSerial
        Write-InfoLog "Using provided device serial: $DeviceSerial"
    }

    Write-InfoLog "Starting scrcpy session (Recording: $IsRecording, InitialPreset: $(
    if ($InitialPresetName) {
        $InitialPresetName
    } else {'None'}))"

    do {
        $relaunch = $false
        Write-DebugLog "Starting scrcpy session loop"
        # 1. Check if device is selected
        if ([string]::IsNullOrWhiteSpace($config.selectedDevice)) {
            Write-DebugLog "No device selected, showing device selection immediately"
            $selectedDevice = Show-DeviceSelection -adbPath $executables.AdbPath
            if ($null -eq $selectedDevice) { 
                $config.selectedDevice = ""
                Save-Config $config
                return 
            }
            $config.selectedDevice = $selectedDevice
            Save-Config $config
            Write-InfoLog "Selected device: $selectedDevice"
        }

        # 2. Validate selected device (only if we have one)
        if (-not [string]::IsNullOrWhiteSpace($config.selectedDevice)) {
            $config.selectedDevice = $config.selectedDevice.Trim()
            Write-DebugLog "Validating device connection: $($config.selectedDevice)"
            $deviceList = Get-AdbDeviceList -adbPath $executables.AdbPath
            $device = $deviceList | Where-Object { $_.Serial -eq $config.selectedDevice } | Select-Object -First 1
            
            if (-not $device -or $device.State -ne 'device') {
                Write-ErrorLog "Device validation failed. Device found: $($null -ne $device), State: $($device.State)"
                Write-DebugLog "Device not ready, attempting to refresh..."
                $maxRetries = 5
                $retryCount = 0
                $deviceReady = $false

                while ($retryCount -lt $maxRetries -and -not $deviceReady) {
                    $retryCount++
                    Write-DebugLog "Waiting for device to become ready (attempt $retryCount/$maxRetries)..."
                    Start-Sleep -Seconds 1

                    $deviceList = Get-AdbDeviceList -adbPath $executables.AdbPath
                    $device = $deviceList | Where-Object { $_.Serial -eq $config.selectedDevice } | Select-Object -First 1

                    if ($device -and $device.State -eq 'device') {
                        $deviceReady = $true
                        Write-DebugLog "Device is now ready!"
                    }
                }
                
                if (-not $deviceReady) {
                    if (-not [string]::IsNullOrEmpty($DeviceSerial)) {
                        Write-ErrorLog "Provided device '$($config.selectedDevice)' is not connected or not in device state."
                        return
                    } else {
                        Write-ErrorLog "Device '$($config.selectedDevice)' is not connected or not in device state. Please select a new one."
                        $config.selectedDevice = ""
                        Save-Config $config
                        Start-Sleep -Seconds 2
                        continue
                    }
                }
            }
        }
        # 3. Select Device if needed
        if ([string]::IsNullOrEmpty($config.selectedDevice)) {
            $selectedDevice = Show-DeviceSelection -adbPath $executables.AdbPath
            if ($null -eq $selectedDevice) { 
                $config.selectedDevice = ""
                Save-Config $config
                return 
            }
            $config.selectedDevice = $selectedDevice
            Save-Config $config
            Write-InfoLog "Selected device: $selectedDevice"
        }
        else {
            $deviceList = Get-AdbDeviceList -adbPath $executables.AdbPath
            $device = $deviceList | Where-Object { $_.Serial -eq $config.selectedDevice -and $_.State -eq 'device' }
            
            if (-not $device) {
                Write-ErrorLog "Stored device '$($config.selectedDevice)' is not available. Please select a new one."
                $config.selectedDevice = ""
                Start-Sleep -Seconds 2
                continue
            }
        }
        # 4. Select Preset
        $selectedPreset = $null
        if ($currentPresetName) {
            Write-DebugLog "Looking for preset: $currentPresetName"
            $selectedPreset = $config.presets | Where-Object { $_.name -eq $currentPresetName } | Select-Object -First 1
            if (-not $selectedPreset) {
                Write-ErrorLog "Preset '$currentPresetName' was not found. Please select one manually."
                Start-Sleep -Seconds 2
                $currentPresetName = $null
            } else {
                Write-DebugLog "Found preset: $($selectedPreset.name)"
            }
        }

        if (-not $selectedPreset) {
            Write-DebugLog "No preset specified, showing selection menu"
            $presetOptions   = Get-FormattedPresetList -Config $config
            $categoryIndices = for ($i = 0; $i -lt $config.presets.Count; $i++) { if (Find-IsCategory -Preset $config.presets[$i]) { $i + 1 } }
            
            while ($true) {
                $menuOptions  = @("Search Presets...") + $presetOptions + @("Back")
                $menuResult   = Show-Menu -Title "Select a Preset to Launch" -Options $menuOptions -CategoryIndices $categoryIndices -SkipCategoriesOnNavigate -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Back")
                $presetChoice = $menuResult.Index

                if ($presetChoice -eq -1 -or $presetChoice -eq ($menuOptions.Count - 1)) { 
                    Write-InfoLog "User canceled preset selection"
                    return 
                }
                
                if ($presetChoice -eq 0) {
                    Write-DebugLog "User selected search option"
                    $selectedPreset = Invoke-PresetSearch -config $config -ReturnSelection
                    if ($null -ne $selectedPreset) {
                        Write-DebugLog "User selected preset from search: $($selectedPreset.name)"
                        break
                    }
                }
                elseif ($categoryIndices -contains $presetChoice) {
                    Write-DebugLog "User selected a category, prompting for valid preset"
                    Write-Host "`nCategories cannot be launched. Please select a valid preset." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
                else {
                    $selectedPreset = $config.presets[$presetChoice - 1]
                    Write-DebugLog "User selected preset: $($selectedPreset.name)"
                    break
                }
            }
        }
        
        $currentPresetName     = $selectedPreset.name
        $config.lastUsedPreset = $currentPresetName
        Save-Config $config
        Write-InfoLog "Using preset: $currentPresetName"

        # 5. Build Command Arguments
        try {
            $finalArgs = Build-ScrcpyArguments -SelectedPreset $selectedPreset -SelectedDevice $config.selectedDevice
        }
        catch {
            Write-ErrorLog "Failed to build scrcpy arguments: $($_.Exception.Message)"
            return
        }
        
        $fullPath = $null

        if ($IsRecording) {
            $recordingPath = $config.recordingPath
            if (-not (Test-Path $recordingPath)) {
                Write-ErrorLog "Recording path '$recordingPath' does not exist. Please set a valid path."
                Wait-Enter
                return
            }
            $fileExt         = "mkv"
            $defaultFilename = "$($selectedPreset.name)_$(Get-Date -Format 'yyyyMMdd_HHmmss').$fileExt"
            $filename        = Read-Input -Prompt "Enter recording filename" -DefaultValue $defaultFilename -HideDefaultValue
            if ($null -eq $filename) { 
                Write-InfoLog "User canceled filename input"
                return 
            }
            $fullPath   = Join-Path -Path $recordingPath -ChildPath $filename
            $finalArgs += "--record", "`"$fullPath`""
            Write-InfoLog "Recording enabled, output path: $fullPath"
        }
        # 6. Launch scrcpy
        if (-not $DisableClearHost) { Clear-Host }
        $deviceDisplayName = Get-DeviceDisplayName -adbPath $executables.AdbPath -deviceSerial $config.selectedDevice
        Write-Host "  STARTING SCRCPY SESSION" -ForegroundColor Cyan
        Write-InfoLog "Preset: '$($selectedPreset.name)'"
        Write-InfoLog "Target Device: $deviceDisplayName"
        if ($IsRecording) {
            Write-InfoLog "Recording to: $fullPath"
            Write-Host "`nExit scrcpy to stop recording" -ForegroundColor Yellow
        }
        Write-InfoLog "Command: scrcpy $($finalArgs -join ' ')"

        try {
            Write-DebugLog "Launching scrcpy process"
                
            if ($DebugPreference -ne "SilentlyContinue" -or $RealTimeCapture) {
                Write-DebugLog "Real-time output capture enabled"
                
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo
                $processInfo.FileName = $executables.ScrcpyPath
                $processInfo.Arguments = $finalArgs
                $processInfo.RedirectStandardOutput = $true
                $processInfo.RedirectStandardError = $true
                $processInfo.UseShellExecute = $false
                $processInfo.CreateNoWindow = $true
                
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo = $processInfo
                $process.Start() | Out-Null
                
                $combinedOutput = [System.Collections.ArrayList]::new()
                
                while (-not $process.HasExited) {
                    Start-Sleep -Milliseconds 200
                    
                    while ($process.StandardOutput.Peek() -gt -1) {
                        $outputLine = $process.StandardOutput.ReadLine()
                        if ($null -ne $outputLine) {
                            Write-Host $outputLine
                            [void]$combinedOutput.Add($outputLine)
                        }
                    }
                    
                    while ($process.StandardError.Peek() -gt -1) {
                        $errorLine = $process.StandardError.ReadLine()
                        if ($null -ne $errorLine) {
                            Write-Host $errorLine -ForegroundColor Red
                            [void]$combinedOutput.Add($errorLine)
                        }
                    }
                }
                
                $remainingOutput = $process.StandardOutput.ReadToEnd()
                if (-not [string]::IsNullOrEmpty($remainingOutput)) {
                    Write-Host $remainingOutput
                    [void]$combinedOutput.Add($remainingOutput)
                }
                
                $remainingError = $process.StandardError.ReadToEnd()
                if (-not [string]::IsNullOrEmpty($remainingError)) {
                    Write-Host $remainingError -ForegroundColor Red
                    [void]$combinedOutput.Add($remainingError)
                }
                
                if ($combinedOutput.Count -gt 0) {
                    Write-LogOnly "Scrcpy output:`n$($combinedOutput -join "`n")"
                }
            }
            else {
                Write-DebugLog "Normal mode - letting scrcpy output directly to console"
                $process = Start-Process -FilePath $executables.ScrcpyPath -ArgumentList $finalArgs -NoNewWindow -PassThru -Wait
            }
            
            Write-DebugLog "Scrcpy exit code: $($process.ExitCode)"
        }
        catch {
            Write-ErrorLog "Failed to launch scrcpy." $_.Exception
        }
        
        # 7. Post-Session Handling
        switch ($process.ExitCode) {
            0 {
                Write-InfoLog "scrcpy session ended."
            }
            1 {
                Write-ErrorLog "Start failure. Review scrcpy output for details."
            }
            2 {
                Write-ErrorLog "Device disconnected while running."
            }
            default {
                Write-ErrorLog "scrcpy exited with an unexpected code: $($process.ExitCode)."
            }
        }
        
        if ($IsRecording -and (Test-Path $fullPath)) {
            if ($config.recordingFormat -eq "RemuxToMP4") {
                Write-InfoLog "Starting remuxing process for: $fullPath"
                Invoke-Remux -SourcePath $fullPath
            }
            else {
                Write-InfoLog "Recording saved to: $fullPath"
            }
        }
        
        Write-Host ""
        Write-Host "[Enter] Return to Main Menu" -ForegroundColor Yellow
        Write-Host "[  R  ] Re-launch with same preset" -ForegroundColor Yellow
        Write-Host "[ESC/X] Exit Script" -ForegroundColor Yellow
        
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Write-DebugLog "Key pressed: $($key.VirtualKeyCode)"
        switch ($key.VirtualKeyCode) {
            13 { 
                Write-DebugLog "User chose to return to main menu"
                $relaunch = $false 
            }
            82 { 
                Write-DebugLog "User chose to relaunch with same preset"
                $relaunch = $true 
            }
            88 { 
                Write-InfoLog "User exited script via X key"
                Exit 
            }
            27 { 
                Write-InfoLog "User exited script via ESC key"
                Exit 
            }
            default {
                Write-DebugLog "User pressed invalid key: $($key.VirtualKeyCode)"
                Write-Host "Invalid key." -ForegroundColor Red
                Start-Sleep -Seconds 1
                $relaunch = $false
            }
        }

    } while ($relaunch)
}
#endregion

#region Main Function
function Main {
    param([string]$Preset)
    
    Write-DebugLog "initializing scrcpy-Automation v$ScriptVersion with parameters:"
    if (-not [string]::IsNullOrEmpty($DeviceSerial)) {
        $config.selectedDevice = $DeviceSerial.Trim()
        Write-DebugLog "-DeviceSerial: $DeviceSerial" 
    }
    if (-not [string]::IsNullOrEmpty($Preset)) { Write-DebugLog "-Preset: $Preset" }
    if ($Log) { Write-DebugLog "-Log: $Log" }
    if ($NoClear) { Write-DebugLog "-NoClear: $NoClear" }
    if ($RealTimeCapture) { Write-DebugLog "-RealTimeCapture: $RealTimeCapture" }
    if (-not [string]::IsNullOrEmpty($ConfigPath) -and $ConfigPath -ne (Join-Path $PSScriptRoot "scrcpy-config.json")) { 
        Write-DebugLog "-ConfigPath: $ConfigPath" 
    }
    if (-not [string]::IsNullOrEmpty($LogPath) -and $LogPath -ne (Join-Path $PSScriptRoot "scrcpy-automation.log")) { 
        Write-DebugLog "-LogPath: $LogPath" 
    }
    $executables = Find-Executables
    if (-not $executables.Success) { Read-Host "Press Enter to exit..."; return }
    
    $config = Get-Config
    if ($null -eq $config) { Read-Host "Press Enter to exit..."; return }

    if (-not [string]::IsNullOrEmpty($Preset) -and $Preset.StartsWith('-')) {
        Write-InfoLog "The argument '$Preset' appears to be a command-line switch" -ForegroundColor Yellow
        Write-Host "For script help, use the standard PowerShell command: Get-Help `"$PSCommandPath`""
        Read-Host "Press Enter to exit..."
        return
    }

    if (-not [string]::IsNullOrEmpty($Preset)) {
        Write-InfoLog "Searching for preset: '$Preset'"
    
        $presetMatches = Find-Presets -SearchTerm $Preset -AllPresets $config.presets
        $targetPreset = $null

        if ($presetMatches.Count -gt 0) {
            $bestMatch = $presetMatches[0]
            $confirm = Read-Input -Prompt "Did you mean '$($bestMatch.name)'? (y/n)" -DefaultValue "y" -HideDefaultValue
            if ($confirm -eq 'y') {
                $targetPreset = $bestMatch
            }
            else {
                Write-InfoLog "Launch canceled."
                return
            }
        }

        if ($targetPreset) {
            Write-InfoLog "Using preset: '$($targetPreset.name)'"
            Start-Scrcpy -executables $executables -config $config -InitialPresetName $targetPreset.name -DeviceSerial $DeviceSerial
            Write-InfoLog "Exiting script after direct launch."
        }
        else {
            Write-ErrorLog "Preset '$Preset' not found, and no close matches were detected."
            Read-Host "Press Enter to exit..."
        }
        return
    }

    $selectedIndex = 0
    while ($true) {
        $options = @()
        $deviceDisplayName = if (-not [string]::IsNullOrEmpty($config.selectedDevice)) {
            Get-DeviceDisplayName -adbPath $executables.AdbPath -deviceSerial $config.selectedDevice
        } else {
            "No device selected"
        }
        $options += "Device: $deviceDisplayName"
        if (-not [string]::IsNullOrEmpty($config.quickLaunchPreset)) { $options += "Quick Launch: $($config.quickLaunchPreset)" }
        if (-not [string]::IsNullOrEmpty($config.lastUsedPreset)) {
            $options += "Last: $($config.lastUsedPreset)"
            $options += "Record Last: $($config.lastUsedPreset)"
        }
        $options += "Start scrcpy", "Record scrcpy", "Manage Presets", "Recording Options", "Exit"
        
        $menuResult = Show-Menu -Title "scrcpy Automation v$ScriptVersion" -Options $options -SelectedIndex $selectedIndex -Footer @("[ ↑/↓ ] Navigate", "[Enter] Select", "[ESC/X] Exit")
        if ($menuResult.Key -in @('Escape', 'x')) { Write-Host "Exiting..."; return }
        $selectedIndex = $menuResult.Index
        $chosenOption = $options[$selectedIndex]
        Write-DebugLog "User selected main menu option: $chosenOption"

        if ($chosenOption.StartsWith("Device")) {
            $selectedDevice = Show-DeviceSelection -adbPath $executables.AdbPath -currentDevice $config.selectedDevice
            if ($null -ne $selectedDevice) {
                $config.selectedDevice = $selectedDevice
                Save-Config $config
            } else {
                $config.selectedDevice = ""
                Save-Config $config
            }
        }
        elseif ($chosenOption.StartsWith("Quick Launch")) { Start-Scrcpy -executables $executables -config $config -InitialPresetName $config.quickLaunchPreset }
        elseif ($chosenOption.StartsWith("Last")) { Start-Scrcpy -executables $executables -config $config -InitialPresetName $config.lastUsedPreset }
        elseif ($chosenOption.StartsWith("Record Last")) { Start-Scrcpy -executables $executables -config $config -IsRecording -InitialPresetName $config.lastUsedPreset }
        elseif ($chosenOption -eq "Start scrcpy") { Start-Scrcpy -executables $executables -config $config }
        elseif ($chosenOption -eq "Record scrcpy") { Start-Scrcpy -executables $executables -config $config -IsRecording }
        elseif ($chosenOption -eq "Manage Presets") { $config = Invoke-PresetManager -config $config }
        elseif ($chosenOption -eq "Recording Options") { $config = Show-RecordingOptions -config $config }
        elseif ($chosenOption -eq "Exit") { Write-Host "Exiting..."; return }
    }
}
#endregion

Main -Preset $Preset