# ===============================
# TabletDesk
# ===============================
# Waits for a Spacedesk monitor to connect, then launches a weather page in kiosk mode on Edge.
# Hotkey and GUI support, robust error handling, and logging.

using namespace System.Management.Automation
using namespace System.Management.Automation.Runspaces

trap {
    Write-Log ("[UNHANDLED ERROR] $_")
    continue
}

# -------------------------------
# Logging & Safe File Operations
# -------------------------------
function Write-Log {
    param([string]$message)
    $logPath = $global:logPath
    if (![string]::IsNullOrWhiteSpace($logPath)) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Safe-AddContent -Path $logPath -Value "[$timestamp] $message"
    }
}
function Safe-AddContent {
    param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Value)
    if (![string]::IsNullOrWhiteSpace($Path)) { Add-Content -Path $Path -Value $Value }
}
function Safe-SetContent {
    param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Value)
    if (![string]::IsNullOrWhiteSpace($Path)) { Set-Content -Path $Path -Value $Value }
}
function Safe-GetContent {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (![string]::IsNullOrWhiteSpace($Path) -and (Test-Path $Path)) { return Get-Content -Path $Path -Raw }
    return $null
}

$ErrorActionPreference = 'Stop'

try {
    Write-Log ("[INIT] Script starting, attempting to load System.Windows.Forms...")
    Add-Type -AssemblyName System.Windows.Forms
    Write-Log ("[INIT] System.Windows.Forms loaded successfully.")
} catch {
    Write-Log ("[INIT][ERROR] Failed to load System.Windows.Forms: {0}" -f $_)
    throw
}
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing

# -------------------------------
# Configuration
# -------------------------------
$global:configPath = Join-Path $env:USERPROFILE "TabletDeskConfig.json"
$global:logPath = Join-Path $env:USERPROFILE "TabletDeskLog.txt"
$defaultConfig = @{
    Url = "http://192.168.1.5:8891/#/app/weatherwaves"
    DisplayIndex = 0
    AutoStart = $true
    ScanTimeoutSec = 10
    DeskThingSleepMs = 100
    DisplayPollInterval = 0.5
    MaxDisplayWaitSec = 30
    DisplayGracePeriodSec = 0.5
    UseEdge = $true
    ServerCheckRetries = 2
    HotKey = "Ctrl+Alt+W"
    CustomKiosk = $false
    KioskArgs = ""
}
$defaultConfigKeys = @('Url','DisplayIndex','AutoStart','ScanTimeoutSec','DeskThingSleepMs','DisplayPollInterval','MaxDisplayWaitSec','DisplayGracePeriodSec','UseEdge','ServerCheckRetries','HotKey','CustomKiosk','KioskArgs')
function Validate-Config {
    param($config)
    $changed = $false
    foreach ($key in $defaultConfig.Keys) {
        if (-not $config.ContainsKey($key)) {
            $config[$key] = $defaultConfig[$key]
            $changed = $true
            Write-Log ("Config missing property '{0}', set to default." -f $key)
        }
    }
    if ($changed) { Save-Config -config $config }
    return $config
}

# Function to load configuration
function Load-Config {
    if (![string]::IsNullOrWhiteSpace($global:configPath) -and (Test-Path $global:configPath)) {
        try {
            $configObj = Safe-GetContent -Path $global:configPath | ConvertFrom-Json
            if ($configObj -isnot [hashtable]) {
                $config = @{}
                foreach ($prop in $configObj.PSObject.Properties) {
                    $config[$prop.Name] = $prop.Value
                }
            } else {
                $config = $configObj
            }
            $config = Validate-Config $config
            Write-Log ("Configuration loaded successfully")
            return $config
        } catch {
            Write-Log ("Error loading configuration: {0}" -f $_)
            return $defaultConfig.Clone()
        }
    } else {
        Write-Log ("No configuration file found or config path is empty, using defaults")
        return $defaultConfig.Clone()
    }
}

# Function to save configuration
function Save-Config {
    param($config)
    try {
        if (![string]::IsNullOrWhiteSpace($global:configPath)) {
            $json = $config | ConvertTo-Json
            Safe-SetContent -Path $global:configPath -Value $json
            Write-Log ("Configuration saved successfully")
        } else {
            Write-Log ("Config path is empty. Not saving configuration.")
        }
    } catch {
        Write-Log ("Error saving configuration: {0}" -f $_)
    }
}

# More robust waiting for Spacedesk/secondary display
function Wait-ForSpacedeskDisplay {
    param(
        [int]$maxWaitSeconds = 30,
        [double]$retryIntervalSeconds = 0.5,
        [int]$minDisplays = 2
    )
    Write-Log ("Entering Wait-ForSpacedeskDisplay")
    $grace = 0.5
    if ($global:config) {
        if ($global:config.ContainsKey('MaxDisplayWaitSec')) { $maxWaitSeconds = [int]$global:config['MaxDisplayWaitSec'] }
        if ($global:config.ContainsKey('DisplayPollInterval')) { $retryIntervalSeconds = [double]$global:config['DisplayPollInterval'] }
        if ($global:config.ContainsKey('DisplayGracePeriodSec')) { $grace = [double]$global:config['DisplayGracePeriodSec'] }
    }
    if ($grace -gt 0) { Start-Sleep -Seconds $grace }
    $displays = [System.Windows.Forms.Screen]::AllScreens
    $nonPrimary = $displays | Where-Object { -not $_.Primary }
    Write-Log ("Initial display count: {0}" -f $displays.Count)
    if ($nonPrimary) {
        foreach ($d in $nonPrimary) {
            Write-Log ("Non-primary display detected at startup: {0}, Bounds=({1},{2},{3},{4})" -f $d.DeviceName, $d.Bounds.X, $d.Bounds.Y, $d.Bounds.Width, $d.Bounds.Height)
        }
        Write-Log ("Non-primary display already present at startup, skipping wait.")
        return $true
    }
    Write-Log ("Waiting for a non-primary display (Spacedesk) to connect...")
    $startTime = Get-Date
    $endTime = $startTime.AddSeconds($maxWaitSeconds)
    $lastCount = $displays.Count
    while ((Get-Date) -lt $endTime) {
        $displays = [System.Windows.Forms.Screen]::AllScreens
        $nonPrimary = $displays | Where-Object { -not $_.Primary }
        Write-Log ("DEBUG: Detected {0} display(s) at {1}" -f $displays.Count, (Get-Date -Format "HH:mm:ss"))
        if ($nonPrimary) {
            foreach ($d in $nonPrimary) {
                Write-Log ("Non-primary display detected: {0}, Bounds=({1},{2},{3},{4})" -f $d.DeviceName, $d.Bounds.X, $d.Bounds.Y, $d.Bounds.Width, $d.Bounds.Height)
            }
            Write-Log ("Non-primary display connected, proceeding.")
            return $true
        }
        if ($lastCount -ne $displays.Count) {
            Write-Log ("DEBUG: Display count changed from {0} to {1}" -f $lastCount, $displays.Count)
            $lastCount = $displays.Count
        }
        Start-Sleep -Seconds $retryIntervalSeconds
    }
    Write-Log ("No non-primary display detected within {0} seconds" -f $maxWaitSeconds)
    return $false
}

# Function to check if server is running
function Test-ServerConnection {
    param(
        [string]$url,
        [int]$retries = 2
    )
    $baseUrl = $url -replace '#.*', ''
    Write-Log ("Testing connection to {0}" -f $baseUrl)
    for ($i = 0; $i -lt $retries; $i++) {
        try {
            $response = Invoke-WebRequest -Uri $baseUrl -TimeoutSec 2 -UseBasicParsing -ErrorAction SilentlyContinue
            if ($response.StatusCode -eq 200) {
                Write-Log ("Server is up and running")
                return $true
            }
        } catch {
            Write-Log ("Attempt {0}/{1}: Server not responding: {2}" -f ($i+1), $retries, $_)
        }
        Start-Sleep -Milliseconds 400
    }
    Write-Log ("Server not available after {0} attempts" -f $retries)
    return $false
}

# Function to find Spacedesk display
function Find-SpacedeskDisplay {
    $displays = [System.Windows.Forms.Screen]::AllScreens
    if ($displays.Count -gt 1) {
        foreach ($display in $displays) {
            if (-not $display.Primary) {
                return $displays.IndexOf($display)
            }
        }
    }
    if ($displays.Count -gt 1) {
        return 1
    }
    return 0
}

# Enhanced debug logging for display detection and Edge launch
function Launch-EdgeKiosk {
    param(
        [string]$url,
        [int]$displayIndex = 1
    )
    $displays = [System.Windows.Forms.Screen]::AllScreens
    Write-Log ("DEBUG: Detected {0} display(s)." -f $displays.Count)
    $i = 0
    foreach ($display in $displays) {
        Write-Log ("DEBUG: Display {0}: Bounds=({1},{2},{3},{4}), Primary={5}" -f $i, $display.Bounds.X, $display.Bounds.Y, $display.Bounds.Width, $display.Bounds.Height, $display.Primary)
        $i++
    }
    # Automatically select the first non-primary display if available
    $targetDisplay = $null
    foreach ($display in $displays) {
        if (-not $display.Primary) {
            $targetDisplay = $display
            break
        }
    }
    if (-not $targetDisplay) {
        $targetDisplay = $displays[0]
        Write-Log ("No non-primary display found. Using primary display (index 0).")
    } else {
        Write-Log ("Using non-primary display: {0}" -f $displays.IndexOf($targetDisplay))
    }
    $bounds = $targetDisplay.Bounds
    Write-Log ("Launching Edge on display {0} ({1}x{2}) for URL: {3}" -f $displays.IndexOf($targetDisplay), $bounds.Width, $bounds.Height, $url)
    try {
        $edgePath = $null
        try {
            $edgeCmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
            if ($edgeCmd) { $edgePath = $edgeCmd.Source }
            Write-Log ("Edge found in PATH: {0}" -f $edgePath)
        } catch {
            Write-Log ("Error searching PATH for Edge: {0}" -f $_)
        }
        if (-not $edgePath) {
            $possiblePaths = @(
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
            )
            foreach ($path in $possiblePaths) {
                Write-Log ("Checking Edge path: {0}" -f $path)
                if (Test-Path $path) {
                    $edgePath = $path
                    Write-Log ("Edge found at: {0}" -f $path)
                    break
                }
            }
        }
        if (-not $edgePath) {
            Write-Log ("Edge browser not found in PATH or default locations. Please check your Edge installation.")
            return
        }
        $extraArgs = ""
        if ($global:config['CustomKiosk'] -and $global:config['KioskArgs']) {
            $extraArgs = $global:config['KioskArgs']
        }
        $args = "-k $url --kiosk-type=fullscreen $extraArgs"
        Write-Log ("Starting Edge process: {0} {1}" -f $edgePath, $args)
        $proc = Start-Process $edgePath -ArgumentList $args -PassThru -ErrorAction Stop
        if ($proc) {
            Write-Log ("Edge process started. PID: {0}" -f $proc.Id)
        } else {
            Write-Log ("Failed to start Edge process.")
            return
        }
        # Event-driven: Wait for Edge window handle to appear (max 4s)
        $edgeProcess = $null
        $waited = 0
        while ($waited -lt 40) {
            $edgeProcess = Get-Process msedge -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
            if ($edgeProcess) { break }
            Start-Sleep -Milliseconds 100
            $waited++
        }
        if ($edgeProcess) {
            $hwnd = $edgeProcess.MainWindowHandle
            Write-Log ("Moving Edge window. HWND: {0}" -f $hwnd)
            Add-Type @"
            using System;
            using System.Runtime.InteropServices;
            public class WindowMover {
                [DllImport("user32.dll")]
                public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
                [DllImport("user32.dll")]
                public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
                [DllImport("user32.dll")]
                public static extern IntPtr GetForegroundWindow();
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool SetForegroundWindow(IntPtr hWnd);
            }
"@
            [WindowMover]::MoveWindow($hwnd, $bounds.X, $bounds.Y, $bounds.Width, $bounds.Height, $true)
            [WindowMover]::ShowWindow($hwnd, 3) # 3 = Maximize
            [WindowMover]::SetForegroundWindow($hwnd)
            Add-Type -AssemblyName System.Windows.Forms
            Start-Sleep -Milliseconds 500
            Write-Log ("Sending F11 to Edge window for fullscreen.")
            [System.Windows.Forms.SendKeys]::SendWait("{F11}")
            Write-Log ("F11 key sent to Edge window (once).")
            Write-Log ("Edge window moved to display {0} and set to full screen" -f $displays.IndexOf($targetDisplay))
            return
        } else {
            Write-Log ("Could not find Edge window to move after starting process. PID: {0}" -f $proc.Id)
            return
        }
    } catch {
        Write-Log ("Error launching Edge: {0}" -f $_)
        return
    }
}

# Function to find weather server
function Find-WeatherServer {
    $ipBase = "192.168.1."
    $ports = @(8891)
    $serverPath = "/"
    $timeoutSec = 2
    $ipRange = 1..254
    $maxParallel = 32
    $foundUrl = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $config = $global:config
    if ($config -and $config.ContainsKey('ScanTimeoutSec')) {
        $maxScanTime = [int]$config['ScanTimeoutSec']
    } else {
        $maxScanTime = 10
    }
    # 1. Synchronously check localhost and 127.0.0.1
    foreach ($port in $ports) {
        foreach ($hostname in @('localhost','127.0.0.1')) {
            $url = "http://${hostname}:${port}${serverPath}"
            Write-Log ("[SCAN] Sync checking: {0}" -f $url)
            try {
                $resp = Invoke-WebRequest -Uri $url -TimeoutSec $timeoutSec -ErrorAction Stop
                if ($resp.StatusCode -eq 200) {
                    Write-Log ("[SCAN] FOUND SERVER (sync): {0}" -f $url)
                    return $url
                }
            } catch {
                Write-Log ("[SCAN] Sync check failed: {0}" -f $_)
            }
        }
    }
    # 2. Parallel scan for network IPs
    $ipQueue = @()
    foreach ($i in $ipRange) {
        foreach ($port in $ports) {
            $ipQueue += ,@($i, $port)
        }
    }
    Write-Log ("[SCAN] Will check these URLs:")
    foreach ($item in $ipQueue) {
        $url = "http://${ipBase}${item[0]}:${item[1]}${serverPath}"
        Write-Log ("[SCAN] Queue: {0}" -f $url)
    }
    $jobs = @()
    while ($ipQueue.Count -gt 0 -or $jobs.Count -gt 0) {
        while ($jobs.Count -lt $maxParallel -and $ipQueue.Count -gt 0) {
            $next = $ipQueue[0]; $ipQueue = $ipQueue[1..($ipQueue.Count-1)]
            $i = $next[0]; $port = $next[1]
            $url = "http://${ipBase}${i}:${port}${serverPath}"
            Write-Log ("[SCAN] Checking: {0}" -f $url)
            $jobs += Start-Job -ScriptBlock {
                param($url, $timeoutSec)
                try {
                    $resp = Invoke-WebRequest -Uri $url -TimeoutSec $timeoutSec -ErrorAction Stop
                    if ($resp.StatusCode -eq 200) { return $url }
                } catch {}
                return $null
            } -ArgumentList $url, $timeoutSec
        }
        $finished = Wait-Job -Job $jobs -Any -Timeout 1
        if ($finished) {
            foreach ($job in @($finished)) {
                $result = Receive-Job $job
                if ($result) {
                    Write-Log ("[SCAN] Job result: {0}" -f $result)
                    $foundUrl = $result
                    Write-Log ("[SCAN] FOUND SERVER: {0}" -f $foundUrl)
                    $jobs | Remove-Job -Force -ErrorAction SilentlyContinue
                    break
                } else {
                    Write-Log ("[SCAN] Job result: null")
                }
            }
            $jobs = $jobs | Where-Object { $_.State -eq 'Running' }
        }
        if ($foundUrl -or $stopwatch.Elapsed.TotalSeconds -ge $maxScanTime) { break }
    }
    if ($foundUrl) {
        Write-Log ("Weather server found at {0} (strict scan, {1:N1}s)" -f $foundUrl, $stopwatch.Elapsed.TotalSeconds)
        return $foundUrl
    } else {
        if ($stopwatch.Elapsed.TotalSeconds -ge $maxScanTime) {
            Write-Log ("Scan exceeded max time ({0:N1}s), aborting all jobs." -f $stopwatch.Elapsed.TotalSeconds)
        }
        Write-Log ("No weather server found in strict scan ({0:N1}s)" -f $stopwatch.Elapsed.TotalSeconds)
        return $null
    }
    Start-Job -ScriptBlock { Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue } | Out-Null
}

# Function to create or update startup shortcut
function Set-StartupShortcut {
    param(
        [bool]$enable
    )
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $shortcutPath = Join-Path $startupFolder "TabletDeskLauncher.lnk"
    if ($enable) {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = "powershell.exe"
        $Shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" --startup"
        $Shortcut.WorkingDirectory = (Split-Path $PSCommandPath)
        $Shortcut.Save()
        Write-Log ("Created startup shortcut")
    } else {
        if (Test-Path $shortcutPath) {
            Remove-Item $shortcutPath -Force
            Write-Log ("Removed startup shortcut")
        }
    }
}

# Function to check if spacedesk.exe is running
function Is-SpacedeskRunning {
    # Use only the actual process names found on this system
    $processNames = @(
        'spacedeskService',
        'spacedeskServiceTray'
    )
    $found = @()
    foreach ($name in $processNames) {
        $proc = Get-Process -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc -and $proc.Id -gt 0) {
            $found += $name
        }
    }
    if ($found.Count -gt 0) {
        Write-Log ("Spacedesk process(es) detected: {0}" -f ($found -join ', '))
        return $true
    } else {
        Write-Log ("No Spacedesk processes found.")
        return $false
    }
}

# Function to launch weather sequence
function Launch-WeatherSequence {
    param($config)
    try {
        Write-Log ("Entering Launch-WeatherSequence")
        Write-Log ("Config dump: {0}" -f ($config | Out-String))
        Write-Log ("Starting Weather Launcher sequence...")
        if (-not $config) {
            Write-Log ("[ERROR] Config object is null!")
            return
        }
        $requiredKeys = @('Url','DisplayIndex','AutoStart','ScanTimeoutSec','DeskThingSleepMs','DisplayPollInterval','MaxDisplayWaitSec','DisplayGracePeriodSec')
        foreach ($key in $requiredKeys) {
            if (-not $config.ContainsKey($key) -or $null -eq $config[$key]) {
                Write-Log ("[ERROR] Config missing or null: {0}" -f $key)
            }
        }
        if (-not (Is-SpacedeskRunning)) {
            Write-Log ("Spacedesk not running. Waiting for display connection...")
            if (-not (Wait-ForSpacedeskDisplay -maxWaitSeconds $config['MaxDisplayWaitSec'] -retryIntervalSeconds $config['DisplayPollInterval'])) {
                Write-Log ("No secondary display detected. Please make sure Spacedesk is connected.")
                return
            }
            Write-Log ("Spacedesk display connection established.")
        } else {
            Write-Log ("Spacedesk is already running.")
        }
        $urlValid = $false
        if ($config.ContainsKey('Url') -and $config['Url']) {
            Write-Log ("Testing connection to {0}" -f $config['Url'])
            $urlValid = Test-ServerConnection -url $config['Url'] -retries 1
        }
        if (-not $urlValid) {
            Write-Log ("Configured weather server not reachable. Scanning for available weather server...")
            $foundBaseUrl = Find-WeatherServer
            if ($foundBaseUrl) {
                # Always use the full weather app path when updating config
                $weatherPath = "/#/app/weatherwaves"
                # Remove any trailing slash before appending path
                if ($foundBaseUrl.EndsWith('/')) {
                    $baseTrimmed = $foundBaseUrl.TrimEnd('/')
                } else {
                    $baseTrimmed = $foundBaseUrl
                }
                $fullUrl = $baseTrimmed + $weatherPath
                Write-Log ("Weather server found by scan: {0}" -f $fullUrl)
                $config['Url'] = $fullUrl
                Save-Config -config $config
                $urlValid = $true
            } else {
                Write-Log ("No weather server found by scan. Aborting launch. Edge will not be started.")
                return
            }
        }
        if ($urlValid -and $config.ContainsKey('DisplayIndex') -and $config['DisplayIndex'] -ne $null) {
            Write-Log ("Launching Edge with URL: {0}" -f $config['Url'])
            Launch-EdgeKiosk -url $config['Url'] -displayIndex $config['DisplayIndex']
        } elseif (-not $urlValid) {
            Write-Log ("Weather server is not responding. Please check if it's running. Edge will not be started.")
        } else {
            Write-Log ("[ERROR] DisplayIndex not set in config. Cannot launch Edge.")
        }
    } catch {
        Write-Log ("Error in Launch-WeatherSequence: {0}" -f $_)
        throw
    }
}

# GUI For Configuration
function Show-ConfigDialog {
    param($config)
    # Defensive: ensure config is always a hashtable
    if ($null -eq $config -or $config -isnot [hashtable]) {
        Write-Log ("[ERROR] Show-ConfigDialog received null or invalid config, using defaults.")
        $config = $defaultConfig.Clone()
    }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "TabletDesk Settings"
    $form.Size = New-Object System.Drawing.Size(750, 750)
    $form.MinimumSize = New-Object System.Drawing.Size(700, 650)
    $form.FormBorderStyle = "Sizable"

    # URL Label
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Location = New-Object System.Drawing.Point(20, 20)
    $urlLabel.Size = New-Object System.Drawing.Size(100, 20)
    $urlLabel.Text = "Weather URL:"
    $form.Controls.Add($urlLabel)
    # URL TextBox
    $urlTextBox = New-Object System.Windows.Forms.TextBox
    $urlTextBox.Location = New-Object System.Drawing.Point(130, 20)
    $urlTextBox.Size = New-Object System.Drawing.Size(320, 20)
    $urlTextBox.Text = $config['Url']
    $form.Controls.Add($urlTextBox)
    # Progress bar for server detection
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(130, 75)
    $progressBar.Size = New-Object System.Drawing.Size(320, 20)
    $progressBar.Style = 'Continuous'
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)
    # Find Server button
    $findServerButton = New-Object System.Windows.Forms.Button
    $findServerButton.Location = New-Object System.Drawing.Point(130, 45)
    $findServerButton.Size = New-Object System.Drawing.Size(180, 23)
    $findServerButton.Text = "Auto-Detect Weather Server"
    $findServerButton.Add_Click({
        $findServerButton.Enabled = $false
        $findServerButton.Text = "Scanning..."
        $progressBar.Value = 0
        $progressBar.Visible = $true
        # Simulate progress
        for ($p = 0; $p -le 100; $p += 10) {
            Start-Sleep -Milliseconds 100
            $progressBar.Value = $p
            $form.Refresh()
        }
        $foundUrl = Find-WeatherServer
        $progressBar.Value = 100
        $findServerButton.Text = "Auto-Detect Weather Server"
        $findServerButton.Enabled = $true
        $progressBar.Visible = $false
        if ($foundUrl) {
            $urlTextBox.Text = $foundUrl
            Write-Log ("Weather server found at {0}" -f $foundUrl)
        } else {
            Write-Log ("No weather server found on your subnet.")
        }
        return $null
    }) | Out-Null
    $form.Controls.Add($findServerButton)
    # Display selection
    $displayLabel = New-Object System.Windows.Forms.Label
    $displayLabel.Location = New-Object System.Drawing.Point(20, 110)
    $displayLabel.Size = New-Object System.Drawing.Size(150, 20)
    $displayLabel.Text = "Target Display:"
    $form.Controls.Add($displayLabel)
    $displayCombo = New-Object System.Windows.Forms.ComboBox
    $displayCombo.Location = New-Object System.Drawing.Point(180, 110)
    $displayCombo.Size = New-Object System.Drawing.Size(300, 20)
    $displayCombo.DropDownStyle = "DropDownList"
    $displays = [System.Windows.Forms.Screen]::AllScreens
    for ($i = 0; $i -lt $displays.Count; $i++) {
        $display = $displays[$i]
        $isPrimary = if ($display.Primary) { " (Primary)" } else { "" }
        $displayCombo.Items.Add(("Display {0} - {1}x{2}{3}" -f $i, $display.Bounds.Width, $display.Bounds.Height, $isPrimary))
    }
    # Remember last display
    $lastDisplayFile = Join-Path $env:USERPROFILE "TabletDeskLastDisplay.txt"
    $lastDisplayIndex = $null
    if (Test-Path $lastDisplayFile) {
        try { $lastDisplayIndex = [int](Get-Content $lastDisplayFile -Raw) } catch {}
    }
    if ($lastDisplayIndex -ne $null -and $lastDisplayIndex -ge 0 -and $lastDisplayIndex -lt $displays.Count) {
        $displayCombo.SelectedIndex = $lastDisplayIndex
    } elseif ($config['DisplayIndex'] -lt $displays.Count -and $config['DisplayIndex'] -ge 0) {
        $displayCombo.SelectedIndex = $config['DisplayIndex']
    } elseif ($displays.Count -ge 3) {
        $displayCombo.SelectedIndex = 2
    } elseif ($displays.Count -ge 2) {
        $displayCombo.SelectedIndex = 1
    } else {
        $displayCombo.SelectedIndex = 0
    }
    $form.Controls.Add($displayCombo)
    # Save last display on save
    $form.add_FormClosing({
        try {
            $selIdx = $displayCombo.SelectedIndex
            if ($selIdx -ge 0) {
                Set-Content -Path $lastDisplayFile -Value $selIdx
            }
        } catch {}
    })
    # Find Spacedesk button
    $findButton = New-Object System.Windows.Forms.Button
    $findButton.Location = New-Object System.Drawing.Point(500, 110)
    $findButton.Size = New-Object System.Drawing.Size(150, 23)
    $findButton.Text = "Find Spacedesk"
    $findButton.Add_Click({
        $spacedeskIndex = Find-SpacedeskDisplay
        if ($spacedeskIndex -lt $displayCombo.Items.Count) {
            $displayCombo.SelectedIndex = $spacedeskIndex
            Write-Log ("Found Spacedesk display at index {0}" -f $spacedeskIndex)
        } else {
            Write-Log ("Could not identify Spacedesk display")
        }
        return $null
    }) | Out-Null
    $form.Controls.Add($findButton)
    # Auto-start checkbox
    $autoStartCheckbox = New-Object System.Windows.Forms.CheckBox
    $autoStartCheckbox.Location = New-Object System.Drawing.Point(20, 150)
    $autoStartCheckbox.Size = New-Object System.Drawing.Size(250, 20)
    $autoStartCheckbox.Text = "Launch automatically at Windows startup"
    $autoStartCheckbox.Checked = $config['AutoStart']
    $form.Controls.Add($autoStartCheckbox)
    # Hotkey label and textbox
    $hotkeyLabel = New-Object System.Windows.Forms.Label
    $hotkeyLabel.Location = New-Object System.Drawing.Point(20, 180)
    $hotkeyLabel.Size = New-Object System.Drawing.Size(80, 20)
    $hotkeyLabel.Text = "Hotkey:"
    $form.Controls.Add($hotkeyLabel)
    $hotkeyBox = New-Object System.Windows.Forms.TextBox
    $hotkeyBox.Location = New-Object System.Drawing.Point(180, 180)
    $hotkeyBox.Size = New-Object System.Drawing.Size(200, 20)
    if ($config -and $config.ContainsKey('HotKey') -and $config['HotKey']) {
        $hotkeyBox.Text = $config['HotKey']
    } else {
        $hotkeyBox.Text = "Ctrl+Alt+W" # default
    }
    $form.Controls.Add($hotkeyBox)
    # Customise Kiosk checkbox (robust)
    $customKioskCheckbox = New-Object System.Windows.Forms.CheckBox
    $customKioskCheckbox.Location = New-Object System.Drawing.Point(20, 500)
    $customKioskCheckbox.Size = New-Object System.Drawing.Size(180, 25)
    $customKioskCheckbox.Text = "Customise Kiosk"
    if ($config -and $config.ContainsKey('CustomKiosk')) {
        $customKioskCheckbox.Checked = $config['CustomKiosk']
    } else {
        $customKioskCheckbox.Checked = $false
    }
    $form.Controls.Add($customKioskCheckbox)
    $customiseButton = New-Object System.Windows.Forms.Button
    $customiseButton.Location = New-Object System.Drawing.Point(220, 500)
    $customiseButton.Size = New-Object System.Drawing.Size(140, 25)
    $customiseButton.Text = "Kiosk Options"
    $customiseButton.Enabled = $customKioskCheckbox.Checked
    $form.Controls.Add($customiseButton)
    $customKioskCheckbox.Add_CheckedChanged({ $customiseButton.Enabled = $customKioskCheckbox.Checked })
    $customiseButton.Add_Click({
        if ($null -eq $config -or $config -isnot [hashtable]) { $config = $defaultConfig.Clone() }
        $config = Show-KioskOptionsDialog $config
        if ($null -eq $config -or $config -isnot [hashtable]) { $config = $defaultConfig.Clone() }
    })
    # Max wait time
    $waitLabel = New-Object System.Windows.Forms.Label
    $waitLabel.Location = New-Object System.Drawing.Point(20, 210)
    $waitLabel.Size = New-Object System.Drawing.Size(300, 20)
    $waitLabel.Text = "Max wait time for display (seconds):"
    $form.Controls.Add($waitLabel)
    $waitNumeric = New-Object System.Windows.Forms.NumericUpDown
    $waitNumeric.Location = New-Object System.Drawing.Point(340, 210)
    $waitNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $waitNumeric.Minimum = 10
    $waitNumeric.Maximum = 300
    $waitValue = [int]$config['MaxDisplayWaitSec']
    if ($waitValue -lt $waitNumeric.Minimum) { $waitValue = $waitNumeric.Minimum }
    if ($waitValue -gt $waitNumeric.Maximum) { $waitValue = $waitNumeric.Maximum }
    $waitNumeric.Value = $waitValue
    $form.Controls.Add($waitNumeric)
    # Server retries
    $retriesLabel = New-Object System.Windows.Forms.Label
    $retriesLabel.Location = New-Object System.Drawing.Point(20, 240)
    $retriesLabel.Size = New-Object System.Drawing.Size(300, 20)
    $retriesLabel.Text = "Server connection retries:"
    $form.Controls.Add($retriesLabel)
    $retriesNumeric = New-Object System.Windows.Forms.NumericUpDown
    $retriesNumeric.Location = New-Object System.Drawing.Point(340, 240)
    $retriesNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $retriesNumeric.Minimum = 1
    $retriesNumeric.Maximum = 20
    $retriesValue = [int]$config['ServerCheckRetries']
    if ($retriesValue -lt $retriesNumeric.Minimum) { $retriesValue = $retriesNumeric.Minimum }
    if ($retriesValue -gt $retriesNumeric.Maximum) { $retriesValue = $retriesNumeric.Maximum }
    $retriesNumeric.Value = $retriesValue
    $form.Controls.Add($retriesNumeric)
    # Scan timeout
    $scanTimeoutLabel = New-Object System.Windows.Forms.Label
    $scanTimeoutLabel.Location = New-Object System.Drawing.Point(20, 270)
    $scanTimeoutLabel.Size = New-Object System.Drawing.Size(300, 20)
    $scanTimeoutLabel.Text = "Scan timeout (seconds):"
    $form.Controls.Add($scanTimeoutLabel)
    $scanTimeoutNumeric = New-Object System.Windows.Forms.NumericUpDown
    $scanTimeoutNumeric.Location = New-Object System.Drawing.Point(340, 270)
    $scanTimeoutNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $scanTimeoutNumeric.Minimum = 1
    $scanTimeoutNumeric.Maximum = 60
    $scanTimeoutValue = [int]$config['ScanTimeoutSec']
    if ($scanTimeoutValue -lt $scanTimeoutNumeric.Minimum) { $scanTimeoutValue = $scanTimeoutNumeric.Minimum }
    if ($scanTimeoutValue -gt $scanTimeoutNumeric.Maximum) { $scanTimeoutValue = $scanTimeoutNumeric.Maximum }
    $scanTimeoutNumeric.Value = $scanTimeoutValue
    $form.Controls.Add($scanTimeoutNumeric)
    # DeskThing sleep
    $deskThingSleepLabel = New-Object System.Windows.Forms.Label
    $deskThingSleepLabel.Location = New-Object System.Drawing.Point(20, 300)
    $deskThingSleepLabel.Size = New-Object System.Drawing.Size(300, 20)
    $deskThingSleepLabel.Text = "DeskThing sleep (ms):"
    $form.Controls.Add($deskThingSleepLabel)
    $deskThingSleepNumeric = New-Object System.Windows.Forms.NumericUpDown
    $deskThingSleepNumeric.Location = New-Object System.Drawing.Point(340, 300)
    $deskThingSleepNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $deskThingSleepNumeric.Minimum = 0
    $deskThingSleepNumeric.Maximum = 2000
    $deskThingSleepValue = [int]$config['DeskThingSleepMs']
    if ($deskThingSleepValue -lt $deskThingSleepNumeric.Minimum) { $deskThingSleepValue = $deskThingSleepNumeric.Minimum }
    if ($deskThingSleepValue -gt $deskThingSleepNumeric.Maximum) { $deskThingSleepValue = $deskThingSleepNumeric.Maximum }
    $deskThingSleepNumeric.Value = $deskThingSleepValue
    $form.Controls.Add($deskThingSleepNumeric)
    # Display poll interval
    $displayPollLabel = New-Object System.Windows.Forms.Label
    $displayPollLabel.Location = New-Object System.Drawing.Point(20, 330)
    $displayPollLabel.Size = New-Object System.Drawing.Size(300, 20)
    $displayPollLabel.Text = "Display poll interval (s):"
    $form.Controls.Add($displayPollLabel)
    $displayPollNumeric = New-Object System.Windows.Forms.NumericUpDown
    $displayPollNumeric.Location = New-Object System.Drawing.Point(340, 330)
    $displayPollNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $displayPollNumeric.Minimum = 0.1
    $displayPollNumeric.Maximum = 5
    $displayPollNumeric.DecimalPlaces = 1
    $displayPollNumeric.Increment = 0.1
    $displayPollValue = [decimal]$config['DisplayPollInterval']
    if ($displayPollValue -lt $displayPollNumeric.Minimum) { $displayPollValue = $displayPollNumeric.Minimum }
    if ($displayPollValue -gt $displayPollNumeric.Maximum) { $displayPollValue = $displayPollNumeric.Maximum }
    $displayPollNumeric.Value = $displayPollValue
    $form.Controls.Add($displayPollNumeric)
    # Max display wait
    $maxDisplayWaitLabel = New-Object System.Windows.Forms.Label
    $maxDisplayWaitLabel.Location = New-Object System.Drawing.Point(20, 360)
    $maxDisplayWaitLabel.Size = New-Object System.Drawing.Size(300, 20)
    $maxDisplayWaitLabel.Text = "Max display wait (s):"
    $form.Controls.Add($maxDisplayWaitLabel)
    $maxDisplayWaitNumeric = New-Object System.Windows.Forms.NumericUpDown
    $maxDisplayWaitNumeric.Location = New-Object System.Drawing.Point(340, 360)
    $maxDisplayWaitNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $maxDisplayWaitNumeric.Minimum = 1
    $maxDisplayWaitNumeric.Maximum = 120
    $maxDisplayWaitValue = [int]$config['MaxDisplayWaitSec']
    if ($maxDisplayWaitValue -lt $maxDisplayWaitNumeric.Minimum) { $maxDisplayWaitValue = $maxDisplayWaitNumeric.Minimum }
    if ($maxDisplayWaitValue -gt $maxDisplayWaitNumeric.Maximum) { $maxDisplayWaitValue = $maxDisplayWaitNumeric.Maximum }
    $maxDisplayWaitNumeric.Value = $maxDisplayWaitValue
    $form.Controls.Add($maxDisplayWaitNumeric)
    # Display grace period
    $displayGracePeriodLabel = New-Object System.Windows.Forms.Label
    $displayGracePeriodLabel.Location = New-Object System.Drawing.Point(20, 390)
    $displayGracePeriodLabel.Size = New-Object System.Drawing.Size(300, 20)
    $displayGracePeriodLabel.Text = "Display grace period (s):"
    $form.Controls.Add($displayGracePeriodLabel)
    $displayGracePeriodNumeric = New-Object System.Windows.Forms.NumericUpDown
    $displayGracePeriodNumeric.Location = New-Object System.Drawing.Point(340, 390)
    $displayGracePeriodNumeric.Size = New-Object System.Drawing.Size(60, 20)
    $displayGracePeriodNumeric.Minimum = 0
    $displayGracePeriodNumeric.Maximum = 10
    $displayGracePeriodNumeric.DecimalPlaces = 1
    $displayGracePeriodNumeric.Increment = 0.1
    $displayGracePeriodValue = [decimal]$config['DisplayGracePeriodSec']
    if ($displayGracePeriodValue -lt $displayGracePeriodNumeric.Minimum) { $displayGracePeriodValue = $displayGracePeriodNumeric.Minimum }
    if ($displayGracePeriodValue -gt $displayGracePeriodNumeric.Maximum) { $displayGracePeriodValue = $displayGracePeriodNumeric.Maximum }
    $displayGracePeriodNumeric.Value = $displayGracePeriodValue
    $form.Controls.Add($displayGracePeriodNumeric)
    # Browser selection
    $browserGroup = New-Object System.Windows.Forms.GroupBox
    $browserGroup.Location = New-Object System.Drawing.Point(20, 420)
    $browserGroup.Size = New-Object System.Drawing.Size(430, 60)
    $browserGroup.Text = "Browser Option"
    $edgeRadio = New-Object System.Windows.Forms.RadioButton
    $edgeRadio.Location = New-Object System.Drawing.Point(10, 20)
    $edgeRadio.Size = New-Object System.Drawing.Size(200, 20)
    $edgeRadio.Text = "Microsoft Edge (Kiosk Mode)"
    $edgeRadio.Checked = $config['UseEdge']
    $browserGroup.Controls.Add($edgeRadio)
    $form.Controls.Add($browserGroup)
    # Save button
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(250, 650)
    $saveButton.Size = New-Object System.Drawing.Size(140, 35)
    $saveButton.Text = "Save Settings"
    $saveButton.Font = New-Object System.Drawing.Font($saveButton.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $saveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $saveButton
    $form.Controls.Add($saveButton)
    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(450, 650)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    # Import/Export Config Buttons
    $importButton = New-Object System.Windows.Forms.Button
    $importButton.Location = New-Object System.Drawing.Point(20, 700)
    $importButton.Size = New-Object System.Drawing.Size(120, 30)
    $importButton.Text = "Import Config"
    $importButton.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $imported = Get-Content -Path $ofd.FileName -Raw | ConvertFrom-Json
                foreach ($prop in $imported.PSObject.Properties) {
                    $config[$prop.Name] = $prop.Value
                }
                Write-Log ("Imported config from {0}" -f $ofd.FileName)
                [System.Windows.Forms.MessageBox]::Show("Configuration imported! Please review settings before saving.", "Import Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                Write-Log ("[ERROR] Failed to import config: {0}" -f $_)
                [System.Windows.Forms.MessageBox]::Show("Failed to import configuration.", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $form.Controls.Add($importButton)
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Location = New-Object System.Drawing.Point(160, 700)
    $exportButton.Size = New-Object System.Drawing.Size(120, 30)
    $exportButton.Text = "Export Config"
    $exportButton.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $sfd.FileName = "TabletDeskConfig.json"
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $json = $config | ConvertTo-Json
                Set-Content -Path $sfd.FileName -Value $json
                Write-Log ("Exported config to {0}" -f $sfd.FileName)
                [System.Windows.Forms.MessageBox]::Show("Configuration exported!", "Export Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                Write-Log ("[ERROR] Failed to export config: {0}" -f $_)
                [System.Windows.Forms.MessageBox]::Show("Failed to export configuration.", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $form.Controls.Add($exportButton)
    # On save, update config fields (robust)
    $form.add_FormClosing({
        if ($hotkeyBox) { $config['HotKey'] = $hotkeyBox.Text }
        if ($customKioskCheckbox) { $config['CustomKiosk'] = $customKioskCheckbox.Checked }
        if ($urlTextBox) { $config['Url'] = $urlTextBox.Text }
        $selectedDisplayIndex = if ($displayCombo) { $displayCombo.SelectedIndex } else { 0 }
        if ($selectedDisplayIndex -lt 0) { $selectedDisplayIndex = 0 }
        if ($displayCombo -and $selectedDisplayIndex -ge $displayCombo.Items.Count) { $selectedDisplayIndex = $displayCombo.Items.Count - 1 }
        $config['DisplayIndex'] = $selectedDisplayIndex
        if ($autoStartCheckbox) { $config['AutoStart'] = $autoStartCheckbox.Checked }
        if ($waitNumeric) { $config['MaxDisplayWaitSec'] = [int]$waitNumeric.Value }
        if ($retriesNumeric) { $config['ServerCheckRetries'] = [int]$retriesNumeric.Value }
        if ($scanTimeoutNumeric) { $config['ScanTimeoutSec'] = [int]$scanTimeoutNumeric.Value }
        if ($deskThingSleepNumeric) { $config['DeskThingSleepMs'] = [int]$deskThingSleepNumeric.Value }
        if ($displayPollNumeric) { $config['DisplayPollInterval'] = [double]$displayPollNumeric.Value }
        if ($displayGracePeriodNumeric) { $config['DisplayGracePeriodSec'] = [double]$displayGracePeriodNumeric.Value }
        if ($edgeRadio) { $config['UseEdge'] = $edgeRadio.Checked }
        Set-StartupShortcut -enable $config['AutoStart']
        return $config
    })
    # Show dialog and capture result
    $dialogResult = $form.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Log ("Settings saved via tray.")
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Show-ConfigDialog returning null or invalid config, using defaults.")
            return $defaultConfig.Clone()
        }
        return $config
    } else {
        Write-Log ("Settings dialog cancelled or closed.")
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Show-ConfigDialog returning null or invalid config, using defaults.")
            return $defaultConfig.Clone()
        }
        return $config
    }
}

# KIOSK OPTIONS DIALOG (CHECKBOX-ONLY, NO STRING BAR)
function Show-KioskOptionsDialog {
    param($config)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Kiosk Options"
    $form.Size = New-Object System.Drawing.Size(400, 270)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $argLabel = New-Object System.Windows.Forms.Label
    $argLabel.Location = New-Object System.Drawing.Point(20, 20)
    $argLabel.Size = New-Object System.Drawing.Size(200, 20)
    $argLabel.Text = "Extra Edge Kiosk Arguments:"
    $form.Controls.Add($argLabel)

    # Common options - move down for clarity
    $chkInPrivate = New-Object System.Windows.Forms.CheckBox
    $chkInPrivate.Location = New-Object System.Drawing.Point(40, 55)
    $chkInPrivate.Size = New-Object System.Drawing.Size(300, 20)
    $chkInPrivate.Text = "InPrivate Mode (--inprivate)"
    $form.Controls.Add($chkInPrivate)

    $chkNoFirstRun = New-Object System.Windows.Forms.CheckBox
    $chkNoFirstRun.Location = New-Object System.Drawing.Point(40, 80)
    $chkNoFirstRun.Size = New-Object System.Drawing.Size(300, 20)
    $chkNoFirstRun.Text = "No First Run (--no-first-run)"
    $form.Controls.Add($chkNoFirstRun)

    $chkDisablePrint = New-Object System.Windows.Forms.CheckBox
    $chkDisablePrint.Location = New-Object System.Drawing.Point(40, 105)
    $chkDisablePrint.Size = New-Object System.Drawing.Size(300, 20)
    $chkDisablePrint.Text = "Disable Printing (--kiosk-printing=false)"
    $form.Controls.Add($chkDisablePrint)

    $chkHideScroll = New-Object System.Windows.Forms.CheckBox
    $chkHideScroll.Location = New-Object System.Drawing.Point(40, 130)
    $chkHideScroll.Size = New-Object System.Drawing.Size(300, 20)
    $chkHideScroll.Text = "Hide Scrollbars (--hide-scrollbars)"
    $form.Controls.Add($chkHideScroll)

    $chkDisablePopup = New-Object System.Windows.Forms.CheckBox
    $chkDisablePopup.Location = New-Object System.Drawing.Point(40, 155)
    $chkDisablePopup.Size = New-Object System.Drawing.Size(300, 20)
    $chkDisablePopup.Text = "Disable Popups (--disable-popup-blocking)"
    $form.Controls.Add($chkDisablePopup)

    # Load initial state from config
    $args = if ($config.ContainsKey('KioskArgs')) { $config['KioskArgs'] } else { "" }
    $chkInPrivate.Checked = $args -match "--inprivate"
    $chkNoFirstRun.Checked = $args -match "--no-first-run"
    $chkDisablePrint.Checked = $args -match "--kiosk-printing=false"
    $chkHideScroll.Checked = $args -match "--hide-scrollbars"
    $chkDisablePopup.Checked = $args -match "--disable-popup-blocking"

    # Save button
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(150, 200)
    $saveButton.Size = New-Object System.Drawing.Size(100, 30)
    $saveButton.Text = "Save"
    $saveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $saveButton
    $form.Controls.Add($saveButton)

    $dialogResult = $form.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $parts = @()
        if ($chkInPrivate.Checked) { $parts += "--inprivate" }
        if ($chkNoFirstRun.Checked) { $parts += "--no-first-run" }
        if ($chkDisablePrint.Checked) { $parts += "--kiosk-printing=false" }
        if ($chkHideScroll.Checked) { $parts += "--hide-scrollbars" }
        if ($chkDisablePopup.Checked) { $parts += "--disable-popup-blocking" }
        $config['KioskArgs'] = ($parts -join " ")
        return $config
    }
    return $config
}

# HOTKEY REGISTRATION
Add-Type @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
public class HotKeyRegister {
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
"@

# --- INLINE HOTKEY HANDLER ---
function Start-HotkeyListener {
    param([string]$hotkey)
    try {
        Write-Log ("[HOTKEY] [INLINE] Entering Start-HotkeyListener for $hotkey")
        Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] [INLINE] Entering Start-HotkeyListener for $hotkey"
        $triggerFile = $global:hotkeyTriggerFile
        $parsed = Parse-Hotkey $hotkey
        $mod = $parsed.mod
        $vk = $parsed.vk
        $key = $parsed.key
        Write-Log ("[HOTKEY] (Inline) Attempting to register hotkey: $hotkey (mod=$mod, vk=$vk, key=$key)")
        Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] (Inline) Attempting to register hotkey: $hotkey (mod=$mod, vk=$vk, key=$key)"
        try {
            Add-Type @'
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
public class HotKeyRegister {
    [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
'@
            $id = 0x1234
            $form = New-Object System.Windows.Forms.Form
            $form.ShowInTaskbar = $false
            $form.WindowState = 'Minimized'
            $form.Visible = $false
            $success = [HotKeyRegister]::RegisterHotKey($form.Handle, $id, $mod, $vk)
            Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] (Inline) RegisterHotKey called: mod=$mod vk=$vk key=$key success=$success"
            if (-not $success) {
                [System.Windows.Forms.MessageBox]::Show("Failed to register hotkey: $key. It may already be in use.", "Hotkey Error")
                Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] (Inline) RegisterHotKey FAILED: mod=$mod vk=$vk key=$key"
            }
            $form.Add_FormClosed({ [HotKeyRegister]::UnregisterHotKey($form.Handle, $id) })
            $form.Add_HandleCreated({ [HotKeyRegister]::RegisterHotKey($form.Handle, $id, $mod, $vk) })
            $form.Add_WndProc({
                param($sender, $e)
                if ($e.Msg -eq 0x0312) {
                    Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] (Inline) WM_HOTKEY received! Writing trigger file."
                    Set-Content -Path $triggerFile -Value ([DateTime]::UtcNow.ToString('o'))
                }
            })
            $form.ShowDialog() | Out-Null
        } catch {
            Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] (Inline) ERROR: $($_.Exception.Message)"
        }
    } catch {
        Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] [INLINE] ERROR in Start-HotkeyListener: $($_.Exception.Message)"
    }
}

# Start hotkey listener at script startup (hardcoded for diagnostics)
Start-HotkeyListener 'Ctrl+Alt+W'

# Function to check if server is running
function Get-ServerStatus {
    param([string]$url)
    try {
        $resp = Invoke-WebRequest -Uri $url -TimeoutSec 2 -UseBasicParsing -ErrorAction Stop
        if ($resp.StatusCode -eq 200) { return $true }
    } catch {}
    return $false
}

# DeskThing early initialization (before any scans or logic)
function Ensure-DeskThing {
    $deskthingProcess = Get-Process -Name DeskThing -ErrorAction SilentlyContinue
    $deskThingPath = "C:\Users\2e0dn\AppData\Local\Programs\deskthing\DeskThing.exe"
    if (-not $deskthingProcess) {
        if (Test-Path $deskThingPath) {
            Write-Log ("DeskThing not running. Launching DeskThing (minimized)...")
            try {
                Start-Process -FilePath $deskThingPath -WindowStyle Minimized
                $sleepMs = 100
                if ($global:config -and $global:config.ContainsKey('DeskThingSleepMs')) { $sleepMs = [int]$global:config['DeskThingSleepMs'] }
                if ($sleepMs -gt 0) { Start-Sleep -Milliseconds $sleepMs }
            } catch {
                Write-Log ("Failed to start DeskThing: {0}" -f $_)
            }
        } else {
            Write-Log ("DeskThing.exe not found at $deskThingPath")
        }
    } else {
        Write-Log ("DeskThing is already running.")
    }
}

# System Tray Application
function Start-TrayApplication {
    $config = Load-Config
    $sequenceTriggered = $false
    # Auto-detect weather server on startup if enabled
    if ($config['AutoDetectOnStartup']) {
        $foundUrl = Find-WeatherServer
        if ($foundUrl -and $foundUrl -ne $config['Url']) {
            $config['Url'] = $foundUrl
            Save-Config -config $config
            Write-Log ("Auto-detected weather server at {0} and updated config." -f $foundUrl)
        } else {
            Write-Log ("No new weather server detected on startup.")
        }
        $sequenceTriggered = $true
    }
    # Only run the full launch sequence after auto-detect or on demand, without another Find-WeatherServer
    if ($sequenceTriggered -or $MyInvocation.Line -match "--startup") {
        Launch-WeatherSequence $config
    }
    # Use system icon for guaranteed compatibility
    $iconPath = "$env:USERPROFILE\Desktop\TabletDeskTrayIcon.ico"
    $icon = if (Test-Path $iconPath) { [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath) } else { [System.Drawing.SystemIcons]::Application }
    # Create the tray icon
    $trayIcon = New-Object System.Windows.Forms.NotifyIcon
    $trayIcon.Icon = $icon
    $trayIcon.Text = "TabletDesk"
    $trayIcon.Visible = $true
    # Create context menu for tray icon
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    # Launch Weather item
    $launchItem = $contextMenu.Items.Add("Launch Desk Thing Display")
    $launchItem.Add_Click({
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Config is null or invalid in Launch Desk Thing Display handler.")
            $config = $defaultConfig.Clone()
        }
        Launch-WeatherSequence $config
        return $null
    }) | Out-Null
    # Settings item
    $settingsItem = $contextMenu.Items.Add("Settings")
    $settingsItem.Add_Click({
        Write-Log ("[EVENT] Settings tray menu clicked.")
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Config is null or invalid in Settings handler.")
            $config = $defaultConfig.Clone()
        }
        $config = Show-ConfigDialog -config $config
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Config is null or invalid after Show-ConfigDialog, using defaults.")
            $config = $defaultConfig.Clone()
        }
        Save-Config -config $config
        return $null
    }) | Out-Null
    # View log item
    $logItem = $contextMenu.Items.Add("View Log")
    $logItem.Add_Click({
        if (![string]::IsNullOrWhiteSpace($global:logPath) -and (Test-Path $global:logPath)) {
            Start-Process notepad.exe -ArgumentList $global:logPath
            Write-Log ("Viewing log file")
        } else {
            Write-Log ("Log path is empty or file does not exist.")
        }
        return $null
    }) | Out-Null
    # Retry Connection item
    $retryItem = $contextMenu.Items.Add("Retry Server Detection")
    $retryItem.Add_Click({
        if ($null -eq $config -or $config -isnot [hashtable]) {
            Write-Log ("[ERROR] Config is null or invalid in Retry Server Detection handler.")
            $config = $defaultConfig.Clone()
        }
        Write-Log ("Manual server detection triggered from tray.")
        $foundUrl = Find-WeatherServer
        if ($foundUrl) {
            $config['Url'] = $foundUrl
            Save-Config -config $config
            Write-Log ("Weather server found at {0}" -f $foundUrl)
        } else {
            Write-Log ("No weather server found on your subnet.")
        }
        return $null
    }) | Out-Null
    # Exit item
    $exitItem = $contextMenu.Items.Add("Exit")
    $exitItem.Add_Click({
        $trayIcon.Visible = $false
        $trayIcon.Dispose()
        [System.Windows.Forms.Application]::Exit()
        Stop-Process $pid
        return $null
    }) | Out-Null
    $trayIcon.ContextMenuStrip = $contextMenu
    $appContext = New-Object System.Windows.Forms.ApplicationContext
    [System.Windows.Forms.Application]::Run($appContext)
}

# Main script
try {
    Write-Log ("[CONFIG] Loading configuration...")
    $config = $null
    try {
        $config = Load-Config
        Write-Log ("[CONFIG] Config loaded: {0}" -f ($config | Out-String))
        $config = Validate-Config $config
        Write-Log ("[CONFIG] Config after validation: {0}" -f ($config | Out-String))
    } catch {
        Write-Log ("[CONFIG][ERROR] Failed to load or validate config: {0}" -f $_)
        $config = $defaultConfig.Clone()
        Write-Log ("[CONFIG] Using default config: {0}" -f ($config | Out-String))
    }

    # Add Write-AutoHotKeyScript function to generate/update an AHK script using the current hotkey from config
    function Write-AutoHotKeyScript {
        param(
            [string]$Hotkey = 'Ctrl+Alt+W',
            [string]$ScriptPath = "$env:USERPROFILE\Desktop\TabletDesk.ahk",
            [string]$PSScriptPath = "$PSScriptRoot\TabletDesk-Launcher.ps1"
        )
        # Convert PowerShell hotkey format to AHK format
        $ahkKey = $Hotkey -replace '(?i)ctrl', '^' -replace '(?i)alt', '!' -replace '(?i)shift', '+' -replace '(?i)win', '#'
        $ahkKey = $ahkKey -replace '\s*\+\s*', ''
        if ($ahkKey -match '([\^\!\+#]+)([A-Za-z0-9]+)') {
            $ahkKey = $Matches[1] + $Matches[2].ToLower()
        } else {
            $ahkKey = '^!w' # fallback
        }
        # Generate AHK v2-compatible script with block braces
        $ahk = @"
#SingleInstance force
$ahkKey::{
    Run('powershell.exe -ExecutionPolicy Bypass -File `"$PSScriptPath`" --hotkey', , "Hide")
}
"@
        Set-Content -Path $ScriptPath -Value $ahk -Encoding UTF8
        Write-Log ("[AHK] AutoHotKey v2 script written to $ScriptPath for hotkey $Hotkey")
        try {
            $ahkExe = "C:\Program Files\AutoHotkey\v2\AutoHotkey.exe"
            Start-Process -FilePath $ahkExe -ArgumentList "`"$ScriptPath`"" -WindowStyle Hidden
            Write-Log ("[AHK] Launched AutoHotKey script: $ScriptPath with ${ahkExe}")
        } catch {
            Write-Log ("[AHK] ERROR: Failed to launch AutoHotKey script with ${ahkExe}: $($_.Exception.Message)")
        }
    }

    # Call Write-AutoHotKeyScript after config is loaded and validated
    if ($global:config['HotKey']) {
        Write-AutoHotKeyScript -Hotkey $global:config['HotKey']
    }

    # --- PATCH START ---
    # Auto-launch weather display if --startup argument is present or AutoStart is enabled
    if ($args -contains "--startup" -or ($config['AutoStart'] -eq $true)) {
        Write-Log ("[AUTO] Auto-launching weather display on startup.")
        Launch-WeatherSequence $config
    }
    # --- PATCH END ---

    # Start the tray application as usual
    Start-TrayApplication
} catch {
    Write-Log ("[FATAL] Unhandled error in main script: {0}" -f $_)
}

try {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("TabletDesk loaded! (WinForms test)", "TabletDesk")
    Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] [MAIN] WinForms MessageBox test shown."
} catch {
    Add-Content -Path "$env:TEMP\TabletDesk_hotkey_debug.log" -Value "[$(Get-Date -Format o)] [MAIN] ERROR: WinForms MessageBox failed: $($_.Exception.Message)"
}

# -------------------------------
# Display Detection & Waiting
# -------------------------------
function Wait-ForSpacedeskDisplay {
    param(
        [int]$maxWaitSeconds = 30,
        [double]$retryIntervalSeconds = 0.5,
        [int]$minDisplays = 2
    )
    $elapsed = 0
    while ($elapsed -lt $maxWaitSeconds) {
        $screens = [System.Windows.Forms.Screen]::AllScreens
        if ($screens.Count -ge $minDisplays) {
            Write-Log ("Detected $($screens.Count) displays. Proceeding.")
            return $screens
        }
        Start-Sleep -Seconds $retryIntervalSeconds
        $elapsed += $retryIntervalSeconds
    }
    Write-Log ("Timeout waiting for Spacedesk display. Only $($screens.Count) detected.")
    return $null
}

# -------------------------------
# Edge Launch (Kiosk Mode)
# -------------------------------
function Launch-EdgeKiosk {
    param(
        [string]$url,
        [int]$displayIndex = 1
    )
    $displays = [System.Windows.Forms.Screen]::AllScreens
    Write-Log ("DEBUG: Detected {0} display(s)." -f $displays.Count)
    $i = 0
    foreach ($display in $displays) {
        Write-Log ("DEBUG: Display {0}: Bounds=({1},{2},{3},{4}), Primary={5}" -f $i, $display.Bounds.X, $display.Bounds.Y, $display.Bounds.Width, $display.Bounds.Height, $display.Primary)
        $i++
    }
    # Select the first non-primary display if available
    $targetDisplay = $displays | Where-Object { -not $_.Primary } | Select-Object -First 1
    if (-not $targetDisplay) {
        $targetDisplay = $displays[0]
        Write-Log ("No non-primary display found. Using primary display (index 0).")
    } else {
        Write-Log ("Using non-primary display: {0}" -f $displays.IndexOf($targetDisplay))
    }
    $bounds = $targetDisplay.Bounds
    Write-Log ("Launching Edge on display {0} ({1}x{2}) for URL: {3}" -f $displays.IndexOf($targetDisplay), $bounds.Width, $bounds.Height, $url)
    try {
        $edgePath = $null
        try {
            $edgeCmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
            if ($edgeCmd) { $edgePath = $edgeCmd.Source }
            Write-Log ("Edge found in PATH: {0}" -f $edgePath)
        } catch {
            Write-Log ("Error searching PATH for Edge: {0}" -f $_)
        }
        if (-not $edgePath) {
            $possiblePaths = @(
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe"
            )
            foreach ($path in $possiblePaths) {
                Write-Log ("Checking Edge path: {0}" -f $path)
                if (Test-Path $path) {
                    $edgePath = $path
                    Write-Log ("Edge found at: {0}" -f $path)
                    break
                }
            }
        }
        if (-not $edgePath) {
            Write-Log ("Edge browser not found in PATH or default locations. Please check your Edge installation.")
            return
        }
        $extraArgs = ""
        if ($global:config['CustomKiosk'] -and $global:config['KioskArgs']) {
            $extraArgs = $global:config['KioskArgs']
        }
        $args = "-k $url --kiosk-type=fullscreen $extraArgs"
        Write-Log ("Starting Edge process: {0} {1}" -f $edgePath, $args)
        $proc = Start-Process $edgePath -ArgumentList $args -PassThru -ErrorAction Stop
        if ($proc) {
            Write-Log ("Edge process started. PID: {0}" -f $proc.Id)
        } else {
            Write-Log ("Failed to start Edge process.")
            return
        }
        # Wait for Edge window handle to appear (max 4s)
        $edgeProcess = $null
        $waited = 0
        while ($waited -lt 40) {
            $edgeProcess = Get-Process msedge -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
            if ($edgeProcess) { break }
            Start-Sleep -Milliseconds 100
            $waited++
        }
        if ($edgeProcess) {
            $hwnd = $edgeProcess.MainWindowHandle
            Write-Log ("Moving Edge window. HWND: {0}" -f $hwnd)
            Add-Type @"
            using System;
            using System.Runtime.InteropServices;
            public class WindowMover {
                [DllImport("user32.dll")]
                public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
                [DllImport("user32.dll")]
                public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
                [DllImport("user32.dll")]
                public static extern IntPtr GetForegroundWindow();
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool SetForegroundWindow(IntPtr hWnd);
            }
"@
            [WindowMover]::MoveWindow($hwnd, $bounds.X, $bounds.Y, $bounds.Width, $bounds.Height, $true)
            [WindowMover]::ShowWindow($hwnd, 3) # 3 = Maximize
            [WindowMover]::SetForegroundWindow($hwnd)
            Add-Type -AssemblyName System.Windows.Forms
            Start-Sleep -Milliseconds 500
            Write-Log ("Sending F11 to Edge window for fullscreen.")
            [System.Windows.Forms.SendKeys]::SendWait("{F11}")
            Write-Log ("F11 key sent to Edge window (once).")
            Write-Log ("Edge window moved to display {0} and set to full screen" -f $displays.IndexOf($targetDisplay))
            return
        } else {
            Write-Log ("Could not find Edge window to move after starting process. PID: {0}" -f $proc.Id)
            return
        }
    } catch {
        Write-Log ("Error launching Edge: {0}" -f $_)
        return
    }
}
