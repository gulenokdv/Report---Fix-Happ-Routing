param(
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        'ui',
        'agent',
        'watch-ui-close',
        'ensure-defaults',
        'status-lines',
        'status-json',
        'start-agent',
        'start-agent-if-enabled',
        'stop-agent',
        'process-monitor',
        'update-rulesets',
        'open-log'
    )]
    [string]$Action
    ,
    [int]$UiPid = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot = $PSScriptRoot
$AppRoot = $ScriptRoot
if ((Split-Path -Path $ScriptRoot -Leaf) -eq 'scripts') {
    $AppRoot = Split-Path -Path $ScriptRoot -Parent
}
$PatchScript = Join-Path -Path $ScriptRoot -ChildPath 'patch-happ-config.ps1'
$ConfigPath = "$env:LOCALAPPDATA\Happ\config.json"
$AppInfoPath = Join-Path -Path $AppRoot -ChildPath 'serv\happ-routing-app.json'
$SettingsPath = Join-Path -Path $AppRoot -ChildPath 'happ-routing-settings.json'
$StatePath = Join-Path -Path $AppRoot -ChildPath 'happ-routing-state.json'
$RemoteConfigPath = Join-Path -Path $AppRoot -ChildPath 'serv\happ-routing-remote-config.json'
$RemoteConfigUrl = 'https://raw.githubusercontent.com/gulenokdv/Report---Fix-Happ-Routing/main/serv/happ-routing-remote-config.json'
$LogPath = Join-Path -Path $AppRoot -ChildPath 'happ-routing.log'

$RulesDir = Join-Path -Path $AppRoot -ChildPath 'rulesets'
$DefaultAppVersion = '0.1.0'
$UiWindowTitle = 'Happ Routing Fix by gulenok91'
$AutostartRunKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$AutostartRunSubKeyPath = 'Software\Microsoft\Windows\CurrentVersion\Run'
$AutostartRunValueName = 'HappRoutingFixByGulenok91'
$ScriptSignature = 'HAPP_ROUTING_SIGNATURE=HRF_GULENOK91'
$LegacyAutostartRunValueNames = @('HappRoutingFix', 'HappRoutingGuard')
$ManagedAutostartCommandMarkers = @('happ-routing-fix.bat', 'start-happ-routing-guard.bat')
$script:LogThrottle = @{}
$script:UiShutdownHandled = $false
$script:UiExitRequested = $false
$script:ShowConfigPathPrompt = $false
$script:LastAgentConfigError = ''
$script:ConfigCriticalErrorCounts = @{}
$script:LastObservedConfigPath = ''
$script:LastObservedConfigStamp = ''



function Get-CommandExecutablePath {
    param([string]$CommandLine)

    if ([string]::IsNullOrWhiteSpace($CommandLine)) {
        return $null
    }

    $trimmed = $CommandLine.Trim()
    if ($trimmed.StartsWith('"')) {
        $closing = $trimmed.IndexOf('"', 1)
        if ($closing -gt 1) {
            return $trimmed.Substring(1, $closing - 1)
        }
    }

    $parts = $trimmed.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -gt 0) {
        return $parts[0]
    }
    return $null
}

function Test-ManagedLauncherFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -ne '.bat' -and $extension -ne '.cmd') {
        return $false
    }

    try {
        $text = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        return ($text -match [regex]::Escape($ScriptSignature))
    }
    catch {
        return $false
    }
}

function Get-LauncherPath {
    if (-not [string]::IsNullOrWhiteSpace($env:HAPP_ROUTING_LAUNCHER)) {
        $candidate = $env:HAPP_ROUTING_LAUNCHER
        if (Test-Path -LiteralPath $candidate) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    $batCandidates = Get-ChildItem -LiteralPath $AppRoot -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -in @('.bat', '.cmd')
    }
    foreach ($file in @($batCandidates)) {
        try {
            $text = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop
            if ($text -match [regex]::Escape($ScriptSignature)) {
                return $file.FullName
            }
        }
        catch {
        }
    }

    $fallback = Join-Path -Path $AppRoot -ChildPath 'happ-routing-fix.bat'
    return $fallback
}

function Write-Utf8BomFile {
    param([string]$Path, [string]$Content)
    $lastError = $null
    $encoding = New-Object System.Text.UTF8Encoding($true)
    foreach ($attempt in 1..8) {
        try {
            [System.IO.File]::WriteAllText($Path, $Content, $encoding)
            return
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds 50
        }
    }
    if ($null -ne $lastError) {
        throw $lastError
    }
}


function Read-Utf8JsonFile {
    param([string]$Path)

    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8) | ConvertFrom-Json
}

function ConvertTo-ReadableJson {
    param([object]$InputObject)

    $json = $InputObject | ConvertTo-Json -Depth 10
    return [regex]::Replace($json, '\\u([0-9a-fA-F]{4})', {
        param($match)
        return [char][Convert]::ToInt32($match.Groups[1].Value, 16)
    })
}

function Write-AppLog {
    param([string]$Message)
    Add-Content -LiteralPath $LogPath -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -Encoding UTF8
}

function Write-AppLogLimited {
    param(
        [string]$Key,
        [string]$Message,
        [int]$Limit = 15
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        Write-AppLog -Message $Message
        return
    }

    if (-not $script:LogThrottle.ContainsKey($Key)) {
        $script:LogThrottle[$Key] = 0
    }

    $script:LogThrottle[$Key] = [int]$script:LogThrottle[$Key] + 1
    $count = [int]$script:LogThrottle[$Key]

    if ($count -le $Limit) {
        Write-AppLog -Message $Message
        if ($count -eq $Limit) {
            Write-AppLog -Message ('Повторы сообщения [' + $Key + '] подавлены после ' + $Limit + ' записей.')
        }
    }
}

function Show-WindowsNotification {
    param(
        [string]$Title,
        [string]$Message
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    try {
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
        $notifyIcon.Visible = $true
        $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::None
        $notifyIcon.BalloonTipTitle = $Title
        $notifyIcon.BalloonTipText = $Message
        $notifyIcon.ShowBalloonTip(5000)

        $deadline = (Get-Date).AddSeconds(6)
        while ((Get-Date) -lt $deadline) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
        }
    }
    finally {
        $notifyIcon.Dispose()
    }
}

function Get-RunRegistryValue {
    param([string]$Name)

    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($AutostartRunSubKeyPath, $false)
        if ($null -eq $key) {
            return $null
        }
        return $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    }
    finally {
        if ($null -ne $key) {
            $key.Close()
        }
    }
}

function Set-RunRegistryValue {
    param([string]$Name, [string]$Value)

    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($AutostartRunSubKeyPath)
        if ($null -eq $key) {
            throw 'Не удалось открыть раздел автозапуска.'
        }
        $key.SetValue($Name, $Value, [Microsoft.Win32.RegistryValueKind]::String)
    }
    finally {
        if ($null -ne $key) {
            $key.Close()
        }
    }
}

function Remove-RunRegistryValue {
    param([string]$Name)

    $key = $null
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($AutostartRunSubKeyPath, $true)
        if ($null -eq $key) {
            return
        }
        if ($null -ne $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)) {
            $key.DeleteValue($Name, $false)
        }
    }
    finally {
        if ($null -ne $key) {
            $key.Close()
        }
    }
}

function Get-DefaultSettings {
    return @{
        auto_update_enabled      = $true
        auto_update_interval_min = 60
        autostart_enabled        = $false
        keep_background_on_close = $true
        custom_config_path       = ''
    }
}

function Save-Settings {
    param([hashtable]$Settings)
    Write-Utf8BomFile -Path $SettingsPath -Content (ConvertTo-ReadableJson -InputObject ([pscustomobject]$Settings))
}

function Get-EffectiveConfigPath {
    $settings = Get-Settings
    $custom = [string]$settings.custom_config_path
    if (-not [string]::IsNullOrWhiteSpace($custom)) {
        if (Test-Path -LiteralPath $custom -PathType Leaf) {
            return $custom
        }
        $joined = Join-Path -Path $custom -ChildPath 'config.json'
        if (Test-Path -LiteralPath $joined -PathType Leaf) {
            return $joined
        }
        if (Test-Path -LiteralPath $custom -PathType Container) {
            return $joined
        }
        return $custom
    }
    return "$env:LOCALAPPDATA\Happ\config.json"
}

function Add-AnnouncementError {
    param([string]$Message)
    $state = Get-State
    $errors = @()
    if ($state.ContainsKey('announcement_errors')) {
        try {
            $loaded = $state.announcement_errors | ConvertFrom-Json
            if ($loaded -is [array]) {
                $errors = @($loaded)
            }
        }
        catch {}
    }
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $errors += "[$timestamp] $Message"
    if ($errors.Count -gt 5) {
        $errors = $errors[-5..-1]
    }
    Update-State -Values @{ announcement_errors = ($errors | ConvertTo-Json -Compress) }
}

function Get-AnnouncementErrors {
    $state = Get-State
    if (-not $state.ContainsKey('announcement_errors')) {
        return @()
    }
    try {
        $loaded = $state.announcement_errors | ConvertFrom-Json
        if ($loaded -is [array]) {
            return @($loaded)
        }
        return @($loaded)
    }
    catch {
        return @()
    }
}

function Clear-AnnouncementErrors {
    Update-State -Values @{ announcement_errors = '' }
}

function Get-StateUiSignature {
    $state = Get-State
    $status = [string]$(if ($state.ContainsKey('status')) { $state.status } else { '' })
    $announcementErrors = [string]$(if ($state.ContainsKey('announcement_errors')) { $state.announcement_errors } else { '' })
    $lastErrorMessage = [string]$(if ($state.ContainsKey('last_error_message')) { $state.last_error_message } else { '' })
    $pidValue = [string]$(if ($state.ContainsKey('pid')) { $state.pid } else { '' })
    return ('{0}|{1}|{2}|{3}' -f $status, $announcementErrors, $lastErrorMessage, $pidValue)
}


function Reset-ConfigCriticalErrorTracking {
    $script:ConfigCriticalErrorCounts = @{}
}

function Register-ConfigCriticalError {
    param([string]$ConfigPath)

    $key = [string]$ConfigPath
    if ([string]::IsNullOrWhiteSpace($key)) {
        $key = '__empty__'
    }

    if (-not $script:ConfigCriticalErrorCounts.ContainsKey($key)) {
        $script:ConfigCriticalErrorCounts[$key] = 0
    }

    $script:ConfigCriticalErrorCounts[$key] = [int]$script:ConfigCriticalErrorCounts[$key] + 1
    return [int]$script:ConfigCriticalErrorCounts[$key]
}

function Get-ConfigBanInfo {
    $state = Get-State
    return [pscustomobject]@{
        path  = [string]$(if ($state.ContainsKey('banned_config_path')) { $state.banned_config_path } else { '' })
        stamp = [string]$(if ($state.ContainsKey('banned_config_stamp')) { $state.banned_config_stamp } else { '' })
    }
}

function Set-ConfigBanInfo {
    param(
        [string]$Path,
        [string]$Stamp
    )

    Update-State -Values @{
        banned_config_path  = [string]$Path
        banned_config_stamp = [string]$Stamp
    }
}

function Clear-ConfigBanInfo {
    Update-State -Values @{
        banned_config_path  = ''
        banned_config_stamp = ''
    }
}

function Test-ConfigBanned {
    param(
        [string]$Path,
        [string]$Stamp
    )

    $ban = Get-ConfigBanInfo
    if ([string]::IsNullOrWhiteSpace($ban.path)) {
        return $false
    }

    if ([string]$ban.path -ne [string]$Path) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$ban.stamp)) {
        return $true
    }

    return ([string]$ban.stamp -eq [string]$Stamp)
}


function Get-DefaultAppInfo {

    return @{
        current_version     = $DefaultAppVersion
        remote_config_url   = 'https://raw.githubusercontent.com/gulenokdv/Report---Fix-Happ-Routing/main/serv/happ-routing-remote-config.json'
        download_page_url   = 'https://github.com/gulenokdv/Report---Fix-Happ-Routing'
    }
}

function Save-AppInfo {
    param([hashtable]$AppInfo)
    Write-Utf8BomFile -Path $AppInfoPath -Content (ConvertTo-ReadableJson -InputObject ([pscustomobject]$AppInfo))
}

function Get-AppInfo {
    $appInfo = Get-DefaultAppInfo
    $needsSave = $false

    if (Test-Path -LiteralPath $AppInfoPath) {
        try {
            $loaded = Read-Utf8JsonFile -Path $AppInfoPath
            if ($loaded.PSObject.Properties['current_version']) {
                $appInfo.current_version = [string]$loaded.current_version
            }
            if ($loaded.PSObject.Properties['remote_config_url']) {
                $appInfo.remote_config_url = [string]$loaded.remote_config_url
            }
            if ($loaded.PSObject.Properties['download_page_url']) {
                $appInfo.download_page_url = [string]$loaded.download_page_url
            }
        }
        catch {
            $needsSave = $true
        }
    }
    else {
        $needsSave = $true
    }

    if ([string]::IsNullOrWhiteSpace([string]$appInfo.current_version)) {
        $appInfo.current_version = $DefaultAppVersion
        $needsSave = $true
    }

    if ($needsSave) {
        Save-AppInfo -AppInfo $appInfo
    }

    return [pscustomobject]$appInfo
}

function Get-AppVersion {
    return [string](Get-AppInfo).current_version
}

function Get-RemoteConfigUrl {
    return [string](Get-AppInfo).remote_config_url
}

function Get-DefaultDownloadPageUrl {
    return [string](Get-AppInfo).download_page_url
}

function Get-Settings {
    $settings = Get-DefaultSettings
    $needsSave = $false

    if (Test-Path -LiteralPath $SettingsPath) {
        try {
            $loaded = Read-Utf8JsonFile -Path $SettingsPath
            foreach ($name in @('auto_update_enabled', 'auto_update_interval_min', 'autostart_enabled', 'keep_background_on_close', 'custom_config_path')) {
                if ($loaded.PSObject.Properties[$name]) {
                    $settings[$name] = $loaded.$name
                }
            }

            if (-not $loaded.PSObject.Properties['keep_background_on_close'] -and $loaded.PSObject.Properties['hide_console_on_fix']) {
                $settings.keep_background_on_close = [bool]$loaded.hide_console_on_fix
                $needsSave = $true
            }
        }
        catch {
            $needsSave = $true
        }
    }
    else {
        $needsSave = $true
    }
    
    if ([int]$settings.auto_update_interval_min -lt 5) {
        $settings.auto_update_interval_min = 5
        $needsSave = $true
    }

    if ($needsSave) {
        Save-Settings -Settings $settings
    }
    return [pscustomobject]$settings
}

function Get-State {
    if (-not (Test-Path -LiteralPath $StatePath)) {
        return @{}
    }

    try {
        $loaded = Read-Utf8JsonFile -Path $StatePath
        $state = @{}
        foreach ($property in $loaded.PSObject.Properties) {
            $state[$property.Name] = $property.Value
        }
        return $state
    }
    catch {
        return @{}
    }
}

function Save-State {
    param([hashtable]$State)
    Write-Utf8BomFile -Path $StatePath -Content (ConvertTo-ReadableJson -InputObject ([pscustomobject]$State))
}

function Get-DefaultRemoteConfig {
    $currentVersion = Get-AppVersion

    return @{
        type            = 'none'
        latest_version  = $currentVersion
        message         = ''
        message_url     = ''
        disable_fix     = $false
    }
}

function Get-RemoteConfig {
    $remoteConfig = Get-DefaultRemoteConfig

    if (Test-Path -LiteralPath $RemoteConfigPath) {
        try {
            $loaded = Read-Utf8JsonFile -Path $RemoteConfigPath
            $remoteConfig = ConvertTo-RemoteConfigHashtable -RemoteConfig $loaded
        }
        catch {
            # fallback to defaults
        }
    }

    if ($null -eq $remoteConfig['message']) {
        $remoteConfig['message'] = ''
    }

    return [pscustomobject]$remoteConfig
}


function ConvertTo-VersionObject {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    try {
        return [version]$Value
    }
    catch {
        return $null
    }
}

function Compare-AppVersions {
    param(
        [string]$LeftVersion,
        [string]$RightVersion
    )

    $left = ConvertTo-VersionObject -Value $LeftVersion
    $right = ConvertTo-VersionObject -Value $RightVersion
    if ($null -eq $left -or $null -eq $right) {
        return 0
    }

    return $left.CompareTo($right)
}

function ConvertTo-RemoteConfigHashtable {
    param([object]$RemoteConfig)

    $defaults = Get-DefaultRemoteConfig
    $normalized = [ordered]@{}

    if ($null -ne $RemoteConfig) {
        if ($RemoteConfig -is [System.Collections.IDictionary]) {
            foreach ($key in $RemoteConfig.Keys) {
                $normalized[[string]$key] = $RemoteConfig[$key]
            }
        }
        else {
            foreach ($property in $RemoteConfig.PSObject.Properties) {
                $normalized[$property.Name] = $property.Value
            }
        }
    }

    if (-not $normalized.Contains('type')) {
        $normalized['type'] = [string]$defaults.type
    }
    if (-not $normalized.Contains('latest_version')) {
        $normalized['latest_version'] = [string]$defaults.latest_version
    }
    if (-not $normalized.Contains('message')) {
        $normalized['message'] = [string]$defaults.message
    }
    if (-not $normalized.Contains('disable_fix')) {
        $normalized['disable_fix'] = [bool]$defaults.disable_fix
    }
    if (-not $normalized.Contains('message_url')) {
        $normalized['message_url'] = [string]$defaults.message_url
    }

    $normalized['type'] = [string]$normalized['type']
    $normalized['latest_version'] = [string]$normalized['latest_version']
    $normalized['message'] = [string]$normalized['message']
    $normalized['message_url'] = [string]$normalized['message_url']
    $normalized['disable_fix'] = [bool]$normalized['disable_fix']

    if ([string]::IsNullOrWhiteSpace($normalized['latest_version'])) {
        $normalized['latest_version'] = Get-AppVersion
    }

    if ($null -eq $normalized['message']) {
        $normalized['message'] = ''
    }

    return $normalized
}

function Get-RemoteConfigDownloadUrl {
    param([object]$RemoteConfig)

    if ($null -eq $RemoteConfig) {
        return Get-DefaultDownloadPageUrl
    }

    if ($RemoteConfig -is [System.Collections.IDictionary]) {
        if ($RemoteConfig.Contains('download_url') -and -not [string]::IsNullOrWhiteSpace([string]$RemoteConfig['download_url'])) {
            return [string]$RemoteConfig['download_url']
        }
    }
    elseif ($RemoteConfig.PSObject.Properties['download_url'] -and -not [string]::IsNullOrWhiteSpace([string]$RemoteConfig.download_url)) {
        return [string]$RemoteConfig.download_url
    }

    return Get-DefaultDownloadPageUrl
}

function Test-UiSessionActive {
    $state = Get-State
    if (-not $state.ContainsKey('ui_session_active') -or [string]$state.ui_session_active -ne 'true') {
        return $false
    }

    if (-not $state.ContainsKey('ui_pid') -or [string]::IsNullOrWhiteSpace([string]$state.ui_pid)) {
        return $false
    }

    try {
        $pidValue = [int]$state.ui_pid
        $process = Get-Process -Id $pidValue -ErrorAction Stop
        return ($null -ne $process)
    }
    catch {
        Update-State -Values @{ ui_session_active = 'false'; ui_pid = '' }
        return $false
    }
}

function Test-UpdatePromptDismissed {
    param([object]$RemoteConfig)

    $config = ConvertTo-RemoteConfigHashtable -RemoteConfig $RemoteConfig
    $state = Get-State
    $currentKey = ('{0}|{1}|{2}' -f [string]$config['latest_version'], [string]$config['message'], [string]$config['type'])
    $dismissedKey = [string]$(if ($state.ContainsKey('dismissed_update_prompt_key')) { $state.dismissed_update_prompt_key } else { '' })
    return ($currentKey -eq $dismissedKey)
}

function Dismiss-UpdatePrompt {
    param([object]$RemoteConfig)

    $config = ConvertTo-RemoteConfigHashtable -RemoteConfig $RemoteConfig
    Update-State -Values @{ dismissed_update_prompt_key = ('{0}|{1}|{2}' -f [string]$config['latest_version'], [string]$config['message'], [string]$config['type']) }
}

function Test-RemoteConfigChanged {
    param(
        [object]$Left,
        [object]$Right
    )

    $leftConfig = ConvertTo-RemoteConfigHashtable -RemoteConfig $Left
    $rightConfig = ConvertTo-RemoteConfigHashtable -RemoteConfig $Right

    foreach ($name in @('type', 'latest_version', 'message', 'message_url', 'disable_fix', 'download_url')) {
        if ([string]$leftConfig[$name] -ne [string]$rightConfig[$name]) {
            return $true
        }
    }

    return $false
}

function Update-State {
    param([hashtable]$Values)

    $state = Get-State
    $changed = $false

    foreach ($key in $Values.Keys) {
        $newValue = [string]$Values[$key]
        if (-not $state.ContainsKey($key) -or [string]$state[$key] -ne $newValue) {
            $state[$key] = $newValue
            $changed = $true
        }
    }

    if ($changed) {
        Save-State -State $state
    }
}

function Remove-LegacyAutostartArtifacts {
    $startupDir = [Environment]::GetFolderPath('Startup')
    $legacyKnown = @(
        'Happ Routing Guard.lnk',
        'start-happ-routing-guard.bat',
        'happ-routing-fix.bat'
    )

    foreach ($name in $legacyKnown) {
        $path = Join-Path -Path $startupDir -ChildPath $name
        if (Test-Path -LiteralPath $path) {
            try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop } catch {}
        }
    }

    $startupScripts = Get-ChildItem -LiteralPath $startupDir -File -ErrorAction SilentlyContinue | Where-Object {
        $_.Extension -in @('.bat', '.cmd')
    }
    foreach ($script in @($startupScripts)) {
        try {
            $text = Get-Content -LiteralPath $script.FullName -Raw -ErrorAction Stop
            if ($text -match [regex]::Escape($ScriptSignature)) {
                Remove-Item -LiteralPath $script.FullName -Force -ErrorAction Stop
            }
        }
        catch {
        }
    }

    foreach ($valueName in @($LegacyAutostartRunValueNames)) {
        try {
            $rawValue = Get-RunRegistryValue -Name $valueName
            if ($null -eq $rawValue) { continue }

            $command = [string]$rawValue

            $isManaged = $false
            if (-not [string]::IsNullOrWhiteSpace($command)) {
                $exePath = Get-CommandExecutablePath -CommandLine $command
                $isManaged = Test-ManagedLauncherFile -Path $exePath

                if (-not $isManaged) {
                    $normalized = $command.ToLowerInvariant()
                    foreach ($marker in @($ManagedAutostartCommandMarkers)) {
                        if ($normalized.Contains($marker)) {
                            $isManaged = $true
                            break
                        }
                    }
                }
            }
            else {
                $isManaged = $true
            }

            if ($isManaged -and $valueName -ne $AutostartRunValueName) {
                Remove-RunRegistryValue -Name $valueName
            }
        }
        catch {
        }
    }
}

function Set-AutostartEntry {
    param([bool]$Enabled)

    try {
        Remove-LegacyAutostartArtifacts

        if ($Enabled) {
            $launcherPath = Get-LauncherPath
            $command = ('"{0}" /background' -f $launcherPath)
            $current = [string](Get-RunRegistryValue -Name $AutostartRunValueName)
            if ([string]$current -ne [string]$command) {
                Set-RunRegistryValue -Name $AutostartRunValueName -Value $command
            }
        }
        else {
            Remove-RunRegistryValue -Name $AutostartRunValueName
        }

        Update-State -Values @{ autostart_last_error = '' }
        return $true
    }
    catch {
        Update-State -Values @{ autostart_last_error = $_.Exception.Message }
        Write-AppLogLimited -Key ('autostart:' + $_.Exception.Message) -Message ('Не удалось обновить автозапуск: ' + $_.Exception.Message)
        return $false
    }
}

function Toggle-Setting {
    param([string]$Name)

    $data = Get-DefaultSettings
    $current = Get-Settings
    foreach ($key in @($data.Keys)) {
        $data[$key] = $current.$key
    }
    $data[$Name] = -not [bool]$current.$Name
    Save-Settings -Settings $data
    return [pscustomobject]$data
}

function Get-DesiredRunning {
    $state = Get-State
    if (-not $state.ContainsKey('desired_running')) {
        return $true
    }

    return ([string]$state.desired_running -ne 'false')
}

function Get-AgentProcess {
    $scriptPaths = @(
        (Join-Path -Path $ScriptRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant(),
        (Join-Path -Path $AppRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant()
    )
    $processes = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" -ErrorAction SilentlyContinue
    foreach ($process in @($processes)) {
        if ($null -eq $process.CommandLine) { continue }
        $commandLine = $process.CommandLine.ToLowerInvariant()
        foreach ($scriptPath in $scriptPaths) {
            if ($commandLine.Contains($scriptPath) -and $commandLine.Contains('-action agent')) {
                return $process
            }
        }
    }
    return $null
}

function Test-LegacyAgentProcess {
    param([object]$Process)

    if ($null -eq $Process -or $null -eq $Process.CommandLine) {
        return $false
    }

    $currentScriptPath = (Join-Path -Path $ScriptRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant()
    $legacyScriptPath = (Join-Path -Path $AppRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant()
    $commandLine = $Process.CommandLine.ToLowerInvariant()
    return ($commandLine.Contains($legacyScriptPath) -and -not $commandLine.Contains($currentScriptPath))
}

function Start-AgentProcess {
    param([bool]$SetDesired = $true)

    if ($SetDesired) {
        Update-State -Values @{ desired_running = 'true' }
    }

    $existing = Get-AgentProcess
    if ($null -ne $existing) {
        if (Test-LegacyAgentProcess -Process $existing) {
            Stop-Process -Id $existing.ProcessId -Force
            Start-Sleep -Milliseconds 300
        }
        else {
            return $existing.ProcessId
        }
    }

    $scriptPath = Join-Path -Path $ScriptRoot -ChildPath 'happ-routing.ps1'
    $argumentLine = ('-NoLogo -NoProfile -ExecutionPolicy Bypass -File "{0}" -Action agent' -f $scriptPath)
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $argumentLine -WindowStyle Hidden -PassThru

    Start-Sleep -Milliseconds 300
    return $process.Id
}

function Stop-AgentProcess {
    param([bool]$SetDesired = $true)

    if ($SetDesired) {
        Update-State -Values @{ desired_running = 'false' }
    }

    $existing = Get-AgentProcess
    if ($null -eq $existing) {
        Update-State -Values @{ pid = ''; status = 'остановлено' }
        return $false
    }

    Stop-Process -Id $existing.ProcessId -Force
    Update-State -Values @{ pid = ''; status = 'остановлено' }
    return $true
}

function Get-RelatedFixProcesses {
    $currentScriptPath = (Join-Path -Path $ScriptRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant()
    $legacyScriptPath = (Join-Path -Path $AppRoot -ChildPath 'happ-routing.ps1').ToLowerInvariant()
    $launcherPath = (Join-Path -Path $AppRoot -ChildPath 'happ-routing-fix.bat').ToLowerInvariant()
    $monitorLauncherPath = (Join-Path -Path $AppRoot -ChildPath 'happ-routing-process-check.bat').ToLowerInvariant()
    $markers = @(
        $currentScriptPath,
        $legacyScriptPath,
        $launcherPath,
        $ScriptSignature.ToLowerInvariant()
    )

    $processes = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe' OR Name = 'cmd.exe'" -ErrorAction Stop
    $related = @()

    foreach ($process in @($processes)) {
        if ([int]$process.ProcessId -eq [int]$PID) { continue }
        if ($null -eq $process.CommandLine) { continue }

        $commandLine = $process.CommandLine.ToLowerInvariant()
        if ($commandLine.Contains($monitorLauncherPath)) { continue }

        $isRelated = $false
        foreach ($marker in $markers) {
            if (-not [string]::IsNullOrWhiteSpace($marker) -and $commandLine.Contains($marker)) {
                $isRelated = $true
                break
            }
        }

        if ($isRelated) {
            $related += $process
        }
    }

    return @($related | Sort-Object ProcessId)
}

function Get-RelatedProcessScriptPath {
    param([object]$Process)

    if ($null -eq $Process -or [string]::IsNullOrWhiteSpace([string]$Process.CommandLine)) {
        return ''
    }

    $commandLine = [string]$Process.CommandLine
    $patterns = @(
        '(?i)-File\s+"([^"]+)"',
        '(?i)-File\s+([^\s]+)',
        '(?i)"([^"]*happ-routing-fix\.bat)"',
        '(?i)([^\s"]*happ-routing-fix\.bat)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($commandLine, $pattern)
        if ($match.Success) {
            return $match.Groups[1].Value
        }
    }

    return $commandLine
}

function Show-ProcessMonitorScreen {
    param(
        [array]$Processes,
        [string]$ErrorMessage = '',
        [int]$KilledCount = -1
    )

    Clear-Host
    Write-Host 'Happ Routing Fix - проверка процессов' -ForegroundColor Gray
    Write-Host ''

    if (-not [string]::IsNullOrWhiteSpace($ErrorMessage)) {
        Write-Host 'СТАТУС: НЕ УДАЛОСЬ ПРОЧИТАТЬ ПРОЦЕССЫ' -ForegroundColor DarkRed
        Write-Host $ErrorMessage -ForegroundColor DarkGray
    }
    elseif ($Processes.Count -gt 0) {
        Write-Host ('СТАТУС: НАЙДЕНО ПРОЦЕССОВ: {0}' -f $Processes.Count) -ForegroundColor DarkYellow
    }
    else {
        Write-Host 'СТАТУС: ВСЕ ПРОЦЕССЫ ФИКСА МЕРТВЫ' -ForegroundColor Green
    }

    if ($KilledCount -ge 0) {
        Write-Host ('Последнее действие: убито процессов: {0}' -f $KilledCount) -ForegroundColor DarkCyan
    }

    Write-Host ''
    if ($Processes.Count -gt 0) {
        foreach ($process in @($Processes)) {
            Write-Host ('PID {0}  {1}' -f $process.ProcessId, $process.Name) -ForegroundColor Cyan
            Write-Host ('  {0}' -f (Get-RelatedProcessScriptPath -Process $process)) -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host 'Запущенных процессов от happ-routing-fix.bat не видно.' -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host '[1] Обновить список' -ForegroundColor Cyan
    Write-Host '[2] Убить все найденные процессы' -ForegroundColor DarkYellow
    Write-Host '[4] Запустить консольку фикса' -ForegroundColor Gray
    Write-Host '[5] Открыть Telegram разработчика' -ForegroundColor DarkGray
    Write-Host '[8] Открыть репозиторий на GitHub' -ForegroundColor DarkGray
    Write-Host '[Esc] Выйти' -ForegroundColor DarkGray
}

function Start-ProcessMonitor {
    Initialize-Console
    $Host.UI.RawUI.WindowTitle = 'Happ Routing Fix - process check'
    $lastKilledCount = -1

    while ($true) {
        $errorMessage = ''
        $processes = @()
        try {
            $processes = @(Get-RelatedFixProcesses)
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        Show-ProcessMonitorScreen -Processes $processes -ErrorMessage $errorMessage -KilledCount $lastKilledCount
        $lastKilledCount = -1
        $key = Read-ProcessMonitorKey

        if ($key.VirtualKeyCode -eq 27) {
            break
        }

        if ($key.VirtualKeyCode -eq 52 -or $key.VirtualKeyCode -eq 100) {
            Start-LauncherWindow
            continue
        }

        if ($key.VirtualKeyCode -eq 53 -or $key.VirtualKeyCode -eq 101) {
            Open-DeveloperTelegram
            continue
        }

        if ($key.VirtualKeyCode -eq 56 -or $key.VirtualKeyCode -eq 104) {
            try {
                Start-Process -FilePath 'https://github.com/gulenokdv/Report---Fix-Happ-Routing' | Out-Null
            }
            catch {
                Write-AppLogLimited -Key 'github-open' -Message 'Не удалось открыть репозиторий на GitHub.'
            }
            continue
        }

        if ($key.VirtualKeyCode -eq 50 -or $key.VirtualKeyCode -eq 98) {
            $killed = 0
            foreach ($process in @($processes)) {
                try {
                    Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
                    $killed++
                }
                catch {
                }
            }

            Update-State -Values @{ desired_running = 'false'; pid = ''; status = 'остановлено' }
            $lastKilledCount = $killed
        }
    }
}

function Invoke-UiShutdown {
    param([bool]$ShowBackgroundNotice = $true)

    if ($script:UiShutdownHandled) {
        return
    }

    $script:UiShutdownHandled = $true
    Update-State -Values @{ ui_session_active = 'false'; ui_pid = '' }
    $settings = Get-Settings

    if ([bool]$settings.keep_background_on_close) {
        $agent = Get-AgentProcess
        if ($null -ne $agent) {
            $state = Get-State
            $noticeShown = $false
            if ($state.ContainsKey('background_notice_shown') -and [string]$state.background_notice_shown -eq 'true') {
                $noticeShown = $true
            }

            if ($ShowBackgroundNotice -and -not $noticeShown) {
                try {
                    Show-WindowsNotification -Title 'Happ Routing Fix' -Message 'Фикс продолжает работать в фоне. Повторный запуск батника раскроет существующую консоль.'
                    Update-State -Values @{ background_notice_shown = 'true' }
                    Write-AppLog 'Показано уведомление о работе фикса в фоне.'
                }
                catch {
                    Write-AppLogLimited -Key ('background-notice-inline:' + $_.Exception.Message) -Message ('Не удалось показать встроенное фоновое уведомление: ' + $_.Exception.Message)
                }
            }
        }
        return
    }

    try {
        $null = Stop-AgentProcess -SetDesired $false
    }
    catch {
    }
}

function Start-UiCloseWatcher {
    $scriptPath = Join-Path -Path $ScriptRoot -ChildPath 'happ-routing.ps1'
    $argumentLine = ('-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "{0}" -Action watch-ui-close -UiPid {1}' -f $scriptPath, $PID)
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argumentLine -WindowStyle Hidden | Out-Null
}

function Wait-UiCloseAndNotify {
    param([int]$WatchedUiPid)

    if ($WatchedUiPid -le 0) {
        return
    }

    while ($true) {
        try {
            $null = Get-Process -Id $WatchedUiPid -ErrorAction Stop
            Start-Sleep -Milliseconds 250
            continue
        }
        catch {
            break
        }
    }

    Start-Sleep -Milliseconds 500

    $settings = Get-Settings
    if (-not [bool]$settings.keep_background_on_close) {
        try {
            $null = Stop-AgentProcess -SetDesired $false
            Write-AppLog 'Консоль закрыта при F2=выкл, агент фикса остановлен watcher-процессом.'
        }
        catch {
            Write-AppLogLimited -Key ('watcher-stop:' + $_.Exception.Message) -Message ('Watcher не смог остановить агент после закрытия консоли: ' + $_.Exception.Message)
        }
        return
    }

    $agent = Get-AgentProcess
    if ($null -eq $agent) {
        return
    }

    $state = Get-State
    if ($state.ContainsKey('background_notice_shown') -and [string]$state.background_notice_shown -eq 'true') {
        return
    }

    try {
        Show-WindowsNotification -Title 'Happ Routing Fix' -Message 'Фикс продолжает работать в фоне. Повторный запуск батника раскроет существующую консоль.'
        Update-State -Values @{ background_notice_shown = 'true' }
        Write-AppLog 'Показано уведомление о работе фикса в фоне.'
    }
    catch {
        Write-AppLogLimited -Key ('background-notice:' + $_.Exception.Message) -Message ('Не удалось показать фоновое уведомление: ' + $_.Exception.Message)
    }
}

function Invoke-RulesUpdate {
    $oldProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    try {
        New-Item -ItemType Directory -Force -Path $RulesDir | Out-Null

        $downloads = @(
            @{ Name = 'geosite-vk.srs'; Url = 'https://raw.githubusercontent.com/SagerNet/sing-geosite/rule-set/geosite-vk.srs' },
            @{ Name = 'geosite-category-ru.srs'; Url = 'https://raw.githubusercontent.com/SagerNet/sing-geosite/rule-set/geosite-category-ru.srs' },
            @{ Name = 'geoip-ru.srs'; Url = 'https://raw.githubusercontent.com/SagerNet/sing-geoip/rule-set/geoip-ru.srs' }
        )

        $state = Get-State
        $downloadedCount = 0
        $unchangedCount = 0
        $stateChanged = $false

        Write-AppLog 'Начало проверки обновления rulesets...'

        foreach ($download in $downloads) {
            $fileKey = 'ruleset_etag_' + $download.Name
            $currentEtag = ''
            if ($state.ContainsKey($fileKey)) {
                $currentEtag = [string]$state[$fileKey]
            }

            $targetPath = Join-Path -Path $RulesDir -ChildPath $download.Name
            $localFileExists = (Test-Path -LiteralPath $targetPath -PathType Leaf)
            if (-not $localFileExists) {
                Write-AppLog ('Локальный ruleset отсутствует, будет скачан заново: {0}' -f $download.Name)
            }

            $serverEtag = $null
            $headSuccess = $false

            # Пробуем получить ETag через HEAD запрос
            try {
                $headResponse = Invoke-WebRequest -Uri $download.Url -Method Head -UseBasicParsing -TimeoutSec 30
                if ($headResponse.Headers) {
                    $headerDict = $headResponse.Headers
                    if ($headerDict.ContainsKey('ETag')) {
                        $serverEtag = [string]$headerDict['ETag']
                    }
                    elseif ($headerDict.ContainsKey('etag')) {
                        $serverEtag = [string]$headerDict['etag']
                    }
                    elseif ($headerDict['ETag']) {
                        $serverEtag = [string]$headerDict['ETag']
                    }
                }
                $headSuccess = $true
                Write-AppLog ('HEAD {0}: ETag={1}' -f $download.Name, $(if ($serverEtag) { $serverEtag.Substring(0, [Math]::Min(20, $serverEtag.Length)) + '...' } else { 'none' }))
            }
            catch {
                Write-AppLog ('HEAD запрос не удался для {0}: {1}' -f $download.Name, $_.Exception.Message)
            }

            # Если ETag совпадает и локальный файл уже существует — пропускаем скачивание
            if ($localFileExists -and $headSuccess -and $serverEtag -and $currentEtag -eq $serverEtag) {
                $unchangedCount++
                continue
            }


            # Скачиваем файл
            $tempPath = [System.IO.Path]::GetTempFileName()
            try {
                Write-AppLog ('Скачивание {0}...' -f $download.Name)
                Invoke-WebRequest -Uri $download.Url -OutFile $tempPath -UseBasicParsing -TimeoutSec 60 | Out-Null
                Copy-Item -LiteralPath $tempPath -Destination (Join-Path -Path $RulesDir -ChildPath $download.Name) -Force

                # Сохраняем ETag (если получили с HEAD, иначе вычисляем хеш файла)
                if ($serverEtag) {
                    $state[$fileKey] = $serverEtag
                    $stateChanged = $true
                }
                else {
                    try {
                        $fileHash = (Get-FileHash -LiteralPath $tempPath -Algorithm SHA256).Hash
                        $state[$fileKey] = $fileHash
                        $stateChanged = $true
                        Write-AppLog ('Сохранен хеш для {0}: {1}...' -f $download.Name, $fileHash.Substring(0, 16))
                    }
                    catch {
                        Write-AppLog ('Не удалось вычислить хеш для {0}' -f $download.Name)
                    }
                }
                $downloadedCount++
            }
            catch {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                throw ('Ошибка скачивания {0}: {1}' -f $download.Name, $_.Exception.Message)
            }
            finally {
                if (Test-Path -LiteralPath $tempPath) {
                    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                }
            }
        }

        # Сохраняем обновлённое state если были изменения ETag
        if ($stateChanged) {
            Save-State -State $state
            Write-AppLog 'ETAG rulesets сохранены в state.'
        }

        Update-State -Values @{
            last_rules_update_at = (Get-Date).ToString('s')
            next_rules_retry_at  = ''
        }

        if ($downloadedCount -gt 0) {
            Write-AppLog ('Rulesets успешно обновлены: скачано {0} из {1} файлов.' -f $downloadedCount, $downloads.Count)
        }
        else {
            Write-AppLog 'Rulesets не требуют обновления (файлы актуальны).'
        }
    }
    finally {
        $ProgressPreference = $oldProgress
    }
}

function Get-ConfigStamp {
    $effectiveConfigPath = Get-EffectiveConfigPath
    if (-not (Test-Path -LiteralPath $effectiveConfigPath)) { return $null }
    $item = Get-Item -LiteralPath $effectiveConfigPath
    return ('{0}|{1}' -f $item.LastWriteTimeUtc.Ticks, $item.Length)
}

function Test-ConfigNeedsPatch {
    $effectiveConfigPath = Get-EffectiveConfigPath
    if (-not (Test-Path -LiteralPath $effectiveConfigPath)) { return $false }

    try {
        $config = Read-Utf8JsonFile -Path $effectiveConfigPath
    }
    catch {
        return $true
    }

    if ($null -eq $config.route -or -not $config.route.PSObject.Properties['rule_set']) { return $true }

    $routeRuleSet = @($config.route.rule_set)
    if ($routeRuleSet.Count -lt 3) { return $true }

    $tags = @($routeRuleSet | ForEach-Object { $_.tag })
    $types = @($routeRuleSet | ForEach-Object { $_.type })
    $expectedRuleSetPaths = @{
        'geosite-vk' = [System.IO.Path]::GetFullPath((Join-Path -Path $RulesDir -ChildPath 'geosite-vk.srs'))
        'geosite-category-ru' = [System.IO.Path]::GetFullPath((Join-Path -Path $RulesDir -ChildPath 'geosite-category-ru.srs'))
        'geoip-ru' = [System.IO.Path]::GetFullPath((Join-Path -Path $RulesDir -ChildPath 'geoip-ru.srs'))
    }

    if (-not ($tags -contains 'geosite-vk' -and $tags -contains 'geosite-category-ru' -and $tags -contains 'geoip-ru')) { return $true }
    if (($types | Where-Object { $_ -eq 'local' }).Count -lt 3) { return $true }

    foreach ($ruleSet in $routeRuleSet) {
        if (-not $ruleSet.PSObject.Properties['tag']) { continue }

        $tag = [string]$ruleSet.tag
        if (-not $expectedRuleSetPaths.ContainsKey($tag)) { continue }
        if (-not $ruleSet.PSObject.Properties['path']) { return $true }

        $actualPath = [string]$ruleSet.path
        if ([string]::IsNullOrWhiteSpace($actualPath)) { return $true }

        try {
            $normalizedActualPath = [System.IO.Path]::GetFullPath($actualPath)
        }
        catch {
            return $true
        }

        if ($normalizedActualPath -ne $expectedRuleSetPaths[$tag]) {
            return $true
        }
    }

    foreach ($rule in @($config.route.rules)) {
        if ($rule.PSObject.Properties['rule_set'] -and $rule.outbound -eq 'direct') {
            $sets = @($rule.rule_set)
            if ($sets -contains 'geosite-vk' -or $sets -contains 'geosite-category-ru' -or $sets -contains 'geoip-ru') {
                return $false
            }
        }
    }

    return $true
}

function Invoke-ConfigPatch {
    $effectiveConfigPath = Get-EffectiveConfigPath
    if (-not (Test-Path -LiteralPath $effectiveConfigPath)) { return $false }

    & $PatchScript -ConfigPath $effectiveConfigPath | Out-Null
    Update-State -Values @{
        last_patch_at = (Get-Date).ToString('s')
        status        = 'пропатчено'
    }
    Write-AppLog 'Конфиг маршрутизации пропатчен.'
    return $true
}

function Invoke-BurstPatch {
    $deadline = (Get-Date).AddMilliseconds(2500)
    $attempts = 0
    $effectiveConfigPath = Get-EffectiveConfigPath

    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $effectiveConfigPath) {
            try {
                if (Invoke-ConfigPatch) {
                    $attempts++
                }
            }
            catch {
            }
        }
        Start-Sleep -Milliseconds 15
    }

    return $attempts
}

function Test-RulesNeedUpdate {
    $settings = Get-Settings
    if (-not [bool]$settings.auto_update_enabled) { return $false }

    $state = Get-State
    if ($state.ContainsKey('next_rules_retry_at') -and -not [string]::IsNullOrWhiteSpace([string]$state.next_rules_retry_at)) {
        try {
            $retryAt = [datetime]::Parse([string]$state.next_rules_retry_at)
            if ((Get-Date) -lt $retryAt) { return $false }
        }
        catch {
        }
    }

    if (-not $state.ContainsKey('last_rules_update_at') -or [string]::IsNullOrWhiteSpace([string]$state.last_rules_update_at)) {
        return $true
    }

    try {
        $lastUpdate = [datetime]::Parse([string]$state.last_rules_update_at)
    }
    catch {
        return $true
    }

    return (((Get-Date) - $lastUpdate).TotalMinutes -ge [int]$settings.auto_update_interval_min)
}

function Test-RemoteConfigNeedsSync {
    $settings = Get-Settings
    if (-not [bool]$settings.auto_update_enabled) { return $false }

    $url = Get-RemoteConfigUrl
    if ([string]::IsNullOrWhiteSpace($url)) { return $false }

    $state = Get-State
    $hasStartupCheck = ($state.ContainsKey('remote_config_boot_check_pending') -and [string]$state.remote_config_boot_check_pending -eq 'true')
    $hasUiSession = Test-UiSessionActive
    if (-not $hasStartupCheck -and -not $hasUiSession) {
        return $false
    }

    if ($hasStartupCheck) {
        return $true
    }

    if ($state.ContainsKey('next_remote_config_retry_at') -and -not [string]::IsNullOrWhiteSpace([string]$state.next_remote_config_retry_at)) {
        try {
            $retryAt = [datetime]::Parse([string]$state.next_remote_config_retry_at)
            if ((Get-Date) -lt $retryAt) { return $false }
        }
        catch {
        }
    }

    if (-not $state.ContainsKey('last_remote_config_sync_at') -or [string]::IsNullOrWhiteSpace([string]$state.last_remote_config_sync_at)) {
        return $true
    }

    try {
        $lastSyncAt = [datetime]::Parse([string]$state.last_remote_config_sync_at)
    }
    catch {
        return $true
    }

    return (((Get-Date) - $lastSyncAt).TotalMinutes -ge [int]$settings.auto_update_interval_min)
}

function Invoke-RemoteConfigSync {
    $url = Get-RemoteConfigUrl
    if ([string]::IsNullOrWhiteSpace($url)) {
        throw 'Не задан remote_config_url в serv/happ-routing-app.json.'
    }

    $tempPath = [System.IO.Path]::GetTempFileName()
    try {
        Invoke-WebRequest -Uri $url -OutFile $tempPath

        $serverConfig = Read-Utf8JsonFile -Path $tempPath
        $serverConfigData = ConvertTo-RemoteConfigHashtable -RemoteConfig $serverConfig
        $localConfig = Get-RemoteConfig
        $changed = Test-RemoteConfigChanged -Left $localConfig -Right $serverConfigData

        if ($changed) {
            # Перезаписываем локальный remote-config файл
            Write-Utf8BomFile -Path $RemoteConfigPath -Content (ConvertTo-ReadableJson -InputObject ([pscustomobject]$serverConfigData))
            Update-State -Values @{
                last_remote_config_sync_at = (Get-Date).ToString('s')
                last_remote_config_change_at = (Get-Date).ToString('s')
                next_remote_config_retry_at = ''
                remote_config_boot_check_pending = 'false'
            }
            Request-UiFlash
            Write-AppLog ('Remote config обновлен: type={0}, latest_version={1}' -f $serverConfigData.type, $serverConfigData.latest_version)
        }
        else {
            Update-State -Values @{
                last_remote_config_sync_at = (Get-Date).ToString('s')
                next_remote_config_retry_at = ''
                remote_config_boot_check_pending = 'false'
            }
        }

        return $changed
    }
    finally {
        if (Test-Path -LiteralPath $tempPath) {
            try {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop
            }
            catch {
            }
        }
    }
}

function Invoke-RemoteConfigSyncManual {
    try {
        $changed = Invoke-RemoteConfigSync
        if ($changed) {
            Update-State -Values @{ status = 'обновления синхронизированы' }
        }
        else {
            Update-State -Values @{ status = 'обновления не найдены' }
        }
    }
    catch {
        Update-State -Values @{
            status             = 'ошибка проверки обновлений'
            last_error_message = $_.Exception.Message
        }
        Write-AppLogLimited -Key ('remote-config-manual:' + $_.Exception.Message) -Message ('Ручная проверка обновлений завершилась ошибкой: ' + $_.Exception.Message)
    }
}

function Format-RuDate {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return 'еще не было'
    }

    try {
        $culture = New-Object System.Globalization.CultureInfo('ru-RU')
        return ([datetime]::Parse($Value)).ToString('HH:mm:ss, d MMMM yyyy', $culture)
    }
    catch {
        return $Value
    }
}

function Get-StatusObject {
    $settings = Get-Settings
    $state = Get-State
    $process = Get-AgentProcess

    $running = $false
    $pidValue = ''
    if ($null -ne $process) {
        $running = $true
        $pidValue = [string]$process.ProcessId
    }

    return [pscustomobject]@{
        running              = $running
        pid                  = $pidValue
        desired_running      = (Get-DesiredRunning)
        status               = [string]$(if ($state.ContainsKey('status')) { $state.status } else { '' })
        last_patch_at        = [string]$(if ($state.ContainsKey('last_patch_at')) { $state.last_patch_at } else { '' })
        last_rules_update_at = [string]$(if ($state.ContainsKey('last_rules_update_at')) { $state.last_rules_update_at } else { '' })
        autostart_error      = [string]$(if ($state.ContainsKey('autostart_last_error')) { $state.autostart_last_error } else { '' })
        autostart_enabled    = [bool]$settings.autostart_enabled
        keep_background_on_close = [bool]$settings.keep_background_on_close
        auto_update_enabled  = [bool]$settings.auto_update_enabled
        auto_update_interval = [int]$settings.auto_update_interval_min
    }
}

function Initialize-Console {
    [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    [Console]::TreatControlCAsInput = $true
    $Host.UI.RawUI.WindowTitle = $UiWindowTitle

    if (-not ('HappRoutingFixConsoleApi' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
public class HappRoutingFixConsoleApi {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
'@
    }
}

$script:UiToggleLine = -1
$script:ShowAnnouncementPrompt = $false

function Activate-UiWindow {
    $activated = $false
    try {
        if (-not ('HappRoutingFixWindowApi' -as [type])) {
            Add-Type @'
using System;
using System.Runtime.InteropServices;
public class HappRoutingFixWindowApi {
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
'@
        }

        $window = Get-Process -ErrorAction SilentlyContinue | Where-Object {
            $_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -eq $UiWindowTitle
        } | Select-Object -First 1

        if ($null -ne $window) {
            $handle = [IntPtr]$window.MainWindowHandle
            [HappRoutingFixWindowApi]::ShowWindowAsync($handle, 9) | Out-Null
            Start-Sleep -Milliseconds 80
            [HappRoutingFixWindowApi]::SetForegroundWindow($handle) | Out-Null
            Start-Sleep -Milliseconds 80
            [HappRoutingFixWindowApi]::SetForegroundWindow($handle) | Out-Null
            $activated = $true
        }
    }
    catch {
    }

    try {
        $shell = New-Object -ComObject WScript.Shell
        if ($shell.AppActivate($UiWindowTitle)) {
            $activated = $true
        }
    }
    catch {
    }

    return $activated
}

function Write-TelegramLink {
    $esc = [char]27
    $label = 'fix by gulenok91 tg: @gulenok91'
    Write-Host ($esc + ']8;;https://t.me/gulenok91' + $esc + '\' + $label + $esc + ']8;;' + $esc + '\') -ForegroundColor DarkGray
}

function Start-LauncherWindow {
    $launcherPath = Join-Path -Path $AppRoot -ChildPath 'happ-routing-fix.bat'
    if (-not (Activate-UiWindow)) {
        Start-Process -FilePath $launcherPath | Out-Null
    }
}

function Restart-FixUi {
    $launcherPath = Join-Path -Path $AppRoot -ChildPath 'happ-routing-fix.bat'
    $null = Stop-AgentProcess -SetDesired $false
    Start-Sleep -Milliseconds 250
    Start-Process -FilePath $launcherPath | Out-Null
    $script:UiExitRequested = $true
}


function Start-ProcessCheckerWindow {
    $checkerPath = Join-Path -Path $AppRoot -ChildPath 'happ-routing-process-check.bat'
    Start-Process -FilePath $checkerPath | Out-Null
}

function Open-DeveloperTelegram {
    try {
        Start-Process -FilePath 'https://t.me/gulenok91' | Out-Null
        return
    }
    catch {
    }

    try {
        Start-Process -FilePath 't.me/gulenok91' | Out-Null
    }
    catch {
        Write-AppLogLimited -Key 'telegram-open' -Message 'Не удалось открыть Telegram разработчика.'
    }
}

function Read-UiKey {
    $key = [Console]::ReadKey($true)
    $state = ''
    if (($key.Modifiers -band [ConsoleModifiers]::Control) -ne 0) {
        $state += 'LeftCtrlPressed,RightCtrlPressed'
    }
    if (($key.Modifiers -band [ConsoleModifiers]::Shift) -ne 0) {
        if ($state.Length -gt 0) { $state += ',' }
        $state += 'ShiftPressed'
    }
    if (($key.Modifiers -band [ConsoleModifiers]::Alt) -ne 0) {
        if ($state.Length -gt 0) { $state += ',' }
        $state += 'LeftAltPressed,RightAltPressed'
    }

    return [pscustomobject]@{
        VirtualKeyCode  = [int]$key.Key
        Character       = $key.KeyChar
        ControlKeyState = $state
    }
}

function Test-IgnoredConsoleCopyKey {
    param([object]$KeyInfo)

    $key = [int]$KeyInfo.VirtualKeyCode
    if ($key -in @(16, 17, 18, 91, 92)) {
        return $true
    }

    $state = [string]$KeyInfo.ControlKeyState
    $ctrlPressed = ($state.Contains('LeftCtrlPressed') -or $state.Contains('RightCtrlPressed'))
    if ($ctrlPressed -and ($key -eq 67 -or $key -eq 45)) {
        return $true
    }

    return $false
}

function Read-ProcessMonitorKey {
    while ($true) {
        $key = Read-UiKey
        if (Test-IgnoredConsoleCopyKey -KeyInfo $key) {
            continue
        }

        return $key
    }
}

function Get-UiDisplayData {
    $status = Get-StatusObject
    $settings = Get-Settings
    $appVersion = Get-AppVersion
    $remoteConfig = Get-RemoteConfig
    $statusText = 'НЕ ЗАПУЩЕНО'
    $statusColor = 'DarkRed'

    $rawStatus = [string]$status.status
    if ($status.running) {
        $statusText = 'ЗАПУЩЕНО'
        $statusColor = 'Green'
    }

    if (-not [string]::IsNullOrWhiteSpace($rawStatus)) {
        if ($rawStatus -eq 'ОШИБКА КОНФИГА') {
            $statusText = 'ОШИБКА КОНФИГА'
            $statusColor = 'Red'
        }
        elseif ($rawStatus -like 'ошибка*') {
            $statusText = $rawStatus.ToUpperInvariant()
            $statusColor = 'DarkRed'
        }
        elseif ($rawStatus -eq 'ожидание config.json от Happ') {
            $statusText = 'ОЖИДАНИЕ CONFIG.JSON'
            $statusColor = 'DarkYellow'
        }
        elseif (-not $status.running) {
            $statusText = $rawStatus.ToUpperInvariant()
            $statusColor = 'DarkRed'
        }
    }


    $f1Text = 'выкл'
    $f2Text = 'выкл'
    $f3Text = 'вкл'
    if ($status.autostart_enabled) { $f1Text = 'вкл' }
    if ($status.keep_background_on_close) { $f2Text = 'вкл' }
    if (-not $status.auto_update_enabled) { $f3Text = 'выкл' }

    $f1Color = 'DarkGray'
    $f2Color = 'DarkGray'
    $f3Color = 'DarkMagenta'
    if ($f1Text -eq 'вкл') { $f1Color = 'Magenta' }
    if ($f2Text -eq 'вкл') { $f2Color = 'Magenta' }
    if ($f3Text -eq 'выкл') { $f3Color = 'DarkGray' }

    $remoteMessageText = '-'
    $remoteMessageColor = 'DarkGray'
    $remoteMessageRaw = [string]$remoteConfig.message
    if (-not [string]::IsNullOrWhiteSpace($remoteMessageRaw)) {
        $remoteMessageText = $remoteMessageRaw.Trim()
    }

    $remoteType = [string]$remoteConfig.type
    $latestVersion = [string]$remoteConfig.latest_version
    $versionCompare = Compare-AppVersions -LeftVersion $appVersion -RightVersion $latestVersion
    $shouldShowUpdate = ($remoteType -eq 'update' -and $versionCompare -lt 0)
    $shouldShowNews = (($remoteType -eq 'news' -or $remoteType -eq 'new') -and -not [string]::IsNullOrWhiteSpace($remoteMessageRaw))
    $disableFix = [bool]$remoteConfig.disable_fix
    $updateDownloadUrl = Get-RemoteConfigDownloadUrl -RemoteConfig $remoteConfig
    $messageUrl = [string]$remoteConfig.message_url
    $hasMessageUrl = -not [string]::IsNullOrWhiteSpace($messageUrl)

    # Логика отображения кнопок
    $showUpdateButtons = $false
    $showNewsButtons = $false
    $yButtonText = 'Скачать обновление'

    if ($shouldShowUpdate) {
        # type=update, версии разные — показываем автоматом (если не dismissed) или при [6]
        if (-not (Test-UpdatePromptDismissed -RemoteConfig $remoteConfig)) {
            $showUpdateButtons = $true
        }
        if ($script:ShowAnnouncementPrompt) {
            $showUpdateButtons = $true
        }
        $yButtonText = 'Скачать обновление'
    }
    elseif ($remoteType -eq 'update' -and $script:ShowAnnouncementPrompt) {
        # type=update, версии одинаковые, но нажали [6] — показываем Y/N
        $showUpdateButtons = $true
        $yButtonText = 'Открыть репозиторий'
    }

    if ($shouldShowNews -and $hasMessageUrl) {
        # type=news, есть ссылка — показываем G/N всегда
        $showNewsButtons = $true
    }

    if ($shouldShowUpdate) {
        $remoteMessageColor = 'Red'
        if ([string]::IsNullOrWhiteSpace($remoteMessageRaw)) {
            $remoteMessageText = ('Доступна новая версия: {0}' -f $latestVersion)
        }
    }
    elseif ($shouldShowNews) {
        $remoteMessageColor = 'DarkGray'
    }
    else {
        $remoteMessageText = '-'
        $remoteMessageColor = 'DarkGray'
    }
    
    if ($disableFix) {
        $statusText = 'ОТКЛЮЧЕН УДАЛЕННО'
        $statusColor = 'DarkRed'
        $showUpdateButtons = $false
        $showNewsButtons = $false
        if ($remoteMessageText -eq '-') {
            $remoteMessageText = 'Фикс временно отключен разработчиком'
        }
        $remoteMessageColor = 'Red'
    }

    $customConfigPath = [string]$settings.custom_config_path
    $effectiveConfigPath = Get-EffectiveConfigPath
    $displayConfigPath = ''
    if (-not [string]::IsNullOrWhiteSpace($customConfigPath)) {
        $displayConfigPath = $effectiveConfigPath
    }

    return [pscustomobject]@{
        status = $status
        remote_config = $remoteConfig
        app_version = $appVersion
        status_text = $statusText
        status_color = $statusColor
        f1_text = $f1Text
        f2_text = $f2Text
        f3_text = $f3Text
        f1_color = $f1Color
        f2_color = $f2Color
        f3_color = $f3Color
        remote_message_text = $remoteMessageText
        remote_message_color = $remoteMessageColor
        show_update_buttons = $showUpdateButtons
        show_news_buttons = $showNewsButtons
        y_button_text = $yButtonText
        update_download_url = $updateDownloadUrl
        disable_fix = $disableFix
        message_url = $messageUrl
        announcement_errors = @(Get-AnnouncementErrors)
        custom_config_path = $customConfigPath
        effective_config_path = [string]$effectiveConfigPath
        display_config_path = [string]$displayConfigPath
    }
}


function Open-UpdateDownloadPage {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return
    }

    try {
        Start-Process -FilePath $Url | Out-Null
    }
    catch {
        Write-AppLogLimited -Key ('update-page-open:' + $_.Exception.Message) -Message ('Не удалось открыть страницу обновления: ' + $_.Exception.Message)
    }
}

function Open-MessageUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return
    }

    # Если ссылка не начинается с http/https, добавляем https://
    $fullUrl = $Url
    if (-not ($Url -match '^https?://')) {
        $fullUrl = 'https://' + $Url
    }

    try {
        Start-Process -FilePath $fullUrl | Out-Null
        Write-AppLog ('Ссылка из объявления открыта: {0}' -f $Url)
    }
    catch {
        Write-AppLogLimited -Key ('message-url-open:' + $_.Exception.Message) -Message ('Не удалось открыть ссылку из объявления: ' + $_.Exception.Message)
    }
}

function Request-UiFlash {
    Update-State -Values @{ ui_flash_request = (Get-Date).ToString('o') }
}

function Render-UiScreen {
    param(
        [pscustomobject]$DisplayData,
        [string]$F1ColorOverride = '',
        [string]$F2ColorOverride = '',
        [string]$F3ColorOverride = ''
    )

    Clear-Host
    Write-TelegramLink
    Write-Host ''
    Write-Host 'Когда статус ЗАПУЩЕНО - переподключитесь в Happ. Иногда может потребоваться повторно переподключиться 2 раза' -ForegroundColor Yellow
    Write-Host '(например при старте Windows)' -ForegroundColor Yellow
    Write-Host 'Если после двух переподключений фикс не работает - полностью выключите Happ (даже в tray),' -ForegroundColor DarkYellow
    Write-Host 'нажмите [9] для полного перезапуска фикса и после переподключитесь, также 1-2 раза.' -ForegroundColor DarkYellow

    Write-Host ''
    Write-Host ('СТАТУС: {0}' -f $DisplayData.status_text) -ForegroundColor $DisplayData.status_color
    Write-Host ('PID: {0}' -f $(if ($DisplayData.status.pid) { $DisplayData.status.pid } else { '-' })) -ForegroundColor Cyan
    Write-Host ('Последний патч: {0}' -f (Format-RuDate $DisplayData.status.last_patch_at)) -ForegroundColor Gray
    Write-Host ('Последнее обновление rulesets: {0}' -f (Format-RuDate $DisplayData.status.last_rules_update_at)) -ForegroundColor Gray
    if ($DisplayData.status.autostart_error) {
        Write-Host ''
        Write-Host 'ВНИМАНИЕ: автозапуск не удалось обновить, но сам фикс продолжает работать.' -ForegroundColor DarkRed
        Write-Host $DisplayData.status.autostart_error -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Host '[1] Обновить rulesets' -ForegroundColor Cyan
    Write-Host '[2] Открыть лог' -ForegroundColor DarkCyan
    Write-Host '[3] Запустить/остановить фикс' -ForegroundColor Red
    Write-Host '[4] Запустить чекер процессов' -ForegroundColor Gray
    Write-Host '[5] Изменить путь до config.json' -ForegroundColor Gray
    if (-not [string]::IsNullOrWhiteSpace($DisplayData.display_config_path) -and -not $script:ShowConfigPathPrompt) {
        Write-Host $DisplayData.display_config_path -ForegroundColor DarkGray
    }

    if ($script:ShowConfigPathPrompt) {
        Write-Host ' Выберите файл config.json или папку, где он находится.' -ForegroundColor DarkYellow
        Write-Host ' Дефолтный путь: AppData\Local\Happ\config.json' -ForegroundColor DarkYellow
        Write-Host ' Отмена выбора файла возвращает дефолт путь.' -ForegroundColor DarkYellow
        Write-Host ''
    }
    Write-Host '[7] Открыть Telegram разработчика' -ForegroundColor DarkGray
    Write-Host '[8] Открыть репозиторий на GitHub' -ForegroundColor DarkGray
    Write-Host '[9] Полный перезапуск фикса' -ForegroundColor DarkMagenta
    Write-Host '[Esc] Выход' -ForegroundColor DarkGray

    Write-Host ''

    $showLink = (-not [string]::IsNullOrWhiteSpace($DisplayData.message_url))
    
    Write-Host '[6] Объявления: ' -ForegroundColor White -NoNewline
    if ($showLink) {
        Write-Host $DisplayData.remote_message_text -ForegroundColor $DisplayData.remote_message_color -NoNewline
        Write-Host ', ' -ForegroundColor DarkGray -NoNewline
        Write-Host $DisplayData.message_url -ForegroundColor DarkCyan
    }
    else {
        Write-Host $DisplayData.remote_message_text -ForegroundColor $DisplayData.remote_message_color
    }

    if ($DisplayData.show_update_buttons) {
        Write-Host ('Y {0}    N Скрыть предложение' -f $DisplayData.y_button_text) -ForegroundColor DarkYellow
    }
    if ($DisplayData.show_news_buttons) {
        Write-Host 'G Открыть ссылку' -ForegroundColor Magenta -NoNewline
        Write-Host '    N Скрыть предложение' -ForegroundColor DarkYellow
    }

    $announcementErrors = @($DisplayData.announcement_errors)
    if ($announcementErrors.Count -gt 0) {
        foreach ($err in $announcementErrors) {
            Write-Host $err -ForegroundColor Red
        }
    }
    Write-Host ''

    $f1Color = $DisplayData.f1_color
    $f2Color = $DisplayData.f2_color
    $f3Color = $DisplayData.f3_color
    if (-not [string]::IsNullOrWhiteSpace($F1ColorOverride)) { $f1Color = $F1ColorOverride }
    if (-not [string]::IsNullOrWhiteSpace($F2ColorOverride)) { $f2Color = $F2ColorOverride }
    if (-not [string]::IsNullOrWhiteSpace($F3ColorOverride)) { $f3Color = $F3ColorOverride }

    $script:UiToggleLine = [Console]::CursorTop
    Write-Host 'F1 Автозапуск Windows: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f1_text -ForegroundColor $f1Color -NoNewline
    Write-Host '    F2 Оставаться в фоне после закрытия: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f2_text -ForegroundColor $f2Color -NoNewline
    Write-Host '    F3 Автообновление rulesets: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f3_text -ForegroundColor $f3Color
}

function Write-UiToggleLine {
    param(
        [pscustomobject]$DisplayData,
        [string]$F1ColorOverride = '',
        [string]$F2ColorOverride = '',
        [string]$F3ColorOverride = ''
    )

    if ($script:UiToggleLine -lt 0) {
        return
    }

    $f1Color = $DisplayData.f1_color
    $f2Color = $DisplayData.f2_color
    $f3Color = $DisplayData.f3_color
    if (-not [string]::IsNullOrWhiteSpace($F1ColorOverride)) { $f1Color = $F1ColorOverride }
    if (-not [string]::IsNullOrWhiteSpace($F2ColorOverride)) { $f2Color = $F2ColorOverride }
    if (-not [string]::IsNullOrWhiteSpace($F3ColorOverride)) { $f3Color = $F3ColorOverride }

    try {
        [Console]::SetCursorPosition(0, $script:UiToggleLine)
        Write-Host (' ' * ([Math]::Max(1, [Console]::WindowWidth - 1))) -NoNewline
        [Console]::SetCursorPosition(0, $script:UiToggleLine)
    }
    catch {
        return
    }

    Write-Host 'F1 Автозапуск Windows: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f1_text -ForegroundColor $f1Color -NoNewline
    Write-Host '    F2 Оставаться в фоне после закрытия: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f2_text -ForegroundColor $f2Color -NoNewline
    Write-Host '    F3 Автообновление rulesets: ' -ForegroundColor White -NoNewline
    Write-Host $DisplayData.f3_text -ForegroundColor $f3Color
}

function Show-UiScreen {
    $display = Get-UiDisplayData
    Render-UiScreen -DisplayData $display
}

function Invoke-UiFlashAnimation {
    $display = Get-UiDisplayData

    foreach ($i in 1..4) {
        Write-UiToggleLine -DisplayData $display -F1ColorOverride 'DarkYellow' -F2ColorOverride 'DarkYellow' -F3ColorOverride 'DarkYellow'
        Start-Sleep -Milliseconds 420
        Write-UiToggleLine -DisplayData $display
        Start-Sleep -Milliseconds 420
    }
}

function Wait-UiKeyOrFlash {
    param([string]$LastFlashRequest)

    $lastUiSignature = Get-StateUiSignature

    while ($true) {
        $state = Get-State
        $currentFlashRequest = ''
        if ($state.ContainsKey('ui_flash_request')) {
            $currentFlashRequest = [string]$state.ui_flash_request
        }

        if (-not [string]::IsNullOrWhiteSpace($currentFlashRequest) -and $currentFlashRequest -ne $LastFlashRequest) {
            Invoke-UiFlashAnimation
            return [pscustomobject]@{
                key = $null
                flash_request = $currentFlashRequest
            }
        }

        $currentUiSignature = Get-StateUiSignature
        if ($currentUiSignature -ne $lastUiSignature) {
            return [pscustomobject]@{
                key = $null
                flash_request = $LastFlashRequest
            }
        }

        if ([Console]::KeyAvailable) {
            return [pscustomobject]@{
                key = Read-UiKey
                flash_request = $LastFlashRequest
            }
        }

        Start-Sleep -Milliseconds 80
    }
}


function Start-UiLoop {
    Initialize-Console

    $uiMutexCreated = $false
    $uiMutex = New-Object System.Threading.Mutex($true, 'Local\HappRoutingUiByGulenok91', [ref]$uiMutexCreated)

    if (-not $uiMutexCreated) {
        Request-UiFlash
        $null = Activate-UiWindow
        return
    }

    try {
        Update-State -Values @{
            ui_session_active = 'true'
            ui_pid            = [string]$PID
        }
        $settings = Get-Settings
        $null = Set-AutostartEntry -Enabled ([bool]$settings.autostart_enabled)
        if (Get-DesiredRunning) {
            $null = Start-AgentProcess -SetDesired $false
        }
        Start-UiCloseWatcher
        Request-UiFlash
        $lastFlashRequest = ''
        $processExitHandler = [System.EventHandler]{
            param($sender, $args)
            try {
                Invoke-UiShutdown -ShowBackgroundNotice (-not $script:UiExitRequested)
            }
            catch {
            }
        }
        [System.AppDomain]::CurrentDomain.add_ProcessExit($processExitHandler)

        Show-UiScreen
        while ($true) {
            $waitResult = Wait-UiKeyOrFlash -LastFlashRequest $lastFlashRequest
            $lastFlashRequest = [string]$waitResult.flash_request
            if ($null -eq $waitResult.key) {
                Show-UiScreen
                continue
            }
            if (-not (Handle-UiKey -KeyInfo $waitResult.key)) {
                break
            }
            Show-UiScreen
        }

    }
    finally {
        try {
            Invoke-UiShutdown -ShowBackgroundNotice (-not $script:UiExitRequested)
        }
        catch {
        }
        if ($null -ne $uiMutex) {
            try {
                $uiMutex.ReleaseMutex() | Out-Null
            }
            catch {
            }
            $uiMutex.Dispose()
        }
    }
}

function Handle-UiKey {
    param([object]$KeyInfo)

    $key = $KeyInfo.VirtualKeyCode
    $display = Get-UiDisplayData

    if ($key -eq 27) {
        $script:UiExitRequested = $true
        return $false
    }

    # Обработка Y — открыть репозиторий/скачать обновление
    if ($display.show_update_buttons -and ($key -eq 89)) {
        Dismiss-UpdatePrompt -RemoteConfig $display.remote_config
        $script:ShowAnnouncementPrompt = $false
        Open-UpdateDownloadPage -Url $display.update_download_url
        return $true
    }

    # Обработка N — скрыть предложение (update или news)
    if (($display.show_update_buttons -or $display.show_news_buttons) -and ($key -eq 78)) {
        if ($display.show_update_buttons) {
            Dismiss-UpdatePrompt -RemoteConfig $display.remote_config
        }
        $script:ShowAnnouncementPrompt = $false
        return $true
    }

    # Обработка G — Открыть ссылку из news
    if ($display.show_news_buttons -and ($key -eq 71)) {
        if (-not [string]::IsNullOrWhiteSpace($display.message_url)) {
            Open-MessageUrl -Url $display.message_url
        }
        return $true
    }

    if ($key -eq 112) {
        $settings = Toggle-Setting -Name 'autostart_enabled'
        $null = Set-AutostartEntry -Enabled ([bool]$settings.autostart_enabled)
        return $true
    }

    if ($key -eq 113) {
        $null = Toggle-Setting -Name 'keep_background_on_close'
        return $true
    }

    if ($key -eq 114) {
        $null = Toggle-Setting -Name 'auto_update_enabled'
        return $true
    }

    if ($key -eq 49 -or $key -eq 97) {
        try {
            Invoke-RulesUpdate
        }
        catch {
            Update-State -Values @{
                status              = 'ошибка обновления rulesets'
                last_error_message  = $_.Exception.Message
                next_rules_retry_at = (Get-Date).AddMinutes(5).ToString('s')
            }
            Write-AppLogLimited -Key ('rules-manual:' + $_.Exception.Message) -Message ('Ручное обновление rulesets завершилось ошибкой: ' + $_.Exception.Message)
        }
        return $true
    }

    if ($key -eq 50 -or $key -eq 98) {
        if (-not (Test-Path -LiteralPath $LogPath)) {
            New-Item -ItemType File -Path $LogPath -Force | Out-Null
        }
        Start-Process notepad.exe -ArgumentList $LogPath | Out-Null
        return $true
    }

    if ($key -eq 51 -or $key -eq 99) {
        if (Get-AgentProcess) {
            $null = Stop-AgentProcess
        }
        else {
            $null = Start-AgentProcess -SetDesired $true
        }
        return $true
    }

    if ($key -eq 52 -or $key -eq 100) {
        Start-ProcessCheckerWindow
        return $true
    }

    if ($key -eq 53 -or $key -eq 101) {
        # [5] - Изменить путь до config.json
        if (-not $script:ShowConfigPathPrompt) {
            # Первое нажатие - показываем текст и ждём перед диалогом
            $script:ShowConfigPathPrompt = $true
            Show-UiScreen
            Start-Sleep -Milliseconds 1800
            
            $selectedPath = ''
            $selectionCancelled = $false
            try {
                Add-Type -AssemblyName System.Windows.Forms | Out-Null
                $dlg = New-Object System.Windows.Forms.OpenFileDialog
                $dlg.Title = 'Выберите config.json Happ'
                $dlg.Filter = 'config.json|config.json|Все файлы|*.*'
                $dlg.FileName = ''
                $dlg.InitialDirectory = "$env:LOCALAPPDATA\Happ"

                # Создаём невидимую форму-родитель с TopMost, чтобы диалог был поверх консоли
                $parentForm = New-Object System.Windows.Forms.Form
                $parentForm.TopMost = $true
                $parentForm.WindowState = 'Minimized'
                $parentForm.ShowInTaskbar = $false
                $parentForm.StartPosition = 'Manual'
                $parentForm.Location = New-Object System.Drawing.Point(-10000, -10000)
                [void]$parentForm.Show()
                [void]$parentForm.Focus()

                $result = $dlg.ShowDialog($parentForm)
                $parentForm.Dispose()

                if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dlg.FileName)) {
                    $selectedPath = $dlg.FileName
                }
                else {
                    $selectionCancelled = $true
                }
            }
            catch {
                Write-AppLog "Ошибка при выборе пути к config.json: $($_.Exception.Message)"
                $selectionCancelled = $true
            }

            # Очищаем накопленные нажатия клавиш, пока был открыт диалог
            while ([Console]::KeyAvailable) {
                $null = [Console]::ReadKey($false)
            }

            if ($selectionCancelled -or [string]::IsNullOrWhiteSpace($selectedPath)) {
                # Отмена выбора - сброс на дефолт
                $settings = Get-DefaultSettings
                $current = Get-Settings
                foreach ($k in @($settings.Keys)) { $settings[$k] = $current.$k }
                $settings['custom_config_path'] = ''
                Save-Settings -Settings $settings
                $script:ShowConfigPathPrompt = $false
                Clear-AnnouncementErrors
                Write-AppLog 'Путь к config.json сброшен на дефолтный.'
                
                # Принудительно обновляем статус в state
                Update-State -Values @{ status = 'дефолтный путь' }
                
                Stop-AgentProcess -SetDesired $false
                Start-Sleep -Milliseconds 200
                Start-AgentProcess -SetDesired $true
                Start-Sleep -Milliseconds 200
                Show-UiScreen
                return $true
            }

            # Проверяем валидность выбранного пути
            $effectivePath = ''
            $pathValid = $false
            
            if (Test-Path -LiteralPath $selectedPath -PathType Leaf) {
                $effectivePath = $selectedPath
                $pathValid = $true
            }
            else {
                $effectivePath = $selectedPath
                $pathValid = $false
            }

            if (-not $pathValid) {
                # Путь не валиден - оставляем текст, не сохраняем путь
                Write-AppLog "Выбран неверный путь к config.json: $effectivePath"
                Show-UiScreen
                return $true
            }

            # Путь валиден - сохраняем, перезапускаем агент, скрываем текст
            $settings = Get-DefaultSettings
            $current = Get-Settings
            foreach ($k in @($settings.Keys)) { $settings[$k] = $current.$k }
            $settings['custom_config_path'] = $effectivePath
            Save-Settings -Settings $settings
            $script:ShowConfigPathPrompt = $false
            Clear-AnnouncementErrors
            Write-AppLog "Установлен путь к config.json: $effectivePath"

            # Принудительно обновляем статус в state
            Update-State -Values @{ status = 'смена конфига' }
            
            Stop-AgentProcess -SetDesired $false
            Start-Sleep -Milliseconds 300
            Start-AgentProcess -SetDesired $true
            Start-Sleep -Milliseconds 500
            Show-UiScreen
            return $true
        }
        else {
            # Повторное нажатие [5] - сброс на дефолт
            $settings = Get-DefaultSettings
            $current = Get-Settings
            foreach ($k in @($settings.Keys)) { $settings[$k] = $current.$k }
            $settings['custom_config_path'] = ''
            Save-Settings -Settings $settings
            $script:ShowConfigPathPrompt = $false
            Clear-AnnouncementErrors
            Write-AppLog 'Путь к config.json сброшен на дефолтный.'
            
            # Принудительно обновляем статус в state
            Update-State -Values @{ status = 'дефолтный путь' }
            
            Stop-AgentProcess -SetDesired $false
            Start-Sleep -Milliseconds 200
            Start-AgentProcess -SetDesired $true
            Start-Sleep -Milliseconds 200
            Show-UiScreen
            return $true
        }
    }

    if ($key -eq 55 -or $key -eq 103) {
        Open-DeveloperTelegram
        return $true
    }

    if ($key -eq 56 -or $key -eq 104) {
        try {
            Start-Process -FilePath 'https://github.com/gulenokdv/Report---Fix-Happ-Routing' | Out-Null
        }
        catch {
            Write-AppLogLimited -Key 'github-open' -Message 'Не удалось открыть репозиторий на GitHub.'
        }
        return $true
    }

    if ($key -eq 57 -or $key -eq 105) {
        Write-AppLog 'Запрошен полный перезапуск фикса из UI.'
        Restart-FixUi
        return $false
    }

    if ($key -eq 54 -or $key -eq 102) {

        # 6 - Объявления: синхронизация и показ кнопок
        try {
            $changed = Invoke-RemoteConfigSync
            if ($changed) {
                Update-State -Values @{ status = 'обновления синхронизированы' }
            }
            else {
                Update-State -Values @{ status = 'обновления не найдены' }
            }
        }
        catch {
            Update-State -Values @{
                status             = 'ошибка синхронизации обновлений'
                last_error_message = $_.Exception.Message
            }
            Write-AppLogLimited -Key ('remote-config-manual:' + $_.Exception.Message) -Message ('Ручная проверка обновлений завершилась ошибкой: ' + $_.Exception.Message)
        }
        $script:ShowAnnouncementPrompt = $true
        return $true
    }

    return $true
}

function Start-AgentLoop {
    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, 'Local\HappRoutingGuardByGulenok91', [ref]$createdNew)

    if (-not $createdNew) {
        exit 0
    }

    try {
        Update-State -Values @{
            pid        = [string]$PID
            started_at = (Get-Date).ToString('s')
            status     = 'запущено'
        }
        Write-AppLog 'Агент фикса запущен.'

        # При старте агента сразу проверяем обновления (даже если прошло мало времени)
        # Сбрасываем last_rules_update_at чтобы принудительно проверить rulesets
        $settings = Get-Settings
        Update-State -Values @{
            last_rules_update_at       = '2000-01-01T00:00:00'
            remote_config_boot_check_pending = 'true'
        }

        $lastStamp = ''
        $lastMode = ''
        $lastAutostartSyncAt = [datetime]::MinValue

        while ($true) {
            try {
            $settings = Get-Settings

            if ([bool]$settings.autostart_enabled -and (((Get-Date) - $lastAutostartSyncAt).TotalSeconds -ge 30)) {
                $null = Set-AutostartEntry -Enabled $true
                $lastAutostartSyncAt = Get-Date
            }

            if (Test-RemoteConfigNeedsSync) {
                try {
                    $remoteConfigChanged = Invoke-RemoteConfigSync
                    if ($remoteConfigChanged) {
                        Update-State -Values @{ status = 'обновления синхронизированы' }
                        $lastMode = 'remote config updated'
                    }
                }
                catch {
                    Update-State -Values @{
                        status                      = 'ошибка синхронизации обновлений'
                        last_error_message          = $_.Exception.Message
                        remote_config_boot_check_pending = 'false'
                        next_remote_config_retry_at = (Get-Date).AddMinutes(5).ToString('s')
                    }
                    if ($lastMode -ne 'remote config failed') {
                        Write-AppLogLimited -Key ('remote-config:' + $_.Exception.Message) -Message ('Синхронизация remote config завершилась ошибкой: ' + $_.Exception.Message)
                    }
                    $lastMode = 'remote config failed'
                    Start-Sleep -Seconds 1
                    continue
                }
            }

            $remoteConfig = Get-RemoteConfig
            if ([bool]$remoteConfig.disable_fix) {
                Update-State -Values @{ status = 'фикс отключен разработчиком' }
                $lastMode = 'remote disabled'
                Start-Sleep -Milliseconds 250
                continue
            }

            if (Test-RulesNeedUpdate) {
                try {
                    Invoke-RulesUpdate
                    Update-State -Values @{ status = 'rulesets обновлены' }
                    $lastMode = 'rulesets updated'
                }
                catch {
                    Update-State -Values @{
                        status              = 'ошибка обновления rulesets'
                        last_error_message  = $_.Exception.Message
                        next_rules_retry_at = (Get-Date).AddMinutes(5).ToString('s')
                    }
                    if ($lastMode -ne 'rulesets failed') {
                        Write-AppLogLimited -Key ('rules-auto:' + $_.Exception.Message) -Message ('Автообновление rulesets завершилось ошибкой: ' + $_.Exception.Message)
                    }
                    $lastMode = 'rulesets failed'
                    Start-Sleep -Seconds 1
                    continue
                }
            }

            $effectiveConfigPath = Get-EffectiveConfigPath
            if (-not (Test-Path -LiteralPath $effectiveConfigPath)) {
                if (-not [string]::IsNullOrWhiteSpace($script:LastObservedConfigPath) -or -not [string]::IsNullOrWhiteSpace($script:LastObservedConfigStamp)) {
                    Reset-ConfigCriticalErrorTracking
                    Clear-ConfigBanInfo
                    $script:LastObservedConfigPath = ''
                    $script:LastObservedConfigStamp = ''
                }
                Update-State -Values @{ status = 'ожидание config.json от Happ' }
                if ($script:LastAgentConfigError -ne 'not-found') {
                    $script:LastAgentConfigError = 'not-found'
                    Write-AppLogLimited -Key 'config-missing' -Message ("config.json не найден: $effectiveConfigPath")
                    Add-AnnouncementError -Message ("config.json не найден: $effectiveConfigPath. Нажмите [5] и выберите правильный путь.")
                }
                $lastMode = 'waiting config'
                Start-Sleep -Seconds 1
                continue
            }


            $currentStamp = Get-ConfigStamp
            $configIdentityChanged = ($script:LastObservedConfigPath -ne [string]$effectiveConfigPath -or $script:LastObservedConfigStamp -ne [string]$currentStamp)
            if ($configIdentityChanged) {
                Reset-ConfigCriticalErrorTracking
                $ban = Get-ConfigBanInfo
                $hasStaleBan = (-not [string]::IsNullOrWhiteSpace($ban.path) -and -not ([string]$ban.path -eq [string]$effectiveConfigPath -and [string]$ban.stamp -eq [string]$currentStamp))
                if ($script:LastAgentConfigError -eq 'disabled' -or $hasStaleBan) {
                    Clear-ConfigBanInfo
                    Clear-AnnouncementErrors
                    Write-AppLog ("Блокировка config.json снята из-за смены пути или файла: $effectiveConfigPath")
                }
                $script:LastAgentConfigError = ''
                $script:LastObservedConfigPath = [string]$effectiveConfigPath
                $script:LastObservedConfigStamp = [string]$currentStamp
            }

            $isBannedConfig = Test-ConfigBanned -Path $effectiveConfigPath -Stamp $currentStamp

            if ($isBannedConfig) {
                if ($script:LastAgentConfigError -ne 'disabled') {
                    $script:LastAgentConfigError = 'disabled'
                    Update-State -Values @{
                        status             = 'ОШИБКА КОНФИГА'
                        last_error_message = 'Выбранный config.json отключен после 6 критических ошибок.'
                    }
                    Add-AnnouncementError -Message ('ОШИБКА КОНФИГА: этот config.json временно отключен после 6 критических ошибок. Выберите другой путь через [5].')
                    Write-AppLog ("Пропуск забаненного config.json: $effectiveConfigPath")
                }
                $lastMode = 'config banned'
                Start-Sleep -Seconds 2
                continue
            }

            if ($script:LastAgentConfigError -eq 'not-found') {
                Clear-AnnouncementErrors
                Write-AppLog "config.json найден: $effectiveConfigPath"
                $script:LastAgentConfigError = ''
            }



            if ([string]::IsNullOrWhiteSpace($lastStamp)) {
                $lastStamp = [string]$currentStamp
            }
            elseif ([string]$currentStamp -ne $lastStamp) {
                $attempts = Invoke-BurstPatch
                $lastStamp = [string](Get-ConfigStamp)
                Update-State -Values @{ status = ('обнаружен rewrite, патчей: {0}' -f $attempts) }
                $lastMode = 'rewrite'
                Write-AppLog ('Обнаружен rewrite config.json, выполнено патчей: {0}' -f $attempts)
                Start-Sleep -Milliseconds 20
                continue
            }

            if (Test-ConfigNeedsPatch) {
                try {
                    $null = Invoke-ConfigPatch
                    $lastStamp = [string](Get-ConfigStamp)
                    Update-State -Values @{ status = 'готово'; last_error_message = '' }
                    $lastMode = 'patched'
                    Clear-AnnouncementErrors
                    Write-AppLog "config.json успешно пропатчен: $effectiveConfigPath"
                }

                catch {
                    Update-State -Values @{
                        status             = 'ошибка патча'
                        last_error_message = $_.Exception.Message
                    }
                    if ($lastMode -ne 'patch failed') {
                        $errorMsg = 'Патч config.json невозможен: ' + $_.Exception.Message
                        Write-AppLogLimited -Key ('patch:' + $_.Exception.Message) -Message $errorMsg
                        Add-AnnouncementError -Message $errorMsg
                    }
                    $lastMode = 'patch failed'
                }
            }
            else {
                Update-State -Values @{ status = 'готово' }
                $lastMode = 'ready'
            }

            Start-Sleep -Milliseconds 250

            }
            catch {
                $errMsg = $_.Exception.Message
                $displayErr = $errMsg
                if ($displayErr -match 'route' -or $displayErr -match 'config') {
                    $displayErr = 'Выбран неправильный config.json.'
                }

                $criticalCount = Register-ConfigCriticalError -ConfigPath $effectiveConfigPath
                Update-State -Values @{
                    status             = 'ОШИБКА КОНФИГА'
                    last_error_message = $errMsg
                }

                if ($criticalCount -le 6) {
                    Write-AppLogLimited -Key ('agent-critical:' + $errMsg) -Message ('Критическая ошибка агента: ' + $errMsg)
                    Add-AnnouncementError -Message ('ОШИБКА КОНФИГА: ' + $displayErr)
                }

                if ($criticalCount -ge 6) {
                    $script:LastAgentConfigError = 'disabled'
                    Set-ConfigBanInfo -Path $effectiveConfigPath -Stamp $currentStamp
                    Update-State -Values @{ status = 'ОШИБКА КОНФИГА' }
                    Write-AppLog ("Конфиг отключен после 6 критических ошибок: $effectiveConfigPath")
                    Add-AnnouncementError -Message ('ОШИБКА КОНФИГА: этот config.json временно отключен после 6 критических ошибок. Выберите другой путь через [5].')
                    Start-Sleep -Seconds 2
                    continue
                }


                Start-Sleep -Seconds 5
            }

        }
    }
    finally {
        Update-State -Values @{ pid = ''; status = 'остановлено' }
        Write-AppLog 'Агент фикса остановлен.'
        if ($null -ne $mutex) {
            try {
                $mutex.ReleaseMutex() | Out-Null
            }
            catch {
            }
            $mutex.Dispose()
        }
    }
}

try {
switch ($Action) {
    'ui' {
        Start-UiLoop
    }

    'agent' {
        Start-AgentLoop
    }

    'watch-ui-close' {
        Wait-UiCloseAndNotify -WatchedUiPid $UiPid
    }

    'ensure-defaults' {
        $settings = Get-Settings
        $null = Set-AutostartEntry -Enabled ([bool]$settings.autostart_enabled)
        [pscustomobject]@{ ok = $true } | ConvertTo-Json -Depth 5
    }

    'status-lines' {
        $status = Get-StatusObject
        'running={0}' -f $status.running
        'pid={0}' -f $status.pid
        'status={0}' -f $status.status
        'last_patch_at={0}' -f $status.last_patch_at
        'last_rules_update_at={0}' -f $status.last_rules_update_at
    }

    'status-json' {
        Get-StatusObject | ConvertTo-Json -Depth 10
    }

    'start-agent' {
        [pscustomobject]@{ ok = $true; pid = Start-AgentProcess } | ConvertTo-Json -Depth 10
    }

    'start-agent-if-enabled' {
        if (Get-DesiredRunning) {
            Update-State -Values @{ remote_config_boot_check_pending = 'true' }
            [pscustomobject]@{ ok = $true; started = $true; pid = (Start-AgentProcess -SetDesired $false) } | ConvertTo-Json -Depth 10
        }
        else {
            [pscustomobject]@{ ok = $true; started = $false; pid = '' } | ConvertTo-Json -Depth 10
        }
    }

    'stop-agent' {
        [pscustomobject]@{ stopped = (Stop-AgentProcess) } | ConvertTo-Json -Depth 10
    }

    'process-monitor' {
        Start-ProcessMonitor
    }

    'update-rulesets' {
        Invoke-RulesUpdate
        [pscustomobject]@{ ok = $true } | ConvertTo-Json -Depth 10
    }

    'open-log' {
        if (-not (Test-Path -LiteralPath $LogPath)) {
            New-Item -ItemType File -Path $LogPath -Force | Out-Null
        }
        Start-Process notepad.exe -ArgumentList $LogPath | Out-Null
        [pscustomobject]@{ ok = $true } | ConvertTo-Json -Depth 10
    }
}
}
catch {
    Write-AppLog ('Критическая ошибка действия [' + $Action + ']: ' + $_.Exception.Message)
    throw
}
