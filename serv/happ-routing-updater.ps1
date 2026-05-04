param(
    [string]$StagingPath,
    [string]$InstallPath,
    [string]$CurrentVersion,
    [string]$UiPid
)

# happ-routing-updater.ps1
# Временный updater для автообновления

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] $Message"
}

function Test-OurProcess {
    param([object]$Process)
    if ($null -eq $Process) { return $false }
    $cmd = $Process.CommandLine.ToLowerInvariant()
    return ($cmd.Contains('happ-routing-fix.bat') -or $cmd.Contains('happ-routing.ps1'))
}

function Get-OurProcesses {
    $processes = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" -ErrorAction SilentlyContinue
    return @($processes | Where-Object { Test-OurProcess -Process $_ })
}

function Wait-ForProcessesToExit {
    param([array]$Processes, [int]$TimeoutMs = 10000)
    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    while ((Get-Date) -lt $deadline) {
        $stillRunning = @()
        foreach ($p in $Processes) {
            try {
                $check = Get-Process -Id $p.ProcessId -ErrorAction Stop
                $stillRunning += $p
            }
            catch {}
        }
        if ($stillRunning.Count -eq 0) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Kill-OurProcesses {
    param([array]$Processes)
    foreach ($p in $Processes) {
        try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Write-Log "Убит процесс PID $($p.ProcessId)"
        }
        catch {
            Write-Log "Не удалось убить процесс $($p.ProcessId): $($_.Exception.Message)"
        }
    }
}

function Copy-Files {
    param([string]$Source, [string]$Dest)
    Get-ChildItem -Path $Source -Recurse -File | ForEach-Object {
        $destPath = Join-Path -Path $Dest -ChildPath $_.FullName.Substring($Source.Length).TrimStart('\')
        $destDir = Split-Path -Path $destPath -Parent
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Force -Path $destDir | Out-Null
        }
        Copy-Item -Path $_.FullName -Destination $destPath -Force
    }
}

try {
    Write-Log "=== happ-routing-updater запущен ==="
    Write-Log "Staging: $StagingPath"
    Write-Log "Install: $InstallPath"
    Write-Log "Target version: $CurrentVersion"

    if (-not (Test-Path -LiteralPath $StagingPath)) {
        throw "Staging path не найден: $StagingPath"
    }
    if (-not (Test-Path -LiteralPath $InstallPath)) {
        throw "Install path не найден: $InstallPath"
    }

    # Проверка структуры
    $requiredItems = @('happ-routing-fix.bat', 'scripts', 'serv', 'rulesets')
    foreach ($item in $requiredItems) {
        if (-not (Test-Path -LiteralPath (Join-Path -Path $StagingPath -ChildPath $item))) {
            throw "В staging нет обязательного элемента: $item"
        }
    }

    # Найти и остановить наши процессы
    Write-Log "Поиск наших процессов..."
    $ourProcesses = Get-OurProcesses
    if ($ourProcesses.Count -gt 0) {
        Write-Log "Найдено процессов: $($ourProcesses.Count)"
        Kill-OurProcesses -Processes $ourProcesses
        
        Write-Log "Ожидание завершения процессов..."
        $gottenAway = -not (Wait-ForProcessesToExit -Processes $ourProcesses -TimeoutMs 8000)
        if ($gottenAway) {
            Write-Log "Не все процессы успели завершиться, продолжаем..."
        }
    }

    # Бэкап runtime-файлов
    Write-Log "Бэкап runtime-файлов..."
    $runtimeFiles = @(
        'happ-routing-settings.json',
        'happ-routing-state.json',
        'happ-routing.log',
        'happ-routing-remote-config-cache.json',
        'serv/update-state.json'
    )
    $backupDir = Join-Path -Path $InstallPath -ChildPath 'backup-runtime'
    if (Test-Path -LiteralPath $backupDir) {
        Remove-Item -LiteralPath $backupDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    foreach ($rf in $runtimeFiles) {
        $rfPath = Join-Path -Path $InstallPath -ChildPath $rf
        if (Test-Path -LiteralPath $rfPath) {
            Copy-Item -Path $rfPath -Destination (Join-Path -Path $backupDir -ChildPath $rf) -Force
        }
    }

    # Копирование новых файлов
    Write-Log "Копирование новых файлов..."
    Copy-Files -Source $StagingPath -Dest $InstallPath

    # Восстановление runtime-файлов
    Write-Log "Восстановление runtime-файлов..."
    foreach ($rf in $runtimeFiles) {
        $rfPath = Join-Path -Path $InstallPath -ChildPath $rf
        $rfBackup = Join-Path -Path $backupDir -ChildPath $rf
        if (Test-Path -LiteralPath $rfBackup) {
            Copy-Item -Path $rfBackup -Destination $rfPath -Force
        }
    }

    # Очистка бэкапа
    if (Test-Path -LiteralPath $backupDir) {
        Remove-Item -LiteralPath $backupDir -Recurse -Force
    }

    # Запуск нового фикса
    Write-Log "Запуск обновлённого фикса..."
    $launcherPath = Join-Path -Path $InstallPath -ChildPath 'happ-routing-fix.bat'
    Start-Process -FilePath $launcherPath -WindowStyle Normal

    # Удаление updater
    Write-Log "Очистка updater..."
    try {
        $self = $MyInvocation.MyCommand.Path
        Start-Sleep -Milliseconds 500
        Remove-Item -LiteralPath $self -Force -ErrorAction SilentlyContinue
    }
    catch {}

    Write-Log "=== Обновление завершено ==="
    exit 0
}
catch {
    Write-Log "ОШИБКА: $($_.Exception.Message)"
    exit 1
}
