$script:desktop = [Environment]::GetFolderPath("Desktop")
$script:backupDir = Join-Path $script:desktop ".desktop_backup"
$script:logFile = Join-Path $env:TEMP "desktop_watchdog.log"
$script:skipPattern = '^\.desktop_backup', '^desktop\.ini$'

function Write-Log {
    param([string]$msg)
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try { "$time $msg" | Out-File -FilePath $script:logFile -Encoding utf8 -Append } catch {}
}

function Backup-File {
    param([string]$path)
    $name = Split-Path $path -Leaf
    foreach ($p in $script:skipPattern) { if ($name -match $p) { return } }
    $bp = Join-Path $script:backupDir $name
    try { if (Test-Path $path) { Copy-Item $path $bp -Force -ErrorAction SilentlyContinue } } catch {}
}

function Restore-File {
    param([string]$name)
    $bp = Join-Path $script:backupDir $name
    $tp = Join-Path $script:desktop $name
    if ((Test-Path $bp) -and -not (Test-Path $tp)) {
        try {
            Copy-Item $bp $tp -Force -ErrorAction SilentlyContinue
            if (Test-Path $tp) { Write-Log "Restored: $name"; return $true }
        } catch {}
    }
    return $false
}

function New-DesktopShortcut {
    param([string]$TargetPath)
    $target = Resolve-Path $TargetPath -ErrorAction SilentlyContinue
    if (-not $target) { Write-Output "路径不存在: $TargetPath"; return }
    $name = Split-Path $target -Leaf
    $shortcutPath = Join-Path $script:desktop "$name.lnk"
    if (Test-Path $shortcutPath) { Write-Output "快捷方式已存在: $shortcutPath"; return }
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $target.Path
        if (Test-Path $target.Path -PathType Container) { $shortcut.Description = "文件夹: $name" }
        $shortcut.Save()
        Write-Output "已创建快捷方式: $shortcutPath"
    } catch {
        Write-Output "创建快捷方式失败: $_"
    }
}
