$utilsPath = Join-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) "desktop_utils.ps1"
. $utilsPath

if (-not (Test-Path $script:backupDir -PathType Container)) {
    New-Item -ItemType Directory -Path $script:backupDir -Force -ErrorAction SilentlyContinue | Out-Null
}

Get-ChildItem $script:desktop -File -ErrorAction SilentlyContinue | ForEach-Object { Backup-File $_.FullName }

$restored = 0
Get-ChildItem $script:backupDir -File -ErrorAction SilentlyContinue | ForEach-Object {
    if (Restore-File $_.Name) { $restored++ }
}
if ($restored -gt 0) { Write-Log "Initial restore: $restored files" }

Write-Log "Watchdog started"

$knownFiles = @{}
Get-ChildItem $script:desktop -File -ErrorAction SilentlyContinue | Where-Object {
    $skip = $false; foreach ($p in $script:skipPattern) { if ($_.Name -match $p) { $skip = $true } }; -not $skip
} | ForEach-Object { $knownFiles[$_.Name] = $_.LastWriteTime.Ticks }

while ($true) {
    Start-Sleep -Seconds 1

    $currentFiles = @{}
    Get-ChildItem $script:desktop -File -ErrorAction SilentlyContinue | Where-Object {
        $skip = $false; foreach ($p in $script:skipPattern) { if ($_.Name -match $p) { $skip = $true } }; -not $skip
    } | ForEach-Object {
        $currentFiles[$_.Name] = $_.LastWriteTime.Ticks
        if (-not $knownFiles.ContainsKey($_.Name)) { Backup-File $_.FullName }
    }

    foreach ($name in $knownFiles.Keys) {
        if (-not $currentFiles.ContainsKey($name)) {
            if (Restore-File $name) { $currentFiles[$name] = $knownFiles[$name] }
        }
    }

    $knownFiles = $currentFiles
}
