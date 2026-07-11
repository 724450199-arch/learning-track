param([string[]]$paths)
if (-not $paths -or $paths.Count -eq 0) { exit }
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell
foreach ($p in $paths) {
    if (-not (Test-Path $p)) { continue }
    $name = Split-Path $p -Leaf
    $lnk = Join-Path $desktop "$name.lnk"
    $shortcut = $shell.CreateShortcut($lnk)
    $shortcut.TargetPath = (Resolve-Path $p).Path
    $shortcut.Save()
}
