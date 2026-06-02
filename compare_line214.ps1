$origPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1.original"
$currPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1"

$origBytes = [System.IO.File]::ReadAllBytes($origPath)
$currBytes = [System.IO.File]::ReadAllBytes($currPath)

$origContent = [System.Text.Encoding]::UTF8.GetString($origBytes)
$currContent = [System.Text.Encoding]::UTF8.GetString($currBytes)

$origLines = $origContent -split "`n"
$currLines = $currContent -split "`n"

# Compare line 214 in both files
Write-Host "=== ORIGINAL Line 214 ==="
$ol = $origLines[213].TrimEnd("`r")
Write-Host ("[" + $ol + "]")
$c = $ol.ToCharArray()
for ($idx2 = 0; $idx2 -lt $c.Length; $idx2++) {
    $ch = $c[$idx2]
    $hx = [Convert]::ToString([int]$ch, 16)
    Write-Host ("  " + $idx2 + ": 0x$hx = '$ch'")
}

Write-Host ""
Write-Host "=== CURRENT Line 214 ==="
$cl = $currLines[213].TrimEnd("`r")
Write-Host ("[" + $cl + "]")
$c2 = $cl.ToCharArray()
for ($idx3 = 0; $idx3 -lt $c2.Length; $idx3++) {
    $ch2 = $c2[$idx3]
    $hx2 = [Convert]::ToString([int]$ch2, 16)
    Write-Host ("  " + $idx3 + ": 0x$hx2 = '$ch2'")
}

# Compare file sizes
Write-Host ""
Write-Host ("Original: " + $origBytes.Length + " bytes, " + $origLines.Length + " lines")
Write-Host ("Current:  " + $currBytes.Length + " bytes, " + $currLines.Length + " lines")
Write-Host ("Line 214 original content: [" + $ol + "]")
Write-Host ("Line 214 current  content: [" + $cl + "]")
Write-Host ("Same: " + ($ol -eq $cl))
