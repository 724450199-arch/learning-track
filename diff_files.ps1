$origPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1.original"
$currPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1"

$origBytes = [System.IO.File]::ReadAllBytes($origPath)
$currBytes = [System.IO.File]::ReadAllBytes($currPath)

$origContent = [System.Text.Encoding]::UTF8.GetString($origBytes)
$currContent = [System.Text.Encoding]::UTF8.GetString($currBytes)

$origLines = $origContent -split "`n"
$currLines = $currContent -split "`n"

$minLen = [Math]::Min($origLines.Length, $currLines.Length)
$diffLines = @()

for ($i = 0; $i -lt $minLen; $i++) {
    $oL = $origLines[$i].TrimEnd("`r")
    $cL = $currLines[$i].TrimEnd("`r")
    if ($oL -ne $cL) {
        $diffLines += $i + 1
        if ($diffLines.Count -le 10) {
            Write-Host ("Line " + ($i+1) + ": DIFFER")
            Write-Host ("  Orig: [" + $oL + "]")
            Write-Host ("  Curr: [" + $cL + "]")
        }
    }
}

Write-Host ("")
Write-Host ("Total differences: " + $diffLines.Count)
if ($diffLines.Count -gt 0) {
    Write-Host ("First 5 diff lines: " + (($diffLines | Select-Object -First 5) -join ", "))
    Write-Host ("Original total lines: " + $origLines.Length)
    Write-Host ("Current total lines: " + $currLines.Length)
}
