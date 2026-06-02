$origPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1.original"
$currPath = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1" 

function Test-Parse($path, $lbl) {
    $errs = $null; $toks = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$toks, [ref]$errs)
    if ($errs.Count -gt 0) {
        Write-Host ($lbl + ": " + $errs.Count + " error(s)")
        foreach ($e in $errs) {
            Write-Host ("  Line " + $e.Extent.StartLineNumber + ": " + $e.Message)
        }
    } else {
        Write-Host ($lbl + ": PARSE OK")
    }
}

Test-Parse $origPath "Original"
Test-Parse $currPath "Current"
