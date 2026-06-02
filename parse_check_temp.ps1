[Console]::OutputEncoding = [Text.Encoding]::UTF8
$file = 'C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1'
$content = [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8)
$tokens = $null
$errors = $null
try {
  $null = [System.Management.Automation.Language.Parser]::ParseInput($content, [ref]$tokens, [ref]$errors)
  if ($errors.Count -gt 0) {
    Write-Host ("Found " + $errors.Count + " errors:")
    foreach ($e in $errors) {
      Write-Host ("  Line " + $e.Extent.StartLineNumber + ": " + $e.Message)
    }
  } else {
    Write-Host "No parse errors"
  }
} catch {
  Write-Host ("Exception: " + $_.Exception.Message)
}
