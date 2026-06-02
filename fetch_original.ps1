$token = [Environment]::GetEnvironmentVariable("GITEE_TOKEN", "User")
$headers = @{ Authorization = "Bearer $token" }
$url = "https://gitee.com/api/v5/repos/yql8981229/learning-track/contents/daily_learning.ps1"
try {
  $resp = Invoke-RestMethod -Uri $url -Headers $headers
  $decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($resp.content))
  $originalPath = Join-Path $PSScriptRoot "daily_learning.ps1.original"
  [System.IO.File]::WriteAllText($originalPath, $decoded, [System.Text.Encoding]::UTF8)
  Write-Host "Saved original to $originalPath"
  Write-Host ("Size: " + $decoded.Length + " chars, " + [Text.Encoding]::UTF8.GetByteCount($decoded) + " bytes")
} catch {
  $errorMsg = $_.Exception.Message
  if ($_.Exception.Response) {
    $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
    $errorBody = $reader.ReadToEnd()
    $reader.Close()
    $errorMsg = $errorBody
  }
  Write-Host "Error: $errorMsg"
}
