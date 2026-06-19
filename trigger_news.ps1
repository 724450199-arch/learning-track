param()

$LogFile = "$PSScriptRoot\news_trigger.log"

function Write-Log {
  param($Msg)
  $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  "$time $Msg" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log "===== 触发 GitHub Actions 开始 ====="

try {
  $token = "protocol=https`nhost=github.com`n`n" | git credential fill 2>&1 | Select-String "password=" | ForEach-Object { $_.ToString().Replace("password=","") }
  if (-not $token) { throw "未能从凭据管理器获取 GitHub token" }

  $headers = @{ Authorization = "Bearer $token"; Accept = "application/vnd.github+json" }
  $body = @{ ref = "main" } | ConvertTo-Json

  Invoke-RestMethod -Uri "https://api.github.com/repos/724450199-arch/learning-track/actions/workflows/daily-news.yml/dispatches" -Headers $headers -Method Post -Body $body -ContentType "application/json" -TimeoutSec 15

  Write-Log "GitHub Actions 触发成功"
} catch {
  Write-Log "触发失败: $_"
}

Write-Log "===== 触发结束 ====="
