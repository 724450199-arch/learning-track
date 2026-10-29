param(
  [Parameter(Position=0)]
  [ValidateSet("on","off","status","env-on","env-off")]
  [string]$Action = "status"
)

$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$proxyServer = "127.0.0.1:7897"
$proxyUrl = "http://127.0.0.1:7897"
# Domains that must always use the local network (bypass Clash).
$noProxy = "127.0.0.1,localhost,::1,opencode.ai,opencodeusercontent.com,deepseek.com,agnes-ai.com,models.dev,gitee.com"

function Show-Status {
  $cfg = Get-ItemProperty -Path $regPath
  Write-Host ("System proxy : ProxyEnable=" + $cfg.ProxyEnable + "  ProxyServer=" + $cfg.ProxyServer)
  $alive = Test-NetConnection -ComputerName 127.0.0.1 -Port 7897 -WarningAction SilentlyContinue -InformationLevel Quiet
  if ($alive) { Write-Host "Clash 7897   : running" } else { Write-Host "Clash 7897   : NOT running" }
  $hp = [Environment]::GetEnvironmentVariable("HTTPS_PROXY","User")
  $np = [Environment]::GetEnvironmentVariable("NO_PROXY","User")
  if ($hp) { Write-Host ("App proxy env: HTTPS_PROXY=" + $hp) } else { Write-Host "App proxy env: (empty)" }
  if ($np) { Write-Host ("NO_PROXY     : " + $np) }
  Write-Host ""
  Write-Host "opencode / git / curl read the env vars; browsers read the system proxy."
  Write-Host "ChatGPT web needs: on    |    opencode OpenAI model needs: Clash running only"
}

switch ($Action) {
  "on" {
    Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 1
    Set-ItemProperty -Path $regPath -Name ProxyServer -Value $proxyServer
    Write-Host ("System proxy ON: " + $proxyServer)
    Show-Status
  }
  "off" {
    Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 0
    Write-Host "System proxy OFF"
    Show-Status
  }
  "env-on" {
    [Environment]::SetEnvironmentVariable("HTTPS_PROXY",$proxyUrl,"User")
    [Environment]::SetEnvironmentVariable("HTTP_PROXY",$proxyUrl,"User")
    [Environment]::SetEnvironmentVariable("NO_PROXY",$noProxy,"User")
    [Environment]::SetEnvironmentVariable("no_proxy",$noProxy,"User")
    Write-Host "App proxy env ON (restart opencode to take effect)"
    Show-Status
  }
  "env-off" {
    [Environment]::SetEnvironmentVariable("HTTPS_PROXY",$null,"User")
    [Environment]::SetEnvironmentVariable("HTTP_PROXY",$null,"User")
    [Environment]::SetEnvironmentVariable("NO_PROXY",$null,"User")
    [Environment]::SetEnvironmentVariable("no_proxy",$null,"User")
    Write-Host "App proxy env OFF (restart opencode to take effect)"
    Show-Status
  }
  default {
    Show-Status
  }
}
