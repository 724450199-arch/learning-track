param(
  [Parameter(Position=0)]
  [ValidateSet("on","off","status")]
  [string]$Action = "status"
)

$regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$proxyServer = "127.0.0.1:7897"

function Show-Status {
  $cfg = Get-ItemProperty -Path $regPath
  Write-Host ("ProxyEnable : " + $cfg.ProxyEnable)
  Write-Host ("ProxyServer : " + $cfg.ProxyServer)
  $alive = Test-NetConnection -ComputerName 127.0.0.1 -Port 7897 -WarningAction SilentlyContinue -InformationLevel Quiet
  if ($alive) { $state = "Clash(7897): running" } else { $state = "Clash(7897): not running" }
  Write-Host $state
}

switch ($Action) {
  "on" {
    Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 1
    Set-ItemProperty -Path $regPath -Name ProxyServer -Value $proxyServer
    Write-Host ("Proxy ON: " + $proxyServer)
    Show-Status
  }
  "off" {
    Set-ItemProperty -Path $regPath -Name ProxyEnable -Value 0
    Write-Host "Proxy OFF"
    Show-Status
  }
  default {
    Show-Status
  }
}