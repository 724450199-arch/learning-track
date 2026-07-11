param([string]$SendKey)

# 添加代理支持
$ProxyUrl = "http://127.0.0.1:7897"

# 读取原脚本，在 Invoke-WebRequest 调用前注入代理
$script = Get-Content "C:\Users\yang\AppData\Local\LearningEnglish\daily_news.ps1" -Raw

# 在文件开头添加代理变量
$proxyBlock = @'
$Script:UseProxy = $null
if (Test-Connection 127.0.0.1 -Count 1 -Quiet) {
  try {
    $Script:UseProxy = New-Object System.Net.WebProxy('http://127.0.0.1:7897')
    Write-Log "本地代理 127.0.0.1:7897 已就绪"
  } catch { }
}

'@

# 替换 Invoke-WebRequest 调用 - 在 -UseBasicParsing 后面加上 -Proxy
$modifiedScript = $script -replace '(Invoke-WebRequest .*?-UseBasicParsing)', '$1 -Proxy $Script:UseProxy'

# 在文件顶部插入代理定义
$modifiedScript = $proxyBlock + $modifiedScript

# 临时脚本路径
$tempScript = "C:\Users\yang\AppData\Local\LearningEnglish\daily_news_proxy.ps1"
$modifiedScript | Out-File -FilePath $tempScript -Encoding utf8

Write-Host "代理脚本已生成，正在运行..."
if ($SendKey) {
  & $tempScript -SendKey $SendKey
} else {
  & $tempScript
}

Remove-Item -LiteralPath $tempScript -Force
Remove-Item -LiteralPath $PSCommandPath -Force