param(
  [string]$SendKey
)

# ====== 配置 ======
$LogFile = "$PSScriptRoot\news.log"
$MaxItems = 30
$RequestTimeout = 15

function Write-Log {
  param($Msg)
  $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  "$time $Msg" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log "===== 新闻采集开始 ====="

# ====== RSS 来源 ======
$Sources = @(
  @{ Name = "36氪"; Url = "https://36kr.com/feed" }
  @{ Name = "IT之家"; Url = "https://www.ithome.com/rss/" }
  @{ Name = "Reuters Tech"; Url = "https://feeds.reuters.com/reuters/technologyNews" }
)

# ====== 关键词 ======
$Keywords = @(
  "AI","人工智能","大模型","LLM","AGI","OpenAI","ChatGPT","GPT","Claude","Gemini",
  "芯片","半导体","封装","Chiplet","先进封装","CoWoS","HBM","TSMC","台积电","ASML",
  "GPU","CPU","算力","数据中心","HPC","英伟达","NVIDIA","AMD","Intel",
  "光刻机","EDA","RISC-V","Arm",
  "科技","互联网","量子计算","5G","6G","机器人","自动驾驶"
)

# ====== 采集 ======
$AllItems = @()
foreach ($Src in $Sources) {
  try {
    Write-Log "正在采集: $($Src.Name)"
    $resp = Invoke-WebRequest -Uri $Src.Url -TimeoutSec $RequestTimeout -UseBasicParsing -ErrorAction Stop
    $xml = [System.Xml.XmlDocument]::new()
    $xml.LoadXml($resp.Content)
    $items = $xml.rss.channel.item | Select-Object -First 25
    foreach ($item in $items) {
      $title = if ($item.title.'#cdata-section') { $item.title.'#cdata-section' } else { "$($item.title)" }
      $desc  = if ($item.description.'#cdata-section') { $item.description.'#cdata-section' } else { "$($item.description)" }
      $link  = if ($item.link.'#cdata-section') { $item.link.'#cdata-section' } else { "$($item.link)" }
      $AllItems += @{
        title   = $title -replace '\s+',' '
        summary = ($desc -replace '<[^>]+>','').Trim() -replace '\s+',' '
        link    = $link.Trim()
        source  = $Src.Name
      }
    }
  } catch {
    Write-Log "采集 $($Src.Name) 失败: $_"
  }
}

Write-Log "共采集 $($AllItems.Count) 条新闻"

# ====== 关键词过滤去重 ======
$Filtered = @()
$KeywordsPattern = "($($Keywords -join '|'))"
$Seen = @{}

foreach ($item in $AllItems) {
  $text = "$($item.title) $($item.summary)"
  if ($text -match $KeywordsPattern) {
    $key = $item.title.Substring(0, [Math]::Min(50, $item.title.Length))
    if (-not $Seen.ContainsKey($key)) {
      $Seen[$key] = $true
      $MatchesList = [regex]::Matches($text, $KeywordsPattern) | ForEach-Object { $_.Value }
      $item.matchCount = ($MatchesList | Select-Object -Unique).Count
      $Filtered += $item
    }
  }
}
$Filtered = $Filtered | Sort-Object matchCount -Descending | Select-Object -First $MaxItems
Write-Log "过滤后 $($Filtered.Count) 条相关新闻"

# ====== 格式化 ======
$DateStr = Get-Date -Format "yyyy-MM-dd"
$TimeStr = Get-Date -Format "HH:mm"
$Body = @"
## 每日科技早报
**$DateStr** | 共 $($Filtered.Count) 条精选

"@

$i = 1
foreach ($item in $Filtered) {
  $summary = if ($item.summary) { $item.summary.Substring(0, [Math]::Min(120, $item.summary.Length)) } else { "" }
  $summary = $summary -replace '[\r\n]+',' '
  $Body += "**$i. $($item.title)**
> [$($item.source)] $summary
$($item.link)

"
  $i++
}

$Body += "---
自动采集于 $DateStr $TimeStr | 来源: 36氪 / IT之家 / Reuters"

# ====== 方糖推送 ======
if (-not $SendKey) {
  $ConfigFile = "$PSScriptRoot\sendkey.txt"
  if (Test-Path $ConfigFile) {
    $SendKey = (Get-Content $ConfigFile -Raw).Trim()
  }
}

if ($SendKey) {
  try {
    $resp = Invoke-RestMethod -Uri "https://sctapi.ftqq.com/$SendKey.send" -Method Post `
      -Body @{ title = "每日科技早报 $DateStr"; content = $Body } `
      -TimeoutSec 30
    if ($resp.code -eq 0) {
      Write-Log "方糖推送成功 (pushid: $($resp.data.pushid))"
    } else {
      Write-Log "方糖推送失败: code=$($resp.code) msg=$($resp.message)"
    }
  } catch {
    Write-Log "方糖推送异常: $_"
  }
} else {
  $out = "$PSScriptRoot\daily_news_$DateStr.md"
  $Body | Out-File -FilePath $out -Encoding utf8
  Write-Log "SendKey 未配置，日报已保存到: $out"
  Write-Host "已保存到 $out"
}

Write-Log "===== 采集结束 ====="
Write-Host "完成，共 $($Filtered.Count) 条"
