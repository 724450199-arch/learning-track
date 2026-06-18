param(
  [string]$PushPlusToken
)

# ====== 配置 ======
$LogFile = "$PSScriptRoot\news.log"
$CacheFile = "$PSScriptRoot\news_cache.json"
$MaxItems = 30

function Write-Log {
  param($Msg)
  $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  "$time $Msg" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log "===== 新闻采集开始 ====="

# ====== 来源配置 ======
$Sources = @(
  @{ Name = "36氪"; Url = "https://36kr.com/feed"; Type = "rss"; Tags = @("科技","AI","芯片") }
  @{ Name = "虎嗅"; Url = "https://www.huxiu.com/rss/0.xml"; Type = "rss"; Tags = @("科技") }
  @{ Name = "Reuters Tech"; Url = "https://feeds.reuters.com/reuters/technologyNews"; Type = "rss"; Tags = @("tech","AI","chip") }
)

# 中文搜索词：AI / 人工智能 / 芯片 / 半导体 / 封装 / 大模型 / 算力 / GPU / 科技
$Keywords = @(
  "AI","人工智能","大模型","LLM","AGI","OpenAI","ChatGPT","GPT","Claude","Gemini",
  "芯片","半导体","封装","Chiplet","先进封装","CoWoS","HBM","TSMC","台积电","ASML",
  "GPU","CPU","算力","数据中心","HPC","英伟达","NVIDIA","AMD","Intel",
  "光刻机","EDA","RISC-V","Arm",
  "科技","互联网","量子计算","5G","6G","机器人","自动驾驶"
)

# ====== 采集新闻 ======
$AllItems = @()

foreach ($Src in $Sources) {
  try {
    Write-Log "正在采集: $($Src.Name)"
    if ($Src.Type -eq "rss") {
      $xml = [System.Xml.XmlDocument]::new()
      $xml.Load($Src.Url)
      $items = $xml.rss.channel.item | Select-Object -First 20
      foreach ($item in $items) {
        $AllItems += @{
          title    = $item.title -replace '\s+',' '
          summary  = ($item.description -replace '<[^>]+>','').Trim() -replace '\s+',' '
          link     = $item.link
          source   = $Src.Name
          pubDate  = if ($item.pubDate) { $item.pubDate } else { (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
        }
      }
    }
  } catch {
    Write-Log "采集 $($Src.Name) 失败: $_"
  }
}

Write-Log "共采集 $($AllItems.Count) 条新闻"

# ====== 关键词过滤 ======
$Filtered = @()
$KeywordsPattern = "($($Keywords -join '|'))"
$SeenTitles = @{}

foreach ($item in $AllItems) {
  $text = "$($item.title) $($item.summary)"
  if ($text -match $KeywordsPattern) {
    $key = $item.title.Substring(0, [Math]::Min(50, $item.title.Length))
    if (-not $SeenTitles.ContainsKey($key)) {
      $SeenTitles[$key] = $true
      $MatchesList = [regex]::Matches($text, $KeywordsPattern) | ForEach-Object { $_.Value }
      $item.matchCount = ($MatchesList | Select-Object -Unique).Count
      $Filtered += $item
    }
  }
}

# 按匹配度排序
$Filtered = $Filtered | Sort-Object matchCount -Descending | Select-Object -First $MaxItems

Write-Log "过滤后 $($Filtered.Count) 条相关新闻"

# ====== 格式化日报 ======
$DateStr = Get-Date -Format "yyyy-MM-dd"
$TimeStr = Get-Date -Format "HH:mm"

$Body = @"
## 📰 每日科技早报
**$DateStr** | 共 $($Filtered.Count) 条精选

---
"@

$i = 1
foreach ($item in $Filtered) {
  $Summary = if ($item.summary) { $item.summary.Substring(0, [Math]::Min(120, $item.summary.Length)) } else { "" }
  $Summary = $Summary -replace '[\r\n]+',' '
  $Body += @"

**$i. $($item.title)**
> [$($item.source)] $Summary
🔗 $($item.link)
"@
  $i++
}

$Body += @"

---
*自动采集于 $DateStr $TimeStr*
*来源: 36氪 / 虎嗅 / Reuters / 半导体行业观察*
"@

# ====== 发送到微信 ======
function Send-PushPlus {
  param($Token, $Title, $Content)
  $body = @{
    token   = $Token
    title   = $Title
    content = $Content
    template = "markdown"
  } | ConvertTo-Json
  try {
    $resp = Invoke-RestMethod -Uri "https://www.pushplus.plus/send" -Method Post -Body $body -ContentType "application/json" -TimeoutSec 30
    if ($resp.code -eq 200) {
      Write-Log "PushPlus 发送成功"
    } else {
      Write-Log "PushPlus 发送失败: $($resp.msg)"
    }
  } catch {
    Write-Log "PushPlus 异常: $_"
  }
}

if (-not $PushPlusToken) {
  $ConfigFile = "$PSScriptRoot\pushplus_token.txt"
  if (Test-Path $ConfigFile) {
    $PushPlusToken = (Get-Content $ConfigFile -Raw).Trim()
  }
}

if ($PushPlusToken) {
  Send-PushPlus -Token $PushPlusToken -Title "每日科技早报 $DateStr" -Content $Body
  Write-Log "已推送到微信"
} else {
  Write-Log "⚠ PushPlus token 未配置，日报已保存到: $PSScriptRoot\daily_news_$DateStr.md"
  $Body | Out-File -FilePath "$PSScriptRoot\daily_news_$DateStr.md" -Encoding utf8
}

Write-Log "===== 采集结束 ====="
Write-Host "日报生成完毕，共 $($Filtered.Count) 条"
