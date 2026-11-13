param(
  [switch]$Push,
  [string]$OutFile = "E:\learning-track\kids_edu\多多学习中心.md",
  [string]$RootPageId = "b80cd768-6ef5-4da6-a257-e2afe8d388ac",
  [int]$MaxDepth = 3
)
$ErrorActionPreference = "Stop"
$token = "zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj"
$mcp   = "https://mcp.flowus.cn/message"
$repo  = "E:\learning-track"
$logFile = Join-Path (Split-Path $OutFile -Parent) "backup.log"

function Log($msg) {
  $line = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $msg
  [IO.File]::AppendAllText($logFile, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
}

$script:Sid = $null

function Invoke-Raw($body, $sid) {
  $req = [System.Net.HttpWebRequest]::Create($mcp)
  $req.Method = "POST"
  $req.ContentType = "application/json"
  $req.Accept = "application/json, text/event-stream"
  $req.Headers.Add("Authorization", "Bearer $token")
  if ($sid) { $req.Headers.Add("Mcp-Session-Id", $sid) }
  $req.Timeout = 120000
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
  $req.ContentLength = $bytes.Length
  $s = $req.GetRequestStream(); $s.Write($bytes, 0, $bytes.Length); $s.Close()
  $resp = $req.GetResponse()
  $sidOut = $null
  try { $sidOut = $resp.Headers["Mcp-Session-Id"] } catch {}
  $r = New-Object System.IO.StreamReader($resp.GetResponseStream(), [System.Text.Encoding]::UTF8)
  return @{ sid = $sidOut; body = $r.ReadToEnd() }
}

function Ensure-Session {
  if ($script:Sid) { return }
  $init = '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"weekly-backup","version":"1.0"}}}'
  $r = Invoke-Raw $init $null
  $script:Sid = $r.sid
  if (-not $script:Sid) { throw "MCP initialize 失败（未拿到 Mcp-Session-Id）" }
  Invoke-Raw '{"jsonrpc":"2.0","method":"notifications/initialized"}' $script:Sid | Out-Null
}

function Call-Mcp($name, $arguments) {
  Ensure-Session
  $params = @{ name = $name; arguments = $arguments } | ConvertTo-Json -Depth 12 -Compress
  $call = '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":' + $params + '}'
  $r = Invoke-Raw $call $script:Sid
  $parts = New-Object System.Collections.ArrayList
  foreach ($line in ($r.body -split "`n")) {
    $t = $line.TrimEnd("`r")
    if ($t.StartsWith("data:")) { [void]$parts.Add($t.Substring(5).TrimStart()) }
  }
  $payload = if ($parts.Count -gt 0) { $parts -join "" } else { $r.body }
  $j = $payload | ConvertFrom-Json
  if ($j.error) { throw ("MCP error: " + ($j.error | ConvertTo-Json -Depth 6 -Compress)) }
  if ($j.result.isError) { throw ("tool error: " + $j.result.content[0].text) }
  return $j.result.content[0].text
}

function Get-PageMarkdown($pageId) {
  $txt = Call-Mcp "flowus_get_page_markdown" @{ page_id = $pageId }
  return ($txt | ConvertFrom-Json).markdown
}

function Get-Blocks($blockId) {
  $txt = Call-Mcp "flowus_get_block_children" @{ block_id = $blockId; page_size = 100 }
  return ($txt | ConvertFrom-Json).results
}

function Get-BlockTitle($b) {
  $t = $b.type
  $val = $b.PSObject.Properties[$t].Value
  if ($t -eq "child_page" -or $t -eq "child_database") { return $val.title }
  if ($val -and $val.PSObject.Properties["rich_text"]) {
    return (($val.rich_text | ForEach-Object { $_.plain_text }) -join "")
  }
  return ""
}

$sb = New-Object System.Text.StringBuilder
$now = Get-Date
[void]$sb.AppendLine("# 多多学习中心（FlowUs 备份）")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("> 数据来源：[多多学习中心](https://flowus.cn/" + $RootPageId + ")　｜　页面 ID：``" + $RootPageId + "``")
[void]$sb.AppendLine("> 自动备份时间：" + $now.ToString("yyyy-MM-dd HH:mm") + "　｜　由 ``kids_edu/backup_flowus_duoduo.ps1`` 生成，请勿手改")
[void]$sb.AppendLine("")

$rootBlocks = Get-Blocks $RootPageId
$rootMd = Get-PageMarkdown $RootPageId
[void]$sb.AppendLine("## 正文")
[void]$sb.AppendLine("")
[void]$sb.AppendLine($rootMd.Trim())
[void]$sb.AppendLine("")

# 目录结构
[void]$sb.AppendLine("## 目录")
[void]$sb.AppendLine("")
foreach ($b in $rootBlocks) {
  if ($b.type -eq "child_page") {
    [void]$sb.AppendLine("- [文档] " + (Get-BlockTitle $b))
  } elseif ($b.type -eq "child_database") {
    [void]$sb.AppendLine("- [数据库] " + (Get-BlockTitle $b))
  }
}
[void]$sb.AppendLine("")

function Export-Node($pageId, $title, $depth) {
  $md = Get-PageMarkdown $pageId
  $head = ("#" * ($depth + 2))
  $pad = "　" * ($depth - 1)
  [void]$sb.AppendLine($head + " " + $pad + $title)
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("> <https://flowus.cn/" + $pageId + ">")
  [void]$sb.AppendLine("")
  if ($md -and $md.Trim().Length -gt 0) {
    [void]$sb.AppendLine($md.Trim())
  } else {
    [void]$sb.AppendLine("_（本页暂无正文）_")
  }
  [void]$sb.AppendLine("")
  if ($depth -ge $MaxDepth) { return }
  $blocks = Get-Blocks $pageId
  foreach ($b in $blocks) {
    if ($b.type -eq "child_page") {
      Export-Node $b.id (Get-BlockTitle $b) ($depth + 1)
    } elseif ($b.type -eq "child_database") {
      [void]$sb.AppendLine(("#" * ($depth + 3)) + " [数据库] " + (Get-BlockTitle $b))
      [void]$sb.AppendLine("")
      [void]$sb.AppendLine("> <https://flowus.cn/" + $b.id + ">　记录未导出，可在 FlowUs 中查看")
      [void]$sb.AppendLine("")
    }
  }
}

[void]$sb.AppendLine("---")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("# 子页面内容")
[void]$sb.AppendLine("")
foreach ($b in $rootBlocks) {
  if ($b.type -eq "child_page") {
    try {
      Export-Node $b.id (Get-BlockTitle $b) 1
    } catch {
      [void]$sb.AppendLine("> ⚠ 导出失败：" + $_.Exception.Message)
      [void]$sb.AppendLine("")
    }
  }
}

$text = $sb.ToString()
[IO.Directory]::CreateDirectory((Split-Path $OutFile -Parent)) | Out-Null
[IO.File]::WriteAllText($OutFile, $text, (New-Object System.Text.UTF8Encoding($false)))
Log ("导出完成: " + $OutFile + "  " + $text.Length + " 字符")

if ($Push) {
  Push-Location $repo
  try {
    git add -- "kids_edu/多多学习中心.md"
    $st = git status --porcelain -- "kids_edu/多多学习中心.md"
    if ([string]::IsNullOrWhiteSpace($st)) {
      Log "无变化，跳过提交"
      Write-Host "no changes"
    } else {
      git commit -m ("weekly flowus backup - " + $now.ToString("yyyy-MM-dd HH:mm")) | Out-Null
      Log "已提交"
      $p1 = git push origin main 2>&1
      if ($LASTEXITCODE -eq 0) { Log "已推送 origin（GitHub + Gitee）" }
      else {
        Log ("origin 推送失败: " + ($p1 -join " "))
        $p2 = git push gitee main 2>&1
        if ($LASTEXITCODE -eq 0) { Log "已推送 gitee" } else { Log ("gitee 推送失败: " + ($p2 -join " ")) }
      }
    }
  } finally {
    Pop-Location
  }
}
Write-Host "done"
