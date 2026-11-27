<#
.SYNOPSIS
  多多题目分析→录入息流（带图片）全流程脚本
.DESCRIPTION
  按 2026-11-27 用户确认的标准格式，把多多（一年级）的题目录入息流数据库并上传题目图片。
  防重复：录入前先查库，同题已存在则只补图片不重复创建。
.PARAMETER ImagePath      题目图片路径
.PARAMETER Db            目标数据库：mistake(重点难点,默认) / ao(浅奥错题)
.PARAMETER Title          条目标题（简述）
.PARAMETER Attr           题目内容/已知条件/孩子答案
.PARAMETER Note           验证过程+判定结论+图片来源
.PARAMETER Date           日期 yyyy-MM-dd，默认今天
.PARAMETER Section        排序/加减法/图形与空间/英语/其他，默认其他
.PARAMETER Mastery        未掌握/需巩固/已掌握
.PARAMETER Cause          粗心/概念不清/审题不清/计算失误/方法不会（做对了留空）
.PARAMETER Subject        语文/数学/英语/其他，默认数学
.EXAMPLE
  .\duoduo_flowus_upload.ps1 -ImagePath "E:\临时文件\多多错题\数独类.png" -Title "3x3填数..." -Attr "..." -Note "..." -Mastery 已掌握
#>
param(
  [Parameter(Mandatory=$true)][string]$ImagePath,
  [string]$Db = "mistake",
  [Parameter(Mandatory=$true)][string]$Title,
  [string]$Attr = "",
  [string]$Note = "",
  [string]$Date = (Get-Date -Format "yyyy-MM-dd"),
  [string]$Section = "其他",
  [string]$Mastery = "需巩固",
  [string]$Cause = "",
  [string]$Subject = "数学",
  [string]$OutFile = "C:\Users\yang\AppData\Local\Temp\opencode\duoduo_upload_result.txt"
)

$ErrorActionPreference = "Stop"
$token = (Select-String -Path "$PSScriptRoot\upload_to_flowus.ps1" -Pattern '^\$token\s*=\s*"([^"]+)"').Matches[0].Groups[1].Value
$dbId = if ($Db -eq "ao") { "da6af8d0-4aeb-43d2-b914-c338035f05de" } else { "b072234a-0207-418c-bdd0-b19ec19b4e98" }
$base = "https://api.flowus.cn/v2"
$out = @()

function Req($method, $path, $bodyBytes, $ctype, $hdrs) {
  $req = [Net.HttpWebRequest]::Create($base + $path)
  $req.Method = $method
  $req.Headers.Add("Authorization", "Bearer $token")
  if ($hdrs) { foreach ($k in $hdrs.Keys) { $req.Headers.Add($k, $hdrs[$k]) } }
  if ($bodyBytes) {
    $req.ContentType = $ctype
    $req.ContentLength = $bodyBytes.Length
    $st = $req.GetRequestStream(); $st.Write($bodyBytes,0,$bodyBytes.Length); $st.Close()
  }
  try {
    $resp = $req.GetResponse()
    $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
    $txt = $sr.ReadToEnd(); $sr.Close()
    return @{ ok=$true; status=[int]$resp.StatusCode; text=$txt }
  } catch [System.Net.WebException] {
    $we = $_.Exception; $msg = ""
    if ($we.Response) { $sr2 = New-Object IO.StreamReader($we.Response.GetResponseStream(), [Text.Encoding]::UTF8); $msg = $sr2.ReadToEnd(); $sr2.Close() }
    return @{ ok=$false; status=[int]$we.Status; text=$msg }
  }
}
function Json($o) { [Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress -Depth 30)) }
function RichArr($t) { if ([string]::IsNullOrEmpty($t)) { return @() } else { return ,@(@{ type="text"; text=@{ content=$t } }) } }

if (-not (Test-Path -LiteralPath $ImagePath)) { throw "图片不存在: $ImagePath" }
$fileBytes = [IO.File]::ReadAllBytes($ImagePath)
$fileName = [IO.Path]::GetFileName($ImagePath)
$out += "题目图片: $fileName ($($fileBytes.Length) bytes)"

# --- 0) 防重复：查库中是否有相同标题的活跃条目 ---
$q = Req "POST" "/databases/$dbId/query" (Json @{page_size=100}) "application/json; charset=utf-8" $null
if ($q.ok) {
  $jq = $q.text | ConvertFrom-Json
  foreach ($r in $jq.results) {
    if ($r.in_trash) { continue }
    $etitle = ""
    foreach ($p in $r.properties.PSObject.Properties) { if ($p.Value.type -eq "title") { $etitle = ($p.Value.title | ForEach-Object { $_.plain_text }) -join "" } }
    if ($etitle -eq $Title) {
      $pageId = $r.id
      $out += "已存在同题记录 $pageId，跳过创建，只补图片"
      $skipCreate = $true
      break
    }
  }
}

# --- 1) 创建条目（若未重复）---
if (-not $skipCreate) {
  $props = @{
    "题目" = @{ type="title"; title=(RichArr $Title) }
    "属性" = @{ type="rich_text"; rich_text=(RichArr $Attr) }
    "订正与备注" = @{ type="rich_text"; rich_text=(RichArr $Note) }
    "日期" = @{ type="date"; date=@{ start=$Date } }
    "Section" = @{ type="select"; select=@{ name=$Section } }
    "掌握情况" = @{ type="select"; select=@{ name=$Mastery } }
    "学科" = @{ type="select"; select=@{ name=$Subject } }
  }
  if ($Cause) { $props["错因分析"] = @{ type="select"; select=@{ name=$Cause } } }
  $c = Req "POST" "/pages" (Json @{ parent=@{database_id=$dbId}; properties=$props }) "application/json; charset=utf-8" $null
  if (-not $c.ok) { throw "创建条目失败: $($c.text)" }
  $pageId = ($c.text | ConvertFrom-Json).id
  $out += "已创建条目: $pageId"
} else {
  $pageId = $pageId
}

# --- 2) 上传题目图片（图片块，不是文件卡片）---
$u = Req "POST" "/files/upload-url" (Json @{ filename=$fileName; content_type="image/png"; content_length=$fileBytes.Length; parent=@{page_id=$pageId} }) "application/json; charset=utf-8" $null
if (-not $u.ok) { throw "获取上传地址失败: $($u.text)" }
$ju = $u.text | ConvertFrom-Json
$hdrs = @{}
if ($ju.headers) { foreach ($hp in $ju.headers.PSObject.Properties) { $hdrs[$hp.Name] = $hp.Value } }
$up = Req $ju.method $ju.upload_url $fileBytes "image/png" $hdrs
if (-not ($up.ok -or ($up.status -ge 200 -and $up.status -lt 300))) { throw "图片字节上传失败: $($up.text)" }
$out += "图片字节已上传 (oss_name=$($ju.oss_name))"

$blk = Req "PATCH" "/blocks/$pageId/children" (Json @{ children=@(@{ object="block"; type="image"; image=@{ type="file"; file=@{ oss_name=$ju.oss_name; size=[int]$ju.size } } }) }) "application/json; charset=utf-8" $null
if (-not $blk.ok) { throw "添加图片块失败: $($blk.text)" }
$out += "已添加图片块到页面"

# --- 3) 验证 ---
$v = Req "GET" "/blocks/$pageId/children?page_size=100" $null $null $null
if ($v.ok) {
  $jv = $v.text | ConvertFrom-Json
  $imgCount = @($jv.results | Where-Object { $_.type -eq "image" }).Count
  $out += "验证: 页面图片块数量=$imgCount"
}
$out += "页面ID: $pageId"
[IO.File]::WriteAllText($OutFile, ($out -join "`r`n"), (New-Object Text.UTF8Encoding($false)))
Write-Output ($out -join "`r`n")
