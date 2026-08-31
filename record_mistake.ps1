<#
.SYNOPSIS
  错题拍照转文字并录入息流(FlowUs)错题数据库
.DESCRIPTION
  拖拽或选择一张错题照片 -> Agnes AI 视觉识别（算式/应用题更准，失败则 Windows OCR 兜底） -> 显示并确认 -> 写入多多错题数据库
.EXAMPLE
  .\record_mistake.ps1 -Image "D:\错题\多多_数学_01.jpg"
  .\record_mistake.ps1   (弹出文件选择框)
#>
param(
  [string]$Image,
  [string]$Db   # "class" 课堂练习错题 | "ao" 浅奥错题，不传则交互选择
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ====== 配置 ======
$token = "zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj"
$base = "https://api.flowus.cn/v2"
$dbClass = "b072234a-0207-418c-bdd0-b19ec19b4e98"   # 课堂练习错题
$dbAo    = "da6af8d0-4aeb-43d2-b914-c338035f05de"   # 浅奥错题

# Agnes AI 视觉识别配置（Agent AI 系统，OpenAI 兼容多模态）
$agnesBase = "https://apihub.agnes-ai.com/v1"
$agnesModel = "agnes-2.0-flash"
function Get-AgnesKey {
  $k = $env:AGNES_API_KEY
  if (-not $k) {
    $kf = Join-Path (Join-Path $env:USERPROFILE ".agnes") "api_key"
    if (Test-Path $kf) { $k = [IO.File]::ReadAllText($kf, [Text.Encoding]::UTF8).Trim() }
  }
  if (-not $k) { throw "未找到 Agnes API key（环境变量 AGNES_API_KEY 或 ~/.agnes/api_key）" }
  return $k
}

# ====== 工具函数 ======

# Agnes AI 视觉识别（识别错题图片文字/算式/应用题）
function Get-AgnesVisionText([string]$imgPath) {
  $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($imgPath))
  $dataUri = "data:image/png;base64,$b64"
  $sys = "你是一年级数学错题识别助手。请完整读出图片中的题目文字与算式，按原样逐条列出。若为计算/应用题请把题目、条件和要求的算式都读出。不要输出思考过程，直接把识别到的文字结果写出来。"
  $body = @{
    model = $agnesModel
    messages = @(@{ role = "user"; content = @(
      @{ type = "text"; text = $sys }
      @{ type = "image_url"; image_url = @{ url = $dataUri } }
    )})
    max_tokens = 1024
    temperature = 0.2
  }
  $r = Invoke-RestMethod -Uri "$agnesBase/chat/completions" -Method Post `
    -Body ($body | ConvertTo-Json -Depth 12 -Compress) -ContentType "application/json" `
    -Headers @{ Authorization = "Bearer $(Get-AgnesKey)" } -TimeoutSec 120
  $content = $r.choices[0].message.content
  if (-not [string]::IsNullOrWhiteSpace($content)) { return $content.Trim() }
  # 兜底：答案落入 reasoning_content
  $reason = $r.choices[0].message.reasoning_content
  if ($reason) { return $reason.Trim() }
  return $null
}

# Windows 自带 OCR（后备方案）
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
function Await-Async($WinRtTask, $ResultType) {
  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
  $netTask = $asTask.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  $netTask.Result
}
[Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null
[Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime] | Out-Null

function Get-OcrText([string]$imgPath) {
  $fileTask = [Windows.Storage.StorageFile]::GetFileFromPathAsync($imgPath)
  $file = Await-Async $fileTask ([Windows.Storage.StorageFile])
  $streamTask = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
  $stream = Await-Async $streamTask ([Windows.Storage.Streams.IRandomAccessStream])
  $decoderTask = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
  $decoder = Await-Async $decoderTask ([Windows.Graphics.Imaging.BitmapDecoder])
  $bitmapTask = $decoder.GetSoftwareBitmapAsync()
  $bitmap = Await-Async $bitmapTask ([Windows.Graphics.Imaging.SoftwareBitmap])
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new("zh-Hans-CN"))
  if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
  if (-not $engine) { throw "无法创建 OCR 引擎" }
  $resultTask = $engine.RecognizeAsync($bitmap)
  $result = Await-Async $resultTask ([Windows.Media.Ocr.OcrResult])
  $stream.Dispose(); [GC]::Collect()
  $sb = New-Object System.Text.StringBuilder
  foreach ($line in $result.Lines) { $sb.AppendLine($line.Text) | Out-Null }
  return $sb.ToString().Trim()
}

function Call-FlowUs($method, $path, $body, $mbOk) {
  $params = @{ Uri = $base + $path; Method = $method; ContentType = "application/json; charset=utf-8"; Headers = @{ "Authorization" = "Bearer $token" }; ErrorAction = "Stop" }
  if ($body) { $params.Body = ($body | ConvertTo-Json -Depth 30 -Compress) }
  try { return Invoke-RestMethod @params -TimeoutSec 30 }
  catch {
    if (($mbOk -eq $true) -and $_.Exception.Response -and ([int]$_.Exception.Response.StatusCode) -eq 303) {
      try { $loc = $_.Exception.Response.Headers["Location"]; Write-Host "redirect: $loc"; return $loc } catch { }
    }
    Write-Host ("FlowUs ERR: " + $_.Exception.Message)
    try { Write-Host ("BODY: " + ($_.ErrorDetails.Message)) } catch { }
    throw
  }
}

function New-MistakeRecord($dbId, $title, $today) {
  $props = @{
    "题目" = @{ type = "title"; title = @(@{ type = "text"; text = @{ content = $title } }) }
    "日期" = @{ type = "date"; date = @{ start = $today } }
  }
  $body = @{
    parent = @{ database_id = $dbId }
    properties = $props
  }
  $r = Call-FlowUs -method POST -path "/pages" -body $body
  return $r
}

# ====== 选择图片 ======
if (-not $Image) {
  Add-Type -AssemblyName System.Windows.Forms
  $ofd = New-Object System.Windows.Forms.OpenFileDialog
  $ofd.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.webp"
  $ofd.Title = "选择错题照片"
  if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host "已取消"; exit 0 }
  $Image = $ofd.FileName
}
if (-not (Test-Path -LiteralPath $Image)) { Write-Error "文件不存在: $Image"; exit 1 }

# ====== 识别（Agnes AI 视觉优先，失败则 Windows OCR 兜底） ======
Write-Host "正在用 Agnes AI 识别图片（算式/应用题更准）..."
$text = $null
try {
  $text = Get-AgnesVisionText -imgPath $Image
} catch {
  Write-Host ("Agnes 识别失败: " + $_.Exception.Message + "，改用 Windows OCR")
}
if (-not $text) {
  Write-Host "正在用 Windows OCR 识别..."
  $text = Get-OcrText -imgPath $Image
}
Write-Host "`n========== 识别结果 =========="
Write-Host $text
Write-Host "============================="

# ====== 确认/编辑题目 ======
Write-Host ""
Write-Host "是否使用以上文字作为错题题目记录？"
Write-Host "  [y] 直接使用"
Write-Host "  [n] 手动输入/修改题目"
$ans = Read-Host "选择 (y/n)"
if ($ans -notmatch '^y') {
  $title = Read-Host "输入题目文字"
} else {
  $title = ($text -replace "`r?`n", "；")
}
if (-not $title.Trim()) { Write-Host "题目为空，已取消"; exit 0 }

# ====== 选择目标库 ======
if ($Db -notmatch '^(class|ao)$') {
  Write-Host ""
  Write-Host "选择目标错题库："
  Write-Host "  [1] 课堂练习错题"
  Write-Host "  [2] 浅奥错题"
  $dbSel = Read-Host "选择 (1/2)"
  $Db = if ($dbSel -eq '2') { 'ao' } else { 'class' }
}

$dbId = if ($Db -eq 'ao') { $dbAo } else { $dbClass }
$dbName = if ($Db -eq 'ao') { "浅奥错题" } else { "课堂练习错题" }

# ====== 写入息流 ======
$today = (Get-Date).ToString("yyyy-MM-dd")
Write-Host "`n正在写入 $dbName 数据库..."
$rec = New-MistakeRecord -dbId $dbId -title $title -today $today
Write-Host "已录入！记录ID: $($rec.id)"
Write-Host "打开查看: https://flowus.cn/$($rec.id)"
