@(
' Agnes-AI API Helper (通义万相 · 阿里云百炼)'
' ============================================'
' Models: wanx-v1 (image), wan2.6-t2v/i2v (video)'
' API Base: https://dashscope.aliyuncs.com'
''

$QuotaFile = Join-Path (Join-Path $env:LOCALAPPDATA "LearningEnglish") "agnes_quota.json"

# -------- Quota Tracking -------- #

function Get-AgnesQuota {
    $default = @{
        image_count = 0
        video_seconds = 0
        quota_image_max = 500
        quota_video_sec_max = 50
        quota_expire_days = 90
        created_at = (Get-Date).ToString("yyyy-MM-dd")
        warned_low_image = $false
        warned_exhausted_image = $false
        warned_low_video = $false
        warned_exhausted_video = $false
    }
    if (Test-Path $QuotaFile) {
        try {
            $data = Get-Content $QuotaFile -Raw -Encoding UTF8 | ConvertFrom-Json
            return $data
        } catch {
            Write-Warning "Failed to read quota file, resetting"
        }
    }
    return $default
}

function Save-AgnesQuota($data) {
    $data | ConvertTo-Json | Set-Content $QuotaFile -Encoding UTF8
}

function Test-AgnesQuota($type, $amount = 1) {
    $q = Get-AgnesQuota
    if ($type -eq "image") {
        $used = [int]$q.image_count
        $max = [int]$q.quota_image_max
        $remaining = $max - $used
        $after = $used + $amount
        if ($after -ge $max -and -not $q.warned_exhausted_image) {
            Write-Host "==============================================" -ForegroundColor Red
            Write-Host "  wanx-v1 免费额度已用完 ($used/$max 张)" -ForegroundColor Red
            Write-Host "  续费: https://bailian.console.aliyun.com/" -ForegroundColor Red
            Write-Host "==============================================" -ForegroundColor Red
            $q.warned_exhausted_image = $true
            Save-AgnesQuota $q
        }
        elseif ($remaining -le 50 -and -not $q.warned_low_image) {
            Write-Host "⚠ wanx-v1 免费额度剩余仅 $remaining/$max 张" -ForegroundColor Yellow
            $q.warned_low_image = $true
            Save-AgnesQuota $q
        }
        return $remaining -gt 0
    }
    elseif ($type -eq "video") {
        $used = [int]$q.video_seconds
        $max = [int]$q.quota_video_sec_max
        $remaining = $max - $used
        $after = $used + $amount
        if ($after -ge $max -and -not $q.warned_exhausted_video) {
            Write-Host "==============================================" -ForegroundColor Red
            Write-Host "  wan2.6 免费额度已用完 ($used/$max 秒)" -ForegroundColor Red
            Write-Host "  续费: https://bailian.console.aliyun.com/" -ForegroundColor Red
            Write-Host "==============================================" -ForegroundColor Red
            $q.warned_exhausted_video = $true
            Save-AgnesQuota $q
        }
        elseif ($remaining -le 10 -and -not $q.warned_low_video) {
            Write-Host "⚠ wan2.6 免费额度剩余仅 $remaining/$max 秒" -ForegroundColor Yellow
            $q.warned_low_video = $true
            Save-AgnesQuota $q
        }
        return $remaining -ge $amount
    }
    return $true
}

function Update-AgnesQuota($type, $amount = 1) {
    $q = Get-AgnesQuota
    if ($type -eq "image") {
        $q.image_count = [int]$q.image_count + $amount
    }
    elseif ($type -eq "video") {
        $q.video_seconds = [int]$q.video_seconds + $amount
    }
    Save-AgnesQuota $q
    $used = if ($type -eq "image") { $q.image_count } else { $q.video_seconds }
    $max = if ($type -eq "image") { $q.quota_image_max } else { $q.quota_video_sec_max }
    Write-Host "已用: $used/$max ($([math]::Round(($max-$used)/$max*100))% 剩余)" -ForegroundColor Cyan
}

function Show-AgnesQuota {
    $q = Get-AgnesQuota
    Write-Host "`n========== 通义万相 免费额度 ==========" -ForegroundColor Cyan
    $imgRemain = [int]$q.quota_image_max - [int]$q.image_count
    Write-Host "  wanx-v1 图片: $($q.image_count)/$($q.quota_image_max) 张 (剩余 $imgRemain)" -ForegroundColor $(if ($imgRemain -le 0){"Red"}elseif($imgRemain -le 50){"Yellow"}else{"Green"})
    $vidRemain = [int]$q.quota_video_sec_max - [int]$q.video_seconds
    Write-Host "  wan2.6  视频: $($q.video_seconds)/$($q.quota_video_sec_max) 秒 (剩余 $vidRemain)" -ForegroundColor $(if ($vidRemain -le 0){"Red"}elseif($vidRemain -le 10){"Yellow"}else{"Green"})
    Write-Host "========================================`n" -ForegroundColor Cyan
}

# -------- API Key -------- #

function Get-DashScopeApiKey {
    $key = $env:DASHSCOPE_API_KEY
    if (-not $key) {
        $keyFile = Join-Path (Join-Path $env:USERPROFILE ".dashscope") "api_key"
        if (Test-Path $keyFile) {
            $key = [System.IO.File]::ReadAllText($keyFile, [Text.Encoding]::UTF8).Trim()
        }
    }
    if (-not $key) {
        Write-Error "DASHSCOPE_API_KEY not found. Set env var or save to ~/.dashscope/api_key"
        return $null
    }
    return $key
}

# -------- Image Generation -------- #

function Invoke-AgnesImage {
    param(
        [string]$Prompt,
        [string]$Model = "wanx-v1",
        [string]$Size = "1024*1024",
        [int]$N = 1,
        [string]$OutputFile,
        [int]$PollInterval = 5
    )

    if (-not (Test-AgnesQuota "image" $N)) {
        Write-Error "wanx-v1 免费额度已用完！"
        return $null
    }

    $key = Get-DashScopeApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        input = @{ prompt = $Prompt }
        parameters = @{ size = $Size; n = $N }
    }

    $json = $body | ConvertTo-Json -Depth 3
    $headers = @{
        Authorization = "Bearer $key"
        "X-DashScope-Async" = "enable"
    }

    try {
        Write-Host "Creating image task..."
        $task = Invoke-RestMethod -Uri "https://dashscope.aliyuncs.com/api/v1/services/aigc/text2image/image-synthesis" `
            -Method Post -Body $json -ContentType "application/json" -Headers $headers

        $taskId = $task.output.task_id
        Write-Host "Task ID: $taskId"

        do {
            Start-Sleep -Seconds $PollInterval
            $result = Invoke-RestMethod -Uri "https://dashscope.aliyuncs.com/api/v1/tasks/$taskId" `
                -Headers @{ Authorization = "Bearer $key" }
            Write-Host "Status: $($result.output.task_status)"
        } while ($result.output.task_status -eq "PENDING" -or $result.output.task_status -eq "RUNNING")

        if ($result.output.task_status -eq "SUCCEEDED" -and $result.output.results) {
            $url = $result.output.results[0].url
            if ($OutputFile) {
                Invoke-WebRequest -Uri $url -OutFile $OutputFile
                Write-Host "Saved: $OutputFile"
            }
            Write-Host "Image URL: $url"
            Update-AgnesQuota "image" $N
            return $result
        } else {
            Write-Error "Generation failed: $($result.output.task_status)"
            return $null
        }
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

# -------- Video Generation -------- #

function Invoke-AgnesVideo {
    param(
        [string]$Prompt,
        [string]$Model = "wan2.6-t2v",
        [int]$Duration = 5,
        [string]$InputImage,
        [string]$OutputFile,
        [int]$PollInterval = 10
    )

    if (-not (Test-AgnesQuota "video" $Duration)) {
        Write-Error "wan2.6 免费额度已用完！"
        return $null
    }

    $key = Get-DashScopeApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        input = @{ prompt = $Prompt }
        parameters = @{ duration = $Duration }
    }

    if ($InputImage) {
        $body.model = "wan2.6-i2v"
        $body.input.img_url = $InputImage
    }

    $json = $body | ConvertTo-Json -Depth 3
    $headers = @{
        Authorization = "Bearer $key"
        "X-DashScope-Async" = "enable"
    }

    try {
        Write-Host "Creating video task..."
        $task = Invoke-RestMethod -Uri "https://dashscope.aliyuncs.com/api/v1/services/aigc/video-generation/video-synthesis" `
            -Method Post -Body $json -ContentType "application/json" -Headers $headers

        $taskId = $task.output.task_id
        Write-Host "Task ID: $taskId"

        do {
            Start-Sleep -Seconds $PollInterval
            $result = Invoke-RestMethod -Uri "https://dashscope.aliyuncs.com/api/v1/tasks/$taskId" `
                -Headers @{ Authorization = "Bearer $key" }
            Write-Host "Status: $($result.output.task_status)"
        } while ($result.output.task_status -eq "PENDING" -or $result.output.task_status -eq "RUNNING")

        if ($result.output.task_status -eq "SUCCEEDED" -and $result.output.results) {
            $videoUrl = $result.output.results[0].url
            if ($OutputFile) {
                Invoke-WebRequest -Uri $videoUrl -OutFile $OutputFile
                Write-Host "Saved: $OutputFile"
            }
            Write-Host "Video URL: $videoUrl"
            Update-AgnesQuota "video" $Duration
            return $result
        } else {
            Write-Error "Video generation failed: $($result.output.task_status)"
            return $null
        }
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

function Show-AgnesHelp {
    @"
Agnes-AI (通义万相) Commands:
  Invoke-AgnesImage -Prompt "desc" [-Size 1024*1024] [-OutputFile path]
  Invoke-AgnesVideo -Prompt "desc" [-Duration 5] [-OutputFile path]
  Show-AgnesQuota                        - 查看剩余额度

Models:
  wanx-v1      - Text-to-image (async)
  wan2.6-t2v   - Text-to-video (async)
  wan2.6-i2v   - Image-to-video (async)

Usage:
  Powershell -ExecutionPolicy Bypass -Command ".
  '$env:USERPROFILE\AppData\Local\LearningEnglish\生图技能.ps1';
  Invoke-AgnesImage -Prompt '...' -OutputFile 'out.png'"
"@
}
)
