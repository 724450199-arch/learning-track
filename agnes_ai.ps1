@(
' Agnes AI API Helper (Sapiens AI)'
' =================================='
' Models: agnes-2.0-flash (chat), agnes-image-2.0-flash (image), agnes-video-v2.0 (video)'
' API Base: https://apihub.agnes-ai.com/v1'
''

# -------- API Key -------- #

function Get-AgnesApiKey {
    $key = $env:AGNES_API_KEY
    if (-not $key) {
        $keyFile = Join-Path (Join-Path $env:USERPROFILE ".agnes") "api_key"
        if (Test-Path $keyFile) {
            $key = [System.IO.File]::ReadAllText($keyFile, [Text.Encoding]::UTF8).Trim()
        }
    }
    if (-not $key) {
        Write-Error "AGNES_API_KEY not found. Set env var or save to ~/.agnes/api_key"
        return $null
    }
    return $key
}

# -------- Chat Completion (OpenAI compatible) -------- #

function Invoke-AgnesChat {
    param(
        [string]$Prompt,
        [string]$Model = "agnes-2.0-flash",
        [string]$SystemPrompt = "You are a helpful assistant.",
        [int]$MaxTokens = 4096,
        [float]$Temperature = 0.7
    )

    $key = Get-AgnesApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        messages = @(
            @{ role = "system"; content = $SystemPrompt }
            @{ role = "user"; content = $Prompt }
        )
        max_tokens = $MaxTokens
        temperature = $Temperature
    }

    try {
        $resp = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/v1/chat/completions" `
            -Method Post `
            -Body ($body | ConvertTo-Json -Depth 5) `
            -ContentType "application/json" `
            -Headers @{ Authorization = "Bearer $key" }

        $content = $resp.choices[0].message.content
        Write-Host $content
        return $content
    } catch {
        Write-Error "API error: $_"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            Write-Host "Response: $($reader.ReadToEnd())"
        }
        return $null
    }
}

# -------- Image Generation -------- #

function Invoke-AgnesImage {
    param(
        [string]$Prompt,
        [string]$Model = "agnes-image-2.0-flash",
        [string]$Size = "1024x1024",
        [string]$OutputFile,
        [string]$InputImage,
        [string]$ResponseFormat = "url"
    )

    $key = Get-AgnesApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        prompt = $Prompt
        size = $Size
        extra_body = @{ response_format = $ResponseFormat }
    }

    if ($InputImage) {
        $body.extra_body.image = @($InputImage)
    }

    try {
        Write-Host "Generating image..."
        $resp = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/v1/images/generations" `
            -Method Post `
            -Body ($body | ConvertTo-Json -Depth 5) `
            -ContentType "application/json" `
            -Headers @{ Authorization = "Bearer $key" }

        if ($resp.data -and $resp.data[0]) {
            $url = $resp.data[0].url
            $b64 = $resp.data[0].b64_json
            if ($url) {
                Write-Host "Image URL: $url"
                if ($OutputFile) {
                    Invoke-WebRequest -Uri $url -OutFile $OutputFile
                    Write-Host "Saved: $OutputFile"
                }
                return $url
            } elseif ($b64) {
                Write-Host "Image received as Base64"
                if ($OutputFile) {
                    $bytes = [Convert]::FromBase64String($b64)
                    [System.IO.File]::WriteAllBytes($OutputFile, $bytes)
                    Write-Host "Saved: $OutputFile"
                }
                return $b64
            }
        }
        return $resp
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

# -------- Video Generation (Async) -------- #

function Invoke-AgnesVideo {
    param(
        [string]$Prompt,
        [string]$Model = "agnes-video-v2.0",
        [int]$Width = 1152,
        [int]$Height = 768,
        [int]$NumFrames = 121,
        [int]$FrameRate = 24,
        [string]$InputImage,
        [string]$OutputFile,
        [int]$PollInterval = 15
    )

    $key = Get-AgnesApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        prompt = $Prompt
        width = $Width
        height = $Height
        num_frames = $NumFrames
        frame_rate = $FrameRate
    }

    if ($InputImage) {
        $body.image = $InputImage
    }

    try {
        Write-Host "Creating video task..."
        $task = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/v1/videos" `
            -Method Post `
            -Body ($body | ConvertTo-Json -Depth 5) `
            -ContentType "application/json" `
            -Headers @{ Authorization = "Bearer $key" }

        $videoId = $task.video_id
        Write-Host "Video ID: $videoId (Status: $($task.status))"

        do {
            Start-Sleep -Seconds $PollInterval
            $result = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/agnesapi?video_id=$videoId" `
                -Headers @{ Authorization = "Bearer $key" }
            Write-Host "Status: $($result.status) ($($result.progress)%)"
        } while ($result.status -eq "queued" -or $result.status -eq "in_progress")

        if ($result.status -eq "completed" -and $result.url) {
            Write-Host "Video URL: $($result.url)"
            if ($OutputFile) {
                Invoke-WebRequest -Uri $result.url -OutFile $OutputFile
                Write-Host "Saved: $OutputFile"
            }
            return $result
        } else {
            Write-Error "Video generation failed: $($result.status) $($result.error)"
            return $null
        }
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

# -------- Check Video Status -------- #

function Get-AgnesVideoStatus {
    param([string]$VideoId)

    $key = Get-AgnesApiKey
    if (-not $key) { return }

    try {
        $result = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/agnesapi?video_id=$VideoId" `
            -Headers @{ Authorization = "Bearer $key" }
        Write-Host "Status: $($result.status) ($($result.progress)%)"
        if ($result.status -eq "completed") {
            Write-Host "Video URL: $($result.url)"
            Write-Host "Duration: $($result.seconds)s, Size: $($result.size)"
        }
        return $result
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

# -------- Help -------- #

function Show-AgnesHelp {
    @"
Agnes AI (Sapiens AI) Commands:
  Invoke-AgnesChat -Prompt "hello"      - AI 对话
  Invoke-AgnesImage -Prompt "desc" [-Size 1024x1024] [-OutputFile path]  - 生成图片
  Invoke-AgnesVideo -Prompt "desc" [-Duration 5] [-OutputFile path]      - 生成视频
  Get-AgnesVideoStatus -VideoId "..."   - 查询视频生成状态

Chat Models:
  agnes-2.0-flash       - 通用对话/推理

Image Models:
  agnes-image-2.0-flash - 文生图/图生图
  agnes-image-2.1-flash - 高密度视觉生成

Video Models:
  agnes-video-v2.0      - 文生视频/图生视频

API Base: https://apihub.agnes-ai.com/v1
Auth: AGNES_API_KEY 环境变量 或 ~/.agnes/api_key 文件

Usage:
  Powershell -ExecutionPolicy Bypass -Command ".
  '$env:USERPROFILE\AppData\Local\LearningEnglish\agnes_ai.ps1';
  Invoke-AgnesImage -Prompt 'a cute cat' -OutputFile 'cat.png'"
"@
}
)
