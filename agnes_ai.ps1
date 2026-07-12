@(
' Agnes AI API Helper (https://agnes-ai.com)'
' =========================================='
' Models: agnes-image-2.0-flash (image), agnes-video-v2.0 (video)'
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

# -------- Image Generation (sync, OpenAI-compatible) -------- #

function Invoke-AgnesImage {
    param(
        [string]$Prompt,
        [string]$Model = "agnes-image-2.0-flash",
        [string]$Size = "1024x1024",
        [string]$OutputFile
    )

    $key = Get-AgnesApiKey
    if (-not $key) { return }

    $body = @{
        model = $Model
        prompt = $Prompt
        size = $Size
        extra_body = @{
            response_format = "url"
        }
    }

    $json = $body | ConvertTo-Json -Depth 3
    $headers = @{
        Authorization = "Bearer $key"
        "Content-Type" = "application/json"
    }

    try {
        Write-Host "Generating image with Agnes AI..."
        Write-Host "Model: $Model | Size: $Size"
        $resp = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/v1/images/generations" `
            -Method Post -Body $json -ContentType "application/json" -Headers $headers

        if ($resp.data -and $resp.data[0].url) {
            $url = $resp.data[0].url
            Write-Host "Image URL: $url"
            if ($OutputFile) {
                Invoke-WebRequest -Uri $url -OutFile $OutputFile
                Write-Host "Saved: $OutputFile"
            }
            return $resp
        } else {
            Write-Error "No image URL in response"
            return $null
        }
    } catch {
        Write-Error "API error: $_"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            Write-Host ("Response: " + $reader.ReadToEnd())
        }
        return $null
    }
}

# -------- Video Generation (async) -------- #

function Invoke-AgnesVideo {
    param(
        [string]$Prompt,
        [string]$Model = "agnes-video-v2.0",
        [int]$Width = 1152,
        [int]$Height = 768,
        [int]$NumFrames = 121,
        [int]$FrameRate = 24,
        [string]$InputImage,  # optional image-to-video
        [string]$OutputFile,
        [int]$PollInterval = 10
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

    $json = $body | ConvertTo-Json -Depth 3
    $headers = @{
        Authorization = "Bearer $key"
        "Content-Type" = "application/json"
    }

    try {
        Write-Host "Creating video task with Agnes AI..."
        $task = Invoke-RestMethod -Uri "https://apihub.agnes-ai.com/v1/videos" `
            -Method Post -Body $json -ContentType "application/json" -Headers $headers

        $videoId = $task.video_id
        $taskId = $task.task_id
        Write-Host "Video ID: $videoId | Task ID: $taskId"

        # Poll for result
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
            Write-Error "Video generation failed: $($result.error)"
            return $null
        }
    } catch {
        Write-Error "API error: $_"
        return $null
    }
}

function Show-AgnesHelp {
    @"
Agnes AI Commands:
  Invoke-AgnesImage -Prompt "desc" [-Size 1024x1024] [-OutputFile path]
  Invoke-AgnesVideo -Prompt "desc" [-Width 1152] [-Height 768] [-NumFrames 121] [-FrameRate 24] [-InputImage url] [-OutputFile path]

Models:
  agnes-image-2.0-flash  - Text-to-image (sync)
  agnes-video-v2.0       - Text/Image-to-video (async)

Setup:
  \$env:AGNES_API_KEY = "sk-..."   # or save to ~\.agnes\api_key
"@
}
)
