$token = "zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj"
$duoduoPageId = "b80cd768-6ef5-4da6-a257-e2afe8d388ac"
$xiaomingPageId = "8169ce9a-efd8-47bd-b671-2d0b8fa4c6e6"
$baseUrl = "https://mcp.flowus.cn/v2"

function Call-FlowUs($method, $path, $body) {
    $url = $baseUrl + $path
    $params = @{
        Uri = $url
        Method = $method
        ContentType = "application/json; charset=utf-8"
        Headers = @{ "Authorization" = "Bearer $token" }
        ErrorAction = "Stop"
    }
    if ($body) {
        $params.Body = ($body | ConvertTo-Json -Depth 10 -Compress)
    }
    try {
        $resp = Invoke-RestMethod @params
        return $resp
    } catch {
        Write-Host ("ERROR: " + $_.Exception.Message)
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            Write-Host ("Response: " + $reader.ReadToEnd())
        }
        throw
    }
}

# Step 1: Get upload URL for duoduo docx
Write-Host "Getting upload URL for duoduo..."
$uploadReq1 = @{
    filename = "多多_第1周_练习题.docx"
    content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    content_length = 84534
    parent = @{
        type = "page_id"
        page_id = $duoduoPageId
    }
}
$uploadResp1 = Call-FlowUs -method POST -path "/files/upload" -body $uploadReq1
Write-Host ($uploadResp1 | ConvertTo-Json -Depth 5)

# Step 2: Upload the file
Write-Host "Uploading duoduo docx..."
$fileBytes = [System.IO.File]::ReadAllBytes("C:\Users\yang\AppData\Local\LearningEnglish\worksheets\printable\duoduo_week1.docx")
Invoke-RestMethod -Uri $uploadResp1.upload_url -Method $uploadResp1.method -Body $fileBytes -ContentType "application/vnd.openxmlformats-officedocument.wordprocessingml.document" -ErrorAction Stop | Out-Null
Write-Host "Uploaded!"
