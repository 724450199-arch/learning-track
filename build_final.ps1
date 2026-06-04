$poemsPath = "C:\Users\yang\AppData\Local\Temp\opencode\poems.json"
$poemsJson = [System.IO.File]::ReadAllText($poemsPath, [System.Text.Encoding]::UTF8)
$poems = $poemsJson | ConvertFrom-Json
$weeks = $poems | Group-Object week

$children = @()
foreach ($wg in $weeks) {
    $w = [int]$wg.Name
    $pList = $wg.Group
    $poemTitles = ($pList | ForEach-Object { $_.title }) -join "、"

    $paragraphs = @()
    foreach ($p in $pList) {
        $text = "$($p.title) — $($p.author)：「$($p.text)」"
        $paragraphs += @{
            object = "block"
            type = "paragraph"
            paragraph = @{
                rich_text = @(
                    @{
                        type = "text"
                        text = @{ content = $text; link = $null }
                        annotations = @{ bold = $false; italic = $false; strikethrough = $false; underline = $false; code = $false; color = "default" }
                        plain_text = $text
                        href = $null
                    }
                )
                color = "default"
            }
        }
    }

    $children += @{
        object = "block"
        type = "toggle"
        toggle = @{
            rich_text = @(
                @{
                    type = "text"
                    text = @{ content = "📜 第${w}周 — $poemTitles"; link = $null }
                    annotations = @{ bold = $true; italic = $false; strikethrough = $false; underline = $false; code = $false; color = "default" }
                    plain_text = "📜 第${w}周 — $poemTitles"
                    href = $null
                }
            )
            color = "default"
        }
        children = $paragraphs
    }
}

$payload = @{ children = $children; after = "8f17cde2-7bc5-4e18-88dd-7c93c99791f1" }
$json = $payload | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText("C:\Users\yang\AppData\Local\Temp\opencode\poem_final.json", $json, [System.Text.Encoding]::UTF8)
Write-Output "OK: $($children.Count) weeks, payload size: $($json.Length) bytes"
