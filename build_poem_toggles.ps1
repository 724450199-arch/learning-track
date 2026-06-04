# Build nested toggle structure for 65 poems (13 weeks x 5 poems)
$poemsPath = "C:\Users\yang\AppData\Local\Temp\opencode\poems.json"
$poemsJson = [System.IO.File]::ReadAllText($poemsPath, [System.Text.Encoding]::UTF8)
$poems = $poemsJson | ConvertFrom-Json

$children = @()

# Group by week
$weeks = $poems | Group-Object week

foreach ($wg in $weeks) {
    $w = [int]$wg.Name
    $pList = $wg.Group

    # Build week-level toggle
    $poemTitles = ($pList | ForEach-Object { $_.title }) -join "、"
    $weekToggleLabel = "第${w}周 — $poemTitles"

    $poemToggles = @()
    foreach ($p in $pList) {
        $poemText = "$($p.author)：「$($p.text)」"
        $poemToggle = @{
            object = "block"
            type = "toggle"
            toggle = @{
                rich_text = @(
                    @{
                        type = "text"
                        text = @{ content = $p.title; link = $null }
                        annotations = @{ bold = $false; italic = $false; strikethrough = $false; underline = $false; code = $false; color = "default" }
                        plain_text = $p.title
                        href = $null
                    }
                )
                color = "default"
            }
            children = @(
                @{
                    object = "block"
                    type = "paragraph"
                    paragraph = @{
                        rich_text = @(
                            @{
                                type = "text"
                                text = @{ content = $poemText; link = $null }
                                annotations = @{ bold = $false; italic = $true; strikethrough = $false; underline = $false; code = $false; color = "gray" }
                                plain_text = $poemText
                                href = $null
                            }
                        )
                        color = "default"
                    }
                }
            )
        }
        $poemToggles += $poemToggle
    }

    $children += @{
        object = "block"
        type = "toggle"
        toggle = @{
            rich_text = @(
                @{
                    type = "text"
                    text = @{ content = "📜 $weekToggleLabel"; link = $null }
                    annotations = @{ bold = $true; italic = $false; strikethrough = $false; underline = $false; code = $false; color = "default" }
                    plain_text = "📜 $weekToggleLabel"
                    href = $null
                }
            )
            color = "default"
        }
        children = $poemToggles
    }
}

$payload = @{ children = $children }
$json = $payload | ConvertTo-Json -Depth 10 -Compress
[System.IO.File]::WriteAllText("C:\Users\yang\AppData\Local\Temp\opencode\poem_payload.json", $json, [System.Text.Encoding]::UTF8)
Write-Output "Payload written: $($json.Length) bytes"
