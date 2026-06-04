$path = "C:\Users\yang\AppData\Local\LearningEnglish\daily_learning.ps1"
$content = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)

$lines = $content -split "`n"
$inPoems = $false
$weekNum = 0
$poems = @()

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    if ($line -match '\$ChineseContent\[(\d+)\]') {
        $weekNum = [int]$Matches[1]
        $inPoems = $false
    }
    if ($inPoems -and $line -match '@{title="(.*?)".*?author="(.*?)".*?text="(.*?)"}') {
        $poems += @{week=$weekNum; title=$Matches[1]; author=$Matches[2]; text=$Matches[3]}
    }
    if ($line -match 'poems=@\(') {
        $inPoems = $true
    }
}

$poems | Format-Table -AutoSize
