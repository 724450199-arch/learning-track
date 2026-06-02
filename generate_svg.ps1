param(
    [int]$DuoDuoWeek,
    [int]$XiaoMingWeek,
    [string]$DDLetter,
    [string]$DDWords,
    [string]$DDAct,
    [string]$XMLetter,
    [string]$XMWords,
    [string]$XMAct,
    [string]$ChineseWeek = "",
    [string]$ChinesePinyin = "",
    [string]$ChineseChars = "",
    [string]$ChineseAct = "",
    [string]$WorksheetDir
)

function New-Svg([string[]]$lines) {
    $lines -join "`n"
}

$ddWordsList = $DDWords -split ',\s*'
$ddDate = Get-Date -Format "yyyy-MM-dd"

$ddLines = @()
$ddLines = @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <rect width="595" height="842" fill="#fff"/>',
  "  <text x='50' y='60' font-size='24' font-weight='bold' fill='#333'>多多 - 牛津自然拼读 第3册 第${DuoDuoWeek}周</text>",
  "  <text x='50' y='100' font-size='18' fill='#555'>当前学习：$DDLetter</text>",
  "  <text x='50' y='130' font-size='16' fill='#777'>单词跟读：$DDWords</text>",
  '  <line x1="50" y1="145" x2="545" y2="145" stroke="#ddd"/>',
  '  <text x="50" y="180" font-size="18" font-weight="bold" fill="#333">按发音排序 (a_e)</text>',
  '  <text x="50" y="210" font-size="14" fill="#999">(在空白处填入正确的单词)</text>',
  '  <text x="50" y="250" font-size="16" fill="#333">1. c_t (猫)</text>',
  '  <text x="50" y="280" font-size="16" fill="#333">2. c_k (蛋糕)</text>',
  '  <text x="50" y="310" font-size="16" fill="#333">3. l_k (湖)</text>',
  '  <text x="50" y="340" font-size="16" fill="#333">4. m_k (制作)</text>',
  '  <text x="50" y="370" font-size="16" fill="#333">5. n_m (名字)</text>',
  '  <line x1="50" y1="390" x2="545" y2="390" stroke="#ddd"/>',
  '  <text x="50" y="420" font-size="18" font-weight="bold" fill="#333">读句子填空</text>',
  '  <text x="50" y="450" font-size="14" fill="#333">I have a red c___. (帽子/cap)</text>',
  '  <text x="50" y="475" font-size="14" fill="#333">She likes to r___ fast. (跑/run)</text>',
  '  <text x="50" y="500" font-size="14" fill="#333">The sun is h___. (热/hot)</text>',
  '  <line x1="50" y1="520" x2="545" y2="520" stroke="#ddd"/>',
  '  <text x="50" y="550" font-size="18" font-weight="bold" fill="#333">数学 - 数数</text>',
  '  <text x="50" y="580" font-size="14" fill="#333">Count the objects and write the number:</text>',
  '  <text x="50" y="610" font-size="48" fill="#333">☆☆☆☆☆ ☆☆☆☆☆ ☆☆☆</text>',
  '  <text x="50" y="650" font-size="16" fill="#333">数一数：___ 个星星</text>',
  '  <line x1="50" y1="670" x2="545" y2="670" stroke="#ddd"/>',
  "  <text x='50' y='700' font-size='14' fill='#999'>生成日期：$ddDate</text>",
  '</svg>'
)

$ddSvg = New-Svg $ddLines

# === 小铭 SVG ===
$xmWordsList = $XMWords -split ',\s*'
$xmLetters = @()
foreach ($w in $xmWordsList) {
    $xmLetters += $w.Substring(0, 1)
}
$xmLettersStr = $xmLetters -join ', '

$xmDate = Get-Date -Format "yyyy-MM-dd"

$xmLines = @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <rect width="595" height="842" fill="#fff"/>',
  "  <text x='50' y='60' font-size='24' font-weight='bold' fill='#333'>小铭 - 牛津自然拼读 第1册 第${XiaoMingWeek}周</text>",
  "  <text x='50' y='100' font-size='18' fill='#555'>当前学习：$XMLetter</text>",
  "  <text x='50' y='130' font-size='16' fill='#777'>单词：$XMWords</text>",
  '  <line x1="50" y1="145" x2="545" y2="145" stroke="#ddd"/>',
  '  <text x="50" y="180" font-size="18" font-weight="bold" fill="#333">字母描写</text>',
  "  <text x='50' y='220' font-size='80' fill='#ccc' font-family='monospace'>$XMLetter $($XMLetter.ToLower())</text>",
  '  <text x="50" y="260" font-size="16" fill="#333">(沿着灰色字母描一描)</text>',
  '  <line x1="50" y1="280" x2="545" y2="280" stroke="#ddd"/>',
  '  <text x="50" y="310" font-size="18" font-weight="bold" fill="#333">找字母</text>',
  "  <text x='50' y='340' font-size='14' fill='#333'>把字母 $XMLetter 圈出来：</text>",
  "  <text x='50' y='380' font-size='30' fill='#333' font-family='monospace'>$xmLettersStr $($xmLettersStr.ToLower()) abc def</text>",
  '  <line x1="50" y1="400" x2="545" y2="400" stroke="#ddd"/>',
  '  <text x="50" y="430" font-size="18" font-weight="bold" fill="#333">单词认读</text>',
  '  <text x="50" y="460" font-size="14" fill="#333">大声读出下面的单词：</text>',
  "  <text x='50' y='500' font-size='24' fill='#333'>$XMWords</text>",
  '  <line x1="50" y1="520" x2="545" y2="520" stroke="#ddd"/>',
  '  <text x="50" y="550" font-size="18" font-weight="bold" fill="#333">发音练习</text>',
  '  <text x="50" y="580" font-size="14" fill="#333">大声说出每个单词的首字母发音：</text>',
  "  <text x='50' y='620' font-size='20' fill='#333'>$($xmWordsList[0]) → /$($XMLetter.ToLower())/</text>",
  '  <line x1="50" y1="640" x2="545" y2="640" stroke="#ddd"/>',
  "  <text x='50' y='700' font-size='14' fill='#999'>生成日期：$xmDate</text>",
  '</svg>'
)

$xmSvg = New-Svg $xmLines

# === 语文 SVG ===
$cnDate = Get-Date -Format "yyyy-MM-dd"
$cnPinyinList = $ChinesePinyin -split '\s+'
$cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }
$cnPinyinStr = $ChinesePinyin
$cnCharsStr = $ChineseChars

$cnLines = @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <rect width="595" height="842" fill="#fff"/>',
  "  <text x='50' y='60' font-size='24' font-weight='bold' fill='#333'>多多 - 语文冲刺 第${ChineseWeek}周</text>",
  "  <text x='50' y='100' font-size='18' fill='#555'>拼音：$cnPinyinStr</text>",
  "  <text x='50' y='130' font-size='16' fill='#777'>汉字：$cnCharsStr</text>",
  '  <line x1="50" y1="145" x2="545" y2="145" stroke="#ddd"/>',
  '  <text x="50" y="180" font-size="18" font-weight="bold" fill="#333">拼音跟读</text>',
  '  <text x="50" y="210" font-size="14" fill="#999">(大声读3遍)</text>',
  "  <text x='50' y='250' font-size='24' fill='#333'>$cnPinyinStr</text>",
  '  <line x1="50" y1="280" x2="545" y2="280" stroke="#ddd"/>',
  '  <text x="50" y="310" font-size="18" font-weight="bold" fill="#333">汉字描红</text>',
  '  <text x="50" y="340" font-size="14" fill="#999">(每个字描2遍)</text>'
)

$yPos = 390
foreach ($c in $cnCharList) {
  $cnLines += "  <text x='50' y='${yPos}' font-size='28' fill='#ccc' font-family='serif'>$c $c $c $c $c</text>"
  $yPos += 50
}

$cnLines += '  <line x1="50" y1="' + $yPos + '" x2="545" y2="' + $yPos + '" stroke="#ddd"/>'
$cnLines += '  <text x="50" y="' + ($yPos + 40) + '" font-size="18" font-weight="bold" fill="#333">今日任务</text>'
$cnLines += '  <text x="50" y="' + ($yPos + 70) + '" font-size="14" fill="#333">拼音跟读3遍 | 汉字认读 | 描红练习 | 在家找字</text>'
$cnLines += '  <text x="50" y="' + ($yPos + 100) + '" font-size="14" fill="#999">活动：' + $ChineseAct + '</text>'
$cnLines += "  <text x='50' y='" + ($yPos + 140) + "' font-size='14' fill='#999'>生成日期：$cnDate</text>"
$cnLines += '</svg>'

$cnSvg = New-Svg $cnLines

# Save files
$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) {
    New-Item -ItemType Directory -Path $printableDir -Force | Out-Null
}

$ddPath = Join-Path $printableDir "duoduo_week${DuoDuoWeek}.svg"
$xmPath = Join-Path $printableDir "xiaoming_week${XiaoMingWeek}.svg"
$cnPath = Join-Path $printableDir "duoduo_chinese_week${ChineseWeek}.svg"

[System.IO.File]::WriteAllText($ddPath, $ddSvg, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($xmPath, $xmSvg, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($cnPath, $cnSvg, [System.Text.Encoding]::UTF8)

Write-Output ("SVG generated: " + $ddPath)
Write-Output ("SVG generated: " + $xmPath)
Write-Output ("SVG generated: " + $cnPath)
