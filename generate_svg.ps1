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

function New-Svg([string[]]$lines) { $lines -join "`n" }

$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) {
    New-Item -ItemType Directory -Path $printableDir -Force | Out-Null
}

$DateStr = Get-Date -Format "yyyy-MM-dd"

# ========== 多多英语 ==========
$ddWordsList = $DDWords -split ',\s*'
$ddSvg = New-Svg @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <defs>',
  '    <linearGradient id="dh" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#4A90D9"/><stop offset="100%" stop-color="#67B8F7"/></linearGradient>',
  '    <linearGradient id="ds1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8F4FD"/><stop offset="100%" stop-color="#F0F8FF"/></linearGradient>',
  '    <linearGradient id="ds2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFF8E1"/><stop offset="100%" stop-color="#FFFDE7"/></linearGradient>',
  '    <linearGradient id="ds3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8F5E9"/><stop offset="100%" stop-color="#F1F8E9"/></linearGradient>',
  '  </defs>',
  '  <rect width="595" height="842" fill="#F8FBFF"/>',
  '  <rect x="0" y="0" width="595" height="85" fill="url(#dh)" rx="0"/>',
  '  <text x="297" y="38" font-size="22" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 多多 · 牛津自然拼读 第3册</text>',
  '  <text x="297" y="65" font-size="16" fill="#E3F2FD" text-anchor="middle">第' + $DuoDuoWeek + '周 &lt;' + $DDLetter + '&gt;  ' + $DDAct + '</text>',
  '  <rect x="30" y="100" width="535" height="45" rx="8" fill="#E3F2FD" opacity="0.6"/>',
  '  <text x="45" y="128" font-size="17" fill="#1565C0" font-weight="bold">&#x1F4D6; 单词跟读：</text>',
  '  <text x="185" y="128" font-size="17" fill="#333">' + $DDWords + '</text>',
  '  <rect x="30" y="160" width="535" height="210" rx="10" fill="url(#ds1)" stroke="#BBDEFB" stroke-width="1.5"/>',
  '  <text x="297" y="188" font-size="17" font-weight="bold" fill="#1565C0" text-anchor="middle">&#x1F3B6; 按发音排序 — a_e</text>',
  '  <text x="297" y="208" font-size="12" fill="#78909C" text-anchor="middle">(在横线上填入正确的单词)</text>',
  '  <text x="60" y="245" font-size="16" fill="#333" font-family="monospace">1. c<tspan fill="#E53935" font-weight="bold">__</tspan>e (蛋糕)</text>',
  '  <text x="60" y="272" font-size="16" fill="#333" font-family="monospace">2. l<tspan fill="#E53935" font-weight="bold">__</tspan>e (湖)</text>',
  '  <text x="60" y="299" font-size="16" fill="#333" font-family="monospace">3. m<tspan fill="#E53935" font-weight="bold">__</tspan>e (制作)</text>',
  '  <text x="60" y="326" font-size="16" fill="#333" font-family="monospace">4. n<tspan fill="#E53935" font-weight="bold">__</tspan>e (名字)</text>',
  '  <text x="60" y="353" font-size="16" fill="#333" font-family="monospace">5. g<tspan fill="#E53935" font-weight="bold">__</tspan>e (游戏)</text>',
  '  <rect x="30" y="385" width="535" height="120" rx="10" fill="url(#ds2)" stroke="#FFE082" stroke-width="1.5"/>',
  '  <text x="297" y="413" font-size="17" font-weight="bold" fill="#F57F17" text-anchor="middle">&#x1F4DD; 读句子填空</text>',
  '  <text x="60" y="448" font-size="15" fill="#333" font-family="monospace">I like to r<tspan fill="#E53935" font-weight="bold">__</tspan>e. (赛跑/race)</text>',
  '  <text x="60" y="478" font-size="15" fill="#333" font-family="monospace">She has a red c<tspan fill="#E53935" font-weight="bold">__</tspan>e. (披肩/cape)</text>',
  '  <text x="60" y="508" font-size="15" fill="#333" font-family="monospace">We m<tspan fill="#E53935" font-weight="bold">__</tspan>e a cake. (制作/make)</text>',
  '  <rect x="30" y="520" width="535" height="210" rx="10" fill="url(#ds3)" stroke="#A5D6A7" stroke-width="1.5"/>',
  '  <text x="297" y="548" font-size="17" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x1F522; 数学 · Numbers to 10</text>',
  '  <text x="297" y="568" font-size="12" fill="#558B2F" text-anchor="middle">数一数，在横线上写出数字</text>',
  '  <text x="70" y="605" font-size="18" fill="#333" font-family="monospace">&#x2605; &#x2605; &#x2605; &#x2605; &#x2605; = </text><text x="250" y="605" font-size="18" fill="#E53935" font-family="monospace" font-weight="bold">__</text>',
  '  <text x="320" y="605" font-size="18" fill="#333" font-family="monospace">&#x2605; &#x2605; &#x2605; = </text><text x="458" y="605" font-size="18" fill="#E53935" font-family="monospace" font-weight="bold">__</text>',
  '  <text x="70" y="640" font-size="18" fill="#333" font-family="monospace">&#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; = </text><text x="310" y="640" font-size="18" fill="#E53935" font-family="monospace" font-weight="bold">__</text>',
  '  <text x="70" y="675" font-size="18" fill="#333" font-family="monospace">&#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; = </text><text x="330" y="675" font-size="18" fill="#E53935" font-family="monospace" font-weight="bold">__</text>',
  '  <text x="70" y="710" font-size="18" fill="#333" font-family="monospace">圈出较大的数：  8    3</text>',
  '  <text x="70" y="730" font-size="18" fill="#333" font-family="monospace">圈出较大的数：  5    7</text>',
  '</svg>'
)

# ========== 小铭英语 ==========
$xmWordsList = $XMWords -split ',\s*'
$xmLetters = @()
foreach ($w in $xmWordsList) { $xmLetters += $w.Substring(0, 1) }
$xmLettersStr = $xmLetters -join ', '
$xmLetterDisplay = $XMLetter.Substring(0,1) + " " + $XMLetter.Substring(1)

$xmLetterUp = $XMLetter.Substring(0,1).ToUpper()
$xmLetterLow = if ($xmLetterUp -eq 'A') { "&#x0251;" } else { $XMLetter.Substring(0,1).ToLower() }

$xmSvg = New-Svg @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <defs>',
  '    <linearGradient id="xh" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#43A047"/><stop offset="100%" stop-color="#66BB6A"/></linearGradient>',
  '    <linearGradient id="xs1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8F5E9"/><stop offset="100%" stop-color="#F1F8E9"/></linearGradient>',
  '    <linearGradient id="xs2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FCE4EC"/><stop offset="100%" stop-color="#FFF0F3"/></linearGradient>',
  '    <linearGradient id="xs3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8EAF6"/><stop offset="100%" stop-color="#F3F4FD"/></linearGradient>',
  '  </defs>',
  '  <rect width="595" height="842" fill="#F9FFF9"/>',
  '  <rect x="0" y="0" width="595" height="85" fill="url(#xh)" rx="0"/>',
  '  <text x="297" y="38" font-size="22" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 小铭 · 牛津自然拼读 第1册</text>',
  '  <text x="297" y="65" font-size="16" fill="#E8F5E9" text-anchor="middle">第' + $XiaoMingWeek + '周 ' + $XMLetter + '  ' + $XMAct + '</text>',
  '  <rect x="30" y="100" width="535" height="45" rx="8" fill="#E8F5E9" opacity="0.6"/>',
  '  <text x="45" y="128" font-size="17" fill="#2E7D32" font-weight="bold">&#x1F4D6; 单词：</text>',
  '  <text x="145" y="128" font-size="17" fill="#333">' + $XMWords + '</text>',
  '  <rect x="30" y="160" width="535" height="120" rx="10" fill="url(#xs1)" stroke="#A5D6A7" stroke-width="1.5"/>',
  '  <text x="297" y="188" font-size="17" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x270F;&#xFE0F; 字母描写</text>',
  '  <text x="297" y="205" font-size="12" fill="#78909C" text-anchor="middle">先看示例（黑色），再描灰色字母</text>',
  '  <text x="200" y="248" font-size="50" fill="#333" text-anchor="middle" font-family="monospace" font-weight="bold">' + $xmLetterUp + '</text>',
  '  <text x="270" y="248" font-size="50" fill="#CFD8DC" text-anchor="middle" font-family="monospace" font-weight="bold">' + $xmLetterUp + ' ' + $xmLetterUp + ' ' + $xmLetterUp + ' ' + $xmLetterUp + '</text>',
  '  <text x="200" y="272" font-size="50" fill="#333" text-anchor="middle" font-family="monospace" font-weight="bold">' + $xmLetterLow + '</text>',
  '  <text x="270" y="272" font-size="50" fill="#CFD8DC" text-anchor="middle" font-family="monospace" font-weight="bold">' + $xmLetterLow + ' ' + $xmLetterLow + ' ' + $xmLetterLow + ' ' + $xmLetterLow + '</text>',
  '  <rect x="30" y="275" width="535" height="90" rx="10" fill="url(#xs2)" stroke="#F48FB1" stroke-width="1.5"/>',
  '  <text x="297" y="303" font-size="17" font-weight="bold" fill="#C62828" text-anchor="middle">&#x1F50D; 找字母</text>',
  '  <text x="297" y="320" font-size="13" fill="#78909C" text-anchor="middle">把字母 ' + $XMLetter.Substring(0,1) + ' 圈出来：</text>',
  '  <text x="297" y="355" font-size="28" fill="#333" text-anchor="middle" font-family="monospace">a a a  a a a  b c d e f g h</text>',
  '  <rect x="30" y="380" width="535" height="100" rx="10" fill="url(#xs3)" stroke="#9FA8DA" stroke-width="1.5"/>',
  '  <text x="297" y="408" font-size="17" font-weight="bold" fill="#283593" text-anchor="middle">&#x1F4E3; 单词认读</text>',
  '  <text x="297" y="425" font-size="13" fill="#78909C" text-anchor="middle">大声读出下面的单词：</text>',
  '  <text x="297" y="468" font-size="22" fill="#333" text-anchor="middle" font-weight="bold">' + $XMWords + '</text>',
  '  <rect x="30" y="495" width="535" height="155" rx="10" fill="url(#xs2)" stroke="#F48FB1" stroke-width="1.5"/>',
  '  <text x="297" y="523" font-size="17" font-weight="bold" fill="#C62828" text-anchor="middle">&#x1F399;&#xFE0F; 发音练习 · 首字母发音</text>',
  '  <text x="297" y="540" font-size="13" fill="#78909C" text-anchor="middle">大声说出每个单词的首字母发音，再读整词：</text>',
  '  <text x="297" y="575" font-size="20" fill="#333" text-anchor="middle">' + $xmWordsList[0] + ' → /' + $XMLetter.Substring(0,1).ToLower() + '/</text>',
  '  <text x="297" y="605" font-size="20" fill="#333" text-anchor="middle">' + $xmWordsList[1] + ' → /' + $XMLetter.Substring(0,1).ToLower() + '/</text>',
  '  <text x="297" y="635" font-size="20" fill="#333" text-anchor="middle">' + $xmWordsList[2] + ' → /' + $XMLetter.Substring(0,1).ToLower() + '/</text>',
  '  <text x="297" y="700" font-size="12" fill="#B0BEC5" text-anchor="middle">生成日期：' + $DateStr + '</text>',
  '</svg>'
)

# ========== 语文 SVG ==========
$cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }

$cnSvgLines = @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <defs>',
  '    <linearGradient id="ch" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#E53935"/><stop offset="100%" stop-color="#FF7043"/></linearGradient>',
  '    <linearGradient id="cs1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFF3E0"/><stop offset="100%" stop-color="#FFF8E1"/></linearGradient>',
  '    <linearGradient id="cs2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FBE9E7"/><stop offset="100%" stop-color="#FFF0E8"/></linearGradient>',
  '  </defs>',
  '  <rect width="595" height="842" fill="#FFFBF5"/>',
  '  <rect x="0" y="0" width="595" height="85" fill="url(#ch)" rx="0"/>',
  '  <text x="297" y="38" font-size="22" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 多多 · 语文冲刺</text>',
  '  <text x="297" y="65" font-size="16" fill="#FFEBEE" text-anchor="middle">第' + $ChineseWeek + '周 · 9月入学准备</text>',
  '  <rect x="30" y="100" width="535" height="45" rx="8" fill="#FFEBEE" opacity="0.6"/>',
  '  <text x="45" y="128" font-size="17" fill="#C62828" font-weight="bold">&#x1F4D6; 拼音：</text>',
  '  <text x="145" y="128" font-size="17" fill="#333">' + $ChinesePinyin + '</text>',
  '  <rect x="30" y="160" width="535" height="105" rx="10" fill="url(#cs1)" stroke="#FFCC80" stroke-width="1.5"/>',
  '  <text x="297" y="188" font-size="17" font-weight="bold" fill="#E65100" text-anchor="middle">&#x1F3A4; 拼音跟读 · 大声读3遍</text>',
  '  <text x="297" y="250" font-size="28" fill="#333" text-anchor="middle" font-weight="bold">' + $ChinesePinyin + '</text>',
  '  <rect x="30" y="280" width="535" height="340" rx="10" fill="url(#cs2)" stroke="#FFAB91" stroke-width="1.5"/>',
  '  <text x="297" y="308" font-size="17" font-weight="bold" fill="#BF360C" text-anchor="middle">&#x270F;&#xFE0F; 汉字描红 · 每个字描2遍</text>'
)

$yPos = 345
foreach ($c in $cnCharList) {
  $cnSvgLines += "  <text x='297' y='${yPos}' font-size='30' fill='#CFD8DC' text-anchor='middle' font-family='STKaiti,serif' font-weight='bold'>$c  $c  $c  $c  $c</text>"
  $yPos += 48
}

$cnSvgLines += '  <rect x="30" y="' + ($yPos + 5) + '" width="535" height="55" rx="10" fill="#FFCC02" opacity="0.2"/>'
$cnSvgLines += '  <text x="297" y="' + ($yPos + 38) + '" font-size="15" fill="#E65100" text-anchor="middle">&#x1F3C6; 今日任务：跟读3遍 | 认读 | 描红 | 在家找字</text>'
$cnSvgLines += '  <text x="297" y="' + ($yPos + 70) + '" font-size="12" fill="#B0BEC5" text-anchor="middle">活动：' + $ChineseAct + '</text>'
$cnSvgLines += '  <text x="297" y="' + ($yPos + 100) + '" font-size="12" fill="#B0BEC5" text-anchor="middle">生成日期：' + $DateStr + '</text>'
$cnSvgLines += '</svg>'

$cnSvg = New-Svg $cnSvgLines

# ====== Save ======
$ddPath = Join-Path $printableDir "duoduo_week${DuoDuoWeek}.svg"
$xmPath = Join-Path $printableDir "xiaoming_week${XiaoMingWeek}.svg"
$cnPath = Join-Path $printableDir "duoduo_chinese_week${ChineseWeek}.svg"

[System.IO.File]::WriteAllText($ddPath, $ddSvg, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($xmPath, $xmSvg, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($cnPath, $cnSvg, [System.Text.Encoding]::UTF8)

Write-Output ("SVG: " + $ddPath)
Write-Output ("SVG: " + $xmPath)
Write-Output ("SVG: " + $cnPath)
