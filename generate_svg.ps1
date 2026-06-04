param(
  [string]$DuoDuoWeek = "1",
  [string]$XiaoMingWeek = "1",
  [string]$DDLetter = "",
  [string]$DDWords = "",
  [string]$DDAct = "",
  [string]$XMLetter = "",
  [string]$XMWords = "",
  [string]$XMAct = "",
  [string]$ChineseWeek = "",
  [string]$ChinesePinyin = "",
  [string]$ChineseChars = "",
  [string]$ChineseAct = "",
  [string]$ChinesePoem1Title = "",
  [string]$ChinesePoem1Author = "",
  [string]$ChinesePoem1Text = "",
  [string]$ChinesePoem2Title = "",
  [string]$ChinesePoem2Author = "",
  [string]$ChinesePoem2Text = "",
  [string]$DDMathWeek = "1",
  [string]$WorksheetDir = ""
)

function New-Svg([string[]]$lines) { $lines -join "`n" }

$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) {
    New-Item -ItemType Directory -Path $printableDir -Force | Out-Null
}

$DateStr = Get-Date -Format "yyyy-MM-dd"

# ========== 多多英语 ==========
$ddWordsList = $DDWords -split ',\s*'

# a_e 单词释义表
$ddMeanings = @{
  cake="蛋糕"; name="名字"; game="游戏"; make="制作"
  late="迟的"; snake="蛇"; plane="飞机"; plate="盘子"
  lake="湖"; race="赛跑"; cape="披肩"; kite="风筝"
  bike="自行车"; ride="骑"; five="五"; home="家"
  nose="鼻子"; rope="绳子"; hope="希望"; cube="立方体"
  tube="管子"; cute="可爱的"; mule="骡子"
}

# 为每个 a_e 单词生成填空练习行
$ddExLines = @()
$ddIdx = 1
$ddY = 268
foreach ($w in $ddWordsList) {
  $w = $w.Trim().ToLower()
  if (-not $w -or $w -eq '-') { continue }
  $pos = $w.IndexOf('a')
  if ($pos -ge 0) {
    $prefix = $w.Substring(0, $pos)
    $suffix = $w.Substring($pos + 1)
    $meaning = if ($ddMeanings.ContainsKey($w)) { $ddMeanings[$w] } else { $w }
    $ddExLines += "  <text x='60' y='${ddY}' font-size='26' fill='#333' font-family='monospace'>${ddIdx}. ${prefix}<tspan fill='#E53935' font-weight='bold'>__</tspan>${suffix} ($meaning)</text>"
    $ddIdx++
    $ddY += 38
  }
}
# 如果当前单词不够填空，补一些通用的
if ($ddIdx -le 2) {
  $ddExLines += "  <text x='60' y='${ddY}' font-size='26' fill='#333' font-family='monospace'>${ddIdx}. c<tspan fill='#E53935' font-weight='bold'>__</tspan>ke (蛋糕)</text>"; $ddIdx++; $ddY += 38
  $ddExLines += "  <text x='60' y='${ddY}' font-size='26' fill='#333' font-family='monospace'>${ddIdx}. n<tspan fill='#E53935' font-weight='bold'>__</tspan>me (名字)</text>"; $ddIdx++; $ddY += 38
}

$ddExStr = $ddExLines -join "`n"

# 为单词生成句子练习
$ddSentences = @{
  late="The train is l__te. (火车晚点)"; snake="A sn__ke is long. (蛇很长)"
  plane="The pl__ne can fly. (飞机能飞)"; plate="Put it on the pl__te. (放盘子里)"
  cake="We m__ke a cake. (做蛋糕)"; name="My n__me is Tom. (名字)"
  game="Let's play a g__me. (玩游戏)"; make="I c__n make it. (我能做)"
  lake="Let's go to the l__ke. (去湖边)"; race="I like to r__ce. (赛跑)"
  cape="She has a red c__pe. (披肩)"
}
$ddSentY = 500
$ddSentLines = @()
$ddSIdx = 1
foreach ($w in $ddWordsList) {
  $w = $w.Trim().ToLower()
  if (-not $w -or $w -eq '-') { continue }
  if ($ddSentences.ContainsKey($w)) {
    $ddSentLines += "  <text x='60' y='${ddSentY}' font-size='24' fill='#333' font-family='monospace'>$($ddSIdx). $($ddSentences[$w])</text>"
    $ddSIdx++
    $ddSentY += 40
    if ($ddSIdx -gt 4) { break }
  }
}
$ddSentStr = $ddSentLines -join "`n"

# ====== 生成数学内容（按周变化） ======
$mathLines = @()
if ([int]$DDMathWeek -le 2) {
  if ([int]$DDMathWeek -eq 1) {
    $mathLines += '  <text x="297" y="630" font-size="20" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x1F522; 数学 &#xB7; Numbers to 10</text>'
    $mathLines += '  <text x="297" y="652" font-size="14" fill="#558B2F" text-anchor="middle">数一数，在横线上写出数字</text>'
    $mathLines += '  <text x="70" y="690" font-size="22" fill="#333" font-family="monospace">&#x2605; &#x2605; &#x2605; &#x2605; &#x2605; = </text><text x="260" y="690" font-size="22" fill="#E53935" font-family="monospace" font-weight="bold">__</text>'
    $mathLines += '  <text x="330" y="690" font-size="22" fill="#333" font-family="monospace">&#x2605; &#x2605; &#x2605; = </text><text x="470" y="690" font-size="22" fill="#E53935" font-family="monospace" font-weight="bold">__</text>'
    $mathLines += '  <text x="70" y="728" font-size="22" fill="#333" font-family="monospace">&#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; &#x25CF; = </text><text x="320" y="728" font-size="22" fill="#E53935" font-family="monospace" font-weight="bold">__</text>'
    $mathLines += '  <text x="70" y="766" font-size="22" fill="#333" font-family="monospace">&#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; &#x25A3; = </text><text x="340" y="766" font-size="22" fill="#E53935" font-family="monospace" font-weight="bold">__</text>'
    $mathLines += '  <text x="70" y="804" font-size="22" fill="#333" font-family="monospace">圈出较大的数：  8    3</text>'
    $mathLines += '  <text x="70" y="824" font-size="22" fill="#333" font-family="monospace">圈出较大的数：  5    7</text>'
  } else {
    $mathLines += '  <text x="297" y="668" font-size="24" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x1F522; 数学 &#xB7; Numbers to 10</text>'
    $mathLines += '  <text x="297" y="694" font-size="16" fill="#558B2F" text-anchor="middle">写出缺少的数字</text>'
    $mathLines += '  <text x="70" y="740" font-size="30" fill="#333" font-family="monospace">1, 2, <tspan fill="#E53935" font-weight="bold">__</tspan>, 4, 5</text>'
    $mathLines += '  <text x="70" y="782" font-size="30" fill="#333" font-family="monospace">5, <tspan fill="#E53935" font-weight="bold">__</tspan>, 7, 8, 9</text>'
    $mathLines += '  <text x="70" y="824" font-size="30" fill="#333" font-family="monospace">8, 7, 6, <tspan fill="#E53935" font-weight="bold">__</tspan>, 4</text>'
  }
} elseif ([int]$DDMathWeek -le 4) {
  $mathLines += '  <text x="297" y="630" font-size="20" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x1F522; 数学 &#xB7; Number Bonds</text>'
  $mathLines += '  <text x="297" y="652" font-size="14" fill="#558B2F" text-anchor="middle">10可以分成几和几？填空</text>'
  $mathLines += '  <text x="70" y="695" font-size="24" fill="#333" font-family="monospace">10 = <tspan fill="#E53935" font-weight="bold">__</tspan> + 3</text>'
  $mathLines += '  <text x="70" y="735" font-size="24" fill="#333" font-family="monospace">10 = 6 + <tspan fill="#E53935" font-weight="bold">__</tspan></text>'
  $mathLines += '  <text x="70" y="775" font-size="24" fill="#333" font-family="monospace">8 = 5 + <tspan fill="#E53935" font-weight="bold">__</tspan></text>'
  $mathLines += '  <text x="70" y="815" font-size="24" fill="#333" font-family="monospace">9 = <tspan fill="#E53935" font-weight="bold">__</tspan> + 4</text>'
} else {
  $mathLines += '  <text x="297" y="630" font-size="20" font-weight="bold" fill="#2E7D32" text-anchor="middle">&#x1F522; 数学</text>'
  $mathLines += '  <text x="297" y="680" font-size="20" fill="#558B2F" text-anchor="middle">数学练习</text>'
}
$ddMathStr = $mathLines -join "`n"

$ddSvg = New-Svg @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <defs>',
  '    <linearGradient id="dh" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#4A90D9"/><stop offset="100%" stop-color="#67B8F7"/></linearGradient>',
  '    <linearGradient id="ds1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8F4FD"/><stop offset="100%" stop-color="#F0F8FF"/></linearGradient>',
  '    <linearGradient id="ds2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFF8E1"/><stop offset="100%" stop-color="#FFFDE7"/></linearGradient>',
  '    <linearGradient id="ds3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8F5E9"/><stop offset="100%" stop-color="#F1F8E9"/></linearGradient>',
  '  </defs>',
  '  <rect width="595" height="842" fill="#F8FBFF"/>',
  '  <rect x="0" y="0" width="595" height="95" fill="url(#dh)" rx="0"/>',
  '  <text x="297" y="42" font-size="32" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 多多 · 牛津自然拼读 第3册</text>',
  '  <text x="297" y="74" font-size="20" fill="#E3F2FD" text-anchor="middle">第' + $DuoDuoWeek + '周 &lt;' + $DDLetter + '&gt;  ' + $DDAct + '</text>',
  '  <rect x="30" y="110" width="535" height="55" rx="8" fill="#E3F2FD" opacity="0.6"/>',
  '  <text x="45" y="146" font-size="24" fill="#1565C0" font-weight="bold">&#x1F4D6; 单词跟读：</text>',
  '  <text x="200" y="146" font-size="24" fill="#333">' + $DDWords + '</text>',
  '  <rect x="30" y="180" width="535" height="260" rx="10" fill="url(#ds1)" stroke="#BBDEFB" stroke-width="1.5"/>',
  '  <text x="297" y="214" font-size="24" font-weight="bold" fill="#1565C0" text-anchor="middle">&#x1F3B6; 按发音排序 — a_e</text>',
  '  <text x="297" y="240" font-size="16" fill="#78909C" text-anchor="middle">(在横线上填入正确的单词)</text>',
$ddExStr,
  '  <rect x="30" y="455" width="535" height="195" rx="10" fill="url(#ds2)" stroke="#FFE082" stroke-width="1.5"/>',
  '  <text x="297" y="486" font-size="24" font-weight="bold" fill="#F57F17" text-anchor="middle">&#x1F4DD; 读句子填空</text>',
$ddSentStr,
  '  <rect x="30" y="662" width="535" height="170" rx="10" fill="url(#ds3)" stroke="#A5D6A7" stroke-width="1.5"/>',
$ddMathStr,
  '</svg>'
)

# ========== 小铭英语 ==========
$xmWordsList = $XMWords -split ',\s*'
$xmLetters = @()
foreach ($w in $xmWordsList) { $xmLetters += $w.Substring(0, 1) }
$xmLettersStr = $xmLetters -join ', '
$xmLetterDisplay = $XMLetter.Substring(0,1) + " " + $XMLetter.Substring(1)

$xmLetterUp = $XMLetter.Substring(0,1).ToUpper()
$xmLetterLowSimple = $XMLetter.Substring(0,1).ToLower()
$xmLetterLow = if ($xmLetterUp -eq 'A') { "&#x0251;" } else { $xmLetterLowSimple }

# 找字母：生成混合字符串
$xmFindLetter = $xmLetterLowSimple
$xmOtherLetters = @('a','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z') | Where-Object { $_ -ne $xmFindLetter }
$xmMixed = @()
$rng = [Random]::new()
# 造一个18个字母的混合串：重复目标字母 + 随机其他字母
for ($i = 0; $i -lt 6; $i++) { $xmMixed += $xmFindLetter }
for ($i = 0; $i -lt 12; $i++) { $xmMixed += $xmOtherLetters[$rng.Next(0, $xmOtherLetters.Count)] }
# 打乱
for ($i = $xmMixed.Count - 1; $i -gt 0; $i--) {
  $j = $rng.Next(0, $i + 1); $tmp = $xmMixed[$i]; $xmMixed[$i] = $xmMixed[$j]; $xmMixed[$j] = $tmp
}
$xmMixedStr = ($xmMixed -join '  ')

# 图片-单词匹配表
$xmEmojiMap = @{
  bear="🐻"; bird="🐦"; book="📖"; apple="🍎"; ant="🐜"; alligator="🐊"
  cat="🐱"; cup="🥤"; car="🚗"; dog="🐶"; duck="🐥"; doll="🧸"
  egg="🥚"; elephant="🐘"; fish="🐟"; frog="🐸"; flower="🌸"
  girl="👧"; goat="🐐"; gift="🎁"; hat="🎩"; horse="🐴"
  igloo="🏠"; insect="🐛"; ink="🖊️"; juice="🧃"; jet="✈️"
  koala="🐨"; key="🔑"; lion="🦁"; leaf="🍃"
  monkey="🐵"; milk="🥛"; moon="🌙";   nut="🥜"; nose="👃"
  ox="🐂"; octopus="🐙"; orange="🍊"; pig="🐷"; pen="✏️"; pizza="🍕"
  queen="👑"; question="❓"; rabbit="🐰"; robot="🤖"
  sun="☀️"; sock="🧦"; tea="🍵"; tiger="🐯"
  umbrella="☂️"; up="⬆️"; under="⬇️"; violin="🎻"; van="🚐"; vet="🏥"
  water="💧"; watch="⌚"; web="🕸️"; fox="🦊"; box="📦"; wax="🕯️"
  yellow="💛"; yarn="🧶"; zebra="🦓"; zoo="🏛️"
}

# 生成连线部分：左图右词
$xmMatchLines = @()
$xmMatchY = 375
# 打乱单词顺序用于右侧
$xmShuffled = @($xmWordsList | Sort-Object { $rng.Next() })
for ($i = 0; $i -lt [Math]::Min(3, $xmWordsList.Count); $i++) {
  $w = $xmWordsList[$i].Trim().ToLower()
  $emoji = if ($xmEmojiMap.ContainsKey($w)) { $xmEmojiMap[$w] } else { "❓" }
  $sw = $xmShuffled[$i].Trim().ToLower()
  $xmMatchLines += "  <text x='80' y='${xmMatchY}' font-size='48'>$emoji</text>"
  $xmMatchLines += "  <text x='320' y='$($xmMatchY+6)' font-size='32' fill='#333' font-weight='bold'>$sw</text>"
  $xmMatchY += 70
}

$xmMatchStr = $xmMatchLines -join "`n"

$xmSvg = New-Svg @(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">',
  '  <defs>',
  '    <linearGradient id="xh" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#43A047"/><stop offset="100%" stop-color="#66BB6A"/></linearGradient>',

  '    <linearGradient id="xs2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FCE4EC"/><stop offset="100%" stop-color="#FFF0F3"/></linearGradient>',
  '    <linearGradient id="xs3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8EAF6"/><stop offset="100%" stop-color="#F3F4FD"/></linearGradient>',
  '  </defs>',
  '  <rect width="595" height="842" fill="#F9FFF9"/>',
  '  <rect x="0" y="0" width="595" height="95" fill="url(#xh)" rx="0"/>',
  '  <text x="297" y="42" font-size="32" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 小铭 · 牛津自然拼读 第1册</text>',
  '  <text x="297" y="74" font-size="20" fill="#E8F5E9" text-anchor="middle">第' + $XiaoMingWeek + '周 ' + $XMLetter + '  ' + $XMAct + '</text>',
  '  <rect x="30" y="110" width="535" height="55" rx="8" fill="#E8F5E9" opacity="0.6"/>',
  '  <text x="45" y="146" font-size="24" fill="#2E7D32" font-weight="bold">&#x1F4D6; 单词：</text>',
  '  <text x="155" y="146" font-size="24" fill="#333">' + $XMWords + '</text>',
  '  <rect x="30" y="180" width="535" height="130" rx="10" fill="url(#xs2)" stroke="#F48FB1" stroke-width="1.5"/>',
  '  <text x="297" y="214" font-size="24" font-weight="bold" fill="#C62828" text-anchor="middle">&#x1F50D; 找字母</text>',
  '  <text x="297" y="236" font-size="16" fill="#78909C" text-anchor="middle">把字母 ' + $xmFindLetter + ' 圈出来：</text>',
  '  <text x="297" y="290" font-size="40" fill="#333" text-anchor="middle" font-family="monospace">' + $xmMixedStr + '</text>',
  '  <rect x="30" y="325" width="535" height="195" rx="10" fill="url(#xs3)" stroke="#9FA8DA" stroke-width="1.5"/>',
  '  <text x="297" y="360" font-size="24" font-weight="bold" fill="#283593" text-anchor="middle">&#x1F517; 图片和单词连线</text>',
$xmMatchStr,
  '  <rect x="30" y="535" width="535" height="165" rx="10" fill="url(#xs2)" stroke="#F48FB1" stroke-width="1.5"/>',
  '  <text x="297" y="566" font-size="24" font-weight="bold" fill="#C62828" text-anchor="middle">&#x1F399;&#xFE0F; 发音练习 · 首字母发音</text>',
  '  <text x="297" y="590" font-size="16" fill="#78909C" text-anchor="middle">大声说出每个单词的首字母发音，再读整词：</text>',
  '  <text x="297" y="630" font-size="28" fill="#333" text-anchor="middle">' + $xmWordsList[0] + ' → /' + $xmLetterLowSimple + '/</text>',
  '  <text x="297" y="660" font-size="28" fill="#333" text-anchor="middle">' + $xmWordsList[1] + ' → /' + $xmLetterLowSimple + '/</text>',
  '  <text x="297" y="690" font-size="28" fill="#333" text-anchor="middle">' + $xmWordsList[2] + ' → /' + $xmLetterLowSimple + '/</text>',
  '</svg>'
)

# ========== 语文 SVG ==========
$cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }

$cnSvg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 1000" width="595" height="1000">
  <defs>
    <linearGradient id="ch" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#E53935"/><stop offset="100%" stop-color="#FF7043"/></linearGradient>
    <linearGradient id="cs1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFF3E0"/><stop offset="100%" stop-color="#FFF8E1"/></linearGradient>
    <linearGradient id="cs2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FBE9E7"/><stop offset="100%" stop-color="#FFF0E8"/></linearGradient>
    <linearGradient id="cs3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8EAF6"/><stop offset="100%" stop-color="#F3F4FD"/></linearGradient>
  </defs>
  <rect width="595" height="1050" fill="#FFFBF5"/>
  <rect x="0" y="0" width="595" height="85" fill="url(#ch)" rx="0"/>
  <text x="297" y="40" font-size="26" font-weight="bold" fill="#fff" text-anchor="middle">&#x2605; 多多 · 语文冲刺</text>
  <text x="297" y="68" font-size="18" fill="#FFEBEE" text-anchor="middle">第' + $ChineseWeek + '周 · 9月入学准备</text>
  <rect x="30" y="100" width="535" height="50" rx="8" fill="#FFEBEE" opacity="0.6"/>
  <text x="45" y="132" font-size="20" fill="#C62828" font-weight="bold">&#x1F4D6; 拼音：</text>
  <text x="150" y="132" font-size="20" fill="#333">' + $ChinesePinyin + '</text>
  <rect x="30" y="165" width="535" height="220" rx="10" fill="url(#cs1)" stroke="#FFCC80" stroke-width="1.5"/>
  <text x="297" y="196" font-size="20" font-weight="bold" fill="#E65100" text-anchor="middle">&#x1F3A4; 拼音跟读 + 写一写</text>
  <text x="297" y="255" font-size="30" fill="#333" text-anchor="middle" font-weight="bold">' + $ChinesePinyin + '</text>
'

# Pinyin tracing lines
$pyList = @($ChinesePinyin -split '\s+') | Where-Object { $_ -and $_ -notmatch '巩固|复习|韵母|声母|全部|整体|拼读' }
$pyY = 285
foreach ($py in $pyList) {
  if ($pyY -gt 375) { break }
  $cnSvg += "  <text x='297' y='${pyY}' font-size='24' fill='#CFD8DC' text-anchor='middle' font-family='monospace' font-weight='bold'>$py  $py  $py  $py  $py</text>"
  $pyY += 28
}
# Tone syllables
if ($pyList.Count -le 3) {
  $syllabes = @()
  foreach ($py in $pyList) {
    $pc = $py -replace '[0-9]',''
    if ($pc -match '^[bpmfdtnlgkhjqxzcsryw]$') { $syllabes += "${pc}ā ${pc}á ${pc}ǎ ${pc}à" }
  }
  if ($syllabes.Count -gt 0) { $cnSvg += "  <text x='297' y='${pyY}' font-size='22' fill='#FF8A65' text-anchor='middle'>" + ($syllabes -join '  ') + '</text>' }
}

$cnSvg += @"
  <rect x="30" y="400" width="535" height="240" rx="10" fill="url(#cs2)" stroke="#FFAB91" stroke-width="1.5"/>
  <text x="297" y="430" font-size="20" font-weight="bold" fill="#BF360C" text-anchor="middle">&#x270F;&#xFE0F; 汉字描红 · 每个字描3遍</text>
"@

$yPos = 470
foreach ($c in $cnCharList) {
  $cnSvg += "  <text x='297' y='${yPos}' font-size='36' fill='#CFD8DC' text-anchor='middle' font-family='STKaiti,serif' font-weight='bold'>$c  $c  $c  $c  $c</text>"
  $yPos += 50
}

# ====== Poem Section ======
function Format-PoemLines {
  param([string]$Text)
  $parts = $Text -replace '[，。！？]', "|" -split '\|' | Where-Object { $_ -and $_.Trim() }
  $lines = @()
  for ($i = 0; $i -lt $parts.Count; $i += 2) {
    $line = $parts[$i].Trim()
    if ($i + 1 -lt $parts.Count) { $line += "，" + $parts[$i+1].Trim() + "。" }
    else { $line += "。" }
    $lines += $line
  }
  return $lines
}
$pBoxH = 250
$cnSvg += '  <rect x="30" y="' + ($yPos + 5) + '" width="535" height="' + $pBoxH + '" rx="10" fill="url(#cs3)" stroke="#9FA8DA" stroke-width="1.5"/>'
$pY = $yPos + 30
if ($ChinesePoem1Title) {
  $cnSvg += '  <text x="297" y="' + $pY + '" font-size="17" font-weight="bold" fill="#283593" text-anchor="middle">&#x1F4DA; ' + $ChinesePoem1Title + ' — ' + $ChinesePoem1Author + '</text>'
  $pY += 26
  $p1Lines = Format-PoemLines $ChinesePoem1Text
  foreach ($l in $p1Lines) {
    $cnSvg += '  <text x="80" y="' + $pY + '" font-size="19" fill="#333" font-family="STKaiti,serif">' + $l + '</text>'
    $pY += 26
  }
  $pY += 10
}
if ($ChinesePoem2Title) {
  $cnSvg += '  <text x="297" y="' + $pY + '" font-size="17" font-weight="bold" fill="#283593" text-anchor="middle">&#x1F4DA; ' + $ChinesePoem2Title + ' — ' + $ChinesePoem2Author + '</text>'
  $pY += 26
  $p2Lines = Format-PoemLines $ChinesePoem2Text
  foreach ($l in $p2Lines) {
    $cnSvg += '  <text x="80" y="' + $pY + '" font-size="19" fill="#333" font-family="STKaiti,serif">' + $l + '</text>'
    $pY += 26
  }
}

$taskY = $yPos + $pBoxH + 20
$cnSvg += '  <rect x="30" y="' + $taskY + '" width="535" height="65" rx="10" fill="#FFCC02" opacity="0.2"/>
  <text x="297" y="' + ($taskY + 30) + '" font-size="17" fill="#E65100" text-anchor="middle">&#x1F3C6; 今日任务：跟读3遍 | 认读 | 描红 | 读古诗 | 在家找字</text>
  <text x="297" y="' + ($taskY + 58) + '" font-size="14" fill="#B0BEC5" text-anchor="middle">活动：' + $ChineseAct + '</text>
  <text x="297" y="' + ($taskY + 88) + '" font-size="14" fill="#B0BEC5" text-anchor="middle">生成日期：' + $DateStr + '</text>
</svg>'

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
