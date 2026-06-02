param(
    [string]$DuoDuoWeek,
    [string]$XiaoMingWeek,
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

$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) {
    New-Item -ItemType Directory -Path $printableDir -Force | Out-Null
}

function Write-RtfFile {
    param($Path, $Body)
    $rtf = @(
        "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset134 Yu Gothic;}}"
        "\paperw11900\paperh16840\margl1134\margr1134\margt1134\margb1134"
        "\pard\f0\fs28"
        $Body
        "}"
    ) -join "\line "
    $utf8Bom = [System.Text.Encoding]::UTF8.GetPreamble()
    $bytes = $utf8Bom + [System.Text.Encoding]::UTF8.GetBytes($rtf)
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

function Escape-Rtf {
    param($s)
    $s -replace '\\', '\\' -replace '{', '\{' -replace '}', '\}' -replace "\n", "\line "
}

$DateStr = Get-Date -Format "yyyy-MM-dd"

# === 多多 DOCX (RTF) ===
$ddLines = @(
    "\pard\b\fs44 ${DuoDuoWeek}周 - 多多 牛津自然拼读 第3册\b0\par"
    "\pard\fs28 字母: ${DDLetter}   单词: ${DDWords}\par"
    "\pard\fs24 活动: ${DDAct}\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 按发音排序 (a_e)\b0\par"
    "\pard\fs24 (在横线上填入正确的单词)\par"
    "\pard\fs28 1. c_t (猫)\par"
    "\pard\fs28 2. c_k (蛋糕)\par"
    "\pard\fs28 3. l_k (湖)\par"
    "\pard\fs28 4. m_k (制作)\par"
    "\pard\fs28 5. n_m (名字)\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 读句子填空\b0\par"
    "\pard\fs28 I have a red c___. (帽子/cap)\par"
    "\pard\fs28 She likes to r___ fast. (跑/run)\par"
    "\pard\fs28 The sun is h___. (热/hot)\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 数学 - 数数\b0\par"
    "\pard\fs24 Count the objects and write the number:\par"
    "\pard\fs48 ☆☆☆☆☆ ☆☆☆☆☆ ☆☆☆\par"
    "\pard\fs28 数一数: ___ 个星星\par"
    "\pard \brdrb\brdrs\par"
    "\pard\fs20 生成日期: ${DateStr}\par"
)
Write-RtfFile -Path (Join-Path $printableDir "duoduo_week${DuoDuoWeek}.doc") -Body ($ddLines -join "")
Write-Output ("DOC generated: duoduo_week${DuoDuoWeek}.doc")

# === 小铭 DOCX (RTF) ===
$xmLines = @(
    "\pard\b\fs44 ${XiaoMingWeek}周 - 小铭 牛津自然拼读 第1册\b0\par"
    "\pard\fs28 字母: ${XMLetter}   单词: ${XMWords}\par"
    "\pard\fs24 活动: ${XMAct}\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 字母描写\b0\par"
    "\pard\fs24 Trace the letters:\par"
    "\pard\fs72\cf2 ${XMLetter} ${XMLetter} ${XMLetter} ${XMLetter} ${XMLetter}\par"
    "\pard\fs72\cf2 $($XMLetter.ToLower()) $($XMLetter.ToLower()) $($XMLetter.ToLower()) $($XMLetter.ToLower()) $($XMLetter.ToLower())\par"
    "\pard\fs24 (沿着灰色字母描一描)\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 找字母\b0\par"
    "\pard\fs28 把字母 ${XMLetter} 圈出来:\par"
    "\pard\fs40 a a a a a b c d e f g h\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 单词认读\b0\par"
    "\pard\fs28 大声读出下面的单词:\par"
    "\pard\fs36\b ${XMWords}\b0\par"
    "\pard \brdrb\brdrs\par"
    "\pard\b\fs36 发音练习\b0\par"
    "\pard\fs28 大声说出每个单词的首字母发音:\par"
    "\pard\fs32 apple → /a/\par"
    "\pard \brdrb\brdrs\par"
    "\pard\fs20 生成日期: ${DateStr}\par"
)
Write-RtfFile -Path (Join-Path $printableDir "xiaoming_week${XiaoMingWeek}.doc") -Body ($xmLines -join "")
Write-Output ("DOC generated: xiaoming_week${XiaoMingWeek}.doc")

# === 语文 DOCX (RTF) ===
if ($ChineseWeek) {
    $cnPinyin = Escape-Rtf $ChinesePinyin
    $cnChars = Escape-Rtf $ChineseChars
    $cnAct = Escape-Rtf $ChineseAct
    $cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }

    $cnLines = @(
        "\pard\b\fs44 ${ChineseWeek}周 - 多多 语文冲刺\b0\par"
        "\pard\fs28 拼音: ${cnPinyin}    汉字: ${cnChars}\par"
        "\pard\fs24 活动: ${cnAct}\par"
        "\pard \brdrb\brdrs\par"
        "\pard\b\fs36 拼音跟读\b0\par"
        "\pard\fs24 (大声读3遍)\par"
        "\pard\fs48\b ${cnPinyin}\b0\par"
        "\pard \brdrb\brdrs\par"
        "\pard\b\fs36 汉字描红\b0\par"
        "\pard\fs24 (每个字描2遍)\par"
    )

    foreach ($c in $cnCharList) {
        $esc = Escape-Rtf $c
        $cnLines += "\pard\fs56\cf2 ${esc}  ${esc}  ${esc}  ${esc}  ${esc}\par"
    }

    $cnLines += "\pard \brdrb\brdrs\par"
    $cnLines += "\pard\b\fs36 今日任务\b0\par"
    $cnLines += "\pard\fs28 拼音跟读3遍 | 汉字认读 | 描红练习 | 在家找字\par"
    $cnLines += "\pard \brdrb\brdrs\par"
    $cnLines += "\pard\fs20 生成日期: ${DateStr}\par"

    Write-RtfFile -Path (Join-Path $printableDir "duoduo_chinese_week${ChineseWeek}.doc") -Body ($cnLines -join "")
    Write-Output ("DOC generated: duoduo_chinese_week${ChineseWeek}.doc")
}

Write-Output "Done."
