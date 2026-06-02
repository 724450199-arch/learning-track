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
    param($Path, $ColorTable, $Body)
    $rtf = @(
        "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset134 Yu Gothic;}}{\colortbl;$ColorTable}"
        "\paperw11900\paperh16840\margl1134\margr1134\margt567\margb567"
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

# === 多多 DOC ===
$ddColor = "red0\green0\blue0;\red21\green96\blue200;\red232\green245\blue253;\red253\green216\blue53;\red46\green125\blue50;\red232\green245\blue233;\red21\green96\blue200;\red255\green255\blue255"
$ddLines = @(
    "\pard\cb3\cf2\b\fs44 ${DuoDuoWeek}周 - 多多 牛津自然拼读 第3册\cf1\cb8\b0\par"
    "\pard\fs28 \cf2单词: ${DDWords}\tab 活动: ${DDAct}\cf1\par"
    "\pard\cb3\cf2\b\fs36 按发音排序 (a_e)\cf1\b0\par"
    "\pard\fs24 (在横线上填入正确的单词)\par"
    "\pard\fs28 1. c__ (猫)\par"
    "\pard\fs28 2. c__ (蛋糕)\par"
    "\pard\fs28 3. l__ (湖)\par"
    "\pard\fs28 4. m__ (制作)\par"
    "\pard\fs28 5. n__ (名字)\par"
    "\pard\cb4\cf2\b\fs36 读句子填空\cf1\b0\par"
    "\pard\fs28 I have a red c__. (帽子/cap)\par"
    "\pard\fs28 She likes to r__ fast. (跑/run)\par"
    "\pard\fs28 The sun is h__. (热/hot)\par"
    "\pard\cb6\cf5\b\fs36 数学 - 数数\cf1\b0\par"
    "\pard\fs24 Count the objects and write the number:\par"
    "\pard\fs48 ☆☆☆☆☆ ☆☆☆☆☆ ☆☆☆\par"
    "\pard\fs28 数一数: ___ 个星星\par"
    "\pard\fs20 生成日期: ${DateStr}\par"
)
Write-RtfFile -Path (Join-Path $printableDir "duoduo_week${DuoDuoWeek}.doc") -ColorTable $ddColor -Body ($ddLines -join "")
Write-Output ("DOC: duoduo_week${DuoDuoWeek}.doc")

# === 小铭 DOC ===
$xmColor = "red0\green0\blue0;\red67\green160\blue71;\red232\green245\blue233;\red252\green228\blue236;\red232\green234\blue246;\red40\green53\blue147;\red67\green160\blue71;\red255\green255\blue255"
$xmLetterUpper = $XMLetter.Substring(0,1)
$xmLines = @(
    "\pard\cb6\cf2\b\fs44 ${XiaoMingWeek}周 - 小铭 牛津自然拼读 第1册\cf1\cb8\b0\par"
    "\pard\fs28 \cf2单词: ${XMWords}\tab 活动: ${XMAct}\cf1\par"
    "\pard\cb2\cf2\b\fs36 字母描写\cf1\b0\par"
    "\pard\fs24 (沿着灰色字母描一描)\par"
    "\pard\fs72\cf2 ${xmLetterUpper} ${xmLetterUpper} ${xmLetterUpper} ${xmLetterUpper} ${xmLetterUpper}\cf1\par"
    "\pard\fs72\cf2 $($xmLetterUpper.ToLower()) $($xmLetterUpper.ToLower()) $($xmLetterUpper.ToLower()) $($xmLetterUpper.ToLower()) $($xmLetterUpper.ToLower())\cf1\par"
    "\pard\cb3\cf5\b\fs36 找字母\cf1\b0\par"
    "\pard\fs28 把字母 ${xmLetterUpper} 圈出来:\par"
    "\pard\fs40 a a a a a b c d e f g h\par"
    "\pard\cb4\cf5\b\fs36 单词认读\cf1\b0\par"
    "\pard\fs28 大声读出下面的单词:\par"
    "\pard\fs36\b ${XMWords}\b0\par"
    "\pard\cb3\cf5\b\fs36 发音练习\cf1\b0\par"
    "\pard\fs28 大声说出每个单词的首字母发音:\par"
    "\pard\fs32 $($xmWordsList[0]) → /$($xmLetterUpper.ToLower())/\par"
    "\pard\fs20 生成日期: ${DateStr}\par"
)
if ($XMWords) { $xmWordsList = $XMWords -split ',\s*' } else { $xmWordsList = @("A") }
Write-RtfFile -Path (Join-Path $printableDir "xiaoming_week${XiaoMingWeek}.doc") -ColorTable $xmColor -Body ($xmLines -join "")
Write-Output ("DOC: xiaoming_week${XiaoMingWeek}.doc")

# === 语文 DOC ===
if ($ChineseWeek) {
    $cnPinyin = Escape-Rtf $ChinesePinyin
    $cnChars = Escape-Rtf $ChineseChars
    $cnAct = Escape-Rtf $ChineseAct
    $cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }
    $cnColor = "red0\green0\blue0;\red229\green57\blue53;\red255\green243\blue224;\red251\green233\blue231;\red230\green81\blue0;\red191\green54\blue12;\red255\green255\blue255"

    $cnLines = @(
        "\pard\cb2\cf2\b\fs44 ${ChineseWeek}周 - 多多 语文冲刺 \b0\par"
        "\pard\fs28 \cf2拼音: ${cnPinyin}\tab 汉字: ${cnChars}\cf1\par"
        "\pard\fs24 活动: ${cnAct}\par"
        "\pard\cb3\cf4\b\fs36 拼音跟读 - 大声读3遍\cf1\b0\par"
        "\pard\fs48\b ${cnPinyin}\b0\par"
        "\pard\cb4\cf5\b\fs36 汉字描红 - 每个字描2遍\cf1\b0\par"
    )

    foreach ($c in $cnCharList) {
        $esc = Escape-Rtf $c
        $cnLines += "\pard\fs56\cf2 ${esc}  ${esc}  ${esc}  ${esc}  ${esc}\cf1\par"
    }

    $cnLines += "\pard\b\fs36 今日任务\b0\par"
    $cnLines += "\pard\fs28 拼音跟读3遍 | 汉字认读 | 描红练习 | 在家找字\par"
    $cnLines += "\pard\fs20 生成日期: ${DateStr}\par"

    Write-RtfFile -Path (Join-Path $printableDir "duoduo_chinese_week${ChineseWeek}.doc") -ColorTable $cnColor -Body ($cnLines -join "")
    Write-Output ("DOC: duoduo_chinese_week${ChineseWeek}.doc")
}

Write-Output "Done."
