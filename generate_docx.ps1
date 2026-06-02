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

$CR = [char]13

function Build-DuoDuo {
    param($week, $letter, $words, $act)
    @(
        "多多 - 牛津自然拼读 第3册 第${week}周"
        "字母: ${letter}    单词: ${words}"
        "活动: ${act}"
        "------------------------------------------------------------"
        "按发音排序 (a_e)"
        "(在横线上填入正确的单词)"
        "1. c_t (猫)"
        "2. c_k (蛋糕)"
        "3. l_k (湖)"
        "4. m_k (制作)"
        "5. n_m (名字)"
        "------------------------------------------------------------"
        "读句子填空"
        "I have a red c___. (帽子/cap)"
        "She likes to r___ fast. (跑/run)"
        "The sun is h___. (热/hot)"
        "------------------------------------------------------------"
        "数学 - 数数"
        "Count the objects and write the number:"
        "☆☆☆☆☆ ☆☆☆☆☆ ☆☆☆"
        "数一数: ___ 个星星"
        "生成日期: $(Get-Date -Format 'yyyy-MM-dd')"
    ) -join $CR
}

function Build-XiaoMing {
    param($week, $letter, $words, $act)
    @(
        "小铭 - 牛津自然拼读 第1册 第${week}周"
        "字母: ${letter}    单词: ${words}"
        "活动: ${act}"
        "------------------------------------------------------------"
        "字母描写 - Trace the letters"
        "${letter}  ${letter.ToLower()}${letter.ToLower()}"
        "(沿着灰色字母描一描)"
        "------------------------------------------------------------"
        "找字母 - Find the letter"
        "把字母 ${letter} 圈出来："
        "a, a, a  a, a, a  abc def"
        "------------------------------------------------------------"
        "单词认读 - Read the words"
        "大声读出下面的单词："
        $words
        "------------------------------------------------------------"
        "发音练习 - Beginning sounds"
        "大声说出每个单词的首字母发音："
        "apple → /aa/"
        "生成日期: $(Get-Date -Format 'yyyy-MM-dd')"
    ) -join $CR
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = $word.CentimetersToPoints(21.0)
    $doc.PageSetup.PageHeight = $word.CentimetersToPoints(29.7)
    $doc.PageSetup.TopMargin = $word.CentimetersToPoints(1.0)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1.0)
    $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(1.5)
    $doc.PageSetup.RightMargin = $word.CentimetersToPoints(1.5)
    $doc.Content.Font.Name = "Yu Gothic"
    $doc.Content.Font.Size = 14
    $doc.Content.Text = Build-DuoDuo $DuoDuoWeek $DDLetter $DDWords $DDAct
    $docxPath = Join-Path $printableDir "duoduo_week${DuoDuoWeek}.docx"
    $doc.SaveAs([ref][object]$docxPath, [ref][object]16)
    $doc.Close()
    Write-Output "DOCX generated: $docxPath"

    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = $word.CentimetersToPoints(21.0)
    $doc.PageSetup.PageHeight = $word.CentimetersToPoints(29.7)
    $doc.PageSetup.TopMargin = $word.CentimetersToPoints(1.0)
    $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1.0)
    $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(1.5)
    $doc.PageSetup.RightMargin = $word.CentimetersToPoints(1.5)
    $doc.Content.Font.Name = "Yu Gothic"
    $doc.Content.Font.Size = 14
    $doc.Content.Text = Build-XiaoMing $XiaoMingWeek $XMLetter $XMWords $XMAct
    $docxPath = Join-Path $printableDir "xiaoming_week${XiaoMingWeek}.docx"
    $doc.SaveAs([ref][object]$docxPath, [ref][object]16)
    $doc.Close()
    Write-Output "DOCX generated: $docxPath"

    if ($ChineseWeek) {
        $doc = $word.Documents.Add()
        $doc.PageSetup.PageWidth = $word.CentimetersToPoints(21.0)
        $doc.PageSetup.PageHeight = $word.CentimetersToPoints(29.7)
        $doc.PageSetup.TopMargin = $word.CentimetersToPoints(1.0)
        $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(1.0)
        $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(1.5)
        $doc.PageSetup.RightMargin = $word.CentimetersToPoints(1.5)
        $doc.Content.Font.Name = "Yu Gothic"
        $doc.Content.Font.Size = 14
        $cnContent = @(
            "多多 - 语文冲刺 第${ChineseWeek}周"
            "拼音: ${ChinesePinyin}    汉字: ${ChineseChars}"
            "活动: ${ChineseAct}"
            "------------------------------------------------------------"
            "拼音跟读 - 大声读3遍"
            $ChinesePinyin
            "------------------------------------------------------------"
            "汉字描红 - 每个字描2遍"
        )
        $cnCharList = @($ChineseChars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }
        foreach ($c in $cnCharList) {
            $cnContent += "${c}  ${c}  ${c}  ${c}  ${c}"
        }
        $cnContent += "------------------------------------------------------------"
        $cnContent += "今日任务"
        $cnContent += "拼音跟读3遍 | 汉字认读 | 描红练习 | 在家找字"
        $cnContent += "生成日期: $(Get-Date -Format 'yyyy-MM-dd')"

        $doc.Content.Text = ($cnContent -join $CR)
        $docxPath = Join-Path $printableDir "duoduo_chinese_week${ChineseWeek}.docx"
        $doc.SaveAs([ref][object]$docxPath, [ref][object]16)
        $doc.Close()
        Write-Output "DOCX generated: $docxPath"
    }

    Write-Output "Done."
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}
finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
