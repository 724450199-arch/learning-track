param(
  [string]$OutputPath = "$PSScriptRoot\一年级上册教学大纲.pptx"
)

Write-Host "正在生成 PPT..."
$ppt = New-Object -ComObject PowerPoint.Application
# $ppt.Visible = msoFalse not allowed in this environment
$pres = $ppt.Presentations.Add()

function Add-Slide {
  param($Pres, [int]$LayoutIndex = 1)
  return $Pres.Slides.Add($Pres.Slides.Count + 1, $LayoutIndex)
}

# ====== 第1页：标题 ======
$s1 = Add-Slide -Pres $pres -LayoutIndex 1
$s1.Shapes.Title.TextFrame.TextRange.Text = "人教版（2024版）一年级上册"
$s1.Shapes.Title.TextFrame.TextRange.Font.Size = 36
$s1.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$s1.Shapes(2).TextFrame.TextRange.Text = "教学大纲`n`n多多 · 6岁`n2026年6月更新"
$s1.Shapes(2).TextFrame.TextRange.Font.Size = 20

# ====== 第2页：语文总览 ======
$s2 = Add-Slide -Pres $pres
$s2.Shapes.Title.TextFrame.TextRange.Text = "语文 · 13周计划"
$s2.Shapes.Title.TextFrame.TextRange.Font.Size = 28
$s2.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$rows2 = 13 + 1
$tbl2 = $s2.Shapes.AddTable($rows2, 3, 30, 80, 680, 26).Table
$tbl2.Cell(1,1).Shape.TextFrame.TextRange.Text = "周次"
$tbl2.Cell(1,2).Shape.TextFrame.TextRange.Text = "单元"
$tbl2.Cell(1,3).Shape.TextFrame.TextRange.Text = "重点内容"
foreach ($c in 1..3) { $tbl2.Cell(1,$c).Shape.TextFrame.TextRange.Font.Bold = $true; $tbl2.Cell(1,$c).Shape.TextFrame.TextRange.Font.Size = 10 }
$d2 = @(
  @("1","入学+识字1-2","天地人 金木水火土"),
  @("2","识字3-5","口耳目 日月水火 对韵歌"),
  @("3","拼音1-4","a o e / i u ü / b p m f / d t n l"),
  @("4","拼音5-9","g k h / j q x / zh ch sh r / z c s y w"),
  @("5","拼音10-14","复韵母+鼻韵母全部"),
  @("6","拼音总复习","全部拼音+拼读"),
  @("7","课文1-2","秋天 + 小小的船"),
  @("8","课文3-4","江南 + 四季"),
  @("9","识字6-8","画 / 大小多少 / 小书包"),
  @("10","识字9-10","日月明 + 升国旗"),
  @("11","课文5-7","影子 / 比尾巴 / 青蛙写诗"),
  @("12","课文8-11","雨点儿 / 明天要远足 / 大还是小 / 项链"),
  @("13","课文12-14","雪地里小画家 / 乌鸦喝水 / 小蜗牛")
)
for ($r=0; $r -lt 13; $r++) {
  for ($c=0; $c -lt 3; $c++) {
    $tbl2.Cell($r+2,$c+1).Shape.TextFrame.TextRange.Text = $d2[$r][$c]
    $tbl2.Cell($r+2,$c+1).Shape.TextFrame.TextRange.Font.Size = 9
  }
}

# ====== 第3页：数学总览 ======
$s3 = Add-Slide -Pres $pres
$s3.Shapes.Title.TextFrame.TextRange.Text = "数学 · 核心单元"
$s3.Shapes.Title.TextFrame.TextRange.Font.Size = 28
$s3.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$rows3 = 7 + 1
$tbl3 = $s3.Shapes.AddTable($rows3, 3, 30, 80, 680, 32).Table
$tbl3.Cell(1,1).Shape.TextFrame.TextRange.Text = "周次"
$tbl3.Cell(1,2).Shape.TextFrame.TextRange.Text = "单元"
$tbl3.Cell(1,3).Shape.TextFrame.TextRange.Text = "内容"
foreach ($c in 1..3) { $tbl3.Cell(1,$c).Shape.TextFrame.TextRange.Font.Bold = $true; $tbl3.Cell(1,$c).Shape.TextFrame.TextRange.Font.Size = 11 }
$d3 = @(
  @("1-2","数学游戏","校园寻宝+教室数数+分类"),
  @("3-5","一、5以内","1~5认识+比大小+第几+分与合+加减法+0"),
  @("6-11","二、6~10","6~9认识+分与合+6/7/8/9加减+10认识+连加连减"),
  @("12-13","三、立体图形","长方体/正方体/圆柱/球"),
  @("14-15","四、11~20","10加几+认识11~20+数位+简单加减"),
  @("16-18","五、进位加法","9/8/7/6/5/4/3/2加几 凑十法"),
  @("19-20","六、总复习","全册回顾+闯关游戏")
)
for ($r=0; $r -lt 7; $r++) {
  for ($c=0; $c -lt 3; $c++) {
    $tbl3.Cell($r+2,$c+1).Shape.TextFrame.TextRange.Text = $d3[$r][$c]
    $tbl3.Cell($r+2,$c+1).Shape.TextFrame.TextRange.Font.Size = 11
  }
}

# ====== 第4页：数学拓展 ======
$s4 = Add-Slide -Pres $pres
$s4.Shapes.Title.TextFrame.TextRange.Text = "数学 · 拓展延伸"
$s4.Shapes.Title.TextFrame.TextRange.Font.Size = 28
$s4.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$rows4 = 10 + 1
$tbl4 = $s4.Shapes.AddTable($rows4, 2, 30, 80, 680, 30).Table
$tbl4.Cell(1,1).Shape.TextFrame.TextRange.Text = "周次"
$tbl4.Cell(1,2).Shape.TextFrame.TextRange.Text = "内容"
foreach ($c in 1..2) { $tbl4.Cell(1,$c).Shape.TextFrame.TextRange.Font.Bold = $true }
$d4 = @(
  @("21","20以内退位减法（下册预备）"),
  @("22","100以内数初步 整十数"),
  @("23","认识人民币（元角分）"),
  @("24","认识钟表（整点半点）"),
  @("25","找规律 图形与数字"),
  @("26","分类与整理"),
  @("27","位置（上下前后左右）"),
  @("28","长度比较 轻重比较"),
  @("29","图文应用题 解题策略"),
  @("30-31","上册总复习挑战")
)
for ($r=0; $r -lt 10; $r++) {
  $tbl4.Cell($r+2,1).Shape.TextFrame.TextRange.Text = $d4[$r][0]
  $tbl4.Cell($r+2,2).Shape.TextFrame.TextRange.Text = $d4[$r][1]
}

# ====== 第5页：每日安排 ======
$s5 = Add-Slide -Pres $pres
$s5.Shapes.Title.TextFrame.TextRange.Text = "每日学习安排"
$s5.Shapes.Title.TextFrame.TextRange.Font.Size = 28
$s5.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$b1 = $s5.Shapes.AddTextbox(1, 30, 80, 300, 250)
$b1.TextFrame.TextRange.Text = "语文`n`n· 跟读拼音/识字 5min`n· 汉字认读+描红 5min`n· 读古诗 3min`n· 在家找字 2min"
$b1.TextFrame.TextRange.Font.Size = 14
$b2 = $s5.Shapes.AddTextbox(1, 380, 80, 300, 250)
$b2.TextFrame.TextRange.Text = "数学`n`n· 数感/计算练习 5min`n· 游戏化活动 5min`n· 练习纸 5min"
$b2.TextFrame.TextRange.Font.Size = 14
$b3 = $s5.Shapes.AddTextbox(1, 30, 320, 680, 100)
$b3.TextFrame.TextRange.Text = "温馨提示`n· 每天轮流语文和数学，隔天交替`n· 周六日做综合复习或游戏`n· 保持15-20分钟，不要超时"
$b3.TextFrame.TextRange.Font.Size = 14

# ====== 第6页：总结 ======
$s6 = Add-Slide -Pres $pres -LayoutIndex 1
$s6.Shapes.Title.TextFrame.TextRange.Text = "每天进步一点点！"
$s6.Shapes.Title.TextFrame.TextRange.Font.Size = 36
$s6.Shapes.Title.TextFrame.TextRange.Font.Bold = $true
$s6.Shapes(2).TextFrame.TextRange.Text = "目标：9月入学前掌握拼音+前100字+20以内加减法`n`n每天15分钟，开开心心学语文！"
$s6.Shapes(2).TextFrame.TextRange.Font.Size = 20

# ====== 保存 ======
$pres.SaveAs([System.IO.Path]::GetFullPath($OutputPath))
$pres.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "PPT 已生成: $OutputPath"
