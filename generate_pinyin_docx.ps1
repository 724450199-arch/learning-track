param(
  [int]$Lesson = 1,
  [string]$WorksheetDir = ""
)

if (-not $WorksheetDir) { $WorksheetDir = "$PSScriptRoot\worksheets" }
$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) { New-Item -ItemType Directory -Path $printableDir -Force | Out-Null }

$DateStr = Get-Date -Format "yyyy-MM-dd"

$Lessons = @{}
$Lessons[1]  = @{title="a o e"; subtitle="单韵母入门"; reading=@("ā á ǎ à","ō ó ǒ ò","ē é ě è"); chars=@("a","o","e"); homework=@("读 a o e 卡片 3 遍","找 a 音物品(苹果→ā)","描红 a o e 各 1 行")}
$Lessons[2]  = @{title="i u ü"; subtitle="单韵母进阶"; reading=@("ī í ǐ ì","ū ú ǔ ù","ǖ ǘ ǚ ǜ"); chars=@("i","u","ü"); homework=@("6 个单韵母四声朗读","描红 i u ü 各 1 行","家长读声调孩子指韵母")}
$Lessons[3]  = @{title="b p m f"; subtitle="声母入门+拼读"; reading=@("b+a→bā bá bǎ bà","p+a→pā pá pǎ pà","m+a→mā má mǎ mà","f+a→fā fá fǎ fà"); chars=@("b","p","m","f"); homework=@("读 b p m f 卡片","拼读 bā(八) bí(鼻) mā(妈)","描红 b p m f 各 1 行")}
$Lessons[4]  = @{title="d t n l"; subtitle="声母进阶"; reading=@("d+a→dā dá dǎ dà","t+a→tā tá tǎ tà","n+a→ná nǎ nà","l+a→lā lá lǎ là"); chars=@("d","t","n","l"); homework=@("复习 b p m f d t n l","拼读 dà(大) tǔ(土) nǐ(你)","描红 d t n l 各 1 行")}
$Lessons[5]  = @{title="g k h"; subtitle="声母进阶"; reading=@("g+a→gā gá gǎ gà","k+a→kā kǎ kà","h+a→hā há hǎ hà","g+u→gū gú gǔ gù"); chars=@("g","k","h"); homework=@("读 g k h 卡片","拼读 gē(哥) kè(课) hé(河)","描红 g k h 各 1 行")}
$Lessons[6]  = @{title="j q x"; subtitle="+ü省写规则"; reading=@("j+i→jī jí jǐ jì","q+i→qī qí qǐ qì","x+i→xī xí xǐ xì","j+ü→jū jú jǔ jù"); chars=@("j","q","x"); homework=@("读 j q x 卡片","拼读 jú(菊) qù(去) xǔ(许)","描红 j q x 各 1 行")}
$Lessons[7]  = @{title="zh ch sh r"; subtitle="翘舌音"; reading=@("zh+a→zhā zhá zhǎ zhà","ch+a→chā chá chǎ chà","sh+a→shā shá shǎ shà","r+e→rě rè"); chars=@("zh","ch","sh","r"); homework=@("翘舌音卡片认读","拼读 zhú(竹) chē(车) shū(书)","平翘舌对比 z-zh c-ch s-sh")}
$Lessons[8]  = @{title="z c s y w"; subtitle="平舌音+整体认读"; reading=@("z+a→zā zá zǎ zà","c+a→cā cǎ cà","s+a→sā sǎ sà","yī yí yǐ yì","wū wú wǔ wù"); chars=@("z","c","s","y","w"); homework=@("23 个声母按序朗读","拼读 zú(足) sù(速) wā(蛙)","默写所有声母")}
$Lessons[9]  = @{title="声母总复习"; subtitle="阶段小测"; reading=@("b p m f d t n l","g k h j q x","zh ch sh r z c s y w"); chars=@("b-d","p-q","zh-ch-sh"); homework=@("声母表朗读打卡","区分 b-d p-q(手势记忆)","默写全部 23 个声母")}
$Lessons[10] = @{title="ai ei ui"; subtitle="复韵母入门"; reading=@("āi ái ǎi ài","ēi éi ěi èi","uī uí uǐ uì","b+ai→bāi bái bǎi bài"); chars=@("ai","ei","ui"); homework=@("ai ei ui 四声朗读","拼读 bēi(杯) huí(回)","描红 ai ei ui 各 1 行")}
$Lessons[11] = @{title="ao ou iu"; subtitle="复韵母"; reading=@("āo áo ǎo ào","ōu óu ǒu òu","iū iú iǔ iù","d+ao→dāo dǎo dào"); chars=@("ao","ou","iu"); homework=@("ao ou iu 四声朗读","拼读 gāo(高) tóu(头) liù(六)","描红 ao ou iu 各 1 行")}
$Lessons[12] = @{title="ie üe er"; subtitle="复韵母+er特殊"; reading=@("iē ié iě iè","üē üé üě üè","ēr ér ěr èr","x+ie→xiē xié xiě xiè"); chars=@("ie","üe","er"); homework=@("ie üe er 四声朗读","拼读 tiē(贴) èr(二)","复习全部复韵母")}
$Lessons[13] = @{title="an en in un ün"; subtitle="前鼻韵母"; reading=@("ān án ǎn àn","ēn én ěn èn","īn ín ǐn ìn","ūn ún ǔn ùn","ǖn ǘn ǚn ǜn"); chars=@("an","en","in","un","ün"); homework=@("前鼻韵母朗读","拼读 shān(山) mén(门) xīn(心)","描红 an en in 各 1 行")}
$Lessons[14] = @{title="ang eng ing ong"; subtitle="后鼻韵母"; reading=@("āng áng ǎng àng","ēng éng ěng èng","īng íng ǐng ìng","ōng óng ǒng òng"); chars=@("ang","eng","ing","ong"); homework=@("后鼻韵母朗读","前后对比 an-ang en-eng in-ing","描红后鼻韵母各 1 行")}
$Lessons[15] = @{title="整体认读音节"; subtitle="16 个整体认读"; reading=@("zhi chi shi ri","zi ci si","yi wu yu","ye yue yuan","yin yun ying"); chars=@("zhi","chi","shi","ri","zi","ci","si"); homework=@("整体认读表朗读打卡","用整体认读音节组词","默写全部 16 个整体认读音节")}
$Lessons[16] = @{title="总复习+闯关"; subtitle="结课"; reading=@("23 个声母","24 个韵母","16 个整体认读音节"); chars=@("名字拼音: ","____"); homework=@("默写全部声母+韵母","拼读 10 个音节","写自己的拼音名字")}

$L = $Lessons[$Lesson]

function Escape-Rtf {
    param($s)
    $s -replace '\\', '\\' -replace '{', '\{' -replace '}', '\}' -replace "\n", "\line "
}

$colorTbl = "red0\green0\blue0;\red229\green57\blue53;\red255\green243\blue224;\red251\green233\blue231;\red230\green81\blue0;\red191\green54\blue12;\red255\green255\blue255;\red232\green245\blue253;\red33\green150\blue243"

$lines = @()
$lines += "\pard\cf2\b\fs44 第${Lesson}课 - $($L.title) 拼音练习\cf1\b0\par"
$lines += "\pard\cf8\fs24 $($L.subtitle)  生成日期: ${DateStr}\cf1\par"
$lines += "\pard\cb3\cf4\b\fs36 朗读练习 - 大声读3遍\cf1\b0\par"

foreach ($rd in $L.reading) {
    $e = Escape-Rtf $rd
    $lines += "\pard\fs32\b ${e}\b0\par"
}

$lines += "\pard\cb2\cf2\b\fs36 书写练习 - 描红\cf1\b0\par"
$lines += "\pard\fs24 (在横线上描红, 注意占中格位置)\par"

foreach ($ch in $L.chars) {
    $e = Escape-Rtf $ch
    $spaces = "      "
    $lines += "\pard\fs60\cf1 ${e}\cf7 ${spaces}${e}  ${e}  ${e}  ${e}\cf1\par"
}

$lines += "\pard\cb3\cf4\b\fs36 课后练习\cf1\b0\par"
$idx = 1
foreach ($hw in $L.homework) {
    $e = Escape-Rtf $hw
    $lines += "\pard\fs28 ${idx}. ${e}\par"
    $idx++
}

$lines += "\pard\fs20 生成日期: ${DateStr}\par"

$rtf = @(
    "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset134 Yu Gothic;}{\f1\fnil\fcharset134 SimSun;}}{\colortbl;$colorTbl}"
    "\paperw11900\paperh16840\margl1134\margr1134\margt567\margb567"
    "\pard\f0\fs28"
    ($lines -join "")
    "}"
) -join "\line "

$utf8Bom = [System.Text.Encoding]::UTF8.GetPreamble()
$bytes = $utf8Bom + [System.Text.Encoding]::UTF8.GetBytes($rtf)

$path = Join-Path $printableDir "pinyin_lesson${Lesson}.doc"
[System.IO.File]::WriteAllBytes($path, $bytes)
Write-Output "✅ 生成拼音练习纸: $path"
