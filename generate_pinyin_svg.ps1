param(
  [int]$Lesson = 1,
  [string]$WorksheetDir = ""
)

if (-not $WorksheetDir) { $WorksheetDir = "$PSScriptRoot\worksheets" }
$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) { New-Item -ItemType Directory -Path $printableDir -Force | Out-Null }

$DateStr = Get-Date -Format "yyyy-MM-dd"

$Lessons = @{}
$Lessons[1]  = @{title="a o e"; subtitle="单韵母入门"; reading=@("ā á ǎ à","ō ó ǒ ò","ē é ě è"); rhymes=@("阿姨好 ā á ǎ à","公鸡喔喔 ō ó ǒ ò","大白鹅 ē é ě è"); chars=@("a","o","e"); homework=@("读 a o e 卡片 3 遍","找 a 音物品(苹果→ā)","描红 a o e 各 1 行")}
$Lessons[2]  = @{title="i u ü"; subtitle="单韵母进阶"; reading=@("ī í ǐ ì","ū ú ǔ ù","ǖ ǘ ǚ ǜ"); rhymes=@("衣服衣服 ī í ǐ ì","乌鸦乌鸦 ū ú ǔ ù","小鱼小鱼 ǖ ǘ ǚ ǜ"); chars=@("i","u","ü"); homework=@("6 个单韵母四声朗读","描红 i u ü 各 1 行","家长读声调孩子指韵母")}
$Lessons[3]  = @{title="b p m f"; subtitle="声母入门+拼读"; reading=@("b+a→bā bá bǎ bà","p+a→pā pá pǎ pà","m+a→mā má mǎ mà","f+a→fā fá fǎ fà"); rhymes=@("广播广播 b b b","泼水泼水 p p p","摸头摸头 m m m","佛像佛像 f f f"); chars=@("b","p","m","f"); homework=@("读 b p m f 卡片","拼读 bā(八) bí(鼻) mā(妈)","描红 b p m f 各 1 行")}
$Lessons[4]  = @{title="d t n l"; subtitle="声母进阶"; reading=@("d+a→dā dá dǎ dà","t+a→tā tá tǎ tà","n+a→ná nǎ nà","l+a→lā lá lǎ là"); rhymes=@("马蹄声声 d d d","伞把弯弯 t t t","大门打开 n n n","棍子直直 l l l"); chars=@("d","t","n","l"); homework=@("复习 b p m f d t n l","拼读 dà(大) tǔ(土) nǐ(你)","描红 d t n l 各 1 行")}
$Lessons[5]  = @{title="g k h"; subtitle="声母进阶"; reading=@("g+a→gā gá gǎ gà","k+a→kā kǎ kà","h+a→hā há hǎ hà","g+u→gū gú gǔ gù"); rhymes=@("鸽子鸽子 g g g","蝌蚪蝌蚪 k k k","喝水喝水 h h h"); chars=@("g","k","h"); homework=@("读 g k h 卡片","拼读 gē(哥) kè(课) hé(河)","描红 g k h 各 1 行")}
$Lessons[6]  = @{title="j q x"; subtitle="+ü省写规则"; reading=@("j+i→jī jí jǐ jì","q+i→qī qí qǐ qì","x+i→xī xí xǐ xì","j+ü→jū jú jǔ jù"); rhymes=@("j q x 真淘气，见到 ü 眼就挖去"); chars=@("j","q","x"); homework=@("读 j q x 卡片","拼读 jú(菊) qù(去) xǔ(许)","描红 j q x 各 1 行")}
$Lessons[7]  = @{title="zh ch sh r"; subtitle="翘舌音"; reading=@("zh+a→zhā zhá zhǎ zhà","ch+a→chā chá chǎ chà","sh+a→shā shá shǎ shà","r+e→rě rè"); rhymes=@("织毛衣 zh zh zh","吃东西 ch ch ch","狮子狮子 sh sh sh","日出日出 r r r"); chars=@("zh","ch","sh","r"); homework=@("翘舌音卡片认读","拼读 zhú(竹) chē(车) shū(书)","平翘舌对比 z-zh c-ch s-sh")}
$Lessons[8]  = @{title="z c s y w"; subtitle="平舌音+整体认读"; reading=@("z+a→zā zá zǎ zà","c+a→cā cǎ cà","s+a→sā sǎ sà","yī yí yǐ yì","wū wú wǔ wù"); rhymes=@("写字写字 z z z","擦桌擦桌 c c c","蚕丝蚕丝 s s s"); chars=@("z","c","s","y","w"); homework=@("23 个声母按序朗读","拼读 zú(足) sù(速) wā(蛙)","默写所有声母")}
$Lessons[9]  = @{title="声母总复习"; subtitle="阶段小测"; reading=@("b p m f d t n l","g k h j q x","zh ch sh r z c s y w"); rhymes=@("平舌: z c s / 翘舌: zh ch sh"); chars=@("b-d","p-q","zh-ch-sh"); homework=@("声母表朗读打卡","区分 b-d p-q(手势记忆)","默写全部 23 个声母")}
$Lessons[10] = @{title="ai ei ui"; subtitle="复韵母入门"; reading=@("āi ái ǎi ài","ēi éi ěi èi","uī uí uǐ uì","b+ai→bāi bái bǎi bài"); rhymes=@("挨着挨着 āi ái ǎi ài","飞飞飞 ēi éi ěi èi","围巾围巾 uī uí uǐ uì"); chars=@("ai","ei","ui"); homework=@("ai ei ui 四声朗读","拼读 bēi(杯) huí(回)","描红 ai ei ui 各 1 行")}
$Lessons[11] = @{title="ao ou iu"; subtitle="复韵母"; reading=@("āo áo ǎo ào","ōu óu ǒu òu","iū iú iǔ iù","d+ao→dāo dǎo dào"); rhymes=@("奥运奥运 āo áo ǎo ào","海鸥海鸥 ōu óu ǒu òu","游泳游泳 iū iú iǔ iù"); chars=@("ao","ou","iu"); homework=@("ao ou iu 四声朗读","拼读 gāo(高) tóu(头) liù(六)","描红 ao ou iu 各 1 行")}
$Lessons[12] = @{title="ie üe er"; subtitle="复韵母+er特殊"; reading=@("iē ié iě iè","üē üé üě üè","ēr ér ěr èr","x+ie→xiē xié xiě xiè"); rhymes=@("椰子椰子 iē ié iě iè","月亮月亮 üē üé üě üè","耳朵耳朵 ēr ér ěr èr"); chars=@("ie","üe","er"); homework=@("ie üe er 四声朗读","拼读 tiē(贴) èr(二)","复习全部复韵母")}
$Lessons[13] = @{title="an en in un ün"; subtitle="前鼻韵母"; reading=@("ān án ǎn àn","ēn én ěn èn","īn ín ǐn ìn","ūn ún ǔn ùn","ǖn ǘn ǚn ǜn"); rhymes=@("天安门 ān án ǎn àn","摁门铃 ēn én ěn èn"); chars=@("an","en","in","un","ün"); homework=@("前鼻韵母朗读","拼读 shān(山) mén(门) xīn(心)","描红 an en in 各 1 行")}
$Lessons[14] = @{title="ang eng ing ong"; subtitle="后鼻韵母"; reading=@("āng áng ǎng àng","ēng éng ěng èng","īng íng ǐng ìng","ōng óng ǒng òng"); rhymes=@("昂首 āng áng ǎng àng","台灯 ēng éng ěng èng","星星 īng íng ǐng ìng","闹钟 ōng óng ǒng òng"); chars=@("ang","eng","ing","ong"); homework=@("后鼻韵母朗读","前后对比 an-ang en-eng in-ing","描红后鼻韵母各 1 行")}
$Lessons[15] = @{title="整体认读音节"; subtitle="16 个整体认读"; reading=@("zhi chi shi ri","zi ci si","yi wu yu","ye yue yuan","yin yun ying"); rhymes=@("四是四，十是十，十四是十四"); chars=@("zhī","chī","shī","rì","zì","cì","sì"); homework=@("整体认读表朗读打卡","用整体认读音节组词","默写全部 16 个整体认读音节")}
$Lessons[16] = @{title="总复习+闯关"; subtitle="结课"; reading=@("23 个声母","24 个韵母","16 个整体认读音节"); rhymes=@("拼音小故事: mā ma qí chē"); chars=@("名字拼音: ____"); homework=@("默写全部声母+韵母","拼读 10 个音节","写自己的拼音名字 ✅")}

$L = $Lessons[$Lesson]

$lines = @()
$lines += '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 595 842" width="595" height="842">'
$lines += '  <defs>'
$lines += '    <linearGradient id="h" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#E53935"/><stop offset="100%" stop-color="#FF7043"/></linearGradient>'
$lines += '    <linearGradient id="s1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFF3E0"/><stop offset="100%" stop-color="#FFF8E1"/></linearGradient>'
$lines += '    <linearGradient id="s2" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FBE9E7"/><stop offset="100%" stop-color="#FFF0E8"/></linearGradient>'
$lines += '    <linearGradient id="s3" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#E8EAF6"/><stop offset="100%" stop-color="#F3F4FD"/></linearGradient>'
$lines += '  </defs>'
$lines += '  <rect width="595" height="842" fill="#FFFBF5"/>'
$lines += '  <rect x="0" y="0" width="595" height="85" fill="url(#h)" rx="0"/>'
$lines += "  <text x='297' y='40' font-size='26' font-weight='bold' fill='#fff' text-anchor='middle'>&#x1F524; 拼音专项练习 &#xB7; 第${Lesson}课</text>"
$lines += "  <text x='297' y='68' font-size='18' fill='#FFEBEE' text-anchor='middle'>$($L.title) &#xB7; $($L.subtitle)</text>"
$lines += '  <rect x="30" y="100" width="535" height="50" rx="8" fill="#FFEBEE" opacity="0.6"/>'
$lines += '  <text x="45" y="132" font-size="20" fill="#C62828" font-weight="bold">&#x1F444; 朗读练习：大声读 3 遍</text>'
$lines += '  <rect x="30" y="165" width="535" height="200" rx="10" fill="url(#s1)" stroke="#FFCC80" stroke-width="1.5"/>'
$lines += '  <text x="297" y="196" font-size="22" font-weight="bold" fill="#E65100" text-anchor="middle">&#x1F4E2; 四声调 / 拼读练习</text>'

$yPos = 230
foreach ($line in $L.reading) {
  $lines += "  <text x='50' y='${yPos}' font-size='26' fill='#333' font-family='monospace'>$line</text>"
  $yPos += 34
}

$lines += '  <rect x="30" y="380" width="535" height="240" rx="10" fill="url(#s2)" stroke="#FFAB91" stroke-width="1.5"/>'
$lines += '  <text x="297" y="410" font-size="22" font-weight="bold" fill="#BF360C" text-anchor="middle">&#x270F;&#xFE0F; 书写练习 &#xB7; 描红</text>'
$lines += '  <text x="297" y="435" font-size="14" fill="#9E9E9E" text-anchor="middle">在四线三格中描红，注意占格位置</text>'

$yPos2 = 470
$colors = @("#D32F2F","#1565C0","#2E7D32","#F57F17","#6A1B9A")
$ci = 0
foreach ($ch in $L.chars) {
  if ($yPos2 -gt 610) { break }
  $c = if ($ci -lt $colors.Count) { $colors[$ci] } else { "#333" }
  $ci++
  $lines += "  <text x='50' y='${yPos2}' font-size='28' fill='${c}' font-weight='bold'>$ch</text>"
  $lines += "  <line x1='90' y1='$($yPos2-15)' x2='560' y2='$($yPos2-15)' stroke='#CFD8DC' stroke-width='0.5'/>"
  $lines += "  <line x1='90' y1='$($yPos2-5)' x2='560' y2='$($yPos2-5)' stroke='#B0BEC5' stroke-width='0.5'/>"
  $lines += "  <line x1='90' y1='$($yPos2+5)' x2='560' y2='$($yPos2+5)' stroke='#B0BEC5' stroke-width='0.5'/>"
  $lines += "  <line x1='90' y1='$($yPos2+15)' x2='560' y2='$($yPos2+15)' stroke='#CFD8DC' stroke-width='0.5'/>"
  $repeat = ($ch,$ch,$ch,$ch,$ch -join "  ")
  $lines += "  <text x='110' y='$($yPos2+5)' font-size='22' fill='#E0E0E0' font-family='monospace'>$repeat</text>"
  $yPos2 += 38
}

$homeworkY = 635
$lines += "  <rect x='30' y='${homeworkY}' width='535' height='195' rx='10' fill='url(#s3)' stroke='#9FA8DA' stroke-width='1.5'/>"
$lines += "  <text x='297' y='$($homeworkY+30)' font-size='20' font-weight='bold' fill='#283593' text-anchor='middle'>&#x1F3E0; 课后练习</text>"

$hy = $homeworkY + 65
$idx = 1
foreach ($hw in $L.homework) {
  $lines += "  <text x='50' y='${hy}' font-size='18' fill='#333'>${idx}. $hw</text>"
  $hy += 32
  $idx++
}

$lines += "  <text x='297' y='830' font-size='14' fill='#B0BEC5' text-anchor='middle'>生成日期：${DateStr} | 多多加油！</text>"
$lines += '</svg>'

$svg = $lines -join "`n"

$path = Join-Path $printableDir "pinyin_lesson${Lesson}.svg"
[System.IO.File]::WriteAllText($path, $svg, [System.Text.Encoding]::UTF8)
Write-Output "✅ 生成拼音练习纸: $path"
