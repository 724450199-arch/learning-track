param(
  [int]$Lesson = 1,
  [string]$WorksheetDir = ""
)

if (-not $WorksheetDir) { $WorksheetDir = "$PSScriptRoot\worksheets" }
$printableDir = Join-Path $WorksheetDir "printable"
if (-not (Test-Path $printableDir)) { New-Item -ItemType Directory -Path $printableDir -Force | Out-Null }

$DateStr = Get-Date -Format "yyyy-MM-dd"

function Convert-ToneMarks {
  param($s)
  $r = $s
  $r = $r -replace 'u:1', ([char]0x01D6)
  $r = $r -replace 'u:2', ([char]0x01D8)
  $r = $r -replace 'u:3', ([char]0x01DA)
  $r = $r -replace 'u:4', ([char]0x01DC)
  $r = $r -replace 'u:e', ([char]0x00FC + "e")
  $r = $r -replace 'u:n', ([char]0x00FC + "n")
  $r = $r -replace 'u:', ([char]0x00FC)
  $r = $r -replace 'a1', ([char]0x0101)
  $r = $r -replace 'a2', ([char]0x00E1)
  $r = $r -replace 'a3', ([char]0x01CE)
  $r = $r -replace 'a4', ([char]0x00E0)
  $r = $r -replace 'o1', ([char]0x014D)
  $r = $r -replace 'o2', ([char]0x00F3)
  $r = $r -replace 'o3', ([char]0x01D2)
  $r = $r -replace 'o4', ([char]0x00F2)
  $r = $r -replace 'e1', ([char]0x0113)
  $r = $r -replace 'e2', ([char]0x00E9)
  $r = $r -replace 'e3', ([char]0x011B)
  $r = $r -replace 'e4', ([char]0x00E8)
  $r = $r -replace 'i1', ([char]0x012B)
  $r = $r -replace 'i2', ([char]0x00ED)
  $r = $r -replace 'i3', ([char]0x01D0)
  $r = $r -replace 'i4', ([char]0x00EC)
  $r = $r -replace 'u1', ([char]0x016B)
  $r = $r -replace 'u2', ([char]0x00FA)
  $r = $r -replace 'u3', ([char]0x01D4)
  $r = $r -replace 'u4', ([char]0x00F9)
  return $r
}

$Lessons = @{}
$Lessons[1]  = @{title="a o e"; subtitle="单韵母入门"; reading=@("a1 a2 a3 a4","o1 o2 o3 o4","e1 e2 e3 e4"); chars=@("a","o","e"); jingle="张大嘴巴 a a a, 公鸡打鸣 o o o, 白鹅倒影 e e e"; homework=@("读 a o e 卡片 3 遍","例词: a→a1yi2(阿姨) ba4ba(爸爸) ma1ma(妈妈) da4(大) ta1(他)","例词: o→wo1(窝) wo3(我) bo1luo2(菠萝)","例词: e→e2(鹅) he2(河) ge1ge(哥哥)","描红 a o e 各 1 行")}
$Lessons[2]  = @{title="i u u:"; subtitle="单韵母进阶"; reading=@("i1 i2 i3 i4","u1 u2 u3 u4","u:1 u:2 u:3 u:4"); chars=@("i","u","u:"); jingle="牙齿对齐 i i i, 嘴巴突出 u u u, 小鱼吐泡 u: u: u:"; homework=@("6 个单韵母四声朗读","例词: i→yi1fu(衣服) li2(梨) yi3zi(椅子) di4(地)","例词: u→wu1(屋) tu3(土) gu3(鼓) du4(肚) lu4(路)","例词: u:→yu2(鱼) qu4(去) lv4(绿) nv3(女)","描红 i u u: 各 1 行")}
$Lessons[3]  = @{title="b p m f"; subtitle="声母入门+拼读"; reading=@("b+a  ba1 ba2 ba3 ba4","p+a  pa1 pa2 pa3 pa4","m+a  ma1 ma2 ma3 ma4","f+a  fa1 fa2 fa3 fa4"); chars=@("b","p","m","f"); jingle="收听广播 b b b, 上山爬坡 p p p, 两个门洞 m m m, 一根拐杖 f f f"; homework=@("读 b p m f 卡片","例词: b→ba1(八) bi2(鼻) ba4ba(爸爸) bu4(布)","例词: p→pa1(趴) pu2tao(葡萄) po4(破)","例词: m→ma1ma(妈妈) mi3(米) ma3(马) mu4(木)","例词: f→fa1(发) fu2(扶) fu4mu3(父母)","描红 b p m f 各 1 行")}
$Lessons[4]  = @{title="d t n l"; subtitle="声母进阶"; reading=@("d+a  da1 da2 da3 da4","t+a  ta1 ta2 ta3 ta4","n+a  na2 na3 na4","l+a  la1 la2 la3 la4"); chars=@("d","t","n","l"); jingle="马蹄声响 d d d, 雨伞把儿 t t t, 一扇小门 n n n, 小棍赶猪 l l l"; homework=@("复习 b p m f d t n l","例词: d→da4(大) di4di(弟弟) du2(读) deng1(灯)","例词: t→ta1(他) tu3(土) ti1(踢) tu4zi(兔子)","例词: n→na3(哪) ni3(你) nv3(女) nu3li4(努力)","例词: l→le4(乐) li2(梨) lu4(路) lv4(绿)","描红 d t n l 各 1 行")}
$Lessons[5]  = @{title="g k h"; subtitle="声母进阶"; reading=@("g+a  ga1 ga2 ga3 ga4","k+a  ka1 ka3 ka4","h+a  ha1 ha2 ha3 ha4","g+u  gu1 gu2 gu3 gu4"); chars=@("g","k","h"); jingle="白鸽飞翔 g g g, 小鸡出壳 k k k, 喝水一杯 h h h"; homework=@("读 g k h 卡片","例词: g→ge1ge(哥哥) gu3(鼓) gua1(瓜) gou3(狗)","例词: k→ka3(卡) ke4(课) ku1(哭)","例词: h→he2(河) hua1(花) hu2(湖) he1(喝)","描红 g k h 各 1 行")}
$Lessons[6]  = @{title="j q x"; subtitle="+u:省写规则"; reading=@("j+i  ji1 ji2 ji3 ji4","q+i  qi1 qi2 qi3 qi4","x+i  xi1 xi2 xi3 xi4","j+u:  ju1 ju2 ju3 ju4"); chars=@("j","q","x"); jingle="母鸡捉虫 j j j, 七彩气球 q q q, 一个大叉 x x x"; homework=@("读 j q x 卡片","例词: j→ji1(鸡) ju2(菊) jia1(家) jiu3(酒)","例词: q→qi2(旗) qu4(去) qiu2(球) qi1(七)","例词: x→xi1(西) xia4(下) xi3(洗) xue2(学)","描红 j q x 各 1 行")}
$Lessons[7]  = @{title="zh ch sh r"; subtitle="翘舌音"; reading=@("zh+a  zha1 zha2 zha3 zha4","ch+a  cha1 cha2 cha3 cha4","sh+a  sha1 sha2 sha3 sha4","r+e  re3 re4"); chars=@("zh","ch","sh","r"); jingle="织件毛衣 zh zh zh, 吃个苹果 ch ch ch, 一只小狮 sh sh sh, 太阳日出 r r r"; homework=@("翘舌音卡片认读","例词: zh→zhu2(竹) zhi3(纸) zhong1(钟) zhuo1(桌)","例词: ch→che1(车) chi1(吃) cha2(茶) chuang1(窗)","例词: sh→shu1(书) shi2(十) shan1(山) shui3(水)","例词: r→ri4(日) re4(热) ren2(人)","平翘舌对比 z-zh c-ch s-sh")}
$Lessons[8]  = @{title="z c s y w"; subtitle="平舌音整体认读"; reading=@("z+a  za1 za2 za3 za4","c+a  ca1 ca3 ca4","s+a  sa1 sa3 sa4","yi1 yi2 yi3 yi4","wu1 wu2 wu3 wu4"); chars=@("z","c","s","y","w"); jingle="小小蚕儿 z z z, 刺猬满身 c c c, 蚕儿吐丝 s s s, 树杈晾衣 y y y, 屋顶相连 w w w"; homework=@("23 个声母按序朗读","例词: z→zu2(足) zi4(字) zuo4(坐) zou3(走)","例词: c→ci4(刺) ca1(擦) cai4(菜)","例词: s→si4(四) sa3(洒) su4(速) sen1(森)","例词: y→yi1(一) yu3(雨) yun2(云) yue4(月)","例词: w→wu1(屋) wa2wa(娃娃) wo3(我) wu3(五)","默写所有声母")}
$Lessons[9]  = @{title="声母总复习"; subtitle="阶段小测"; reading=@("b p m f d t n l","g k h j q x","zh ch sh r z c s y w"); chars=@("b-d","p-q","zh-ch-sh"); jingle="声母一共 23 个, b p m f 开头来, d t n l 跟后面, g k h j q x 排, zh ch sh r z c s, y w 两位最后来"; homework=@("声母表朗读打卡","例词对比: b-d→ba1(八) da1(搭) bi2(鼻) di2(笛)","例词对比: p-q→pa1(趴) qi1(七) pu2(葡) qu2(渠)","例词对比: zh-z→zhi3(纸) zi3(紫) chi3(尺) ci3(此)","默写全部 23 个声母")}
$Lessons[10] = @{title="ai ei ui"; subtitle="复韵母入门"; reading=@("ai1 ai2 ai3 ai4","ei1 ei2 ei3 ei4","ui1 ui2 ui3 ui4","b+ai  bai1 bai2 bai3 bai4"); chars=@("ai","ei","ui"); jingle="a+i 好朋友 ai ai ai, e+i 在一起 ei ei ei, u+i 手拉手 ui ui ui"; homework=@("ai ei ui 四声朗读","例词: ai→bai2(白) kai1(开) cai4(菜) ai4(爱)","例词: ei→bei1(杯) mei2(没) fei1(飞) hei1(黑)","例词: ui→hui2(回) gui4(贵) zui3(嘴) shui4(睡)","描红 ai ei ui 各 1 行")}
$Lessons[11] = @{title="ao ou iu"; subtitle="复韵母"; reading=@("ao1 ao2 ao3 ao4","ou1 ou2 ou3 ou4","iu1 iu2 iu3 iu4","d+ao  dao1 dao3 dao4"); chars=@("ao","ou","iu"); jingle="a+o 大声喊 ao ao ao, o+u 欧洲走 ou ou ou, i+u 邮递信 iu iu iu"; homework=@("ao ou iu 四声朗读","例词: ao→gao1(高) mao1(猫) pao3(跑) dao1(刀)","例词: ou→gou3(狗) tou2(头) zou3(走)","例词: iu→liu4(六) qiu2(球) niu2(牛) jiu3(酒)","描红 ao ou iu 各 1 行")}
$Lessons[12] = @{title="ie u:e er"; subtitle="复韵母+er特殊"; reading=@("ie1 ie2 ie3 ie4","u:e1 u:e2 u:e3 u:e4","er1 er2 er3 er4","x+ie  xie1 xie2 xie3 xie4"); chars=@("ie","u:e","er"); jingle="i+e 椰子甜 ie ie ie, u:+e 月亮圆 u:e u:e u:e, e+r 耳朵听 er er er"; homework=@("ie u:e er 四声朗读","例词: ie→xie4(谢) tie1(贴) bie2(别) die2(蝶)","例词: u:e→yue4(月) xue2(学) que4(雀) jue2(觉)","例词: er→er4(二) er2(儿) er3(耳)","复习全部复韵母")}
$Lessons[13] = @{title="an en in un u:n"; subtitle="前鼻韵母"; reading=@("an1 an2 an3 an4","en1 en2 en3 en4","in1 in2 in3 in4","un1 un2 un3 un4","u:n1 u:n2 u:n3 u:n4"); chars=@("an","en","in","un","u:n"); jingle="a+n 天安门 an an an, e+n 摁门铃 en en en, i+n 树阴凉 in in in, u+n 春天暖 un un un, u:+n 飘白云 u:n u:n u:n"; homework=@("前鼻韵母朗读","例词: an→shan1(山) ban1(班) lan2(蓝) kan4(看)","例词: en→men2(门) ben3(本) ren2(人)","例词: in→jin1(金) xin1(心) lin2(林)","例词: un→chun1(春) lun2(轮) kun4(困)","例词: u:n→yun2(云) qun2(裙) jun1(军)","描红 an en in 各 1 行")}
$Lessons[14] = @{title="ang eng ing ong"; subtitle="后鼻韵母"; reading=@("ang1 ang2 ang3 ang4","eng1 eng2 eng3 eng4","ing1 ing2 ing3 ing4","ong1 ong2 ong3 ong4"); chars=@("ang","eng","ing","ong"); jingle="a+ng 大灰狼 ang ang ang, e+ng 开台灯 eng eng eng, i+ng 老鹰飞 ing ing ing, o+ng 时钟响 ong ong ong"; homework=@("后鼻韵母朗读","例词: ang→yang2(羊) fang1(方) zhang1(张)","例词: eng→feng1(风) deng1(灯) leng3(冷)","例词: ing→xing1(星) ming2(明) ting1(听)","例词: ong→hong2(红) long2(龙) zhong1(钟)","描红后鼻韵母各 1 行")}
$Lessons[15] = @{title="整体认读音节"; subtitle="16 个整体认读"; reading=@("zhi chi shi ri","zi ci si","yi wu yu","ye yue yuan","yin yun ying"); chars=@("zhi","chi","shi","ri","zi","ci","si"); jingle="zhi chi shi ri 翘舌音, zi ci si 平舌音, yi wu yu 不用拼, ye yue yuan 整体记, yin yun ying 要牢记"; homework=@("整体认读表朗读打卡","例词: zhi→zhi1dao4(知道) zhi3(纸)","例词: chi→chi1fan4(吃饭) chi2(迟) chi3(尺)","例词: shi→shi2(十) shi4(是)","例词: ri→ri4(日)","例词: zi→zi4(字) zi3(紫)","例词: ci→ci2(词) ci4(次)","例词: si→si4(四) si1(丝)","默写全部 16 个整体认读音节")}
$Lessons[16] = @{title="总复习+闯关"; subtitle="结课"; reading=@("23 个声母","24 个韵母","16 个整体认读音节"); chars=@("名字拼音: ","____"); jingle="拼音王国真有趣, 声母韵母在一起, 整体认读要记住, 我能读又能写"; homework=@("默写全部声母+韵母","拼读: shan1 shui3(山水) hua1 duo1(花朵) tai2 yang2(太阳) xiao3 gou3(小狗) gong1 ji1(公鸡)","写自己的拼音名字")}

$L = $Lessons[$Lesson]

function Escape-RtfUnicode {
  param($s)
  $sb = New-Object System.Text.StringBuilder
  for ($i = 0; $i -lt $s.Length; $i++) {
    $c = $s[$i]; $code = [int]$c
    if ($code -le 127) {
      if ($code -eq 92) { [void]$sb.Append("\\") }
      elseif ($code -eq 123) { [void]$sb.Append("\{") }
      elseif ($code -eq 125) { [void]$sb.Append("\}") }
      else { [void]$sb.Append($c) }
    } else { [void]$sb.Append("\u${code}?") }
  }
  return $sb.ToString()
}

$colors = "red0\green0\blue0;\red229\green57\blue53;\red255\green243\blue224;\red251\green233\blue231;\red230\green81\blue0;\red191\green54\blue12;\red255\green255\blue255;\red232\green245\blue253;\red33\green150\blue243"

function Fmt([string]$s) {
  return Escape-RtfUnicode (Convert-ToneMarks $s)
}

$lines = @()
$lines += "\pard\cf2\b\fs44 " + (Fmt "第${Lesson}课 - $($L.title) 拼音练习") + "\cf1\b0\par"
$lines += "\pard\cf8\fs24 " + (Fmt "$($L.subtitle)  生成日期: ${DateStr}") + "\cf1\par"
$lines += "\pard\cb3\cf4\b\fs36 " + (Fmt "朗读练习 - 大声读3遍") + "\cf1\b0\par"

foreach ($rd in $L.reading) {
  $lines += "\pard\fs32\b " + (Fmt $rd) + "\b0\par"
}

$lines += "\pard\cf6\b\fs28 " + (Fmt "顺口溜: $($L.jingle)") + "\cf1\b0\par"

$lines += "\pard\cb2\cf2\b\fs36 " + (Fmt "书写练习 - 描红") + "\cf1\b0\par"
$lines += "\pard\fs24 " + (Fmt "(在横线上描红, 注意占中格位置)") + "\par"

foreach ($ch in $L.chars) {
  $e = Fmt $ch
  $lines += "\pard\fs60\cf1\f1 ${e}\cf7       ${e}  ${e}  ${e}  ${e}\cf1\f0\par"
}

$lines += "\pard\cb3\cf4\b\fs36 " + (Fmt "课后练习") + "\cf1\b0\par"
$idx = 1
foreach ($hw in $L.homework) {
  $lines += "\pard\fs28 " + (Fmt "${idx}. ${hw}") + "\par"
  $idx++
}

$lines += "\pard\fs20 " + (Fmt "生成日期: ${DateStr}") + "\par"

$header = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset134 SimSun;}{\f1\fnil\fcharset0 Century Gothic;}}{\colortbl;$colors}\paperw11900\paperh16840\margl1134\margr1134\margt567\margb567\pard\f0\fs28"
$body = $lines -join ""
$rtf = $header + $body + "}"

$bytes = [System.Text.Encoding]::ASCII.GetBytes($rtf)

$path = Join-Path $printableDir "pinyin_lesson${Lesson}.doc"
[System.IO.File]::WriteAllBytes($path, $bytes)
Write-Output "生成拼音练习纸: $path"
