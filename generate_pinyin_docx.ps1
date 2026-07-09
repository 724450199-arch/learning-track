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
    if ($code -eq 92) { [void]$sb.Append("\\") }
    elseif ($code -eq 123) { [void]$sb.Append("\{") }
    elseif ($code -eq 125) { [void]$sb.Append("\}") }
    elseif ($code -le 127) { [void]$sb.Append($c) }
    else { [void]$sb.Append($c) }
  }
  return $sb.ToString()
}

$colors = "\red0\green0\blue0;\red229\green57\blue53;\red255\green243\blue224;\red230\green81\blue0;\red191\green54\blue12;\red232\green245\blue253;\red245\green248\blue252;\red0\green51\blue153;\red255\green255\blue255"

function Fmt([string]$s) {
  return Escape-RtfUnicode (Convert-ToneMarks $s)
}

function Get-WordIllustration {
  param($word)
  $last = $word[-1]
  $d = @{
    "姨"='<g fill="white" fill-opacity="0.3"><circle cx="40" cy="38" r="10"/></g><circle cx="37" cy="36" r="1.5" fill="white"/><circle cx="43" cy="36" r="1.5" fill="white"/><path d="M36 42 Q40 46 44 42" stroke="white" stroke-width="1.5" fill="none"/><circle cx="32" cy="30" r="3" fill="white" fill-opacity="0.2"/>'
    "爸"='<g fill="white" fill-opacity="0.3"><circle cx="40" cy="38" r="10"/></g><circle cx="37" cy="36" r="1.5" fill="white"/><circle cx="43" cy="36" r="1.5" fill="white"/><path d="M36 42 Q40 44 44 42" stroke="white" stroke-width="1.5" fill="none"/>'
    "妈"='<g fill="white" fill-opacity="0.3"><circle cx="40" cy="38" r="10"/></g><circle cx="37" cy="36" r="1.5" fill="white"/><circle cx="43" cy="36" r="1.5" fill="white"/><path d="M36 42 Q40 46 44 42" stroke="white" stroke-width="1.5" fill="none"/>'
    "哥"='<g fill="white" fill-opacity="0.3"><circle cx="40" cy="38" r="10"/></g><circle cx="37" cy="36" r="1.5" fill="white"/><circle cx="43" cy="36" r="1.5" fill="white"/><path d="M36 42 Q40 44 44 42" stroke="white" stroke-width="1.5" fill="none"/><rect x="25" y="32" width="6" height="4" rx="2" fill="white" fill-opacity="0.2"/><rect x="49" y="32" width="6" height="4" rx="2" fill="white" fill-opacity="0.2"/>'
    "弟"='<g fill="white" fill-opacity="0.3"><circle cx="40" cy="40" r="8"/></g><circle cx="37" cy="36" r="1.5" fill="white"/><circle cx="43" cy="36" r="1.5" fill="white"/><path d="M36 43 Q40 45 44 43" stroke="white" stroke-width="1.5" fill="none"/><rect x="26" y="30" width="5" height="3" rx="1.5" fill="white" fill-opacity="0.2"/><rect x="49" y="30" width="5" height="3" rx="1.5" fill="white" fill-opacity="0.2"/>'
    "鹅"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="42" rx="8" ry="6"/></g><path d="M40 36 Q42 30 46 32" stroke="white" stroke-width="1.5" fill="none"/><circle cx="46" cy="31" r="1" fill="white"/>'
    "鱼"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="40" rx="12" ry="5"/></g><path d="M52 40 L56 36 L56 44 Z" fill="white" fill-opacity="0.3"/><circle cx="34" cy="39" r="1" fill="white"/>'
    "猫"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="42" rx="8" ry="7"/></g><path d="M32 36 L34 30 L37 35" stroke="white" stroke-width="1" fill="none"/><path d="M48 36 L46 30 L43 35" stroke="white" stroke-width="1" fill="none"/><circle cx="37" cy="40" r="1" fill="white"/><circle cx="43" cy="40" r="1" fill="white"/>'
    "狗"='<g fill="white" fill-opacity="0.3"><ellipse cx="35" cy="44" rx="10" ry="6"/></g><circle cx="45" cy="40" r="6" fill="white" fill-opacity="0.2"/><circle cx="44" cy="39" r="1" fill="white"/>'
    "牛"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="44" rx="8" ry="5"/></g><circle cx="40" cy="38" r="6" fill="white" fill-opacity="0.2"/><path d="M36 33 L34 28 M44 33 L46 28" stroke="white" stroke-width="1.5" fill="none"/><circle cx="37" cy="37" r="1" fill="white"/><circle cx="43" cy="37" r="1" fill="white"/>'
    "马"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="44" rx="10" ry="4"/></g><path d="M38 40 Q40 32 44 30 Q46 28 48 30" stroke="white" stroke-width="1.5" fill="none"/><circle cx="43" cy="31" r="1" fill="white"/>'
    "鸡"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="44" rx="7" ry="5"/></g><circle cx="40" cy="36" r="5" fill="white" fill-opacity="0.2"/><path d="M47 34 L52 32 L49 36" fill="white" fill-opacity="0.2"/><circle cx="38" cy="35" r="1" fill="white"/>'
    "兔"='<g fill="white" fill-opacity="0.3"><ellipse cx="40" cy="44" rx="6" ry="5"/></g><circle cx="40" cy="36" r="6" fill="white" fill-opacity="0.2"/><path d="M34 32 L32 26 M46 32 L48 26" stroke="white" stroke-width="1.5" fill="none"/><circle cx="38" cy="35" r="1" fill="white"/><circle cx="42" cy="35" r="1" fill="white"/>'
    "星"='<polygon points="40,30 42,36 48,36 44,40 46,46 40,42 34,46 36,40 32,36 38,36" fill="white" fill-opacity="0.6"/>'
    "月"='<path d="M30 38 Q30 30 40 30 Q32 34 32 44 Q32 50 40 50 Q30 46 30 38" fill="white" fill-opacity="0.5"/>'
    "山"='<path d="M18 55 L35 30 L48 45 L60 25 L72 55" stroke="white" stroke-width="2" fill="none"/><line x1="18" y1="55" x2="72" y2="55" stroke="white" stroke-opacity="0.3" stroke-width="1"/>'
    "花"='<g fill="white" fill-opacity="0.4"><circle cx="36" cy="38" r="3"/><circle cx="40" cy="34" r="3"/><circle cx="44" cy="38" r="3"/><circle cx="36" cy="42" r="3"/><circle cx="44" cy="42" r="3"/></g><circle cx="40" cy="38" r="2.5" fill="white" fill-opacity="0.7"/>'
    "屋"='<path d="M25 50 L25 38 L40 28 L55 38 L55 50 Z" stroke="white" stroke-width="1.5" fill="none"/><rect x="34" y="42" width="12" height="8" rx="1" fill="white" fill-opacity="0.2"/>'
    "灯"='<rect x="35" y="45" width="10" height="4" rx="1" fill="white" fill-opacity="0.3"/><path d="M38 45 L37 38 L43 38 L42 45" fill="white" fill-opacity="0.2"/><circle cx="40" cy="35" r="4" fill="white" fill-opacity="0.6"/>'
    "钟"='<circle cx="40" cy="40" r="10" fill="none" stroke="white" stroke-opacity="0.5" stroke-width="1.5"/><line x1="40" y1="40" x2="40" y2="34" stroke="white" stroke-width="1"/><line x1="40" y1="40" x2="44" y2="40" stroke="white" stroke-width="1"/>'
    "车"='<rect x="22" y="40" width="36" height="12" rx="3" fill="white" fill-opacity="0.3"/><rect x="28" y="36" width="24" height="8" rx="2" fill="white" fill-opacity="0.2"/><circle cx="30" cy="53" r="3" fill="white" fill-opacity="0.3"/><circle cx="50" cy="53" r="3" fill="white" fill-opacity="0.3"/>'
    "书"='<rect x="30" y="28" width="20" height="24" rx="2" fill="white" fill-opacity="0.3"/><line x1="35" y1="33" x2="45" y2="33" stroke="white" stroke-width="1"/><line x1="35" y1="37" x2="45" y2="37" stroke="white" stroke-width="1"/>'
    "河"='<path d="M18 45 Q30 38 40 45 Q50 52 62 45" stroke="white" stroke-opacity="0.4" stroke-width="2" fill="none"/><path d="M18 48 Q30 41 40 48 Q50 55 62 48" stroke="white" stroke-opacity="0.2" stroke-width="1.5" fill="none"/>'
    "云"='<ellipse cx="35" cy="42" rx="10" ry="6" fill="white" fill-opacity="0.4"/><ellipse cx="45" cy="40" rx="8" ry="5" fill="white" fill-opacity="0.3"/><ellipse cx="38" cy="38" rx="6" ry="4" fill="white" fill-opacity="0.5"/>'
    "水"='<path d="M30 50 Q35 36 40 28 Q45 36 50 50" fill="white" fill-opacity="0.2" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/><line x1="34" y1="42" x2="46" y2="42" stroke="white" stroke-opacity="0.3" stroke-width="1"/>'
    "火"='<path d="M40 28 Q36 38 32 42 Q36 40 38 44 Q36 48 40 52 Q44 48 42 44 Q44 40 48 42 Q44 38 40 28" fill="white" fill-opacity="0.5"/>'
    "衣"='<path d="M30 35 L40 28 L50 35 L48 52 L32 52 Z" fill="white" fill-opacity="0.3" stroke="white" stroke-opacity="0.4" stroke-width="1"/>'
    "果"='<circle cx="40" cy="40" r="8" fill="white" fill-opacity="0.3"/><path d="M40 32 L40 28 M36 30 L44 30" stroke="white" stroke-opacity="0.4" stroke-width="1.5" fill="none"/>'
    "人"='<circle cx="40" cy="34" r="6" fill="white" fill-opacity="0.3"/><line x1="40" y1="40" x2="40" y2="50" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/><line x1="34" y1="44" x2="46" y2="44" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/>'
    "日"='<circle cx="40" cy="40" r="10" fill="white" fill-opacity="0.3"/><circle cx="40" cy="40" r="4" fill="white" fill-opacity="0.5"/><line x1="40" y1="26" x2="40" y2="30" stroke="white" stroke-opacity="0.5" stroke-width="1"/><line x1="40" y1="50" x2="40" y2="54" stroke="white" stroke-opacity="0.5" stroke-width="1"/><line x1="26" y1="40" x2="30" y2="40" stroke="white" stroke-opacity="0.5" stroke-width="1"/><line x1="50" y1="40" x2="54" y2="40" stroke="white" stroke-opacity="0.5" stroke-width="1"/>'
    "耳"='<ellipse cx="34" cy="40" rx="4" ry="7" fill="white" fill-opacity="0.3"/><ellipse cx="46" cy="40" rx="4" ry="7" fill="white" fill-opacity="0.3"/>'
    "鼻"='<ellipse cx="40" cy="42" rx="5" ry="4" fill="white" fill-opacity="0.3"/><circle cx="38" cy="38" r="2.5" fill="white" fill-opacity="0.2"/><circle cx="42" cy="38" r="2.5" fill="white" fill-opacity="0.2"/>'
    "足"='<ellipse cx="40" cy="48" rx="8" ry="3" fill="white" fill-opacity="0.3"/><line x1="40" y1="48" x2="34" y2="34" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/>'
    "手"='<ellipse cx="40" cy="42" rx="6" ry="4" fill="white" fill-opacity="0.3"/><line x1="28" y1="42" x2="52" y2="42" stroke="white" stroke-opacity="0.3" stroke-width="1"/>'
    "旗"='<line x1="30" y1="28" x2="30" y2="52" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/><rect x="30" y="28" width="16" height="10" fill="white" fill-opacity="0.3"/>'
    "鼓"='<rect x="32" y="38" width="16" height="14" rx="2" fill="white" fill-opacity="0.3"/><line x1="32" y1="45" x2="48" y2="45" stroke="white" stroke-opacity="0.4" stroke-width="1"/><circle cx="40" cy="45" r="2" fill="white" fill-opacity="0.5"/>'
    "球"='<circle cx="40" cy="40" r="8" fill="white" fill-opacity="0.3"/><path d="M34 35 Q40 40 46 35" stroke="white" stroke-width="1" fill="none"/><path d="M34 45 Q40 40 46 45" stroke="white" stroke-width="1" fill="none"/>'
    "梨"='<ellipse cx="40" cy="42" rx="7" ry="9" fill="white" fill-opacity="0.3"/><path d="M40 33 L40 30" stroke="white" stroke-opacity="0.4" stroke-width="1.5" fill="none"/>'
    "瓜"='<ellipse cx="40" cy="42" rx="10" ry="7" fill="white" fill-opacity="0.3"/><path d="M34 38 Q40 42 46 38" stroke="white" stroke-opacity="0.4" stroke-width="1" fill="none"/>'
    "米"='<line x1="40" y1="30" x2="40" y2="50" stroke="white" stroke-width="1.5"/><line x1="32" y1="36" x2="48" y2="36" stroke="white" stroke-width="1"/><line x1="32" y1="44" x2="48" y2="44" stroke="white" stroke-width="1"/>'
    "竹"='<line x1="36" y1="30" x2="34" y2="52" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/><line x1="44" y1="30" x2="46" y2="52" stroke="white" stroke-opacity="0.4" stroke-width="1.5"/><line x1="34" y1="38" x2="46" y2="38" stroke="white" stroke-opacity="0.3" stroke-width="1"/>'
    "桌"='<line x1="26" y1="38" x2="54" y2="38" stroke="white" stroke-opacity="0.4" stroke-width="2"/><line x1="30" y1="38" x2="28" y2="54" stroke="white" stroke-opacity="0.3" stroke-width="1.5"/><line x1="50" y1="38" x2="52" y2="54" stroke="white" stroke-opacity="0.3" stroke-width="1.5"/>'
  }
  if ($d.ContainsKey($last)) { return $d[$last] }
  if ($d.ContainsKey($word)) { return $d[$word] }
  $c = [int]([char]$last) % 4
  @('<circle cx="40" cy="40" r="10" fill="white" fill-opacity="0.25"/>','<rect x="32" y="32" width="16" height="16" rx="2" fill="white" fill-opacity="0.25"/>','<polygon points="40,30 30,50 50,50" fill="white" fill-opacity="0.25"/>','<path d="M32 40 Q40 28 48 40 Q40 52 32 40" fill="white" fill-opacity="0.25"/>')[$c]
}

function New-LessonIconsPng {
  param($lessonNum)
  $L2 = $Lessons[$lessonNum]
  $pairs = @()
  foreach ($hw in $L2.homework) {
    $m = [regex]::Matches($hw, '([\u4e00-\u9fff]{1,4})')
    foreach ($word in $m.Value) {
      if ($pairs.Count -ge 8) { break }
      if ($word -match '^(读|例|描|复|默|用|平|区|翘|前|后|整|写|拼|顺|复习|前后对比|区分|翘舌音|平翘舌|23|默写|声母|韵母|全部|名字|自己|的)') { continue }
      if ($pairs -notcontains $word) { $pairs += $word }
    }
  }
  if ($pairs.Count -eq 0) { return $null }
  $n = [Math]::Min($pairs.Count, 8); $w = 80 * $n; $h = 80
  $colors = @("#FF6B35","#2196F3","#4CAF50","#9C27B0","#FF9800","#E91E63","#00BCD4","#795548")
  $svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + $w + '" height="' + $h + '" viewBox="0 0 ' + $w + ' ' + $h + '"><rect width="' + $w + '" height="' + $h + '" fill="white"/>'
  for ($i = 0; $i -lt $n; $i++) {
    $word = $pairs[$i]; $bg = $colors[$i % 8]
    $svg += '<g transform="translate(' + ($i*80) + ',5)"><rect x="2" y="0" width="76" height="70" rx="10" fill="' + $bg + '"/>'
    $svg += Get-WordIllustration $word
    $svg += '<text x="40" y="62" text-anchor="middle" font-family="sans-serif" font-weight="bold" font-size="10" fill="white" fill-opacity="0.8">' + $word + '</text></g>'
  }
  $svg += '</svg>'
  $tmpDir = "$env:TEMP\pinyin_img"
  if (!(Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null }
  $htmlPath = "$tmpDir\lesson${lessonNum}.html"
  Set-Content -Path $htmlPath -Value "<!DOCTYPE html><html><body style='margin:0'>$svg</body></html>" -Encoding UTF8
  Add-Type -AssemblyName System.Windows.Forms,System.Drawing
  $wb = New-Object System.Windows.Forms.WebBrowser
  $wb.ScriptErrorsSuppressed = $true; $wb.ScrollBarsEnabled = $false
  $wb.Size = New-Object System.Drawing.Size($w, $h)
  $wb.Navigate("file:///$htmlPath")
  for ($t = 0; $t -lt 50 -and $wb.ReadyState -ne [System.Windows.Forms.WebBrowserReadyState]::Complete; $t++) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 100 }
  Start-Sleep -Milliseconds 300
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $wb.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0,0,$w,$h)))
  $wb.Dispose()
  $ms = New-Object System.IO.MemoryStream
  $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  $hex = [System.BitConverter]::ToString($ms.ToArray()) -replace '-',''
  $ms.Close(); Remove-Item $htmlPath -Force -ErrorAction SilentlyContinue
  return @{hex=$hex; w=$w; h=$h}
}

$lines = @()
$lines += "\pard\cf2\b\fs44 " + (Fmt "第${Lesson}课 - $($L.title) 拼音练习") + "\cf1\b0\par"
$lines += "\pard\cf6\fs24 " + (Fmt "$($L.subtitle)  生成日期: ${DateStr}") + "\cf1\par"
$lines += "\pard\cb3\cf8\b\fs36 " + (Fmt "朗读练习 - 大声读3遍") + "\cf1\b0\par"

foreach ($rd in $L.reading) {
  $lines += "\pard\fs36\b " + (Fmt $rd) + "\b0\par"
}

$lines += "\pard\cf5\b\fs28 " + (Fmt "顺口溜: $($L.jingle)") + "\cf1\b0\par"

$lines += "\pard\cb3\cf2\b\fs36 " + (Fmt "书写练习 - 描红") + "\cf1\b0\par"
$lines += "\pard\fs24 " + (Fmt "(在横线上描红, 注意占中格位置)") + "\par"

foreach ($ch in $L.chars) {
  $e = Fmt $ch
  $lines += "\pard\fs60\cf1\f1 ${e}\cf7       ${e}  ${e}  ${e}  ${e}\cf1\f0\par"
}

$lines += "\pard\cb3\cf4\b\fs36 " + (Fmt "课后练习") + "\cf1\b0\par"
$idx = 1
foreach ($hw in $L.homework) {
  $lines += "\pard\sl420\fs28 " + (Fmt "${idx}. ${hw}") + "\par"
  $idx++
}

# 连一连：拼音配汉字
$pairs = @()
foreach ($hw in $L.homework) {
  $m = [regex]::Matches($hw, '([a-zü:0-9]+)\(([\u4e00-\u9fff]+)\)')
  foreach ($match in $m) { $pairs += @{pinyin=$match.Groups[1].Value; word=$match.Groups[2].Value} }
}
if ($pairs.Count -ge 3) {
  $lines += "\pard\cb3\cf4\b\fs36 " + (Fmt "连一连：将拼音与对应的汉字连线") + "\cf1\b0\par"
  $shuffled = $pairs | Sort-Object { Get-Random }
  $lines += "{\trowd\cellx3500\cellx7000\intbl"
  $lines += "\pard\intbl\cf2\b\fs28 " + (Fmt "拼音") + "\cell \pard\intbl\cf2\b\fs28 " + (Fmt "汉字") + "\cell\row\b0\cf1"
  for ($i = 0; $i -lt $pairs.Count; $i++) {
    $py = (Convert-ToneMarks $pairs[$i].pinyin)
    $wd = $shuffled[$i].word
    $lines += "\pard\intbl\fs28\cf1\f1 " + (Escape-RtfUnicode $py) + "\cell \pard\intbl\fs28\cf7 " + (Escape-RtfUnicode $wd) + "\cell\row"
  }
  $lines += "}"
}

$lines += "\pard\fs20 " + (Fmt "生成日期: ${DateStr}") + "\par"

$header = "{\rtf1\ansi\ansicpg65001\deff0{\fonttbl{\f0\fnil\fcharset134 SimSun;}{\f1\fnil\fcharset0 Century Gothic;}}{\colortbl;$colors}\paperw11900\paperh16840\margl1134\margr1134\margt567\margb567\pard\f0\fs28"
$body = $lines -join ""
$rtf = $header + $body + "}"

$bytes = [System.Text.Encoding]::UTF8.GetBytes($rtf)

$path = Join-Path $printableDir "pinyin_lesson${Lesson}.doc"
[System.IO.File]::WriteAllBytes($path, $bytes)
Write-Output "生成拼音练习纸: $path"
