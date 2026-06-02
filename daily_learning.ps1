param(
  [DateTime]$RunDate = (Get-Date)
)

# ====== 配置 ======
$ProgressFile = "$PSScriptRoot\progress.json"
$LogFile = "$PSScriptRoot\learning.log"
$WorksheetDir = "$PSScriptRoot\worksheets"
$DuoDuoUrl = "https://flowus.cn/b80cd768-6ef5-4da6-a257-e2afe8d388ac"
$XiaoMingUrl = "https://flowus.cn/8169ce9a-efd8-47bd-b671-2d0b8fa4c6e6"
$MathUrl = "https://flowus.cn/47549cdb-db3c-4f67-b902-e0d947cdc60c"
$ChineseUrl = "https://flowus.cn/74179223-c084-4be9-a08f-0dbebe16a55f"
$ServerChanKey = "SCT357593T6RJQiHZRwcLRN61jTWM5Dg3L"

if (-not (Test-Path $WorksheetDir)) { New-Item -ItemType Directory -Path $WorksheetDir -Force | Out-Null }

function Write-Log {
  param($Msg)
  $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  "$time $Msg" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log "===== 开始执行 ====="

# ====== 读取或初始化进度 ======
if (-not (Test-Path $ProgressFile)) {
  $Progress = @{
    start_date = "2026-05-31"
    duo_duo = @{ current_week = 1; total_weeks = 12; name = "多多" }
    xiao_ming = @{ current_week = 1; total_weeks = 33; name = "小铭" }
    duo_duo_math = @{ current_week = 1; total_weeks = 31; name = "多多数学" }
    duo_duo_chinese = @{ current_week = 1; total_weeks = 13; name = "多多语文" }
  } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
} else {
  $Progress = Get-Content $ProgressFile -Encoding utf8 | ConvertFrom-Json
  if (-not $Progress.duo_duo_math) {
    $Progress | Add-Member -NotePropertyName "duo_duo_math" -NotePropertyValue ([PSCustomObject]@{ current_week = 1; total_weeks = 31; name = "多多数学" })
  }
  if (-not $Progress.duo_duo_chinese) {
    $Progress | Add-Member -NotePropertyName "duo_duo_chinese" -NotePropertyValue ([PSCustomObject]@{ current_week = 1; total_weeks = 13; name = "多多语文" })
  }
  if (-not $Progress.issues) {
    $Progress | Add-Member -NotePropertyName "issues" -NotePropertyValue ([PSCustomObject]@{})
  }
}

# ====== 计算当前周 ======
$StartDate = [DateTime]::Parse($Progress.start_date)
$DaysSinceStart = [Math]::Floor(($RunDate - $StartDate).TotalDays)

$RawWeek = [Math]::Floor($DaysSinceStart / 7) + 1
$DuoDuoWeek = [Math]::Min([Math]::Max(1, $RawWeek), $Progress.duo_duo.total_weeks)
$XiaoMingWeek = [Math]::Min([Math]::Max(1, $RawWeek), $Progress.xiao_ming.total_weeks)

Write-Log "多多: 第${DuoDuoWeek}周 / 小铭: 第${XiaoMingWeek}周"

# ====== 多多语文（拼音+识字 6-8月冲刺） ======
$ChineseWeek = [Math]::Min([Math]::Max(1, $RawWeek), 13)
$ChineseContent = @{}
# ====== G1（一年级 5言绝句起步）=====
$ChineseContent[1]  = @{pinyin="b p m f"; chars="人 口 手 大 小"; act="看拼音视频b p m f，识字卡:人口手大小"; poems=@(
  @{title="咏鹅"; author="骆宾王"; text="鹅鹅鹅，曲项向天歌。白毛浮绿水，红掌拨清波。"}
  @{title="静夜思"; author="李白"; text="床前明月光，疑是地上霜。举头望明月，低头思故乡。"}
  @{title="画"; author="王维"; text="远看山有色，近听水无声。春去花还在，人来鸟不惊。"}
  @{title="悯农（其一）"; author="李绅"; text="锄禾日当午，汗滴禾下土。谁知盘中餐，粒粒皆辛苦。"}
  @{title="悯农（其二）"; author="李绅"; text="春种一粒粟，秋收万颗子。四海无闲田，农夫犹饿死。"}
)}
$ChineseContent[2]  = @{pinyin="d t n l"; chars="上 下 天 地"; act="拼音跳房子，在家找认识的字"; poems=@(
  @{title="春晓"; author="孟浩然"; text="春眠不觉晓，处处闻啼鸟。夜来风雨声，花落知多少。"}
  @{title="古朗月行（节选）"; author="李白"; text="小时不识月，呼作白玉盘。又疑瑶台镜，飞在青云端。"}
  @{title="池上"; author="白居易"; text="小娃撑小艇，偷采白莲回。不解藏踪迹，浮萍一道开。"}
  @{title="江南"; author="汉乐府"; text="江南可采莲，莲叶何田田。鱼戏莲叶间。鱼戏莲叶东，鱼戏莲叶西，鱼戏莲叶南，鱼戏莲叶北。"}
  @{title="风"; author="李峤"; text="解落三秋叶，能开二月花。过江千尺浪，入竹万竿斜。"}
)}
$ChineseContent[3]  = @{pinyin="g k h j q x"; chars="一 二 三 四 五"; act="拼音卡片抢答，数字描红"; poems=@(
  @{title="小池"; author="杨万里"; text="泉眼无声惜细流，树阴照水爱晴柔。小荷才露尖尖角，早有蜻蜓立上头。"}
  @{title="寻隐者不遇"; author="贾岛"; text="松下问童子，言师采药去。只在此山中，云深不知处。"}
  @{title="所见"; author="袁枚"; text="牧童骑黄牛，歌声振林樾。意欲捕鸣蝉，忽然闭口立。"}
  @{title="村居"; author="高鼎"; text="草长莺飞二月天，拂堤杨柳醉春烟。儿童散学归来早，忙趁东风放纸鸢。"}
  @{title="登鹳雀楼"; author="王之涣"; text="白日依山尽，黄河入海流。欲穷千里目，更上一层楼。"}
)}
# ====== G2（二年级 7言渐多）=====
$ChineseContent[4]  = @{pinyin="zh ch sh r z c s y w"; chars="全部声母复习"; act="所有声母总复习，BINGO游戏"; poems=@(
  @{title="望庐山瀑布"; author="李白"; text="日照香炉生紫烟，遥看瀑布挂前川。飞流直下三千尺，疑是银河落九天。"}
  @{title="敕勒歌"; author="北朝民歌"; text="敕勒川，阴山下。天似穹庐，笼盖四野。天苍苍，野茫茫，风吹草低见牛羊。"}
  @{title="咏柳"; author="贺知章"; text="碧玉妆成一树高，万条垂下绿丝绦。不知细叶谁裁出，二月春风似剪刀。"}
  @{title="赠汪伦"; author="李白"; text="李白乘舟将欲行，忽闻岸上踏歌声。桃花潭水深千尺，不及汪伦送我情。"}
  @{title="绝句（迟日江山）"; author="杜甫"; text="迟日江山丽，春风花草香。泥融飞燕子，沙暖睡鸳鸯。"}
)}
$ChineseContent[5]  = @{pinyin="a o e i u ü"; chars="日 月 水 火 山 石 田 土"; act="韵母歌，字卡配图"; poems=@(
  @{title="绝句（两个黄鹂）"; author="杜甫"; text="两个黄鹂鸣翠柳，一行白鹭上青天。窗含西岭千秋雪，门泊东吴万里船。"}
  @{title="望天门山"; author="李白"; text="天门中断楚江开，碧水东流至此回。两岸青山相对出，孤帆一片日边来。"}
  @{title="江雪"; author="柳宗元"; text="千山鸟飞绝，万径人踪灭。孤舟蓑笠翁，独钓寒江雪。"}
  @{title="山行"; author="杜牧"; text="远上寒山石径斜，白云生处有人家。停车坐爱枫林晚，霜叶红于二月花。"}
  @{title="回乡偶书"; author="贺知章"; text="少小离家老大回，乡音无改鬓毛衰。儿童相见不相识，笑问客从何处来。"}
)}
$ChineseContent[6]  = @{pinyin="a o e i u ü 巩固"; chars="花 鸟 虫 鱼 马 牛 羊 狗 猫"; act="动物字卡游戏"; poems=@(
  @{title="早发白帝城"; author="李白"; text="朝辞白帝彩云间，千里江陵一日还。两岸猿声啼不住，轻舟已过万重山。"}
  @{title="饮湖上初晴后雨"; author="苏轼"; text="水光潋滟晴方好，山色空蒙雨亦奇。欲把西湖比西子，淡妆浓抹总相宜。"}
  @{title="惠崇春江晚景"; author="苏轼"; text="竹外桃花三两枝，春江水暖鸭先知。蒌蒿满地芦芽短，正是河豚欲上时。"}
  @{title="元日"; author="王安石"; text="爆竹声中一岁除，春风送暖入屠苏。千门万户曈曈日，总把新桃换旧符。"}
  @{title="清明"; author="杜牧"; text="清明时节雨纷纷，路上行人欲断魂。借问酒家何处有，牧童遥指杏花村。"}
)}
# ====== G3-G4 ======
$ChineseContent[7]  = @{pinyin="ai ei ui"; chars="新字4个"; act="韵母配对游戏"; poems=@(
  @{title="九月九日忆山东兄弟"; author="王维"; text="独在异乡为异客，每逢佳节倍思亲。遥知兄弟登高处，遍插茱萸少一人。"}
  @{title="凉州词"; author="王之涣"; text="黄河远上白云间，一片孤城万仞山。羌笛何须怨杨柳，春风不度玉门关。"}
  @{title="枫桥夜泊"; author="张继"; text="月落乌啼霜满天，江枫渔火对愁眠。姑苏城外寒山寺，夜半钟声到客船。"}
  @{title="凉州词"; author="王翰"; text="葡萄美酒夜光杯，欲饮琵琶马上催。醉卧沙场君莫笑，古来征战几人回。"}
  @{title="出塞"; author="王昌龄"; text="秦时明月汉时关，万里长征人未还。但使龙城飞将在，不教胡马度阴山。"}
)}
$ChineseContent[8]  = @{pinyin="ao ou iu ie ue er"; chars="复习全部"; act="全部韵母大闯关"; poems=@(
  @{title="鹿柴"; author="王维"; text="空山不见人，但闻人语响。返景入深林，复照青苔上。"}
  @{title="别董大"; author="高适"; text="千里黄云白日曛，北风吹雁雪纷纷。莫愁前路无知己，天下谁人不识君。"}
  @{title="题西林壁"; author="苏轼"; text="横看成岭侧成峰，远近高低各不同。不识庐山真面目，只缘身在此山中。"}
  @{title="夏日绝句"; author="李清照"; text="生当作人杰，死亦为鬼雄。至今思项羽，不肯过江东。"}
  @{title="墨梅"; author="王冕"; text="吾家洗砚池头树，朵朵花开淡墨痕。不要人夸好颜色，只留清气满乾坤。"}
)}
# ====== G4→G5 ======
$ChineseContent[9]  = @{pinyin="zhi chi shi ri zi ci si"; chars="复习"; act="整体认读音节认读"; poems=@(
  @{title="送元二使安西"; author="王维"; text="渭城朝雨浥轻尘，客舍青青柳色新。劝君更尽一杯酒，西出阳关无故人。"}
  @{title="黄鹤楼送孟浩然之广陵"; author="李白"; text="故人西辞黄鹤楼，烟花三月下扬州。孤帆远影碧空尽，唯见长江天际流。"}
  @{title="独坐敬亭山"; author="李白"; text="众鸟高飞尽，孤云独去闲。相看两不厌，只有敬亭山。"}
  @{title="泊船瓜洲"; author="王安石"; text="京口瓜洲一水间，钟山只隔数重山。春风又绿江南岸，明月何时照我还。"}
  @{title="示儿"; author="陆游"; text="死去元知万事空，但悲不见九州同。王师北定中原日，家祭无忘告乃翁。"}
)}
$ChineseContent[10] = @{pinyin="yi wu yu ye yue yuan yin yun ying"; chars="复习"; act="绕口令游戏"; poems=@(
  @{title="题临安邸"; author="林升"; text="山外青山楼外楼，西湖歌舞几时休。暖风熏得游人醉，直把杭州作汴州。"}
  @{title="秋夜将晓出篱门迎凉有感"; author="陆游"; text="三万里河东入海，五千仞岳上摩天。遗民泪尽胡尘里，南望王师又一年。"}
  @{title="四时田园杂兴（其三十一）"; author="范成大"; text="昼出耘田夜绩麻，村庄儿女各当家。童孙未解供耕织，也傍桑阴学种瓜。"}
  @{title="晓出净慈寺送林子方"; author="杨万里"; text="毕竟西湖六月中，风光不与四时同。接天莲叶无穷碧，映日荷花别样红。"}
  @{title="蜂"; author="罗隐"; text="不论平地与山尖，无限风光尽被占。采得百花成蜜后，为谁辛苦为谁甜。"}
)}
# ====== G5 ======
$ChineseContent[11] = @{pinyin="ba ma di tu fo gu ka"; chars="拼读组合"; act="拼音+字组合拼读"; poems=@(
  @{title="江上渔者"; author="范仲淹"; text="江上往来人，但爱鲈鱼美。君看一叶舟，出没风波里。"}
  @{title="游园不值"; author="叶绍翁"; text="应怜屐齿印苍苔，小扣柴扉久不开。春色满园关不住，一枝红杏出墙来。"}
  @{title="江南春"; author="杜牧"; text="千里莺啼绿映红，水村山郭酒旗风。南朝四百八十寺，多少楼台烟雨中。"}
  @{title="观书有感"; author="朱熹"; text="半亩方塘一鉴开，天光云影共徘徊。问渠那得清如许，为有源头活水来。"}
  @{title="石灰吟"; author="于谦"; text="千锤万凿出深山，烈火焚烧若等闲。粉骨碎身全不怕，要留清白在人间。"}
)}
# ====== G5→G6（最长、最难）=====
$ChineseContent[12] = @{pinyin="复习全部拼音"; chars="前100字总复习(一)"; act="字卡全部过一遍"; poems=@(
  @{title="春夜喜雨"; author="杜甫"; text="好雨知时节，当春乃发生。随风潜入夜，润物细无声。"}
  @{title="竹石"; author="郑燮"; text="咬定青山不放松，立根原在破岩中。千磨万击还坚劲，任尔东西南北风。"}
  @{title="己亥杂诗"; author="龚自珍"; text="浩荡离愁白日斜，吟鞭东指即天涯。落红不是无情物，化作春泥更护花。"}
  @{title="七步诗"; author="曹植"; text="煮豆燃豆萁，豆在釜中泣。本是同根生，相煎何太急。"}
  @{title="长歌行（节选）"; author="汉乐府"; text="青青园中葵，朝露待日晞。阳春布德泽，万物生光辉。常恐秋节至，焜黄华叶衰。百川东到海，何时复西归。少壮不努力，老大徒伤悲。"}
)}
$ChineseContent[13] = @{pinyin="复习全部拼音"; chars="前100字总复习(二)+写名字"; act="学写自己名字"; poems=@(
  @{title="书湖阴先生壁"; author="王安石"; text="茅檐长扫净无苔，花木成畦手自栽。一水护田将绿绕，两山排闼送青来。"}
  @{title="六月二十七日望湖楼醉书"; author="苏轼"; text="黑云翻墨未遮山，白雨跳珠乱入船。卷地风来忽吹散，望湖楼下水如天。"}
  @{title="芙蓉楼送辛渐"; author="王昌龄"; text="寒雨连江夜入吴，平明送客楚山孤。洛阳亲友如相问，一片冰心在玉壶。"}
  @{title="送杜少府之任蜀州"; author="王勃"; text="城阙辅三秦，风烟望五津。与君离别意，同是宦游人。海内存知己，天涯若比邻。无为在歧路，儿女共沾巾。"}
  @{title="望月怀远"; author="张九龄"; text="海上生明月，天涯共此时。情人怨遥夜，竟夕起相思。"}
)}

$C = $ChineseContent[[int]$ChineseWeek]

# ====== 每日语文内容分配 ======
$Dow = [int]$RunDate.DayOfWeek
$DayIndex = if ($Dow -eq 0 -or $Dow -eq 6) { 5 } else { $Dow }

# ====== 选当日古诗词（每日 2 首，轮转展示） ======
$cnPoems = $C.poems
$ChinesePoem1Title = ""; $ChinesePoem1Author = ""; $ChinesePoem1Text = ""
$ChinesePoem2Title = ""; $ChinesePoem2Author = ""; $ChinesePoem2Text = ""
if ($cnPoems -and $cnPoems.Count -ge 2) {
    $idx1 = (($DayIndex - 1) * 2) % $cnPoems.Count
    $idx2 = (($DayIndex - 1) * 2 + 1) % $cnPoems.Count
    $ChinesePoem1Title = $cnPoems[$idx1].title
    $ChinesePoem1Author = $cnPoems[$idx1].author
    $ChinesePoem1Text = $cnPoems[$idx1].text
    $ChinesePoem2Title = $cnPoems[$idx2].title
    $ChinesePoem2Author = $cnPoems[$idx2].author
    $ChinesePoem2Text = $cnPoems[$idx2].text
} elseif ($cnPoems -and $cnPoems.Count -eq 1) {
    $ChinesePoem1Title = $cnPoems[0].title
    $ChinesePoem1Author = $cnPoems[0].author
    $ChinesePoem1Text = $cnPoems[0].text
}

function Get-ChineseDailyAct {
  param($C, $DayIndex)
  $pList = @($C.pinyin -split '\s+') | Where-Object { $_ -and $_ -notmatch '巩固|复习|韵母|声母|全部|整体' }
  $cText = $C.chars
  $hasReview = $cText -match '复习|总'
  $pCount = $pList.Count
  $mP = [Math]::Max(1, [Math]::Ceiling($pCount / 2))

  $parts = @()
  switch ($DayIndex) {
    1 {
      if ($pCount -gt 0) { $parts += "视频学: $($pList[0..($mP-1)] -join ' ') 跟读" }
      elseif ($cText -match '整体认读|拼读') { $parts += "学 $($C.pinyin) 跟读" }
      else { $parts += "复习 $($C.pinyin)" }
      if (-not $hasReview -and $cText -notmatch '复习|字母|声母|韵母') { $parts += "新字认+描红: $cText" }
      else { $parts += $cText }
      $parts += "读古诗: $($C.poems[0].title)"
      break
    }
    2 {
      if ($pCount -gt $mP) { $parts += "复习→学: $($pList[$mP..($pCount-1)] -join ' ')" }
      elseif ($pCount -gt 0) { $parts += "复习巩固: $($C.pinyin)" }
      else { $parts += "复习: $($C.pinyin)" }
      if (-not $hasReview -and $cText -notmatch '复习|字母|声母|韵母') { $parts += "新字: $cText 在家找" }
      $parts += "读古诗: $($C.poems[[Math]::Min(1, $C.poems.Count-1)].title)"
      break
    }
    3 { $parts += "总复习拼音: $($C.pinyin)"; $parts += "字卡配对: $cText"; $parts += "背古诗: $($C.poems[[Math]::Min(2, $C.poems.Count-1)].title)"; break }
    4 { $parts += "拼音抢答PK"; $parts += "字卡闪卡挑战: $cText"; $parts += "背古诗: $($C.poems[[Math]::Min(3, $C.poems.Count-1)].title)"; $parts += "在家找到认识的字"; break }
    5 { $parts += "本周总复习+小测验"; $parts += "全部拼音和字过一遍"; $parts += "背本周所有古诗"; $parts += "奖励贴纸!"; break }
  }
  return $parts -join ' | '
}

$DailyAct = Get-ChineseDailyAct -C $C -DayIndex $DayIndex

# ====== 多多内容（第3册 长元音） ======
$DuoDuoContent = @{}
$DuoDuoContent[1]  = @{letter="a_e"; words="cake, name, game, make"; act="烤蛋糕学a_e，用字母磁铁拼单词"}
$DuoDuoContent[2]  = @{letter="a_e 巩固"; words="late, snake, plane, plate"; act="做纸飞机写单词，画画配对"}
$DuoDuoContent[3]  = @{letter="i_e"; words="kite, bike, ride, five"; act="放风筝学i_e，骑自行车模仿"}
$DuoDuoContent[4]  = @{letter="i_e 巩固"; words="nine, like, hide, time"; act="躲猫猫hide游戏，数字卡片"}
$DuoDuoContent[5]  = @{letter="复习 a_e + i_e"; words="复习"; act="单词BINGO、拼读比赛"}
$DuoDuoContent[6]  = @{letter="o_e"; words="home, nose, rope, hope"; act="用绳子摆字母，画鼻子游戏"}
$DuoDuoContent[7]  = @{letter="o_e 巩固"; words="note, bone, rose, stone"; act="收集石头写单词"}
$DuoDuoContent[8]  = @{letter="u_e"; words="cube, tube, cute, mule"; act="用冰块cube学u_e，积木搭单词"}
$DuoDuoContent[9]  = @{letter="u_e 巩固"; words="rule, flute, June, tune"; act="唱曲子tune，用尺子rule拼读"}
$DuoDuoContent[10] = @{letter="复习 o_e + u_e"; words="复习"; act="单词接龙、闪卡挑战"}
$DuoDuoContent[11] = @{letter="e_e"; words="these, eve, here, theme"; act="读句子找e_e单词"}
$DuoDuoContent[12] = @{letter="总复习"; words="全部长元音"; act="阅读小故事、拼读大闯关"}

$D = $DuoDuoContent[[int]$DuoDuoWeek]

# ====== 小铭内容（第1册 字母音） ======
$XiaoMingContent = @{}
$XiaoMingContent[1]  = @{letter="Aa"; words="apple, ant, alligator"; act="苹果拓印画A，唱Apple Song"}
$XiaoMingContent[2]  = @{letter="Bb"; words="bear, bird, book"; act="小熊玩偶学Bb，豆子摆字母B"}
$XiaoMingContent[3]  = @{letter="Cc"; words="cat, cup, car"; act="卡纸剪字母C，猫捉杯子游戏"}
$XiaoMingContent[4]  = @{letter="Dd"; words="dog, duck, doll"; act="橡皮泥捏字母D和小狗"}
$XiaoMingContent[5]  = @{letter="复习A-D"; words="-"; act="字母BINGO、拍苍蝇大赛"}
$XiaoMingContent[6]  = @{letter="Ee"; words="egg, elbow, elephant"; act="鸡蛋盒手工，Elephant Song"}
$XiaoMingContent[7]  = @{letter="Ff"; words="fish, frog, flower"; act="羽毛沾颜料写F，折纸青蛙"}
$XiaoMingContent[8]  = @{letter="Gg"; words="girl, goat, gift"; act="包礼物盒贴G，Goat歌"}
$XiaoMingContent[9]  = @{letter="Hh"; words="hat, hand, horse"; act="做字母H帽子，手掌印画"}
$XiaoMingContent[10] = @{letter="复习E-H"; words="-"; act="字母钓鱼、跳房子认字母"}
$XiaoMingContent[11] = @{letter="Ii"; words="igloo, insect, ink"; act="冰块搭Igloo，手指画小虫"}
$XiaoMingContent[12] = @{letter="Jj"; words="juice, jar, jet"; act="喝果汁学Jj，折纸飞机"}
$XiaoMingContent[13] = @{letter="Kk"; words="kite, koala, key"; act="做风筝手工，钥匙拓印"}
$XiaoMingContent[14] = @{letter="Ll"; words="lion, lamp, leaf"; act="树叶贴字母L，学狮子吼"}
$XiaoMingContent[15] = @{letter="复习I-L"; words="-"; act="字母接龙、藏卡片找字母"}
$XiaoMingContent[16] = @{letter="Mm"; words="monkey, milk, moon"; act="喝牛奶画M，学猴子动作"}
$XiaoMingContent[17] = @{letter="Nn"; words="nest, nut, nose"; act="树枝搭鸟巢，贴鼻子游戏"}
$XiaoMingContent[18] = @{letter="Oo"; words="ox, octopus, orange"; act="吃橙子学Oo，橡皮泥搓圆"}
$XiaoMingContent[19] = @{letter="Pp"; words="pig, pen, pizza"; act="泡泡纸印P，画小粉猪"}
$XiaoMingContent[20] = @{letter="复习M-P"; words="-"; act="字母保龄球、跳跳认字母"}
$XiaoMingContent[21] = @{letter="Qq"; words="queen, quilt, question"; act="做纸皇冠，碎布贴字母Q"}
$XiaoMingContent[22] = @{letter="Rr"; words="rabbit, rug, robot"; act="学兔子跳Rr，纸巾筒机器人"}
$XiaoMingContent[23] = @{letter="Ss"; words="sun, sock, sandwich"; act="画太阳Ss，沙子写字母"}
$XiaoMingContent[24] = @{letter="Tt"; words="tea, tiger, table"; act="喝茶时间学Tt，画老虎条纹"}
$XiaoMingContent[25] = @{letter="复习Q-T"; words="-"; act="字母拼图大赛、你说我拍"}
$XiaoMingContent[26] = @{letter="Uu"; words="umbrella, up, under"; act="用雨伞学Uu，往上跳喊up"}
$XiaoMingContent[27] = @{letter="Vv"; words="violin, van, vet"; act="蔬菜印章印V，学小提琴"}
$XiaoMingContent[28] = @{letter="Ww"; words="water, watch, web"; act="水彩画W，毛线织蜘蛛网"}
$XiaoMingContent[29] = @{letter="Xx"; words="fox, box, wax"; act="盒子做X形状，蜡笔画X"}
$XiaoMingContent[30] = @{letter="Yy"; words="yellow, yo-yo, yarn"; act="黄色颜料画Y，玩悠悠球"}
$XiaoMingContent[31] = @{letter="Zz"; words="zebra, zoo, zero"; act="画斑马条纹Z，蜜蜂嗡嗡声"}
$XiaoMingContent[32] = @{letter="总复习"; words="26个字母"; act="字母大游行、闪卡挑战"}
$XiaoMingContent[33] = @{letter="总复习"; words="26个字母"; act="字母大游行、唱字母歌"}

$X = $XiaoMingContent[[int]$XiaoMingWeek]

# ====== 多多数学（新加坡数学 G1 共31周） ======
$MathContent = @{}
$MathContent[1]  = @{ch="Ch1 Numbers to 10"; topic="数1-10 认读写 比较多少"; act="🎲 骰子 🖐️ 手指 🧸 乐高塔"}
$MathContent[2]  = @{ch="Ch1 Numbers to 10"; topic="数1-10 认读写 比较多少"; act="🎲 骰子 🖐️ 手指 🧸 乐高塔"}
$MathContent[3]  = @{ch="Ch2 Number Bonds"; topic="整体与部分 10以内分解"; act="🧩 手指分 🥣 扣子分碗 🏠 数字房子"}
$MathContent[4]  = @{ch="Ch2 Number Bonds"; topic="整体与部分 10以内分解"; act="🧩 手指分 🥣 扣子分碗 🏠 数字房子"}
$MathContent[5]  = @{ch="Ch3 Addition ≤10"; topic="加法含义 交换律 加法表"; act="🐸 跳数轴 🎯 投篮计分 🃏 闪卡"}
$MathContent[6]  = @{ch="Ch3 Addition ≤10"; topic="加法含义 交换律 加法表"; act="🐸 跳数轴 🎯 投篮计分 🃏 闪卡"}
$MathContent[7]  = @{ch="Ch4 Subtraction ≤10"; topic="减法含义 加减互逆 减法表"; act="🍪 吃饼干 🎈 气球飞 🐛 倒退"}
$MathContent[8]  = @{ch="Ch4 Subtraction ≤10"; topic="减法含义 加减互逆 减法表"; act="🍪 吃饼干 🎈 气球飞 🐛 倒退"}
$MathContent[9]  = @{ch="R1 复习 Ch1-4"; topic="数字 分解 加减法"; act="🔄 数字歌→分解游戏→闪卡混练"}
$MathContent[10] = @{ch="Ch5-6 Shapes & Ordinal"; topic="图形 规律 第1-第10"; act="✂️ 剪纸 🏠 寻宝 🚂 火车"}
$MathContent[11] = @{ch="Ch5-6 Shapes & Ordinal"; topic="图形 规律 第1-第10"; act="✂️ 剪纸 🏠 寻宝 🚂 火车"}
$MathContent[12] = @{ch="Ch7 Numbers to 20"; topic="11-20 十位个位 顺序"; act="🧮 十格框 📿 串珠 🎲 比大小"}
$MathContent[13] = @{ch="Ch7 Numbers to 20"; topic="11-20 十位个位 顺序"; act="🧮 十格框 📿 串珠 🎲 比大小"}
$MathContent[14] = @{ch="Ch8 Add/Sub ≤20"; topic="两位数±一位数 凑十法"; act="💰 硬币 🎪 套圈 🧮 接龙"}
$MathContent[15] = @{ch="Ch8 Add/Sub ≤20"; topic="两位数±一位数 凑十法"; act="💰 硬币 🎪 套圈 🧮 接龙"}
$MathContent[16] = @{ch="R2 复习 1A全册"; topic="Ch1-8全面回顾"; act="🔄 薄弱环节闪卡战"}
$MathContent[17] = @{ch="Ch9-10 Length & Weight"; topic="比较长短轻重 cm kg"; act="✋ 拃 📏 身高 ⚖️ 天平"}
$MathContent[18] = @{ch="Ch9-10 Length & Weight"; topic="比较长短轻重 cm kg"; act="✋ 拃 📏 身高 ⚖️ 天平"}
$MathContent[19] = @{ch="Ch11-12 Graphs & 数≤40"; topic="读图 统计 21-40"; act="🍎 水果统计 🧮 十格框"}
$MathContent[20] = @{ch="Ch11-12 Graphs & 数≤40"; topic="读图 统计 21-40"; act="🍎 水果统计 🧮 十格框"}
$MathContent[21] = @{ch="Ch13 Add/Sub ≤40"; topic="两位数±个位/整十 竖式"; act="🧱 积木 🧮 计数器 🏦 存钱罐"}
$MathContent[22] = @{ch="Ch13 Add/Sub ≤40"; topic="两位数±个位/整十 竖式"; act="🧱 积木 🧮 计数器 🏦 存钱罐"}
$MathContent[23] = @{ch="R3 复习 Ch9-13"; topic="测量 统计 数 加减"; act="🔄 混练闪卡"}
$MathContent[24] = @{ch="Ch14-15 ×÷入门"; topic="重复加 均分 2/5/10倍"; act="👟 鞋配对 🍬 排糖 🍪 分饼"}
$MathContent[25] = @{ch="Ch14-15 ×÷入门"; topic="重复加 均分 2/5/10倍"; act="👟 鞋配对 🍬 排糖 🍪 分饼"}
$MathContent[26] = @{ch="Ch16 Time"; topic="整点/半点 推算 一周"; act="⏰ 纸盘钟 🎬 行程 🗓️ 计划"}
$MathContent[27] = @{ch="Ch17 Numbers to 100"; topic="41-100 百数表 排序"; act="🧮 百数表 🏗️ 积木100 🔢 猜数"}
$MathContent[28] = @{ch="Ch18 Add/Sub ≤100"; topic="两位数±整十/一位"; act="💰 钱包 🛒 购物 🧮 算盘"}
$MathContent[29] = @{ch="Ch19 Money"; topic="纸币硬币 兑换 找零"; act="🏪 商店 💵 凑钱 🧾 找零"}
$MathContent[30] = @{ch="R4 复习 1B全册"; topic="Ch9-19全面回顾"; act="🔄 强化薄弱环节"}
$MathContent[31] = @{ch="R4 复习 1B全册"; topic="Ch9-19全面回顾"; act="🔄 强化薄弱环节"}
$MathWeek = [Math]::Min([Math]::Max(1, $RawWeek), 31)
$M = $MathContent[[int]$MathWeek]

# ====== 记录每日学习内容 ======
$DateStr = Get-Date -Format "yyyy-MM-dd"
$DayLog = "  ===== $DateStr 第$([Math]::Floor($DaysSinceStart / 7) + 1)周 ====="
$DayLog += "`n  多多（第${DuoDuoWeek}周）：$($D.letter) 单词：$($D.words) 活动：$($D.act)"
$DayLog += "`n  小铭（第${XiaoMingWeek}周）：$($X.letter) 单词：$($X.words) 活动：$($X.act)"
$DayLog += "`n  多多数学（第${MathWeek}周）：$($M.ch) 内容：$($M.topic) 活动：$($M.act)"
$DayLog += "`n  多多语文（第${ChineseWeek}周）：$($C.pinyin) $($C.chars) 每日：$DailyAct"
Out-File -FilePath $LogFile -Encoding utf8 -Append -InputObject $DayLog

# ====== 生成每日练习卡片（Markdown） ======
function New-Worksheet {
  param($ChildName, $Week, $Content)
  $letter = $Content.letter
  $words = $Content.words
  $act = $Content.act
  $wordList = $words -split ',\s*'

  $lines = @()
  $lines += "# $ChildName 英语练习 - $DateStr"
  $lines += ""
  $lines += "## 本周学习：第${Week}周 - $letter"
  $lines += ""
  $lines += "### 单词跟读"
  $lines += ""
  foreach ($w in $wordList) {
    $lines += "- [ ] $w"
  }
  $lines += ""
  $lines += "### 活动"
  $lines += ""
  $lines += "- $act"
  $lines += ""
  $lines += "### 写一写"
  $lines += ""
  foreach ($w in $wordList) {
    $lines += "| $w | | | |"
  }
  $lines += ""
  $lines += "---"
  $lines += "*每天15分钟，开心学习！*"

  return $lines -join "`n"
}

$DuoDuoSheet = New-Worksheet -ChildName "多多" -Week $DuoDuoWeek -Content $D
$DuoDuoSheet | Out-File -FilePath "$WorksheetDir\多多_第${DuoDuoWeek}周_${DateStr}.md" -Encoding utf8

$XiaoMingSheet = New-Worksheet -ChildName "小铭" -Week $XiaoMingWeek -Content $X
$XiaoMingSheet | Out-File -FilePath "$WorksheetDir\小铭_第${XiaoMingWeek}周_${DateStr}.md" -Encoding utf8

# ====== 生成数学练习卡片 ======
function New-MathWorksheet {
  param($ChildName, $Week, $Content)
  $ch = $Content.ch
  $topic = $Content.topic
  $act = $Content.act

  $lines = @()
  $lines += "# $ChildName 数学练习 - $DateStr"
  $lines += ""
  $lines += "## 本周学习：第${Week}周 - $ch"
  $lines += ""
  $lines += "### 学习内容"
  $lines += ""
  $lines += "- $topic"
  $lines += ""
  $lines += "### 活动"
  $lines += ""
  $lines += "- $act"
  $lines += ""
  $lines += "### 写一写 / 算一算"
  $lines += ""
  $lines += "| 题目 | 答案 |"
  $lines += "|------|------|"
  $lines += "| | |"
  $lines += "| | |"
  $lines += "| | |"
  $lines += "| | |"
  $lines += "| | |"
  $lines += ""
  $lines += "---"
  $lines += "*每天15分钟，开心学习！*"

  return $lines -join "`n"
}

$MathSheet = New-MathWorksheet -ChildName "多多数学" -Week $MathWeek -Content $M
$MathSheet | Out-File -FilePath "$WorksheetDir\多多数学_第${MathWeek}周_${DateStr}.md" -Encoding utf8

# ====== 生成语文练习卡片 ======
function New-ChineseWorksheet {
  param($ChildName, $Week, $Content, $DayIdx, $DailyAct)
  $lines = @()
  $lines += "# $ChildName 语文练习 - $DateStr"
  $lines += ""
  $lines += "## 第${Week}周 · $($Content.pinyin) | $($Content.chars)"
  $lines += ""
  $lines += "### 今日任务"
  $lines += ""
  $lines += "- $DailyAct"
  $lines += ""
  $lines += "### 拼音写一写"
  $lines += ""
  $lines += "| 拼音 | 写3遍 |"
  $lines += "|------|-------|"
  foreach ($p in ($Content.pinyin -split '\s+')) { if ($p) { $lines += "| $p | |" } }
  $lines += ""
  $lines += "### 汉字描红"
  $lines += ""
  $lines += "| 汉字 | 描红 |"
  $lines += "|------|------|"
  $cList = @($Content.chars -split '\s+') | Where-Object { $_ -and $_ -notmatch '复习|字母|声母|韵母|组合|全部|新字' }
  foreach ($c in $cList) { $lines += "| $c | |" }
  $lines += ""
  $lines += "### 今日古诗"
  $lines += ""
  $poems = $Content.poems
  if ($poems) {
    $pIdx = [Math]::Max(0, [Math]::Min($DayIdx - 1, $poems.Count - 1))
    $lines += "**" + $poems[$pIdx].title + " - " + $poems[$pIdx].author + "**"
    $lines += ""
    $lines += $poems[$pIdx].text
  }
  $lines += ""
  $lines += "### 今日完成"
  $lines += ""
  $lines += "- [ ] 拼音跟读3遍"
  $lines += "- [ ] 汉字认读"
  $lines += "- [ ] 描红练习"
  $lines += "- [ ] 读古诗"
  $lines += "- [ ] 在家找字"
  $lines += ""
  $lines += "---"
  $lines += "*每天15分钟，开开心心学语文！*"
  return $lines -join "`n"
}

$ChineseSheet = New-ChineseWorksheet -ChildName "多多语文" -Week $ChineseWeek -Content $C -DayIdx $DayIndex -DailyAct $DailyAct
$ChineseSheet | Out-File -FilePath "$WorksheetDir\多多语文_第${ChineseWeek}周_${DateStr}.md" -Encoding utf8

Write-Log "练习卡片已生成"

# ====== SVG 打印练习纸 ======
$svgScript = Join-Path $PSScriptRoot "generate_svg.ps1"
if (Test-Path $svgScript) {
  try {
    & $svgScript -DuoDuoWeek $DuoDuoWeek -XiaoMingWeek $XiaoMingWeek -DDLetter $D.letter -DDWords $D.words -DDAct $D.act -XMLetter $X.letter -XMWords $X.words -XMAct $X.act -ChineseWeek $ChineseWeek -ChinesePinyin $C.pinyin -ChineseChars $C.chars -ChineseAct $C.act -ChinesePoem1Title $ChinesePoem1Title -ChinesePoem1Author $ChinesePoem1Author -ChinesePoem1Text $ChinesePoem1Text -ChinesePoem2Title $ChinesePoem2Title -ChinesePoem2Author $ChinesePoem2Author -ChinesePoem2Text $ChinesePoem2Text -WorksheetDir $WorksheetDir
  } catch {
    Write-Log ("SVG生成失败: " + $_.Exception.Message)
  }
  } else {
    Write-Log "SVG生成脚本不存在: $svgScript"
  }

  # ====== DOCX 打印文档 ======
  $docxScript = Join-Path $PSScriptRoot "generate_docx.ps1"
  if (Test-Path $docxScript) {
    try {
      & $docxScript -DuoDuoWeek $DuoDuoWeek -XiaoMingWeek $XiaoMingWeek -DDLetter $D.letter -DDWords $D.words -DDAct $D.act -XMLetter $X.letter -XMWords $X.words -XMAct $X.act -ChineseWeek $ChineseWeek -ChinesePinyin $C.pinyin -ChineseChars $C.chars -ChineseAct $C.act -ChinesePoem1Title $ChinesePoem1Title -ChinesePoem1Author $ChinesePoem1Author -ChinesePoem1Text $ChinesePoem1Text -ChinesePoem2Title $ChinesePoem2Title -ChinesePoem2Author $ChinesePoem2Author -ChinesePoem2Text $ChinesePoem2Text -WorksheetDir $WorksheetDir
    } catch {
      Write-Log ("DOCX生成失败: " + $_.Exception.Message)
    }
  } else {
    Write-Log "DOCX生成脚本不存在: $docxScript"
  }

# ====== 更新进度文件 ======
$Progress.duo_duo.current_week = $DuoDuoWeek
$Progress.xiao_ming.current_week = $XiaoMingWeek
$Progress.duo_duo_math.current_week = $MathWeek
$Progress.duo_duo_chinese.current_week = $ChineseWeek
$Progress | ConvertTo-Json -Depth 3 | Set-Content -Path $ProgressFile -Encoding utf8
Write-Log "进度文件已更新"

# ====== 生成FlowUs同步数据 ======
$SyncData = @{
  generated_at = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  duo_duo = @{
    page_id = "b80cd768-6ef5-4da6-a257-e2afe8d388ac"
    week = $DuoDuoWeek
    items = @(
      "按发音排序填空 (c_e, l_e, m_e, n_e, g_e)",
      "读句子填空 (race, cape, make a cake)",
      "数学：数数写数字 + 比较大小"
    )
  }
  xiao_ming = @{
    page_id = "8169ce9a-efd8-47bd-b671-2d0b8fa4c6e6"
    week = $XiaoMingWeek
    letter = $($X.letter)
    words = $($X.words)
    items = @(
      "字母描写: $($X.letter)",
      "找字母: 圈出 $($X.letter.Substring(0,1))",
      "单词认读: $($X.words)",
      "发音练习: apple → /a/"
    )
  }
  duo_duo_chinese = @{
    page_id = "74179223-c084-4be9-a08f-0dbebe16a55f"
    week = $ChineseWeek
    pinyin = $($C.pinyin)
    chars = $($C.chars)
    poem1_title = $ChinesePoem1Title
    poem1_author = $ChinesePoem1Author
    poem1_text = $ChinesePoem1Text
    poem2_title = $ChinesePoem2Title
    poem2_author = $ChinesePoem2Author
    poem2_text = $ChinesePoem2Text
    synced = $false
  }
  synced = $false
}
$SyncPath = Join-Path -Path $PSScriptRoot -ChildPath "worksheets\flowus_sync.json"
$SyncData | ConvertTo-Json -Depth 5 | Set-Content -Path $SyncPath -Encoding utf8
Write-Log "FlowUs同步数据已保存到: $SyncPath"

# ====== Windows 通知 ======
try {
  Add-Type -AssemblyName System.Windows.Forms

  $Notify = New-Object System.Windows.Forms.NotifyIcon
  $Notify.Icon = [System.Drawing.SystemIcons]::Information
  $Notify.BalloonTipIcon = "Info"
  $Notify.BalloonTipTitle = "学习时间到！"

  $b1 = " 多多 第${DuoDuoWeek}周：$($D.letter)`n   单词：$($D.words)`n   活动：$($D.act)"
  $b2 = " 小铭 第${XiaoMingWeek}周：$($X.letter)`n   单词：$($X.words)`n   活动：$($X.act)"
  $b3 = " 多多数学 第${MathWeek}周：$($M.ch)`n   内容：$($M.topic)`n   活动：$($M.act)"
  $b4 = " 多多语文 第${ChineseWeek}周：$($C.pinyin) + $($C.chars)`n   今日：$DailyAct"
  $b5 = " 练习卡片已保存至：$WorksheetDir`n 点击查看息流学习页面"
  $Body = "$b1`n`n$b2`n`n$b3`n`n$b4`n`n$b5"

  $Notify.BalloonTipText = $Body.Trim()
  $Notify.Visible = $true
  $Notify.ShowBalloonTip(20000)

  Register-ObjectEvent -InputObject $Notify -EventName BalloonTipClicked -Action {
    Start-Process "$DuoDuoUrl"
    Start-Process "$XiaoMingUrl"
    Start-Process "$ChineseUrl"
  } | Out-Null

  Start-Sleep -Seconds 22
  $Notify.Dispose()
  Write-Log "通知已发送"
} catch {
  Write-Log "通知发送失败：$($_.Exception.Message)"
}

# ====== 微信推送（Server酱） ======
$WeChatTitle = "学习提醒 - 第$([Math]::Floor($DaysSinceStart / 7) + 1)周"
$WeChatBody = "**多多 第${DuoDuoWeek}周：$($D.letter)**`n单词：$($D.words)`n活动：$($D.act)`n`n**小铭 第${XiaoMingWeek}周：$($X.letter)**`n单词：$($X.words)`n活动：$($X.act)`n`n**多多数学 第${MathWeek}周：$($M.ch)**`n内容：$($M.topic)`n活动：$($M.act)`n`n**多多语文 第${ChineseWeek}周：$($C.pinyin) + $($C.chars)**`n今日：$DailyAct`n`n查看息流详情：$ChineseUrl"
try {
  $jsonBody = @{ title = $WeChatTitle; desp = $WeChatBody } | ConvertTo-Json
  $utf8Body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
  $scUrl = "https://sctapi.ftqq.com/${ServerChanKey}.send"
  $resp = Invoke-RestMethod -Uri $scUrl -Method Post -Body $utf8Body -ContentType "application/json; charset=utf-8"
  if ($resp.code -eq 0) { Write-Log "微信推送成功" } else { Write-Log "微信推送返回：$($resp.message)" }
} catch {
  Write-Log "微信推送失败：$($_.Exception.Message)"
}

# ====== Gitee Issue 自动创建（仅周一） ======
try {
  $WeekNum = [Math]::Floor($DaysSinceStart / 7) + 1
  $IssueKey = "week_$WeekNum"
  if ($RunDate.DayOfWeek -eq 'Monday' -and -not $Progress.issues.$IssueKey) {
    $GiteeToken = [Environment]::GetEnvironmentVariable("GITEE_TOKEN")
    if ($GiteeToken) {
      $IssueTitle = "第${WeekNum}周学习计划（$($RunDate.ToString('yyyy-MM-dd'))起）"
      $p1 = "## 多多 · 牛津自然拼读 第3册 第${DuoDuoWeek}周`n- [ ] 字母/音：$($D.letter)`n- [ ] 单词：$($D.words)`n- [ ] 活动：$($D.act)"
      $p2 = "## 小铭 · 牛津自然拼读 第1册 第${XiaoMingWeek}周`n- [ ] 字母/音：$($X.letter)`n- [ ] 单词：$($X.words)`n- [ ] 活动：$($X.act)"
      $p3 = "## 多多 · 新加坡数学 G1 第${MathWeek}周`n- [ ] 章节：$($M.ch)`n- [ ] 内容：$($M.topic)`n- [ ] 活动：$($M.act)"
      $p4 = "## 多多 · 语文冲刺 第${ChineseWeek}周`n- [ ] 拼音：$($C.pinyin)`n- [ ] 汉字：$($C.chars)`n- [ ] 本周安排：$($C.act)`n- [ ] 古诗：$($C.poems[0].title)、$($C.poems[1].title)、$($C.poems[2].title)、$($C.poems[3].title)、$($C.poems[4].title)"
      $IssueBody = "$p1`n`n$p2`n`n$p3`n`n$p4"
      $bodyUtf8 = [System.Text.Encoding]::UTF8.GetBytes((@{ title = $IssueTitle; body = $IssueBody; labels = "学习计划" } | ConvertTo-Json))
      $resp = Invoke-RestMethod -Uri "https://gitee.com/api/v5/repos/yql8981229/learning-track/issues" -Method Post -Body $bodyUtf8 -ContentType "application/json; charset=utf-8" -Headers @{ Authorization = "Bearer $GiteeToken" }
      if ($resp.html_url) {
        $Progress.issues | Add-Member -NotePropertyName $IssueKey -NotePropertyValue $resp.html_url
        Write-Log "Gitee Issue 已创建：$($resp.html_url)"
      }
    }
  }
} catch {
  Write-Log "Gitee Issue 创建失败：$($_.Exception.Message)"
}

# ====== 保存进度 ======
$Progress | ConvertTo-Json -Depth 5 | Out-File -FilePath $ProgressFile -Encoding utf8

# ====== Gitee 自动同步 ======
try {
  $RepoDir = $PSScriptRoot
  Set-Location $RepoDir
  git add -A
  $CommitMsg = "daily sync - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
  git commit -m $CommitMsg
  git push
  Write-Log "Gitee 同步完成"
} catch {
  Write-Log "Gitee 同步失败：$($_.Exception.Message)"
}

Write-Log "===== 执行完成 ====="
