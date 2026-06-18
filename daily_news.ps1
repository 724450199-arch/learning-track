param(
  [string]$SendKey
)

# ====== 配置 ======
$LogFile = "$PSScriptRoot\news.log"
$MaxItems = 30
$RequestTimeout = 15

function Write-Log {
  param($Msg)
  $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  "$time $Msg" | Out-File -FilePath $LogFile -Encoding utf8 -Append
}

Write-Log "===== 新闻采集开始 ====="

# ====== 古诗词（65首，13周x5首，同 daily_learning.ps1） ======
$ChineseContent = @{}
$ChineseContent[1]  = @{poems=@(
  @{title="咏鹅"; author="骆宾王"; text="鹅鹅鹅，曲项向天歌。白毛浮绿水，红掌拨清波。"}
  @{title="静夜思"; author="李白"; text="床前明月光，疑是地上霜。举头望明月，低头思故乡。"}
  @{title="画"; author="王维"; text="远看山有色，近听水无声。春去花还在，人来鸟不惊。"}
  @{title="悯农（其一）"; author="李绅"; text="锄禾日当午，汗滴禾下土。谁知盘中餐，粒粒皆辛苦。"}
  @{title="悯农（其二）"; author="李绅"; text="春种一粒粟，秋收万颗子。四海无闲田，农夫犹饿死。"}
)}
$ChineseContent[2]  = @{poems=@(
  @{title="春晓"; author="孟浩然"; text="春眠不觉晓，处处闻啼鸟。夜来风雨声，花落知多少。"}
  @{title="古朗月行（节选）"; author="李白"; text="小时不识月，呼作白玉盘。又疑瑶台镜，飞在青云端。"}
  @{title="池上"; author="白居易"; text="小娃撑小艇，偷采白莲回。不解藏踪迹，浮萍一道开。"}
  @{title="江南"; author="汉乐府"; text="江南可采莲，莲叶何田田。鱼戏莲叶间。鱼戏莲叶东，鱼戏莲叶西，鱼戏莲叶南，鱼戏莲叶北。"}
  @{title="风"; author="李峤"; text="解落三秋叶，能开二月花。过江千尺浪，入竹万竿斜。"}
)}
$ChineseContent[3]  = @{poems=@(
  @{title="小池"; author="杨万里"; text="泉眼无声惜细流，树阴照水爱晴柔。小荷才露尖尖角，早有蜻蜓立上头。"}
  @{title="寻隐者不遇"; author="贾岛"; text="松下问童子，言师采药去。只在此山中，云深不知处。"}
  @{title="所见"; author="袁枚"; text="牧童骑黄牛，歌声振林樾。意欲捕鸣蝉，忽然闭口立。"}
  @{title="村居"; author="高鼎"; text="草长莺飞二月天，拂堤杨柳醉春烟。儿童散学归来早，忙趁东风放纸鸢。"}
  @{title="登鹳雀楼"; author="王之涣"; text="白日依山尽，黄河入海流。欲穷千里目，更上一层楼。"}
)}
$ChineseContent[4]  = @{poems=@(
  @{title="望庐山瀑布"; author="李白"; text="日照香炉生紫烟，遥看瀑布挂前川。飞流直下三千尺，疑是银河落九天。"}
  @{title="敕勒歌"; author="北朝民歌"; text="敕勒川，阴山下。天似穹庐，笼盖四野。天苍苍，野茫茫，风吹草低见牛羊。"}
  @{title="咏柳"; author="贺知章"; text="碧玉妆成一树高，万条垂下绿丝绦。不知细叶谁裁出，二月春风似剪刀。"}
  @{title="赠汪伦"; author="李白"; text="李白乘舟将欲行，忽闻岸上踏歌声。桃花潭水深千尺，不及汪伦送我情。"}
  @{title="绝句（迟日江山）"; author="杜甫"; text="迟日江山丽，春风花草香。泥融飞燕子，沙暖睡鸳鸯。"}
)}
$ChineseContent[5]  = @{poems=@(
  @{title="绝句（两个黄鹂）"; author="杜甫"; text="两个黄鹂鸣翠柳，一行白鹭上青天。窗含西岭千秋雪，门泊东吴万里船。"}
  @{title="望天门山"; author="李白"; text="天门中断楚江开，碧水东流至此回。两岸青山相对出，孤帆一片日边来。"}
  @{title="江雪"; author="柳宗元"; text="千山鸟飞绝，万径人踪灭。孤舟蓑笠翁，独钓寒江雪。"}
  @{title="山行"; author="杜牧"; text="远上寒山石径斜，白云生处有人家。停车坐爱枫林晚，霜叶红于二月花。"}
  @{title="回乡偶书"; author="贺知章"; text="少小离家老大回，乡音无改鬓毛衰。儿童相见不相识，笑问客从何处来。"}
)}
$ChineseContent[6]  = @{poems=@(
  @{title="早发白帝城"; author="李白"; text="朝辞白帝彩云间，千里江陵一日还。两岸猿声啼不住，轻舟已过万重山。"}
  @{title="饮湖上初晴后雨"; author="苏轼"; text="水光潋滟晴方好，山色空蒙雨亦奇。欲把西湖比西子，淡妆浓抹总相宜。"}
  @{title="惠崇春江晚景"; author="苏轼"; text="竹外桃花三两枝，春江水暖鸭先知。蒌蒿满地芦芽短，正是河豚欲上时。"}
  @{title="元日"; author="王安石"; text="爆竹声中一岁除，春风送暖入屠苏。千门万户曈曈日，总把新桃换旧符。"}
  @{title="清明"; author="杜牧"; text="清明时节雨纷纷，路上行人欲断魂。借问酒家何处有，牧童遥指杏花村。"}
)}
$ChineseContent[7]  = @{poems=@(
  @{title="九月九日忆山东兄弟"; author="王维"; text="独在异乡为异客，每逢佳节倍思亲。遥知兄弟登高处，遍插茱萸少一人。"}
  @{title="凉州词"; author="王之涣"; text="黄河远上白云间，一片孤城万仞山。羌笛何须怨杨柳，春风不度玉门关。"}
  @{title="枫桥夜泊"; author="张继"; text="月落乌啼霜满天，江枫渔火对愁眠。姑苏城外寒山寺，夜半钟声到客船。"}
  @{title="凉州词"; author="王翰"; text="葡萄美酒夜光杯，欲饮琵琶马上催。醉卧沙场君莫笑，古来征战几人回。"}
  @{title="出塞"; author="王昌龄"; text="秦时明月汉时关，万里长征人未还。但使龙城飞将在，不教胡马度阴山。"}
)}
$ChineseContent[8]  = @{poems=@(
  @{title="鹿柴"; author="王维"; text="空山不见人，但闻人语响。返景入深林，复照青苔上。"}
  @{title="别董大"; author="高适"; text="千里黄云白日曛，北风吹雁雪纷纷。莫愁前路无知己，天下谁人不识君。"}
  @{title="题西林壁"; author="苏轼"; text="横看成岭侧成峰，远近高低各不同。不识庐山真面目，只缘身在此山中。"}
  @{title="夏日绝句"; author="李清照"; text="生当作人杰，死亦为鬼雄。至今思项羽，不肯过江东。"}
  @{title="墨梅"; author="王冕"; text="吾家洗砚池头树，朵朵花开淡墨痕。不要人夸好颜色，只留清气满乾坤。"}
)}
$ChineseContent[9]  = @{poems=@(
  @{title="送元二使安西"; author="王维"; text="渭城朝雨浥轻尘，客舍青青柳色新。劝君更尽一杯酒，西出阳关无故人。"}
  @{title="黄鹤楼送孟浩然之广陵"; author="李白"; text="故人西辞黄鹤楼，烟花三月下扬州。孤帆远影碧空尽，唯见长江天际流。"}
  @{title="独坐敬亭山"; author="李白"; text="众鸟高飞尽，孤云独去闲。相看两不厌，只有敬亭山。"}
  @{title="泊船瓜洲"; author="王安石"; text="京口瓜洲一水间，钟山只隔数重山。春风又绿江南岸，明月何时照我还。"}
  @{title="示儿"; author="陆游"; text="死去元知万事空，但悲不见九州同。王师北定中原日，家祭无忘告乃翁。"}
)}
$ChineseContent[10] = @{poems=@(
  @{title="题临安邸"; author="林升"; text="山外青山楼外楼，西湖歌舞几时休。暖风熏得游人醉，直把杭州作汴州。"}
  @{title="秋夜将晓出篱门迎凉有感"; author="陆游"; text="三万里河东入海，五千仞岳上摩天。遗民泪尽胡尘里，南望王师又一年。"}
  @{title="四时田园杂兴（其三十一）"; author="范成大"; text="昼出耘田夜绩麻，村庄儿女各当家。童孙未解供耕织，也傍桑阴学种瓜。"}
  @{title="晓出净慈寺送林子方"; author="杨万里"; text="毕竟西湖六月中，风光不与四时同。接天莲叶无穷碧，映日荷花别样红。"}
  @{title="蜂"; author="罗隐"; text="不论平地与山尖，无限风光尽被占。采得百花成蜜后，为谁辛苦为谁甜。"}
)}
$ChineseContent[11] = @{poems=@(
  @{title="江上渔者"; author="范仲淹"; text="江上往来人，但爱鲈鱼美。君看一叶舟，出没风波里。"}
  @{title="游园不值"; author="叶绍翁"; text="应怜屐齿印苍苔，小扣柴扉久不开。春色满园关不住，一枝红杏出墙来。"}
  @{title="江南春"; author="杜牧"; text="千里莺啼绿映红，水村山郭酒旗风。南朝四百八十寺，多少楼台烟雨中。"}
  @{title="观书有感"; author="朱熹"; text="半亩方塘一鉴开，天光云影共徘徊。问渠那得清如许，为有源头活水来。"}
  @{title="石灰吟"; author="于谦"; text="千锤万凿出深山，烈火焚烧若等闲。粉骨碎身全不怕，要留清白在人间。"}
)}
$ChineseContent[12] = @{poems=@(
  @{title="春夜喜雨"; author="杜甫"; text="好雨知时节，当春乃发生。随风潜入夜，润物细无声。"}
  @{title="竹石"; author="郑燮"; text="咬定青山不放松，立根原在破岩中。千磨万击还坚劲，任尔东西南北风。"}
  @{title="己亥杂诗"; author="龚自珍"; text="浩荡离愁白日斜，吟鞭东指即天涯。落红不是无情物，化作春泥更护花。"}
  @{title="七步诗"; author="曹植"; text="煮豆燃豆萁，豆在釜中泣。本是同根生，相煎何太急。"}
  @{title="长歌行（节选）"; author="汉乐府"; text="青青园中葵，朝露待日晞。阳春布德泽，万物生光辉。常恐秋节至，焜黄华叶衰。百川东到海，何时复西归。少壮不努力，老大徒伤悲。"}
)}
$ChineseContent[13] = @{poems=@(
  @{title="书湖阴先生壁"; author="王安石"; text="茅檐长扫净无苔，花木成畦手自栽。一水护田将绿绕，两山排闼送青来。"}
  @{title="六月二十七日望湖楼醉书"; author="苏轼"; text="黑云翻墨未遮山，白雨跳珠乱入船。卷地风来忽吹散，望湖楼下水如天。"}
  @{title="芙蓉楼送辛渐"; author="王昌龄"; text="寒雨连江夜入吴，平明送客楚山孤。洛阳亲友如相问，一片冰心在玉壶。"}
  @{title="送杜少府之任蜀州"; author="王勃"; text="城阙辅三秦，风烟望五津。与君离别意，同是宦游人。海内存知己，天涯若比邻。无为在歧路，儿女共沾巾。"}
  @{title="望月怀远"; author="张九龄"; text="海上生明月，天涯共此时。情人怨遥夜，竟夕起相思。"}
)}

# ====== 选当日古诗 ======
$StartDate = [DateTime]"2026-05-31"
$DaysSinceStart = [Math]::Floor(((Get-Date) - $StartDate).TotalDays)
$RawWeek = [Math]::Floor($DaysSinceStart / 7) + 1
$ChineseWeek = [Math]::Min([Math]::Max(1, $RawWeek), 13)
$Dow = [int](Get-Date).DayOfWeek
$DayIndex = if ($Dow -eq 0 -or $Dow -eq 6) { 5 } else { $Dow }
$cnPoems = $ChineseContent[$ChineseWeek].poems
$PoemTitle = ""; $PoemAuthor = ""; $PoemText = ""
if ($cnPoems -and $cnPoems.Count -ge 1) {
  $idx = (($DayIndex - 1) * 2) % $cnPoems.Count
  $PoemTitle = $cnPoems[$idx].title
  $PoemAuthor = $cnPoems[$idx].author
  $PoemText = $cnPoems[$idx].text
}
$WeekCN = @{1="一";2="二";3="三";4="四";5="五";6="六";7="七";8="八";9="九";10="十";11="十一";12="十二";13="十三"}[$ChineseWeek]

# ====== RSS 来源 ======
$Sources = @(
  @{ Name = "36氪"; Url = "https://36kr.com/feed" }
  @{ Name = "IT之家"; Url = "https://www.ithome.com/rss/" }
  @{ Name = "Reuters Tech"; Url = "https://feeds.reuters.com/reuters/technologyNews" }
)

# ====== 关键词 ======
$Keywords = @(
  "AI","人工智能","大模型","LLM","AGI","OpenAI","ChatGPT","GPT","Claude","Gemini",
  "芯片","半导体","封装","Chiplet","先进封装","CoWoS","HBM","TSMC","台积电","ASML",
  "GPU","CPU","算力","数据中心","HPC","英伟达","NVIDIA","AMD","Intel",
  "光刻机","EDA","RISC-V","Arm",
  "科技","互联网","量子计算","5G","6G","机器人","自动驾驶"
)

# ====== 采集 ======
$AllItems = @()
foreach ($Src in $Sources) {
  try {
    Write-Log "正在采集: $($Src.Name)"
    $resp = Invoke-WebRequest -Uri $Src.Url -TimeoutSec $RequestTimeout -UseBasicParsing -ErrorAction Stop
    $xml = [System.Xml.XmlDocument]::new()
    $xml.LoadXml($resp.Content)
    $items = $xml.rss.channel.item | Select-Object -First 25
    foreach ($item in $items) {
      $title = if ($item.title.'#cdata-section') { $item.title.'#cdata-section' } else { "$($item.title)" }
      $desc  = if ($item.description.'#cdata-section') { $item.description.'#cdata-section' } else { "$($item.description)" }
      $link  = if ($item.link.'#cdata-section') { $item.link.'#cdata-section' } else { "$($item.link)" }
      $AllItems += @{
        title   = $title -replace '\s+',' '
        summary = ($desc -replace '<[^>]+>','').Trim() -replace '\s+',' '
        link    = $link.Trim()
        source  = $Src.Name
      }
    }
  } catch {
    Write-Log "采集 $($Src.Name) 失败: $_"
  }
}

Write-Log "共采集 $($AllItems.Count) 条新闻"

# ====== 关键词过滤去重 ======
$Filtered = @()
$KeywordsPattern = "($($Keywords -join '|'))"
$Seen = @{}

foreach ($item in $AllItems) {
  $text = "$($item.title) $($item.summary)"
  if ($text -match $KeywordsPattern) {
    $key = $item.title.Substring(0, [Math]::Min(50, $item.title.Length))
    if (-not $Seen.ContainsKey($key)) {
      $Seen[$key] = $true
      $MatchesList = [regex]::Matches($text, $KeywordsPattern) | ForEach-Object { $_.Value }
      $item.matchCount = ($MatchesList | Select-Object -Unique).Count
      $Filtered += $item
    }
  }
}
$Filtered = $Filtered | Sort-Object matchCount -Descending | Select-Object -First $MaxItems
Write-Log "过滤后 $($Filtered.Count) 条相关新闻"

# ====== 格式化 ======
$DateStr = Get-Date -Format "yyyy-MM-dd"
$TimeStr = Get-Date -Format "HH:mm"
$Body = @"
## 每日科技早报
**$DateStr** | 共 $($Filtered.Count) 条精选

"@

# ====== 添加今日古诗 ======
if ($PoemText) {
  $Body += "---`n📜 **今日古诗（语文第${WeekCN}周）**`n**${PoemTitle} - ${PoemAuthor}**`n${PoemText}`n---`n`n"
}

$i = 1
foreach ($item in $Filtered) {
  $summary = if ($item.summary) { $item.summary.Substring(0, [Math]::Min(120, $item.summary.Length)) } else { "" }
  $summary = $summary -replace '[\r\n]+',' '
  $Body += "**$i. $($item.title)**
> [$($item.source)] $summary
$($item.link)

"
  $i++
}

$Body += "---
自动采集于 $DateStr $TimeStr | 来源: 36氪 / IT之家 / Reuters"

# ====== 方糖推送 ======
if (-not $SendKey) {
  $ConfigFile = "$PSScriptRoot\sendkey.txt"
  if (Test-Path $ConfigFile) {
    $SendKey = (Get-Content $ConfigFile -Raw).Trim()
  }
}

if ($SendKey) {
  try {
    $resp = Invoke-RestMethod -Uri "https://sctapi.ftqq.com/$SendKey.send" -Method Post `
      -Body @{ title = "每日科技早报 $DateStr"; content = $Body } `
      -TimeoutSec 30
    if ($resp.code -eq 0) {
      Write-Log "方糖推送成功 (pushid: $($resp.data.pushid))"
    } else {
      Write-Log "方糖推送失败: code=$($resp.code) msg=$($resp.message)"
    }
  } catch {
    Write-Log "方糖推送异常: $_"
  }
} else {
  $out = "$PSScriptRoot\daily_news_$DateStr.md"
  $Body | Out-File -FilePath $out -Encoding utf8
  Write-Log "SendKey 未配置，日报已保存到: $out"
  Write-Host "已保存到 $out"
}

Write-Log "===== 采集结束 ====="
Write-Host "完成，共 $($Filtered.Count) 条"
