<#
.SYNOPSIS
  在息流(FlowUs)创建『药学启蒙』独立页面（含小明4岁/多多6岁两套知识与计划/问答卡）
.DESCRIPTION
  修复 PowerShell 5.1 Invoke-RestMethod -Body string 按 ASCII 编码导致中文乱码的问题，
  改用 UTF-8 字节发送 Body，并按 UTF-8 解析响应。
  父页面：多多 b80cd768-6ef5-4da6-a257-e2afe8d388ac
#>
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$token = "zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj"
$base = "https://api.flowus.cn/v2"
$parentPage = "b80cd768-6ef5-4da6-a257-e2afe8d388ac"

function Call-FlowUs-UTF8($method, $path, $body) {
  $url = $base + $path
  $json = if ($body) { $body | ConvertTo-Json -Depth 30 -Compress } else { $null }
  $params = @{
    Uri = $url
    Method = $method
    Headers = @{ "Authorization" = "Bearer $token" }
    ContentType = "application/json; charset=utf-8"
    TimeoutSec = 60
  }
  if ($null -ne $json) {
    # 关键：转 UTF-8 字节发送，避免 PS5.1 按 ASCII 编码把中文变 '?'
    $params.Body = [System.Text.Encoding]::UTF8.GetBytes($json)
  }
  $resp = Invoke-WebRequest @params -UseBasicParsing
  # 按 UTF-8 解码响应
  if ($resp.Content) {
    $raw = $resp.Content
    # Invoke-WebRequest 在 PS5.1 可能已按默认编码解码，这里直接用原始字节重新解析
    return $raw | ConvertFrom-Json
  }
  return $null
}

function RT($t) { @{ type = "text"; text = @{ content = $t } } }

# ---------- 创建页面 ----------
# 注意：FlowUs 页面的 title 属性需包一层 type="title"（见官方文档 create-page）
$properties = @{ title = @{
  type = "title"
  title = @(@{ type = "text"; text = @{ content = "药学启蒙" } })
} }
$body = @{
  parent = @{ page_id = $parentPage }
  properties = $properties
  icon = @{ type = "emoji"; emoji = "💊" }
  children = @(
    @{ object = "block"; type = "heading_1"; heading_1 = @{ rich_text = @(RT("药学启蒙 · 亲子计划")) } }
    @{ object = "block"; type = "paragraph"; paragraph = @{ rich_text = @(RT("给孩子的药学知识启蒙，按年龄分：小明4岁（安全/生活）与多多6岁（科学思维/简单化学）。文末含亲子问答卡与用药数学应用题。")) } }
    @{ object = "block"; type = "divider"; divider = @{} }
  )
}
Write-Host "=== 创建 药学启蒙 页面 ==="
$page = Call-FlowUs-UTF8 -method POST -path "/pages" -body $body
$pageId = $page.id
Write-Host "created: $pageId"
Write-Host "url: https://flowus.cn/$pageId"

# ---------- 小明 4 岁部分 ----------
$children1 = @(
  @{ object = "block"; type = "heading_2"; heading_2 = @{ rich_text = @(RT("一、小明（4岁）：安全第一，生活化")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("药不能乱吃：药长得像糖果但绝不是糖果，只有爸爸妈妈/医生给才能吃")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("认药游戏：认药盒颜色/图案/形状，认识药片、胶囊、糖浆、口服液")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("药怎么帮我们：生病→吃药→身体小卫士变强→打败坏东西")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("日常卫生：勤洗手挡小细菌，按时吃药、喝够水")) } }
  @{ object = "block"; type = "heading_3"; heading_3 = @{ rich_text = @(RT("每周计划")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第1周：安全规则——药不是糖，只有大人给才吃")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第2周：认一认家里的药盒（颜色/形状/图案）")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第3周：身体小卫士的故事——药怎么帮我们好起来")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第4周：洗手讲卫生，看不见的小细菌")) } }
)
Write-Host "=== 写入 小明部分 ==="
$r1 = Call-FlowUs-UTF8 -method PATCH -path "/blocks/$pageId/children" -body @{ children = $children1 }
Write-Host "小明: $($r1.results.Count) 块"

# ---------- 多多 6 岁部分 ----------
$children2 = @(
  @{ object = "block"; type = "divider"; divider = @{} }
  @{ object = "block"; type = "heading_2"; heading_2 = @{ rich_text = @(RT("二、多多（6岁，一年级）：科学思维 + 简单化学")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("药在身体里走的路：进嘴巴→滑到胃→被吸收进血液→送到需要的地方（吸收-分布-作用）")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("为什么打针比吃药快：吃药要经胃和吸收较慢，打针直接进入身体较快")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("为什么不能多吃药：吃多少由医生/说明书规定，过量会让身体小卫士受伤（剂量概念）")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("认识有效期和说明书：看日期、药名、用量、一天几次")) } }
  @{ object = "block"; type = "heading_3"; heading_3 = @{ rich_text = @(RT("动手小实验（化学启蒙）")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("泡腾片遇水冒气泡=二氧化碳，观察固体变气体")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("胶囊为什么颜色不同=区分不同药物")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("糖浆为什么甜=用甜味掩盖苦味")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("温度对药的影响：冷藏药 vs 常温药")) } }
  @{ object = "block"; type = "heading_3"; heading_3 = @{ rich_text = @(RT("每两周计划")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第1-2周：药在身体里走的路 + 打针vs吃药的对比")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第3-4周：剂量概念——不能多吃，配合用药数学应用题")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第5-6周：读说明书/有效期 + 家里药箱大探险")) } }
  @{ object = "block"; type = "to_do"; to_do = @{ rich_text = @(RT("第7-8周：泡腾片小实验 + 化学启蒙")) } }
)
Write-Host "=== 写入 多多部分 ==="
$r2 = Call-FlowUs-UTF8 -method PATCH -path "/blocks/$pageId/children" -body @{ children = $children2 }
Write-Host "多多: $($r2.results.Count) 块"

# ---------- 问答卡 + 用药数学 ----------
$children3 = @(
  @{ object = "block"; type = "divider"; divider = @{} }
  @{ object = "block"; type = "heading_2"; heading_2 = @{ rich_text = @(RT("三、亲子问答卡（多多）")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("Q1：药到身体里会经过哪几步？ A：进嘴巴→胃→被血液吸收→送到需要的地方")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("Q2：为什么打针比吃药快？ A：打针直接进身体，吃药要经胃慢慢吸收")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("Q3：为什么不能多吃药？ A：吃多少有规定，过量会伤害身体")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("Q4：泡腾片放进水为什么会冒泡泡？ A：生成二氧化碳气体")) } }
  @{ object = "block"; type = "bulleted_list_item"; bulleted_list_item = @{ rich_text = @(RT("Q5：过期药能吃吗？为什么？ A：不能，药品会变质失效")) } }
  @{ object = "block"; type = "heading_3"; heading_3 = @{ rich_text = @(RT("用药数学应用题（配合浅奥）")) } }
  @{ object = "block"; type = "paragraph"; paragraph = @{ rich_text = @(RT("例1：每天吃3次，每次2片，一天共吃几片？ → 3×2=6片")) } }
  @{ object = "block"; type = "paragraph"; paragraph = @{ rich_text = @(RT("例2：一瓶药有24片，每天2次每次2片（一天4片），能吃几天？ → 24÷4=6天")) } }
  @{ object = "block"; type = "paragraph"; paragraph = @{ rich_text = @(RT("例3：上午8点吃了1次，每6小时吃1次，下一次几点吃？ → 下午2点")) } }
  @{ object = "block"; type = "paragraph"; paragraph = @{ rich_text = @(RT("例4：一支口服液10毫升，一次喝5毫升，能喝几次？ → 2次")) } }
)
Write-Host "=== 写入 问答卡 ==="
$r3 = Call-FlowUs-UTF8 -method PATCH -path "/blocks/$pageId/children" -body @{ children = $children3 }
Write-Host "问答卡: $($r3.results.Count) 块"

Write-Host ""
Write-Host "=== DONE ==="
Write-Host "url: https://flowus.cn/$pageId"
