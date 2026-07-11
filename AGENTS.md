
# 🧑 用户信息
- **称呼**: 磊哥
- **GitHub**: 724450199-arch
- **Gitee**: yql8981229 (TOTORO)
- **孩子**: 多多（6岁，2026.9入小学）、小铭（4岁）
- **项目**: learning-track（985成才计划）
- **系统**: Windows, 用户名 yang

# ⚙️ 关键决策
- 代码托管以 **Gitee** 为主（GitHub 网络不可达）
- Gitee 仓库: `yql8981229/learning-track`，Token 在 `GITEE_TOKEN` 环境变量
- FlowUs 息流用于学习页面管理
- FlowUs Token: `zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj`（API: `https://api.flowus.cn/v2`）
- FlowUs 页面: 多多 `b80cd768-6ef5-4da6-a257-e2afe8d388ac`，小铭 `8169ce9a-efd8-47bd-b671-2d0b8fa4c6e6`
- AI 服务：通义万相（阿里云百炼）+ Agnes AI (Sapiens AI)
- 微信通知通过 Server酱 推送

# 🛡️ 桌面防删除看门狗
- 仓库位置: `E:\learning-track\`（桌面有快捷方式）
- `protect_desktop.ps1`: 后台运行（Startup 启动），每秒轮询桌面文件
- 发现文件被删除，立即从 `.desktop_backup` 隐藏目录恢复
- 启动方式：`C:\Users\yang\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\protect_desktop.bat`
- `desktop_utils.ps1`: 工具函数集（dot-source 后使用）
  - `New-DesktopShortcut "路径"` — 给文件夹/文件在桌面创建快捷方式
- `send_to_desktop.ps1`: SendTo 集成，右击→发送到→发送到桌面快捷方式
  - SendTo 位置: `%APPDATA%\Microsoft\Windows\SendTo\发送到桌面快捷方式.bat`
  - 需重启 Explorer 或重启电脑后生效

# 🚫 360 安全卫士（重要）
- 安装路径: `C:\Program Files (x86)\360\360Safe\`（自保护，不可写）
- 核心进程: `360tray.exe`（PID 13408，自保护不可杀）
- 清理配置: `C:\Program Files (x86)\360\360Safe\sweeper\cleancfg.xml`
- 桌面清理配置: `C:\Program Files (x86)\360\360Safe\sweeper\DesktopGC.xml`
- 用户配置: `%APPDATA%\360Safe\Clean\CleanSoft\okcleannew_config.dat`（可写但加密）
- 禁用文件: `%PROGRAMDATA%\360Safe\360Disabled`（可写但不可读）
- 注册表: `HKLM\SOFTWARE\WOW6432Node\360Safe\`（含 clean/CleanSoft/safemon/360krnlsvc 等子项）
- 自动清理: `smartcleanday=365`, `smartcleaninterval=7`（配置被锁定无法修改）
- 已被移除启动项: `360huabao`

# 📦 夸克
- **夸克网盘** v6.9.7.761: 已安装于 `%LOCALAPPDATA%\Programs\QuarkCloudDrive\quark_cloud_drive.exe`
- **夸克浏览器**: 腾讯软件中心 API 链接已获取 v6.9.6.896（331MB，下载 11.7MB 后中断需重试）

# 🌐 网络连接
- **Gitee**: ✅ 正常，remote URL 内嵌 Token，push/pull 正常
- **GitHub**: ⚠️ 443 端口不通（需 Clash Verge VPN 启动后才可访问）
- **Clash Verge**: 程序 `E:\VPN\安装文件\clash-verge.exe`，配置 `%APPDATA%\io.github.clash-verge-rev.clash-verge-rev\`

# ⚙️ 服务连接
- **FlowUs**: Token `zQC5s3N1auj4cFKXCAxZ8WLLJi1g7Wy80QZBuTfj`，硬编码在脚本和 AGENTS.md
- **FlowUs API**: `https://api.flowus.cn/v2`
- **FlowUs 页面**: 多多 `b80cd768-6ef5-4da6-a257-e2afe8d388ac`，小铭 `8169ce9a-efd8-47bd-b671-2d0b8fa4c6e6`
- **Agnes AI**: `AGNES_API_KEY` 环境变量已设 + `~/.agnes/api_key` 文件存在
- **Agnes AI API**: `https://apihub.agnes-ai.com/v1`
- **通义万相**: DashScope Key 在 `DASHSCOPE_API_KEY`（阿里云百炼）

# ⚠️ PowerShell 注意事项
- **字符串插值**: 在双引号字符串中访问哈希表/对象的属性，必须用 $(.prop) 子表达式，不能用 ${obj.prop}（后者会查找字面名为 obj.prop 的变量）。例如 ${p.title} 是错的，$(.title) 才对。