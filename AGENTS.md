
# 🧑 用户信息
- **称呼**: 磊哥
- **GitHub/Gitee**: 724450199-arch
- **孩子**: 多多（6岁，2026.9入小学）、小铭（4岁）
- **项目**: learning-track（985成才计划）
- **系统**: Windows, 用户名 yang

# ⚙️ 关键决策
- 代码托管以 **Gitee** 为主（GitHub 网络不可达）
- FlowUs 息流用于学习页面管理
- AI 服务：通义万相（阿里云百炼）+ Agnes AI (Sapiens AI)
- 微信通知通过 Server酱 推送

# ⚠️ PowerShell 注意事项
- **字符串插值**: 在双引号字符串中访问哈希表/对象的属性，必须用 $(.prop) 子表达式，不能用 ${obj.prop}（后者会查找字面名为 obj.prop 的变量）。例如 ${p.title} 是错的，$(.title) 才对。