
- **PowerShell 字符串插值 ⚠️**: 在双引号字符串中访问哈希表/对象的属性，必须用 $(.prop) 子表达式，不能用 ${obj.prop}（后者会查找字面名为 obj.prop 的变量）。例如 ${p.title} 是错的，$(.title) 才对。