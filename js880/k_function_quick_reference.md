# k 函数快速参考卡

## 🚀 立即可用的公式

### 1️⃣ 数组过滤
```
=k("data=>data.filter((x,i)=>i>0)",A1:B100)
跳过标题行，返回剩余行
```

### 2️⃣ 数组转换
```
=k("data=>data.map(x=>[x[0]*2, x[1]])",A1:B100)
第一列翻倍，第二列不变
```

### 3️⃣ 链式过滤+转换
```
=k("data=>data.filter((x,i)=>i>0 && x[1]>1000).map(x=>[x[0],x[1]*1.1])",A1:C100)
过滤出大于1000的，增加10%
```

### 4️⃣ 条件筛选（标题+特定条件）
```
=k("data=>data.filter((x,i)=>i==0 || x[1]=='特定值')",A1:D100)
保留标题行 + 满足条件的行
```

### 5️⃣ 透视表筛选 ⭐
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",
   A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
保留标题 + Product1 数据
```

### 6️⃣ 数值汇总
```
=k("data=>data.reduce((sum,x,i)=>i>0?sum+x[2]:sum,0)",A1:C100)
计算第3列的总和
```

### 7️⃣ 排序后取前N个
```
=k("data=>data.filter((x,i)=>i>0).sort((a,b)=>b[2]-a[2]).slice(0,6)",A1:D1000)
按第3列降序排列，取前5条（+标题）
```

---

## 📚 方法速查表

| 方法 | 语法 | 返回 |
|------|------|------|
| filter | `data.filter((x,i)=>条件)` | 过滤后的数组 |
| map | `data.map(x=>新值)` | 转换后的数组 |
| slice | `data.slice(起,止)` | 切片数组 |
| sort | `data.sort((a,b)=>比较)` | 排序数组 |
| reduce | `data.reduce((累积,x)=>结果,初值)` | 单个聚合值 |
| find | `data.find(x=>条件)` | 第一个匹配 |
| some | `data.some(x=>条件)` | 是否存在 |
| every | `data.every(x=>条件)` | 是否全部 |
| reverse | `data.reverse()` | 倒序数组 |

---

## 🎯 常见场景

### 数据验证：只保留有效记录
```
=k("data=>data.filter((x,i)=>i==0||x[1]>0&&x[2]!='')",A1:D100)
```

### 数据清理：移除空值并转换
```
=k("data=>data.filter(x=>x[0]&&x[1]).map(x=>[x[0].trim(),parseInt(x[1])])",A1:B100)
```

### 分组统计：汇总数据
```
=k("data=>data.filter((x,i)=>i>0).reduce((acc,x)=>{var idx=acc.findIndex(e=>e[0]==x[0]);if(idx>-1)acc[idx][1]+=x[1];else acc.push([x[0],x[1]]);return acc},[])",A1:B100)
按第1列分组求和
```

### 合并字段：创建新列
```
=k("data=>data.map((x,i)=>i==0?[x[0],x[1],'合并']:['合并','新增'])",A1:B100)
```

### 排序+限制：Top N
```
=k("data=>[...data.slice(0,1),...data.slice(1).sort((a,b)=>b[2]-a[2]).slice(0,9)]",A1:D100)
保留标题+排名前9个
```

---

## ⚡ 性能贴士

✅ **好的做法**
```javascript
.filter(...).map(...)          // ✓ 先过滤再转换
.filter(...).sort(...).slice()  // ✓ 按此顺序优化
[...数据].reduce(...)          // ✓ 使用展开符确保可迭代
```

❌ **避免**
```javascript
.map(...).filter(...)          // ✗ 先转换再过滤（浪费）
.sort(...).sort(...)           // ✗ 多次排序
.filter(...).filter(...)       // ✗ 多次过滤同样条件
```

---

## 🔍 调试技巧

### 测试表达式
在 WPS 开发者工具的控制台中：
```javascript
// 测试 filter
JSA.jsaLambda("data=>data.filter((x,i)=>i>0)", [[1,'a'],[2,'b']])

// 测试 map
JSA.jsaLambda("data=>data.map(x=>[x[0]*2])", [[1],[2],[3]])

// 测试链式
JSA.jsaLambda("data=>data.filter((x,i)=>i>0).map(x=>[x[0]*2])", [[1,'a'],[2,'b']])
```

### 打印调试
```javascript
// 在 map 中插入日志
.map(x=>{console.log('当前行:',x); return x;})

// 验证过滤条件
.filter((x,i)=>{var result = i==0 || x[1]>100; console.log(x, result); return result;})
```

---

## 🛡️ 错误处理

| 错误 | 原因 | 修复 |
|------|------|------|
| `#K_ERR` | 表达式语法错误 | 检查括号、箭头函数语法 |
| `#VALUE!` | 数据类型不匹配 | 检查数组维度（一维vs二维） |
| `#NAME?` | $$ 未定义 | 确保 WPS 版本 ≥ 15990 |
| 结果为空 | 过滤条件过严 | 用 `console.log` 验证条件 |
| 溢出未生效 | 单元格右侧有内容 | 清空下方单元格或向右移动 |

---

## 📱 手机快速参考

**最常用的 5 个**:
1. `data.filter((x,i)=>i>0)` - 跳过标题
2. `data.map(x=>[...])` - 转换数据
3. `.sort((a,b)=>b[0]-a[0])` - 降序
4. `.slice(0,10)` - 取前10个
5. `.reduce((s,x)=>s+x[0],0)` - 求和

**关键变量**:
- `x` - 当前行（数组）
- `i` - 当前行号（从0开始）
- `a,b` - 排序时的比较对象

**快速判断**:
- 返回布尔值 → 用在 `filter` 中
- 返回新值 → 用在 `map` 中  
- 返回单值 → 用在 `reduce` 中

---

## 💾 保存此卡为快捷参考

推荐操作：
1. 将此内容保存到文本编辑器
2. 在 WPS 中新建一个"参考表"工作表
3. 粘贴常用公式到该表中
4. 需要时快速复制使用

---

**最后更新**: 2026-06-05  
**版本**: 1.0  
**需要帮助**: 查看 README_k_enhancement.md 或 k_function_implementation_guide.md
