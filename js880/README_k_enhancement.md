# 增强 k 函数 - 完整设计方案

## 📋 概览

您需要的 **增强 k 函数** 已完整设计完成。该方案允许在 WPS 单元格中使用链式方法调用（如 `.filter()`, `.map()` 等），并返回数组结果自动溢出到单元格。

## 🎯 核心能力

### 原有支持
```javascript
=k("JSA.getIndexs", 1, 10, 2)           // 路径调用
=k("x => x*2", 5)                       // Lambda 表达式
=k("$0+$1", [10, 20])                   // 索引选择器
=k("Array2D.z超级透视", A1:H40, ...)    // 方法调用
```

### 新增支持 ✨
```javascript
// 链式调用 - 单一数组
=k("data=>data.filter((x,i)=>i>0).map(x=>[x[0]*2, x[1]])", A1:B10)

// 链式调用 - 透视表
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>x.f2=='Product1')",
   A1:H40,"f3,f2","","count(),sum(`f4`)")

// 复杂筛选组合
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1' && x.f3>100)",
   A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```

## 📦 交付物

本方案包含以下文件：

| 文件 | 用途 |
|------|------|
| `k_function_design.md` | 完整设计文档 |
| `k_function_implementation_guide.md` | 实现步骤指南 |
| `enhanced_k_function.js` | 实现代码（直接复制） |
| `test_enhanced_k_function.js` | 测试套件 |
| `README_k_enhancement.md` | 此文件 |

## 🚀 快速开始

### 方案 A: 完整实现（推荐）
1. 复制 `enhanced_k_function.js` 中的代码
2. 按 `k_function_implementation_guide.md` 的指步骤集成到 JSA880.js
3. 在 WPS 中运行 `test_enhanced_k_function.js` 验证

### 方案 B: 最小化实现
如果只需要基本链式调用，只需修改两个地方：

**位置 1** (约 440 行): 添加全局初始化
```javascript
(function initGlobalAliases() {
    if (typeof globalThis !== 'undefined') {
        globalThis.$$ = undefined;
    }
})();
```

**位置 2** (约 18820 行): 修改 k 函数
```javascript
function k(fn, ...args) {
    try {
        if (typeof globalThis !== 'undefined' && Array2D) {
            globalThis.$$ = Array2D;
        }
        return JSA.jsaLambda(fn, ...args);
    } catch (e) {
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}
```

## 💡 使用示例

### 例子 1: 基础数组过滤
**需求**: 从 A1:B100 中提取销售额 > 1000 的记录，销售额翻倍

**公式**:
```
=k("data=>data.filter((x,i)=>i==0||x[1]>1000).map(x=>[x[0],x[1]*2])", A1:B100)
```

**结果**: 标题行 + 满足条件的数据，第2列为原值的2倍

---

### 例子 2: 透视表条件筛选 ⭐
**需求**: 对销售数据按产品/地区进行透视，只显示 Product1 的数据

**源数据**: A1:H40
- f1: 日期
- f2: 产品 (Product1, Product2, ...)
- f3: 地区 (North, South, ...)
- f4: 销售额

**公式**:
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",
   A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```

**参数说明**:
- 行字段: f3 (地区)
- 列字段: f2 (产品)
- 数据: count(), sum(f4), textjoin(f4, "+")

**结果**: 
```
           Product1  Product2  Product3
North      2         -         1
South      3         2         -
East       1         1         1
```
（自动按 Product1 列筛选）

---

### 例子 3: 多条件复杂筛选
**需求**: 保留标题 + Product1 且销售额 > 1500 的记录，销售额增加 10%

**公式**:
```
=k("(...args)=>$$.superPivot(...args)
    .filter((x,i)=>i==0 || (x.f2=='Product1' && x.f4>1500))
    .map(x=>i==0?x:({...x, f4: x.f4*1.1}))", 
   A1:H40,"f3,f2","","count(),sum(`f4`)")
```

**流程**:
1. superPivot 生成透视表
2. 第一个 filter: 保留标题 + Product1 > 1500
3. map: 第2列起的销售额增加 10%

---

### 例子 4: 数据聚合求和
**需求**: 计算所有销售数据的总额（跳过标题）

**公式**:
```
=k("data=>data.reduce((sum,x,i)=>i>0?sum+(x[2]||0):sum, 0)", A1:C100)
```

**返回**: 单个数值（总销售额）

---

### 例子 5: 排序 + 限制行数
**需求**: 获取销售额前 5 的产品

**公式**:
```
=k("data=>data.filter((x,i)=>i>0)
    .sort((a,b)=>b[2]-a[2])
    .slice(0, 6)", A1:D1000)
```

**说明**: 
- filter: 跳过标题
- sort: 按销售额降序
- slice: 取前6行（标题+前5条数据）

## 🔧 工作原理

```
用户公式
    ↓
┌─────────────────────────────────────┐
│ k("expr", ...args)                  │
└─────────────────────────────────────┘
    ↓ UDF 调用
┌─────────────────────────────────────┐
│ 1. 初始化 $$ = Array2D              │
│ 2. 调用 JSA.jsaLambda(expr, args)   │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│ 3. 检测链式调用 (.filter / .map)    │
│    ✓ 是 → parseChainableExpression  │
│    ✗ 否 → z解析函数表达式           │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│ 4. 创建函数 Function('$$', expr)    │
│    确保 $$ 在作用域中                │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│ 5. 执行链式操作                      │
│    superPivot(...) → filter(...) →  │
│    map(...) → 返回结果数组           │
└─────────────────────────────────────┘
    ↓
┌─────────────────────────────────────┐
│ 6. WPS 15990+ 自动处理数组溢出      │
│    结果填充到下面的单元格            │
└─────────────────────────────────────┘
```

## 📊 支持的方法

### 标准数组方法
| 方法 | 描述 | 示例 |
|------|------|------|
| `.filter(fn)` | 过滤行 | `.filter((x,i)=>x[1]>100)` |
| `.map(fn)` | 转换行 | `.map(x=>[x[0], x[1]*2])` |
| `.reduce(fn, init)` | 汇总 | `.reduce((sum,x)=>sum+x[1],0)` |
| `.slice(start, end)` | 提取范围 | `.slice(1, 10)` |
| `.sort(fn)` | 排序 | `.sort((a,b)=>b[1]-a[1])` |
| `.reverse()` | 倒序 | `.reverse()` |
| `.find(fn)` | 查找第一个 | `.find(x=>x[0]=='target')` |
| `.some(fn)` | 是否存在 | `.some(x=>x[1]<0)` |
| `.every(fn)` | 是否全部 | `.every(x=>x[1]>0)` |

### JSA880 Array2D 方法
| 方法 | 描述 | 示例 |
|------|------|------|
| `$$.superPivot(...)` | 超级透视 | 见例子 2 |
| `$$.z筛选(fn)` | 筛选 | `.filter()` 的 JSA 版本 |
| `$$.z映射(fn)` | 映射 | `.map()` 的 JSA 版本 |
| `$$.z排序(fn)` | 排序 | `.sort()` 的 JSA 版本 |
| `$$.z计数()` | 计数 | 返回行数 |

## ⚠️ 注意事项

### 1. WPS 版本要求
- **最低版本**: WPS 15990+
- 数组溢出功能需要较新版本的 WPS
- 建议使用 2023 年及以后版本

### 2. 性能考虑
- 大数据量（>100万行）的链式操作可能很慢
- 建议先 filter 减少数据，再执行其他操作
- 表达式会被缓存，重复使用不会重新解析

### 3. 内存限制
- 透视结果可能超过 WPS 单元格容量
- 单个工作表最多约 104 万行（Excel 限制）
- 溢出结果默认填充到右侧和下方

### 4. 特殊字符处理
- Lambda 表达式中的反引号 `` ` `` 会自动转换为 `"` 
- 需要真正的模板字符串时，使用 `${...}` 语法

### 5. 错误处理
- 执行错误返回 `#K_ERR: 错误信息`
- 检查 WPS 开发者工具的控制台查看详细错误

## 🧪 测试清单

实现后请验证：

- [ ] **基础功能**
  - [ ] 原有公式（无链式）仍正常工作
  - [ ] `=k("x=>x*2", 5)` 返回 10
  - [ ] `=k("JSA.getIndexs", 1, 5, 1)` 返回 [1,2,3,4,5]

- [ ] **全局别名**
  - [ ] WPS 开发者工具中 `$$` 被定义
  - [ ] `$$ === Array2D` 返回 true

- [ ] **简单链式**
  - [ ] `=k("data=>data.filter((x,i)=>i>0)", A1:B5)` 跳过标题
  - [ ] `=k("data=>data.map(x=>[x[0]*2])", A1:A5)` 翻倍

- [ ] **复杂链式**
  - [ ] 链式 filter + map 正常执行
  - [ ] 结果自动溢出到右下方单元格

- [ ] **透视表集成**
  - [ ] `=k("(...args)=>$$.superPivot(...args)", ...)` 生成透视
  - [ ] `=k("(...args)=>$$.superPivot(...args).filter(...)", ...)` 筛选生效

- [ ] **性能**
  - [ ] 表达式缓存有效（重复调用不变慢）
  - [ ] 大数据集（>10000 行）可以处理

## 📝 文件清单

```
js880/
├── JSA880.js                          (需要修改)
├── k_function_design.md               (设计文档)
├── k_function_implementation_guide.md (实现指南)
├── enhanced_k_function.js             (实现代码)
├── test_enhanced_k_function.js        (测试代码)
└── README_k_enhancement.md            (此文件)
```

## 🎓 学习资源

### 理解 Lambda 表达式
```javascript
// 基础
x => x * 2                    // 一参数简写
(x) => x * 2                  // 标准形式
(x, y) => x + y              // 多参数
(x, y, z) => { return x+y+z; } // 代码块

// JSA 环境中
"x => x*2"                    // 字符串形式（在 k() 中使用）
(...args) => $$.method(...args)  // 可变参数 + 对象方法
```

### 理解数组操作
```javascript
// 原始数组
[[1,'a'], [2,'b'], [3,'c']]

// Filter: 保留符合条件的行
.filter((x,i)=>i>0)           // [[2,'b'],[3,'c']]
.filter((x,i)=>x[0]>1)        // [[2,'b'],[3,'c']]

// Map: 转换每一行
.map(x=>[x[0]*2, x[1]])       // [[4,'a'],[6,'b'],[8,'c']]
.map(x=>({id:x[0], name:x[1]}))  // 转对象格式

// 链式: 依次执行
.filter(...).map(...).sort(...)
```

## 📞 故障排查

| 问题 | 症状 | 解决方案 |
|------|------|---------|
| $$ 未定义 | 返回 `#NAME?` | 检查 finalizeGlobalAliases() 是否被调用 |
| 链式无效 | filter/map 不生效 | 确保表达式中有 `.filter(` 等关键词 |
| 性能下降 | 操作变得很慢 | 减少数据量或简化表达式 |
| 溢出不工作 | 结果未填充到右下 | 需要 WPS 15990+ |
| 错误信息不清 | `#K_ERR` 但看不到原因 | 打开 WPS 开发者工具查看日志 |

## 🔗 相关资源

- JSA880 框架: https://github.com/...（您的 repo）
- WPS 二次开发文档: WPS 官方文档
- JavaScript Lambda: MDN Web Docs - Arrow Functions

---

**版本**: 1.0  
**日期**: 2026-06-05  
**作者**: Claude Agent  
**状态**: ✅ 设计完成，待集成测试
