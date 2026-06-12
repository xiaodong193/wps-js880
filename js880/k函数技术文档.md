# k() 函数技术文档与使用指南

> 适用范围：JSA880 v5.0.0+ / WPS JSA (macOS & Windows)

---

## 1. 概述

`k()` 是 JSA880 框架提供的 **WPS 单元格公式 UDF（用户自定义函数）**，核心能力是把字符串表达式编译为可执行函数，在 WPS 公式中直接调用 JSA880/Array2D 的全部能力——数组生成、透视、筛选、连接、聚合等。

```
=k("JSA.getIndexs", 1, 10, 2)   →  [1, 3, 5, 7, 9] 自动 spill 到相邻单元格
=k("x => x * 2", 5)             →  10
=k("$$.superPivot", A1:F100, "f1", "f2", "count()")  →  2D 交叉透视表
```

## 2. 调用签名

```
k(fn, ...args)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| `fn` | string | 函数表达式字符串，支持 6 种语法（见 §3） |
| `...args` | any | 传递给 fn 的参数，可为数字、字符串、Range 引用 |

返回值根据 fn 的执行结果自动 unwrap：1×1 返回标量，N×1 / 1×N 返回 1D 数组（支持 SUM 等聚合），N×M 返回 2D 数组（WPS 自动 spill）。

## 3. fn 支持的 6 种语法

### 3.1 路径调用

直接调用 JSA880 命名空间下的方法。

```
=k("JSA.getIndexs", 1, 10, 2)        →  [1,3,5,7,9]
=k("$$.superPivot", A1:F100, "f1", "f2", "count()")
=k("Array2D.leftjoin", A1:B10, D1:E10, "f1", "f1")
=k("$$.distinct", A1:B100, 1)
```

路径解析支持 3 个根命名空间（`JSA.` / `$$.` / `Array2D.`），底层通过别名注册自动互认，三条路径等效。

**固定前置参数**：可以在路径中预填部分参数：

```
=k("JSA.getIndexs(1,10)", 2)   →  等效于 k("JSA.getIndexs", 1, 10, 2)
```

### 3.2 Lambda 箭头函数

```
=k("x => x * 2", 5)                     →  10
=k("(a, b) => a + b", 3, 4)             →  7
=k("(...args) => args.join(',')", 1,2,3) →  "1,2,3"
=k("(arr, v) => arr.filter(x => x > v)", A1:A10, 5)
```

### 3.3 索引选择器

用 `$0`, `$1` 引用参数位置。

```
=k("$0 * 2", 5)           →  10
=k("$0 + $1", [1,2,3], 1) →  [2,3,4]
```

### 3.4 列选择器（依赖 Array2D fN 代理）

Array2D 实例支持 `f1`, `f2`... 列属性代理。

```
=k("row => [row.f2, row.f3]", A1:C10)
```

在 chainable 表达式中更常用（见 §4）。

### 3.5 多行 JSA 代码块

以 `return` 语句结束的完整代码块。

```
=k("var s=0; for (var i=0; i<args.length; i++) s += args[i]; return s;", 1,2,3,4,5)  →  15
```

### 3.6 Chainable 链式表达式

见 §4。

## 4. Chainable 链式表达式

当 fn 字符串包含 `.filter(` `.map(` `.slice(` `.sort(` `.reduce(` `.take(` `.skip(` 等链式调用时，`k()` 自动编译为 IIFE 包装的执行体。

```
=k("arr => arr.filter(x => x.f2 === 'Product1').map(x => [x.f2, x.f3])", A1:H100)
```

`$$` 作为 Array2D 的自由变量在 IIFE 内自动捕获，首参数如果是 plain array 会自动 wrap 成 Array2D 实例以支持 `f1`, `f2`... 列属性代理。

```
=k("(...args) => $$.superPivot(...args).filter((x,i) => i===0 || x.f2==='Product1')", A1:H40, "f1", "f2", "count()")
```

## 5. __KJ_ARGS__ 机制

WPS 公式引擎在多字符串参数场景下有时会吞掉参数。`__KJ_ARGS__` 允许在第一个 fn 字符串中内嵌 JSON 标记来显式传递被吞的参数。

```
=k("__KJ_ARGS__={\"rowFields\":\"f3,f2\",\"colFields\":\"f6\"} (...args)=>$$.superPivot(...args)", A1:H40, "count()")
```

支持字段：`rowFields` `colFields` `dataFields` `headerRows`。

## 6. 内部架构

```
单元格 =k(fn, ...args)
  │
  ├─ function k(fn)           ← 顶层 UDF shim (JSA880.js:19752)
  │
  └─ JSA.k(fn, ...)           ← 包装层 (JSA880.js:3454)
       │                        校验 fn + 错误定位
       │
       └─ JSA.jsaLambda(fn, ...)  ← 核心引擎 (JSA880.js:3125)
            │
            ├── __KJ_ARGS__ 提取与注入
            ├── smartUnwrap: Range → Value2, host array → 真 Array
            ├── 路径解析:  JSA.xxx / $$.xxx / Array2D.xxx
            ├── Lambda 编译: 箭头函数 / 多行代码 / 索引选择器
            ├── Chainable 编译器: _kParseChainableExpression (JSA880.js:3074)
            └── 执行 → unwrap 结果 (N×1→1D, N×M→2D, 1×1→标量)
```

**关键设计决策：**

- WPS macOS 上 Range 作为 `function` 类型传入，`smartUnwrap` 有专门路径 `typeof v === 'function' && v.Value2`
- 别名注册表 `_registerJSAAliases`（~3527 行）确保 `JSA.superPivot` / `$$.superPivot` / `Array2D.superPivot` 全部等效
- Chainable IIFE 内 `var $$ = Array2D` 解决自由变量作用域
- 返回结果根据 shape 智能 unwrap：保留 1×1 标量、flatten N×1 为 1D（兼容 SUM）、保留 N×M 做 WPS spill

## 7. 错误码速查

| 错误 | 含义 | 解决方向 |
|------|------|---------|
| `#K_ERR: pos=0, FN` | fn 字符串解析失败 | 检查语法、引号配对 |
| `#K_ERR: pos=0, TYPE` | 类型错误 | 检查参数类型，Range 是否正确传参 |
| `#K_ERR: pos=1, INTERNAL` | JSA880 框架未加载 | 确认 JSA880.js 已作为加载项启用 |
| `#NAME?` | WPS 找不到 UDF | 确认 ThisWorkbook 模块已加载 JSA880.js |
| `#SPILL!` | spill 范围被阻挡 | 检查 sheet XML 中目标区域是否有静态 cell |
| `#NUM!` | 数值计算错误 | 检查 chainable 表达式中的列引用 |

## 8. 常用公式速查

### 数组生成

```
=k("JSA.getIndexs", 1, 10, 1)     →  1,2,3,4,5,6,7,8,9,10
=k("JSA.getIndexs", 1, 10, 2)     →  1,3,5,7,9
```

### 筛选

```
=k("arr => arr.filter(x => x.f2 === 'Product1')", A1:H100)
=k("$$.filter", A1:H100, "f2", "Product1")
=k("$$.z筛选", A1:H100, "f2", "Product1")
```

### 透视 (superPivot)

```
=k("$$.superPivot", A1:H40, "f1", "f2", "count(),sum(\"f3\")")
=k("$$.superPivot", A1:H40, ["f1,f2","产品,月份"], ["f5","地区"], ["count(),sum(\"f3\")","计数,金额"])
```

### 去重 / 分组 / 排序

```
=k("$$.distinct", A1:B100, 1)                    →  按 f1 去重
=k("$$.sort", A1:H100, ["f1+","f2-"], 1)         →  多列排序
=k("$$.z分组", A1:H100, "f1", "sum(\"f3\")")     →  分组汇总
```

### 连接

```
=k("$$.leftjoin", A1:B10, D1:E10, "f1", "f1")    →  左连接
=k("$$.textjoin", A1:A10, ",")                    →  文本合并
```

### 链式组合

```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i===0||x.f2==='Product1')", A1:H40, "f1", "f2", "count()")
```

## 9. 调试技巧

**最简探针**：从最简单的调用开始逐层验证。

```
=k("a => 1 + 1")                  →  测 lambda 路径通不通
=k("a => typeof a", A2:B7)        →  看参数类型（macOS 返回 "function"）
=k("a => Array.isArray(a)", A2:B7) →  看是不是数组
=k("JSA.getIndexs", 1, 5, 1)     →  测路径调用 + 1D spill
```

**排查顺序（遇到 `#SPILL!`）：**

1. 检查 sheet XML 中 spill 范围有无静态 cell 阻挡
2. 检查 JSA.k 末尾 return 的是不是 2D 数组（而不是 wrappedResult 对象）
3. 确认 `fullCalcOnLoad="1"` 已设置

## 10. 代码位置索引

| 组件 | 文件 | 行号 | 行数 |
|------|------|------|------|
| 顶层 `function k()` | JSA880.js | 19752 | 3 |
| `JSA.k` | JSA880.js | 3454 | ~40 |
| `JSA.jsaLambda` | JSA880.js | 3125 | ~330 |
| `_kParseChainableExpression` | JSA880.js | 3074 | ~51 |
| `JSA.k.help` | JSA880.js | 3494 | ~18 |
| `_registerJSAAliases` | JSA880.js | 3527 | ~57 |
| `JSA.z解析函数表达式` | JSA880.js | 3598 | ~235 |

k() 函数体系总代码量约 **730 行**（含注释和空行）。

## 11. 最佳实践

- **优先用路径调用**（`k("JSA.xxx", ...)`）而非手写 lambda，代码更短且久经测试
- **多参场景用 `__KJ_ARGS__`** 标记避免 WPS 吞参
- **macOS 用户注意 Range 是 function 类型** — 这是 WPS JSA macOS 的已知行为，`smartUnwrap` 已处理
- **每次修改 xlsm 文件前 `cp .bak` 备份** — JSA880.js 整体替换会覆盖 codemodule 增量改动
- **用 ET（ElementTree）操作 codemodule codetext**，不要用 regex 直接改 XML

