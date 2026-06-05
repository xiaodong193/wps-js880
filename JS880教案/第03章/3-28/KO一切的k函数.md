# KO 一切的 k 函数 · WPS JSA 自定义函数(UDF)

> 目标:让任何 JSA 框架方法(包括超级透视 / 筛选 / leftjoin / textjoin / 自定义 lambda)都能在 WPS 表格里**像 SUM/VLOOKUP 一样当公式用**。

---

## 🎯 30 秒上手

### 🔑 关键机制(必看)

WPS 公式引擎识别 UDF 的**唯一方式**:**ThisWorkbook 代码模块**里的**顶层 `function` 声明**。

| 部署位置 | 公式能否识别? |
|---|---|
| ThisWorkbook 模块顶层 `function k(){}` | ✅ 能 |
| 加载项(JSA880.js)的顶层 `function` | ❌ WPS 公式引擎**不扫描加载项** |
| `this.k = JSA.k` / `var k = function(){}` | ❌ 都不是顶层 function |

**所以,正确的部署姿势是两步** —— 缺一不可:

| 步骤 | 动作 | 作用 |
|---|---|---|
| ① 加载 JSA880.js | WPS → 选项 → 加载项 → 加载 `js880/JSA880.js` | 提供 `JSA` / `Array2D` / `IO` 等所有框架 API |
| ② 粘 UDF 模块到 **ThisWorkbook** | JSA 编辑器 → ThisWorkbook → 粘入 `KO一切的k函数_UDF模块.js` 全部内容 | 公式引擎扫描到顶层 `function k()`,注册成 UDF |

> ⚠️ **只做 ① 不做 ②** → 框架能用,但 `=k(...)` 公式报 `#NAME?`  
> ⚠️ **只做 ② 不做 ①** → 公式能识别 k(),但 `=k("JSA.xxx",...)` 报 `#K_ERR: undefined`(JSA 命名空间不存在)

---

### 🚀 完整操作步骤

打开你的 `xxx.xlsm` 工作簿,按顺序:

1. **加载 JSA880.js 框架**(一次性,以后所有 xlsm 都共享)
   - WPS → 选项 → 加载项 → 添加 `js880/JSA880.js` 文件 → 重启 WPS
   - 看到 Console 打印 `[JSA880 v4.2.2] KO一切k函数 UDF 已整合到主框架` 即成功
2. **粘 3-5 行 wrapper 到 **ThisWorkbook** (v5.0 新方案)**
   - WPS → 开发工具 → JSA 编辑器
   - 在左侧工程树找到 **ThisWorkbook** 模块(注意不是"模块1"那种普通模块!)
   - 双击打开 → 粘入以下 3-5 行代码(就这 3-5 行,k() 实现全在 JSA880.js):
     ```javascript
     function k(fn) { return JSA.k.apply(null, arguments); }
     function jsaLambda(fn) { return JSA.k.apply(null, arguments); }
     ```
   - `Ctrl + S` 保存 → 关闭 JSA 编辑器
3. **验证**: 在任意单元格输入公式

```
=k("JSA.getIndexs", 1, 10, 2)
```

看到 `1 3 5 7 9` 数组溢出 → **大功告成!** 🎉

### 🆘 找不到 ThisWorkbook 模块?

WPS → 开发工具 → JSA 编辑器,如果左侧工程树没看到 ThisWorkbook:
1. 看左侧工程树顶部是否有"工程"名
2. 工程名右键 → 属性 → 把"加载项类型"切到 **"文档级"**(而不是"用户级"或"应用级")
3. 看到 ThisWorkbook 后双击即可

---

## 📐 设计原理

WPS 公式引擎只扫描 **ThisWorkbook 代码模块** 的**顶层 `function`** 作为 UDF。
**不扫描加载项**(即使加载项里定义了 `function k()`,WPS 公式也找不到)。

| 写法 | 公式能识别? | 原因 |
|---|---|---|
| `this.k = JSA.k` | ❌ | 不是顶层 function,WPS 不注册 |
| `var k = function(){}` | ❌ | var 声明,WPS 不注册 |
| 加载项里的 `function k(){}` | ❌ | 公式引擎不扫描加载项 |
| ThisWorkbook 顶层 `function k(fn, ...args){}` | ✅ | 顶层 function → UDF |

所以最关键的两行:

```javascript
function k(fn, ...args) {
    return JSA.jsaLambda(fn, ...args);
}
function jsaLambda(fn, ...args) {
    return k(fn, ...args);
}
```

内部把"字符串 lambda"丢给 `JSA.jsaLambda` 处理(它已经实现了 6 种语法、缓存、容错、Range 自动 Value2 等所有逻辑)。

---

## 🚀 支持的 6 种 fn 语法

`JSA.jsaLambda(fn, ...args)` 是核心调度器,`fn` 可以是以下任一种:

### ① 路径调用(最常用)

```javascript
k("JSA.getIndexs", 1, 10, 2)           // → [1,3,5,7,9]
k("JSA.cint", 3.7)                       // → 3
k("JSA.today")                           // → "2026-06-05"
k("Array2D.superPivot", data, [...])     // 超级透视
```

> 三个根命名空间都会自动回退查找:`JSA.xx` / `$$` / `Array2D` / 全局函数。  
> 公式里写 `k("JSA.getIndexs", ...)` 和 `k("Array2D.getIndexs", ...)` 都能跑!

### ② Lambda 箭头函数

```javascript
k("x => x * 2", 5)                      // → 10
k("(a, b) => a + b", 3, 4)               // → 7
k("arr => arr.length", [1,2,3,4,5])      // → 5
```

### ③ `$0/$1` 索引选择器

```javascript
k("$0 * $1", [3, 4])                    // → 12  (新写法:接收 1D 数组)
k("$0 + $1", [10, 20])                  // → 30
```

### ④ `f1/f2` 列选择器

```javascript
k("f1 + f2", [10, 20])                  // → 30
```

### ⑤ `(...args)` 多参数

```javascript
k("(...args) => args.join(',')", 1, 2, 3)  // → "1,2,3"
```

### ⑥ `-r` Range 模式

```javascript
k("rng => rng.Address()", -r, "A1:B3")  // 拿到 A1:B3 的 Range 对象
k("(rng, fn) => fn(rng.Value2)", -r, "A1", "(v) => v.length")  // 拿到 A1 值数组长度
```

> `-r` 后面的所有字符串如果长得像 Excel 区域地址(`A1`, `A1:B3`),会被自动 `Range()` 转换。  
> 多数情况下你**不需要**用 `-r`,因为 `JSA.jsaLambda` 内部有智能 Range 检测(见下面容错章节)。

---

## 🛡️ WPS 公式容错(写公式不用操心)

WPS 公式引擎把单元格传进 UDF 时有 3 个坑,`JSA.jsaLambda` 全部处理好了:

### 坑 1:Range 对象直接传 → 多数函数不认识

WPS 把 `A1:H40` 传进公式时,得到的是 Range 对象,而 `filter`/`map`/`superPivot` 都需要 `.Value2` 二维数组。

✅ 自动检测: 看到 `.Address` 和 `.Value2` 就自动取 `.Value2`。

```javascript
// 你写: =k("arr => arr.length", A1:H40)
// 内部: 自动把 Range 转成 [[..],[..]],arr 是 2D 数组
```

### 坑 2:`=k("f1", A1)` 传过来是 `[[1]]`,不是 `[1]`

WPS 把单个单元格传过来是 1x1 二维数组。但很多函数(如 `textjoin` 的列选择器)需要 1D 数组。

✅ 智能 flatten: **只对 1x1 二维数组 flatten 为单值**,其他(1xN / Nx1 / NxM)都保持原样。

```javascript
// 你写: =k("f1", A1)         // 1x1 [[1]] → flatten 为 1
// 你写: =k("...", A1:A10)    // Nx1 [[1],[2],...] → 保持 2D (否则 filter/map 全部瘫痪!)
// 你写: =k("...", A1:H40)    // NxM → 保持 2D
```

> ⚠️ 历史版本曾对 1xN 也 flatten,导致 `A1:C1` 这种 1 行 N 列的 Range 数据被破坏,**已修复**。

### 坑 3:反引号(模板字符串)在 WPS 公式里报 `#NAME?`

公式引擎看到 `` ` `` 就懵,无法识别为 JS 模板字符串。

✅ 自动转换: 不含 `${}` 的反引号包围短串自动转成双引号(`\`f4*f5\` → "f4*f5"`)。

```javascript
// 你写: =k("Array2D.superPivot", A1:H40, "f3,f2", "f6", "sum(\`f4*f5\`),textjoin(\`f4+'+'+f5\`, \`+\`)")
// 内部: 反引号包围的 sum("f4*f5") 字符串能正确解析
```

---

## 🧪 真实公式样例

```excel
' 1. 快速算等差数列
=k("JSA.getIndexs", 1, 10, 2)
' → 1 3 5 7 9 (数组溢出)

' 2. 单元格数据乘 2
=k("x => x * 2", A1)

' 3. 求 A 列总和
=k("JSA.z求和", A1:A100)             ' ← 需要 JSA880.js 加载

' 4. 把 A 列转大写
=k("x => String(x).toUpperCase()", A1)

' 5. 算单元格值在某个范围的比例
=k("x => x > 100 ? '高' : '低'", A1)

' 6. 同时给多个单元格算
=k("JSA.rndIntArray", 1, 100, 50)    ' ← 生成 50 个 1-100 的随机数
```

---

## 📦 配套文件

| 文件 | 作用 |
|---|---|
| `js880/JSA880.js` | **作为加载项加载 v5.0+**(提供 `JSA` / `Array2D` / `JSA.k` 等所有 API) |
| ThisWorkbook 代码模块 | **粘 3-5 行 wrapper**(注册 `k()` / `jsaLambda()` UDF 转发到 JSA.k) |
| `KO一切的k函数.md` | 本文档,使用说明 |
| `_test_k_v5.js` | Node 模拟测试 v5.0(无需 WPS),跑 9 个公式 + 5 个错误注入 + 1 个 meta,16/16 PASS |
| `[DEPRECATED]_KO_k_udf.js` | **已废弃**(v1 ES5 兜底,保留作为历史参考) |
| `[DEPRECATED]_KO一切的k函数_UDF模块.js` | **已废弃**(v2 启动器版本,保留作为历史参考) |

---

## ❓ 常见问题

### Q1: 公式报 `#NAME?`
**原因**: WPS 公式引擎找不到 `k` 函数,说明 UDF 模块没生效。  
**解决**:
- 确认 UDF 模块**粘到了 ThisWorkbook**(不是"模块1"这种普通模块!)
- 确认 `function k(fn, ...args) {}` 是模块**顶层**声明,不是 `var k = ...`
- 保存模块,关掉 JSA 编辑器,重启 xlsm(或按 F5 / 切换工作表触发 Workbook_Open)

### Q2: 公式返回 `#K_ERR: JSA is undefined` 或 `JSA.xxx is not a function`
**原因**: JSA880.js 没加载。  
**解决**: WPS → 选项 → 加载项 → 添加 `js880/JSA880.js` 并重启 WPS。  
或粘 UDF 模块到 ThisWorkbook 后会看到 JSA 控制台打印 JSA880.js 加载成功提示。

### Q3: 公式返回 `#K_ERR: xxx` (其他错误)
**原因**: `fn` 表达式语法错 或 参数不匹配。  
**解决**: 在 JSA 编辑器直接调 `JSA.jsaLambda(...)` 看完整错误,逐步调试。

### Q4: 改了源数据,k() 公式结果不更新
**解决**: 在 JSA 编辑器 `ThisWorkbook` 上,选 `SheetChange` 事件绑定到模块的 `k_onChange`,会自动标记 `=k(...)` 单元格为 dirty 触发重算。

### Q5: 公式返回单个值时不显示(数组没 spill)
**原因**: WPS 版本 < 15990,不支持数组溢出。  
**解决**: 升级 WPS Office 到 15990+ 版本。

### Q6: 怎么验证 k() 已就绪?
打开工作簿,看 JSA 控制台:
```
✅ k() UDF 已就绪!(JSA880 v4.x.x)
   自检:k('JSA.getIndexs', 1, 5, 1) = [1,2,3,4,5]
```
看到这行就说明 `=k(...)` 公式可以用了。

如果没看到,调一下 `k_help()` 看排错指南。

---

## 🏗️ 进阶:直接调用 jsaLambda

不需要 UDF 包装,在 JSA 代码里直接用更灵活:

```javascript
// 1. 路径调用
var seq = JSA.jsaLambda("JSA.getIndexs", 1, 10, 2);    // [1,3,5,7,9]
var csv = JSA.jsaLambda("Array2D.toJson", data, 2);     // JSON 字符串

// 2. Lambda
var doubled = JSA.jsaLambda("x => x * 2", 5);           // 10

// 3. 表达式求值
var sum = JSA.jsaLambda("$0 + $1", [10, 20]);            // 30
var total = JSA.jsaLambda("f1 * f2", [3, 4]);            // 12

// 4. Range 自动处理
var arr = JSA.jsaLambda("rng => rng.Value2", -r, "A1:H40");  // 显式 -r
// 或者 直接传 Range 对象(智能检测)
var arr2 = JSA.jsaLambda("v => v.length", Range("A1:H40"));   // 智能 .Value2

// 5. 自定义函数批量执行
JSA.jsaLambda("Array2D.superPivot",
    Range("A1:H40").Value2,
    ["f3,f2"], ["f6"], ["sum(`f4*f5`)"]
);
```

---

## 📜 更新日志

- **v5.0.0** (2026-06-05)
  - **重大重构:k() 实现代码全量合并到 JSA880.js**(作为 `JSA.k`)
  - ThisWorkbook 只需 3-5 行 wrapper:`function k() { return JSA.k.apply(null, arguments); }`
  - 新增 `JSA.k.help()` 排错指南
  - 错误信息格式化为 `#K_ERR: pos=N, KIND, msg="..."`(KIND ∈ FN / INTERNAL / TYPE)
  - 加 `$$` 全局别名 = Array2D
  - 加 `_kParseChainableExpression` 支持 `$$.superPivot().filter()` 链式调用
  - 顶层 `function k()` / `function jsaLambda()` 改为转发到 JSA.k(去掉 ES6 `...args`,改用 ES5 `arguments`)
  - 3 个旧文件废弃:`KO_k_udf.js` / `KO一切的k函数_UDF模块.js` / JDEData.bin 内嵌 k() 模块
  - **9 个 Excel 测试公式全部跑通**:Sheet5×5 + Sheet6×2 + Sheet8×1 + Sheet9×1
  - 配套 Node 测试 `_test_k_v5.js`:16/16 PASS(9 公式 + 5 错误注入 + 1 meta + 1 smoke)
  - 文档 / 部署流程 / changelog 全部更新

- **v4.2.2** (2026-06-05)
  - 修正 1xN / Nx1 2D 数组 flatten 的 bug(Nx1 Range 数据不能再被破坏)
  - 智能反引号转换(`\`f4\`` → `"f4"`,但 `${...}` 模板字符串保留)
  - UDF 模块增强:自动加载 JSA880.js + Workbook_Open 自检 + k_help 排错
  - JSA 命名空间补齐 20+ 别名(支持 `JSA.getIndexs` / `JSA.superPivot` / `JSA.textjoin` 等)
  - 文档明确两步部署(加载 JSA880.js + 粘 UDF 到 ThisWorkbook)

- **v4.2.1** (2026-06-04)
  - JSA 命名空间补齐 17 个缺失函数
  - 路径解析支持三根命名空间(JSA / $$ / Array2D)自动回退

- **v4.2.0**
  - 初版 `jsaLambda` 6 种语法(Lambda / 路径 / $0 / f1 / 多参数 / -r Range)
