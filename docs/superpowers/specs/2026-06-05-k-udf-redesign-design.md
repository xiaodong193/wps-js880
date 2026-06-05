# KO k() UDF 重新设计 · 设计文档

> 日期:2026-06-05
> 目标文件:`js880/JSA880.js`(主改) + `KO一切的k函数.xlsm` 的 ThisWorkbook(3-5 行 wrapper)
> 数据来源:`KO一切的k函数.xlsm`(9 个测试公式)
> **v2 修订**:用户要求「重写框架 + 都合并到 jsa880.js」。本次只重写 k() 相关部分(jsaLambda + $$ 路径 + 4 个预处理 + JSA.k),JSA880 其他部分(Array2D 方法 / JSA 命名空间工具 / DateUtils / IO)保持不动。

---

## 1. 背景与动机

### 1.1 现状

`/第03章/3-28/` 目录里 k() UDF 有 **3 个并存版本** + JSA880 内嵌版,逻辑分散:

| 文件 | 作用 | 问题 |
|---|---|---|
| `KO_k_udf.js` | ES5 兜底 | 单独文件,k() shim 逻辑和 JSA880 分离 |
| `KO一切的k函数_UDF模块.js` | 走 JSA.jsaLambda 调度 | IIFE 启动副作用 + `...args` 写法 |
| `JDEData.bin` 内嵌的 v4.2.2 | 在 xlsm 里直接生效 | k() 逻辑和 JSA880 框架代码混在一起 |
| `js880/JSA880.js` v4.2.2 | 主框架,含 jsaLambda | 不导出 `JSA.k`,外部只能粘独立文件 |

**问题**:k() 的实现代码被分散到 4 个地方,改一处必须改多处。

### 1.2 测试信号(Excel 里的 9 个公式)

| 位置 | 公式 | 当前报错 |
|---|---|---|
| Sheet5!E2 | `=k("$$.leftjoin",D2:D4,A2:B7,"f1","f1","a.f1,b.f2")` | `#NAME?` |
| Sheet5!H2 | `=k("(a,v)=>a.filter(x=>x[1]>v)",A2:B7,O1)` | `#NAME?` |
| Sheet5!N2 | `=SUM(k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1))` | `#NAME?` |
| Sheet5!Q2 | `=k("arr=>$$.distinct(arr,'f1')",A2:B7)` | `#NAME?` |
| Sheet5!T2 | `=k("(arr,v)=>arr.map(x=>[x.f1,x.f2+v])",A2:B7,O1)` | `#NAME?` |
| Sheet6!F9 | `=k("(...args)=>$$.insertCols($$.superPivot(...args),-1,x=>x.sum())",A1:C17,"f1","f2","sum('f3')",1,0)` | `#NAME?` |
| Sheet6!F15 | `=k("(x,y)=>{y=y+1;return x.map(a=>[a[0],a[2]+'-'+y])}",A2:C17,1)` | `#NAME?` |
| Sheet8!J1 | `=k("$$.superPivot",A1:H40,"f3,f2","f6","sum(\`f4*f5\`),textjoin(\`f4+'*'+f5\`,\`+\`)")` | `#NAME?` |
| Sheet9!J1 | `=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 \|\| x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(\`f4\`),textjoin(\`f4\`,\`+\`)")` | `#NAME?` |

### 1.3 设计目标

1. **让 9 个公式全部跑通**
2. **k() 实现代码全部合并到 JSA880.js**(`JSA.k` 作为新公共 API)
3. **ThisWorkbook 只需 3-5 行 wrapper**(`function k() { return JSA.k.apply(null, arguments); }`)
4. **3 个外部 k() 文件 + JDEData.bin 内嵌版全部消失**
5. **错误信息带位置 + 类别**
6. **只重写 k 相关部分**(jsaLambda + $$ 路径 + 4 个预处理 + JSA.k)

---

## 2. 架构

### 2.1 总体结构(修订后)

**核心原则**:k() 的实现代码全在 JSA880.js 里,ThisWorkbook 只是个 3-5 行的转发 wrapper。

```
js880/JSA880.js  (单文件,770KB → 771KB,新增约 300 行)
├── [现有] Array2D.leftjoin / superPivot / distinct / insertCols
├── [现有] JSA 命名空间工具(text / date / number / vlookup 等)
├── [现有] JSA.jsaLambda(内部 dispatcher)
├── [新增] JSA.k(fn, ...args)        ← k() 完整实现(含 4 个预处理)
├── [新增] $$ = Array2D               ← 顶部自动兜底
├── [新增] _kNormalizeBacktick       ← 反引号转换
├── [新增] _kNormalizeArg            ← Range→Value2
├── [新增] _kNormalizePath           ← $$ 路径解析
├── [新增] _kWrapError               ← 错误位置化
└── [新增] _kRunTest                 ← JSA.k.test() 自检入口

KO一切的k函数.xlsm 的 ThisWorkbook 代码模块(3-5 行)
├── function k(fn) { return JSA.k.apply(null, arguments); }       ← UDF 注册
└── function jsaLambda(fn) { return JSA.k.apply(null, arguments); }  ← UDF 别名
```

**为什么 ThisWorkbook 还需要 wrapper**:
- WPS 公式引擎**只扫描 ThisWorkbook 代码模块的顶层 function** 作为 UDF
- 加载项(JSA880.js)里的顶层 function 不会被注册
- 所以 `function k(){}` 必须出现在 ThisWorkbook,但实现逻辑全在 JSA.k

### 2.2 数据流(修订后)

```
WPS 公式 =k(fn, arg1, arg2, ...)
    ↓
ThisWorkbook: k(fn, ...args)  (3-5 行 wrapper)
    ↓
JSA.k(fn, ...args)  (JSA880.js 内)
    ↓
1. _kNormalizeBacktick(fn)        ← "sum(`f4`)" → 'sum("f4")'
2. args.forEach(_kNormalizeArg)    ← Range → Value2; 1x1 → scalar
3. JSA.jsaLambda(normalizedFn, ...normalizedArgs)  (现有 dispatcher)
    ↓
4. _kWrapError(result)             ← 错误加 pos + kind
    ↓
返回结果 / "#K_ERR: ..."
```

### 2.3 依赖(修订后)

- **JSA880.js**(作为 WPS 加载项加载,提供 `JSA.k` / `JSA.jsaLambda` / `Array2D.leftjoin` 等所有 API)
- **ThisWorkbook 代码模块** 的 3-5 行 wrapper(UDF 注册必需)
- **不依赖** 任何 `KO_k_udf.js` / `KO一切的k函数_UDF模块.js` / JDEData.bin 内嵌版(全部废弃)

---

## 3. API 设计

### 3.1 公式 UDF(在 ThisWorkbook 注册)

#### `k(fn, ...args)` — UDF

**签名**:`k(fn, ...args) → any`

**ThisWorkbook wrapper**(3-5 行):
```javascript
function k(fn) { 
    return JSA.k.apply(null, arguments); 
}
```

#### `jsaLambda(fn, ...args)` — UDF 全名别名

**ThisWorkbook wrapper**(3-5 行):
```javascript
function jsaLambda(fn) { 
    return JSA.k.apply(null, arguments); 
}
```

### 3.2 内部实现(在 JSA880.js 里,作为 `JSA.k`)

#### `JSA.k(fn, ...args)` — 实际实现

**签名**:`JSA.k(fn, ...args) → any | "#K_ERR: ..."`

**实现流程**:
1. 调 `_kNormalizeBacktick(fn)` 处理反引号
2. 遍历 args 调 `_kNormalizeArg(arg)` 处理 Range / 1x1
3. 调 `JSA.jsaLambda(normalizedFn, ...normalizedArgs)`
4. 调 `_kWrapError(result)` 处理 null / undefined / 错误
5. 返回结果 / 错误字符串

**fn 接受的 6 种语法**(由 `JSA.jsaLambda` 处理,本次重写 dispatcher):

1. 路径调用:`k("JSA.getIndexs", 1, 10, 2)`
2. 路径别名:`k("$$.leftjoin", ...)` ← 自动 fallback 到 `Array2D.leftjoin`
3. Lambda 箭头:`k("x => x * 2", 5)`
4. 多参 Lambda:`k("(a, v) => a.filter(x => x[1] > v)", arr, val)`
5. Block Body:`k("(x, y) => { y = y + 1; return x.map(...); }", arr, 1)`
6. 索引选择器:`k("$0 + $1", [10, 20])` / `k("f1 + f2", row)`

**args 接受的类型**:

| 类型 | 处理方式 |
|---|---|
| Range 对象(WPS 公式传进来的) | 自动取 `.Value2` |
| 1x1 二维数组 | 展平为 scalar |
| 1xN / Nx1 / NxM 二维数组 | 保持原样 |
| 字符串 | 原样透传(允许空串 `""`) |
| 数字 / boolean | 原样 |
| null / undefined | 抛错 `#K_ERR: pos=N, ARG, msg="参数为空"` |

#### `JSA.k.help()` — 排错指南(不是公式)

JSA 编辑器里手动调 `JSA.k.help()`,控制台打印排错清单。

#### `JSA.k.test()` — 自检脚本(不是公式)

JSA 编辑器里手动调 `JSA.k.test()`,跑 9 个测试公式的 Node 仿真版,打印每条通过 / 失败。

---

## 4. 预处理细节

### 4.1 `_kNormalizeBacktick(s)`

**目的**:WPS 公式引擎不会截断 `\`...\`` 在双引号内,但 `JSA.jsaLambda` 内部编译成 `new Function('return ...')` 时,反引号会被当成模板字符串(可能解析失败或被某些 WPS 版本禁止)。

**规则**:
- 找到所有 `` `([^`]*?)` ``(非贪婪)
- 如果 inner 不含 `${` 也不含 `\` 转义,转成 `"..."`
- 否则原样保留(可能是真模板字符串)

**示例**:
```js
_kNormalizeBacktick('sum(`f4*f5`)')           // 'sum("f4*f5")'
_kNormalizeBacktick('`a${b}c`')              // "`a${b}c`"  保留
_kNormalizeBacktick('`f1`+`f2`')             // '"f1"+"f2"'
```

### 4.2 `_kNormalizeArg(a)`

**目的**:WPS 把单元格区域传给 UDF 时是 Range 对象,不是二维数组;单个单元格是 1x1 二维数组,但 `JSA.jsaLambda` 内部很多函数(去重 / 筛选 / 透视)需要 1D 数组。

**规则**(按顺序):
1. 如果 a 有 `.Address` 属性且 `.Value2` 可访问 → 返回 `a.Value2`
2. 如果 a 是 1x1 二维数组(只有一个元素,且该元素是单元素数组) → 返回 `a[0][0]`
3. 其他 → 原样返回

**不处理**的情况:
- 1xN / Nx1 / NxM 二维数组:保持原样
- 空串 `""`:保持原样(关键,不能被当成"空参数")

### 4.3 `_kNormalizePath(fn)`(可选,如果 JSA.jsaLambda 不自动处理)

**目的**:`$$.leftjoin` 中的 `$$` 必须是 `Array2D` 的别名,但 JSA880 v4.2.2 的 `JSA.jsaLambda` 应该已经做了这层 fallback。**本设计先不实现**,验证 jsaLambda 行为后决定。

### 4.4 `_kWrapError(e, ctx)` / `_kWrapResult(r, ctx)`

**错误格式**(用户确认采用紧凑格式):
```
#K_ERR: pos=0, FN, msg="无法解析 fn:'...'"
#K_ERR: pos=2, ARG, msg="Range 转数组失败"
#K_ERR: pos=1, INTERNAL, msg="JSA880 框架未加载"
#K_ERR: pos=0, FN, msg="jsaLambda 返回 null"
```

**`pos` 含义**:
- `pos=0` → fn 本身
- `pos=N>0` → 第 N 个参数(从 1 开始)

**`kind` 枚举**:
- `FN` — fn 字符串解析失败
- `ARG` — 参数处理失败
- `RANGE` — Range 转 Value2 失败
- `TYPE` — 参数类型不匹配(如期望数组,给了对象)
- `INTERNAL` — JSA880 框架未加载 / 其他内部错误

**正常结果**:
- scalar → 原样
- 1D 数组 → 让 WPS 数组溢出(spill)
- 2D 数组 → 同上
- null / undefined → `#K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined"`

---

## 5. 错误处理(完整示例)

| 触发 | 错误信息 |
|---|---|
| `=k()` 不带参数 | `#K_ERR: pos=0, FN, msg="fn 不能为空"` |
| `=k("JSA.xxx", A1)` 但 JSA880 没加载 | `#K_ERR: pos=1, INTERNAL, msg="JSA880 框架未加载,请加载 JSA880.js 加载项"` |
| `=k("(a)=>a.b", A1)` A1 是个字符串但代码要 .b | `#K_ERR: pos=0, INTERNAL, msg="TypeError: a.b is not a function"` |
| `=k("$$xxx")` $$xxx 不存在 | `#K_ERR: pos=0, FN, msg="找不到路径 '$$xxx'"` |
| `=k("...args => ...", A1)` WPS 旧版不支持 rest | `#K_ERR: pos=0, FN, msg="fn 语法不被支持:'...args'"` |

---

## 6. 测试策略

### 6.1 Node 仿真测试 — `_test_k_v5.js`

新建文件 `第03章/3-28/_test_k_v5.js`,在 Node 环境下跑 9 个测试公式的仿真版(把 Range 换成 JSON 数据)。

**9 个测试用例**(对应 Excel 里的 9 个公式,见 §1.2):

| # | 公式摘要 | 仿真数据 |
|---|---|---|
| 1 | `k("$$.leftjoin", d2d4, a2b7, "f1", "f1", "a.f1,b.f2")` | 2 个 JSON 数组 |
| 2 | `k("(a,v)=>a.filter(x=>x[1]>v)", a2b7, o1)` | 2D 数组 + scalar |
| 3 | `k("(arr,v)=>arr.map(x=>[x.f2+v])", a2b7, o1)` | 2D 数组 + scalar |
| 4 | `k("arr=>$$.distinct(arr,'f1')", a2b7)` | 2D 数组 |
| 5 | `k("(arr,v)=>arr.map(x=>[x.f1,x.f2+v])", a2b7, o1)` | 2D 数组 + scalar |
| 6 | `k("(...args)=>$$.insertCols($$.superPivot(...args),-1,x=>x.sum())", ...)` | 6 个参数 |
| 7 | `k("(x,y)=>{y=y+1;return x.map(a=>[a[0],a[2]+'-'+y])}", a2c17, 1)` | 2D 数组 + scalar |
| 8 | `k("$$.superPivot", a1h40, "f3,f2", "f6", "sum(\`f4*f5\`),...")` | backtick |
| 9 | `k("(...args)=>$$.superPivot(...args).filter(...)", ..., "", ...)` | rest + spread + filter + 空串 |

**测试运行方式**:
```bash
cd "/Users/daidai193/Library/CloudStorage/SynologyDrive-code/JS880教案/第03章/3-28/"
node _test_k_v5.js
# 期望:9 PASSED,0 FAILED
```

**测试要求**:
- 加载 `JSA880.js`(在 WPS 外模拟)
- 加载 `KO_k_udf_v5.js`
- 把 9 个公式的 `k(...)` 调用剥出来跑
- 打印每条结果(成功 / 错误)

### 6.2 错误注入测试 — 5 个用例

| # | 注入 | 期望错误信息 |
|---|---|---|
| E1 | `k()` 空调用 | `#K_ERR: pos=0, FN, msg="fn 不能为空"` |
| E2 | `k("JSA.xxx")` JSA880 未加载 | `#K_ERR: pos=1, INTERNAL, msg="JSA880 框架未加载..."` |
| E3 | `k("$$yyy")` $$yyy 不存在 | `#K_ERR: pos=0, FN, msg="找不到路径 '$$yyy'"` |
| E4 | `k("x=>x.b", "abc")` 类型错 | `#K_ERR: pos=0, INTERNAL, msg="TypeError: x.b is not a function"` |
| E5 | `k("`a${b}c`", 1)` 模板字符串(不应被改) | 保留 `\`a${b}c\`` |

### 6.3 WPS 实际验证

把 `KO_k_udf_v5.js` 粘到 `KO一切的k函数.xlsm` 的 ThisWorkbook 模块,重启 WPS,打开 9 个公式 cell,确认全部不报 `#NAME?` / `#VALUE!`。

---

## 7. 部署流程(修订后)

### 7.1 一次性操作(用户)

1. WPS → 选项 → 加载项 → 加载 `js880/JSA880.js` v5.0(已包含 JSA.k)
2. 重启 WPS

### 7.2 每个 xlsm 文件需要做的

1. WPS → 开发工具 → JSA 编辑器
2. 左侧工程树 → ThisWorkbook
3. 粘入 **3-5 行 wrapper**:
   ```javascript
   function k(fn) { return JSA.k.apply(null, arguments); }
   function jsaLambda(fn) { return JSA.k.apply(null, arguments); }
   ```
4. Ctrl+S 保存
5. 关闭 JSA 编辑器

### 7.3 验证

打开 xlsm,在任意单元格输入 `=k("JSA.getIndexs", 1, 5, 1)`,期望显示 `1 2 3 4 5`(数组溢出)。

---

## 8. 文件清单(修订后)

### 8.1 改动

- `js880/JSA880.js` — **v4.2.2 → v5.0**,新增约 300 行:
  - 顶部:`$$ = Array2D` 兜底
  - `JSA.k(fn, ...args)` — k() 完整实现
  - `JSA.k.help()` / `JSA.k.test()`
  - `_kNormalizeBacktick` / `_kNormalizeArg` / `_kNormalizePath` / `_kWrapError`
  - 重写 `JSA.jsaLambda` 的 dispatcher(更稳的 ES5 编译)
  - Console 启动日志加上 `JSA.k` ready 自检

- `KO一切的k函数.xlsm` — **ThisWorkbook 模块**替换为 3-5 行 wrapper

### 8.2 新建

- `JS880教案/第03章/3-28/_test_k_v5.js`(~150 行,Node 仿真测试)

### 8.3 标记废弃(保留作为参考,不动)

- `JS880教案/第03章/3-28/KO_k_udf.js`(v1 ES5 兜底)→ 文件名加 `[DEPRECATED]` 前缀
- `JS880教案/第03章/3-28/KO一切的k函数_UDF模块.js`(v2 启动器版本)→ 同上

### 8.4 文档更新

- `JS880教案/第03章/3-28/KO一切的k函数.md`:
  - 更新"30 秒上手"段:ThisWorkbook 只需 3-5 行 wrapper,k() 实现已并入 JSA880.js
  - 更新"配套文件"表格(去掉 KO_k_udf.js / UDF模块.js,加上 JSA.k 说明)
  - 更新 FAQ 的 Q1 / Q2(指向 JSA880.js 加载项)
  - 添加"v5.0 变更日志"段(说明合并动作)

- `js880/JSA880.js` 内部:
  - 顶部加版本号 `v5.0` 和 changelog 段
  - JSA.k 段加 docstring

### 8.5 Excel 文件

- `KO一切的k函数.xlsm`:
  - ThisWorkbook 替换为 3-5 行 wrapper
  - (可选)把 9 个测试公式 cell 的 value 改正确(让用户看到正确结果而非 #NAME?)

---

## 9. 风险与边界

### 9.1 已知风险

| 风险 | 影响 | 缓解 |
|---|---|---|
| WPS 15990 以下不支持数组溢出 | 3、4、5、6、9 公式显示为单个值 | 文档注明需 WPS 15990+ |
| WPS 旧版 `new Function()` 不支持 ES6+ | 6、9 公式 rest/spread 失败 | `JSA.jsaLambda` 内部已用 `try/catch` 兜底;失败时返回明确错误信息 |
| WPS 公式引擎对反引号处理不一致 | 8 公式行为不可预测 | 文档测试,若 WPS 真的截断,在公式里改用双引号 |
| $$ 命名空间被 jsaLambda 抢占 | 公式 1、4 找不到 $$ | 顶部自动 `$$ = Array2D` 兜底 |

### 9.2 不在本次范围

- 重新设计 `JSA.jsaLambda` 本身(只调,不改)
- 重新设计 `JSA880` 框架其他 API
- 写新版本的 k() 不依赖 JSA880
- 改 Excel 里的数据 / 公式

---

## 10. 验收标准(修订后)

✅ 全部满足才算完成:

1. `js880/JSA880.js` v5.0 文件存在,Console 启动日志有 `JSA.k ready` 字样
2. `JSA.k` 包含 4 个预处理 + 错误包装,9 个测试公式在 JSA.k 下不报 `null/undefined` 类错误
3. `_test_k_v5.js` 在 Node 下跑 9 个公式 + 5 个错误注入,**全 PASS**
4. `KO一切的k函数.xlsm` 的 ThisWorkbook 粘入 3-5 行 wrapper,9 个公式 cell **不报 #NAME? / #VALUE!**
5. `KO一切的k函数.md` 文档更新完成,3 个旧 k() 文件标记 `[DEPRECATED]`
6. JSA880.js v5.0 加载后,`Application.JSA880` / `JSA` / `Array2D` 全部可用,既有功能不退化
