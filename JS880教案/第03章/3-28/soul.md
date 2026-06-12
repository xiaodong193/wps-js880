# KO一切k函数 WPS JSA 调试经验总结

> **背景**:把 JSA880 框架 + k() UDF 整合到 WPS JSA on macOS,让 9 个 Excel 公式能正常工作。
> **结果**:从最初的 #NAME? 走到完整 2D 数组自动 spill。跨越 49 个 T 编号,踩了 18+ 个坑。

---

## 🎯 最终成果

| 功能 | 状态 | 关键 T 编号 |
|------|------|------------|
| `=k("JSA.getIndexs", ...)` 1D 数组 spill | ✅ | T29 |
| `=k("$$.superPivot", ...)` 2D 数组 spill | ✅ | T48+T72 |
| `=k("a=>a.filter(...)", ...)` chainable | ✅ | T29 |
| `=k("$$.leftjoin", ...)` 路径调用 | ✅ | T27 |
| `=k("arr=>$$.distinct(arr, ...)", ...)` Array2D 方法 | ✅ | T27 |
| `=k("(...)=>$$...", ...)` 复杂 lambda | ✅ | T27 |
| superPivot 数据字段 f5 求值 | ✅(`parseResultSelector` 解析时去反引号) | T49 |
| `qctextjoin` 别名支持 | ✅(在 `executeAggregation` switch 加 case) | T49 |
| superPivot 2 行表头 | ✅(options.headerRowCount=2) | T56+T70 |
| superPivot 角标题"产品" | ✅(colConfig.fields[0] 标题) | T70 |
| `SUM(k(...))` 兼容 | ✅(N×1 → 1D 让 SUM 加总) | T71 |

**`inject_t27.py`** 是最终注入脚本,保留 t15.bak 原始结构,只修 JSA880 codemodule 的 smartUnwrap + 加 spill 逻辑。

### 9. T58 — superPivot 2 行表头硬编码 + 列数修复(2026-06-06)
- **症状**:T57 全局开关有时序问题,T56 写法有 #NAME? 风险
- **修法**:
  1. JSA.k 入口无条件检测 superPivot 调用,自动 push `{headerRowCount: 2}`
  2. **关键修复**:2 行表头行 0 列值必须 `numDataFields` 次重复(否则列数不对)
     ```
     行 0: [国家, Product1, Product1, Product1, Product2, Product2, Product2]  // 7 列
     行 1: [国家, 计数, 求和, 多项合并, 计数, 求和, 多项合并]                  // 7 列
     ```
- **J1 公式**:`=k("$$.superPivot", A1:H40, "f3", "f2", "count(),sum('f4'),textjoin('f4',+'+')")` ref `J1:P8`
- **结果**:`2 行 × 7 列` 匹配期望图(德国/美国/英国/中国/test/usa 6 国,每国 3 data fields)

### 10. T57 — superPivot 2 行表头全局开关(2026-06-06)
- **症状**:T56 方案需要 `JSON.stringify({headerRowCount:2})` 写在 WPS 公式里,但 WPS JSA 公式不识别 `JSON.stringify` 和 `{...}` 对象字面量,报 #NAME?
- **修法**:
  1. JSA.k 函数体入口加 hook:`if (JSA.__superPivot2row && /superPivot/i.test(fn)) { ... args.push({headerRowCount: 2}); }`
  2. codetext 末尾初始化 `JSA.__superPivot2row = true;`
  3. J1 公式简化为标准 `k("$$.superPivot", ...)` 不再需要 JSON.stringify
- **优势**:用户公式保持简洁,所有 superPivot 调用自动用 2 行表头;可通过 `JSA.__superPivot2row = false` 关闭

### 10. T56 — superPivot 加 2 行表头支持(2026-06-06)
- **症状**:JSA880 superPivot 单列字段输出 3 行表头(列值/列字段标题/数据名),跟 WPS 原生透视的 2 行表头(列值/数据名+行字段)不一致
- **根因**:`numColFieldLevels === 1` 分支固定 `headerRowCount = 3`,无 option 控制
- **修法**:
  1. JSA880.js `Array2D.z超级透视` line 13341:`options` 接受 string 时自动 `JSON.parse`(因为 k() UDF 公式只能用 string)
  2. JSA880.js `headerRowCount` line 14277:`options.headerRowCount === 2` 时走 2 行
  3. JSA880.js 单列字段表头生成分支 line 14400+:加 `__2row` 分支,行 0=列值,行 1=数据名(行字段标题在两行都 push 便于合并)
  4. `opNameMap.textjoin`: "连接" → "多项合并"(匹配 WPS 原生透视标题)
- **J1 公式**:
  ```
  =k("$$.superPivot", A1:H40, "f3", "f2", "count(),sum('f4'),textjoin('f4',+'+')", 1, 1, "@^@", JSON.stringify({headerRowCount:2}))
  ```
  ref = J1:P8(2 行表头 + 6 国家 = 8 行)
- **结果**:跟期望图一致(2 行表头,Product1/Product2 × 3 data fields,德国/美国/英国/中国/test/usa 全 6 国)

### 11. T67-T72 — 完整 spill 解决方案(2026-06-06 最终版)
- **症状**:T56 superPivot 2 行表头 + 角标题"产品"完成,但 WPS spill 不稳定
- **根因链**:
  1. T70 注入 JSA880.js 源文件时,把 JSA.k 末尾的 T60-T68 改动覆盖了(只替换了 superPivot 函数,没替换 JSA.k 整体)
  2. T70 后 JSA.k 末尾回到 `return result;` → wrappedResult 对象被 toString 当文本
  3. T69 superPivot push options 同样被覆盖
- **修法**:
  1. **T71**:JSA.k 末尾重新加智能 unwrap(1×1→标量,N×1/1×N→1D,N×M→2D)
  2. **T72**:JSA.k 入口重新加 superPivot push `{headerRowCount: 2}` options
- **关键教训**:**任何重新注入 JSA880.js 源文件的操作**都会覆盖 JSA880 codemodule 里所有基于 inject_t27.py 的注入(T60-T69 改动)。每次重新注入后,必须**重新加回**这些改动
- **完整状态(2026-06-06)**:
  - JSA.k 入口 superPivot 自动 push options(T72)
  - JSA.k 末尾智能 unwrap(T71)
  - JSA880.js superPivot 2 行表头 + 角标题"产品"(T56+T70)
  - T49 superPivot 数据字段求值(反引号 + qctextjoin 别名)
  - `fullCalcOnLoad="1"` + J1:P8 范围无阻挡

### 10. T48-T55 — WPS UDF spill 修复系列(2026-06-06)
- **症状**:T49 superPivot 数据字段求值修好后,J1:S14 报 `#SPILL!`(显示"国家",但周围没 spill)
- **第一根因(错的)**:以为是 WPS JSA UDF 返回 2D 数组不能 spill
- **T48** — 直接 return 2D 数组 ❌ WPS 报 #SPILL!
- **T50-T53** — 改"主动写 cell"(caller.Address / 多策略 / caller.Column/Row / Array2D.toRange) — 都失败,WPS UDF 上下文禁止写 caller 之外的 cell
- **T54** — 真正根因发现:**J1:S14 范围被静态 cell 阻挡**(K1=产品, L1=2021, M1=2022...)。WPS 想 spill 但第一行/第一列已有数据,不能覆盖
- **T54b** — 安全版:cell 含 `<f>` 公式就保留(只清 `<v>`),cell 只有 `<v>` 才删整段。T54 误删了 sheet6/sheet9 公式
- **修法**:
  1. **T54b** — 清空所有 spill 范围(J1:S14 等 7 个 sheet × N 个 cell = 226 个静态 cell)里除 anchor 之外的所有静态 cell
  2. **T55** — JSA.k 末尾**回滚**到 T48 行为(直接 return 2D 数组),让 WPS 自己 spill(阻挡已清)
  3. 加 `<calcPr ... fullCalcOnLoad="1"/>` 让 WPS 打开时强制重算

**关键认知**:WPS UDF 是**能**返回 2D 数组并 spill 的,只要 spill 范围无阻挡。不要再去写 cell 绕路。

### 10. ⭐ **WPS `#SPILL!` 根因分析框架(2026-06-06 终极认知)**
| 误诊 | 真因 | 检测方法 |
|------|------|---------|
| "UDF 返回 2D 数组不能 spill" | spill 范围有静态 cell 阻挡 | `python3` 读 xlsm 的 sheet XML,看 spill 范围内除 anchor 外是否还有 cell |
| "JSA.k 写 cell 失败" | **WPS UDF 写 caller 之外的 cell 本来就被禁** | 试 `caller.Cells(2, 1).Value2 = ...` 必然抛错 |
| "Range 传错" | Range 是 callable function 对象(macOS WPS) | `=k("a=>typeof a", A2:B7)` 返回 `"function"` |
| "WPS 公式语法不认" | WPS 倒序加载 codemodule,JSA880 内部 wrapper 覆盖 ThisWorkbook wrapper | 用 `__T39_DIAG__` 探针 + console.log |

**根因排查流程**:
1. 先看 sheet XML 里 spill 范围有没有阻挡 cell(`<c r="X">...` 但 anchor 之外)
2. 再看 JSA.k 末尾是 return 2D 还是 return 标量(T48 应该 return 2D)
3. 最后看 caller.Column/Row 是不是单 cell(不是多 cell 范围)

**绝对不要做的**:
- ❌ 用 regex `<c r="X"...>...</c>` 删 cell — 会把公式 cell 也删了(2026-06-06 真实事故)
- ❌ 让 UDF 写 caller 之外的 cell — WPS 硬限制
- ❌ 用 console.log 调试 UDF — 可能静默

**绝对要做的**:
- ✅ 安全清 cell:先检查 `'<f' in cell_inner`,有公式就保留只清 `<v>`
- ✅ 改 JSA.k 的 codetext 用 ET(ElementTree)而不是 regex
- ✅ 每次破坏性操作前 `cp file.xlsm file.xlsm.tNN.bak` 备份
- ✅ 加 `fullCalcOnLoad="1"` 让 WPS 强制重算
- ✅ JSA.k 检测 superPivot 自动 push `{headerRowCount: 2}` options(2026-06-06 T69)

---

## 🐛 关键 bug 时间线(按发现顺序)

### 1. CDATA 编码问题
- **症状**:wrapper 注入后 WPS 看不到代码
- **根因**:WPS JDEData.bin 用 CDATA 时 `]]>` 字符串会被错误分割
- **修法**:用普通 text 节点 + ElementTree 自动处理 `& < >` 编码(T16)

### 2. JS 函数声明顺序陷阱
- **症状**:wrapper 不生效,WPS 还能调 k
- **根因**:JS 函数声明 hoisting,后定义者赢。wrapper 放末尾时被 JSA880.js 内部 `function k` 遮蔽
- **修法**:wrapper 放最前 + 重命名 JSA880.js 内部 `function k()` → `__JSA_k_internal()`(T20)

### 3. `smartUnwrap` 反逻辑 bug
- **症状**:`v !== asRange(v)` 当 v 已是 Range 时为 false,跳过 `.Value2` 转换
- **根因**:`asRange(v)` 对 Range 返回 v 自身,导致 `v !== asRange(v)` 永远是 false
- **修法**:去掉 `v !== asRange(v)` 检查(T23)

### 4. ⭐ **WPS JSA on macOS 的核心怪异:Range 作为 function 传**
- **症状**:`typeof A2:B7 === "function"`,不是 object 也不是 Array
- **根因**:WPS JSA macOS 版本把 Range 编码为 callable function 对象(同时具有 Range 的所有属性:Address, Value2, Cells, Rows 等)
- **修法**:smartUnwrap 增加 `typeof v === 'function' && v.Value2` 的路径(T29)
- **诊断关键**:`=k("a=>typeof a", A2:B7)` → `"function"`,让人顿悟

### 5. WPS UDF scanner 只扫开头 / 只扫 JSA880
- **症状**:wrapper 在 codemodule 里但 WPS 不调它
- **根因**:WPS 倒序加载 codemodules,且 JSA880 codemodule 的 `function k` 会赢
- **修法**:最终发现 **JSA.k 是 WPS 的真正 UDF**,不是 `function k`!(T39)
- **诊断**:`JSA.k` 加 `__T39_DIAG__` 探针,看到 `T39_JSA_K_CALLED:__T39_DIAG__` 证明 JSA.k 是 UDF

### 6. WPS UDF 上下文禁止修改 calling cell 之外的 cell
- **症状**:JSA.k 里用 `caller.Cells(r+1, c+1).Value2 = vv` 写其他 cell,只有 J1 有值
- **根因**:WPS JSA UDF 安全限制,只能改 calling cell
- **修法**:T48 终极方案 — **直接返回 2D 数组**,让 WPS 自己 spill(不写 cell)

### 7. superPivot 返回 wrappedResult 不是 2D 数组
- **症状**:JSA.k 拿到的是 `wrappedResult` 对象,有 `.val()` `.res()` `.toRange()` 方法
- **根因**:JSA880.js 的 z超级透视 用 IIFE 包装,返回带方法的对象
- **修法**:JSA.k 加 `_kExtract2D()` 尝试 `v.val()` / `v.res()` 等方法

### 8. superPivot 数据字段求值 bug(已修 — T49)
- **症状**:
  - `textjoin(\`f4+'+'+f5\`,\`+\`)` 输出 `"10*5\`+\`30*5"` — 反引号没去掉
  - `qctextjoin(\`f12\`)` 不被识别,被忽略,该列填 0
- **根因**:
  - `parseResultSelector` (JSA880.js:13626) 解析后没去掉 args 里的反引号
  - `executeAggregation` (JSA880.js:14073) switch case 没列 `qctextjoin` 别名
  - `parseAggString` regex 也没列 `qctextjoin`
- **修法**(2026-06-06):
  1. `parseResultSelector` line 13638 后:`op.args.push(argValue.replace(/`/g, ""))`
  2. `parseResultSelector` line 13642 后:`if (op.name === 'qctextjoin') op.name = 'textjoin';`
  3. `executeAggregation` switch:case `'textjoin'` fallthrough 到 `case 'qctextjoin'`
  4. `parseAggString` regex 改:`(sum|count|average|avg|max|min|textjoin|qctextjoin|平方和)`
- **验证**:Node `vm.runInContext` 4 个公式(sheet1/sheet6/sheet8/sheet9)全部输出正确数值

---

## 🛠️ 调试方法论(从这次学到的)

### ✅ 有效方法
1. **从简单到复杂**:先 `=k("JSA.getIndexs", 1, 5, 1)` 测基础(无 Range arg)
2. **逐层加复杂度**:chainable → 多 arg → Array2D 方法
3. **全方位探针**:用 lambda 返回 `typeof`/`Array.isArray`/`Object.getOwnPropertyNames` 摸清 WPS 实际传什么
4. **最简单的 lambda 测路径**:`=k("a=>1+1", A2:B7)` 看 jsaLambda 路径通不通
5. **错误信息反向追溯**:`#K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined"` → 看 JSA.k 的判断逻辑(line 2977)
6. **覆盖式注入**:用 `__T37_DIAG__` 探针式 fn 字符串,throw/return 看到 WPS 到底调了谁
7. **直接修源码,加 _kExtract2D 处理 wrappedResult**:不靠 toRange,直接 2D 数组 + WPS 自动 spill

### ❌ 失败方法
1. **依赖 `Console.log` 在 UDF 上下文** — WPS 在 UDF 上下文里可能静默 console
2. **往 cell 写调试信息** — `Application.ActiveSheet` 在模块加载时未就绪;UDF 上下文不能写其他 cell
3. **加诊断 wrapper 然后用 if 分支** — wrapper 没被调到时分支永远不进
4. **修改多处然后期待"某处生效"** — 改动太多难定位
5. **throw 错误让用户看** — WPS JSA UDF 可能不传播 throw,只显示函数 return 值
6. **依赖 `Application.Caller` 在 UDF 上下文能写 cell** — WPS 限制,只能改 caller 自身

---

## 💡 经验教训(11 条)

### 1. **WPS JSA macOS vs Windows 行为差异巨大**
- Windows:Range 作为 2D array of VARIANTs 传
- **macOS:Range 作为 callable function 对象传**
- 跨平台 JSA 代码必须同时处理两种情况

### 2. **WPS JSA macOS 的 UDF 是 JSA.k,不是 function k**
- 不是 top-level function declaration
- 是 JSA 命名空间下的方法
- 改 `function k` 永远没用,必须改 `JSA.k`

### 3. **WPS UDF 上下文的安全限制**
- UDF 只能改 calling cell (`Application.Caller`)
- 不能写其他 cell — 所有 spill 尝试都失败
- **唯一可行的 spill:返回 2D 数组,让 WPS 自己 spill**

### 4. **JS 函数声明 hoisting + codemodule 加载顺序**
- WPS 倒序加载 codemodules
- 后定义者赢
- ThisWorkbook 的 wrapper 会被 JSA880 内部的同名函数覆盖
- **解决:JSA.k 才是 UDF,改 JSA.k 而非 function k**

### 5. **诊断要从"问 WPS 问什么"开始**
不要假设 WPS 传什么(我以为传 2D array),用 probe 让 WPS 告诉你:
```js
=k("a=>typeof a", A2:B7)  // "function" ← macOS 把 Range 当 function 传
=k("a=>Object.getOwnPropertyNames(a).join(',')", A2:B7)  // Address,Value2,Cells,...
```

### 6. **T27 是关键转折点**
**回滚到能跑的状态(t15.bak)+ 最小改动(只换 JSA880 codemodule 的 codetext)** 是正确的回归策略。之前加 kWrapper 改 ThisWorkbook 都是越改越乱。

### 7. **smartUnwrap 6 路径最终方案**
```
路径 0: typeof === 'function' && v.Value2 → v = v.Value2   (WPS macOS Range-as-function)
路径 1: typeof === 'object' && v.Address && v.Value2 → v = v.Value2  (标准 Range)
路径 2: v.Cells / v.Rows / v.Columns → 手动 .Cells(r,c).Value2  (Range-like)
路径 3: Array.isArray(v) && !v.filter → JSON.parse(JSON.stringify(v))  (host array)
路径 4: !Array.isArray && typeof v.length === 'number' → 遍历转 2D  (类数组)
路径 5: 1D Array → 2D n×1  (压扁的 1D 数组)
路径 6: 1x1 2D → 标量  (WPS 单 cell 引用)
```

### 8. **WPS UDF 自动 spill 机制**
- 1D 数组:`[1,2,3,4,5]` 自动 spill 到 N1:N5
- 2D 数组:必须用 1) 数组公式(Ctrl+Shift+Enter)预选范围 或 2) UDF 直接 return 2D 数组
- WrappedResult 对象:不能 spill,必须 `.val()` 提取

### 9. **JSA.k 加 spill 的正确位置**
```js
JSA.k = function(fn) {
    var result = JSA.jsaLambda.apply(...);
    // T48:wrappedResult 提取 2D 数组,让 WPS 自己 spill
    if (typeof result === "object" && result !== null) {
        var arr = result.val ? result.val() : null;
        if (Array.isArray(arr) && arr[0] && Array.isArray(arr[0])) {
            return arr;  // 直接 return 2D,WPS spill
        }
    }
    return result;
};
```

### 10. **Node 仿真和 WPS 行为差异**
Node 仿真 16/16 通过 ≠ WPS 能用。WPS 有自己的引擎/限制,必须实测。

### 11. **永远测试最简形式**
- `=k("anything")` 测 wrapper 是否被调
- `=k("a=>typeof a", A2:B7)` 测 arg 类型
- `=k("a=>1+1", A2:B7)` 测 lambda 路径
- `=k("__T39_DIAG__", ...)` 测具体 UDF 是哪个

### 12. **WPS #SPILL! 三大根因(2026-06-06 总结)**
1. **Spill 范围被静态 cell 阻挡**(最常见)— 检查 sheet XML 看 spill 范围内除 anchor 外有没有 `<c r="X">...</c>`
2. **UDF 返回 wrappedResult 等对象而非 2D 数组** — 用 `.val()` / `.res()` 提取
3. **WPS 没启用动态数组** — 检查 WPS 选项 → 高级 → 启用动态数组

### 13. **破坏性操作前必须 `cp file.xlsm file.xlsm.tNN.bak`**
- T54 误删公式事故(2026-06-06):regex `<c r="X"...>...</c>` 误删 9 个公式
- 教训:每次写破坏性 Python 脚本,开头加 `if not backup.exists(): shutil.copy2(...)`
- 教训:删 cell 前先 `'<f' in cell_inner` 检查,有公式就保留只清 `<v>`

### 14. **JSA880 codemodule codetext 操作要用 ET 不要用 regex**
- ET 解析 → 修改 → 写回 = 安全,保留所有 XML 属性
- regex 操作 codetext = 容易破坏 CDATA、HTML entity 编码、换行符
- 即使只是改一行代码,先用 ET 找到目标 codemodule,再 ET 修改子元素

### 15. **JSA880.js 源文件注入会覆盖 codemodule 里的所有改动(2026-06-06)**
- inject_t27.py 注入的 T60-T68 改动只在 JSA880 codemodule 里
- 如果用 `ct.text = jsa880_code` 整个替换 codemodule,会把这些改动覆盖
- **解法**:每次替换 codemodule 后,必须**重新加回** inject_t27.py 注入的 T60-T69 改动
- 或:把 T60-T69 改动直接合并到 JSA880.js 源文件里,避免重复

### 16. **WPS JSA Workbook_Open 事件名(2026-06-06)**
- WPS JSA 用 `function Workbook_Open() { ... }`(跟 VBA 一样)
- 不是 `ThisWorkbook_Open()`(那是 VBA 风格,有些 IDE 显示但 WPS 不触发)
- WPS 默认禁宏,事件不会跑。要 spill 必须依靠 WPS 自然 spill(UDF return 2D 数组)

### 17. **JSA.k 末尾 unwrap 逻辑根据 shape 决定(2026-06-06)**
- wrappedResult → 2D 数组
- 1×1 → 标量
- N×1 / 1×N → flatten 成 1D(让 SUM 加总)
- N×M (M>1) → 2D 数组让 WPS spill

### 18. **superPivot 2 行表头角标题从 cornerTitle 改成 colConfig 标题(2026-06-06)**
- 期望:J1 = "产品"(列字段标题),J2 = "国家"(行字段标题)
- 修法:`var __cornerTitle = cornerTitle || (colConfig.fields.length > 0 ? getFieldTitle(colConfig.fields[0], 0, 'col') : title);`
- 3 行表头模式有 `if (cornerTitle && numRowFieldLevels === 1) headerRows[0][0] = cornerTitle;`,2 行模式需要类似逻辑

---

## 📁 关键文件

- **`inject_t27.py`** — 最终注入脚本(T49 数据字段 bug 已修)
- **`/tmp/t54b_safe_clear.py`** — T54b 安全清 spill 阻挡脚本(只清 v 保留 f)
- **`/tmp/add_fullcalc.py`** — 加 fullCalcOnLoad="1" 脚本
- **`/tmp/clear_j1s14.py`** — 误删公式版本(❌ 不要再用)
- **`/tmp/t70_corner_title.py`** — T70 2 行表头角标题修复
- **`/tmp/t71_re_add_unwrap.py`** — T71 重新加 unwrap 逻辑
- **`/tmp/t72_re_add_push.py`** — T72 重新加 superPivot push options
- **`KO一切的k函数.xlsm.t15.bak`** — 注入前的最后干净状态(救命锚点)
- **`KO一切的k函数.xlsm.t48.bak`** — T48 修 2D spill 后的状态
- **`KO一切的k函数.xlsm.t49.bak`** — T49 修 superPivot 数据字段后的状态
- **`KO一切的k函数.xlsm.t50.bak`** — T50 修复前(9 个公式,array mode)
- **`KO一切的k函数.xlsm.t54b.bak`** — T54b 完整修复状态(公式全在 + 阻挡清空 + fullCalcOnLoad=1)
- **`KO一切的k函数.xlsm.t54.bak`** — T54 误删公式坏状态(❌ 不要回滚到此)
- **`KO一切的k函数.xlsm.t72.bak`** — T72 完整 spill + 2 行表头(当前最终)
- **`KO一切的k函数.xlsm.bak`** — 原始文件
- **`KO一切的k函数.xlsm.v4.2.2.bak`** — 中间状态

---

## 🚀 留给未来

如果再遇到 WPS JSA 公式不工作:
1. **不要碰 `function k`** — WPS UDF 是 `JSA.k`
2. **5 个最简探针 cell**(`typeof` / `isArray` / `length` / `keys` / `Value2`)
3. **不要假设 WPS 传什么** — 用 probe 问 WPS
4. **macOS vs Windows 行为差异大** — 优先在测试环境复现
5. **smartUnwrap 一定要有"function 路径"**(WPS macOS 特有)
6. **WPS UDF 上下文不能改其他 cell** — 2D 数组直接 return 让 WPS spill
7. **JSA.k 改 wrapper 逻辑**(而非 function k)— 函数 hoisting + codemodule 倒序加载
8. **遇到 #SPILL! 优先查 sheet XML 看 spill 范围有没有静态 cell 阻挡**(2026-06-06 真因)
9. **破坏性操作前 `cp file.xlsm file.xlsm.tNN.bak`**(2026-06-06 T54 误删公式教训)
10. **JSA880 codemodule codetext 用 ET 改,不用 regex**(2026-06-06)
11. **加 `fullCalcOnLoad="1"`** 让 WPS 打开时强制重算(2026-06-06)
12. **JSA.k 末尾 unwrap 根据 shape 决定**(2026-06-06 T71):1×1 标量,N×1/1×N 1D(SUM),N×M 2D(spill)
13. **superPivot 调用自动 push `{headerRowCount: 2}`**(2026-06-06 T72)
14. **JSA880.js 源文件注入会覆盖 codemodule 里的所有改动**(2026-06-06)— 必须重新加
15. **WPS JSA Workbook_Open 事件名是 `Workbook_Open()`**(2026-06-06)— 不是 `ThisWorkbook_Open()`
16. **2 行表头角标题从 cornerTitle 改成 getFieldTitle(colConfig.fields[0])**(2026-06-06 T70)
