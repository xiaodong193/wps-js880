# k() 修复 TODO

## 现状盘点

xlsm 中有 **10 个 k/jsaLambda 公式** 分布在 6 个工作表：

| Sheet | 单元格 | 公式摘要 |
|-------|--------|---------|
| 1 (Sheet1) | N3 | `jsaLambda("$$.superPivot",A2:L23,{"f2"},{""},{...})` |
| 2 | N3 | `jsaLambda("$$.superPivot",A2:L23,{"f2"},{"f12"},{...})` |
| 5 (test) | E2 | `k("$$.leftjoin",D2:D4,A2:B7,"f1","f1","a.f1,b.f2")` |
| 5 | H2 | `k("(a,v)=>a.filter(x=>x[1]>v)",A2:B7,O1)` |
| 5 | N2 | `SUM(k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1))` |
| 5 | Q2 | `k("arr=>$$.distinct(arr,'f1')",A2:B7)` |
| 5 | T2 | `k("(arr,v)=>arr.map(x=>[x.f1,x.f2+v])",A2:B7,O1)` |
| 6 | F9 | `k("(...args)=>$$.insertCols($$.superPivot(...args),-1,x=>x.sum())",A1:C17,"f1","f2","sum('f3')",1,0)` |
| 6 | F15 | `k("(x,y)=>{y=y+1;return x.map(a=>[a[0],a[2]+'-'+y])}",A2:C17,1)` |
| 8 (多层透视) | J1 | `k("$$.superPivot",A1:H40,"f3,f2","f6","sum(`f4*f5`),textjoin(`f4+'*'+f5`,`+`)")` |
| 9 (test) | J1 | `k("(...args)=>$$.superPivot(...args).filter(...)",A1:H40,"f3,f2","",...)` |

## 问题诊断

| # | 问题 | 现状 | 影响 |
|---|------|------|------|
| P1 | 公式路径前缀 `$$` | JSA880 里有 `$$` 别名指向 `Array2D`，但 `z解析函数表达式` 路径解析时**不识别** `$$.` | 所有路径调用都报 null |
| P2 | WPS 公式 `{"f2"}` 是数组常量 | WPS 把它当 1x1 2D 数组 `[["f2"]]` 传给 JSA；框架期望 1D `["f2"]` | z超级透视 拿不到行字段 |
| P3 | `qctextjoin` 函数 | 框架只有 `textjoin`，无 `qctextjoin` 别名 | sheet1 N3 数据字段算错 |
| P4 | `Range` 对象 vs `Value2` 数组 | WPS 公式里 `A2:L23` 传的是 Range，z超级透视 期望 2D 数组 | 数据拿不到 |
| P5 | UDF 注册 | `JSA.k = JSA.jsaLambda` 是运行时赋值，WPS 公式引擎**不认** | 公式直接 #NAME? |

## 修复方案

### A. 改 `JSA.z解析函数表达式` (js880/JSA880.js)
1. 加 `$$.` 路径分支（root = `$$`，否则 Array2D）
2. 容错：如果 `JSA.xxx` 找不到，尝试在 `Array2D` / 全局找

### B. 改 `JSA.jsaLambda` (js880/JSA880.js)
1. 智能 Range 检测：参数是 Range 对象时自动 `.Value2` 转 2D 数组
   （但**不破坏**链式：先看 z超级透视/z去重 实际要什么）
   实际策略：把 Range 替换成 `Range2D` 包装 `{__range: true, value2: ..., address: ...}`，但这样要改太多。
   更稳妥：jsaLambda 把 Range 转 `.Value2` 一次性消化。
2. WPS 数组常量 → JS 数组：1x1 2D 数组自动 flatten
3. 1xN 2D 数组自动 flatten 为 1D（行字段/列字段/数据字段 场景）

### C. 改 `z超级透视` 数据字段解析 (js880/JSA880.js)
1. 把 `qctextjoin` 识别为 `textjoin` 别名

### D. 改 UDF 模块 (KO一切的k函数_UDF模块.js)
1. UDF 的 `k` 函数不再依赖 `JSA.jsaLambda`（避免 null 静默失败）
2. UDF 内部实现完整逻辑：智能 Range、数组常量、$$ 别名、qctextjoin 等
3. 加 `function _k_self_test()` 在 Workbook_Open 里跑自检

### E. 同步修复到 xlsm 的 JDEData.bin
1. 把 JSA880.js 修复内容同步到 bin 里的 JSA880 module
2. 重新打包 xlsm

## 验收

- node 单测 13 组断言全过
- 10 个公式语义在 node 模拟下能跑通（或返回正确类型）
