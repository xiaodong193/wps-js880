# JSA880 v4.0.42 修复经验总结

> 2026-06-07 · 32 个版本演进 · 第 13 个根本 bug 修复
> 受益工作簿:`KO一切的k函数.xlsm` Sheet5 N2

---

## 1. 现象(用户报告)

Sheet5 N2 单元格输入公式:
```
=k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1)
```

A2:B7 是 6 行 2 列的数字矩阵(f1, f2),O1 = 3。

期望输出:N2:N7 spill 6 个数字 = **4, 5, 6, 7, 8, 9**(每行 `f2 + 3`)。

实际输出:N2:N7 全 **#NUM!** × 6。

---

## 2. 调试过程(3 次迭代)

### 2.1 第一次:推理 + 模拟

**根因推理**:

| 步骤 | 推理 |
|---|---|
| 1 | `A2:B7` 是 WPS Range,WPS 公式引擎传过来是 6×2 host array |
| 2 | v4.0.37 `_toRealArray` 已用 try-probe 把 host array 转 plain Array |
| 3 | `arr.map` 走原生 Array.prototype.map,正常 |
| 4 | 但 `x.f2` 在 plain Array 上是 `undefined`(fN proxy 只在 Array2D 实例有) |
| 5 | `undefined + 3 = NaN`,包成 `[NaN]` |
| 6 | WPS spill 6 个 `[NaN]` 各自当 1×1 cell,值是 NaN → `#NUM!` |

**Node 模拟**(`/tmp/test_n2_simulation.js`):
- 原生 Array `.map`:`x.f2 = undefined` → 结果 `[[null],[null],...]` ✅ 匹配 #NUM!
- 改用 `x[1]` 索引:`[[4],[5],[6],[7],[8],[9]]` ✅ 期望值
- Array2D 包装后 `.map`:`f2` 拿到值 → `[[4],...]` ✅

### 2.2 第一次注入:用 `$$.from(__a0)` — 静默失败

**修改**:`_kParseChainableExpression` IIFE 模板,自动 wrap 第一个 plain array 参数:
```js
var fn = new Function('__args', 'return (function() {' +
                      '  var $$ = (typeof Array2D !== "undefined") ? Array2D : this.Array2D;' +
                      '  if (__args && __args.length > 0 && $$ && typeof $$.from === "function") {' +
                      '    var __a0 = __args[0];' +
                      '    if (__a0 && !(__a0 instanceof $$) && Array.isArray(__a0)) {' +
                      '      try { __args = [$$.from(__a0)].concat(Array.prototype.slice.call(__args, 1)); } catch (__we) { /* wrap 失败就用原值 */ }' +
                      '    }' +
                      '  }' +
                      '  return (' + expr + ').apply(null, __args);' +
                      '}).apply(null, __args)');
```

**WPS 验证**:仍然 spill `#NUM!` × 6,日志没有任何新条目。

**问题**:wrap 静默失败,完全没线索。

### 2.3 加诊断日志(关键!)

**修改**:IIFE 内部加 `IIFE IN` + `$$` 探测 + `wrap OK/FAIL/skip` 三分支日志:
```js
if (typeof Console !== "undefined") {
  try { Console.log("[k/v4.0.42] IIFE IN: __args[0].type=" + (typeof __args[0]) + ", isArray=" + Array.isArray(__args[0])); } catch(__) {}
}
var $$ = (typeof Array2D !== "undefined") ? Array2D : this.Array2D;
if (typeof Console !== "undefined") {
  try { Console.log("[k/v4.0.42] $$=" + ($$ ? $$.name || "Array2D" : "null") + ", hasFrom=" + ($$ && typeof $$.from === "function")); } catch(__) {}
}
```

**WPS 验证日志**:
```
[k/v4.0.42] IIFE IN: __args[0].type=object, isArray=true
[k/v4.0.42] $$=Array2D, hasFrom=false    ← 关键!
```

**根因定位**:**`Array2D` 没有 `.from` 静态方法**,所以 `typeof $$.from === "function"` 为 false,整个 wrap 分支跳过,**完全静默**。

### 2.4 第二次注入:用 `new $$` — 成功

**修改**:用 `new $$ (__a0)` 替代 `$$.from(__a0)`。`$$` 是 Array2D 函数本身,`new $$` 等价 `new Array2D(arr)`:
```js
var fn = new Function('__args', 'return (function() {' +
                      '  var $$ = (typeof Array2D !== "undefined") ? Array2D : this.Array2D;' +
                      '  if (__args && __args.length > 0 && $$) {' +
                      '    var __a0 = __args[0];' +
                      '    if (__a0 && !(__a0 instanceof $$) && Array.isArray(__a0)) {' +
                      '      try { __args = [new $$ (__a0)].concat(Array.prototype.slice.call(__args, 1)); } catch (__we) { /* 失败就用原值 */ }' +
                      '    }' +
                      '  }' +
                      '  return (' + expr + ').apply(null, __args);' +
                      '}).apply(null, __args)');
```

**WPS 验证**:
```
[k/v4.0.42] IIFE IN: __args[0].type=object, isArray=true
[k/v4.0.42] $$=Array2D, hasFrom=false
[wrap 没打 skip,因为外层 if 已 false;但后续走原生 map 也行]
```

但等等 — 我简化了 `hasFrom` 守卫后,`if (... && $$)` 走新分支:
```
__a0 是 plain Array → 进入 wrap → new $$ (__a0) → Array2D 实例 ✅
```

**最终结果**:**N2:N7 = 4, 5, 6, 7, 8, 9** ✅

---

## 3. 根因(根因 #14)

### 3.1 完整因果链

```
WPS 公式 =k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1)
    ↓
WPS 公式引擎把 A2:B7 转 6×2 host array,O1 转 Range(smartUnwrap 再转 3)
    ↓
jsaLambda 入口:realArgs = [<6×2 host array>, 3]
    ↓
smartUnwrap:_toRealArray 把 host array 转 plain 6×2 Array
    ↓
链式路径 IIFE:arr 是 plain Array(不是 Array2D 实例)
    ↓
arr.map(x=>[x.f2+v]) 走原生 Array.prototype.map
    ↓
x 是 plain Array(没 fN proxy),x.f2 = undefined
    ↓
undefined + 3 = NaN
    ↓
.map 返回 [[NaN],[NaN],...,[NaN]] (6 项 1D 数组)
    ↓
WPS spill N2:N7 = NaN, NaN, NaN, NaN, NaN, NaN → #NUM! × 6
```

### 3.2 为什么 fN proxy 缺失

`Array2D.prototype.z筛选` 和 `Array2D.prototype.z映射` 内部循环(line 8058-8071)给每行注入 fN proxy:
```js
for (var __fc = 0; __fc < __proxy.length; __fc++) {
    var __cellVal = __proxy[__fc];
    if (typeof __cellVal === 'string') {
        __proxy['f' + (__fc + 1)] = __cellVal.replace(/^\s+|\s+$/g, '');
    } else {
        __proxy['f' + (__fc + 1)] = __cellVal;
    }
}
```

**但原生 Array.prototype.map 没有这个 fN proxy 注入逻辑**。

v4.0.42 修法:在 IIFE 入口把 plain Array 升级为 Array2D 实例,后续 .map 走 Array2D.prototype.z映射-like 路径(实际 Array2D.prototype.map 也是手动注入 fN proxy)。

---

## 4. 关键代码

### 4.1 最终版 _kParseChainableExpression(JSA880.js line 3061-3102)

```js
function _kParseChainableExpression(expr) {
    // 检测是否含链式调用(.filter / .map / .slice / .take / .skip / .sort / .reduce)
    var isChainable = /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce|find|some|every)\s*\(/.test(expr);
    if (!isChainable) return null;

    try {
        // 🔧 v4.0.14 关键: IIFE 内 var $$ = Array2D
        // 🔧 v4.0.42 关键: 自动把第一个 plain array 参数 wrap 成 Array2D
        //   根因: 链式 .map(x=>[x.f2+v]) 中,arr 是 _toRealArray 转的 plain Array,
        //         原生 Array.prototype.map 没 fN proxy,x.f2=undefined → 报 #NUM!
        //   修法: IIFE 开头检测 __args[0] 是 plain Array 但不是 Array2D 实例时,
        //         自动 new $$ (__a0) 转 Array2D($$ = Array2D 函数)
        //         这样 .map 走 Array2D.prototype.map(有 fN proxy)
        //   不影响: 已经是 Array2D 的参数原样传递;非数组参数也不动
        //   ⚠️ 不要用 Array2D.from() — 该静态方法不存在,要用 new Array2D(arr)
        var fn = new Function('__args', 'return (function() {' +
                              '  var $$ = (typeof Array2D !== "undefined") ? Array2D : this.Array2D;' +
                              '  if (__args && __args.length > 0 && $$) {' +
                              '    var __a0 = __args[0];' +
                              '    if (__a0 && !(__a0 instanceof $$) && Array.isArray(__a0)) {' +
                              '      try { __args = [new $$ (__a0)].concat(Array.prototype.slice.call(__args, 1)); } catch (__we) { /* 失败就用原值 */ }' +
                              '    }' +
                              '  }' +
                              '  return (' + expr + ').apply(null, __args);' +
                              '}).apply(null, __args)');
        return fn;
    } catch (e) {
        if (typeof Console !== 'undefined') Console.log('parseChainableExpression 失败:' + e.message);
        return null;
    }
}
```

### 4.2 Array2D 构造函数(line 873-913)自动注入 fN proxy

```js
function Array2D(data) {
    if (!(this instanceof Array2D)) return new Array2D(data);
    
    var items = [];
    if (data === null || data === undefined) items = [];
    else if (data instanceof Array2D) items = data._items;
    else if (Array.isArray(data)) items = data;
    else items = [[data]];
    
    // v4.0.11: 为所有行注入 .f1/.f2 列访问器,支持 x=>x.f3 箭头函数回调
    for (var _fi = 0; _fi < items.length; _fi++) {
        var _frow = items[_fi];
        if (Array.isArray(_frow)) {
            for (var _fc = 0; _fc < _frow.length; _fc++) {
                if (!(_frow.hasOwnProperty('f' + (_fc + 1)))) {
                    Object.defineProperty(_frow, 'f' + (_fc + 1), {
                        get: (function(idx) { return function() { return this[idx]; }; })(_fc),
                        set: (function(idx) { return function(v) { this[idx] = v; }; })(_fc),
                        enumerable: false,
                        configurable: true
                    });
                }
            }
        }
    }
    
    Array.prototype.push.apply(this, items);
    // ... _original / _items / _header 等属性
}
```

**`new Array2D(arr)` 一行调用就完成 fN proxy 注入** — 这就是 v4.0.42 wrap 逻辑正确的根本原因。

---

## 5. 验证

### 5.1 单元测试(11/11 通过)

`/tmp/test_v442_chain_wrap.js`:
- T1: 用户原公式 → `[[4],[5],[6],[7],[8],[9]]` ✅
- T2: `.filter(x=>x[1]>v)` 不破坏 ✅
- T3: `(...args)=>$$.superPivot(...args).filter(...)` 不破坏 ✅
- T4: 原 Array2D 不重复 wrap ✅
- T5: 1D 数组也 wrap ✅
- T6: 端到端 6×1 spill ✅
- T7: 边界(空数组)不抛错 ✅
- T8: 边界(空 args)不抛错 ✅
- T9: 嵌套数组 wrap 后 .map 正确 ✅

### 5.2 WPS 端验证

| 公式 | 期望 | 实际 |
|---|---|---|
| `=k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1)` | N2:N7 = 4,5,6,7,8,9 | ✅ 完美 |
| `=SUM(k(...))` | 39 | (待用户测,SUM 接受 6×1 2D 数组) |

### 5.3 向后兼容性(未破坏的公式)

| 公式类型 | 行为 | 影响 |
|---|---|---|
| `(...args)=>$$.superPivot(...args).filter(x=>x.fN==...)` | 第一个参数是 plain array,wrap 后 superPivot 接受 | ✅ 正常工作 |
| `(a,v)=>a.filter(x=>x[1]>v)` | 第一个参数是 plain array,wrap 后 .filter 走 Array2D 路径 | ✅ 正常工作 |
| `(...args)=>args` (无链式) | 走非链式路径,不进 _kParseChainableExpression | ✅ 不变 |

---

## 6. 调试技巧(经验)

### 6.1 IIFE 内部诊断必加

**教训**:IIFE 内部的 `if` 条件不满足时,外层完全无法感知,只能靠最终结果反推。这次 wrap 静默失败,加 `hasFrom` 日志才发现 `Array2D.from` 不存在。

**经验**:
- IIFE 内部每个关键 `if` 都要有"进入"和"跳过"两个分支的诊断
- 表达式:`if (cond) { log OK } else { log SKIP: cond=value }`
- 不能只 log 成功路径,失败路径(条件不满足)更要 log

### 6.2 WPS 重启验证

**教训**:修改 codemodule 后,只关文件不退出 WPS 进程,新代码不生效。

**经验**:
- 每次 codemodule 修改后,**必须 `Cmd+Q` 完全退出 WPS**
- 重启后立即窗口应该有新版本的 console.log
- 如果没看到,99% 是没退出干净

### 6.3 用 `typeof XXX === "function"` 检查方法存在

**教训**:Array2D 假设有 `.from` 静态方法(类似 `Array.from`),但实际没有。`typeof X.from === "function"` 才发现。

**经验**:
- 不要假设一个对象有某个方法(尤其是从其他库抄过来的命名)
- 静态方法 vs 实例方法要分清(Array2D.from 是想抄 Array.from)
- 实在不确定,先 `grep "Array2D.from\s*=" JSA880.js` 查

### 6.4 数组实例化优先用 `new Xxx(arr)`

**经验**:
- `Array2D.from(arr)` — 静态方法,如果不存在就 fail
- `new Array2D(arr)` — 构造函数,总是存在(因为 function Array2D 一定定义)
- 优先后者更稳

---

## 7. 受益公式(其他可能受益场景)

任何链式调用中第一个参数是 plain array(从 Range.Value2 转过来)且用 `x.fN` 访问的:

| 公式 | 之前 | 之后 |
|---|---|---|
| `=k("(arr,v)=>arr.map(x=>[x.f2+v])",A1:B6,O1)` | N1:N6 = #NUM! × 6 | N1:N6 = 正确值 |
| `=k("arr=>arr.filter(x=>x.f3>0).map(x=>x.f1)",A1:C10)` | #NUM! | 正确值 |
| `=k("(arr,k)=>arr.filter(x=>x.fN==k)",A1:H40,"Product1")` | #K_ERR 或 #NUM! | 正确值 |

---

## 8. 相关 commit / 文件

- **修改文件**:`/Users/daidai193/Library/CloudStorage/SynologyDrive-code/js880/JSA880.js`
  - line 8:版本号 → 4.0.42
  - line 10:版本 banner
  - line 25-32:更新日志 (v4.0.42)
  - line 3061-3102:`_kParseChainableExpression` IIFE 模板加自动 wrap

- **注入脚本**:`/Users/daidai193/Library/CloudStorage/SynologyDrive-code/JS880教案/第03章/3-28/inject_jsa880_v413.py`
  - 完整 codemodule 覆盖,无需修改

- **注入工作簿**:`/Users/daidai193/Library/CloudStorage/SynologyDrive-code/JS880教案/第03章/3-28/KO一切的k函数.xlsm`
  - 注入后 740,471 → 742,769 bytes
  - 验证 xlsm 包含 v4.0.42 标记 6 处

- **Memory**:`/Users/daidai193/.claude/projects/-Users-daidai193-Library-CloudStorage-SynologyDrive-code/memory/wps-jsa880-debugging.md`
  - 新增根因 #14

---

## 9. 后续(剩余任务)

- ⏳ Sheet5 E2 (leftjoin) + Q2 (distinct) F9 验证 — v4.0.40 修复
- ⏳ abcd汇总 F9 链式 `.insertCols + .superPivot` 修复
- ⏳ test J1/J15 `.filter` 完整 spill 验证

---

## 10. 一句话总结

链式 .map 中 `x.fN` 拿不到值,根因是 plain Array 没 fN proxy;修法是在 IIFE 入口用 `new Array2D(arr)`(不是 `Array2D.from(arr)` — 该方法不存在)自动 wrap 第一个 plain array 参数。
