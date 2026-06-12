# JSA880.js 修改应用指南

## 快速总结
您需要的增强 k 函数现已设计完成。核心改进允许在 WPS 单元格中使用链式方法调用，如：
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```

## 应用步骤

### 步骤 1: 初始化 $$ 全局别名
**文件**: `JSA880.js`  
**位置**: 约 440 行（在 LAMBDA_PATTERNS 定义前）  
**代码**:
```javascript
(function initGlobalAliases() {
    if (typeof globalThis !== 'undefined') {
        if (typeof globalThis.$$ === 'undefined') {
            globalThis.$$ = undefined;
        }
    }
    if (typeof window !== 'undefined' && typeof window.$$ === 'undefined') {
        window.$$ = undefined;
    }
})();
```

### 步骤 2: 添加链式表达式解析器
**文件**: `JSA880.js`  
**位置**: 约 2850 行（在 JSA.jsaLambda 函数定义前）  
**代码**: 见 `enhanced_k_function.js` 中的 `parseChainableExpression()` 函数

### 步骤 3: 在 jsaLambda 中集成链式调用检测
**文件**: `JSA880.js`  
**位置**: 约 2857 行（在 `var func = JSA.z解析函数表达式(fn);` 之前）  
**代码**:
```javascript
// 【v4.2.4 增强】支持链式方法调用
if (typeof fn === 'string' && /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce)\s*\(/.test(fn)) {
    var chainParser = parseChainableExpression(fn);
    if (chainParser) {
        try {
            return chainParser.apply(null, [Array2D].concat(realArgs));
        } catch (e) {
            console.warn('链式调用执行失败:', e.message);
        }
    }
}
```

### 步骤 4: 增强 k() UDF 函数
**文件**: `JSA880.js`  
**位置**: 约 18820 行（替换 `function k(fn, ...args)` 的实现）  
**代码**:
```javascript
function k(fn, ...args) {
    try {
        // 确保 $$ 全局对象在 UDF 执行上下文中可用
        if (typeof globalThis !== 'undefined' && Array2D) {
            globalThis.$$ = Array2D;
        }
        if (typeof window !== 'undefined' && Array2D) {
            window.$$ = Array2D;
        }
        return JSA.jsaLambda(fn, ...args);
    } catch (e) {
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}
```

### 步骤 5: 为 Array2D 添加流式方法（可选）
**文件**: `JSA880.js`  
**位置**: 约 14700 行（Array2D 定义完成后）  
**代码**: 见 `enhanced_k_function.js` 中的流式方法包装

### 步骤 6: 完成全局别名初始化
**文件**: `JSA880.js`  
**位置**: 最后一行（约 18900+ 行）  
**代码**:
```javascript
function finalizeGlobalAliases() {
    if (typeof Array2D !== 'undefined') {
        if (typeof globalThis !== 'undefined') {
            globalThis.$$ = Array2D;
        }
        if (typeof window !== 'undefined') {
            window.$$ = Array2D;
        }
    }
}

// 调用完成函数
finalizeGlobalAliases();
```

## 工作原理

### 链式调用解析流程
1. **检测**: 在 jsaLambda 中检测表达式是否包含 `.filter(`, `.map(` 等方法链
2. **创建作用域**: 使用 `Function` 构造器创建一个函数，其中 `$$` 作为参数
3. **执行**: 调用创建的函数，将 Array2D 作为 `$$` 参数传入
4. **返回**: 链式操作的结果直接返回到单元格，WPS 自动处理数组溢出

### 示例流程
```javascript
// 用户输入
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")

// 解析步骤
1. 检测到 .filter( → 启用链式模式
2. 创建函数: new Function('$$', 'return (...args)=>$$.superPivot(...args).filter(...)')
3. 执行: 函数.call(null, Array2D, A1:H40, "f3,f2", "", ...)
4. superPivot 返回透视表
5. filter 筛选行
6. 结果数组回传到单元格，WPS 15990+ 自动溢出
```

## 支持的方法链

| 方法 | 用途 | 示例 |
|------|------|------|
| `.filter(predicate)` | 按条件筛选 | `.filter((x,i)=>i==0 \|\| x.f2=='Product1')` |
| `.map(mapper)` | 转换/映射 | `.map(x=>({...x, f1: x.f1*2}))` |
| `.slice(start, end)` | 提取范围 | `.slice(1, 10)` |
| `.take(n)` | 取前 N 行 | `.take(5)` |
| `.skip(n)` | 跳过 N 行 | `.skip(1)` |
| `.sort(comparator)` | 排序 | `.sort((a,b)=>a.f1-b.f1)` |
| `.forEach(fn)` | 遍历执行 | `.forEach(x=>console.log(x))` |
| `.reduce(fn, init)` | 汇总 | `.reduce((sum,x)=>sum+x.f1, 0)` |

## 实际应用示例

### 例子 1: 透视表 + 产品筛选
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",
   A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```
结果：只显示标题行 + Product1 的透视统计

### 例子 2: 透视表 + 前 5 行 + 字段倍增
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i<5).map(x=>({...x,f1:x.f1*2}))",
   A1:H40,"f3,f2","","sum(`f4`)")
```
结果：前 5 行的透视数据，f1 列值翻倍

### 例子 3: 简单数据数组筛选
```
=k("arr=>arr.filter((x,i)=>i>0 && x[1]>100).map(x=>[x[0].toUpperCase(), x[1]*2])",
   -r, A1:B100)
```
结果：跳过标题行，筛选第2列>100，并转换数据

### 例子 4: 文本数据处理链
```
=k("data=>data.filter(x=>x.indexOf('Error')>-1).map(x=>x.toUpperCase()).slice(0,10)",
   A1:A1000)
```
结果：包含"Error"的文本，转大写，取前10行

## 向后兼容性

所有现有 k 函数公式继续有效：
- ✅ `=k("JSA.getIndexs", 1, 10, 2)`
- ✅ `=k("x => x*2", 5)`
- ✅ `=k("$0+$1", [10, 20])`
- ✅ `=k("Array2D.z超级透视", A1:H40, "f3,f2", "", "count(),sum(...)")`

新增功能仅当检测到方法链时激活，不影响现有公式。

## 性能考虑

1. **缓存**: 链式表达式会被缓存，重复调用避免重新解析
2. **内存**: 大型数据透视结果可能超过 WPS 溢出限制（通常 10-20 万行）
3. **优化建议**: 
   - 先用 `.filter()` 缩小数据量
   - 然后再用 `.map()` 转换
   - 避免重复遍历

## 故障排查

| 症状 | 原因 | 解决方案 |
|------|------|---------|
| `#K_ERR` 错误 | Lambda 表达式语法错误 | 检查箭头函数、括号匹配 |
| 返回 `#NAME?` | $$ 未定义 | 确保 finalizeGlobalAliases() 被调用 |
| 链式调用无效 | 方法不支持 | 只支持数组原生方法，不支持自定义方法 |
| 数组未溢出 | WPS 版本太旧 | 需要 WPS 15990 或更新版本 |
| 结果不符预期 | 筛选条件错误 | 在单元格 / 开发者工具中逐步测试 |

## 验证清单

部署后请检查：
- [ ] `$$ === Array2D` 在开发者工具中返回 true
- [ ] 基础 k 函数仍然正常工作（无链式）
- [ ] 链式表达式能正确解析和执行
- [ ] 单元格中数组能正确溢出
- [ ] 没有性能下降（表达式缓存有效）
