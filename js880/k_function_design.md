# 增强型 k 函数设计文档

## 目标
支持在 WPS 单元格公式中执行复杂的链式操作和 Lambda 表达式，并自动溢出数组结果。

## 用例
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```

## 核心改进点

### 1. 全局 $$ 对象定义
- 创建 `$$` 作为 `Array2D` 的快捷别名
- 在全局作用域中注册，供 Lambda 表达式访问
- 支持所有 Array2D 的方法链

```javascript
// 在 JSA 初始化时
if (typeof $$ === 'undefined') {
    globalThis.$$ = Array2D;
    if (typeof window !== 'undefined') window.$$ = Array2D;
}
```

### 2. 链式方法调用支持
- 修改 Lambda 解析器，允许方法链（如 `.filter()`, `.map()`, `.sort()` 等）
- 返回值应保持为数组或二维数组对象，以便继续链式调用

```javascript
// 当前不支持的表达式
"(...args)=>$$.superPivot(...args).filter(...)"

// 需要改进的解析流程
1. 识别 $$ 对象
2. 调用 superPivot() 方法
3. 将结果转换为可链式调用的对象
4. 执行 .filter() 等后续方法
```

### 3. Array2D 增强的 Fluent API
为 Array2D 的返回值添加以下方法，实现流式编程：

| 方法 | 功能 | 示例 |
|------|------|------|
| `.filter(fn)` | 按条件筛选行 | `.filter((x,i)=>i==0\|\|x.f2=='Product1')` |
| `.map(fn)` | 映射/转换行 | `.map(x=>({...x, f1: x.f1*2}))` |
| `.slice(start, end)` | 提取行范围 | `.slice(1, 10)` |
| `.take(n)` | 取前 N 行 | `.take(5)` |
| `.skip(n)` | 跳过前 N 行 | `.skip(1)` |
| `.sort(fn)` | 排序 | `.sort((a,b)=>a.f1-b.f1)` |
| `.forEach(fn)` | 遍历执行 | `.forEach(x=>console.log(x))` |

### 4. 改进的 jsaLambda 解析流程

```javascript
// 伪代码流程
JSA.jsaLambda = function(fn, ...args) {
    // 1. 确保 $$ 全局对象存在
    if (typeof $$ === 'undefined') {
        globalThis.$$ = Array2D;
    }
    
    // 2. 解析 Lambda 表达式
    var expr = JSA.z解析函数表达式(fn);
    
    // 3. 执行表达式
    var result = expr.apply(null, args);
    
    // 4. 如果结果是二维数组，确保返回值能正确溢出
    if (Array.isArray(result)) {
        return result;  // WPS 15990+ 自动处理数组溢出
    }
    
    return result;
}
```

### 5. Lambda 表达式的语法支持

| 表达式类型 | 示例 |
|----------|------|
| 箭头函数（简单） | `x => x * 2` |
| 箭头函数（多参） | `(x, i) => i === 0 \|\| x.f2 === 'Product1'` |
| 多行箭头函数 | `(...args) => { var x = $$.superPivot(...args); return x.filter(...); }` |
| 路径调用 | `$$.superPivot` / `JSA.getIndexs` / `Array2D.z筛选` |
| 链式调用 | `$$.superPivot(...args).filter(...).map(...)` |
| 特殊变量 | `$0, $1, ...` (参数索引) |
| 列选择器 | `f1, f2, f3` (二维数组列) |

## 实现细节

### 核心修改位置

#### 1. JSA880.js - 全局初始化（第 2700 行附近）
```javascript
// 添加 $$ 全局对象映射
(function initGlobalAliases() {
    if (typeof $$ === 'undefined') {
        if (typeof globalThis !== 'undefined') {
            globalThis.$$ = Array2D;
        }
        if (typeof window !== 'undefined') {
            window.$$ = Array2D;
        }
    }
})();
```

#### 2. z解析函数表达式 中的链式方法支持（第 2964 行）
```javascript
// 在箭头函数解析后，添加链式调用检测
// 检查是否包含 .filter( / .map( 等方法
if (expr.indexOf('.filter(') > -1 || 
    expr.indexOf('.map(') > -1 || 
    expr.indexOf('.sort(') > -1) {
    // 需要特殊处理：确保 $$ 被正确绑定到作用域
    try {
        var fnWithScope = new Function('$$', 'return ' + expr);
        _lambdaCache[expr] = function() {
            return fnWithScope.apply(null, [Array2D].concat([].slice.call(arguments)));
        };
        return _lambdaCache[expr];
    } catch (e) {
        console.warn('链式调用解析失败:', expr, e);
    }
}
```

#### 3. k 函数的 UDF 包装（第 18820 行）
```javascript
function k(fn, ...args) {
    try {
        // 确保 $$ 全局对象存在
        if (typeof $$ === 'undefined') {
            globalThis.$$ = Array2D;
        }
        return JSA.jsaLambda(fn, ...args);
    } catch (e) {
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}
```

## 测试用例

### 测试 1: 基础链式调用
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0)",A1:H40,"f3,f2","","count(),sum(`f4`)")
```
期望：返回第一行（标题）

### 测试 2: 条件筛选
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`)")
```
期望：返回所有 Product1 的透视结果

### 测试 3: 链式 map + filter
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i<5).map(x=>({...x,f1:x.f1*2}))",A1:H40,"f3,f2","","count()")
```
期望：返回前 5 行，并将第一列值翻倍

### 测试 4: 复合条件
```
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
```
期望：返回标题行 + 所有 Product1 记录

## 性能考虑

1. **缓存策略**：已有 `_lambdaCache`，链式表达式也应缓存
2. **内存溢出**：大数据透视结果可能很大，需测试 WPS 数组溢出限制
3. **循环优化**：避免重复遍历数组（如已 filter，不要再 map 同样的条件）

## 向后兼容性

- 现有的 k 函数调用保持不变
- 新增的链式方法是可选的，旧公式不受影响
- $$ 别名不与现有命名冲突

## 参考资源

- 现有 jsaLambda 实现：JSA880.js:2801-2874
- Lambda 表达式解析：JSA880.js:2964-3100+
- Array2D 方法列表：JSA880.js 中 Array2D 对象定义
- WPS UDF 数组溢出：WPS 15990+
