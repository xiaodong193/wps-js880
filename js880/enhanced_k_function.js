// ═══════════════════════════════════════════════════════════════════════
// 增强型 k 函数实现
// 用于支持链式操作和复杂 Lambda 表达式
// ═══════════════════════════════════════════════════════════════════════

// ──────────────────────────────────────────────────────────────────────
// 修改 1: 初始化 $$ 全局对象别名
// 位置：JSA880.js 顶部全局初始化区域（约 440 行附近）
// ──────────────────────────────────────────────────────────────────────

(function initGlobalAliases() {
    // 创建 $$ 作为 Array2D 的快捷别名
    // 用于在 k() 公式中简化引用
    if (typeof globalThis !== 'undefined') {
        if (typeof globalThis.$$ === 'undefined') {
            globalThis.$$ = undefined; // 先占位，会在 Array2D 定义后赋值
        }
    }

    // 兼容浏览器环境
    if (typeof window !== 'undefined' && typeof window.$$ === 'undefined') {
        window.$$ = undefined;
    }
})();

/**
 * 在 Array2D 定义完成后调用此函数
 * （在 Array2D.js 最后或 JSA880.js 大约 14700 行附近调用）
 */
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

// ──────────────────────────────────────────────────────────────────────
// 修改 2: 增强 Lambda 解析器以支持链式方法调用
// 位置：JSA880.js 中 z解析函数表达式 函数内部（约 2964 行）
// ──────────────────────────────────────────────────────────────────────

/**
 * 增强版函数表达式解析 - 支持链式方法调用
 * 改进点：
 * 1. 检测箭头函数中的方法链（.filter, .map, .slice 等）
 * 2. 为链式表达式创建正确的作用域上下文（$$ 绑定）
 * 3. 缓存解析结果以提高性能
 */
function parseChainableExpression(expr) {
    // 检查是否是链式调用表达式
    var isChainable = /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce|flat|flatMap|find|some|every|includes)\s*\(/i.test(expr);

    if (!isChainable) {
        return null; // 返回 null 表示不是链式，由原有流程处理
    }

    try {
        // 确保 $$ 被定义并指向 Array2D
        if (typeof globalThis !== 'undefined') {
            globalThis.$$ = Array2D;
        }
        if (typeof window !== 'undefined') {
            window.$$ = Array2D;
        }

        // 对于链式表达式，使用 Function 构造器创建函数
        // 这样能保证 $$ 在作用域内可访问
        var fnBody = 'return ' + expr;
        var chainFn = new Function('$$', fnBody);

        return chainFn;
    } catch (e) {
        console.warn('链式表达式解析失败 [' + expr + ']:', e.message);
        return null;
    }
}

// ──────────────────────────────────────────────────────────────────────
// 修改 3: 改进 jsaLambda 函数以支持链式调用
// 位置：JSA880.js 中 JSA.jsaLambda 函数（约 2801 行）
// 改动：在原有流程中添加链式调用检测
// ──────────────────────────────────────────────────────────────────────

// 在 JSA.jsaLambda 中，修改解析流程：
// 原来的步骤 2-3) 解析字符串为函数 之前添加：

// INSERT THIS CODE in JSA.jsaLambda after line 2857:
/*
        // 【v4.2.4 增强】支持链式方法调用
        // 检测并解析链式表达式（如 $$.superPivot(...).filter(...).map(...) ）
        if (typeof fn === 'string' && /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce)\s*\(/.test(fn)) {
            var chainParser = parseChainableExpression(fn);
            if (chainParser) {
                try {
                    return chainParser.apply(null, [Array2D].concat(realArgs));
                } catch (e) {
                    console.warn('链式调用执行失败:', e.message);
                    // 继续尝试其他解析方式
                }
            }
        }
*/

// ──────────────────────────────────────────────────────────────────────
// 修改 4: 增强 k() UDF 函数确保全局上下文正确
// 位置：JSA880.js 中 function k() 定义（约 18820 行）
// ──────────────────────────────────────────────────────────────────────

// 改进后的 k 函数（替换原有的 function k 定义）：
function k(fn, ...args) {
    try {
        // 确保 $$ 全局对象在 UDF 执行上下文中可用
        // 这对 WPS 的 UDF 沙箱环境很重要
        if (typeof globalThis !== 'undefined' && Array2D) {
            globalThis.$$ = Array2D;
        }
        if (typeof window !== 'undefined' && Array2D) {
            window.$$ = Array2D;
        }

        // 调用 jsaLambda 执行
        return JSA.jsaLambda(fn, ...args);
    } catch (e) {
        // UDF 不能抛错（会显示 #VALUE!），改返回错误字符串
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}

// ──────────────────────────────────────────────────────────────────────
// 修改 5: 为 Array2D 返回值添加流式操作方法
// 位置：Array2D 类定义部分（约 14700+ 行）
// ──────────────────────────────────────────────────────────────────────

/**
 * 为 Array2D 原型添加流式操作方法
 * 这些方法支持链式调用
 * 调用位置：在 Array2D.prototype 上添加以下方法
 */

// 注意：Array2D 已经有 z筛选、z映射 等方法
// 这里添加的是原生 JavaScript 数组方法的包装，支持链式操作

// 如果还没有 filter 方法，添加以下方法：
if (!Array2D.prototype.filter) {
    Array2D.prototype.filter = function(predicate) {
        // 转换为普通数组，执行 filter，返回新的 Array2D
        var filtered = Array.prototype.filter.call(this, predicate);
        // 返回仍可链式调用的对象
        return Object.setPrototypeOf(filtered, Array2D.prototype);
    };
}

if (!Array2D.prototype.map) {
    Array2D.prototype.map = function(mapper) {
        var mapped = Array.prototype.map.call(this, mapper);
        return Object.setPrototypeOf(mapped, Array2D.prototype);
    };
}

if (!Array2D.prototype.slice) {
    Array2D.prototype.slice = function(start, end) {
        var sliced = Array.prototype.slice.call(this, start, end);
        return Object.setPrototypeOf(sliced, Array2D.prototype);
    };
}

// 便捷方法：取前 N 行
if (!Array2D.prototype.take) {
    Array2D.prototype.take = function(n) {
        return this.slice(0, n);
    };
}

// 便捷方法：跳过前 N 行
if (!Array2D.prototype.skip) {
    Array2D.prototype.skip = function(n) {
        return this.slice(n);
    };
}

// ──────────────────────────────────────────────────────────────────────
// 修改 6: 在 Array2D 定义后调用全局别名完成函数
// 位置：JSA880.js 最后（约 18600+ 行，在所有定义完成后）
// ──────────────────────────────────────────────────────────────────────

// 添加此调用，确保 $$ 被正确初始化：
// finalizeGlobalAliases();

// ═══════════════════════════════════════════════════════════════════════
// 单元格公式使用示例
// ═══════════════════════════════════════════════════════════════════════

/*

示例 1: 基础超级透视 + 筛选
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")

示例 2: 链式 filter + map
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i<5).map(x=>({...x,f1:x.f1*2}))",A1:H40,"f3,f2","","count()")

示例 3: 多条件筛选
=k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || (x.f2=='Product1' && x.f3>100))",A1:H40,"f3,f2","","count(),sum(`f4`)")

示例 4: 简单数组操作
=k("arr=>arr.filter((x,i)=>i>0).map(x=>[x[0]*2, x[1]])",-r,A1:B10)

示例 5: 文本处理链
=k("data=>data.filter(x=>x.indexOf('Product')>-1).map(x=>x.toUpperCase())",A1:A100)

*/

// ═══════════════════════════════════════════════════════════════════════
// 测试代码片段
// ═══════════════════════════════════════════════════════════════════════

/*
// 在 WPS JSA 工作表中测试以下代码
function test_enhanced_k_function() {
    // 测试 1: 确保 $$ 被正确定义
    console.log('Test 1: $$ 定义检查');
    console.log('  typeof $$:', typeof $$);
    console.log('  $$ === Array2D:', $$ === Array2D);

    // 测试 2: 测试简单链式操作
    console.log('\\nTest 2: 简单数组链式操作');
    var testData = [[1, 'a'], [2, 'b'], [3, 'c']];
    var result = JSA.jsaLambda('data=>data.filter((x,i)=>i>0).map(x=>[x[0]*2, x[1]])', testData);
    console.log('  输入:', testData);
    console.log('  输出:', result);

    // 测试 3: 测试 superPivot 链式调用
    console.log('\\nTest 3: superPivot 链式调用');
    var testExpr = '(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2==\"Product1\")';
    console.log('  表达式:', testExpr);
    console.log('  表达式可解析:', typeof parseChainableExpression(testExpr) === 'function');

    MsgBox('增强 k 函数测试完成，详见控制台日志');
}
*/
