/**
 * ⚠️ [DEPRECATED] 此文件已废弃 (2026-06-05)
 *
 * k() 实现已合并到 JSA880.js 的 JSA.k 中(v5.0)
 * 请改用: ThisWorkbook 里 3-5 行 wrapper:
 *   function k(fn) { return JSA.k.apply(null, arguments); }
 *   function jsaLambda(fn) { return JSA.k.apply(null, arguments); }
 *
 * 此文件保留作为历史参考。不要在新工作簿中粘入。
 */

/**
 * ═══════════════════════════════════════════════════════════════════════
 *  KO_k_udf.js · WPS 公式 UDF 兜底层(独立、零依赖)
 * ═══════════════════════════════════════════════════════════════════════
 *
 *  目的:让 WPS 公式栏里直接可用 =k(...) / =jsaLambda(...)
 *       即使主 JSA880 框架加载失败,这个 shim 也能保证 UDF 被注册。
 *
 *  设计原则:
 *    • 全部用 ES5 语法(function / var / arguments),不依赖 const/let、...args、new Function
 *    • 不在顶层执行任何可能有兼容性问题的代码(连 typeof 判断都包在 try-catch 里)
 *    • 依赖懒加载:只在用户实际调用 =k(...) 时才尝试找 JSA.jsaLambda
 *    • 失败时优雅返回 "#K_ERR: ..." 而不是抛错(抛错会显示 #VALUE!)
 *
 *  注入方式:
 *    python3 inject_jsa880_main.py --target KO一切的k函数.xlsm \
 *        --source KO_k_udf.js --module-name KO_k_udf --purge-prefix KO_k_udf
 *
 * ═══════════════════════════════════════════════════════════════════════
 */

// ╔════════════════════════════════════════════════════════════════╗
// ║  [v4.2.2 bug 修复] $$ 路径别名 = Array2D                      ║
// ║  原 z解析函数表达式(line 2967):                                  ║
// ║    root = (typeof $$ !== 'undefined') ? $$ : null;                ║
// ║  框架从未把 $$ 定义为全局对象,导致 typeof $$ 永远 undefined,   ║
// ║  $$.superPivot / $$.getIndexs / $$.z超级透视 全部失败,           ║
// ║  公式返 0。我们在兜底层顶层把 $$ 设到 globalThis 上,             ║
// ║  修复后所有 $$.xxx 公式都能用。                                  ║
// ╚════════════════════════════════════════════════════════════════╝
try {
    (function() {
        var __g = (new Function('return this'))();
        if (typeof __g.$$ === 'undefined' && typeof Array2D !== 'undefined') {
            __g.$$ = Array2D;
        }
    })();
} catch (e) { /* 静默失败不致命 */ }

// ╔════════════════════════════════════════════════════════════════╗
// ║  WPS 公式 UDF 顶层 function(最关键,必须是顶层声明!)           ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * k 函数 — WPS 公式 UDF
 * 单元格示例:
 *   =k("JSA.getIndexs", 1, 10, 2)
 *   =k("JSA.z超级透视", A1:H40, "f3,f2", "f6", "sum(`f4*f5`)", "textjoin(`f4+'+'+f5`,`+`)")
 *
 * 用 arguments 替代 ...args,保证在 WPS 15990 之前的老版本也能加载
 */
function k(fn) {
    try {
        // 收集所有参数(从第 2 个开始,fn 是第 1 个)
        var args = [];
        for (var i = 1; i < arguments.length; i++) {
            args.push(arguments[i]);
        }

        // 1) 优先调 JSA.jsaLambda(主框架 v4.2.2 里的实现,功能最全)
        if (typeof JSA !== 'undefined' && typeof JSA.jsaLambda === 'function') {
            var __result = JSA.jsaLambda.apply(null, [fn].concat(args));
            if (__result === undefined || __result === null) {
                // 返回明确错误,避免 WPS 把 undefined 当 0 显示
                return "#K_ERR: jsaLambda 返回 null/undefined。"
                     + "可能是 [1] $$. 路径仍不识别(检查 JSA Console 有无 [JSA880 v4.2.2] 自检),"
                     + "[2] superPivot 参数格式不对,"
                     + "[3] 公式里用了反引号 \`foo\` 但 WPS 公式栏可能不识别,试试改用双引号 \"foo\"";
            }
            return __result;
        }

        // 2) 次选 Array2D.z (框架的核心 namespace)
        if (typeof Array2D !== 'undefined' && typeof Array2D.z !== 'undefined') {
            return Array2D.z(fn, args);
        }

        // 3) 都没有:优雅报错(不抛错,避免显示 #VALUE!)
        return "#K_ERR: JSA880 框架未加载(k UDF 兜底层提示)";

    } catch (e) {
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}

/**
 * jsaLambda 函数 — k() 的全名版本
 * 单元格公式: =jsaLambda("JSA.getIndexs", 1, 5)
 */
function jsaLambda(fn) {
    return k.apply(null, arguments);
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  Workbook_Open:模块加载时打印确认信息                          ║
// ╚════════════════════════════════════════════════════════════════╝

function Workbook_Open() {
    try {
        if (typeof Console === 'undefined') return;
        Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        Console.log("✅ k() UDF 兜底层已加载(KO_k_udf · ES5 兼容版)");
        Console.log("   在任意单元格输入 =k(\"JSA.getIndexs\", 1, 10, 2)");
        Console.log("   看到 1 3 5 7 9 = 成功");
        Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

        // 自检(必须包在 try-catch 里,避免影响 Workbook_Open 整体执行)
        try {
            var test = k("JSA.getIndexs", 1, 5, 1);
            if (test && test.length) {
                Console.log("   自检:k('JSA.getIndexs', 1, 5, 1) = [" + test.join(",") + "]");
            } else {
                Console.log("   ⚠️ 自检:k() 返回了空或非数组(框架可能没加载)");
            }
        } catch (e) {
            Console.log("   ⚠️ 自检失败:" + e.message);
        }
    } catch (e) {
        // Workbook_Open 抛错不会影响 UDF 注册
    }
}
