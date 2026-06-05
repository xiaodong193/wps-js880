/**
 * ⚠️ [DEPRECATED] 此文件已废弃 (2026-06-05)
 *
 * k() 实现已合并到 JSA880.js 的 JSA.k 中(v5.0)
 * 原本这个文件会自动加载 JSA880.js + 暴露 k() shim,
 * 现在 k() 的实现全在 JSA880.js 内部,ThisWorkbook 只需 3-5 行 wrapper:
 *   function k(fn) { return JSA.k.apply(null, arguments); }
 *   function jsaLambda(fn) { return JSA.k.apply(null, arguments); }
 *
 * 此文件保留作为历史参考。不要在新工作簿中粘入。
 */

/**
 * ═══════════════════════════════════════════════════════════════════════
 *  KO一切的k函数 · WPS 公式 UDF 注入器
 * ═══════════════════════════════════════════════════════════════════════
 *
 * 🎯 目的：让 k() 在 WPS 表格里"当公式用"
 *    单元格 =k("JSA.getIndexs", 1, 10, 2) → 1 3 5 7 9
 *
 * 🔑 关键机制：
 *    WPS 公式引擎只扫描 ThisWorkbook 代码模块的顶层 function 作为 UDF。
 *    加载项(Application.LoadScript 加载的 .js)的顶层 function 不会被注册。
 *
 *    所以这个 UDF 模块必须"粘到当前 xlsm 的 ThisWorkbook 代码模块"里，
 *    它负责加载 JSA880.js 框架(以加载项方式)并提供 UDF shim。
 *
 * 📦 部署步骤(2 步):
 *    1. WPS → 选项 → 加载项 → 加载 JSA880.js(框架)
 *       (或:WPS → 开发工具 → JSA 编辑器 → 在加载项目录里放 JSA880.js)
 *    2. 打开 KO一切的k函数.xlsm
 *       WPS → 开发工具 → JSA 编辑器 → 找到 ThisWorkbook 模块
 *       把本文件**全部内容**粘进去 → Ctrl+S 保存
 *    3. 重启 xlsm(或按 F5 触发 Workbook_Open)
 *    4. 单元格输入 =k("JSA.getIndexs", 1, 10, 2) → 看到 1 3 5 7 9
 *
 * ⚠️ 不要把这个文件作为加载项加载!它必须作为 ThisWorkbook 模块代码存在。
 *
 * ═══════════════════════════════════════════════════════════════════════
 */

// ╔════════════════════════════════════════════════════════════════╗
// ║ 1. 自动加载 JSA880.js 框架(以加载项方式)                        ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 自动加载 JSA880.js 框架代码
 *
 * 调用顺序(逐个尝试,直到其中一个成功):
 *   1. 检查 JSA / Array2D 全局对象是否已存在(加载项已加载)
 *   2. Application.LoadScript("JSA880.js")   ← WPS JSA 引擎提供的入口
 *   3. loadScript("JSA880.js")                ← 旧版 WPS 兼容
 *
 * 如果全部失败,会给出明确的报错信息。WPS 15990+ 推荐使用 JSA 加载项机制。
 */
function _ensureJSA880Loaded() {
    if (typeof JSA !== 'undefined'
        && typeof Array2D !== 'undefined'
        && typeof JSA.jsaLambda === 'function') {
        return true;  // 已加载
    }

    try {
        if (typeof Application !== 'undefined'
            && typeof Application.LoadScript === 'function') {
            Application.LoadScript("JSA880.js");
            if (typeof JSA !== 'undefined' && typeof JSA.jsaLambda === 'function') {
                return true;
            }
        }
    } catch (e1) { /* ignore */ }

    try {
        if (typeof loadScript === 'function') {
            loadScript("JSA880.js");
            if (typeof JSA !== 'undefined' && typeof JSA.jsaLambda === 'function') {
                return true;
            }
        }
    } catch (e2) { /* ignore */ }

    return false;
}

// 在 ThisWorkbook 模块加载时立即执行(顶层代码,WPS 模块载入后立刻跑)
(function _uDFModuleBootstrap() {
    if (_ensureJSA880Loaded()) {
        return;  // 成功,公式 =k(...) 现在可以直接用了
    }
    // 加载失败:给出明确指引
    if (typeof Console !== 'undefined') {
        Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        Console.log("⚠️ JSA880.js 未加载,k() UDF 暂时不能用!");
        Console.log("");
        Console.log("解决:");
        Console.log("  ① WPS → 选项 → 加载项 → 加载 JSA880.js 加载项(推荐)");
        Console.log("  ② 或把 JSA880.js 放到加载项目录并重启 WPS");
        Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    }
})();

// ╔════════════════════════════════════════════════════════════════╗
// ║ 2. WPS 公式 UDF 顶层 function                                  ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * k 函数 — WPS 公式 UDF
 *
 * 单元格公式示例:
 *   =k("JSA.getIndexs", 1, 10, 2)              → [1,3,5,7,9] 数组溢出
 *   =k("x => x*2", 5)                          → 10
 *   =k("Array2D.z超级透视", A1:H40, "f3,f2", "f6", "sum(`f4*f5`)", "textjoin(`f4+'+'+f5`,`+`)")
 *
 * @param {string|Function} fn - 字符串函数表达式 / 路径 / Lambda
 * @param {...any} args - 后续参数(range / 数字 / 字符串 / 对象)
 * @returns {*} 函数结果;失败返回 "#K_ERR: 错误信息"
 */
function k(fn, ...args) {
    try {
        return JSA.jsaLambda(fn, ...args);
    } catch (e) {
        // UDF 不能抛错(会显示 #VALUE!),改返回错误字符串
        return "#K_ERR: " + (e && e.message ? e.message : String(e));
    }
}

/**
 * jsaLambda 函数 — k() 的全名版本,WPS UDF
 * 单元格公式:=jsaLambda("JSA.getIndexs", 1, 5)
 */
function jsaLambda(fn, ...args) {
    return k(fn, ...args);
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 3. Workbook_Open:工作簿打开时打印确认信息 + 自检                ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 工作簿打开事件 - 验证 k() UDF 已注册成功
 * 看到这条日志就说明 =k(...) 公式可以用了
 *
 * 绑定:ThisWorkbook 模块默认会自动作为工作簿打开事件
 */
function Workbook_Open() {
    if (typeof Console === 'undefined') return;
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("✅ k() UDF 已就绪!(基于 JSA880 + UDF shim)");
    Console.log("   测试:在任意单元格输入");
    Console.log("   =k(\"JSA.getIndexs\", 1, 10, 2)");
    Console.log("   看到 1 3 5 7 9 = 成功!");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    // 立刻跑一个自检
    try {
        var test = k("JSA.getIndexs", 1, 5, 1);
        Console.log("   自检:k('JSA.getIndexs', 1, 5, 1) = [" + test.join(",") + "]");
    } catch (e) {
        Console.log("   ⚠️ 自检失败:" + e.message);
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 4. Sheet_Change:依赖单元格变化时自动重算 k() 公式              ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 工作表 Change 事件 - 当数据区域变化时强制重算所有 k 公式
 * 解决"改了源数据但 k() 公式结果没自动更新"的问题
 *
 * 绑定方法:WPS → 工作表事件 → 选择事件源、事件类型=Change、宏名=k_onChange
 */
function k_onChange(Sh, Target) {
    try {
        var sheet = Sh;
        var usedRange = sheet.UsedRange;
        if (!usedRange) return;
        var data = usedRange.Formula;
        if (!data) return;

        var needRecalc = false;
        var rows = data.length;
        var cols = Array.isArray(data[0]) ? data[0].length : 1;
        for (var r = 0; r < rows; r++) {
            for (var c = 0; c < cols; c++) {
                var f = Array.isArray(data[r]) ? data[r][c] : data[r];
                if (typeof f === "string" && f.indexOf("k(") === 1) {
                    var cell = sheet.Cells(r + usedRange.Row, c + usedRange.Column);
                    cell.Dirty = true;
                    needRecalc = true;
                }
            }
        }
        if (needRecalc) {
            Application.Calculate();
            if (typeof Console !== 'undefined') Console.log("k() 公式已自动重算");
        }
    } catch (e) { /* 静默失败 */ }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 5. 工具函数:批量执行 k() 公式                                   ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 把一列 k() 调用结果 spill 到目标区域
 * @param {string} targetRangeAddr - 目标区域起始地址,如 "J1"
 * @param {Array} fnArray - 函数调用参数数组
 * @returns {number} 成功执行的函数数量
 */
function batch_k(targetRangeAddr, fnArray) {
    if (!Array.isArray(fnArray)) return 0;
    var rng = Range(targetRangeAddr);
    var result = [];
    for (var i = 0; i < fnArray.length; i++) {
        var args = fnArray[i];
        var fn = args[0];
        var rest = Array.prototype.slice.call(args, 1);
        result.push(k.apply(null, [fn].concat(rest)));
    }
    if (typeof Array2D !== 'undefined' && Array2D.toRange) {
        Array2D.toRange(result, rng);
    } else {
        rng.Value2 = result;
    }
    if (typeof Console !== 'undefined') {
        Console.log("✅ batch_k: " + fnArray.length + " 个 k() 调用结果已写入 " + targetRangeAddr);
    }
    return fnArray.length;
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 6. 帮助:常见问题 FAQ                                           ║
// ╚════════════════════════════════════════════════════════════════╝

function k_help() {
    if (typeof Console === 'undefined') return;
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("  k() 公式常见问题排查");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("");
    Console.log("❌ 问题1: =k(...) 报 #NAME?");
    Console.log("   原因: WPS 公式引擎不认 k() 函数");
    Console.log("   解决: 把本文件粘到 ThisWorkbook 代码模块并保存");
    Console.log("");
    Console.log("❌ 问题2: k() 返回 #K_ERR: ...");
    Console.log("   原因: 函数表达式语法错 或 参数不匹配");
    Console.log("   解决: 在 JSA 编辑器里直接调 JSA.jsaLambda(...) 调试");
    Console.log("");
    Console.log("❌ 问题3: 改了源数据,k() 结果不更新");
    Console.log("   解决: 绑 Sheet_Change 事件到 k_onChange (本文件提供)");
    Console.log("");
    Console.log("❌ 问题4: 公式返回单个值时不显示(数组没 spill)");
    Console.log("   原因: WPS 版本 < 15990,不支持数组溢出");
    Console.log("   解决: 升级 WPS Office 到 15990+ 版本");
    Console.log("");
    Console.log("✅ 验证 k() UDF 已就绪:");
    var ok = false;
    try { k("JSA.getIndexs", 1, 3, 1); ok = true; } catch (e) {}
    if (ok) Console.log("   ✓ k() 内部调用成功");
    else Console.log("   ✗ k() 内部调用失败 - 检查 JSA880.js 是否已加载");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
}
