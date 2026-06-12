/**
 * KO一切的k函数 · 一键运行 JSA 脚本
 * ────────────────────────────────────────────────────────────
 * 用法：WPS → 开发工具 → JSA 编辑器 → 粘入本文件 → F5 运行
 *       想跑哪个公式就调用对应的 run_xxx()
 *
 * 关键说明（结合 xlsm 实际公式）：
 *   - WPS 公式栏里写 =k(...) / =jsaLambda(...) 报 #NAME?
 *     原因：WPS 公式引擎不识别 JSA 全局函数 k()
 *   - 解决：把"k(...)"翻译成等价的"new Array2D(...).方法(...)"
 *     在 JSA 编辑器里直接跑，结果 spill 到目标单元格
 * ────────────────────────────────────────────────────────────
 * 框架代码（JSA880.js）已在 xlsm 的 JDEData.bin 里内置，直接可用。
 * 如果提示找不到 JSA / Array2D / RngUtils，请先：
 *   1. WPS → 选项 → 加载项 → 选中 JSA880 加载项
 *   2. 或在 JSA 编辑器第一行加：loadScript("JSA880.js")
 */

// ╔════════════════════════════════════════════════════════════════╗
// ║  工具函数：通用结果输出                                          ║
// ╚════════════════════════════════════════════════════════════════╝

/** 把二维数组 spill 到目标单元格 */
function writeTo(rng, arr) {
    var target = typeof rng === "string" ? Range(rng) : rng;
    return Array2D.toRange(arr, target);
}

/** 读取 A1:H40 这种范围到 JS 数组 */
function readRange(addr) {
    return Range(addr).Value2;
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  对应 sheet5 (test) 第 2 行的 6 个 k() 公式                      ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 对应公式 E2: k("$$.leftjoin",D2:D4,A2:B7,"f1","f1","a.f1,b.f2")
 * 左连接 D2:D4 和 A2:B7，按 f1 关联，输出 a.f1 + b.f2
 */
function run_e2() {
    var left  = Range("D2:D4").Value2;   // [["x"], ["x"], ...]
    var right = Range("A2:B7").Value2;   // [["1","a"], ["2","b"], ...]
    var result = new Array2D(left).z左连接(
        right, "f1", "f1", "a.f1,b.f2"
    );
    writeTo("E2", result.val());
    Console.log("E2 ✅ 写出 " + result.length + " 行到 E2");
}

/**
 * 对应公式 H2: k("(a,v)=>a.filter(x=>x[1]>v)",A2:B7,O1)
 * 把 A2:B7 中"第2列 > O1 单元格值"的行筛出来
 * O1 单元格存放阈值（看 sheet5 是数字 3）
 */
function run_h2() {
    var arr = Range("A2:B7").Value2;
    var threshold = Range("O1").Value2;
    var result = new Array2D(arr).z筛选(
        "(x) => x[1] > " + threshold
    );
    writeTo("H2", result.val());
    Console.log("H2 ✅ 筛出 " + result.length + " 行（> " + threshold + "）");
}

/**
 * 对应公式 K2: testfilter(A2:B7)
 * 这是一个 testfilter 自定义函数（用户期望调用）
 * 我们用一个普通 JSA 函数实现：筛选 A 列 > 阈值
 */
function testfilter(rng) {
    var arr = typeof rng === "string" ? Range(rng).Value2 : rng.Value2;
    return new Array2D(arr).z筛选("x => x[0] > 1").val();
}
function run_k2() {
    var result = testfilter("A2:B7");
    writeTo("K2", result);
    Console.log("K2 ✅ testfilter 跑通");
}

/**
 * 对应公式 N2: SUM(k("(arr,v)=>arr.map(x=>[x.f2+v])",A2:B7,O1))
 * 把 A2:B7 每行第 2 列 + O1 单元格值，然后求和
 */
function run_n2() {
    var arr = Range("A2:B7").Value2;
    var v = Range("O1").Value2;
    var mapped = new Array2D(arr).z映射(
        "x => [x.f2 + v]".replace("v", v)
    );
    // 把映射结果的第 1 列求和
    var sum = new Array2D(mapped).z求和("f1");
    writeTo("N2", [[sum]]);
    Console.log("N2 ✅ 求和结果 = " + sum);
}

/**
 * 对应公式 Q2: k("arr=>$$.distinct(arr,'f1')",A2:B7)
 * 按第 1 列去重
 */
function run_q2() {
    var arr = Range("A2:B7").Value2;
    var result = new Array2D(arr).z去重("f1");
    writeTo("Q2", result.val());
    Console.log("Q2 ✅ 去重后 " + result.length + " 行");
}

/**
 * 对应公式 T2: k("(arr,v)=>arr.map(x=>[x.f1,x.f2+v])",A2:B7,O1)
 * 把 A2:B7 每行变成 [x[0], x[1]+v]
 */
function run_t2() {
    var arr = Range("A2:B7").Value2;
    var v = Range("O1").Value2;
    var result = new Array2D(arr).z映射(
        "x => [x.f1, x.f2 + v]".replace("v", v)
    );
    writeTo("T2", result.val());
    Console.log("T2 ✅ 映射结果 " + result.length + " 行");
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  对应 sheet8 (多层透视) J1 公式                                  ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 对应公式 J1: k("$$.superPivot",A1:H40,"f3,f2","f6","sum(`f4*f5`),textjoin(`f4+'*'+f5`,`+`)")
 * 数据 A1:H40：第 1 行表头
 * 行字段: f3, f2  (产品, 国家)
 * 列字段: f6     (年)
 * 数据字段: sum(f4*f5) = 数量×价格总和
 *           textjoin(f4+'*'+f5, '+') = 把 "数量*价格" 拼起来
 *
 * 数据范围：A1:H40 → 第 1 行表头，2-40 行是数据
 */
function run_j1_superPivot() {
    // 1. 读数据（含表头）
    var data = Range("多层透视!A1:H40").Value2;
    if (!Array.isArray(data) || data.length < 2) {
        Console.log("多层透视!A1:H40 是空表");
        return;
    }

    // 2. 调用 superPivot
    //    - arr  = 数据（含表头）
    //    - rowFields = ['f3,f2']  行字段
    //    - colFields = ['f6']     列字段
    //    - dataFields = 'sum("f4*f5"),textjoin("f4+'+'+f5","+")'  字符串形式
    //                  ↑ 注意：JSA 里字符串可以用 '+'，不需要 WPS 公式转义
    //    - headerRows = 1  1 行表头
    var result = Array2D.z超级透视(
        data,
        ["f3,f2"],     // 行字段
        ["f6"],        // 列字段
        ['sum("f4*f5"),textjoin("f4+\'*\'+f5","+")'],  // 数据字段
        1
    );

    // 3. 写到 J1（spill 到 J1:S40）
    var rng = Range("多层透视!J1");
    var resultArr = result.val();
    if (resultArr && resultArr.length > 0) {
        Array2D.toRange(resultArr, rng);
        Console.log("J1 ✅ 透视完成，写出 " + resultArr.length + " 行");
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  对应 sheet9 (test) J1 公式 — 你截图里那个"还是不好使"的公式      ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 对应公式 J1: k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')",A1:H40,"f3,f2","","count(),sum(`f4`),textjoin(`f4`,`+`)")
 *
 * 拆解：
 *   - 先调 superPivot(A1:H40, ['f3,f2'], [''], ['count(),sum(f4),textjoin(f4,+)'])
 *   - 然后 filter 掉表头(i!=0) 和非 Product1 的行
 *   - 数据字段里你可以自由写 `f4` `f5` `f4+'*'+f5` 等等
 */
function run_j1_filter() {
    var data = Range("test!A1:H40").Value2;
    if (!Array.isArray(data) || data.length < 2) {
        Console.log("test!A1:H40 是空表");
        return;
    }

    // Step 1: 先 superPivot
    var pivot = Array2D.z超级透视(
        data,
        ["f3,f2"],  // 行字段
        [""],        // 无列字段
        ['count(),sum("f4"),textjoin("f4","+")'],  // 数据字段
        1
    );

    // Step 2: 过滤（保留表头 + f2 == "Product1"）
    var arr2d = pivot;  // Array2D 实例
    var filtered = arr2d.z筛选("(x, i) => i == 0 || x.f2 == 'Product1'");

    // Step 3: 写到 test!J1
    var rng = Range("test!J1");
    if (filtered.length > 0) {
        Array2D.toRange(filtered.val(), rng);
        Console.log("J1 ✅ Product1 筛选完成，写出 " + filtered.length + " 行");
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  对应 sheet1 (Sheet1) N3 公式                                   ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 对应公式 N3: jsaLambda("$$.superPivot",A2:L23,{"f2"},{""},{"sum(`f5`),sum(`f7`),sum(`f9`),sum(`f10`),qctextjoin(`f12`)"...},...)
 *
 * 你的 jsaLambda 里 qctextjoin 看起来是个新函数，我们用标准 textjoin 替代
 * 数据范围：A2:L23（注意 A2 起，表头在 row 1）
 */
function run_n3_sheet1() {
    var data = Range("Sheet1!A2:L23").Value2;
    if (!Array.isArray(data) || data.length < 1) {
        Console.log("Sheet1!A2:L23 是空表");
        return;
    }
    // 取出第一行作为表头（因为 A2 起，所以要从 A1 读表头）
    var header = Range("Sheet1!A1:L1").Value2;
    if (Array.isArray(header) && Array.isArray(header[0])) {
        data = [].concat([header[0]], data);
    }

    var result = Array2D.z超级透视(
        data,
        ["f2"],     // 行字段
        [""],        // 无列字段
        ['sum("f5"),sum("f7"),sum("f9"),sum("f10"),textjoin("f12","+")'],
        1
    );

    var rng = Range("Sheet1!N3");
    if (result.length > 0) {
        Array2D.toRange(result.val(), rng);
        Console.log("N3 ✅ 写出 " + result.length + " 行透视结果");
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║  一键跑全部（演示用）                                          ║
// ╚════════════════════════════════════════════════════════════════╝

function run_all() {
    Console.log("═══════ 开始跑全部 8 个公式 ═══════");
    try { run_e2();          } catch (e) { Console.log("E2 ❌ " + e.message); }
    try { run_h2();          } catch (e) { Console.log("H2 ❌ " + e.message); }
    try { run_k2();          } catch (e) { Console.log("K2 ❌ " + e.message); }
    try { run_n2();          } catch (e) { Console.log("N2 ❌ " + e.message); }
    try { run_q2();          } catch (e) { Console.log("Q2 ❌ " + e.message); }
    try { run_t2();          } catch (e) { Console.log("T2 ❌ " + e.message); }
    try { run_j1_superPivot(); } catch (e) { Console.log("J1(多层透视) ❌ " + e.message); }
    try { run_j1_filter();     } catch (e) { Console.log("J1(test) ❌ " + e.message); }
    try { run_n3_sheet1();    } catch (e) { Console.log("N3(Sheet1) ❌ " + e.message); }
    Console.log("═══════ 完成 ═══════");
}
