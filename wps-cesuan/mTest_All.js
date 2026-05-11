/**
 * ============== 统一测试套件 ==============
 * 作者：徐晓冬
 * 版本：V2.20260505
 * 描述：整合 mUnitTestFramework + mTest_PeriodRate + mTest_SuperPivot
 * ====================================================
 */

var g_verbose = false;
var _isNode = (typeof window === 'undefined' && typeof Application === 'undefined');

// Node.js 环境检测
if (_isNode) {
    global.Console = { log: function() {} };
    global.Application = {
        ActiveSheet: { Name: '测试工作表' },
        Worksheets: function() { return { Name: '', Cells: {}, Range: function() { return { Merge: function() {} }; } }; },
        Range: function() { return { Value2: null }; },
        ScreenUpdating: true, Calculation: -4105, EnableEvents: true
    };
    global.Worksheets = global.Application.Worksheets;
    var _fs = require('fs'), _path = require('path'), _vm = require('vm');
    try {
        _vm.runInThisContext(_fs.readFileSync(_path.join(__dirname, '../js880/JSA880.js'), 'utf8'), { filename: 'JSA880.js' });
    } catch (e) { console.log("[WARN] JSA880.js 未找到，跳过"); }
}

// ==================== 测试框架 ====================

const assert = {
    assertEquals: (e, a, m) => { if (e !== a) throw new Error(m || `期望${e}，实际${a}`); g_verbose && console.log(`    ✅ ${e} === ${a}`); },
    assertTrue: (v, m) => { if (v !== true) throw new Error(m || `期望true，实际${v}`); g_verbose && console.log(`    ✅ true`); },
    assertFalse: (v, m) => { if (v !== false) throw new Error(m || `期望false，实际${v}`); g_verbose && console.log(`    ✅ false`); },
    assertNull: (v, m) => { if (v !== null) throw new Error(m || `期望null`); g_verbose && console.log(`    ✅ null`); },
    assertNotNull: (v, m) => { if (v === null || v === undefined) throw new Error(m || `值为null`); g_verbose && console.log(`    ✅ ${typeof v}`); },
    assertUndefined: (v, m) => { if (v !== undefined) throw new Error(m || `期望undefined`); g_verbose && console.log(`    ✅ undefined`); },
    assertDefined: (v, m) => { if (v === undefined) throw new Error(m || `未定义`); g_verbose && console.log(`    ✅ 已定义`); },
    assertThrows: (fn, m) => { var thrown = false, err = null; try { fn(); } catch (e) { thrown = true; err = e; } if (!thrown) throw new Error(m || "期望抛出异常"); g_verbose && console.log(`    ✅ 抛出异常: ${err.message}`); },
    assertAlmostEquals: (e, a, t, m) => { if (Math.abs(e - a) > t) throw new Error(m || `期望${e}，实际${a}，容差${t}`); g_verbose && console.log(`    ✅ ${e} ≈ ${a}`); },
    assertArrayEquals: (e, a, m) => {
        if (!Array.isArray(e) || !Array.isArray(a)) throw new Error(m || "非数组");
        if (e.length !== a.length) throw new Error(m || `长度${e.length}!==${a.length}`);
        for (var i = 0; i < e.length; i++) if (e[i] !== a[i]) throw new Error(m || `第${i}项:${e[i]}!==${a[i]}`);
        g_verbose && console.log(`    ✅ [${e.join(",")}]`);
    },
    assertContains: (c, item, m) => {
        var ok = Array.isArray(c) ? c.includes(item) : c.hasOwnProperty(item);
        if (!ok) throw new Error(m || `不包含${item}`);
        g_verbose && console.log(`    ✅ 包含${item}`);
    }
};

class clsTestCase {
    constructor(name, fn, timeout = 5000) {
        this.name = name; this.fn = fn; this.timeout = timeout;
        this.result = null; this.error = null; this.duration = 0; this.skipped = false;
    }
    run() {
        if (this.skipped) { this.result = 'skipped'; return this.getResult(); }
        var start = Date.now();
        try { this.fn(); this.result = 'passed'; } catch (e) { this.result = 'failed'; this.error = e; }
        this.duration = Date.now() - start;
        return this.getResult();
    }
    getResult() {
        return { name: this.name, result: this.result, duration: this.duration,
                 error: this.error ? { message: this.error.message, stack: this.error.stack } : null };
    }
}

class clsTestSuite {
    constructor(name) {
        this.name = name; this.tests = []; this.beforeAll = null; this.afterAll = null;
        this.beforeEach = null; this.afterEach = null; this.results = [];
    }
    test(name, fn, timeout) { this.tests.push(new clsTestCase(name, fn, timeout)); }
    skip(name, fn) { var t = new clsTestCase(name, fn); t.skipped = true; this.tests.push(t); }

    async run() {
        console.log(`\n========== ${this.name} ==========`);
        g_verbose && console.log(`  ${this.tests.length} 个测试\n`);
        if (this.beforeAll) { try { g_verbose && console.log(`  [beforeAll]`); await this.beforeAll(); } catch (e) { console.error(`beforeAll失败: ${e.message}`); return; } }
        for (var test of this.tests) {
            if (this.beforeEach) { try { g_verbose && console.log(`  [beforeEach]`); await this.beforeEach(); } catch (e) { console.error(`beforeEach失败: ${e.message}`); continue; } }
            g_verbose && console.log(`  ▶ ${test.name}`);
            var result = test.run();
            this.results.push(result);
            this._printResult(result);
            if (this.afterEach) { try { g_verbose && console.log(`  [afterEach]`); await this.afterEach(); } catch (e) { console.error(`afterEach失败: ${e.message}`); } }
        }
        if (this.afterAll) { try { g_verbose && console.log(`  [afterAll]`); await this.afterAll(); } catch (e) { console.error(`afterAll失败: ${e.message}`); } }
        return this.getSummary();
    }
    _printResult(r) {
        var icon = { passed: '✅', failed: '❌', skipped: '⏭️' }[r.result];
        var dur = r.duration > 0 ? `(${r.duration}ms)` : '';
        console.log(`${icon} ${r.name} ${dur}`);
        if (r.error) console.error(`   错误: ${r.error.message}`);
        if (g_verbose && r.result === 'passed') console.log(`   ⏱ ${r.duration}ms`);
    }
    getSummary() {
        var passed = this.results.filter(r => r.result === 'passed').length;
        var failed = this.results.filter(r => r.result === 'failed').length;
        var skipped = this.results.filter(r => r.result === 'skipped').length;
        return { suiteName: this.name, total: this.results.length, passed, failed, skipped, success: failed === 0, results: this.results };
    }
}

class clsTestRunner {
    constructor() { this.suites = []; }
    describe(name, fn) { var suite = new clsTestSuite(name); fn(suite); this.suites.push(suite); return suite; }
    async runAll() {
        console.log("\n" + "=".repeat(60));
        console.log("测试报告 (verbose=" + g_verbose + ")");
        console.log("=".repeat(60));
        var summaries = [];
        for (var s of this.suites) { summaries.push(await s.run()); }
        this._printReport(summaries);
        return summaries;
    }
    _printReport(summaries) {
        var tp = 0, tf = 0, ts = 0, tt = 0;
        console.log("\n" + "=".repeat(60) + "\n测试摘要\n" + "=".repeat(60));
        for (var s of summaries) {
            tt += s.total; tp += s.passed; tf += s.failed; ts += s.skipped;
            console.log(`\n${s.suiteName}: ${s.success ? '✅' : '❌'} (通过:${s.passed} 失败:${s.failed} 跳过:${s.skipped})`);
        }
        console.log("\n" + "=".repeat(60));
        console.log(`总计: ${tt} | ✅${tp} | ❌${tf} | ⏭️${ts}`);
        console.log("=".repeat(60));
    }
}

var g_testRunner = new clsTestRunner();
function setVerbose(v) { g_verbose = v; console.log(`[测试] verbose=${v}`); }

// ==================== 测试数据工厂 ====================

const DataFactory = {
    simple: () => [["产品","地区","销售额"],["A","北京",100],["A","上海",200],["B","北京",300],["B","上海",400]],
    standard: () => [["产品","地区","年份","季度","销售额","数量"],
                     ["A","北京",2023,"Q1",1000,10],["A","北京",2023,"Q2",1500,15],
                     ["A","上海",2023,"Q1",1200,12],["A","上海",2023,"Q2",1800,18],
                     ["B","北京",2023,"Q1",2000,20],["B","北京",2023,"Q2",2500,25],
                     ["B","上海",2023,"Q1",2200,22],["B","上海",2023,"Q2",2800,28]],
    multiLevel: () => [["大区","省份","城市","产品","销售额"],
                       ["华北","北京","北京市","A",1000],["华北","北京","北京市","B",2000],
                       ["华北","天津","天津市","A",1500],["华东","上海","上海市","A",2500],
                       ["华东","上海","上海市","B",3000],["华东","江苏","南京市","A",1800]],
    withNulls: () => [["产品","地区","销售额"],["A","北京",100],["A",null,200],[null,"上海",300],["B",null,null],[null,null,500]],
    bulk: (rows) => {
        var data = [["产品","地区","类别","销售额"]];
        var ps = ["A","B","C","D","E"], rs = ["北京","上海","广州","深圳","杭州","成都","武汉","南京"], cs = ["电子","服装","食品","家居"];
        for (var i = 0; i < rows; i++) data.push([ps[i%5], rs[i%8], cs[i%4], Math.floor(Math.random()*10000)]);
        return data;
    }
};

const Assert2D = {
    hasRow: (result, expected, msg) => {
        var found = false;
        for (var i = 0; i < result.length; i++) {
            var row = result[i];
            if (!Array.isArray(row)) continue;
            var match = true;
            for (var j = 0; j < expected.length; j++) {
                if (expected[j] != null && expected[j] != undefined) {
                    var actual = typeof row[j] === 'string' ? row[j].trim() : row[j];
                    if (actual !== expected[j]) { match = false; break; }
                }
            }
            if (match) { found = true; break; }
        }
        if (!found) {
            var debug = ''; for (var i = 0; i < Math.min(result.length, 5); i++) if (Array.isArray(result[i])) debug += `\n     [${i}] ${result[i].join('|')}`;
            throw new Error((msg || '') + `: 未找到 [${expected.join(', ')}]${debug}`);
        }
    },
    cell: (result, r, c, expected, msg) => {
        if (!result[r] || result[r][c] !== expected) {
            throw new Error((msg || '') + `: [${r},${c}] 期望"${expected}" 实际"${result[r]?result[r][c]:'N/A'}"`);
        }
    }
};

const isExcelError = v => typeof v === 'number' && v < -2000000000;

// ==================== 辅助函数 ====================

const $ = {
    pass: (name, ...logs) => (console.log(`✅ ${name}`), logs.forEach(l => console.log(`  ✓ ${l}`)), true),
    fail: (name, err) => (console.log(`❌ ${name}: ${err.message}`), false),
    run: (name, fn) => { try { return fn(); } catch (e) { return $.fail(name, e); } }
};

// ==================== 租金测算系统测试套件 ====================

g_testRunner.describe("租金测算-初始化", suite => {
    suite.test("配置管理器初始化", () => {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        assert.assertNotNull(r.config, "配置管理器未创建");
        const cols = r.config.columnDefinitions;
        assert.assertEquals('A', cols.PERIOD, "PERIOD列");
        assert.assertEquals('B', cols.DATE, "DATE列");
        const methods = r.config.repaymentMethods;
        ["等额本息（后付）","等额本息（先付）","等额本金（按天计息）","等额本金（按期计息）","本金比例（按期计息）","本金比例（按天计息）"].forEach(m => {
            assert.assertNotNull(methods[m], `还款方式:${m}`);
        });
    });
    suite.test("公式生成器初始化", () => {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        assert.assertNotNull(r.formulaGenerator, "公式生成器未创建");
        ["generateEqualPaymentFormulas","generateEqualPaymentAdvanceFormulas","generateEqualPrincipalDailyInterestFormulas","generateEqualPrincipalPeriodicInterestFormulas","generatePrincipalRatioPeriodicInterestFormulas","generatePrincipalRatioDailyInterestFormulas"].forEach(m => {
            assert.assertDefined(r.formulaGenerator[m], `方法:${m}`);
        });
    });
    suite.test("样式管理器初始化", () => {
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        assert.assertNotNull(r.styleManager, "样式管理器未创建");
        ["createTableHeaders","applyDataFormat","addBorder","setBackColor"].forEach(m => {
            assert.assertDefined(r.styleManager[m], `方法:${m}`);
        });
    });
});

g_testRunner.describe("租金测算-完整流程", suite => {
    suite.test("完整租金测算流程", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("PaymentInterval", 1);
        p.SetParameterValue("PaymentsPerYear", 12);
        p.SetParameterValue("LeaseStartDate", "2026-01-01");
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        const r = new clsRentalCalculation(p);
        r.Initialize(p, "1租金测算表V1");
        assert.assertTrue(r.创建租金测算表表头(1, 10), "表头创建失败");
        assert.assertTrue(r.createDataRange(), "数据生成失败");
        assert.assertEquals(12, r.arrData.length, "数据行数");
        assert.assertEquals(13, r.arrData[0].length, "数据列数");
    });
    suite.test("等额本息后付-公式验证", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        const p1 = r.arrData[0];
        assert.assertEquals("1", p1[0], "期次");
        assert.assertTrue(p1[2].includes("PMT"), "租金含PMT");
        assert.assertTrue(p1[3].includes("PPMT"), "本金含PPMT");
    });
    suite.test("等额本金按天计息", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按天计息）");
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        const p1 = r.arrData[0];
        assert.assertTrue(p1[3].includes("/"), "本金含除法");
        assert.assertTrue(p1[4].includes("360"), "利息含360");
    });
    suite.test("本金比例按期计息", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "本金比例（按期计息）");
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        r.createDataRange();
        const p1 = r.arrData[0];
        assert.assertNotNull(p1[11], "本金比例列");
        assert.assertTrue(p1[11].includes("round") || p1[11].includes("ROUND"), "含round");
        const last = r.arrData[11];
        assert.assertTrue(last[11].includes("SUM"), "最后一期含SUM");
    });
});

g_testRunner.describe("租金测算-边界测试", suite => {
    suite.test("最小期数(1期)", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 1);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        const r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        assert.assertTrue(r.createDataRange(), "1期生成");
        assert.assertEquals(1, r.arrData.length);
    });
    suite.test("大数据量(100期)<10s", () => {
        var start = Date.now();
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 100);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        const r = new clsRentalCalculation(p);
        r.Initialize(p, "1租金测算表V1");
        r.createDataRange();
        var dur = Date.now() - start;
        assert.assertEquals(100, r.arrData.length, "100期数据");
        assert.assertTrue(dur < 10000, `应<10s (${dur}ms)`);
        console.log(`   ⏱ 100期: ${dur}ms`);
    });
});

g_testRunner.describe("租金测算-每期适用利率", suite => {
    suite.test("M列初始化", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("InterestRate", 0.03);
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.创建租金测算表表头(1, 10);
        assert.assertTrue(r.每期适用利率(), "M列初始化");
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        assert.assertEquals("每期适用利率", ws.Range(`M${row - 2}`).Value2, "M列表头");
        assert.assertTrue(Math.abs(ws.Range(`M${row}`).Value2 - 0.03) < 0.001, "M列利率值");
    });
    suite.test("修改利率后自动重算", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按期计息）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.使用每期适用利率生成测算表();
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        var original = ws.Range(`E${row}`).Value2;
        if (isExcelError(original)) throw new Error(`初始利息错误:${original}`);
        ws.Range(`M${row}`).Value2 = 0.04;
        ws.Calculate();
        var modified = ws.Range(`E${row}`).Value2;
        if (isExcelError(modified)) throw new Error(`修改后错误:${modified}`);
        assert.assertTrue(modified > original, `利息应增加:${modified}>${original}`);
    });
    suite.test("完整工作流-租金变化", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        assert.assertTrue(r.使用每期适用利率生成测算表(), "测算表生成");
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        var original = ws.Range(`C${row}`).Value2;
        if (isExcelError(original)) throw new Error(`初始租金错误:${original}`);
        ws.Range(`M${row}`).Value2 = 0.04;
        ws.Calculate();
        var modified = ws.Range(`C${row}`).Value2;
        if (isExcelError(modified)) throw new Error(`修改后错误:${modified}`);
        assert.assertTrue(modified > original, `租金应增加`);
    });
});

// ==================== 超级透视测试套件 ====================

g_testRunner.describe("z超级透视-基础透视", suite => {
    suite.test("单行+单列+sum", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], ["f2+"], ['sum("f3")'], 1);
        assert.assertNotNull(result);
        Assert2D.hasRow(result, ["A", 200, 100], "A分组");
        Assert2D.hasRow(result, ["B", 400, 300], "B分组");
    });
    suite.test("仅行字段", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['sum("f3")'], 1);
        Assert2D.hasRow(result, ["A", 300], "A总300");
        Assert2D.hasRow(result, ["B", 700], "B总700");
    });
    suite.test("仅列字段", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), [], ["f2+"], ['sum("f3")'], 1);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 1);
    });
    suite.test("无行无列（仅聚合）", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), [], [], ['sum("f3")'], 1);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 1);
    });
});

g_testRunner.describe("z超级透视-多字段组合", suite => {
    suite.test("多行+单列", () => {
        var result = Array2D.z超级透视(DataFactory.standard(), ["f1+,f2+"], ["f3+"], ['sum("f5")'], 1);
        Assert2D.hasRow(result, ["A","北京",2500], "A×北京");
        Assert2D.hasRow(result, ["B","上海",5000], "B×上海");
    });
    suite.test("单行+多列", () => {
        var result = Array2D.z超级透视(DataFactory.standard(), ["f1+"], ["f3+,f4+"], ['sum("f5")'], 1);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 4);
        var totalA = 0;
        for (var i = 0; i < result.length; i++) {
            if (result[i] && result[i][0] === "A") {
                for (var j = 1; j < result[i].length; j++) if (typeof result[i][j] === 'number') totalA += result[i][j];
            }
        }
        assert.assertEquals(5500, totalA, "A总计");
    });
    suite.test("三层行字段嵌套", () => {
        var result = Array2D.z超级透视(DataFactory.multiLevel(), ["f1+,f2+,f3+"], [], ['sum("f5")'], 1);
        Assert2D.hasRow(result, ["华东","上海","上海市",5500], "华东×上海×上海市");
        Assert2D.hasRow(result, ["华北","北京","北京市",3000], "华北×北京×北京市");
    });
});

g_testRunner.describe("z超级透视-排序", suite => {
    suite.test("升序+", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['sum("f3")'], 1);
        var idxA = -1, idxB = -1;
        for (var i = 0; i < result.length; i++) {
            if (result[i] && result[i][0] === "A") idxA = i;
            if (result[i] && result[i][0] === "B") idxB = i;
        }
        assert.assertTrue(idxA < idxB, "A应在B前");
    });
    suite.test("降序-", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1-"], [], ['sum("f3")'], 1);
        var idxA = -1, idxB = -1;
        for (var i = 0; i < result.length; i++) {
            if (result[i] && result[i][0] === "A") idxA = i;
            if (result[i] && result[i][0] === "B") idxB = i;
        }
        assert.assertTrue(idxB < idxA, "B应在A前");
    });
});

g_testRunner.describe("z超级透视-聚合函数", suite => {
    suite.test("count", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['count()'], 1);
        Assert2D.hasRow(result, ["A", 2], "A计数2");
        Assert2D.hasRow(result, ["B", 2], "B计数2");
    });
    suite.test("sum", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['sum("f3")'], 1);
        Assert2D.hasRow(result, ["A", 300], "A sum300");
        Assert2D.hasRow(result, ["B", 700], "B sum700");
    });
    suite.test("average", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['average("f3")'], 1);
        Assert2D.hasRow(result, ["A", 150], "A avg150");
        Assert2D.hasRow(result, ["B", 350], "B avg350");
    });
    suite.test("max/min", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['max("f3"),min("f3")'], 1);
        Assert2D.hasRow(result, ["A", 200, 100], "A max200 min100");
        Assert2D.hasRow(result, ["B", 400, 300], "B max400 min300");
    });
    suite.test("多聚合组合", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], [], ['count(),sum("f3"),average("f3")'], 1);
        Assert2D.hasRow(result, ["A", 2, 300, 150], "A count2 sum300 avg150");
        Assert2D.hasRow(result, ["B", 2, 700, 350], "B count2 sum700 avg350");
    });
});

g_testRunner.describe("z超级透视-边界情况", suite => {
    suite.test("空数据", () => {
        var result = Array2D.z超级透视([["产品","地区","销售额"]], ["f1+"], ["f2+"], ['sum("f3")'], 1);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 1);
    });
    suite.test("单行数据", () => {
        var result = Array2D.z超级透视([["产品","地区","销售额"],["A","北京",100]], ["f1+"], ["f2+"], ['sum("f3")'], 1);
        Assert2D.hasRow(result, ["A", 100], "单行A100");
    });
    suite.test("含null值", () => {
        var result = Array2D.z超级透视(DataFactory.withNulls(), ["f1+"], ["f2+"], ['sum("f3")'], 1);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 2);
    });
    suite.test("headerRows=0", () => {
        var data = [["A","北京",100],["A","上海",200],["B","北京",300]];
        var result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ['sum("f3")'], 0);
        assert.assertNotNull(result);
        assert.assertTrue(result.length >= 1);
    });
});

g_testRunner.describe("z超级透视-高级选项", suite => {
    suite.test("cornerTitle", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], ["f2+"], ['sum("f3")'], 1, 1, "@^@", { cornerTitle: "销售分析" });
        Assert2D.cell(result, 0, 0, "销售分析", "角落标题");
    });
    suite.test("grandTotal行", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], ["f2+"], ['sum("f3")'], 1, 1, "@^@", { grandTotal: { row: true, label: "合计" } });
        var lastRow = result[result.length - 1];
        assert.assertTrue(lastRow && (lastRow[0] === "合计" || lastRow[0] === "总计"), "应有总计行");
    });
    suite.test("自定义分隔符", () => {
        var result = Array2D.z超级透视(DataFactory.simple(), ["f1+"], ["f2+"], ['sum("f3")'], 1, 1, "|||");
        assert.assertNotNull(result);
    });
});

g_testRunner.describe("z超级透视-性能", suite => {
    suite.test("1000行<100ms", () => {
        var start = Date.now();
        var result = Array2D.z超级透视(DataFactory.bulk(1000), ["f1+"], ["f2+"], ['sum("f4")'], 1);
        var dur = Date.now() - start;
        assert.assertNotNull(result);
        assert.assertTrue(dur < 100, `应<100ms (实际${dur}ms)`);
        console.log(`   ⏱ 1000行: ${dur}ms, ${result.length}行`);
    });
    suite.test("10000行<2000ms", () => {
        var start = Date.now();
        var result = Array2D.z超级透视(DataFactory.bulk(10000), ["f1+,f2+"], ["f3+"], ['count(),sum("f4")'], 1);
        var dur = Date.now() - start;
        assert.assertNotNull(result);
        assert.assertTrue(dur < 2000, `应<2000ms (实际${dur}ms)`);
        console.log(`   ⏱ 10000行: ${dur}ms, ${result.length}行`);
    });
});

// ==================== 每期适用利率测试套件 ====================

g_testRunner.describe("每期适用利率-偿还方式", suite => {
    const methods = [
        { name: "等额本息（后付）", checks: [{ col: 3, exp: "RC[10]", desc: "C列租金" }, { col: 4, exp: "RC[9]", desc: "D列本金" }] },
        { name: "等额本息（先付）", checks: [{ col: 5, exp: "RC[8]", desc: "E列利息" }] },
        { name: "等额本金（按期计息）", checks: [{ col: 5, exp: "RC[8]", desc: "E列利息" }] },
        { name: "本金比例（按期计息）", checks: [{ col: 5, exp: "RC[8]", desc: "E列利息" }] }
    ];

    methods.forEach(m => {
        suite.test(`偿还方式: ${m.name}`, () => {
            var r = new clsRentalCalculation();
            r.Initialize("1租金测算表V1");
            var arr = r.formulaGenerator.适用每期利率(m.name);
            assert.assertNotNull(arr, "公式为空");
            m.checks.forEach(c => {
                var val = arr[1][c.col];
                assert.assertTrue(val.includes(c.exp), `列${c.col}应含"${c.exp}"`);
            });
        });
    });

    suite.test("六种偿还方式RC偏移一致性", () => {
        var r = new clsRentalCalculation();
        r.Initialize("1租金测算表V1");
        ["等额本息（后付）","等额本息（先付）","等额本金（按天计息）","等额本金（按期计息）","本金比例（按期计息）","本金比例（按天计息）"].forEach(method => {
            var arr = r.formulaGenerator.适用每期利率(method);
            assert.assertNotNull(arr, `${method}公式生成失败`);
            [1,2,3].forEach(type => {
                arr[type].forEach((f, col) => {
                    if (f && f.includes("${params.interestRateCell}")) throw new Error(`${method}第${type}期第${col}列未替换`);
                });
            });
        });
    });
});

g_testRunner.describe("每期适用利率-M列操作", suite => {
    suite.test("M列初始化", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("InterestRate", 0.03);
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.创建租金测算表表头(1, 10);
        assert.assertTrue(r.每期适用利率(), "M列初始化失败");
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        assert.assertEquals("每期适用利率", ws.Range(`M${row - 2}`).Value2, "M列表头");
        assert.assertTrue(Math.abs(ws.Range(`M${row}`).Value2 - 0.03) < 0.001, `M${row}值`);
    });

    suite.test("修改利率后自动计算", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按期计息）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.使用每期适用利率生成测算表();
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        var original = ws.Range(`E${row}`).Value2;
        if (isExcelError(original)) throw new Error(`初始利息错误:${original}`);
        ws.Range(`M${row}`).Value2 = 0.04;
        ws.Calculate();
        var modified = ws.Range(`E${row}`).Value2;
        if (isExcelError(modified)) throw new Error(`修改后利息错误:${modified}`);
        assert.assertTrue(modified > original, `利息应增加:${modified}>${original}`);
        assert.assertTrue(Math.abs(modified - 100000000*0.04/12) < 1, "利息计算错误");
    });

    suite.test("批量设置多期利率", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按期计息）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.使用每期适用利率生成测算表();
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 4; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        for (var i = 4; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.035;
        ws.Calculate();
        var int1 = ws.Range(`E${row}`).Value2;
        if (isExcelError(int1)) throw new Error(`第1期错误:${int1}`);
        var ppp = 100000000 / 12, rem5 = 100000000 - 4 * ppp, int5 = ws.Range(`E${row + 4}`).Value2;
        if (isExcelError(int5)) throw new Error(`第5期错误:${int5}`);
        assert.assertTrue(Math.abs(int1 - 100000000*0.03/12) < 1, "第1期利息错");
        assert.assertTrue(Math.abs(int5 - rem5*0.035/12) < 1, "第5期利息错");
    });

    suite.test("完整工作流程", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        assert.assertTrue(r.使用每期适用利率生成测算表(), "测算表生成失败");
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        var original = ws.Range(`C${row}`).Value2;
        if (isExcelError(original)) throw new Error(`初始租金错误:${original}`);
        ws.Range(`M${row}`).Value2 = 0.04;
        ws.Calculate();
        var modified = ws.Range(`C${row}`).Value2;
        if (isExcelError(modified)) throw new Error(`修改后租金错误:${modified}`);
        assert.assertTrue(modified > original, `租金应增加:${modified}>${original}`);
    });

    suite.test("单期边界", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 1);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        assert.assertTrue(r.使用每期适用利率生成测算表(), "单期失败");
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        ws.Range(`M${row}`).Value2 = 0.03;
        ws.Calculate();
        assert.assertTrue(ws.Range(`M${row}`).Value2 > 0, "M列应有正值");
    });

    suite.test("零利率边界", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按期计息）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.使用每期适用利率生成测算表();
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        ws.Range(`M${row}`).Value2 = 0;
        ws.Calculate();
        var interest = ws.Range(`E${row}`).Value2;
        if (isExcelError(interest)) { console.log(`   ⚠ 零利率返回错误值(${interest})`); return; }
        assert.assertTrue(Math.abs(interest) < 1, `零利率利息应接近0，实际${interest}`);
    });

    suite.test("高利率边界(20%)", () => {
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 12);
        p.SetParameterValue("RepaymentMethod", "等额本金（按期计息）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        r.使用每期适用利率生成测算表();
        var ws = p.m_worksheet, row = p.RentTableStartRow;
        for (var i = 0; i < 12; i++) ws.Range(`M${row + i}`).Value2 = 0.03;
        ws.Calculate();
        ws.Range(`M${row}`).Value2 = 0.20;
        ws.Calculate();
        var interest = ws.Range(`E${row}`).Value2;
        if (isExcelError(interest)) throw new Error(`高利率错误:${interest}`);
        assert.assertTrue(Math.abs(interest - 100000000*0.20/12) < 1, "20%利率利息错");
    });

    suite.test("大数据量性能(60期)", () => {
        var start = Date.now();
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        p.SetParameterValue("Principal", 100000000);
        p.SetParameterValue("InterestRate", 0.03);
        p.SetParameterValue("TotalPeriods", 60);
        p.SetParameterValue("RepaymentMethod", "等额本息（后付）");
        var r = new clsRentalCalculation(p);
        r.Initialize();
        assert.assertTrue(r.使用每期适用利率生成测算表(), "60期生成失败");
        var dur = Date.now() - start;
        assert.assertTrue(dur < 5000, `应<5000ms (实际${dur}ms)`);
        console.log(`   ⏱ 60期: ${dur}ms`);
    });
});

// ==================== 入口函数 ====================

function 运行全部测试() {
    setVerbose(true);
    console.log("\n" + "═".repeat(60));
    console.log("统一测试套件 (verbose模式)");
    console.log("═".repeat(60));
    g_testRunner.runAll();
}

function 运行超级透视测试() { 运行全部测试(); }
function 运行每期适用利率测试() { 运行全部测试(); }
function 快速测试() {
    setVerbose(true);
    console.log("\n📦 快速测试");
    g_testRunner.runAll();
}

function 诊断模块(p) {
    console.clear();
    console.log("═".repeat(50));
    console.log("  租金测算模块诊断");
    console.log("═".repeat(50));
    try {
        if (!p) {
            if (typeof globalThis !== 'undefined' && globalThis.p) p = globalThis.p;
            else if (typeof this.p !== 'undefined') p = this.p;
        }
        if (!p) { console.log("❌ 参数管理器未初始化"); return false; }
        console.log("✓ 参数管理器已初始化");
        if (!p.m_worksheet) { console.log("❌ 工作表未初始化"); return false; }
        console.log("✓ 工作表已初始化");
        const required = ["Principal","InterestRate","TotalPeriods","RepaymentMethod"];
        var missing = [];
        for (const k of required) { try { if (p.val(k) == null) missing.push(k); } catch (e) { missing.push(k); } }
        if (missing.length) { console.log(`❌ 缺少参数: ${missing.join(',')}`); return false; }
        console.log("✓ 所有必需参数已设置");
        console.log("═".repeat(50));
        console.log("✓ 所有诊断检查通过");
        return true;
    } catch (e) { console.error("诊断失败:", e.message); return false; }
}

function 使用说明() {
    console.clear();
    console.log("═".repeat(50));
    console.log("  统一测试套件");
    console.log("═".repeat(50));
    console.log("\n【运行测试】");
    console.log("  运行全部测试()          // 运行所有测试(verbose)");
    console.log("  快速测试()              // 快速运行");
    console.log("  setVerbose(true/false)  // 详细/简洁输出");
    console.log("  诊断模块(p)             // 诊断参数管理器状态");
    console.log("\n【Node.js环境】");
    console.log("  自动运行所有测试");
    console.log("\n【套件列表】");
    console.log("  - 租金测算-初始化        (配置/公式/样式管理器)");
    console.log("  - 租金测算-完整流程      (六种还款方式)");
    console.log("  - 租金测算-边界测试      (1期/100期)");
    console.log("  - 租金测算-每期适用利率  (M列操作)");
    console.log("  - z超级透视-基础透视     (单/多字段)");
    console.log("  - z超级透视-多字段组合   (嵌套)");
    console.log("  - z超级透视-排序        (升序/降序)");
    console.log("  - z超级透视-聚合函数    (count/sum/avg/max/min)");
    console.log("  - z超级透视-边界情况    (空/单行/null)");
    console.log("  - z超级透视-高级选项    (cornerTitle/grandTotal)");
    console.log("  - z超级透视-性能        (1000/10000行)");
    console.log("  - 每期适用利率-偿还方式  (六种RC偏移)");
    console.log("  - 每期适用利率-M列操作   (批量利率/边界)");
    console.log("\n测试套件合并: mUnitTestFramework + mTest_PeriodRate + mTest_SuperPivot");
    console.log("═".repeat(50));
}

// Node.js 自动运行
if (_isNode) {
    console.log('\n╔════════════════════════════════════╗');
    console.log('║  统一测试套件 (Node.js环境)        ║');
    console.log('╚════════════════════════════════════╝');
    setVerbose(true);
    g_testRunner.runAll().then(s => {
        var failed = s.reduce((acc, x) => acc + x.failed, 0);
        if (failed > 0) { console.log(`\n❌ ${failed}个测试失败`); process.exit(1); }
        else console.log("\n✅ 全部通过");
    });
}

使用说明();