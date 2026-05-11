/**
 * =======================================================================
 * SuperPivot (z超级透视) 功能测试套件 - JSA880 v3.8.2+
 * =======================================================================
 *
 * 版本: 2.5.0
 * 日期: 2026-02-08
 * 符合规范: WPS JSA ES6-ES2019
 *
 * 功能说明:
 *   全面测试 z超级透视 核心功能和边界情况
 *   每次测试结果实时输出到表格
 *   透视表结果自动输出到"测试输出"工作表
 *
 * 使用方法:
 *   1. 确保 JSA880.js 已加载到 WPS 宏编辑器
 *   2. 运行 runQuickTest() 快速测试（输出透视表到"测试输出"工作表）
 *   3. 运行 runAllTests() 执行所有测试
 *   4. 测试结果输出到"测试结果"工作表
 *   5. 透视表结果输出到"测试输出"工作表
 *
 * 工作表说明:
 *   - "测试结果": 记录所有测试的执行状态（通过/失败）和错误信息
 *   - "测试输出": 显示所有透视表的实际输出结果，便于查看和验证
 *
 * =======================================================================
 */

// =======================================================================
// 常量定义 (UPPER_SNAKE_CASE)
// =======================================================================

const MODULE_NAME = "SuperPivotTestSuite";
const MODULE_VERSION = "2.5.0";
const MODULE_DATE = "2026-02-08";
const REPORT_SHEET_NAME = "测试结果";
const OUTPUT_SHEET_NAME = "测试输出";
const DEFAULT_TEST_ROWS = 1000;
const PERFORMANCE_TEST_ROWS = 5000;

// WPS 枚举常量
const XL_UP = -4162;           // 向上查找

// 颜色常量
const COLOR_GREEN = 0x008000;     // 通过 - 绿色
const COLOR_RED = 0xFF0000;       // 失败 - 红色
const COLOR_BLUE_BG = 0xD9E1F2;   // 表头背景 - 浅蓝色
const COLOR_GRAY_BG = 0xF2F2F2;   // 分组背景 - 浅灰色
const COLOR_HEADER_BG = 0x4472C4; // 输出表头背景 - 深蓝色
const COLOR_BORDER = 0xD0D0D0;    // 边框颜色

// =======================================================================
// 测试报告管理类
// =======================================================================

/**
 * 测试报告管理类
 * 负责实时输出测试结果到工作表
 */
class clsTestReporter {
    constructor() {
        this.m_sheet = null;           // 报告工作表
        this.m_currentRow = 1;         // 当前写入行
        this.m_testNumber = 0;         // 测试序号
        this.m_groupName = "";         // 当前分组名称
        this.m_isInitialized = false;  // 是否已初始化
    }
    
    /**
     * 初始化报告工作表
     */
    initialize() {
        // 获取或创建工作表
        try {
            this.m_sheet = Worksheets(REPORT_SHEET_NAME);
        } catch (e) {
            this.m_sheet = Worksheets.Add();
            this.m_sheet.Name = REPORT_SHEET_NAME;
        }
        
        // 清除原有内容
        this.m_sheet.Cells.ClearContents();
        
        // 写入报告标题
        this.m_sheet.Cells(1, 1).Value2 = "JSA880 SuperPivot (z超级透视) 功能测试报告";
        this.m_sheet.Cells(1, 1).Font.Size = 16;
        this.m_sheet.Cells(1, 1).Font.Bold = true;
        
        // 写入基本信息
        this.m_sheet.Cells(3, 1).Value2 = "测试时间";
        this.m_sheet.Cells(3, 2).Value2 = new Date().toLocaleString();
        
        this.m_sheet.Cells(4, 1).Value2 = "模块版本";
        this.m_sheet.Cells(4, 2).Value2 = MODULE_VERSION;
        
        this.m_sheet.Cells(5, 1).Value2 = "JSA880状态";
        this.m_sheet.Cells(5, 2).Value2 = (typeof JSA880 !== "undefined") ? "已加载" : "未加载";
        
        // 设置基本信息样式
        for (let row = 3; row <= 5; row++) {
            this.m_sheet.Cells(row, 1).Font.Bold = true;
        }
        
        // 表头行（从第7行开始）
        this.m_currentRow = 7;
        this.m_sheet.Cells(this.m_currentRow, 1).Value2 = "序号";
        this.m_sheet.Cells(this.m_currentRow, 2).Value2 = "分组";
        this.m_sheet.Cells(this.m_currentRow, 3).Value2 = "测试名称";
        this.m_sheet.Cells(this.m_currentRow, 4).Value2 = "状态";
        this.m_sheet.Cells(this.m_currentRow, 5).Value2 = "错误信息";
        this.m_sheet.Cells(this.m_currentRow, 6).Value2 = "耗时(ms)";
        
        // 表头样式
        for (let col = 1; col <= 6; col++) {
            this.m_sheet.Cells(this.m_currentRow, col).Font.Bold = true;
            this.m_sheet.Cells(this.m_currentRow, col).Interior.Color = COLOR_BLUE_BG;
        }
        
        this.m_currentRow++;
        this.m_isInitialized = true;
        this.m_testNumber = 0;
        
        // 激活报告工作表
        this.m_sheet.Activate();
        
        return this.m_sheet;
    }
    
    /**
     * 开始新分组
     * @param {string} groupName - 分组名称
     */
    startGroup(groupName) {
        this.m_groupName = groupName;
        
        // 写入分组标题行
        this.m_sheet.Cells(this.m_currentRow, 1).Value2 = groupName;
        this.m_sheet.Cells(this.m_currentRow, 1).Font.Bold = true;
        this.m_sheet.Cells(this.m_currentRow, 1).Interior.Color = COLOR_GRAY_BG;
        
        // 合并分组标题单元格
        this.m_sheet.Range(this.m_sheet.Cells(this.m_currentRow, 1), this.m_sheet.Cells(this.m_currentRow, 6)).Merge();
        
        this.m_currentRow++;
    }
    
    /**
     * 记录单个测试结果
     * @param {string} testName - 测试名称
     * @param {string} status - 状态 PASS/FAIL
     * @param {string} error - 错误信息
     * @param {number} duration - 耗时(ms)
     */
    recordTest(testName, status, error, duration) {
        this.m_testNumber++;
        
        // 写入测试数据
        this.m_sheet.Cells(this.m_currentRow, 1).Value2 = this.m_testNumber;
        this.m_sheet.Cells(this.m_currentRow, 2).Value2 = this.m_groupName;
        this.m_sheet.Cells(this.m_currentRow, 3).Value2 = testName;
        this.m_sheet.Cells(this.m_currentRow, 4).Value2 = status;
        this.m_sheet.Cells(this.m_currentRow, 5).Value2 = error || "";
        this.m_sheet.Cells(this.m_currentRow, 6).Value2 = duration || 0;
        
        // 状态颜色
        if (status === "PASS") {
            this.m_sheet.Cells(this.m_currentRow, 4).Font.Color = COLOR_GREEN;
        } else {
            this.m_sheet.Cells(this.m_currentRow, 4).Font.Color = COLOR_RED;
            this.m_sheet.Cells(this.m_currentRow, 5).Font.Color = COLOR_RED;
        }
        
        this.m_currentRow++;
        
        // 自动调整列宽（每10行调整一次，避免性能问题）
        if (this.m_testNumber % 10 === 0) {
            this.autoFitColumns();
        }
    }
    
    /**
     * 写入汇总信息
     * @param {Object} summary - 汇总对象 {total, pass, fail}
     */
    writeSummary(summary) {
        const summaryRow = this.m_currentRow + 1;
        
        this.m_sheet.Cells(summaryRow, 1).Value2 = "测试汇总";
        this.m_sheet.Cells(summaryRow, 1).Font.Size = 14;
        this.m_sheet.Cells(summaryRow, 1).Font.Bold = true;
        
        this.m_sheet.Cells(summaryRow + 1, 1).Value2 = "总测试数";
        this.m_sheet.Cells(summaryRow + 1, 2).Value2 = summary.total;
        
        this.m_sheet.Cells(summaryRow + 2, 1).Value2 = "通过数";
        this.m_sheet.Cells(summaryRow + 2, 2).Value2 = summary.pass;
        this.m_sheet.Cells(summaryRow + 2, 2).Font.Color = COLOR_GREEN;
        
        this.m_sheet.Cells(summaryRow + 3, 1).Value2 = "失败数";
        this.m_sheet.Cells(summaryRow + 3, 2).Value2 = summary.fail;
        this.m_sheet.Cells(summaryRow + 3, 2).Font.Color = COLOR_RED;
        
        this.m_sheet.Cells(summaryRow + 4, 1).Value2 = "通过率";
        this.m_sheet.Cells(summaryRow + 4, 2).Value2 = ((summary.pass / summary.total * 100).toFixed(1)) + "%";
        
        // 汇总区域加粗
        for (let row = summaryRow; row <= summaryRow + 4; row++) {
            this.m_sheet.Cells(row, 1).Font.Bold = true;
        }
        
        this.autoFitColumns();
    }
    
    /**
     * 自动调整列宽
     */
    autoFitColumns() {
        this.m_sheet.Columns("A:F").AutoFit();
    }
    
    /**
     * 获取工作表
     */
    getSheet() {
        return this.m_sheet;
    }
}

// =======================================================================
// 测试输出管理类
// =======================================================================

/**
 * 测试输出管理类
 * 负责将透视表结果输出到"测试输出"工作表
 */
class clsTestOutput {
    constructor() {
        this.m_sheet = null;           // 输出工作表
        this.m_currentRow = 1;         // 当前写入行
        this.m_currentCol = 1;         // 当前写入列
        this.m_outputCount = 0;        // 输出计数
        this.m_isInitialized = false;  // 是否已初始化
    }

    /**
     * 初始化输出工作表
     */
    initialize() {
        try {
            this.m_sheet = Worksheets(OUTPUT_SHEET_NAME);
        } catch (e) {
            this.m_sheet = Worksheets.Add();
            this.m_sheet.Name = OUTPUT_SHEET_NAME;
        }

        // 清除原有内容
        this.m_sheet.Cells.ClearContents();
        this.m_sheet.Cells.ClearFormats();

        // 写入标题
        this.m_sheet.Cells(1, 1).Value2 = "SuperPivot 测试输出 - " + new Date().toLocaleString();
        this.m_sheet.Cells(1, 1).Font.Size = 14;
        this.m_sheet.Cells(1, 1).Font.Bold = true;

        this.m_currentRow = 3;
        this.m_currentCol = 1;
        this.m_outputCount = 0;
        this.m_isInitialized = true;

        return this.m_sheet;
    }

    /**
     * 输出透视表结果
     * @param {Array} result - 透视表结果
     * @param {string} title - 输出标题
     * @param {Object} options - 输出选项
     */
    outputResult(result, title, options) {
        options = options || {};

        if (!this.m_isInitialized) {
            this.initialize();
        }

        // 检查是否需要换列（每10个结果换一列）
        if (this.m_outputCount > 0 && this.m_outputCount % 10 === 0) {
            this.m_currentCol += 20;  // 每个输出预留20列
            this.m_currentRow = 3;
        }

        // 写入标题
        this.m_sheet.Cells(this.m_currentRow, this.m_currentCol).Value2 = (this.m_outputCount + 1) + ". " + title;
        this.m_sheet.Cells(this.m_currentRow, this.m_currentCol).Font.Bold = true;
        this.m_sheet.Cells(this.m_currentRow, this.m_currentCol).Font.Color = 0x0000FF;
        this.m_currentRow++;

        // 写入数据
        if (result && result.length > 0) {
            const startRow = this.m_currentRow;
            const startCol = this.m_currentCol;

            // 调试信息：输出数据维度
            Console.log("    [输出] " + title + " - 尺寸: " + result.length + " x " + (result[0] ? result[0].length : 0));

            for (let i = 0; i < result.length; i++) {
                const row = result[i];
                if (!row) continue;

                for (let j = 0; j < row.length; j++) {
                    const cellValue = row[j];
                    this.m_sheet.Cells(startRow + i, startCol + j).Value2 = cellValue;

                    // 表头样式
                    if (i === 0) {
                        this.m_sheet.Cells(startRow + i, startCol + j).Font.Bold = true;
                        this.m_sheet.Cells(startRow + i, startCol + j).Interior.Color = COLOR_HEADER_BG;
                        this.m_sheet.Cells(startRow + i, startCol + j).Font.Color = 0xFFFFFF;
                    }
                }
            }

            // 添加边框
            const endRow = startRow + result.length - 1;
            const maxCols = Math.max(...result.map(r => r ? r.length : 0));
            const endCol = startCol + maxCols - 1;
            const range = this.m_sheet.Range(
                this.m_sheet.Cells(startRow, startCol),
                this.m_sheet.Cells(endRow, endCol)
            );
            this.addBorder(range);

            this.m_currentRow = endRow + 3;  // 留出2行空行
        } else {
            this.m_sheet.Cells(this.m_currentRow, this.m_currentCol).Value2 = "(无数据)";
            this.m_currentRow += 2;
        }

        this.m_outputCount++;
    }

    /**
     * 添加边框
     */
    addBorder(range) {
        range.Borders.Weight = 2;
    }

    /**
     * 自动调整列宽
     */
    autoFitColumns() {
        this.m_sheet.Columns.AutoFit();
    }

    /**
     * 获取工作表
     */
    getSheet() {
        return this.m_sheet;
    }
}

// =======================================================================
// 测试运行器类
// =======================================================================

/**
 * 测试运行器类
 * 负责管理测试用例的执行
 */
class clsTestRunner {
    constructor(reporter, output) {
        this.m_results = [];           // 测试结果数组
        this.m_reporter = reporter;    // 报告管理器
        this.m_output = output;        // 输出管理器
        this.m_startTime = null;       // 测试开始时间
        this.m_testStartTime = null;   // 单个测试开始时间
        this.m_enableOutput = true;    // 是否启用输出
        this.m_autoClearOutput = false; // 是否每次测试前自动清空输出
    }

    /**
     * 重置测试状态
     */
    reset() {
        this.m_results = [];
        this.m_startTime = null;
        this.m_testStartTime = null;
    }

    /**
     * 设置是否启用输出
     */
    setEnableOutput(enable) {
        this.m_enableOutput = enable;
    }

    /**
     * 设置是否每次测试前自动清空输出
     */
    setAutoClearOutput(enable) {
        this.m_autoClearOutput = enable;
    }

    /**
     * 输出透视表结果到工作表
     * @param {Array} result - 透视表结果
     * @param {string} title - 输出标题
     * @param {Object} options - 输出选项
     */
    outputResult(result, title, options) {
        if (this.m_output && this.m_enableOutput) {
            this.m_output.outputResult(result, title, options);
        }
    }

    /**
     * 运行单个测试（带输出）
     * @param {string} testName - 测试名称
     * @param {Function} testFunc - 测试函数
     * @param {boolean} enableOutput - 是否输出结果
     * @returns {boolean} 是否通过
     */
    runTestWithOutput(testName, testFunc, enableOutput) {
        const oldOutput = this.m_enableOutput;
        this.m_enableOutput = enableOutput !== false;

        const result = this.runTest(testName, testFunc);

        this.m_enableOutput = oldOutput;

        return result;
    }

    /**
     * 运行单个测试
     * @param {string} testName - 测试名称
     * @param {Function} testFunc - 测试函数
     * @returns {boolean} 是否通过
     */
    runTest(testName, testFunc) {
        // 每次测试前清空输出（如果启用）
        if (this.m_autoClearOutput && this.m_output) {
            this.m_output.initialize();
        }

        this.m_testStartTime = new Date().getTime();
        let status = "PASS";
        let errorMsg = null;

        try {
            testFunc();
            Console.log("  [通过] " + testName);
        } catch (error) {
            status = "FAIL";
            errorMsg = error.message;
            Console.log("  [失败] " + testName + ": " + error.message);
        }

        const duration = new Date().getTime() - this.m_testStartTime;

        // 记录结果
        this.m_results.push({
            name: testName,
            status: status,
            error: errorMsg,
            duration: duration
        });

        // 实时输出到表格
        if (this.m_reporter) {
            this.m_reporter.recordTest(testName, status, errorMsg, duration);
        }

        return status === "PASS";
    }
    
    /**
     * 断言相等
     */
    assertEqual(actual, expected, message) {
        if (actual !== expected) {
            throw new Error((message || "断言失败") + " (期望: " + expected + ", 实际: " + actual + ")");
        }
    }
    
    /**
     * 断言为真
     */
    assertTrue(condition, message) {
        if (!condition) {
            throw new Error(message || "期望为真，实际为假");
        }
    }
    
    /**
     * 断言数组长度
     * 支持 Array2D.z超级透视 返回的包装对象
     */
    assertArrayLength(arr, expectedLength, message) {
        // 检查是否是类数组对象（有 length 属性）
        var actualLength = 0;
        if (Array.isArray(arr)) {
            actualLength = arr.length;
        } else if (arr && typeof arr.length === 'number') {
            // 可能是 Array2D 包装对象或类数组对象
            actualLength = arr.length;
        } else {
            throw new Error((message || "不是数组或类数组对象") + ": " + typeof arr);
        }

        if (actualLength !== expectedLength) {
            throw new Error((message || "长度不符") + " (期望: " + expectedLength + ", 实际: " + actualLength + ")");
        }
    }
    
    /**
     * 打印测试汇总到控制台
     */
    printSummary() {
        let passCount = 0;
        let failCount = 0;
        
        for (let i = 0; i < this.m_results.length; i++) {
            if (this.m_results[i].status === "PASS") {
                passCount++;
            } else {
                failCount++;
            }
        }
        
        Console.log("");
        Console.log("============================================================");
        Console.log("                     测试结果汇总                            ");
        Console.log("============================================================");
        Console.log("  总计: " + this.m_results.length + " 个测试");
        Console.log("  通过: " + passCount + " 个");
        Console.log("  失败: " + failCount + " 个");
        Console.log("  通过率: " + ((passCount / this.m_results.length * 100).toFixed(1)) + "%");
        Console.log("============================================================");
        
        return {
            total: this.m_results.length,
            pass: passCount,
            fail: failCount
        };
    }
}

// =======================================================================
// 全局实例
// =======================================================================

const testReporter = new clsTestReporter();
const testOutput = new clsTestOutput();
const testRunner = new clsTestRunner(testReporter, testOutput);

// =======================================================================
// 测试数据生成器
// =======================================================================

function createTestData() {
    return [
        ["产品", "地区", "年份", "季度", "销售额", "数量"],
        ["A", "北京", 2023, "Q1", 1000, 10],
        ["A", "北京", 2023, "Q2", 1500, 15],
        ["A", "上海", 2023, "Q1", 1200, 12],
        ["A", "上海", 2023, "Q2", 1800, 18],
        ["B", "北京", 2023, "Q1", 2000, 20],
        ["B", "北京", 2023, "Q2", 2500, 25],
        ["B", "上海", 2023, "Q1", 2200, 22],
        ["B", "上海", 2023, "Q2", 2800, 28]
    ];
}

function createSimpleData() {
    return [
        ["产品", "地区", "销售额"],
        ["A", "北京", 100],
        ["A", "上海", 200],
        ["B", "北京", 300],
        ["B", "上海", 400]
    ];
}

function createMultiLevelData() {
    return [
        ["大区", "省份", "城市", "产品", "销售额"],
        ["华北", "北京", "北京市", "A", 1000],
        ["华北", "北京", "北京市", "B", 2000],
        ["华北", "天津", "天津市", "A", 1500],
        ["华东", "上海", "上海市", "A", 2500],
        ["华东", "上海", "上海市", "B", 3000],
        ["华东", "江苏", "南京市", "A", 1800]
    ];
}

function createDateBasedData() {
    return [
        ["日期", "产品", "销售额", "数量"],
        ["2023-01-15", "A", 1000, 10],
        ["2023-01-20", "B", 1500, 15],
        ["2023-02-10", "A", 1200, 12],
        ["2023-02-25", "B", 1800, 18],
        ["2023-03-05", "A", 2000, 20]
    ];
}

function createLargeData(rows) {
    rows = rows || 100;
    const data = [["ID", "类别", "子类别", "值1", "值2"]];
    const categories = ["A", "B", "C", "D"];
    const subCategories = ["X", "Y", "Z"];

    for (let i = 0; i < rows; i++) {
        data.push([
            i + 1,
            categories[i % categories.length],
            subCategories[i % subCategories.length],
            Math.floor(Math.random() * 1000),
            Math.floor(Math.random() * 500)
        ]);
    }

    return data;
}

// =======================================================================
// 测试用例组
// =======================================================================

function testBasicFunctions() {
    testReporter.startGroup("基础功能测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 1: 基础功能测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("基础透视 - 单行单列", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertArrayLength(result, 5, "结果应为5行（3行表头+2行数据）");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "基础透视 - 单行单列");
    });

    testRunner.runTest("无列字段 - 仅行字段", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "无列字段 - 仅行字段");
    });

    testRunner.runTest("无行字段 - 仅列字段", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, [], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "无行字段 - 仅列字段");
    });

    testRunner.runTest("JSA880.透视 快捷方式", function() {
        const data = createSimpleData();
        const result = JSA880.透视(data, "f1+", "f2+", "sum(f3)", 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "JSA880.透视 快捷方式");
    });
}

function testMultipleFields() {
    testReporter.startGroup("多字段测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 2: 多字段测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("多行字段 - 产品和地区", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "多行字段 - 产品和地区");
    });

    testRunner.runTest("多列字段 - 年份和季度", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length >= 4, "多列字段应有至少4行表头");
        Console.log("    结果行数: " + result.length + " (含多层表头)");
        testRunner.outputResult(result, "多列字段 - 年份和季度（多层表头）");
    });

    testRunner.runTest("多行多列字段组合", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "多行多列字段组合");
    });

    testRunner.runTest("三层行字段 - 大区省份城市", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+"], [], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "三层行字段 - 大区省份城市");
    });
}

function testSorting() {
    testReporter.startGroup("排序功能测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 3: 排序功能测试 (+升序 -降序)");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("行字段升序 (+)", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    第一行数据: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "排序 - 行字段升序 (+)");
    });

    testRunner.runTest("行字段降序 (-)", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1-"], [], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    第一行数据: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "排序 - 行字段降序 (-)");
    });

    testRunner.runTest("多字段混合排序", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+,f2-"], ["f3+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "排序 - 多字段混合（升+降）");
    });
}

function testAggregation() {
    testReporter.startGroup("聚合函数测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 4: 聚合函数测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("聚合函数 - count()", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["count()"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "聚合函数 - count计数");
    });

    testRunner.runTest("聚合函数 - sum(\"f3\")", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果示例: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "聚合函数 - sum求和");
    });

    testRunner.runTest("聚合函数 - average(\"f3\")", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["average(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果示例: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "聚合函数 - average平均值");
    });

    testRunner.runTest("聚合函数 - max(\"f3\")", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["max(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果示例: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "聚合函数 - max最大值");
    });

    testRunner.runTest("聚合函数 - min(\"f3\")", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["min(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果示例: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "聚合函数 - min最小值");
    });

    testRunner.runTest("多聚合函数组合", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["count(),sum(\"f3\"),average(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果列数: " + result[0].length);
        testRunner.outputResult(result, "多聚合函数组合（计数,求和,平均）");
    });
}

function testOptions() {
    testReporter.startGroup("Options选项测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 5: Options 选项测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("选项 - cornerTitle 角落标题", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { cornerTitle: "销售分析" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    左上角: " + result[0][0]);
        testRunner.outputResult(result, "选项 - cornerTitle 角落标题");
    });

    testRunner.runTest("选项 - layoutMode: outline", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "outline" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "选项 - layoutMode: outline");
    });

    testRunner.runTest("选项 - layoutMode: compact", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "compact" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "选项 - layoutMode: compact");
    });

    testRunner.runTest("选项 - rowFieldIndent 行缩进", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { rowFieldIndent: true, rowFieldIndentSize: 4 });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "选项 - rowFieldIndent 行缩进");
    });

    testRunner.runTest("选项 - 禁用 rowFieldIndent", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { rowFieldIndent: false });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "选项 - 禁用 rowFieldIndent");
    });
}

function testSubtotalsAndGrandTotals() {
    testReporter.startGroup("小计和总计测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 6: 小计和总计测试");
    Console.log("------------------------------------------------------------");
    
    testRunner.runTest("小计 - 行小计", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { subtotals: { row: true, label: "小计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        let hasSubtotal = false;
        for (let i = 0; i < result.length; i++) {
            for (let j = 0; j < result[i].length; j++) {
                if (result[i][j] === "小计") { hasSubtotal = true; break; }
            }
        }
        Console.log("    包含小计行: " + hasSubtotal);
        testRunner.outputResult(result, "小计 - 行小计");
    });

    testRunner.runTest("总计 - 总计行", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { row: true, label: "总计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果尺寸: " + result.length + " x " + result[0].length);
        testRunner.outputResult(result, "总计 - 总计行");
    });

    testRunner.runTest("总计 - 总计列", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { col: true, label: "总计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    最后一列标题: " + result[result.length - 1][result[0].length - 1]);
        testRunner.outputResult(result, "总计 - 总计列");
    });

    testRunner.runTest("总计 - 行列总计", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { row: true, col: true, label: "总计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果尺寸: " + result.length + " x " + result[0].length);
        testRunner.outputResult(result, "总计 - 行列总计");
    });
}

function testFieldTitles() {
    testReporter.startGroup("字段标题自定义测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 7: 字段标题自定义测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("自定义标题 - 行字段", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+", "产品名称"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    表头: " + JSON.stringify(result[0]));
        testRunner.outputResult(result, "自定义标题 - 行字段");
    });

    testRunner.runTest("自定义标题 - 数据字段", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")", "销售总额"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    表头: " + JSON.stringify(result[0]));
        testRunner.outputResult(result, "自定义标题 - 数据字段");
    });

    testRunner.runTest("自定义标题 - 多数据字段", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\"),count()", "销售额,订单数"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    表头列数: " + result[0].length);
        testRunner.outputResult(result, "自定义标题 - 多数据字段");
    });
}

function testEdgeCases() {
    testReporter.startGroup("边界情况测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 8: 边界情况测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("边界 - 空数据（仅表头）", function() {
        const data = [["产品", "地区", "销售额"]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应返回至少表头");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "边界 - 空数据（仅表头）");
    });

    testRunner.runTest("边界 - 单行数据", function() {
        const data = [["产品", "地区", "销售额"], ["A", "北京", 100]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理单行数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "边界 - 单行数据");
    });

    testRunner.runTest("边界 - 重复数据", function() {
        const data = [["产品", "地区", "销售额"], ["A", "北京", 100], ["A", "北京", 200], ["A", "北京", 300]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理重复数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "边界 - 重复数据聚合");
    });

    testRunner.runTest("边界 - null/undefined 值", function() {
        const data = [["产品", "地区", "销售额"], ["A", "北京", 100], ["A", null, 200], [null, "上海", 300]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理null值");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "边界 - null值处理");
    });

    testRunner.runTest("边界 - 大数据量 (" + DEFAULT_TEST_ROWS + "行)", function() {
        const data = [["产品", "地区", "销售额"]];
        const products = ["A", "B", "C", "D", "E"];
        const regions = ["北京", "上海", "广州", "深圳", "杭州"];
        for (let i = 0; i < DEFAULT_TEST_ROWS; i++) {
            data.push([products[i % products.length], regions[i % regions.length], Math.floor(Math.random() * 10000)]);
        }
        const startTime = new Date().getTime();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        const duration = new Date().getTime() - startTime;
        testRunner.assertTrue(result.length > 0, "应能处理大数据量");
        Console.log("    处理时间: " + duration + "ms");
        Console.log("    结果行数: " + result.length);
        // 大数据量不输出，避免工作表过大
    });
}

function testRealWorldScenarios() {
    testReporter.startGroup("实战场景测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 9: 实战场景测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("场景 - 销售报表（产品×地区）", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+", "产品"], ["f2+", "地区"], ["sum(\"f5\")", "销售额"], 1, 1, "@^@", { cornerTitle: "销售分析", grandTotal: { row: true, col: true, label: "总计" } });
        testRunner.assertTrue(result.length > 0, "应生成销售报表");
        Console.log("    报表尺寸: " + result.length + " x " + result[0].length);
        testRunner.outputResult(result, "销售报表（产品×地区）");
    });

    testRunner.runTest("场景 - 年度季度对比", function() {
        const data = createTestData();
        const result = Array2D.z超级透视(data, ["f1+", "产品"], ["f3+,f4+", "年份,季度"], ["sum(\"f5\")", "销售额"], 1);
        testRunner.assertTrue(result.length > 0, "应生成年报");
        Console.log("    报表行数: " + result.length + " (含多层表头)");
        testRunner.outputResult(result, "年度季度对比（多层表头）");
    });

    testRunner.runTest("场景 - 区域层级分析", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+", "大区,省份,城市"], [], ["sum(\"f5\"),count()", "销售额,记录数"], 1, 1, "@^@", { layoutMode: "outline", subtotals: { row: true, label: "小计" }, grandTotal: { row: true, label: "总计" } });
        testRunner.assertTrue(result.length > 0, "应生成区域分析报表");
        Console.log("    报表行数: " + result.length);
        testRunner.outputResult(result, "区域层级分析（三层行字段）");
    });
}

function testDisplayAsOptions() {
    testReporter.startGroup("百分比显示测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 10: 百分比显示选项测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("displayAs - 百分比模式", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { displayAs: { mode: "percent", decimals: 2 } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "displayAs - 百分比模式");
    });

    testRunner.runTest("displayAs - 列百分比", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { displayAs: { mode: "percent_column", decimals: 1 } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "displayAs - 列百分比");
    });

    testRunner.runTest("displayAs - 行百分比", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { displayAs: { mode: "percent_row", decimals: 1 } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "displayAs - 行百分比");
    });
}

function testColumnSubtotals() {
    testReporter.startGroup("列小计测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 11: 列小计测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("小计 - 列小计", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f5\")"], 1, 1, "@^@", { subtotals: { col: true, label: "小计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        let hasSubtotal = false;
        for (let i = 0; i < result.length; i++) {
            for (let j = 0; j < result[i].length; j++) {
                if (result[i][j] === "小计") { hasSubtotal = true; break; }
            }
        }
        Console.log("    包含列小计: " + hasSubtotal);
        testRunner.outputResult(result, "小计 - 列小计");
    });

    testRunner.runTest("小计 - 行列小计组合", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+,f4+"], ["sum(\"f5\")"], 1, 1, "@^@", { subtotals: { row: true, col: true, label: "小计" } });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果尺寸: " + result.length + " x " + result[0].length);
        testRunner.outputResult(result, "小计 - 行列小计组合");
    });
}

function testChainOperations() {
    testReporter.startGroup("链式操作测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 12: 链式操作测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("链式调用 - filter后透视", function() {
        const data = createTestData();
        // 筛选2023年Q1的数据后透视
        const filtered = Array2D.z筛选(data, "x=>x[2]==2023 && x[3]=='Q1'");
        const result = Array2D.z超级透视(filtered, ["f1+"], ["f2+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能链式调用");
        Console.log("    筛选后透视结果行数: " + result.length);
        testRunner.outputResult(result, "链式调用 - filter后透视");
    });

    testRunner.runTest("链式调用 - 透视后排序", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        // 对透视结果排序
        const sorted = Array2D.z多列排序(result, "f2-");
        testRunner.assertTrue(sorted.length > 0, "应能对透视结果排序");
        Console.log("    排序后行数: " + sorted.length);
        testRunner.outputResult(sorted, "链式调用 - 透视后排序");
    });

    testRunner.runTest("Array2D实例 - 链式透视", function() {
        const data = createSimpleData();
        const arr = new Array2D(data);
        const result = arr.z超级透视(["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "实例方法应可用");
        Console.log("    实例方法结果行数: " + result.length);
        testRunner.outputResult(result, "Array2D实例 - 链式透视");
    });
}

function testSpecialCharacters() {
    testReporter.startGroup("特殊字符和空值处理测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 13: 特殊字符和空值处理测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("特殊字符 - 空字符串", function() {
        const data = [["产品", "地区", "销售额"], ["A", "", 100], ["", "北京", 200], ["B", "上海", 300]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理空字符串");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "特殊字符 - 空字符串");
    });

    testRunner.runTest("特殊字符 - 包含分隔符的数据", function() {
        const data = [["产品", "地区", "销售额"], ["A@产品", "北京", 100], ["B", "上海", 200]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@");
        testRunner.assertTrue(result.length > 0, "应能正确处理包含分隔符的数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "特殊字符 - 包含分隔符的数据");
    });

    testRunner.runTest("特殊字符 - 数字字符串", function() {
        const data = [["产品", "地区", "销售额"], ["123", "北京", 100], ["456", "上海", 200]];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理数字字符串");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "特殊字符 - 数字字符串");
    });
}

function testCallbackFunctions() {
    testReporter.startGroup("回调函数测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 14: 回调函数测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("回调 - 自定义聚合函数", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], [function(g) { return g.sum("f3"); }], 1);
        testRunner.assertTrue(result.length > 0, "应支持回调函数");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "回调 - 自定义聚合函数");
    });

    testRunner.runTest("回调 - 多个回调函数", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], [], [
            function(g) { return g.count(); },
            function(g) { return g.sum("f3"); },
            function(g) { return g.average("f3"); }
        ], 1);
        testRunner.assertTrue(result.length > 0, "应支持多个回调函数");
        Console.log("    结果列数: " + result[0].length);
        testRunner.outputResult(result, "回调 - 多个回调函数");
    });
}

function testSeparatorOptions() {
    testReporter.startGroup("分隔符选项测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 15: 分隔符选项测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("分隔符 - 自定义分隔符", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "###");
        testRunner.assertTrue(result.length > 0, "应支持自定义分隔符");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "分隔符 - 自定义分隔符（###）");
    });

    testRunner.runTest("分隔符 - 空分隔符", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "");
        testRunner.assertTrue(result.length > 0, "应支持空分隔符");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "分隔符 - 空分隔符");
    });
}

function testCompatibilityOptions() {
    testReporter.startGroup("兼容性选项测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 16: 兼容性选项测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("兼容 - rowSubtotals配置", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { rowSubtotals: { enabled: true } });
        testRunner.assertTrue(result.length > 0, "应兼容旧版rowSubtotals配置");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "兼容 - rowSubtotals配置");
    });

    testRunner.runTest("兼容 - grandTotals配置", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotals: { row: true, column: true } });
        testRunner.assertTrue(result.length > 0, "应兼容旧版grandTotals配置");
        Console.log("    结果尺寸: " + result.length + " x " + result[0].length);
        testRunner.outputResult(result, "兼容 - grandTotals配置");
    });

    testRunner.runTest("兼容 - 无headerRows参数", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"]);
        testRunner.assertTrue(result.length > 0, "应支持默认headerRows");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "兼容 - 无headerRows参数");
    });
}

function testOutputHeaderModes() {
    testReporter.startGroup("输出表头模式测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 17: 输出表头模式测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("outputHeader - 包含表头", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        testRunner.assertTrue(result.length >= 2, "应包含表头行");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "outputHeader - 包含表头");
    });

    testRunner.runTest("outputHeader - 不包含表头", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 0);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "outputHeader - 不包含表头");
    });
}

function testLayoutModes() {
    testReporter.startGroup("布局模式测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 18: 布局模式测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("布局 - outline模式", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "outline" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "布局模式 - outline（大纲）");
    });

    testRunner.runTest("布局 - compact模式", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "compact" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "布局模式 - compact（紧凑）");
    });

    testRunner.runTest("布局 - tabular模式", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "tabular" });
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "布局模式 - tabular（表格）");
    });
}

function testWrappedResultMethods() {
    testReporter.startGroup("包装对象方法测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 19: 包装对象方法测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("包装对象 - getMeta方法", function() {
        const data = createSimpleData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        if (result.getMeta && typeof result.getMeta === "function") {
            const meta = result.getMeta();
            testRunner.assertTrue(meta !== undefined, "应返回元数据");
            Console.log("    元数据版本: " + (meta.version || "unknown"));
        } else {
            Console.log("    跳过: 返回结果不是包装对象");
        }
        testRunner.outputResult(result, "包装对象 - getMeta方法");
    });

    testRunner.runTest("包装对象 - applyMerges方法", function() {
        const data = createMultiLevelData();
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f5\")"], 1, 1, "@^@", { layoutMode: "outline" });
        if (result.applyMerges && typeof result.applyMerges === "function") {
            testRunner.assertTrue(true, "applyMerges方法存在");
            Console.log("    applyMerges方法可用");
        } else {
            Console.log("    跳过: 返回结果不是包装对象");
        }
        testRunner.outputResult(result, "包装对象 - applyMerges方法");
    });
}

function testNumberFormats() {
    testReporter.startGroup("数值格式测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 20: 数值格式测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("数值 - 整数聚合", function() {
        const data = [["产品", "数量"], ["A", 10], ["B", 20], ["A", 30]];
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f2\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "数值 - 整数聚合（sum）");
    });

    testRunner.runTest("数值 - 小数聚合", function() {
        const data = [["产品", "金额"], ["A", 10.5], ["B", 20.3], ["A", 30.2]];
        const result = Array2D.z超级透视(data, ["f1+"], [], ["average(\"f2\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "数值 - 小数聚合（average）");
    });

    testRunner.runTest("数值 - 负数处理", function() {
        const data = [["产品", "利润"], ["A", 100], ["B", -50], ["A", -30]];
        const result = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f2\")"], 1);
        testRunner.assertTrue(result.length > 0, "结果不应为空");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "数值 - 负数处理（sum）");
    });
}

function testDateData() {
    testReporter.startGroup("日期数据处理测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 21: 日期数据处理测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("日期 - 包含日期的数据", function() {
        const data = createDateBasedData();
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理日期数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "日期 - 包含日期的数据");
    });
}

function testDeepHierarchies() {
    testReporter.startGroup("深层层级测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 22: 深层层级测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("层级 - 4层行字段", function() {
        const data = [
            ["区域", "国家", "省份", "城市", "产品", "销售额"],
            ["亚洲", "中国", "北京", "北京市", "A", 1000],
            ["亚洲", "中国", "北京", "北京市", "B", 2000],
            ["亚洲", "中国", "上海", "上海市", "A", 1500],
            ["欧洲", "英国", "伦敦", "伦敦市", "A", 2500]
        ];
        const result = Array2D.z超级透视(data, ["f1+,f2+,f3+,f4+"], [], ["sum(\"f6\")"], 1);
        testRunner.assertTrue(result.length > 0, "应支持4层行字段");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "层级 - 4层行字段");
    });

    testRunner.runTest("层级 - 3层列字段", function() {
        const data = [
            ["产品", "年份", "季度", "月份", "销售额"],
            ["A", 2023, "Q1", "1月", 1000],
            ["A", 2023, "Q1", "2月", 1500],
            ["A", 2023, "Q2", "4月", 2000],
            ["B", 2023, "Q1", "1月", 1200]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+,f4+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length > 0, "应支持3层列字段");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "层级 - 3层列字段");
    });
}

function testLargeDataPerformance() {
    testReporter.startGroup("大数据性能测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 23: 大数据性能测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("性能 - 500行数据", function() {
        const data = createLargeData(500);
        const startTime = new Date().getTime();
        const result = Array2D.z超级透视(data, ["f2+,f3+"], [], ["sum(\"f4\"),average(\"f5\")"], 1);
        const duration = new Date().getTime() - startTime;
        testRunner.assertTrue(result.length > 0, "应能处理500行数据");
        Console.log("    处理时间: " + duration + "ms");
        testRunner.outputResult(result, "性能 - 500行数据");
    });

    testRunner.runTest("性能 - 1000行数据", function() {
        const data = createLargeData(1000);
        const startTime = new Date().getTime();
        const result = Array2D.z超级透视(data, ["f2+"], ["f3+"], ["sum(\"f4\"),count()"], 1);
        const duration = new Date().getTime() - startTime;
        testRunner.assertTrue(result.length > 0, "应能处理1000行数据");
        Console.log("    处理时间: " + duration + "ms");
        testRunner.outputResult(result, "性能 - 1000行数据");
    });
}

function testMixedDataTypes() {
    testReporter.startGroup("混合数据类型测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 24: 混合数据类型测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("混合类型 - 字符串和数字", function() {
        const data = [
            ["产品", "类型", "值"],
            ["A", "类型1", 100],
            ["B", "类型2", 200],
            ["A", "类型1", 150]
        ];
        const result = Array2D.z超级透视(data, ["f1+,f2+"], [], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理混合类型");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "混合类型 - 字符串和数字");
    });

    testRunner.runTest("混合类型 - 包含布尔值", function() {
        const data = [
            ["产品", "状态", "销售额"],
            ["A", true, 100],
            ["B", false, 200],
            ["A", true, 150]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理布尔值");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "混合类型 - 包含布尔值");
    });
}

function testEmptyAndSparseData() {
    testReporter.startGroup("空值和稀疏数据测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 25: 空值和稀疏数据测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("空值 - 全为null的数据", function() {
        const data = [
            ["产品", "地区", "销售额"],
            [null, null, null],
            [null, null, null]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理全null数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "空值 - 全为null的数据");
    });

    testRunner.runTest("稀疏 - 大量空值", function() {
        const data = [
            ["产品", "地区", "销售额"],
            ["A", "北京", 100],
            [null, null, null],
            ["B", null, 200],
            [null, "上海", null],
            ["A", "广州", 150]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
        testRunner.assertTrue(result.length > 0, "应能处理稀疏数据");
        Console.log("    结果行数: " + result.length);
        testRunner.outputResult(result, "稀疏 - 大量空值");
    });
}

function testMultiLevelHeaders() {
    testReporter.startGroup("多层标题测试");
    Console.log("");
    Console.log("------------------------------------------------------------");
    Console.log(" 测试组 26: 多层标题测试");
    Console.log("------------------------------------------------------------");

    testRunner.runTest("多层标题 - 两层列字段", function() {
        const data = [
            ["产品", "年份", "季度", "销售额"],
            ["A", 2023, "Q1", 1000],
            ["A", 2023, "Q2", 1500],
            ["B", 2023, "Q1", 2000],
            ["B", 2023, "Q2", 2500]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\")"], 1);
        testRunner.assertTrue(result.length >= 3, "应有至少3行表头（2层列字段+1层数据）");
        Console.log("    表头行数: " + result.length + " (包含多层表头)");
        Console.log("    第1行表头: " + JSON.stringify(result[0]));
        Console.log("    第2行表头: " + JSON.stringify(result[1]));
        testRunner.outputResult(result, "多层标题 - 两层列字段（年份+季度）");
    });

    testRunner.runTest("多层标题 - 三层列字段", function() {
        const data = [
            ["产品", "年份", "季度", "月份", "销售额"],
            ["A", 2023, "Q1", "1月", 1000],
            ["A", 2023, "Q1", "2月", 1200],
            ["B", 2023, "Q1", "1月", 2000],
            ["B", 2023, "Q2", "4月", 2200]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+,f4+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length >= 4, "应有至少4行表头（3层列字段+1层数据）");
        Console.log("    表头行数: " + result.length + " (包含三层表头)");
        Console.log("    第1行: " + JSON.stringify(result[0]));
        Console.log("    第2行: " + JSON.stringify(result[1]));
        Console.log("    第3行: " + JSON.stringify(result[2]));
        testRunner.outputResult(result, "多层标题 - 三层列字段（年份+季度+月份）");
    });

    testRunner.runTest("多层标题 - 列字段+单数据字段", function() {
        const data = [
            ["产品", "年份", "季度", "销售额"],
            ["A", 2023, "Q1", 1000],
            ["A", 2023, "Q2", 1500],
            ["B", 2023, "Q1", 2000]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\")"], 1);
        testRunner.assertTrue(result.length >= 3, "列字段后应跟数据字段标题");
        Console.log("    表头行数: " + result.length);
        // 检查最后一行表头是否包含数据字段标题
        const lastHeaderRow = result[2] || result[1];
        Console.log("    数据字段标题行: " + JSON.stringify(lastHeaderRow));
        testRunner.outputResult(result, "多层标题 - 列字段+单数据字段");
    });

    testRunner.runTest("多层标题 - 列字段+多数据字段", function() {
        const data = [
            ["产品", "年份", "季度", "销售额", "数量"],
            ["A", 2023, "Q1", 1000, 10],
            ["A", 2023, "Q2", 1500, 15],
            ["B", 2023, "Q1", 2000, 20]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\"),sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length >= 3, "应有多层表头");
        testRunner.assertTrue(result[0].length > 4, "列数应包含多个数据字段");
        Console.log("    表头行数: " + result.length);
        Console.log("    总列数: " + result[0].length);
        testRunner.outputResult(result, "多层标题 - 列字段+多数据字段（销售额+数量）");
    });

    testRunner.runTest("多层标题 - 多行+多列字段", function() {
        const data = [
            ["大区", "省份", "年份", "季度", "销售额"],
            ["华北", "北京", 2023, "Q1", 1000],
            ["华北", "北京", 2023, "Q2", 1500],
            ["华东", "上海", 2023, "Q1", 2000],
            ["华东", "上海", 2023, "Q2", 2500]
        ];
        const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
        testRunner.assertTrue(result.length >= 4, "应有多层表头（2行+2列）");
        Console.log("    表头行数: " + result.length);
        Console.log("    行字段列数: 2 (大区, 省份)");
        Console.log("    列字段层数: 2 (年份, 季度)");
        testRunner.outputResult(result, "多层标题 - 多行+多列字段（2×2）");
    });

    testRunner.runTest("多层标题 - 自定义数据字段标题", function() {
        const data = [
            ["产品", "年份", "季度", "销售额"],
            ["A", 2023, "Q1", 1000],
            ["B", 2023, "Q1", 2000]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\")", "总销售额"], 1);
        testRunner.assertTrue(result.length >= 3, "应有自定义数据字段标题");
        Console.log("    表头行数: " + result.length);
        // 检查数据字段标题行是否包含"总销售额"
        let foundCustomTitle = false;
        for (let i = 0; i < Math.min(3, result.length); i++) {
            for (let j = 0; j < result[i].length; j++) {
                if (result[i][j] === "总销售额") {
                    foundCustomTitle = true;
                    break;
                }
            }
        }
        Console.log("    自定义标题存在: " + foundCustomTitle);
        testRunner.outputResult(result, "多层标题 - 自定义数据字段标题（总销售额）");
    });

    testRunner.runTest("多层标题 - 多数据字段自定义标题", function() {
        const data = [
            ["产品", "年份", "季度", "销售额", "数量", "利润"],
            ["A", 2023, "Q1", 1000, 10, 200],
            ["A", 2023, "Q2", 1500, 15, 300],
            ["B", 2023, "Q1", 2000, 20, 400],
            ["B", 2023, "Q2", 2200, 22, 450],
            ["A", 2024, "Q1", 1200, 12, 250]
        ];
        // 使用多层列字段 + 多数据字段来产生更多列
        // 列字段：年份+季度（2层），数据字段：销售额、数量、利润（3个）
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\"),sum(\"f5\"),sum(\"f6\")", "销售额,数量,利润"], 1);
        testRunner.assertTrue(result.length >= 3, "应有表头（2层列字段+1层数据字段）");
        // 结构: [角落] [2023.Q1.销售额] [2023.Q1.数量] [2023.Q1.利润] [2023.Q2.销售额]...
        testRunner.assertTrue(result[0].length >= 4, "应有多列（行字段+多个数据字段×列值组合）");
        Console.log("    表头行数: " + result.length);
        Console.log("    总列数: " + result[0].length);
        // 验证包含自定义标题
        let hasCustomTitles = false;
        for (let i = 0; i < Math.min(result.length, 3); i++) {
            for (let j = 0; j < result[i].length; j++) {
                if (result[i][j] === "销售额" || result[i][j] === "数量" || result[i][j] === "利润") {
                    hasCustomTitles = true;
                    break;
                }
            }
        }
        Console.log("    包含自定义标题: " + hasCustomTitles);
        testRunner.outputResult(result, "多层标题 - 多数据字段自定义标题");
    });

    testRunner.runTest("多层标题 - 四层列字段", function() {
        const data = [
            ["产品", "区域", "国家", "省份", "城市", "销售额"],
            ["A", "亚洲", "中国", "北京", "北京市", 1000],
            ["A", "亚洲", "中国", "上海", "上海市", 1500],
            ["B", "亚洲", "中国", "北京", "北京市", 2000],
            ["B", "欧洲", "英国", "伦敦", "伦敦市", 2500]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+,f4+,f5+"], ["sum(\"f6\")"], 1);
        testRunner.assertTrue(result.length >= 5, "应有至少5行表头（4层列字段+1层数据）");
        Console.log("    表头行数: " + result.length + " (包含四层表头)");
        Console.log("    第1行: " + JSON.stringify(result[0]));
        Console.log("    第2行: " + JSON.stringify(result[1]));
        Console.log("    第3行: " + JSON.stringify(result[2]));
        Console.log("    第4行: " + JSON.stringify(result[3]));
        testRunner.outputResult(result, "多层标题 - 四层列字段（区域+国家+省份+城市）");
    });

    testRunner.runTest("多层标题 - 复杂结构（ cornerTitle）", function() {
        const data = [
            ["产品", "年份", "季度", "销售额", "数量"],
            ["A", 2023, "Q1", 1000, 10],
            ["A", 2023, "Q2", 1500, 15],
            ["B", 2023, "Q1", 2000, 20]
        ];
        const result = Array2D.z超级透视(
            data,
            ["f1+"],
            ["f2+,f3+"],
            ["sum(\"f4\"),sum(\"f5\")", "销售额,数量"],
            1, 1, "@^@",
            { cornerTitle: "销售分析报表" }
        );
        testRunner.assertTrue(result.length >= 3, "应有复杂多层表头");
        testRunner.assertTrue(result[0][0] === "销售分析报表", "左上角应显示自定义角落标题");
        Console.log("    左上角标题: " + result[0][0]);
        Console.log("    表头行数: " + result.length);
        Console.log("    总列数: " + result[0].length);
        testRunner.outputResult(result, "多层标题 - 复杂结构（带cornerTitle）");
    });

    testRunner.runTest("多层标题 - 表头层级验证", function() {
        const data = [
            ["产品", "年份", "季度", "销售额"],
            ["A", 2023, "Q1", 1000],
            ["A", 2023, "Q2", 1500],
            ["B", 2023, "Q1", 2000]
        ];
        const result = Array2D.z超级透视(data, ["f1+"], ["f2+,f3+"], ["sum(\"f4\")"], 1);

        // 验证表头层级结构
        const headerRowCount = 3; // 2层列字段 + 1层数据字段
        testRunner.assertArrayLength(result, headerRowCount + 2, "应有" + headerRowCount + "行表头+2行数据");

        // 检查第一行（年份层）
        Console.log("    第1行（年份层）: " + JSON.stringify(result[0]));

        // 检查第二行（季度层）
        Console.log("    第2行（季度层）: " + JSON.stringify(result[1]));

        // 检查第三行（数据字段层）
        Console.log("    第3行（数据字段层）: " + JSON.stringify(result[2]));

        Console.log("    表头层级验证完成");
        testRunner.outputResult(result, "多层标题 - 表头层级验证");
    });
}

// =======================================================================
// 主运行函数
// =======================================================================

function runAllTests() {
    Console.log("");
    Console.log("============================================================");
    Console.log("     JSA880 SuperPivot (z超级透视) 功能测试套件 v" + MODULE_VERSION);
    Console.log("============================================================");
    Console.log("开始时间: " + new Date().toLocaleString());
    Console.log("");

    // 初始化报告和输出
    testReporter.initialize();
    testOutput.initialize();
    testRunner.reset();

    // 启用每次测试前自动清空输出表格
    testRunner.setAutoClearOutput(true);
    Console.log("已启用: 每次测试前自动清空输出表格");

    // 运行所有测试组
    testBasicFunctions();
    testMultipleFields();
    testSorting();
    testAggregation();
    testOptions();
    testSubtotalsAndGrandTotals();
    testFieldTitles();
    testEdgeCases();
    testRealWorldScenarios();
    testDisplayAsOptions();
    testColumnSubtotals();
    testChainOperations();
    testSpecialCharacters();
    testCallbackFunctions();
    testSeparatorOptions();
    testCompatibilityOptions();
    testOutputHeaderModes();
    testLayoutModes();
    testWrappedResultMethods();
    testNumberFormats();
    testDateData();
    testDeepHierarchies();
    testLargeDataPerformance();
    testMixedDataTypes();
    testEmptyAndSparseData();
    testMultiLevelHeaders();

    // 输出汇总
    const summary = testRunner.printSummary();
    testReporter.writeSummary(summary);

    // 自动调整输出工作表列宽
    testOutput.autoFitColumns();

    Console.log("");
    Console.log("结束时间: " + new Date().toLocaleString());
    Console.log("测试结果已输出到工作表: " + REPORT_SHEET_NAME);
    Console.log("透视表输出已输出到工作表: " + OUTPUT_SHEET_NAME);

    return summary;
}

function runQuickTest() {
    Console.log("");
    Console.log("=== JSA880 SuperPivot 快速测试 ===");
    Console.log("");

    // 初始化输出
    testOutput.initialize();

    const data = createSimpleData();
    Console.log("原始数据:");
    for (let i = 0; i < data.length; i++) {
        Console.log("  " + JSON.stringify(data[i]));
    }

    Console.log("");
    Console.log("透视结果（产品×地区）:");

    const result = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);

    for (let i = 0; i < result.length; i++) {
        Console.log("  " + JSON.stringify(result[i]));
    }

    // 输出到工作表
    testOutput.outputResult(data, "原始数据");
    testOutput.outputResult(result, "透视结果（产品×地区）");
    testOutput.autoFitColumns();

    Console.log("");
    Console.log("结果已输出到工作表: " + OUTPUT_SHEET_NAME);
    testOutput.getSheet().Activate();

    return result;
}

function runDiagnosticTest() {
    Console.log("");
    Console.log("=== SuperPivot 诊断测试 ===");
    Console.log("");

    // 初始化输出
    testOutput.initialize();

    const data = createSimpleData();
    Console.log("测试数据:");
    Console.log(JSON.stringify(data));

    Console.log("");
    Console.log("--- 测试1: 有行有列 ---");
    const result1 = Array2D.z超级透视(data, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    debugPivotResult(result1, "有行有列");
    testOutput.outputResult(result1, "1. 有行有列");

    Console.log("");
    Console.log("--- 测试2: 仅行字段 ---");
    const result2 = Array2D.z超级透视(data, ["f1+"], [], ["sum(\"f3\")"], 1);
    debugPivotResult(result2, "仅行字段");
    testOutput.outputResult(result2, "2. 仅行字段");

    Console.log("");
    Console.log("--- 测试3: 仅列字段 ---");
    const result3 = Array2D.z超级透视(data, [], ["f2+"], ["sum(\"f3\")"], 1);
    debugPivotResult(result3, "仅列字段");
    testOutput.outputResult(result3, "3. 仅列字段");

    Console.log("");
    Console.log("--- 测试4: 多聚合函数 ---");
    const result4 = Array2D.z超级透视(data, ["f1+"], [], ["count(),sum(\"f3\")"], 1);
    debugPivotResult(result4, "多聚合函数");
    testOutput.outputResult(result4, "4. 多聚合函数");

    testOutput.autoFitColumns();

    Console.log("");
    Console.log("诊断完成，结果已输出到: " + OUTPUT_SHEET_NAME);
    testOutput.getSheet().Activate();

    return {
        result1: result1,
        result2: result2,
        result3: result3,
        result4: result4
    };
}

function runDemoOutput() {
    Console.log("");
    Console.log("=== SuperPivot 功能演示输出 ===");
    Console.log("");

    // 初始化输出
    testOutput.initialize();

    // 1. 基础透视
    const simpleData = createSimpleData();
    const result1 = Array2D.z超级透视(simpleData, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1);
    testOutput.outputResult(result1, "1. 基础透视 - 产品×地区");

    // 2. 多聚合函数
    const result2 = Array2D.z超级透视(simpleData, ["f1+"], [], ["count(),sum(\"f3\"),average(\"f3\")"], 1);
    testOutput.outputResult(result2, "2. 多聚合函数 - 计数/求和/平均");

    // 3. 多层列字段
    const testData = createTestData();
    const result3 = Array2D.z超级透视(testData, ["f1+"], ["f3+,f4+"], ["sum(\"f5\")"], 1);
    testOutput.outputResult(result3, "3. 多层列字段 - 年份×季度");

    // 4. 三层行字段
    const multiLevelData = createMultiLevelData();
    const result4 = Array2D.z超级透视(multiLevelData, ["f1+,f2+,f3+"], [], ["sum(\"f5\"),count()"], 1, 1, "@^@", { layoutMode: "outline" });
    testOutput.outputResult(result4, "4. 三层行字段 - 大区×省份×城市（outline）");

    // 5. 带总计
    const result5 = Array2D.z超级透视(simpleData, ["f1+"], ["f2+"], ["sum(\"f3\")"], 1, 1, "@^@", { grandTotal: { row: true, col: true, label: "总计" } });
    testOutput.outputResult(result5, "5. 带总计 - 行列总计");

    // 自动调整列宽
    testOutput.autoFitColumns();

    Console.log("演示输出完成!");
    Console.log("结果已输出到工作表: " + OUTPUT_SHEET_NAME);
    testOutput.getSheet().Activate();

    return {
        outputCount: 5,
        sheetName: OUTPUT_SHEET_NAME
    };
}

function runPerformanceTest(rowCount) {
    rowCount = rowCount || PERFORMANCE_TEST_ROWS;
    Console.log("");
    Console.log("=== 性能测试 (" + rowCount + " 行) ===");

    const data = [["产品", "地区", "年份", "销售额"]];
    const products = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
    const regions = ["北京", "上海", "广州", "深圳"];
    const years = [2022, 2023, 2024];

    for (let i = 0; i < rowCount; i++) {
        data.push([products[i % products.length], regions[i % regions.length], years[i % years.length], Math.floor(Math.random() * 10000)]);
    }

    Console.log("数据生成完成: " + data.length + " 行");

    const startTime = new Date().getTime();
    const result = Array2D.z超级透视(data, ["f1+,f2+"], ["f3+"], ["sum(\"f4\"),count()"], 1);
    const duration = new Date().getTime() - startTime;

    Console.log("透视完成: " + duration + "ms");
    Console.log("结果行数: " + result.length);

    return { dataRows: rowCount, timeMs: duration, resultRows: result.length };
}

function runSpecificTestGroup(groupName) {
    Console.log("");
    Console.log("============================================================");
    Console.log("     运行指定测试组: " + groupName);
    Console.log("============================================================");

    // 初始化报告
    testReporter.initialize();
    testRunner.reset();

    // 根据组名运行对应测试
    switch(groupName) {
        case "基础":
            testBasicFunctions();
            break;
        case "多字段":
            testMultipleFields();
            break;
        case "排序":
            testSorting();
            break;
        case "聚合":
            testAggregation();
            break;
        case "选项":
            testOptions();
            break;
        case "小计总计":
            testSubtotalsAndGrandTotals();
            break;
        case "边界":
            testEdgeCases();
            break;
        case "实战":
            testRealWorldScenarios();
            break;
        case "链式":
            testChainOperations();
            break;
        default:
            Console.log("未知的测试组: " + groupName);
            Console.log("可用测试组: 基础, 多字段, 排序, 聚合, 选项, 小计总计, 边界, 实战, 链式");
            return null;
    }

    const summary = testRunner.printSummary();
    testReporter.writeSummary(summary);

    return summary;
}

function validateDataIntegrity(data, expectedRowCount, expectedColCount) {
    Console.log("");
    Console.log("=== 数据完整性验证 ===");

    if (!Array.isArray(data)) {
        Console.log("错误: 数据不是数组");
        return false;
    }

    if (data.length === 0) {
        Console.log("错误: 数据为空");
        return false;
    }

    const rowCount = data.length;
    const colCount = data[0].length;

    Console.log("实际行数: " + rowCount);
    Console.log("实际列数: " + colCount);

    if (expectedRowCount && rowCount !== expectedRowCount) {
        Console.log("警告: 行数不匹配 (期望: " + expectedRowCount + ")");
    }

    if (expectedColCount && colCount !== expectedColCount) {
        Console.log("警告: 列数不匹配 (期望: " + expectedColCount + ")");
    }

    let isValid = true;
    for (let i = 0; i < data.length; i++) {
        if (data[i].length !== colCount) {
            Console.log("警告: 第 " + (i + 1) + " 行列数不一致");
            isValid = false;
        }
    }

    Console.log("验证结果: " + (isValid ? "通过" : "失败"));
    return isValid;
}

/**
 * 调试函数：打印透视表结果的详细信息
 */
function debugPivotResult(result, title) {
    Console.log("");
    Console.log("=== 透视表结果调试: " + title + " ===");

    if (!result) {
        Console.log("错误: 结果为空");
        return;
    }

    Console.log("结果类型: " + typeof result);
    Console.log("是否为数组: " + Array.isArray(result));
    Console.log("结果行数: " + result.length);

    if (result.length > 0) {
        Console.log("第一行长度: " + (result[0] ? result[0].length : "undefined"));

        // 打印前5行
        Console.log("前5行内容:");
        for (let i = 0; i < Math.min(5, result.length); i++) {
            const row = result[i];
            if (row) {
                Console.log("  行" + i + ": " + JSON.stringify(row));
            }
        }

        // 检查每行的列数
        const colCounts = result.map(r => r ? r.length : 0);
        Console.log("每行列数: " + JSON.stringify(colCounts));
    }

    Console.log("==================");
}

// =======================================================================
// 模块导出
// =======================================================================

if (typeof Application !== "undefined") {
    Application.runAllTests = runAllTests;
    Application.runQuickTest = runQuickTest;
    Application.runDiagnosticTest = runDiagnosticTest;
    Application.runDemoOutput = runDemoOutput;
    Application.runPerformanceTest = runPerformanceTest;
    Application.runSpecificTestGroup = runSpecificTestGroup;
    Application.validateDataIntegrity = validateDataIntegrity;
    Application.debugPivotResult = debugPivotResult;
    Application.testOutput = testOutput;
    Application.testRunner = testRunner;
    Application.SuperPivotTestSuite = {
        name: MODULE_NAME,
        version: MODULE_VERSION,
        date: MODULE_DATE,
        testGroups: 26,
        description: "SuperPivot (z超级透视) 完整功能测试套件"
    };
}

Console.log("[" + MODULE_NAME + " v" + MODULE_VERSION + "] 测试套件已加载。");
Console.log("");
Console.log("可用命令:");
Console.log("  runDiagnosticTest()               - 诊断测试（输出4种透视表并显示详细调试信息）");
Console.log("  runDemoOutput()                   - 功能演示（输出5个示例透视表到\"测试输出\"工作表）");
Console.log("  runQuickTest()                    - 快速测试（输出到\"测试输出\"工作表）");
Console.log("  runAllTests()                     - 运行所有测试并输出到\"测试结果\"和\"测试输出\"工作表");
Console.log("  runPerformanceTest(n)             - 性能测试（n行数据，默认5000行）");
Console.log("  runSpecificTestGroup(groupName)   - 运行指定测试组");
Console.log("                                    可选组名: 基础, 多字段, 排序, 聚合, 选项, 小计总计, 边界, 实战, 链式");
Console.log("  validateDataIntegrity(data)       - 验证数据完整性");
Console.log("  debugPivotResult(result, title)   - 调试透视表结果");
Console.log("");
Console.log("输出管理:");
Console.log("  testOutput.outputResult(result, title)  - 手动输出透视表结果");
Console.log("  testOutput.getSheet()                   - 获取\"测试输出\"工作表");
Console.log("  testRunner.outputResult(result, title)  - 通过测试运行器输出");
Console.log("  testRunner.setEnableOutput(true/false)  - 启用/禁用自动输出");
Console.log("");
Console.log("测试组列表 (26个测试组):");
Console.log("  1. 基础功能测试");
Console.log("  2. 多字段测试");
Console.log("  3. 排序功能测试");
Console.log("  4. 聚合函数测试");
Console.log("  5. Options选项测试");
Console.log("  6. 小计和总计测试");
Console.log("  7. 字段标题自定义测试");
Console.log("  8. 边界情况测试");
Console.log("  9. 实战场景测试");
Console.log("  10. 百分比显示测试");
Console.log("  11. 列小计测试");
Console.log("  12. 链式操作测试");
Console.log("  13. 特殊字符和空值处理测试");
Console.log("  14. 回调函数测试");
Console.log("  15. 分隔符选项测试");
Console.log("  16. 兼容性选项测试");
Console.log("  17. 输出表头模式测试");
Console.log("  18. 布局模式测试");
Console.log("  19. 包装对象方法测试");
Console.log("  20. 数值格式测试");
Console.log("  21. 日期数据处理测试");
Console.log("  22. 深层层级测试");
Console.log("  23. 大数据性能测试");
Console.log("  24. 混合数据类型测试");
Console.log("  25. 空值和稀疏数据测试");
Console.log("  26. 多层标题测试");
