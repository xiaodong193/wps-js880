/**
 * =======================================================================
 * superPivot v3.9.0 一键测试
 * =======================================================================
 * 
 * 最简单的方式验证 superPivot 是否正常工作
 * 
 * 使用方法：
 *   1. 按 Alt+F11 打开宏编辑器
 *   2. 粘贴此代码
 *   3. 运行 一键测试()
 * =======================================================================
 */

function 一键测试() {
    Console.log("正在测试 superPivot v3.9.0...");
    Console.log("");
    
    // 准备简单测试数据
    var data = [
        ['产品', '年份', '销售额'],
        ['A', '2023', 100],
        ['A', '2024', 200],
        ['B', '2023', 300],
        ['B', '2024', 400]
    ];
    
    var allPassed = true;
    
    // 测试1: 基础功能
    try {
        var result = Array2D.z超级透视(
            data,
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '销售额']
        );
        
        if (result && result.length > 0) {
            Console.log("✅ 基础透视 - 通过");
        } else {
            Console.log("❌ 基础透视 - 失败：结果为空");
            allPassed = false;
        }
    } catch (e) {
        Console.log("❌ 基础透视 - 错误: " + e.message);
        allPassed = false;
    }
    
    // 测试2: 小计功能
    try {
        var result = Array2D.z超级透视(
            data,
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '销售额'],
            1, 1, '@^@',
            {
                rowSubtotals: { enabled: true, label: '小计' },
                grandTotals: { row: true, column: true }
            }
        );
        Console.log("✅ 小计总计 - 通过");
    } catch (e) {
        Console.log("❌ 小计总计 - 错误: " + e.message);
        allPassed = false;
    }
    
    // 测试3: 百分比
    try {
        var result = Array2D.z超级透视(
            data,
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '占比'],
            1, 1, '@^@',
            {
                displayAs: { mode: 'percentOfGrandTotal', decimals: 2 }
            }
        );
        Console.log("✅ 百分比显示 - 通过");
    } catch (e) {
        Console.log("❌ 百分比显示 - 错误: " + e.message);
        allPassed = false;
    }
    
    // 测试4: 角标题
    try {
        var result = Array2D.z超级透视(
            data,
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '销售额'],
            1, 1, '@^@',
            {
                cornerTitle: '测试表'
            }
        );
        var meta = result.getMeta();
        if (meta.options.cornerTitle === '测试表') {
            Console.log("✅ 角标题功能 - 通过");
        } else {
            Console.log("⚠️ 角标题功能 - 可能异常");
        }
    } catch (e) {
        Console.log("❌ 角标题功能 - 错误: " + e.message);
        allPassed = false;
    }
    
    // 测试5: 输出到工作表
    try {
        var result = Array2D.z超级透视(
            data,
            ['f1', '产品'],
            ['f2', '年份'],
            ['sum("f3")', '销售额']
        );
        
        var ws;
        try {
            ws = Application.ActiveWorkbook.Worksheets("superPivot测试");
            ws.Cells.Clear();
        } catch (e) {
            ws = Application.ActiveWorkbook.Worksheets.Add();
            ws.Name = "superPivot测试";
        }
        
        result.toRange("superPivot测试!A1");
        Console.log("✅ 输出到工作表 - 通过");
        Console.log("   请在【superPivot测试】工作表中查看结果");
    } catch (e) {
        Console.log("❌ 输出到工作表 - 错误: " + e.message);
        allPassed = false;
    }
    
    Console.log("");
    if (allPassed) {
        Console.log("🎉 所有测试通过！superPivot v3.9.0 工作正常");
    } else {
        Console.log("⚠️ 部分测试失败，请检查错误信息");
    }
    
    return allPassed;
}

// 运行测试
一键测试();
