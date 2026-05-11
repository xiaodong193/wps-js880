/**
 * ========================================================================
 * mergeCells 功能演示
 * ========================================================================
 *
 * 此文件演示 JSA880 框架中 mergeCells 的各种用法
 *
 * 运行环境: WPS Office + JSA880 框架
 * 运行方式: 在 WPS 宏编辑器中运行此脚本
 * ========================================================================
 */

function main() {
    Console.log("=== mergeCells 功能演示 ===\n");

    // 清空演示区域
    demo_clearArea();

    // 演示1: 简单合并
    demo1_simpleMerge();

    // 演示2: 智能合并透视表标题
    demo2_smartMerge();

    // 演示3: 自定义选项合并
    demo3_customOptions();

    Console.log("\n=== 演示完成！===");
}

// ==================== 演示1: 简单合并 ====================
/**
 * 最基础的合并单元格用法
 */
function demo1_simpleMerge() {
    Console.log("\n--- 演示1: 简单合并 ---");

    // 写入测试数据
    Range("A2").Value2 = "这是一个简单合并示例";
    Range("A2:D2").Interior.Color = 0x4472C4;  // 蓝色背景
    Range("A2:D2").Font.Color = 0xFFFFFF;     // 白色文字

    // 执行合并
    $.mergeCells("A2:D2");

    Console.log("✓ 已合并 A2:D2");

    // 再来一个
    Range("A4").Value2 = "简单合并演示 2";
    Range("A4:C4").Interior.Color = 0x70AD47;  // 绿色背景
    Range("A4:C4").Font.Color = 0xFFFFFF;

    $.z合并单元格("A4:C4");  // 使用中文方法名

    Console.log("✓ 已合并 A4:C4 (使用中文方法名)");
}

// ==================== 演示2: 智能合并透视表标题 ====================
/**
 * 智能识别并合并透视表的多层表头
 */
function demo2_smartMerge() {
    Console.log("\n--- 演示2: 智能合并透视表标题 ---");

    // 创建一个透视表样式的数据
    var pivotData = [
        ["", "", "2023", "", "", "", "2024", "", "", ""],
        ["产品", "国家", "Q1", "Q2", "Q3", "Q4", "Q1", "Q2", "Q3", "Q4"],
        ["Product1", "中国", 100, 120, 130, 140, 150, 160, 170, 180],
        ["Product1", "美国", 90, 110, 120, 130, 140, 150, 160, 170],
        ["Product2", "中国", 80, 100, 110, 120, 130, 140, 150, 160],
        ["Product2", "美国", 70, 90, 100, 110, 120, 130, 140, 150],
    ];

    // 写入数据 (从 A8 开始)
    var startRow = 8;
    var startCol = 1;
    for (var i = 0; i < pivotData.length; i++) {
        for (var j = 0; j < pivotData[i].length; j++) {
            Cells(startRow + i, startCol + j).Value2 = pivotData[i][j];
        }
    }

    // 添加边框
    var rng = Range("A8:K13");
    rng.Borders.LineStyle = 1;
    rng.Borders.Weight = 2;

    // 添加表头样式
    Range("A8:K9").Interior.Color = 0xD9E1F2;
    Range("A8:K9").Font.Bold = true;

    // 执行智能合并
    $.mergeCells("A8:K9", "cm");

    Console.log("✓ 已智能合并透视表标题 A8:K9");
    Console.log("  - 第一行: 2023 和 2024 各自合并");
    Console.log("  - 第二行: Q1-Q4 不合并（值不同）");
}

// ==================== 演示3: 自定义选项合并 ====================
/**
 * 使用自定义选项进行更精细的合并控制
 */
function demo3_customOptions() {
    Console.log("\n--- 演示3: 自定义选项合并 ---");

    // 创建一个更复杂的透视表
    var complexData = [
        ["", "", "年份", "", "", "", "", "", ""],
        ["", "", "2023", "", "", "2024", "", "", ""],
        ["类别", "产品", "Q1", "Q2", "Q3", "Q1", "Q2", "Q3", "销量"],
        ["电子", "手机", 500, 600, 700, 800, 900, 1000, 4500],
        ["电子", "电脑", 300, 350, 400, 450, 500, 550, 2550],
        ["电子", "平板", 200, 250, 300, 350, 400, 450, 1950],
        ["家电", "冰箱", 150, 180, 210, 240, 270, 300, 1350],
        ["家电", "洗衣机", 120, 150, 180, 210, 240, 270, 1170],
    ];

    // 写入数据 (从 A16 开始)
    var startRow = 16;
    var startCol = 1;
    for (var i = 0; i < complexData.length; i++) {
        for (var j = 0; j < complexData[i].length; j++) {
            Cells(startRow + i, startCol + j).Value2 = complexData[i][j];
        }
    }

    // 添加样式
    var rng = Range("A16:I23");
    rng.Borders.LineStyle = 1;

    // 标题行样式
    Range("A16:I19").Interior.Color = 0xE7E6E6;
    Range("A16:I19").Font.Bold = true;

    // 执行智能合并（带选项）
    $.mergeCells("A16:I19", "cm", {
        dataRowStart: 3,      // 数据从第4行开始
        skipLastRow: true,    // 跳过最后一行（数据标题行）
        titleRowCount: 3      // 标题有3行
    });

    Console.log("✓ 已使用自定义选项合并 A16:I19");
    Console.log("  - dataRowStart: 3 (数据从第4行开始)");
    Console.log("  - skipLastRow: true (跳过数据标题行)");
    Console.log("  - titleRowCount: 3 (3行标题)");

    // 合并行标题（纵向合并）- 类别列
    Console.log("\n✓ 纵向合并: 相同类别自动合并");
}

// ==================== 辅助函数 ====================

/**
 * 清空演示区域
 */
function demo_clearArea() {
    // 清空 A 列到 K 列的前30行
    Range("A1:K30").Clear();
    Range("A1").Value2 = "mergeCells 功能演示";
    Range("A1").Font.Size = 16;
    Range("A1").Font.Bold = true;
}
