/**
 * @module mInvoiceExtractor
 * @description 抽取发票数据，使用Array2D.selectCols选取指定列
 * @version 2.0.0
 * @date 2026-05-06
 */
function extractInvoiceData() {
    var filePath = "/Users/daidai193/Downloads/阳山中核-第2批光伏发票_2026-05-06_143120.xlsm";

    Console.log("=== 开始抽取发票数据 ===");

    // 打开工作簿
    var wb = Application.Workbooks.Open(filePath);
    var ws = wb.Sheets("汇总");

    // 使用$.maxArray读取全部数据
    var data = $.maxArray(ws.Range("A1"));

    Console.log("数据行数: " + data.length);

    // 获取表头（第1行）
    var headers = data[0];

    // 目标列名
    var targetCols = ["序号", "发票号码", "开票日期", "货物名称", "单位", "数量", "单价", "金额", "税额", "税率", "发票金额（价税合计）"];

    // 查找目标列的索引
    var colIndexes = [];

    for (var t = 0; t < targetCols.length; t++) {
        var targetName = targetCols[t];
        for (var j = 0; j < headers.length; j++) {
            var h = String(headers[j] || "").trim();
            if (h === targetName) {
                colIndexes.push(j);
                Console.log("找到列: " + targetName + " -> " + j);
                break;
            }
        }
    }

    Console.log("选中列数: " + colIndexes.length);

    // 使用Array2D.skip跳过表头
    var dataRows = Array2D.skip(data, 1);

    // 使用Array2D.selectCols选取需要的列
    var selectedData = Array2D.selectCols(dataRows, colIndexes);

    // 使用Array2D.map处理每一行
    var result = Array2D.map(selectedData, function(row, index) {
        var amount = parseFloat(row[7]) || 0;  // 金额列
        var tax = parseFloat(row[8]) || 0;     // 税额列
        return [index + 1, row[1], row[2], row[3], row[4], row[5], row[6], amount, tax, row[9], row[10]];
    });

    // 输出表头
    var outputHeaders = [targetCols];
    outputHeaders.toRange(ws.Range("I1"));

    // 输出数据
    result.toRange(ws.Range("I2"));

    // 计算并输出总额
    var totalAmount = 0;
    var totalTax = 0;
    for (var i = 0; i < result.length; i++) {
        totalAmount = totalAmount + (result[i][7] || 0);
        totalTax = totalTax + (result[i][8] || 0);
    }

    var totalRow = [["", "", "", "", "", "", "", "合计:", totalTax]];
    totalRow.toRange(ws.Range("I" + (result.length + 2)));

    Console.log("=== 完成 ===");
    Console.log("共处理 " + result.length + " 条记录");
    Console.log("金额合计: " + totalAmount.toFixed(2));
    Console.log("税额合计: " + totalTax.toFixed(2));

    wb.Close(true);
    return result;
}