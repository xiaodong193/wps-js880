/**
 * 合同余额合并模块
 * 功能：合并3个合同工作表到"汇总合同"工作表
 *
 * 使用方法：
 * 1. 在WPS中打开 沧州南大港项目3个合同实时余额_WPS版.xlsm
 * 2. 按 Alt+F11 打开代码编辑器
 * 3. 新建模块，粘贴此代码
 * 4. F5 运行此函数
 */

function 合并合同表() {
    var workbook = Application.Workbooks(1);
    var sheets = workbook.Worksheets;

    // 合同表列表
    var contractSheets = [
        "WMJZZL22025006702",
        "WMJZZL22025006801",
        "WMJZZL22025006901"
    ];

    // 删除已存在的"汇总合同"表
    for (var i = 1; i <= sheets.Count; i++) {
        if (sheets(i).Name === "汇总合同") {
            sheets(i).Delete();
            break;
        }
    }

    // 创建汇总表
    var summarySheet = sheets.Add();
    summarySheet.Name = "汇总合同";

    // 获取表头（从第一个合同表）
    var headerSheet = workbook.Worksheets(contractSheets[0]);
    var lastHeaderCol = headerSheet.Cells(1, headerSheet.Columns.Count).End(-4162).Column; // -4162 = xlToLeft

    // 复制表头
    headerSheet.Range("A1").Offset(0, lastHeaderCol - 1).Resize(1, 1).Copy(summarySheet.Range("A1"));
    summarySheet.Range("A1").Value = "合同编号";

    var currentRow = 2;

    // 合并每个合同表
    for (var s = 0; s < contractSheets.length; s++) {
        var sheetName = contractSheets[s];
        var sourceSheet = workbook.Worksheets(sheetName);

        // 获取最后一行
        var lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(-4162).Row;
        var lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(-4162).Column;

        if (lastRow < 2) continue; // 跳过空表

        // 复制数据行
        if (currentRow === 2) {
            // 第一张表：复制表头+数据
            sourceSheet.Range("A1:V" + lastRow).Copy(summarySheet.Range("A1"));
        } else {
            // 后续表：只复制数据
            sourceSheet.Range("A2:V" + lastRow).Copy(summarySheet.Range("A" + currentRow));
        }

        // 填入合同编号
        for (var r = currentRow; r <= currentRow + lastRow - 2; r++) {
            summarySheet.Cells(r, lastCol + 1).Value = sheetName;
        }

        currentRow = currentRow + lastRow - 1;
    }

    // 调整列宽
    summarySheet.Columns.AutoFit();

    // 冻结首行
    summarySheet.Range("A2").Activate();
    summarySheet.FreezePanes = true;

    return "完成！已合并 " + contractSheets.length + " 个合同表，共 " + (currentRow - 2) + " 行数据到【汇总合同】工作表";
}

/**
 * 刷新实时余额表
 * 功能：根据截至日期刷新"实时余额表"中的汇总数据
 */
function 刷新实时余额表() {
    var workbook = Application.Workbooks(1);
    var summarySheet = workbook.Worksheets("汇总合同");

    if (!summarySheet) {
        return "错误：找不到【汇总合同】工作表，请先运行【合并合同表】";
    }

    // 获取截至日期
    var dateCell = workbook.Worksheets("实时余额表").Range("B2");
    var endDate = dateCell.Value;

    if (!endDate) {
        return "错误：实时余额表中未设置截至日期";
    }

    // 筛选符合条件的数据
    var lastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(-4162).Row;
    var count = 0;

    for (var r = 2; r <= lastRow; r++) {
        var rowDate = summarySheet.Cells(r, 2).Value; // B列是承付日期
        if (rowDate && rowDate >= endDate) {
            count++;
        }
    }

    return "截至 " + endDate.toLocaleDateString() + "，共有 " + count + " 期未回笼";
}
