/**
 * 合同余额合并模块 - 增强版
 */

function 合并合同表() {
    var workbook = Application.Workbooks(1);
    var sheets = workbook.Worksheets;
    var contractSheets = [
        "WMJZZL22025006702",
        "WMJZZL22025006801",
        "WMJZZL22025006901"
    ];
    for (var i = 1; i <= sheets.Count; i++) {
        if (sheets(i).Name === "汇总合同") {
            sheets(i).Delete();
            break;
        }
    }
    var summarySheet = sheets.Add();
    summarySheet.Name = "汇总合同";
    var headerSheet = asSheet(contractSheets[0]);
    //var headerSheet = workbook.Worksheets(contractSheets[0]);
    headerSheet.Range("A1:V1");
    headerSheet.Range("A1:V1").Copy(summarySheet.Range("A1"));
    summarySheet.Range("W1").Value2 = "合同编号";
    var currentRow = 2;
    for (var s = 0; s < contractSheets.length; s++) {
        var sheetName = contractSheets[s];
        var sourceSheet = workbook.Worksheets(sheetName);
        var lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(-4162).Row;
        if (lastRow < 2) continue;
        var rowCount = lastRow - 1;
        if (currentRow === 2) {
            sourceSheet.Range("A2:V" + lastRow).Copy(summarySheet.Range("A2"));
        } else {
            sourceSheet.Range("A2:V" + lastRow).Copy(summarySheet.Range("A" + currentRow));
        }
        for (var r = currentRow; r <= currentRow + rowCount - 1; r++) {
            summarySheet.Cells(r, 23).Value2 = sheetName;
        }
        currentRow = currentRow + rowCount;
    }
    summarySheet.Columns.AutoFit();
    summarySheet.Range("A2").Select();
    summarySheet.FreezePanes = true;
    return "完成！已合并 " + contractSheets.length + " 个合同表，共 " + (currentRow - 2) + " 行";
}

function OA日期转JS日期(oaDate) {
    if (typeof oaDate !== "number") return new Date(oaDate);
    return new Date((oaDate - 25569) * 86400 * 1000);
}

function 日期转字符串(dateObj) {
    return dateObj.getFullYear() + "-" + String(dateObj.getMonth() + 1).padStart(2, "0");
}

/**
 * 显示指定月份明细
 * @param {number} year - 年份，如 2026
 * @param {number} month - 月份，如 5
 */
function 显示月份明细(year, month) {
    var workbook = Application.Workbooks(1);
    var summarySheet = workbook.Worksheets("汇总合同");
    if (!summarySheet) return "错误：请先运行【合并合同表】";

    var sheetName = "月份明细_" + year + "年" + month + "月";
    for (var i = 1; i <= workbook.Worksheets.Count; i++) {
        if (workbook.Worksheets(i).Name === sheetName) {
            workbook.Worksheets(i).Delete();
            break;
        }
    }

    var detailSheet = workbook.Worksheets.Add();
    detailSheet.Name = sheetName;
    summarySheet.Range("A1:W1").Copy(detailSheet.Range("A1"));

    var currentRow = 2;
    var lastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(-4162).Row;

    for (var r = 2; r <= lastRow; r++) {
        var dateVal = summarySheet.Cells(r, 2).Value2;
        if (!dateVal) continue;
        var dateObj = OA日期转JS日期(dateVal);
        if (dateObj.getFullYear() === year && dateObj.getMonth() + 1 === month) {
            summarySheet.Range("A" + r + ":W" + r).Copy(detailSheet.Range("A" + currentRow));
            var status = summarySheet.Cells(r, 3).Value2;
            if (status === "未回笼") {
                detailSheet.Range("A" + currentRow + ":W" + currentRow).Interior.ColorIndex = 45;
            }
            currentRow++;
        }
    }

    detailSheet.Columns.AutoFit();
    detailSheet.Range("A2").Select();
    detailSheet.FreezePanes = true;
    return year + "年" + month + "月明细已生成，共 " + (currentRow - 2) + " 行";
}

/**
 * 显示指定日期区间明细
 * @param {number} startYear - 起始年份
 * @param {number} startMonth - 起始月份
 * @param {number} endYear - 结束年份
 * @param {number} endMonth - 结束月份
 */
function 显示区间明细(startYear, startMonth, endYear, endMonth) {
    var workbook = Application.Workbooks(1);
    var summarySheet = workbook.Worksheets("汇总合同");
    if (!summarySheet) return "错误：请先运行【合并合同表】";

    var startDate = new Date(startYear, startMonth - 1, 1);
    var endDate = new Date(endYear, endMonth, 0);

    var sheetName = "区间明细_" + startYear + "." + startMonth + "-" + endYear + "." + endMonth;
    for (var i = 1; i <= workbook.Worksheets.Count; i++) {
        if (workbook.Worksheets(i).Name === sheetName) {
            workbook.Worksheets(i).Delete();
            break;
        }
    }

    var detailSheet = workbook.Worksheets.Add();
    detailSheet.Name = sheetName;
    summarySheet.Range("A1:W1").Copy(detailSheet.Range("A1"));

    var currentRow = 2;
    var lastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(-4162).Row;

    var count = { total: 0, returned: 0, pending: 0 };

    for (var r = 2; r <= lastRow; r++) {
        var dateVal = summarySheet.Cells(r, 2).Value2;
        if (!dateVal) continue;
        var dateObj = OA日期转JS日期(dateVal);

        if (dateObj >= startDate && dateObj <= endDate) {
            summarySheet.Range("A" + r + ":W" + r).Copy(detailSheet.Range("A" + currentRow));
            var status = summarySheet.Cells(r, 3).Value2;
            if (status === "未回笼") {
                detailSheet.Range("A" + currentRow + ":W" + currentRow).Interior.ColorIndex = 45;
                count.pending++;
            } else {
                count.returned++;
            }
            currentRow++;
            count.total++;
        }
    }

    detailSheet.Columns.AutoFit();
    detailSheet.Range("A2").Select();
    detailSheet.FreezePanes = true;

    return startYear + "." + startMonth + " 至 " + endYear + "." + endMonth + " 明细已生成\n共 " + count.total + " 期（已回笼 " + count.returned + "，未回笼 " + count.pending + "）";
}

/**
 * 月度分类汇总
 */
function 月度分类汇总() {
    var workbook = Application.Workbooks(1);
    var summarySheet = workbook.Worksheets("汇总合同");
    if (!summarySheet) return "错误：请先运行【合并合同表】";

    for (var i = 1; i <= workbook.Worksheets.Count; i++) {
        if (workbook.Worksheets(i).Name === "月度汇总分析") {
            workbook.Worksheets(i).Delete();
            break;
        }
    }

    var now = new Date();
    var currentYear = now.getFullYear();
    var currentMonth = now.getMonth() + 1;
    var lastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(-4162).Row;

    var monthStats = {
        "已回笼": { count: 0, rent: 0, principal: 0, interest: 0 },
        "未回笼": { count: 0, rent: 0, principal: 0, interest: 0 }
    };

    for (var r = 2; r <= lastRow; r++) {
        var dateVal = summarySheet.Cells(r, 2).Value2;
        var status = summarySheet.Cells(r, 3).Value2;
        var rent = summarySheet.Cells(r, 5).Value2 || 0;
        var principal = summarySheet.Cells(r, 6).Value2 || 0;
        var interest = summarySheet.Cells(r, 7).Value2 || 0;

        if (!dateVal) continue;
        var dateObj = OA日期转JS日期(dateVal);
        if (dateObj.getFullYear() === currentYear && dateObj.getMonth() + 1 === currentMonth) {
            var key = (status === "已回笼") ? "已回笼" : "未回笼";
            monthStats[key].count++;
            monthStats[key].rent += rent;
            monthStats[key].principal += principal;
            monthStats[key].interest += interest;
        }
    }

    var analysisSheet = workbook.Worksheets.Add();
    analysisSheet.Name = "月度汇总分析";
    analysisSheet.Range("A1").Value2 = "项目";
    analysisSheet.Range("B1").Value2 = "金额";
    analysisSheet.Range("A1:B1").Font.Bold = true;
    analysisSheet.Range("A1:B1").Interior.ColorIndex = 15;

    var row = 2;
    analysisSheet.Cells(row, 1).Value2 = currentYear + "年" + currentMonth + "月 汇总";
    analysisSheet.Cells(row, 1).Font.Bold = true;
    row++;

    analysisSheet.Cells(row, 1).Value2 = "已回笼 笔数";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].count;
    row++;
    analysisSheet.Cells(row, 1).Value2 = "已回笼 租金合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].rent;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;
    analysisSheet.Cells(row, 1).Value2 = "已回笼 本金合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].principal;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;
    analysisSheet.Cells(row, 1).Value2 = "已回笼 租息合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].interest;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;

    analysisSheet.Cells(row, 1).Value2 = "未回笼 笔数";
    analysisSheet.Cells(row, 2).Value2 = monthStats["未回笼"].count;
    row++;
    analysisSheet.Cells(row, 1).Value2 = "未回笼 租金合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["未回笼"].rent;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;
    analysisSheet.Cells(row, 1).Value2 = "未回笼 本金合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["未回笼"].principal;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;
    analysisSheet.Cells(row, 1).Value2 = "未回笼 租息合计";
    analysisSheet.Cells(row, 2).Value2 = monthStats["未回笼"].interest;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    row++;

    row++;
    analysisSheet.Cells(row, 1).Value2 = "当月总计 租金";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].rent + monthStats["未回笼"].rent;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    analysisSheet.Cells(row, 1).Font.Bold = true;
    analysisSheet.Cells(row, 2).Font.Bold = true;
    row++;
    analysisSheet.Cells(row, 1).Value2 = "当月总计 本金";
    analysisSheet.Cells(row, 2).Value2 = monthStats["已回笼"].principal + monthStats["未回笼"].principal;
    analysisSheet.Cells(row, 2).NumberFormat = "#,##0.00";
    analysisSheet.Cells(row, 1).Font.Bold = true;
    analysisSheet.Cells(row, 2).Font.Bold = true;

    analysisSheet.Columns("A:A").ColumnWidth = 25;
    analysisSheet.Columns("B:B").ColumnWidth = 18;

    for (var r = 2; r <= row; r++) {
        var label = analysisSheet.Cells(r, 1).Value2;
        if (label && label.indexOf("未回笼") >= 0) {
            analysisSheet.Cells(r, 1).Interior.ColorIndex = 45;
            analysisSheet.Cells(r, 2).Interior.ColorIndex = 45;
        }
    }

    analysisSheet.Activate();
    var totalCount = monthStats["已回笼"].count + monthStats["未回笼"].count;
    return currentYear + "年" + currentMonth + "月汇总完成！共 " + totalCount + " 期（已回笼 " + monthStats["已回笼"].count + "，未回笼 " + monthStats["未回笼"].count + "）";
}

/**
 * 显示当月明细
 */
function 显示当月明细() {
    var now = new Date();
    return 显示月份明细(now.getFullYear(), now.getMonth() + 1);
}

/**
 * 刷新实时余额表
 */
function 刷新实时余额表() {
    var workbook = Application.Workbooks(1);
    var summarySheet = workbook.Worksheets("汇总合同");
    if (!summarySheet) return "错误：请先运行【合并合同表】";

    var dateCell = workbook.Worksheets("实时余额表").Range("B2");
    var endDate = dateCell.Value2;
    if (!endDate) return "错误：未设置截至日期";

    var lastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(-4162).Row;
    var count = 0;
    for (var r = 2; r <= lastRow; r++) {
        var rowDate = summarySheet.Cells(r, 2).Value2;
        if (rowDate && rowDate >= endDate) count++;
    }
    return "截至 " + OA日期转JS日期(endDate).toLocaleDateString() + "，共有 " + count + " 期未回笼";
}

/**
 * 主函数 - 一键执行全部操作
 */
function Main() {
    合并合同表();
    月度分类汇总();
    显示当月明细();
    刷新实时余额表();
    显示月份明细(2026, 5);
    显示区间明细(2026, 1, 2026, 5);
}
}