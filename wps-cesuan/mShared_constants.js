
/**
 * ============== 共享常量定义 ==============
 * 作者：徐晓冬
 * 描述：定义所有模块共享的常量，避免重复声明
 * ====================================================
 */

// ============== 模块级常量声明 ==============
const MODULE_NAME = "租金测算";
const AUTHOR = "徐晓冬";
const VERSION = "4.2025.11.24";

// ============== 错误代码常量 ==============
const ERR_INITIALIZE = 1000;
const ERR_VALIDATION = 2000;
const ERR_FORMULA = 3000;
const ERR_RANGE = 4000;

// ============== 格式常量 ==============
const FORMAT_STANDARD = "#,##0.00";    // 标准数字格式
const FORMAT_INTEGER = "0";            // 整数格式
const FORMAT_DATE = "yyyy-mm-dd";      // 日期格式
const FORMAT_PERCENTAGE = "0.00%";     // 百分比格式
const FORMAT_TEXT = "@";               // 文本格式

// 建议：在模块顶部自定义最常用的枚举值（JSA）
// 注意：xlDown, xlContinuous, xlThin, xlCenter 是Excel/WPS JSA内置常量，无需重复声明
const XL = {
  HCenter: -4108,   // 水平居中
  VCenter: -4108,   // 垂直居中（同值）
  Left:   -4131,
  Right:  -4152,
  Top:    -4160,
  Bottom: -4107
};

// ============== 字体常量 ==============
const FONT_DEFAULT = "黑体";
const FONT_ENGLISH = "Arial";
const FONT_CHINESE = "微软雅黑";
const FONT_SIZE_TITLE = 14;
const FONT_SIZE_HEADER = 12;
const FONT_SIZE_NORMAL = 10;
const FONT_SIZE_LARGE = 26;


function saveWithFileSystem(fileName, content) {
    let path = "D:\\" + fileName;
    
    // 检查路径是否存在（可选）
    if (FileSystem.Exists(path)) {
        console.log("文件已存在，将覆盖");
    }

    // 直接写入文本
    // FileSystem.WriteTextFile(路径, 内容, 编码)
    // 编码通常默认 utf-8
    FileSystem.WriteFile(path, content);
    
    
    console.log("保存成功");
}
/**
 * 获取静态值转换设置
 */
function GetStaticValueConversion() {
    try {
        // 目前固定返回false，未来可从配置单元格读取
        return false;
    } catch (error) {
        console.log(`[${MODULE_NAME}] 获取静态值转换设置失败，使用默认值true：${error.message}`);
        return true;
    }
}

/**
 * 批量生成R1C1公式并设置格式（优化版）
 */
function GenerateFormulasR1C1(rng, baseFormula, staticValueConversion = false, formatType = "Standard") {
    try {
        // 验证输入参数
        if (rng === null) {
            throw new Error("目标范围不能为空");
        }
        if (baseFormula.length === 0) {
            throw new Error("公式不能为空");
        }
        if (!IsInArray(formatType, ["Standard", "Integer", "Date"])) {
            throw new Error("无效的格式类型，支持类型为：Standard, Integer, Date");
        }
        
        // 设置公式
        rng.FormulaR1C1 = baseFormula;
        
        // 静态值转换（仅在需要时执行）
        if (staticValueConversion) {
            rng.Value2 = rng.Value2;
        }
        
        // 根据 formatType 设置格式
        switch (formatType) {
            case "Standard":
                rng.NumberFormat = "#,##0.00";
                break;
            case "Integer":
                rng.NumberFormat = "0";
                break;
            case "Date":
                rng.NumberFormat = "yyyy-mm-dd";
                break;
        }
        
        return true;
    } catch (error) {
        console.log(`生成公式失败：${error.message}`);
        alert(`生成公式失败：${error.message}`);
        return false;
    }
}

/**
 * 判断字符串是否在数组中
 */
function IsInArray(str, arr) {
    return arr.includes(str);
}

/**
 * 批量生成公式并设置格式
 */
function GenerateFormulas(rng, baseFormula, staticValueConversion = false) {
    try {
        // 验证输入参数
        if (rng === null) {
            throw new Error("目标范围不能为空");
        }
        if (baseFormula.length === 0) {
            throw new Error("公式不能为空");
        }
        
        // 批量生成公式
        rng.Formula = baseFormula;
        if (staticValueConversion) {
            rng.Value2 = rng.Value2;
        }
        rng.NumberFormat = "#,##0.00";
        
        return true;
    } catch (error) {
        console.log(`生成公式失败：${error.message}`);
        alert(`生成公式失败：${error.message}`);
        return false;
    }
}
/**
 * 给单元格或者范围应用格式
 */
function 应用格式(rng, formatType) {
    try {
        // 验证输入参数
        if (rng === null) {
            throw new Error("目标范围不能为空");
        }
        
        // 验证 formatType 类型
        if (!IsInArray(formatType, ["Standard", "Integer", "Date", "Text", "Percentage"])) {
            throw new Error("无效的格式类型，支持类型为：Standard, Integer, Date, Text, Percentage");
        }
        
        // 根据 formatType 设置格式
        switch (formatType) {
            case "Standard":
                rng.NumberFormat = "#,##0.00";
                break;
            case "Integer":
                rng.NumberFormat = "0";
                break;
            case "Date":
                rng.NumberFormat = "yyyy-mm-dd";
                break;
            case "Text":
                rng.NumberFormat = "@";
                break;
            case "Percentage":
                rng.NumberFormat = "0.00%";
                break;
        }
        
        return true;
    } catch (error) {
        console.log(`应用格式失败：${error.message} 格式类型：${formatType}`);
        return false;
    }
}
function 设置表格样式(rng) {
    try {
        // 设置字体
        rng.Font.Name = FONT_ENGLISH;
        rng.Font.Size = FONT_SIZE_NORMAL;
        
        // 设置对齐方式
        rng.HorizontalAlignment = XL.HCenter;
        rng.VerticalAlignment = XL.VCenter;
        
        // 设置行高和列宽
        rng.Rows.AutoFit();
        //rng.Columns.AutoFit();

        return true;
    } catch (error) {
        console.log(`设置表格样式失败：${error.message}`);
        return false;
    }
}
function 设置背景颜色(rng, color){
	rng.Interior.Color = color; 
}
/**
 * 生成综合利率计算结果
 */
function 测试自定义月间隔(){
	r = new RentalCalculation();
	r.Initialize();
	r.生成月间隔();
}

function arrDataFromRng(sheet, RentTableStartRow, arrHeaders){
    	//采用数组编写,生成表数据
        let arr = [];//存放银承现金流量表数据，设计成二维数组
        //计算生成arr的行数（row值），等于 放款形成的row + 还款形成的row，
        // 计算从银承现金流表数据第1行开始，到最后一行的row值。
        //列值（col值）为headers的长度，即15列
		//RentTableStartRow是表格数据值开始的行号
        let rngStart= sheet.Range("A"+String(RentTableStartRow));
        let rngEnd= rngStart.End(xlDown);
        let rows= rngEnd.Row - RentTableStartRow +1;//总行数,作为arr的row值，row值就是数组的第1维脚标数字
        console.log("总行数,作为arr的row值，row值就是数组的第1维脚标数字:"+ rows);
        let cols = arrHeaders.length || 0; // 简化：列值（col值）为headers的长度，即15列，作为arr的第2维脚标数字
        
        console.log("列值（col值）为headers的长度，作为arr的第2维脚标数字:"+ cols);
        
        const nameOfsheet = sheet.Name;
        //let range = sheet.getRange(RentTableStartRow, 1, rows, cols); 
        // 修复：应该是 RentTableStartRow+rows-1，避免多读一行
        let range = sheet.Range(sheet.Cells(RentTableStartRow, 1), sheet.Cells(RentTableStartRow+rows-1, cols)) 
        let sheetData = range.Value2; // 获取范围内的所有值
        for (let i = 0; i < rows; i++){
            let row = [];
            for (let j = 0; j < cols;j++){
                //初始化二维数组arr[i][j]的值
                //arr[i][j] = "";//单元格值
                row.push({
                    value: sheetData[i][j], //单元格值
					Formula: "", //单元格公式
                    row: i, //行索引
                    col: j, //列索引
                    address: `R${i+1}C${j+1}`, //单元格地址
                    sheetAddress: `${nameOfsheet}!R${i+1}C${j+1}`, //在Sheet中的实际地址,工作表+单元格地址
                    header: arrHeaders[j] //对应的表头名称
                });
            }
            arr.push(row);
        }
        //console.log("arr:"+ JSON.stringify(arr, null, 2));//调试信息，打印二维数组内容
        return arr;
}

/**
 * WPS JSA数组扩展函数 - 实现列、行、单元格级别的读写操作
 * @param {Object} sheet - Excel工作表对象
 * @param {number} RentTableStartRow - 表格数据开始行号
 * @param {Array} arrHeaders - 表头数组
 * @returns {Object} 包含数据和操作方法的对象
 */
/**
 * WPS JSA数组扩展函数 - 实现列、行、单元格级别的读写操作
 * @param {Object} sheet - Excel工作表对象
 * @param {number} RentTableStartRow - 表格数据开始行号
 * @param {Array} arrHeaders - 表头数组
 * @returns {Object} 包含数据和操作方法的对象
 */
function arrDataFromRngExtended(sheet, RentTableStartRow, arrHeaders) {
    // 1. 读取数据到数组（基于原始函数逻辑）
    var arr = [];
    var rngStart = sheet.Range("A" + String(RentTableStartRow));
    var rngEnd = rngStart.End(xlDown);
    var rows = rngEnd.Row - RentTableStartRow + 1;
    var cols = arrHeaders.length || 0;

    var nameOfsheet = sheet.Name;
    var range = sheet.Range(sheet.Cells(RentTableStartRow, 1), sheet.Cells(RentTableStartRow + rows - 1, cols));
    var sheetData = range.Value2;
    
    // 创建数据数组（保持与原始函数一致的结构）
    for (var i = 0; i < rows; i++) {
        var row = [];
        for (var j = 0; j < cols; j++) {
            row.push({
                value: sheetData[i][j],
                Formula: "",
                row: i,
                col: j,
                address: "R" + (i+1) + "C" + (j+1),
                sheetAddress: nameOfsheet + "!R" + (i+1) + "C" + (j+1),
                header: arrHeaders[j],
                ratePerPeriod: null
            });
        }
        arr.push(row);
    }
    
    // 返回包含扩展功能的对象
    return {
        // 基础数据数组
        data: arr,
        
        // 1. 读取指定区域的某列数据（兼容原始函数名）
        getColumn: function(colIndex) {
            // 参数验证
            if (colIndex < 0 || colIndex >= cols) {
                throw new Error('列索引超出范围');
            }
            // 返回指定列的所有单元格数据
            var columnData = [];
            for (var i = 0; i < rows; i++) {
                columnData.push(arr[i][colIndex].value);
            }
            return columnData;
        },
        
        // 2. 修改指定区域的某列数据（兼容原始函数名）
        setColumn: function(colIndex, newData) {
            // 参数验证
            if (colIndex < 0 || colIndex >= cols) {
                throw new Error('列索引超出范围');
            }
            if (!Array.isArray(newData) || newData.length !== rows) {
                throw new Error('数据长度不匹配，应为' + rows + '个元素');
            }
            // 修改指定列的数据
            for (var i = 0; i < rows; i++) {
                arr[i][colIndex].value = newData[i];
            }
            // 同步更新到工作表
            var colRange = sheet.Range(sheet.Cells(RentTableStartRow, colIndex + 1), sheet.Cells(RentTableStartRow + rows - 1, colIndex + 1));
            var newDataArray = [];
            for (var i = 0; i < rows; i++) {
                newDataArray.push([newData[i]]);
            }
            colRange.Value2 = newDataArray;
            return arr;
        },
        
        // 3. 读取指定区域的某行数据（兼容原始函数名）
        getRow: function(rowIndex) {
            // 参数验证
            if (rowIndex < 0 || rowIndex >= rows) {
                throw new Error('行索引超出范围');
            }
            // 返回指定行的所有单元格数据
            var rowData = [];
            for (var j = 0; j < cols; j++) {
                rowData.push(arr[rowIndex][j].value);
            }
            return rowData;
        },
        
        // 4. 修改指定区域的某行数据（兼容原始函数名）
        setRow: function(rowIndex, newData) {
            // 参数验证
            if (rowIndex < 0 || rowIndex >= rows) {
                throw new Error('行索引超出范围');
            }
            if (!Array.isArray(newData) || newData.length !== cols) {
                throw new Error('数据长度不匹配，应为' + cols + '个元素');
            }
            // 修改指定行的数据
            for (var j = 0; j < cols; j++) {
                arr[rowIndex][j].value = newData[j];
            }
            // 同步更新到工作表
            var rowRange = sheet.Range(sheet.Cells(RentTableStartRow + rowIndex, 1), sheet.Cells(RentTableStartRow + rowIndex, cols));
            rowRange.Value2 = [newData];
            return arr;
        },
        
        // 5. 读取特定区域的数值（兼容原始函数名）
        getCell: function(rowIndex, colIndex) {
            // 参数验证
            if (rowIndex < 0 || rowIndex >= rows) {
                throw new Error('行索引超出范围');
            }
            if (colIndex < 0 || colIndex >= cols) {
                throw new Error('列索引超出范围');
            }
            // 返回指定单元格的值
            return arr[rowIndex][colIndex].value;
        },
        
        // 6. 修改特定区域的数值（兼容原始函数名）
        setCell: function(rowIndex, colIndex, value) {
            // 参数验证
            if (rowIndex < 0 || rowIndex >= rows) {
                throw new Error('行索引超出范围');
            }
            if (colIndex < 0 || colIndex >= cols) {
                throw new Error('列索引超出范围');
            }
            // 修改指定单元格的数据
            arr[rowIndex][colIndex].value = value;
            // 同步更新到工作表
            var cellRange = sheet.Range(sheet.Cells(RentTableStartRow + rowIndex, colIndex + 1));
            cellRange.Value2 = [[value]];
            return arr;
        },
        
        // 7. 获取数据维度
        getDimensions: function() {
            return {
                rows: rows,
                cols: cols
            };
        }
    };
};

/**
 * 重写二维数组中多列的指定行范围的数据
 * @param {Array} arr - 二维数组
 * @param {number} startRow - 开始行索引（包含）
 * @param {number} endRow - 结束行索引（包含）
 * @param {Array} colConfigs - 列配置数组，格式：[{colIndex: 1, newValue: 'X'}, {colIndex: 2, newValue: 'Y'}]
 * @returns {Array} 修改后的数组
 */

function arr重写列数据(arr, startRow, endRow, colConfigs) {
    // 验证输入参数
    if (!Array.isArray(arr) || arr.length === 0) {
        throw new Error('输入必须是非空数组');
    }
    
    if (startRow < 0 || endRow >= arr.length || startRow > endRow) {
        throw new Error('行索引范围无效');
    }
    
    if (!Array.isArray(colConfigs) || colConfigs.length === 0) {
        throw new Error('列配置不能为空');
    }
    
    // 创建数组副本
    const result = arr.map(row => [...row]);
    
    // 遍历每一行
    for (let i = startRow; i <= endRow; i++) {
        // 遍历每个列配置
        colConfigs.forEach(config => {
            const { colIndex, newValue } = config;
            
            if (colIndex < 0) {
                throw new Error(`列索引 ${colIndex} 不能为负数`);
            }
            
            // 扩展数组如果需要
            if (result[i].length <= colIndex) {
                while (result[i].length <= colIndex) {
                    result[i].push(undefined);
                }
            }
            
            // 设置新值
            if (typeof newValue === 'function') {
                result[i][colIndex] = newValue(result[i][colIndex], i, colIndex);
            } else {
                result[i][colIndex] = newValue;
            }
        });
    }
    
    return result;
}
function logjson(arr){
	const s = JSON.stringify(arr);
	console.log(s);
}

/**
 * 处理列数据转换和清除
 * @param {Range} rng - Excel范围对象
 * @param {Array} convertColumns - 需要转换为数值的列索引数组（从1开始）
 * @param {Array} clearColumns - 需要清除内容的列索引数组（从1开始）
 */
function ProcessColumnData(rng, convertColumns = [], clearColumns = []) {
    try {
        // 转换指定列为数值
        convertColumns.forEach(colIndex => {
            if (colIndex > 0) {
                const colRange = rng.Columns.Item(colIndex);
                colRange.Value2 = colRange.Value2;
            }
        });
        
        // 清除指定列的内容
        clearColumns.forEach(colIndex => {
            if (colIndex > 0) {
                const colRange = rng.Columns.Item(colIndex);
                colRange.ClearContents();
            }
        });
        
        return true;
    } catch (error) {
        console.log(`处理列数据失败：${error.message}`);
        return false;
    }
}
