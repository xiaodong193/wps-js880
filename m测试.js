Attribute Module_Name = "m测试"
/**
 * 银行承兑汇票模块测试文件
 * 用于测试银行承兑汇票功能是否正常工作
 */

/**
 * 银行承兑汇票测试函数
 */
function 银行承兑汇票测试() {
	try {
		const bankModule = new BankAcceptanceModule();
		bankModule.生成银承现金流量表();
		console.log("银行承兑汇票功能执行完成");
		return true;
	} catch (error) {
		console.log("银行承兑汇票功能执行失败：" + error.message);
		return false;
	}
}

function t1(){
	let rows = 11;
	let cols = 15;
	let excelArray = [];
	for (let i = 0; i < rows; i++) {
	    let row = [];
	    for (let j = 0; j < cols; j++) {
	        // 可以根据需要设置不同的初始值
	        row.push({
	            value: "",           // 单元格值
	            row: i,             // 行索引
	            col: j,             // 列索引
	            address: `R${i+1}C${j+1}` // 单元格地址
	        });
	    }
	    excelArray.push(row);
	}
	let usersJson = JSON.stringify(excelArray, null, 2)
	console.log(usersJson);
}


function t2(){
	生成租金表();
	let arr = r.arr租金测算表数据();
	console.clear();
	console.log(JSON.stringify(arr).toLocaleString);


}
function 测试租金调整(){
	// 创建调整对象
	const adjuster = new RentalAdjustment();

	// 初始化
	adjuster.Initialize();

	// 【配置调息】第7期调整为5%利率，第10期调整为4.8%
	adjuster.SetRateAdjustment(7, 0.05);   // 期号，新利率（小数形式）
	adjuster.SetRateAdjustment(10, 0.048);

	// 【配置租金变更】第8期起租金下降10%（乘以0.9）
	adjuster.SetRentalAdjustment(8, 0.90);  // 开始期号，倍数

	// 生成新Sheet中的调整表
	adjuster.GenerateAdjustedRentalTable();

	// 可选：创建调息配置表（方便查看每期利率）
	adjuster.CreateRateAdjustmentTable();
}



function test重写列数据(){
	r = new RentalCalculationWithAdjustment();
	r.Initialize();
	r.清除原有表中数据();
	r.创建租金测算表表头(1, 10);
	r.createDataRange();
	let arr2D = r.写入每期利率();
	r.每期适用利率();

	let arrFormula= r.等额租金法arr();//存储的公式二维数组
	let arrData = r.arrToArrData(arrFormula);//扩展数据区域的二维数组
	//rewriteMultipleColumnsRange配置参数
	let arr =arrData;
	let startRow = 3;
	let endRow =19;
	let colConfigs =[
    { colIndex: 1, newValue: 'COLUMN1_MODIFIED' },
    { colIndex: 2, newValue: 'COLUMN2_MODIFIED' }
	];

	arr = arr重写列数据(arr, startRow, endRow, colConfigs);
	console.clear();
	logjson(arr);
}

/**
 * 示例：如何使用处理列数据函数
 * 替代原有的重复代码：
 * const col1Range = rng.Columns.Item(1);// 期次列
 * const col2Range = rng.Columns.Item(2);// 日期列
 * const col4Range = rng.Columns.Item(4);// 本金列
 * const col12Range = rng.Columns.Item(12);// 本金比例列
 * col1Range.Value2 = col1Range.Value2;// 将期次列转换为数值
 * col2Range.Value2 = col2Range.Value2;// 将日期列转换为数值
 * col4Range.Value2 = col4Range.Value2;// 将本金列转换为数值
 * col12Range.ClearContents();// 清除本金比例列内容
 */
function 示例_处理列数据(rng) {
    // 使用新的通用函数替代原有的重复代码
    // 参数1: Excel范围对象
    // 参数2: 需要转换为数值的列索引数组（从1开始）
    // 参数3: 需要清除内容的列索引数组（从1开始）
    处理列数据(rng, [1, 2, 4], [12]);
}

/**
 * 在Sheet2上创建测试数据并运行maxRange测试
 * 表格场景：
 * - 标题: A1 数据：A2:C6
 * - 标题: A9  数据：A10:C13
 * - 数据: F1单个单元格
 */
function testMaxRangeOnSheet2() {
    console.clear();
    console.log("========== 在Sheet2上测试 maxRange ==========\n");

    const sheet = Application.ActiveWorkbook.Sheets.Item("Sheet2");
    sheet.Activate();

    // 清空Sheet2
    sheet.Cells.Clear();

    // 创建测试数据
    // 表一：A1:C6
    sheet.Range("A1").Value = "表一";
    sheet.Range("A2").Value = "城市";
    sheet.Range("A3").Value = "南京";
    sheet.Range("A4").Value = "上海";
    sheet.Range("A5").Value = "北京";
    sheet.Range("A6").Value = "海南";

    sheet.Range("B2").Value = "数量";
    sheet.Range("B3").Value = 10;
    sheet.Range("B4").Value = 8;
    sheet.Range("B5").Value = 9;
    sheet.Range("B6").Value = 11;

    sheet.Range("C2").Value = "金额";
    sheet.Range("C3").Value = 200;
    sheet.Range("C4").Value = 210;
    sheet.Range("C5").Value = 345;
    sheet.Range("C6").Value = 600;

    // 表二：A9:C13
    sheet.Range("A9").Value = "表二";
    sheet.Range("A10").Value = "城市";
    sheet.Range("A11").Value = "海南";
    sheet.Range("A12").Value = "北京";
    sheet.Range("A13").Value = "西藏";

    sheet.Range("B10").Value = "数量";
    sheet.Range("B11").Value = 10;
    sheet.Range("B12").Value = 5;
    sheet.Range("B13").Value = 8;

    sheet.Range("C10").Value = "金额";
    sheet.Range("C11").Value = 900;
    sheet.Range("C12").Value = 500;
    sheet.Range("C13").Value = 300;

    // F1单个单元格
    sheet.Range("F1").Value = "测试";

    console.log("测试数据已创建到Sheet2\n");
    console.log("数据结构：");
    console.log("  - 表一: A1:C6");
    console.log("  - 表二: A9:C13");
    console.log("  - F1: 单个单元格\n");

    // ==================== 开始测试 ====================

    // 示例1
    console.log("--- 示例1: z最大行区域('A1:A1000') ---");
    var endRowCell = RngUtils.z最大行区域("A1:A1000");
    console.log("结果: " + endRowCell.Address());
    console.log("预期: $A$1:$A$13");
    console.log("测试" + (endRowCell.Address() === "$A$1:$A$13" ? "通过" : "失败") + "\n");

    // 示例2
    console.log("--- 示例2: z最大行区域(Range('F:F')) ---");
    var endRowCell2 = RngUtils.z最大行区域(sheet.Range("F:F"));
    console.log("结果: " + endRowCell2.Address());
    console.log("预期: $F$1");
    console.log("测试" + (endRowCell2.Address() === "$F$1" ? "通过" : "失败") + "\n");

    // 示例3
    console.log("--- 示例3: z最大行区域('1:1000','A') ---");
    console.log("结果: " + RngUtils.z最大行区域("1:1000","A").Address());
    console.log("预期: $1:$13\n");

    // 示例4
    console.log("--- 示例4: z最大行区域('1:1000','F') ---");
    console.log("结果: " + RngUtils.z最大行区域("1:1000","F").Address());
    console.log("预期: $1:$1\n");

    // 示例5
    console.log("--- 示例5: z最大行区域('a1','-c') 连续区域 ---");
    console.log("结果: " + RngUtils.z最大行区域("a1","-c").Address());
    console.log("预期: $A$1:$C$6 (A1的CurrentRegion)\n");

    // 示例6
    console.log("--- 示例6: z最大行区域('a1','-u') UsedRange ---");
    console.log("结果: " + RngUtils.z最大行区域("a1","-u").Address());
    console.log("预期: $A$1:$A$13\n");

    console.log("========== 测试结束 ==========");
}

/**
 * 快速测试 maxRange 基本功能
 */
function testMaxRangeQuick() {
    console.clear();

    console.log(">>> 快速测试 maxRange <<<\n");

    // 测试1
    const r1 = RngUtils.maxRange("A1:A1000");
    console.log("maxRange('A1:A1000') = " + r1.Address());

    // 测试2
    const r2 = RngUtils.maxRange("1:1000", "A");
    console.log("maxRange('1:1000', 'A') = " + r2.Address());

    // 测试3
    const r3 = RngUtils.maxRange("a1", "-c");
    console.log("maxRange('a1', '-c') = " + r3.Address());

    console.log("\n>>> 测试完成 <<<");
}
