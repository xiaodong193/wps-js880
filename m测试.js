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
