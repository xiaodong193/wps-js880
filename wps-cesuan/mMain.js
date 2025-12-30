/**
 * 租金测算系统主文件
 * 整合所有租金测算功能的主模块
 */

// 确保参数管理器已加载
if (typeof ParameterManager === 'undefined') {
    console.log("错误：ParameterManager未定义，请检查mParameterManager.js是否正确加载");
}

/**
 * 主计算函数
 */
function 计算main(){
	// 输入租金成本在B4单元格
    p.m_worksheet.Range("B4").Value2 = 50000000;// 租金成本
    //租赁票面利率
    p.m_worksheet.Range("B5").Value2 = 4.5;// 租赁票面利率
    //名义货价
    p.m_worksheet.Range("B6").Value2 = 1;// 名义货价
    //总期数
    p.m_worksheet.Range("B7").Value2 = 12;// 总期数
    //偿还方式
    p.m_worksheet.Range("B8").Value2 = "等额本息（后付）";// 偿还方式
    //支付间隔
    p.m_worksheet.Range("D10").Value2 = 6;// 支付间隔
    //放款日
    p.m_worksheet.Range("B9").Value2 =  p.GetDefaultDate();// 放款日
    //经纪人费用比例
    p.m_worksheet.Range("G16").Value2 = 0.001;// 经纪人费用比例


	生成租金表();
	生成现金流量表();
}

/**
 * 清除函数
 */
function 清除(){
	const r = new RentalCalculation();
	r.Initialize();
	r.清除原有表中数据();
}

/**
 * 调期函数
 */
function 调期(){
	const r = new RentalCalculation();
	r.Initialize();
	r.生成月间隔();
}


/**
 * 初始化系统
 */
function 初始化系统() {
    try {
        console.log("=== 租金测算系统初始化 ===");
        console.log("系统版本：" + VERSION);
        console.log("作者：" + AUTHOR);
        console.log("系统名称：" + MODULE_NAME);
        console.log("系统初始化完成");
        return true;
    } catch (error) {
        console.log("系统初始化失败：" + error.message);
        return false;
    }
}

function 生成租金表(){
	const r = new RentalCalculation();
	r.Initialize();
	r.清除原有表中数据();
	r.创建租金测算表表头(1, 10);
	r.createDataRange();
	r.每期适用利率();
	}

function 生成现金流量表() {
    let cashFlowGen = null;
    try {


        // 在WPS JS环境中，我们直接调用CashFlowGenerator类
        cashFlowGen = new CashFlowGenerator();
        cashFlowGen.Initialize()
        if (cashFlowGen.GenerateCashFlowTable()) {
        	cashFlowGen.综合利率一览()
            return true;
        } else {
            console.log(`[${cashFlowGen.MODULE_NAME}] 现金流量表生成失败`);
            return false;
        }
    } catch (error) {
        const moduleName = cashFlowGen ? cashFlowGen.MODULE_NAME : "CashFlowGenerator";
        console.log(`[${moduleName}] 生成现金流量表失败：${error.message}`);
        alert(`生成现金流量表时发生错误：${error.message}`);
        return false;
    }
}	
function 生成银承现金流量表() {
    try {
        console.log("开始测试完整银行承兑汇票现金流量表生成...");

        const bankModule = new cls银行承兑汇票();
        //let arr = bankModule.arrRngCashFlow();
//        Console.clear();
//        Console.log(arr[0][0].value);
//        arr[0][0].value = 1;
//        Console.log(arr[0][0].value);
        const result = bankModule.生成银承现金流量表();
        
        console.log("完整银行承兑汇票现金流量表生成结果: " + result);
        return result;
        
    } catch (error) {
        console.log("完整银行承兑汇票现金流量表生成测试失败: " + error.message);
        return false;
    }
}