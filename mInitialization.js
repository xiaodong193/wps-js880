Attribute Module_Name = "mInitialization"
// ============== 初始化模块：租金测算表参数区域初始化（完全集成参数管理器版本） ==============
// 作者：徐晓冬
// 描述：专门处理租金测算表各个参数区域的初始化功能，完全集成ParameterManager参数管理器
// 优化内容：直接调用参数管理器方法、统一参数管理、增强可维护性
// ===================================================================
// ============== 主初始化函数 ==============
class rentCalculationFillinArea {
	constructor(){
		this.MODULE_NAME = "测算表填写区域"
		this.ModuleModifyDate = "20251228"
		this.ws = p.m_worksheet
		this.p = p; // 存储参数管理器实例
		console.log(`[${this.MODULE_NAME}] 类实例创建`);
	}
	main(){
		console.log(`[${this.MODULE_NAME}] 开始初始化填写区域...`);
        // 使用参数管理器初始化各区域
        this.初始化价格参数区域();
        this.初始化利率要素();
        this.初始化租赁项目基本信息区域();
        this.初始化经纪人费用参数区域();
        
        console.log(`[${this.MODULE_NAME}] 填写区域初始化完成`);
        return true;
	}
	初始化价格参数区域() {
    // 功能：使用参数管理器初始化价格参数填写区域
    
	    try {
	    	const ws = this.ws
	        this.生成主标题("租赁项目预算租金偿还表","A1",ws);
	        this.生成区域标题("价格参数填写区域", "A3", ws);
	        // 使用参数管理器获取参数配置并设置单元格
	        this.设置单元格值("Principal", ws); // 租赁成本
	        this.设置单元格公式("ActualPayment", ws); // 实际付款
	        this.设置单元格值("InterestRate", ws); // 利率
	        this.设置单元格值("RateOption", ws); // 利率选项
	        this.设置参数数据有效性("K2:K3", "RateOption");
	        this.设置单元格公式("Deposit", ws); // 保证金
	        this.设置单元格值("DepositMarginRate", ws); // 押金保证金比例
	        this.设置单元格值("NominalPrice", ws); // 名义货价
	        this.设置单元格值("TotalPeriods", ws); // 总期数
	        this.设置单元格值("PaymentInterval", ws); // 支付间隔
	        this.设置参数数据有效性("B2:B5", "PaymentInterval");
	        this.设置单元格公式("PaymentsPerYear", ws); // 每年还款次数
	        this.设置单元格公式("ProjectDurationYears", ws); // 项目时长/年
	        this.设置单元格值("RepaymentMethod", ws); // 还款方式
	        this.设置参数数据有效性("A2:A9", "RepaymentMethod");
	        this.设置单元格值("LeaseStartDate", ws); // 放款日
	        
	        console.log(`[${this.MODULE_NAME}] 价格参数区域初始化完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 价格参数区域初始化失败：${error.message}`);
	        return false;
	    }
	}
	初始化利率要素() {
    // 功能：使用参数管理器初始化利率要素填写区域
    
	    try {
	        // 利率要素填写区域标题
	        const ws = this.ws
	        this.生成区域标题("利率要素填写区域", "F9", ws);
	        
	        // 使用参数管理器设置利率要素参数
	        this.设置单元格公式("LPRBenchmarkRate", ws); // LPR基准利率
	        this.设置单元格公式("FloatingBasisPoints", ws); // 浮动基点
	        this.设置单元格值("LPRPeriod", ws); // LPR期限选择
	        this.设置参数数据有效性("L2:L3", "LPRPeriod");
	        
	        // 先设置数据有效性
	        this.设置单元格值("LPRDate", ws); // LPR发布日期
	        this.设置参数数据有效性("N2:N13", "LPRDate");
	        
	        // 然后设置默认值为数据有效性列表的第一个值
	        const wsRepay = Application.Worksheets.Item("还款设置");
	        const firstValue = wsRepay.Range("N2").Value2;
	        const cellAddressA1 = p.GetCellAddressA1("LPRDate");
	
	        
	        const lprDateRange = ws.Range(cellAddressA1);
	        lprDateRange.Value2 = firstValue;
	        lprDateRange.Font.Size = 12;
	        lprDateRange.NumberFormat = "yyyy-mm-dd";
	        lprDateRange.Font.Bold = false;
	        lprDateRange.Font.Name = "黑体";
	        lprDateRange.Interior.Color = this.p.m_COLOR_YELLOW; // 设置单元格背景为黄色
	        lprDateRange.Font.Color = this.p.m_COLOR_BLACK; // 设置字体颜色为黑色
	        lprDateRange.HorizontalAlignment = -4152; // xlRight
	        lprDateRange.VerticalAlignment = -4108;   // xlCenter
	        wsRepay.Calculate();
	        
	        this.设置单元格值("LPRRateDescription", ws); 
	        const descCellAddress = p.GetCellAddressA1("LPRRateDescription");
	        const desc = this.获取利率描述();
	        
	        const descRange = ws.Range(descCellAddress);
	        descRange.Formula = desc;
	        descRange.Font.Size = 12;
	        descRange.NumberFormat = "@";
	        descRange.Font.Bold = false;
	        descRange.Font.Name = "黑体";
	        descRange.Interior.Color = this.p.m_COLOR_LIGHT_GREEN; // 设置单元格背景为浅绿色
	        descRange.Font.Color = this.p.m_COLOR_BLACK; // 设置字体颜色为黑色
	        descRange.HorizontalAlignment = -4131; // xlLeft
	        descRange.VerticalAlignment = -4108;   // xlCenter
	        descRange.WrapText = false; // 自动换行
	        
	        console.log(`[${this.MODULE_NAME}] 利率要素区域初始化完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 利率要素区域初始化失败：${error.message}`);
	        return false;
	    }
	}
	初始化租赁项目基本信息区域() {
    // 功能：使用参数管理器初始化租赁项目基本信息填写区域
    
	    try {
	    	const ws = this.ws
	        this.生成区域标题("租赁项目基本信息区域", "F3", ws);
	        
	        // 使用参数管理器设置项目信息参数
	        this.设置单元格值("Lessee", ws); // 承租人
	        this.设置单元格值("Guarantor", ws); // 担保人
	        this.设置单元格值("GuaranteeMethod", ws); // 担保方式
	        this.设置单元格值("LeaseMethod", ws); // 租赁方式
	        
	        console.log(`[${this.MODULE_NAME}] 租赁项目基本信息区域初始化完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 租赁项目基本信息区域初始化失败：${error.message}`);
	        return false;
	    }
	}
	初始化经纪人费用参数区域() {
    // 功能：使用参数管理器初始化经纪人费用参数填写区域
	    try {
	    	const ws = this.ws
	        this.生成区域标题("经纪人费用参数区域", "F13", ws);
	        
	        // 使用参数管理器设置经纪人费用参数
	        this.设置单元格值("BrokerPaymentMethod", ws); // 经纪人费用支付方式
	        this.设置参数数据有效性("D2:D6", "BrokerPaymentMethod");
	        this.设置单元格值("BrokerFeeRate", ws); // 经纪人费用比例
	        this.设置单元格公式("BrokerTotalFee", ws); // 经纪人总费用
	        
	        // 提示信息
	        const infoRange = ws.Range("H14");
	        infoRange.Value2 = "其他支付方式请在K43之后自行输入经纪人费用";
	        infoRange.Font.Size = 12;
	        infoRange.Font.Bold = false;
	        infoRange.Font.Name = "黑体";
	        infoRange.Interior.Color = this.p.m_COLOR_WHITE;
	        infoRange.Font.Color = this.p.m_COLOR_BLACK;
	        infoRange.HorizontalAlignment = -4131; // xlLeft
	        infoRange.VerticalAlignment = -4108;   // xlCenter
	        
	        console.log(`[${this.MODULE_NAME}] 经纪人费用参数区域初始化完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 经纪人费用参数区域初始化失败：${error.message}`);
	        return false;
	    }
	}
		// ============== 核心辅助函数（完全集成参数管理器） ==============
	
	// 参数A1样式地址（从参数管理器获取）
	参数A1样式地址(parameterName) {
	    // 功能：根据参数名称从参数管理器获取对应的值单元格地址
	    // 参数：parameterName - 参数名称
	    // 返回：A1样式单元格地址字符串
	    
		    try {
		        return p.GetCellAddressA1(parameterName);
		    } catch (error) {
		        console.log(`[${this.MODULE_NAME}] 获取参数地址失败：${error.message}`);
		        return null;
		    }
	}
	
	// ============== 辅助生成函数 ==============
	生成主标题(heading, cellAddressA1, worksheet){
		try {
	        // 设置参数标题单元格
	        const titleRange = worksheet.Range(cellAddressA1);
	        titleRange.Value2 = heading;
	        titleRange.Font.Size = 26;
	        titleRange.Font.Bold = true;
	        titleRange.Font.Name = "黑体";
	        titleRange.HorizontalAlignment = xlLeft; // xlLeft
	        titleRange.VerticalAlignment = xlCenter;   // xlCenter
	        titleRange.Rows.AutoFit();
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 生成区域标题失败：${error.message}`);
	        return false;
	    }
	}
	生成区域标题(heading, cellAddressA1, worksheet) {
	    // 功能：返回参数填写区域标题字符串
	    
	    try {
	        // 设置参数标题单元格
	        const titleRange = worksheet.Range(cellAddressA1);
	        titleRange.Value2 = heading;
	        titleRange.Font.Size = 14;
	        titleRange.Font.Bold = false;
	        titleRange.Font.Name = "黑体";
	        titleRange.HorizontalAlignment = -4131; // xlLeft
	        titleRange.VerticalAlignment = -4108;   // xlCenter
	        
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 生成区域标题失败：${error.message}`);
	        return false;
	    }
	}
	
	// 设置单元格值（完全集成参数管理器）
	设置单元格值(parameterName, worksheet) {
	    // 功能：直接调用参数管理器方法设置单个参数单元格
	    // 参数：parameterName - 参数名称
	    //       worksheet - 工作表对象
	    
	    try {
	        // 直接从参数管理器获取配置信息
	        const cellAddressA1 = p.GetCellAddressA1(parameterName);
	        const displayName = p.GetConfigValue(parameterName, "DisplayName");
	        const defaultValue = p.GetConfigValue(parameterName, "DefaultValue");
	        const dataType = p.GetConfigValue(parameterName, "DataType");
	        const vbaFormat = p.GetConfigValue(parameterName, "VbaFormat");
	        
	        // 设置参数值单元格
	        const valueRange = worksheet.Range(cellAddressA1);
	        valueRange.Value2 = defaultValue;
	        valueRange.Font.Size = 12;
	        valueRange.NumberFormat = vbaFormat;
	        valueRange.Font.Bold = false;
	        valueRange.Font.Name = "黑体";
	        valueRange.Interior.Color = this.p.m_COLOR_YELLOW; // 设置单元格背景为黄色
	        valueRange.Font.Color = this.p.m_COLOR_BLACK; // 设置字体颜色为黑色
	        valueRange.HorizontalAlignment = -4152; // xlRight
	        valueRange.VerticalAlignment = -4108;   // xlCenter
	        
	        // 设置参数名称单元格（左侧单元格）
	        const nameRange = valueRange.Offset(0, -1);
	        nameRange.Value2 = displayName;
	        nameRange.Font.Size = 12;
	        nameRange.Font.Bold = false;
	        nameRange.Font.Name = "黑体";
	        nameRange.HorizontalAlignment = -4131; // xlLeft
	        nameRange.VerticalAlignment = -4108;   // xlCenter
	        
	        console.log(`[${this.MODULE_NAME}] 参数'${parameterName}'设置完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 设置单元格值失败：${error.message}`);
	        return false;
	    }
	}
	
	// 设置单元格公式（完全集成参数管理器）
	设置单元格公式(parameterName, worksheet) {
	    // 功能：直接调用参数管理器方法设置单个参数单元格
	    // 参数：parameterName - 参数名称
	    //       worksheet - 工作表对象
	    
	    try {
	        // 直接从参数管理器获取配置信息
	        const cellAddressA1 = p.GetCellAddressA1(parameterName);
	        const displayName = p.GetConfigValue(parameterName, "DisplayName");
	        const dataType = p.GetConfigValue(parameterName, "DataType");
	        const defaultFormula = p.GetConfigValue(parameterName, "DefaultFormula");
	        const vbaFormat = p.GetConfigValue(parameterName, "VbaFormat");
	        
	        // 设置参数公式单元格
	        const formulaRange = worksheet.Range(cellAddressA1);
	        formulaRange.Formula = defaultFormula;
	        formulaRange.Font.Size = 12;
	        formulaRange.NumberFormat = vbaFormat;
	        formulaRange.Font.Bold = false;
	        formulaRange.Font.Name = "黑体";
	        formulaRange.Interior.Color = this.p.m_COLOR_LIGHT_GREEN; // 设置单元格背景为浅绿色
	        formulaRange.Font.Color = this.p.m_COLOR_BLACK; // 设置字体颜色为黑色
	        formulaRange.HorizontalAlignment = -4152; // xlRight
	        formulaRange.VerticalAlignment = -4108;   // xlCenter
	        
	        // 设置参数名称单元格（左侧单元格）
	        const nameRange = formulaRange.Offset(0, -1);
	        nameRange.Value2 = displayName;
	        nameRange.Font.Size = 12;
	        nameRange.Font.Bold = false;
	        nameRange.Font.Name = "黑体";
	        nameRange.Interior.Color = this.p.m_COLOR_WHITE; // 设置单元格背景为白色
	        nameRange.Font.Color = this.p.m_COLOR_BLACK; // 设置字体颜色为黑色
	        nameRange.HorizontalAlignment = -4131; // xlLeft
	        nameRange.VerticalAlignment = -4108;   // xlCenter
	        
	        console.log(`[${this.MODULE_NAME}] 参数'${parameterName}'公式设置完成`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 设置单元格公式失败：${error.message}`);
	        return false;
	    }
	}
	
	设置参数数据有效性(rngSourceA1, parameterName) {
	    // 功能：根据参数配置设置数据有效性
	    // 参数：parameterName - 参数名称
	    
	    try {
	    	const ws = this.ws
	        const activeWorkbook = Application.ActiveWorkbook;
	        const wsRepay = activeWorkbook.Worksheets.Item("还款设置");
	        const rngSource = wsRepay.Range(rngSourceA1);
	        
	        // 参数A1样式地址
	        const cellAddressA1 = p.GetCellAddressA1(parameterName);
	        const targetRange = ws.Range(cellAddressA1);
	        const Address1 = rngSource.Address(true, true, xlA1, true);
	        const rngSourceValue = rngSource.Value2;
	        
	        // 清除原有数据有效性
	        try {
	            targetRange.Validation.Delete();
	        } catch (e) {
	            // 忽略删除错误
	        }
	        
	        // 最简化的写法
	        targetRange.Validation.Add({
	            Type: xlValidateList,  // xlValidateList
	            Operator: xlBetween,  // xlBetween
	            Formula1: "=" + Address1
	        });
	         // 智能设置默认值（选择第一个选项）
	        //targetRange.Value2 = rngSourceValue[0];
	        targetRange.Value2 = p.GetConfigValue(parameterName, "DefaultValue");
	        targetRange.HorizontalAlignment = -4131; // xlLeft
	        targetRange.VerticalAlignment = -4108;   // xlCenter
	        console.log("默认值设置为: " + rngSourceValue[0]);
	        targetRange.Interior.Color = this.p.m_COLOR_YELLOW;
	        //targetRange.value2 = dataArray[1]
	        
	        console.log(`[${this.MODULE_NAME}] 参数'${parameterName}'数据有效性设置完毕`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 设置参数数据有效性失败：${error.message}`);
	        return false;
	    }
	}
	// 获取利率描述
	获取利率描述() {
	    // 构建更详细的利率描述文本
	    
	    // 获取单元格地址
	    const lCell = p.GetCellAddressA1("LPRDate"); // LPR发布日期单元格地址
	    const pCell = p.GetCellAddressA1("LPRPeriod");
	    const rCell = p.GetCellAddressA1("LPRBenchmarkRate");
	    const bpCell = p.GetCellAddressA1("FloatingBasisPoints");
	    
	    const desc = `= "即 " & TEXT(${lCell},"yyyy年mm月dd日") & "全国银行间同业拆借中心公布的 " & ${pCell} & " 年期人民币贷款基础利率（LPR）" & ${rCell}*100 & "%，加" & ROUND(${bpCell},2) & "BPS（1BP=0.01%）"`;
	    
	    return desc;
	}
}


// ============== 快捷调用函数 ==============
function 测算表填写区域模块调用() {
    // 功能：重新初始化所有区域
    
    try {
        console.log("=== 开始测算表填写区域生成 ===");
        r = new rentCalculationFillinArea()
        r.main()
        
        // 重新初始化
        return ;
        
    } catch (error) {
        console.log(`测算表填写区域生成：${error.message}`);
        return false;
    }
}
// console.log(`[mInitialization] 模块加载完成 - 版本 ${VERSION}`);
// ============== 结束 ==============

