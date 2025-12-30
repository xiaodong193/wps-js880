Attribute Module_Name = "mCashFlowGenerator"
/**
 * ============== 现金流量表生成器类模块 ==============
 * 作者：徐晓冬
 * 描述：专业的现金流量表生成器，基于VBA最佳实践重构
 * 特性：数组批量处理、错误处理、性能监控、兼容性支持
 * WPS JS环境转换版本
 * ====================================================
 */

/**
 * CashFlowGenerator类 - 现金流量表生成器
 */
class CashFlowGenerator {
    /**
     * 构造函数
     */
    constructor() {
    	this.MODULE_NAME = "CashFlowGenerator";
        // 私有变量
        // 设置参数管理器
        this.p = (typeof p !== 'undefined') ? p : null;

        if (this.p === null) {
            throw new Error("参数管理器p未初始化，请确保mParameterManager.js已正确加载");
        }

        this.pWorksheet = this.p.m_worksheet;     // 目标工作表
        this.pIsInitialized = true;               // 初始化状态
        this.pTotalPeriods = this.p.TotalPeriodsCellValue;
        this.pCashFlowStartRow = this.p.CashFlowTablerowStart;  // 现金流表起始行
        this.pRentStartRow = this.p.RentTableStartRow; // 租金表起始行
        this.pPerformanceTimer = {
            startTime: 0,
            endTime: 0,
            duration: 0,
            operation: ""
        }; // 性能计时器
                    
        // 错误代码常量
        this.cfValidationError = 1001;
        this.cfCalculationError = 1002;
        this.cfDataMissingError = 1003;
        this.cfFormulaError = 1004;
        this.cfInitializationError = 1005;
        this.cfCompatibilityError = 1006;
        

        
        console.log(`[${this.MODULE_NAME}] 类实例创建`);
    }
    
    /**
     * 获取初始化状态
     */
    get IsInitialized() {
        return this.pIsInitialized;
    }
    
    /**
     * 获取总期数
     */
    get TotalPeriods() {
        return this.p.TotalPeriodsCellValue;
    }
    
    /**
     * 获取现金流表起始行
     */
    get CashFlowStartRow() {
        return this.p.CashFlowTablerowStart;
    }
    
    /**
     * 初始化现金流量表生成器
     */
    Initialize(sheetName = "1租金测算表V1") {
        try {
      
            this.endPerformanceTimer();
            console.log(`[${this.MODULE_NAME}] 初始化完成，总期数: ${this.pTotalPeriods}`);
            return true;
            
        } catch (error) {
            this.pIsInitialized = false;
            const errMsg = `现金流量表生成器初始化失败: ${error.message}`;
            console.log(`[${this.MODULE_NAME}] ${errMsg}`);
            throw new Error(errMsg);
        }
    }
    
    /**
     * 生成完整的现金流量表
     */
    GenerateCashFlowTable() {
        try {
            if (!this.pIsInitialized) {
                throw new Error("生成器未初始化");
            }
            
            this.startPerformanceTimer("现金流量表生成");
            
            // 步骤1：验证参数
            if (!this.ValidateParameters()) {
                return false;
            }
            
            // 步骤2：创建表头
            if (!this.CreateCashFlowHeaders()) {
                return false;
            }
            
            // 步骤3：生成现金流数据
            if (!this.GenerateCashFlowData()) {
                return false;
            }
            
            // 步骤4：处理经纪人费用
            if (!this.ProcessBrokerFees()) {
                return false;
            }
            
            // 步骤5：生成备注
            if (!this.GenerateRemarks()) {
                return false;
            }
            
            // 步骤6：应用格式
            if (!this.ApplyFormatting()) {
                return false;
            }
            
            this.endPerformanceTimer();
            console.log(`[${this.MODULE_NAME}] 现金流量表生成完成，耗时: ${this.pPerformanceTimer.duration.toFixed(3)} 秒`);
            return true;
            
        } catch (error) {
            this.HandleCashFlowError(error.name || "GenerateCashFlowTable", error.message);
            return false;
        }
    }
    
    /**
     * 验证参数
     */
    ValidateParameters() {
        try {
            // 验证总期数
            if (this.pTotalPeriods <= 0) {
                throw new Error("总期数必须大于0");
            }
            
            // 验证工作表
            if (this.pWorksheet === null) {
                throw new Error("工作表对象无效");
            }
            
            // 验证关键参数
            const principal = this.p.ReadParameterValue("Principal");
            if (principal <= 0) {
                throw new Error("租赁成本必须大于0");
            }
            
            console.log(`[${this.MODULE_NAME}] 参数验证通过`);
            return true;
        } catch (error) {
            this.HandleCashFlowError("ValidateParameters", error.message);
            return false;
        }
    }
    
    /**
     * 创建现金流量表表头
     */
    CreateCashFlowHeaders() {
        try {
            // 定义表头数组
            const headers = ["期次", "日期", "净现金流1（1+2+3+4+5）", "净现金流1-备注",
                            "净现金流2（1+2+3+4）", "净现金流2-备注", "（1）电汇放款",
                            "（2)租金偿付", "（3）保证金", "（4）名义货价", "（5）经纪人费用"];
            
            // 设置总标题
            const titleCell = this.p.m_worksheet.Range(`A${this.p.CashFlowTablerowStart - 2}`);
            titleCell.Value2 = "现金流及综合利率测算";
            titleCell.Interior.Color = this.p.m_COLOR_WHITE; // 使用共享常量
            titleCell.Font.Name = FONT_DEFAULT; // 使用共享常量
            titleCell.Font.Size = FONT_SIZE_TITLE; // 使用共享常量
            titleCell.Font.Color = this.p.m_COLOR_BLACK; // 使用共享常量
           	// 动态计算列范围
	        const lastCol = headers.length;
            // 设置表头
            const headerRange = this.p.m_worksheet.Range(
                this.p.m_worksheet.Cells(this.p.CashFlowTablerowStart - 1, 1),
                this.p.m_worksheet.Cells(this.p.CashFlowTablerowStart - 1, lastCol)
            );
            // 修正为二维数组格式
            const headerArray = [headers];
            headerRange.Value2 = headerArray;
            
            headerRange.Interior.Color = this.p.m_COLOR_BLUE; // 使用共享常量
            headerRange.Font.Name = FONT_DEFAULT; // 使用共享常量
            headerRange.Font.Size = FONT_SIZE_HEADER; // 使用共享常量
            headerRange.Font.Color = this.p.m_COLOR_BLACK; // 使用共享常量
            headerRange.HorizontalAlignment = xlCenter;
            headerRange.VerticalAlignment = xlCenter;
            headerRange.WrapText = true;
	        // 设置边框
	        headerRange.Borders.LineStyle = xlContinuous;
	        headerRange.Borders.Weight = xlThin;
	        headerRange.Borders.Color = this.p.m_COLOR_GRAY;
            
            console.log(`[${this.MODULE_NAME}] 表头创建完成`);
            return true;
        } catch (error) {
            this.HandleCashFlowError("CreateCashFlowHeaders", error.message);
            return false;
        }
    }
    
    /**
     * 生成现金流数据（数组批量处理）
     */
    GenerateCashFlowData() {
        try {
            // 创建二维数组存储现金流数据
            const cashFlowArray = [];
            for (let i = 0; i <= this.pTotalPeriods; i++) {
                cashFlowArray[i] = new Array(11);
            }
            
            // 批量生成现金流数据
            for (let i = 0; i <= this.pTotalPeriods; i++) {
                cashFlowArray[i][0] = i; // 期次
                
                // 日期
                if (i === 0) {
                    //cashFlowArray[i][1] = this.p.LeaseStartDateCellValue;
                    cashFlowArray[i][1] = `=${this.p.LeaseStartDateCellA1}`;
                } else {
                    //cashFlowArray[i][1] = this.p.m_worksheet.Range(`B${this.p.RentTableStartRow + i - 1}`).Value2;
                    cashFlowArray[i][1] = `=B${this.p.RentTableStartRow + i - 1}`;
                }
                
                // 净现金流公式
                cashFlowArray[i][2] = "=SUM(RC[4]:RC[8])"; // 净现金流1
                cashFlowArray[i][4] = "=SUM(RC[1]:RC[5])"; // 净现金流2
                
                // 现金流项目
                if (i === 0) {
                    // 第一期：放款 + 保证金收取
                    cashFlowArray[i][6] = `=-${this.p.PrincipalCellA1}`;
                    cashFlowArray[i][8] = `=${this.p.DepositCellA1}`;
                } else if (i < this.pTotalPeriods) {
                    // 中间期次：租金偿付
                    cashFlowArray[i][7] = `=C${this.p.RentTableStartRow + i - 1}`;
                } else {
                    // 最后一期：租金偿付 + 保证金退还 + 名义货价
                    cashFlowArray[i][7] = `=C${this.p.RentTableStartRow + i - 1}`;
                    cashFlowArray[i][8] = `=-${this.p.DepositCellA1}`;
                    cashFlowArray[i][9] = `=${this.p.NominalPriceCellA1}`;
                }
            }
            
            // 批量写入数据
            const targetRange = this.p.m_worksheet.Range(
                `A${this.p.CashFlowTablerowStart}`
            ).Resize(this.pTotalPeriods + 1, 11);
            targetRange.Value2 = cashFlowArray;
            
            console.log(`[${this.MODULE_NAME}] 现金流数据生成完成，共${this.pTotalPeriods + 1}期`);
            return true;
        } catch (error) {
            this.HandleCashFlowError("GenerateCashFlowData", error.message);
            return false;
        }
    }
    
    /**
     * 处理经纪人费用
     */
    ProcessBrokerFees() {
        try {
            // 读取经纪人费用参数
            const brokerPaymentMethod = this.p.BrokerPaymentMethodCellValue;
            const brokerFeeRate = this.p.BrokerFeeRateCellValue;
            const principal = this.p.principalCellValue;
            
            const brokerFeeRange = this.p.m_worksheet.Range(
                `K${this.p.CashFlowTablerowStart}:K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            
            // 根据支付方式设置经纪人费用
            switch (brokerPaymentMethod) {
                case "一次性支付-放款时":
                    this.p.m_worksheet.Range(`K${this.p.CashFlowTablerowStart}`).Formula = "=-$B$4*$G$16";
                    break;
                case "一次性支付-第1期租金":
                    this.p.m_worksheet.Range(`K${this.p.CashFlowTablerowStart + 1}`).Formula = "=-$B$4*$G$16";
                    break;
                case "分三次支付（第1\\中\\末期）":
                    const midPeriod = this.p.CashFlowTablerowStart + 1 + Math.round(this.p.TotalPeriodsCellValue / 2);
                    this.p.m_worksheet.Range(`K${this.p.CashFlowTablerowStart + 1}`).Formula = "=-$B$4*$G$16/3";
                    this.p.m_worksheet.Range(`K${midPeriod}`).Formula = "=-$B$4*$G$16/3";
                    this.p.m_worksheet.Range(`K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`).Formula = "=-$B$4*$G$16/3";
                    break;
            }
            
            // 设置格式
            brokerFeeRange.NumberFormat = FORMAT_STANDARD; // 使用共享常量
            
            console.log(`[${this.MODULE_NAME}] 经纪人费用处理完成，支付方式: ${brokerPaymentMethod}`);
            return true;
        } catch (error) {
            this.HandleCashFlowError("ProcessBrokerFees", error.message);
            return false;
        }
    }
    
    /**
     * 生成备注
     */
    GenerateRemarks() {
        try {
            // 读取现金流数据
            const cashFlowData = this.p.m_worksheet.Range(
                `G${this.p.CashFlowTablerowStart}:K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            ).Value2;
            
            // 生成备注数组
            const remark1Array = [];
            const remark2Array = [];
            
            // 生成备注
            for (let i = 0; i <= this.p.TotalPeriodsCellValue; i++) {
                remark1Array[i] = this.GenerateRemark1(cashFlowData, i + 1);
                remark2Array[i] = this.GenerateRemark2(cashFlowData, i + 1);
            }
            
            // 批量写入备注
            this.p.m_worksheet.Range(`D${this.p.CashFlowTablerowStart}`).Resize(this.p.TotalPeriodsCellValue + 1, 1).Value2 =
                this.transposeArray(remark1Array);
            this.p.m_worksheet.Range(`F${this.p.CashFlowTablerowStart}`).Resize(this.p.TotalPeriodsCellValue + 1, 1).Value2 =
                this.transposeArray(remark2Array);
            
            console.log(`[${this.MODULE_NAME}] 备注生成完成`);
            return true;
        } catch (error) {
            this.HandleCashFlowError("GenerateRemarks", error.message);
            return false;
        }
    }
    
    /**
     * 生成净现金流1备注
     */
    GenerateRemark1(cashFlowData, rowIndex) {
        let remark = "";
        const row = cashFlowData[rowIndex - 1] || [];
        
        if (row[0] !== undefined && row[0] !== 0) remark += "电汇放款/";
        if (row[1] !== undefined && row[1] !== 0) remark += `第${rowIndex - 1}期租金/`;
        if (row[2] > 0) {
            remark += "出租人收取保证金/";
        } else if (row[2] < 0) {
            remark += "出租人退还保证金/";
        }
        if (row[3] !== undefined && row[3] !== 0) remark += "出租人收取名义货价/";
        if (row[4] !== undefined && row[4] !== 0) remark += "经纪人费用/";
        
        return remark;
    }
    
    /**
     * 生成净现金流2备注
     */
    GenerateRemark2(cashFlowData, rowIndex) {
        let remark = "";
        const row = cashFlowData[rowIndex - 1] || [];
        
        if (row[0] !== undefined && row[0] !== 0) remark += "电汇放款/";
        if (row[1] !== undefined && row[1] !== 0) remark += `第${rowIndex - 1}期租金/`;
        if (row[2] > 0) {
            remark += "承租人支付保证金/";
        } else if (row[2] < 0) {
            remark += "承租人收回保证金/";
        }
        if (row[3] !== undefined && row[3] !== 0) remark += "出租人名义货价收取/";
        
        return remark;
    }
    
    /**
     * 转置数组
     */
    transposeArray(arr) {
        if (!Array.isArray(arr)) return arr;
        return arr.map(item => [item]);
    }
    
    /**
     * 应用格式
     */
    ApplyFormatting() {
        try {
            // Value2
            const dataRange = this.p.m_worksheet.Range(
                `A${this.p.CashFlowTablerowStart}:K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            
            设置表格样式(dataRange)
            // 设置日期格式
            const dateRange = this.p.m_worksheet.Range(
                `B${this.p.CashFlowTablerowStart}:B${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            //dateRange.NumberFormat = FORMAT_DATE; // 使用共享常量
            应用格式(dateRange, "Date")
            设置表格样式(dateRange)
            // 设置数字格式
            const numberRange = this.p.m_worksheet.Range(
                `C${this.p.CashFlowTablerowStart}:K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            //numberRange.NumberFormat = FORMAT_STANDARD; // 使用共享常量
            //numberRange.Font.Name = FONT_ENGLISH;
            //应用格式(numberRange, "Standard")
            设置表格样式(numberRange)
            const textRange = this.p.m_worksheet.Range(
                `D${this.p.CashFlowTablerowStart}:D${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue},` +
                `F${this.p.CashFlowTablerowStart}:F${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            应用格式(textRange, "Standard")
            设置表格样式(textRange)         
            console.log(`[${this.MODULE_NAME}] 格式应用完成`);
            //选取现金流量表全部区域，包括标题
            const allRange = this.p.m_worksheet.Range(
                `A${this.p.CashFlowTablerowStart - 1}:K${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue}`
            );
            this.添加框线(allRange);
            return true;
        } catch (error) {
            this.HandleCashFlowError("ApplyFormatting", error.message);
            return false;
        }
    }
    
    /**
     * 开始性能计时
     */
    startPerformanceTimer(operation) {
        this.pPerformanceTimer.startTime = new Date().getTime();
        this.pPerformanceTimer.operation = operation;
        console.log(`[${this.MODULE_NAME}] 开始: ${operation}`);
    }
    
    /**
     * 结束性能计时
     */
    endPerformanceTimer() {
        this.pPerformanceTimer.endTime = new Date().getTime();
        this.pPerformanceTimer.duration = (this.pPerformanceTimer.endTime - this.pPerformanceTimer.startTime) / 1000;
        console.log(`[${this.MODULE_NAME}] 完成: ${this.pPerformanceTimer.operation}，耗时: ${this.pPerformanceTimer.duration.toFixed(3)} 秒`);
    }
    
    /**
     * 处理现金流错误
     */
    HandleCashFlowError(errSource, errDescription) {
        const errorMsg = `[${errSource}] ${errDescription}`;
        console.log(`[${this.MODULE_NAME}] 错误: ${errorMsg}`);
        
        // 根据错误类型进行特定处理
        if (errDescription.includes("验证")) {
            alert(`现金流数据验证失败：${errDescription}`);
        } else if (errDescription.includes("计算")) {
            alert(`现金流计算失败：${errDescription}`);
        } else if (errDescription.includes("数据缺失")) {
            alert(`现金流数据缺失：${errDescription}`);
        } else {
            alert(`现金流生成失败：${errDescription}`);
        }
    }
    添加框线(rng) {
	    try {

	        rng.Borders.LineStyle = xlContinuous;
	        rng.Borders.Color = this.p.m_COLOR_BLACK;
	        rng.Borders.Weight = xlThin;
	        rng.Borders.TintAndShade = 0;
	        return true;
	    } catch (error) {
	        console.log(`添加框线失败：${error.message}`);
	        return false;
	    }
	}
	综合利率一览() {
	    try {
	        // 检查是否已初始化
	        if (p === null || !p.Isinitialized) {
	            p.Initialize();
	        }
	        
	        const titleCell = p.m_worksheet.Range("A15");
	        titleCell.Value2 = "综合利率一览";
	        titleCell.Interior.Color = this.p.m_COLOR_WHITE;
	        titleCell.Font.Name = FONT_DEFAULT;
	        titleCell.Font.Size = FONT_SIZE_TITLE;
	        titleCell.Font.Color = this.p.m_COLOR_BLACK;
	        
	        const xirr1Cell = p.ConvertR1C1ToA1("R16C4"); // D16
	        const xirr1Label = p.m_worksheet.Range(xirr1Cell).Offset(0, -1);
	        xirr1Label.Value2 = "XIRR净内含报酬率";
	        xirr1Label.Font.Name = FONT_DEFAULT;
	        xirr1Label.Font.Size = FONT_SIZE_HEADER;
	        xirr1Label.Font.Color = this.p.m_COLOR_BLACK;
	        
	        const xirr2Cell = p.ConvertR1C1ToA1("R17C4"); // D17
	        const xirr2Label = p.m_worksheet.Range(xirr2Cell).Offset(0, -1);
	        xirr2Label.Value2 = "（1）企业看XIRR";
	        xirr2Label.Font.Name = FONT_DEFAULT;
	        xirr2Label.Font.Size = FONT_SIZE_HEADER;
	        xirr2Label.Font.Color = this.p.m_COLOR_BLACK;
	        
	        const xirrDiffCell = p.ConvertR1C1ToA1("R18C4"); // D18
	        const xirrDiffLabel = p.m_worksheet.Range(xirrDiffCell).Offset(0, -1);
	        xirrDiffLabel.Value2 = "（2）经纪人费用影响";
	        xirrDiffLabel.Font.Name = FONT_DEFAULT;
	        xirrDiffLabel.Font.Size = FONT_SIZE_HEADER;
	        xirrDiffLabel.Font.Color = this.p.m_COLOR_BLACK;
	        
	        const irr1Cell = p.ConvertR1C1ToA1("R16C2"); // B16
	        const irr1Label = p.m_worksheet.Range(irr1Cell).Offset(0, -1);
	        irr1Label.Value2 = "IRR内含报酬率";
	        irr1Label.Font.Name = FONT_DEFAULT;
	        irr1Label.Font.Size = FONT_SIZE_HEADER;
	        irr1Label.Font.Color = this.p.m_COLOR_BLACK;
	        
	        const irr2Cell = p.ConvertR1C1ToA1("R17C2"); // B17
	        const irr2Label = p.m_worksheet.Range(irr2Cell).Offset(0, -1);
	        irr2Label.Value2 = "(1)企业看IRR";
	        irr2Label.Font.Name = FONT_DEFAULT;
	        irr2Label.Font.Size = FONT_SIZE_HEADER;
	        irr2Label.Font.Color = this.p.m_COLOR_BLACK;
	        
	        // 生成XIRR公式
	        let formula = `=XIRR(C${p.CashFlowTablerowStart}:C${p.CashFlowTablerowStart + p.TotalPeriodsCellValue},` +
	                     `B${p.CashFlowTablerowStart}:B${p.CashFlowTablerowStart + p.TotalPeriodsCellValue})`;
	        
	        const xirr1Range = p.m_worksheet.Range(xirr1Cell);
	        xirr1Range.Formula = formula;
	        xirr1Range.Font.Name = FONT_DEFAULT;
	        xirr1Range.Font.Size = FONT_SIZE_HEADER;
	        xirr1Range.Font.Color = this.p.m_COLOR_BLACK;
	        xirr1Range.Interior.Color = this.p.m_COLOR_LIGHT_GREEN;
	        
	        formula = `=XIRR(E${p.CashFlowTablerowStart}:E${p.CashFlowTablerowStart + p.TotalPeriodsCellValue},` +
	                 `B${p.CashFlowTablerowStart}:B${p.CashFlowTablerowStart + p.TotalPeriodsCellValue})`;
	        
	        const xirr2Range = p.m_worksheet.Range(xirr2Cell);
	        xirr2Range.Formula = formula;
	        xirr2Range.Font.Name = FONT_DEFAULT;
	        xirr2Range.Font.Size = FONT_SIZE_HEADER;
	        xirr2Range.Font.Color = this.p.m_COLOR_BLACK;
	        xirr2Range.Interior.Color = this.p.m_COLOR_LIGHT_GREEN;
	        
	        // 生成XIRR差异公式
	        formula = "=R[-2]C-R[-1]C";
	        const xirrDiffRange = p.m_worksheet.Range(xirrDiffCell);
	        xirrDiffRange.FormulaR1C1 = formula;
	        xirrDiffRange.Font.Name = FONT_DEFAULT;
	        xirrDiffRange.Font.Size = FONT_SIZE_HEADER;
	        xirrDiffRange.Font.Color = this.p.m_COLOR_BLACK;
	        xirrDiffRange.Interior.Color = this.p.m_COLOR_LIGHT_RED;
	        
	        // 生成IRR公式
	        formula = `=IRR(C${p.CashFlowTablerowStart}:C${p.CashFlowTablerowStart + p.TotalPeriodsCellValue})*${p.PaymentsPerYearCellA1}`;
	        const irr1Range = p.m_worksheet.Range(irr1Cell);
	        irr1Range.Formula = formula;
	        irr1Range.Font.Name = FONT_DEFAULT;
	        irr1Range.Font.Size = FONT_SIZE_HEADER;
	        irr1Range.Font.Color = this.p.m_COLOR_BLACK;
	        irr1Range.Interior.Color = this.p.m_COLOR_LIGHT_GREEN;
	        
	        formula = `=IRR(E${p.CashFlowTablerowStart}:E${p.CashFlowTablerowStart + p.TotalPeriodsCellValue})*${p.PaymentsPerYearCellA1}`;
	        const irr2Range = p.m_worksheet.Range(irr2Cell);
	        irr2Range.Formula = formula;
	        irr2Range.Font.Name = FONT_DEFAULT;
	        irr2Range.Font.Size = FONT_SIZE_HEADER;
	        irr2Range.Font.Color = this.p.m_COLOR_BLACK;
	        irr2Range.Interior.Color = this.p.m_COLOR_LIGHT_GREEN;
	        
	        // 设置数字格式
	        const formatRange = p.m_worksheet.Range(`${irr1Cell}:${irr2Cell},${xirr1Cell}:${xirrDiffCell}`);
	        formatRange.NumberFormatLocal = "0.00%";
	        
	        console.log(`[${MODULE_NAME}] 综合利率一览计算完成`);
	        return true;
	    } catch (error) {
	        console.log(`[${MODULE_NAME}] 综合利率一览计算失败：${error.message}`);
	        return false;
	    }
}
    /**
     * 检测是否为WPS环境
     */
    IsWPS() {
        try {
            const appName = Application.Name.toLowerCase();
            return appName.includes("wps");
        } catch (error) {
            return false;
        }
    }
    
    /**
     * 设置兼容性模式
     */
    SetCompatibilityMode() {
        if (this.IsWPS()) {
            Application.Calculation = xlCalculationManual;
            Application.ScreenUpdating = false;
        } else {
            Application.Calculation = xlCalculationAutomatic;
            Application.ScreenUpdating = true;
        }
    }
}
