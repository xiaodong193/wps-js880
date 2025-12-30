/**
 * ============== 参数管理器类模块 ==============
 * 作者：徐晓冬
 * 描述：增强版参数管理器，基于对象配置，支持动态参数管理
 * ====================================================
 */
// ============== 修正：RGB颜色函数 ==============

class ParameterManager {
    constructor() {
        this.MODULE_NAME = "ParameterManager";
        
        // 私有变量
        this.m_worksheet = null;
        this._isInitialized = false;
        this.m_config = this.CreateDefaultConfig();
        this.targetSheetName = null;
        this.m_staticValueConversion = false;
        
        this.m_sourceSheetName = "1租金测算表V1";// 1租金测算表V1表格名称
        this.m_sourceSheet = Application.Worksheets(this.m_sourceSheetName);// 1租金测算表V1表格
        this.m_billSheetName = "银行承兑汇票";//银行承兑汇票表格名称
        this.m_billSheet = Application.Worksheets(this.m_billSheetName);;//银行承兑汇票表格
        this.m_repaymentSettingSheetName = "还款设置";//还款设置表格名称
        this.m_repaymentSettingSheet = Application.Worksheets(this.m_repaymentSettingSheetName);//还款设置表格
        this.m_loanRateSheetName = "贷款基础利率";//贷款基础利率表格名称
        this.m_loanRateSheet = Application.Worksheets(this.m_loanRateSheetName);//贷款基础利率表格

        // 表格结构相关常量
        this.m_RowStart = 26; // 租金表标题起始行
        this.m_SheetNamerow = 1; // 表头行数
        // ============== 尺寸常量 ==============
		this.m_MinRowHeight = 15;      // 最小行高
		this.m_MinColumnWidth = 20;   // 最小列宽


        
        // 列定义常量
        this.m_COL_PERIOD = "A";
        this.m_COL_DATE = "B";
        this.m_COL_RENT = "C";
        this.m_COL_PRINCIPAL = "D";
        this.m_COL_INTEREST = "E";
        this.m_COL_RENT_BALANCE = "F";
        this.m_COL_PRINCIPAL_BALANCE = "G";
        this.m_COL_REMAINING_BALANCE = "H";
        this.m_COL_PAID_RENT = "I";
        this.m_COL_MONTH_INTERVAL = "J";
        this.m_COL_CUSTOM_INTERVAL = "K";
        this.m_COL_PRINCIPAL_RATIO = "L";
        this.m_COL_RatePerPeriod = "M";
        
        this.m_COLOR_WHITE = this.RGB(255, 255, 255);     // RGB(255, 255, 255) 白色
        this.m_COLOR_BLUE = this.RGB(0, 174, 240);      // RGB(0, 174, 240) 蓝色
        this.m_COLOR_YELLOW = this.RGB(255, 255, 0);       // RGB(255, 255, 0) 黄色
        this.m_COLOR_LIGHT_GREEN = this.RGB(204, 255, 204); // RGB(204, 255, 204) 浅绿色
        this.m_COLOR_LIGHT_RED = this.RGB(255, 204, 204);   // RGB(255, 204, 204) 浅红色
        this.m_COLOR_BLACK = this.RGB(0, 0, 0);            // RGB(0, 0, 0) 黑色
        this.m_COLOR_GRAY = this.RGB(128, 128, 128);       // RGB(128, 128, 128) 灰色
        console.log("[" + this.MODULE_NAME + "] 类实例创建");
    }

    // ============== 属性定义 ==============

    // 获取工作表对象
    get WsTarget() {
        if (!this._isInitialized) {
            throw new Error("参数管理器未初始化");
        }
        return this.m_worksheet;
    }
    // 获取初始化状态
    get IsInitialized() {
        return this._isInitialized;
    }

    // 表格构成要素属性
    get RowStart() {
        return this.m_RowStart;
    }
	
    get SheetNamerow() {
        return this.m_SheetNamerow;
    }

    // 租金测算表数据第1行Row值
    get RentTableStartRow() {
        return this.m_RowStart + this.m_SheetNamerow * 2;
    }

    get CashFlowTablerowStart() {
        return this.RentTableStartRow + this.TotalPeriodsCellValue + 6;
    }

    // ============== 初始化方法 ==============

    /**
     * 初始化参数管理器
     * @param {string|Worksheet} sheetName - 工作表名称或Worksheet对象
     * @param {Object} config - 参数配置对象
     */
	// 修改：原来的简单赋值改为调用新方法
	Initialize(sheetName = "1租金测算表V1") {
	    try {

	        console.log(`[${this.MODULE_NAME}] 开始初始化工作表`);
	        
	        // 使用GetOrCreateWorksheet替代直接获取
	        this.m_worksheet = this.GetOrCreateWorksheet(sheetName);
	        this.targetSheetName = this.m_worksheet.Name;
	        
	        this._isInitialized = true;
	        console.log(`[${this.MODULE_NAME}] 初始化成功，工作表: ${this.targetSheetName}`);
	        return true;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 初始化失败: ${error.message}`);
	        return false;
	    }
	}

    // 检查工作表是否存在的辅助函数
    WorksheetExists(sheetName) {
        try {
            const ws = Application.Worksheets(sheetName);
            return true;
        } catch (error) {
            return false;
        }
    }


    // ============== 配置管理 ==============

    /**
     * 创建默认配置
     */
    CreateDefaultConfig() {
        const config = {};
        
        // ==================== 价格参数配置 ====================
        config.Principal = {
            CellAddress: "R4C2",
            DefaultValue: 100000000,
            VbaFormat: "#,##0.00",
            DisplayName: "租赁成本",
            ValidationRule: ">0",
            DataType: "Double",
            IsRequired: true,
            Description: "租赁资产的总成本"
        };

        config.InterestRate = {
            CellAddress: "R5C2",
            DefaultValue: 0.03,
            VbaFormat: "0.0000%",
            DisplayName: "租赁票面利率",
            ValidationRule: ">0",
            DataType: "Double",
            IsRequired: true,
            Description: "租赁合同约定的年利率"
        };

        config.Deposit = {
            CellAddress: "R6C2",
            DefaultValue: 0,
            VbaFormat: "#,##0.00",
            DisplayName: "租赁保证金",
            ValidationRule: ">=0",
            DataType: "Double",
            IsRequired: false,
            Description: "租赁保证金金额",
            DefaultFormula: "=(B4*D6)"
        };

        config.DepositMarginRate = {
            CellAddress: "R6C4",
            DefaultValue: 0.01,
            VbaFormat: "0.00%",
            DisplayName: "保证金费率",
            ValidationRule: ">=0",
            DataType: "Double",
            IsRequired: false,
            Description: "租赁保证金金额"
        };

        config.NominalPrice = {
            CellAddress: "R8C2",
            DefaultValue: 1,
            VbaFormat: "#,##0.00",
            DisplayName: "名义货价",
            ValidationRule: ">=0",
            DataType: "Double",
            IsRequired: true,
            Description: "租赁合同中的名义货价"
        };

        config.TotalPeriods = {
            CellAddress: "R10C2",
            DefaultValue: 10,
            VbaFormat: "0",
            DisplayName: "总期数",
            ValidationRule: ">0",
            DataType: "Long",
            IsRequired: true,
            Description: "租金测算的总期数"
        };

        config.PaymentsPerYear = {
            CellAddress: "R11C2",
            DefaultValue: 2,
            VbaFormat: "0",
            DisplayName: "每年还款次数",
            ValidationRule: ">0",
            DataType: "Long",
            IsRequired: true,
            Description: "每年的还款次数",
            DefaultFormula: "=12/D10"
        };

        config.RepaymentMethod = {
            CellAddress: "R12C2",
            DefaultValue: "等额本息（后付）",
            VbaFormat: "",
            DisplayName: "偿还方式",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: true,
            Description: "租金偿还方式"
        };

        config.LeaseStartDate = {
            CellAddress: "R13C2",
            DefaultValue: this.GetDefaultDate(),
            VbaFormat: "yyyy-mm-dd",
            DisplayName: "预算起租日（放款日）",
            ValidationRule: "TRUE",
            DataType: "Date",
            IsRequired: true,
            Description: "租赁合同约定的起租日期"
        };

        config.PaymentInterval = {
            CellAddress: "R10C4",
            DefaultValue: 6,
            VbaFormat: "0",
            DisplayName: "支付间隔",
            ValidationRule: ">0",
            DataType: "Long",
            IsRequired: true,
            Description: "支付间隔(月)"
        };

        config.ProjectDurationYears = {
            CellAddress: "R11C4",
            DefaultValue: 3,
            VbaFormat: "0",
            DisplayName: "项目时长/年",
            ValidationRule: ">0",
            DataType: "Long",
            IsRequired: false,
            Description: "项目时长/年",
            DefaultFormula: "=D10*B10/12"
        };

        config.ProjectDurationMonths = {
            CellAddress: "R12C4",
            DefaultValue: 36,
            VbaFormat: "0",
            DisplayName: "项目时长/月",
            ValidationRule: ">0",
            DataType: "Long",
            IsRequired: false,
            Description: "项目时长/月",
            DefaultFormula: "=D10*B10"
        };

        config.StaticValueConversion = {
            CellAddress: "R1C4",
            DefaultValue: false,
            VbaFormat: "General",
            DisplayName: "静态值转换",
            ValidationRule: "",
            DataType: "Boolean",
            IsRequired: false,
            Description: "是否将公式转换为静态值"
        };

        // ==================== 利率参数配置 ====================
        config.LPRDate = {
            CellAddress: "R11C9",
            DefaultValue: "",
            VbaFormat: "yyyy-mm-dd",
            DisplayName: "LPR发布日期",
            ValidationRule: "TRUE",
            DataType: "Date",
            IsRequired: true,
            Description: "LPR利率发布日期"
        };

        config.LPRBenchmarkRate = {
            CellAddress: "R10C7",
            DefaultValue: 0.0425,
            VbaFormat: "0.00%",
            DisplayName: "LPR基准利率",
            ValidationRule: ">=0,<=0.2",
            DataType: "Double",
            IsRequired: false,
            Description: "LPR基准利率",
            FormulaDependency: "LPRPeriod,LPRDate",
            DataSource: "贷款基础利率!$B$2:$D$14",
            IsCalculated: true,
            DefaultFormula: "=IFS(G11=5,VLOOKUP(I11,贷款基础利率!$B$2:$D$14,3,FALSE)/100,G11=1,VLOOKUP(I11,贷款基础利率!$B$2:$D$14,2,FALSE)/100)"
        };

        config.LPRPeriod = {
            CellAddress: "R11C7",
            DefaultValue: 5,
            VbaFormat: "0",
            DisplayName: "LPR期限选择",
            ValidationRule: "1,5",
            DataType: "Long",
            IsRequired: true,
            Description: "LPR期限选择（1=1年期，5=5年期）"
        };

        config.FloatingBasisPoints = {
            CellAddress: "R10C9",
            DefaultValue: 0,
            VbaFormat: "0.00",
            DisplayName: "浮动基点(BP)",
            ValidationRule: "",
            DataType: "Double",
            IsRequired: false,
            Description: "相对于LPR的浮动基点数",
            DefaultFormula: "=(B5-G10)*10000"
        };

        config.RateOption = {
            CellAddress: "R5C4",
            DefaultValue: "固定利率",
            VbaFormat: "",
            DisplayName: "利率选择",
            ValidationRule: "",
            DataType: "String",
            IsRequired: false,
            Description: "利率选择",
            DataSource: "还款设置!$K$2:$K$3",
            IsCalculated: true
        };

        config.ActualPayment = {
            CellAddress: "R4C4",
            DefaultValue: 0,
            VbaFormat: "#,##0.00",
            DisplayName: "实际付款",
            ValidationRule: "",
            DataType: "Double",
            IsRequired: false,
            Description: "实际付款是租赁本金减去保证金",
            FormulaDependency: "Principal,Deposit",
            IsCalculated: true,
            DefaultFormula: "=B4-B6"
        };

        config.LPRRateDescription = {
            CellAddress: "R12C7",
            DefaultValue: "",
            VbaFormat: "",
            DisplayName: "LPR利率描述",
            ValidationRule: "",
            DataType: "String",
            IsRequired: false,
            Description: "LPR利率描述",
            FormulaDependency: "LPRPeriod,LPRBenchmarkRate",
            IsCalculated: true,
            DefaultFormula: ""
        };

        config.ActualInterestRate = {
            CellAddress: "R10C8",
            DefaultValue: 0.03,
            VbaFormat: "0.00%",
            DisplayName: "实际利率",
            ValidationRule: ">0",
            DataType: "Double",
            IsRequired: false,
            Description: "实际执行利率",
            FormulaDependency: "LPRBenchmarkRate,FloatingBasisPoints",
            IsCalculated: true,
            DefaultFormula: "=G10+I10/10000"
        };

        // ==================== 方案要素配置 ====================
        config.Lessee = {
            CellAddress: "R5C7",
            DefaultValue: "",
            VbaFormat: "",
            DisplayName: "承租人",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: false,
            Description: "承租人名称"
        };

        config.Guarantor = {
            CellAddress: "R6C7",
            DefaultValue: "",
            VbaFormat: "",
            DisplayName: "担保人",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: false,
            Description: "担保人名称"
        };

        config.GuaranteeMethod = {
            CellAddress: "R7C7",
            DefaultValue: "",
            VbaFormat: "",
            DisplayName: "担保方式",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: false,
            Description: "担保方式"
        };

        config.LeaseMethod = {
            CellAddress: "R8C7",
            DefaultValue: "",
            VbaFormat: "",
            DisplayName: "租赁方式",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: false,
            Description: "租赁方式"
        };

        // ==================== 经纪人参数配置 ====================
        config.BrokerPaymentMethod = {
            CellAddress: "R14C7",
            DefaultValue: "一次性支付-放款时",
            VbaFormat: "",
            DisplayName: "经纪人费用支付方式",
            ValidationRule: "TRUE",
            DataType: "String",
            IsRequired: true,
            Description: "经纪人费用的支付方式"
        };

        config.BrokerFeeRate = {
            CellAddress: "R16C7",
            DefaultValue: 0.001,
            VbaFormat: "0.00%",
            DisplayName: "经纪人费用比例",
            ValidationRule: ">=0",
            DataType: "Double",
            IsRequired: true,
            Description: "经纪人费用比例（如0.02表示2%）"
        };

        config.BrokerTotalFee = {
            CellAddress: "R15C7",
            DefaultValue: 0,
            VbaFormat: "#,##0.00",
            DisplayName: "经纪人总费用",
            ValidationRule: ">=0",
            DataType: "Long",
            IsRequired: false,
            Description: "经纪人总费用(BrokerTotalFee)",
            DefaultFormula: "=B4*G16"
        };

        return config;
    }
    //转换颜色值
	RGB(r, g, b) { return r + g * 256 + b * 65536; }
    // 获取默认日期（Date后1个月的第15日）
    GetDefaultDate() {
        const now = new Date();
        const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 15);
        
        // 格式化为 YYYY-MM-DD
        const year = nextMonth.getFullYear();
        const month = String(nextMonth.getMonth() + 1).padStart(2, '0');
        const day = String(nextMonth.getDate()).padStart(2, '0');
        
        return `${year}-${month}-${day}`;
    }
	// ============== 修正：CreateParameterInputArea方法 ==============
	CreateParameterInputArea() {
	    try {
	        if (!this._isInitialized) {
	            console.log(`[${this.MODULE_NAME}] 参数管理器未初始化`);
	            return false;
	        }
	        
	        console.log(`[${this.MODULE_NAME}] 开始创建参数输入区域`);
	        
	        let row = 13; // 从第13行开始
	        
	        // 添加参数区域标题
	        try {
	            const areaTitle = this.m_worksheet.Range(`A${row}`);
	            areaTitle.Value2 = "参数输入区域";
	            areaTitle.Font.Bold = true;
	            areaTitle.Font.Size = 14;
	            
	            try {
	                areaTitle.Interior.Color = this.RGB(173, 216, 230); // 浅蓝色背景
	            } catch (error) {
	                console.log(`[${this.MODULE_NAME}] 无法设置标题背景色: ${error.message}`);
	            }
	            
	            this.m_worksheet.Range(`A${row}:C${row}`).Merge();
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 设置参数区域标题失败: ${error.message}`);
	        }
	        
	        row += 2;
	        
	        // 为每个参数创建输入行
	        const paramNames = this.GetParameterNames();
	        let createdCount = 0;
	        
	        for (const paramName of paramNames) {
	            const config = this.m_config[paramName];
	            
	            try {
	                // 参数名称
	                const nameCell = this.m_worksheet.Range(`A${row}`);
	                nameCell.Value2 = config.DisplayName;
	                nameCell.Font.Bold = true;
	                
	                // 参数值单元格
	                const valueCell = this.m_worksheet.Range(`B${row}`);
	                valueCell.Value2 = config.DefaultValue;
	                
	                // 设置单元格格式
	                if (config.VbaFormat) {
	                    try {
	                        valueCell.NumberFormat = config.VbaFormat;
	                    } catch (error) {
	                        console.log(`[${this.MODULE_NAME}] 设置格式失败 ${paramName}: ${error.message}`);
	                    }
	                }
	                
	                // 说明文字
	                const descCell = this.m_worksheet.Range(`C${row}`);
	                descCell.Value2 = config.Description || "";
	                try {
	                    descCell.Font.Color = this.RGB(128, 128, 128); // 灰色文字
	                    descCell.Font.Italic = true;
	                } catch (error) {
	                    console.log(`[${this.MODULE_NAME}] 设置说明文字样式失败: ${error.message}`);
	                }
	                
	                // 更新配置中的单元格地址（使用A1格式）
	                const a1Address = `B${row}`;
	                config.CellAddressA1 = a1Address;
	                
	                // 添加边框
	                try {
	                    this.m_worksheet.Range(`A${row}:C${row}`).Borders.LineStyle = 1;
	                } catch (error) {
	                    console.log(`[${this.MODULE_NAME}] 设置边框失败: ${error.message}`);
	                }
	                
	                createdCount++;
	                
	            } catch (error) {
	                console.log(`[${this.MODULE_NAME}] 创建参数行失败 ${paramName}: ${error.message}`);
	            }
	            
	            row++;
	        }
	        
	        console.log(`[${this.MODULE_NAME}] 参数输入区域创建完成，成功创建 ${createdCount}/${paramNames.length} 个参数`);
	        return createdCount > 0;
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 创建参数输入区域失败: ${error.message}`);
	        return false;
	    }
	}

    /**
     * 在新建的工作表中创建参数输入区域
     * @returns {boolean} 是否创建成功
     */	
     
    // ============== 地址转换方法 ==============

    /**
     * 根据参数名称获取单元格的A1样式地址
     * @param {string} parameterName - 参数名称
     * @returns {string} A1样式单元格地址
     */
    GetCellAddressA1(parameterName) {
        const r1c1Address = this.GetConfigValue(parameterName, "CellAddress");
        if (!r1c1Address) {
            return "";
        }
        return this.ConvertR1C1ToA1(r1c1Address, parameterName);
    }

    /**
     * 将R1C1样式地址转换为A1样式地址
     * @param {string} r1c1Address - R1C1样式地址
     * @param {string} parameterName - 参数名称
     * @returns {string} A1样式地址
     */
    ConvertR1C1ToA1(r1c1Address, parameterName) {
        try {
            // 解析R1C1格式
            const rPos = r1c1Address.indexOf("R");
            const cPos = r1c1Address.indexOf("C");
            
            if (rPos === -1 || cPos === -1) {
                return this.GetDefaultCellAddress(parameterName);
            }
            
            const rowPart = r1c1Address.substring(rPos + 1, cPos);
            const colPart = r1c1Address.substring(cPos + 1);
            
            const rowNumber = parseInt(rowPart);
            const colNumber = parseInt(colPart);
            
            // 转换为A1样式
            return this.m_worksheet.Cells(rowNumber, colNumber).Address(false, false);
            
        } catch (error) {
            return this.GetDefaultCellAddress(parameterName);
        }
    }

    // 获取参数对应的默认单元格地址
    GetDefaultCellAddress(parameterName) {
        if (parameterName && this._isInitialized) {
            if (this.m_config[parameterName]) {
                const address = this.GetConfigValue(parameterName, "CellAddress");
                console.log("[" + this.MODULE_NAME + "] GetDefaultCellAddress: 参数'" + parameterName + "'返回配置地址: " + address);
                return address;
            }
        }
        console.log("[" + this.MODULE_NAME + "] GetDefaultCellAddress: 无法获取参数地址，返回空字符串");
        return "";
    }
	// ============== 修正：获取或创建工作表方法 ==============
	GetOrCreateWorksheet(sheetName) {
	    try {
	        // 尝试获取现有工作表
	        let worksheet = null;
	        try {
	            worksheet = Application.Sheets(sheetName);
	            console.log(`[${this.MODULE_NAME}] 找到现有工作表: ${sheetName}`);
	            return worksheet;
	        } catch (error) {
	            // 工作表不存在，创建新工作表
	            console.log(`[${this.MODULE_NAME}] 工作表不存在，正在创建: ${sheetName}`);
	            worksheet = Application.Sheets.Add();
	            worksheet.Name = sheetName;
	            
	            // 等待工作表完全创建
	            Application.Calculate();
	            
	            // 尝试初始化格式，如果失败就继续使用工作表
	            try {
	                this.InitializeWorksheetFormat(worksheet);
	            } catch (formatError) {
	                console.log(`[${this.MODULE_NAME}] 格式初始化失败，但工作表已创建: ${formatError.message}`);
	            }
	            
	            return worksheet;
	        }
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 获取或创建工作表失败: ${error.message}`);
	        return null;
	    }
	}
	// 新增：初始化工作表的格式和结构
	InitializeWorksheetFormat(worksheet) {
	    try {
	        console.log(`[${this.MODULE_NAME}] 开始初始化工作表格式`);
	        
	        // 先检查工作表是否可写
	        try {
	            // 测试写入一个临时值
	            const testCell = worksheet.Range("A1");
	            testCell.Value2 = "测试";
	            if (testCell.Value2 !== "测试") {
	                throw new Error("工作表只读");
	            }
	            // 清理测试值
	            testCell.Value2 = "";
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 工作表可能为只读状态，跳过格式设置`);
	            return; // 如果是只读的，跳过格式设置
	        }
	        
	        // 清空单元格（如果可写）
	        try {
	            worksheet.Cells.Clear();
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 无法清空单元格: ${error.message}`);
	        }
	        
	        // 添加标题行
	        try {
	            const titleCell = worksheet.Range("A1");
	            titleCell.Value2 = "参数配置表";
	            titleCell.Font.Bold = true;
	            titleCell.Font.Size = 16;
	            
	            // 合并标题单元格
	            worksheet.Range("A1:C1").Merge();
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 设置标题失败: ${error.message}`);
	        }
	        
	        // 添加表头
	        try {
	            const headers = ["参数名称", "参数值", "说明"];
	            for (let i = 0; i < headers.length; i++) {
	                const headerCell = worksheet.Range(`A${i + 3}`);
	                headerCell.Value2 = headers[i];
	                headerCell.Font.Bold = true;
	                try {
	                    headerCell.Interior.Color = this.RGB(200, 200, 200);
	                } catch (error) {
	                    console.log(`[${this.MODULE_NAME}] 设置表头颜色失败: ${error.message}`);
	                }
	            }
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 设置表头失败: ${error.message}`);
	        }
	        
	        // 设置列宽
	        try {
	            worksheet.Columns("A:A").ColumnWidth = 20;
	            worksheet.Columns("B:B").ColumnWidth = 15;
	            worksheet.Columns("C:C").ColumnWidth = 30;
	        } catch (error) {
	            console.log(`[${this.MODULE_NAME}] 设置列宽失败: ${error.message}`);
	        }
	        
	        console.log(`[${this.MODULE_NAME}] 工作表格式初始化完成`);
	        
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 工作表格式初始化失败: ${error.message}`);
	    }
	}

    // ============== 配置值读取方法 ==============

    /**
     * 读取配置字典中的值
     * @param {string} parameterName - 参数名称
     * @param {string} configKey - 配置键名
     * @returns {any} 配置值
     */
    GetConfigValue(parameterName, configKey) {
        try {
            if (!this._isInitialized) {
                return "";
            }

            // 检查参数是否存在
            if (!this.m_config[parameterName]) {
                return "";
            }

            const paramConfig = this.m_config[parameterName];
            
            // 检查配置键是否存在
            if (paramConfig.hasOwnProperty(configKey)) {
                return paramConfig[configKey];
            } else {
                return "";
            }
            
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 读取配置值出错: " + error.message + " (参数: " + parameterName + ", 键: " + configKey + ")");
            return "";
        }
    }

    // ============== 参数值读取方法 ==============

    /**
     * 从工作表读取指定参数的值
     * @param {string} parameterName - 参数名称
     * @returns {any} 参数值
     */
    ReadParameterValue(parameterName) {
        try {
            // 1. 基础验证
            if (!this.m_worksheet) {
                return null;
            }

            // 2. 获取参数配置
            if (!this.m_config[parameterName]) {
                return null;
            }

            const paramConfig = this.m_config[parameterName];
            const defaultValue = paramConfig.DefaultValue;
            const dataType = paramConfig.DataType;
            const validationRule = paramConfig.ValidationRule;

            // 3. 读取单元格值
            const cellAddressA1 = this.ConvertR1C1ToA1(paramConfig.CellAddress, parameterName);
            
            if (!cellAddressA1) {
                return this.GetSimpleDefaultValue(defaultValue, dataType);
            }

            const cellValue = this.m_worksheet.Range(cellAddressA1).Value2;

            // 4. 处理空值情况
            if (cellValue === null || cellValue === undefined || cellValue === "") {
                return this.GetSimpleDefaultValue(defaultValue, dataType);
            }

            // 5. 根据数据类型进行转换并验证
            let resultValue;
            switch (dataType) {
                case "Long":
                    if (!isNaN(cellValue)) {
                        resultValue = parseInt(cellValue);
                    } else {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    if (!this.ValidateValueByRule(resultValue, validationRule)) {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    break;
                    
                case "Double":
                    if (!isNaN(cellValue)) {
                        resultValue = parseFloat(cellValue);
                    } else {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    if (!this.ValidateValueByRule(resultValue, validationRule)) {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    break;
                    
                case "Date":
                    if (this.IsValidDate(cellValue)) {
                        resultValue = new Date(cellValue);
                    } else {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    if (!this.ValidateValueByRule(resultValue, validationRule)) {
                        resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                    }
                    break;
                    
                case "Boolean":
                    const strValue = String(cellValue).toUpperCase().trim();
                    switch (strValue) {
                        case "TRUE": case "1": case "是": case "YES": case "Y": case "T":
                            resultValue = true;
                            break;
                        case "FALSE": case "0": case "否": case "NO": case "N": case "F":
                            resultValue = false;
                            break;
                        default:
                            if (!isNaN(cellValue)) {
                                resultValue = parseInt(cellValue) === 1;
                            } else {
                                resultValue = this.GetSimpleDefaultValue(defaultValue, dataType);
                            }
                    }
                    break;
                    
                case "String":
                    resultValue = String(cellValue);
                    break;
                    
                default:
                    resultValue = cellValue;
            }

            return resultValue;
            
        } catch (error) {
            const defaultValue = this.m_config[parameterName]?.DefaultValue;
            const dataType = this.m_config[parameterName]?.DataType;
            return this.GetSimpleDefaultValue(defaultValue, dataType);
        }
    }

    // 检查是否为有效日期
    IsValidDate(date) {
        return date instanceof Date && !isNaN(date);
    }

    // 默认值获取函数
    GetSimpleDefaultValue(defaultValue, dataType) {
        try {
            switch (dataType) {
                case "Long":
                    return !isNaN(defaultValue) ? parseInt(defaultValue) : 0;
                case "Double":
                    return !isNaN(defaultValue) ? parseFloat(defaultValue) : 0.0;
                case "Date":
                    return this.IsValidDate(defaultValue) ? defaultValue : new Date();
                case "Boolean":
                    return defaultValue === true || defaultValue === "TRUE" || defaultValue === "1";
                case "String":
                    return defaultValue !== null && defaultValue !== undefined ? String(defaultValue) : "";
                default:
                    return defaultValue;
            }
        } catch (error) {
            switch (dataType) {
                case "Long": return 0;
                case "Double": return 0.0;
                case "Date": return new Date();
                case "Boolean": return false;
                case "String": return "";
                default: return null;
            }
        }
    }

    // ============== 验证函数 ==============

    /**
     * 根据规则验证值
     * @param {any} value - 要验证的值
     * @param {string} rule - 验证规则
     * @returns {boolean} 验证是否通过
     */
    ValidateValueByRule(value, rule) {
        if (!rule) return true; // 没有规则则通过

        switch (rule) {
            case ">0":
                return !isNaN(value) && value > 0;
            case ">=0":
                return !isNaN(value) && value >= 0;
            case "Boolean":
                return typeof value === 'boolean' || 
                       value === "TRUE" || value === "FALSE" || 
                       value === "1" || value === "0";
            case "1,5":
                return value === 1 || value === 5;
            default:
                return true;
        }
    }

    // ============== 参数属性访问器 ==============

    // 价格参数属性
    get PrincipalCellR1C1() { return this.GetConfigValue("Principal", "CellAddress"); }
    get PrincipalCellA1() { return this.GetCellAddressA1("Principal"); }
    get principalCellValue() { return this.ReadParameterValue("Principal"); }

    get ActualPaymentCellR1C1() { return this.GetConfigValue("ActualPayment", "CellAddress"); }
    get ActualPaymentCellA1() { return this.GetCellAddressA1("ActualPayment"); }

    get InterestRateCellR1C1() { return this.GetConfigValue("InterestRate", "CellAddress"); }
    get InterestRateCellA1() { return this.GetCellAddressA1("InterestRate"); }
    get InterestRateCellValue() { return this.ReadParameterValue("InterestRate"); }
    get DepositCellR1C1() { return this.GetConfigValue("Deposit", "CellAddress"); }
    get DepositCellA1() { return this.GetCellAddressA1("Deposit"); }

    get DepositMarginRateCellR1C1() { return this.GetConfigValue("DepositMarginRate", "CellAddress"); }
    get DepositMarginRateCellA1() { return this.GetCellAddressA1("DepositMarginRate"); }

    get NominalPriceCellR1C1() { return this.GetConfigValue("NominalPrice", "CellAddress"); }
    get NominalPriceCellA1() { return this.GetCellAddressA1("NominalPrice"); }

    get TotalPeriodsCellR1C1() { return this.GetConfigValue("TotalPeriods", "CellAddress"); }
    get TotalPeriodsCellA1() { return this.GetCellAddressA1("TotalPeriods"); }
    get TotalPeriodsCellValue() { return this.ReadParameterValue("TotalPeriods"); }

    get PaymentsPerYearCellR1C1() { return this.GetConfigValue("PaymentsPerYear", "CellAddress"); }
    get PaymentsPerYearCellA1() { return this.GetCellAddressA1("PaymentsPerYear"); }
    get PaymentsPerYearValue() {return this.ReadParameterValue("PaymentsPerYear")}

    get RepaymentMethodCellR1C1() { return this.GetConfigValue("RepaymentMethod", "CellAddress"); }
    get RepaymentMethodCellA1() { return this.GetCellAddressA1("RepaymentMethod"); }
    get RepaymentMethodCellValue() { return this.ReadParameterValue("RepaymentMethod"); }

    get LeaseStartDateCellR1C1() { return this.GetConfigValue("LeaseStartDate", "CellAddress"); }
    get LeaseStartDateCellA1() { return this.GetCellAddressA1("LeaseStartDate"); }
    get LeaseStartDateCellValue() { return this.ReadParameterValue("LeaseStartDate"); }

    get PaymentIntervalCellR1C1() { return this.GetConfigValue("PaymentInterval", "CellAddress"); }
    get PaymentIntervalCellA1() { return this.GetCellAddressA1("PaymentInterval"); }
    get PaymentIntervalCellValue() { return this.ReadParameterValue("PaymentInterval"); }

    // 项目参数属性
    get ProjectDurationYearsCellR1C1() { return this.GetConfigValue("ProjectDurationYears", "CellAddress"); }
    get ProjectDurationYearsCellA1() { return this.GetCellAddressA1("ProjectDurationYears"); }

    get ProjectDurationMonthsCellR1C1() { return this.GetConfigValue("ProjectDurationMonths", "CellAddress"); }
    get ProjectDurationMonthsCellA1() { return this.GetCellAddressA1("ProjectDurationMonths"); }

    get StaticValueConversionCellR1C1() { return this.GetConfigValue("StaticValueConversion", "CellAddress"); }
    get StaticValueConversionCellA1() { return this.GetCellAddressA1("StaticValueConversion"); }

    // 利率参数属性
    get LPRDateCellR1C1() { return this.GetConfigValue("LPRDate", "CellAddress"); }
    get LPRDateCellA1() { return this.GetCellAddressA1("LPRDate"); }

    get LPRBenchmarkRateCellR1C1() { return this.GetConfigValue("LPRBenchmarkRate", "CellAddress"); }
    get LPRBenchmarkRateCellA1() { return this.GetCellAddressA1("LPRBenchmarkRate"); }

    get LPRPeriodCellR1C1() { return this.GetConfigValue("LPRPeriod", "CellAddress"); }
    get LPRPeriodCellA1() { return this.GetCellAddressA1("LPRPeriod"); }

    get FloatingBasisPointsCellR1C1() { return this.GetConfigValue("FloatingBasisPoints", "CellAddress"); }
    get FloatingBasisPointsCellA1() { return this.GetCellAddressA1("FloatingBasisPoints"); }

    get RateOptionCellR1C1() { return this.GetConfigValue("RateOption", "CellAddress"); }
    get RateOptionCellA1() { return this.GetCellAddressA1("RateOption"); }

    get LPRRateDescriptionCellR1C1() { return this.GetConfigValue("LPRRateDescription", "CellAddress"); }
    get LPRRateDescriptionCellA1() { return this.GetCellAddressA1("LPRRateDescription"); }

    get ActualInterestRateCellR1C1() { return this.GetConfigValue("ActualInterestRate", "CellAddress"); }
    get ActualInterestRateCellA1() { return this.GetCellAddressA1("ActualInterestRate"); }

    // 方案要素属性
    get LesseeCellR1C1() { return this.GetConfigValue("Lessee", "CellAddress"); }
    get LesseeCellA1() { return this.GetCellAddressA1("Lessee"); }

    get GuarantorCellR1C1() { return this.GetConfigValue("Guarantor", "CellAddress"); }
    get GuarantorCellA1() { return this.GetCellAddressA1("Guarantor"); }

    get GuaranteeMethodCellR1C1() { return this.GetConfigValue("GuaranteeMethod", "CellAddress"); }
    get GuaranteeMethodCellA1() { return this.GetCellAddressA1("GuaranteeMethod"); }

    get LeaseMethodCellR1C1() { return this.GetConfigValue("LeaseMethod", "CellAddress"); }
    get LeaseMethodCellA1() { return this.GetCellAddressA1("LeaseMethod"); }

    // 经纪人参数属性
    get BrokerPaymentMethodCellR1C1() { return this.GetConfigValue("BrokerPaymentMethod", "CellAddress"); }
    get BrokerPaymentMethodCellA1() { return this.GetCellAddressA1("BrokerPaymentMethod"); }
    get BrokerPaymentMethodCellValue() { return this.ReadParameterValue("BrokerPaymentMethod"); }

    get BrokerFeeRateCellR1C1() { return this.GetConfigValue("BrokerFeeRate", "CellAddress"); }
    get BrokerFeeRateCellA1() { return this.GetCellAddressA1("BrokerFeeRate"); }
    get BrokerFeeRateCellValue() { return this.ReadParameterValue("BrokerFeeRate"); }
    
    get BrokerTotalFeeCellR1C1() { return this.GetConfigValue("BrokerTotalFee", "CellAddress"); }
    get BrokerTotalFeeCellA1() { return this.GetCellAddressA1("BrokerTotalFee"); }
    get BrokerTotalFeeCellValue() { return this.ReadParameterValue("BrokerTotalFee"); }
    
    // ==============类变量访问器 ==============
    get RowStartValue() { return this.m_RowStart; }
    set RowStartValue(value) { this.m_RowStart = value; }


    // ============== 高级功能方法 ==============

    /**
     * 获取所有参数名称列表
     * @returns {Array} 参数名称数组
     */
    GetParameterNames() {
        if (!this._isInitialized) {
            return [];
        }
        return Object.keys(this.m_config);
    }

    /**
     * 获取指定参数的完整配置信息
     * @param {string} parameterName - 参数名称
     * @returns {Object} 参数配置对象
     */
    GetParameterConfig(parameterName) {
        if (!this._isInitialized || !this.m_config[parameterName]) {
            return null;
        }
        return this.m_config[parameterName];
    }

    /**
     * 获取所有参数的配置摘要
     * @returns {string} 配置摘要字符串
     */
    GetAllParametersSummary() {
        if (!this._isInitialized) {
            return "参数管理器未初始化";
        }

        let summary = "=== 所有参数配置摘要 ===\n";
        const paramNames = this.GetParameterNames();
        summary += "参数总数: " + paramNames.length + "\n\n";

        for (const paramName of paramNames) {
            const config = this.m_config[paramName];
            summary += "【" + config.DisplayName + "】\n";
            summary += "  参数名称: " + paramName + "\n";
            summary += "  单元格地址: " + config.CellAddress + "\n";
            summary += "  默认值: " + config.DefaultValue + "\n";
            summary += "  数据类型: " + config.DataType + "\n";
            summary += "  是否必需: " + (config.IsRequired ? "是" : "否") + "\n";
            summary += "  验证规则: " + config.ValidationRule + "\n";
            summary += "  描述: " + config.Description + "\n";

            // 检查是否有默认公式
            if (config.DefaultFormula) {
                summary += "  默认公式: " + config.DefaultFormula + "\n";
            }

            summary += "\n";
        }

        return summary;
    }

    /**
     * 设置参数值到工作表
     * @param {string} parameterName - 参数名称
     * @param {any} value - 要设置的值
     * @returns {boolean} 是否设置成功
     */
    SetParameterValue(parameterName, value) {
        try {
            if (!this._isInitialized || !this.m_config[parameterName]) {
                return false;
            }

            const cellAddressA1 = this.GetCellAddressA1(parameterName);
            if (!cellAddressA1) {
                return false;
            }

            this.m_worksheet.Range(cellAddressA1).Value2 = value;
            return true;
            
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 设置参数值失败: " + error.message);
            return false;
        }
    }

    /**
     * 批量设置参数值
     * @param {Object} parameters - 参数对象 {参数名: 值}
     * @returns {Object} 设置结果 {成功数: number, 失败数: number}
     */
    SetParametersBatch(parameters) {
        let successCount = 0;
        let failureCount = 0;

        for (const [paramName, value] of Object.entries(parameters)) {
            if (this.SetParameterValue(paramName, value)) {
                successCount++;
            } else {
                failureCount++;
            }
        }

        return {
            successCount: successCount,
            failureCount: failureCount
        };
    }

    /**
     * 验证所有必需参数是否已填写
     * @returns {Object} 验证结果 {isValid: boolean, missingParams: Array}
     */
    ValidateRequiredParameters() {
        const missingParams = [];
        
        for (const paramName of this.GetParameterNames()) {
            const config = this.m_config[paramName];
            if (config.IsRequired) {
                const value = this.ReadParameterValue(paramName);
                if (value === null || value === undefined || value === "") {
                    missingParams.push(config.DisplayName);
                }
            }
        }

        return {
            isValid: missingParams.length === 0,
            missingParams: missingParams
        };
    }

    // ============== 
    // ============== 工具方法 ==============

    /**
     * 导出参数配置到JSON文件
     * @param {string} filePath - 文件路径
     * @returns {boolean} 是否导出成功
     */
    ExportConfigToJson(filePath) {
        try {
            // 使用WPS JSA的文件操作API
            const configData = {
                exportTime: new Date().toISOString(),
                sheetName: this.m_worksheet.Name,
                parameters: this.m_config
            };
            
            // 这里需要根据WPS JSA的实际API来写入文件
            console.log("JSON配置内容:", JSON.stringify(configData, null, 2));
            console.log("文件路径:", filePath);
            // 实际文件写入需要根据WPS JSA API实现
            
            return true;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 导出配置失败: " + error.message);
            return false;
        }
    }

    /**
     * 从JSON文件导入参数配置
     * @param {string} filePath - 文件路径
     * @returns {boolean} 是否导入成功
     * @deprecated WPS JSA不支持ActiveXObject，此方法暂不可用
     */
    ImportConfigFromJson(filePath) {
        try {
            // WPS JSA不支持ActiveXObject，此功能暂不可用
            console.log("[" + this.MODULE_NAME + "] 警告: WPS JSA不支持ActiveXObject，ImportConfigFromJson功能暂不可用");
            console.log("[" + this.MODULE_NAME + "] 提示: 如需导入配置，请手动编辑参数单元格或使用ExportConfigToJson查看当前配置");
            return false;
            
            /* 
            // 以下代码仅适用于VBA环境，在WPS JSA中无法使用
            const fso = new ActiveXObject("Scripting.FileSystemObject");
            if (!fso.FileExists(filePath)) {
                console.log("[" + this.MODULE_NAME + "] 配置文件不存在: " + filePath);
                return false;
            }

            const file = fso.OpenTextFile(filePath, 1);
            const jsonContent = file.ReadAll();
            file.Close();

            const configData = JSON.parse(jsonContent);
            this.m_config = configData.parameters;
            
            console.log("[" + this.MODULE_NAME + "] 配置已从文件导入: " + filePath);
            return true;
            */
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 导入配置失败: " + error.message);
            return false;
        }
    }

    /**
     * 重置参数值为默认值
     * @param {string} parameterName - 参数名称（可选，不传则重置所有参数）
     * @returns {Object} 重置结果 {successCount: number, failureCount: number}
     */
    ResetToDefault(parameterName) {
        let successCount = 0;
        let failureCount = 0;

        try {
            if (parameterName) {
                // 重置单个参数
                const config = this.m_config[parameterName];
                if (config) {
                    if (this.SetParameterValue(parameterName, config.DefaultValue)) {
                        successCount++;
                    } else {
                        failureCount++;
                    }
                }
            } else {
                // 重置所有参数
                for (const paramName of this.GetParameterNames()) {
                    const config = this.m_config[paramName];
                    if (this.SetParameterValue(paramName, config.DefaultValue)) {
                        successCount++;
                    } else {
                        failureCount++;
                    }
                }
            }
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 重置参数失败: " + error.message);
        }

        return { successCount, failureCount };
    }

    /**
     * 应用默认公式到计算参数
     * @returns {Object} 应用结果 {successCount: number, failureCount: number}
     */
    ApplyDefaultFormulas() {
        let successCount = 0;
        let failureCount = 0;

        try {
            for (const paramName of this.GetParameterNames()) {
                const config = this.m_config[paramName];
                if (config.DefaultFormula) {
                    const cellAddressA1 = this.GetCellAddressA1(paramName);
                    if (cellAddressA1) {
                        this.m_worksheet.Range(cellAddressA1).Formula = config.DefaultFormula;
                        successCount++;
                    } else {
                        failureCount++;
                    }
                }
            }
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 应用默认公式失败: " + error.message);
        }

        return { successCount, failureCount };
    }

    // ============== 数据验证方法 ==============

    /**
     * 验证参数值的有效性
     * @param {string} parameterName - 参数名称
     * @returns {Object} 验证结果 {isValid: boolean, message: string, value: any}
     */
    ValidateParameter(parameterName) {
        try {
            if (!this.m_config[parameterName]) {
                return {
                    isValid: false,
                    message: "参数不存在: " + parameterName,
                    value: null
                };
            }

            const config = this.m_config[parameterName];
            const value = this.ReadParameterValue(parameterName);
            const validationRule = config.ValidationRule;

            // 检查必需性
            if (config.IsRequired && (value === null || value === undefined || value === "")) {
                return {
                    isValid: false,
                    message: "必需参数未填写: " + config.DisplayName,
                    value: value
                };
            }

            // 检查验证规则
            if (!this.ValidateValueByRule(value, validationRule)) {
                return {
                    isValid: false,
                    message: "参数值不符合验证规则: " + config.DisplayName,
                    value: value
                };
            }

            return {
                isValid: true,
                message: "参数验证通过",
                value: value
            };

        } catch (error) {
            return {
                isValid: false,
                message: "验证过程中发生错误: " + error.message,
                value: null
            };
        }
    }

    /**
     * 批量验证所有参数
     * @returns {Object} 验证结果汇总
     */
    ValidateAllParameters() {
        const results = {
            total: 0,
            valid: 0,
            invalid: 0,
            details: []
        };

        for (const paramName of this.GetParameterNames()) {
            const validation = this.ValidateParameter(paramName);
            results.total++;
            
            if (validation.isValid) {
                results.valid++;
            } else {
                results.invalid++;
            }

            results.details.push({
                parameter: paramName,
                displayName: this.m_config[paramName].DisplayName,
                ...validation
            });
        }

        return results;
    }

    // ============== 调试和日志方法 ==============

    /**
     * 打印参数调试信息
     */
    PrintDebugInfo() {
        console.log("=== ParameterManager 调试信息 ===");
        console.log("初始化状态: " + this._isInitialized);
        console.log("工作表: " + (this.m_worksheet ? this.m_worksheet.Name : "未设置"));
        console.log("参数总数: " + this.GetParameterNames().length);
        
        const validation = this.ValidateAllParameters();
        console.log("参数验证: " + validation.valid + "/" + validation.total + " 通过");
        
        if (validation.invalid > 0) {
            console.log("未通过验证的参数:");
            for (const detail of validation.details) {
                if (!detail.isValid) {
                    console.log("  - " + detail.displayName + ": " + detail.message);
                }
            }
        }
    }

    /**
     * 生成参数报告
     * @returns {string} 参数报告文本
     */
    GenerateParameterReport() {
        let report = "=== 参数管理器报告 ===\n";
        report += "生成时间: " + new Date().toLocaleString() + "\n";
        report += "工作表: " + (this.m_worksheet ? this.m_worksheet.Name : "未设置") + "\n\n";

        const paramNames = this.GetParameterNames();
        report += "参数概览 (" + paramNames.length + "个参数):\n";

        for (const paramName of paramNames) {
            const config = this.m_config[paramName];
            const value = this.ReadParameterValue(paramName);
            const validation = this.ValidateParameter(paramName);

            report += "【" + config.DisplayName + "】\n";
            report += "  值: " + value + "\n";
            report += "  状态: " + (validation.isValid ? "✓ 有效" : "✗ 无效") + "\n";
            if (!validation.isValid) {
                report += "  错误: " + validation.message + "\n";
            }
            report += "  位置: " + this.GetCellAddressA1(paramName) + "\n";
            report += "\n";
        }

        return report;
    }

    // ============== 事件处理方法 ==============

    /**
     * 设置参数变化监听
     * @param {Function} callback - 回调函数
     */
    SetParameterChangeListener(callback) {
        // 在WPS JSA中，可以通过Worksheet的Change事件来监听参数变化
        if (this.m_worksheet) {
            this.m_worksheet.Change = function(changeRange) {
                const changedAddress = changeRange.Address;
                
                // 检查变化的单元格是否对应某个参数
                for (const paramName of this.GetParameterNames()) {
                    const paramAddress = this.GetCellAddressA1(paramName);
                    if (paramAddress === changedAddress) {
                        const newValue = this.ReadParameterValue(paramName);
                        callback(paramName, newValue, changedAddress);
                        break;
                    }
                }
            }.bind(this);
        }
    }

    // ============== 静态工具方法 ==============

    /**
     * 创建参数管理器实例（工厂方法）
     * @param {string} sheetName - 工作表名称
     * @param {Object} config - 配置对象
     * @returns {ParameterManager} 参数管理器实例
     */
    static Create(sheetName, config) {
        const manager = new ParameterManager();
        manager.Initialize(sheetName, config);
        return manager;
    }

    /**
     * 从现有工作表自动检测并创建参数管理器
     * @param {string} sheetName - 工作表名称
     * @returns {ParameterManager} 参数管理器实例
     */
    static AutoCreateFromSheet(sheetName) {
        const manager = new ParameterManager();
        manager.Initialize(sheetName);
        
        // 这里可以添加自动检测逻辑
        // 比如扫描工作表中的特定标记来识别参数
        
        return manager;
    }
}
