/**
 * ============== 初始化模块：租金测算表参数区域初始化（配置驱动架构） ==============
 * 作者：徐晓冬
 * 描述：完全重构版初始化模块 - 配置驱动架构
 * 
 * 核心改进：
 * - 配置驱动：所有初始化配置集中在类静态属性 CONFIG 中
 * - 零重复代码：提取所有公共逻辑和辅助方法
 * - 适配新版参数管理器：使用 val(), addr() 等快捷方法
 * - 职责单一：每个方法只做一件事
 * - 注释完善：所有关键逻辑都有设计意图说明
 * 
 * P0 修复：
 * - P0-1: CONFIG 从全局常量改为类静态属性，避免全局污染
 * - P0-2: 数据有效性配置从参数管理器 DataSource 读取，消除双重来源
 * - P0-3: 区域 backgroundColor 配置实际生效，不再误导
 * 
 * 依赖：
 * - mParameterManager.js (参数管理器 V3)
 * - mRentalCalculation.js (租金测算)
 * ====================================================
 */

// ============== 初始化主类 ==============
/**
 * clsRentCalculationFillinArea - 租金测算表填写区域初始化类
 * 
 * 作用：整合所有初始化功能，提供统一的接口
 * 功能：初始化各个参数区域、设置单元格样式、应用数据有效性
 * 设计：采用配置驱动模式，所有配置集中在类静态属性 CONFIG 中
 */
class clsRentCalculationFillinArea {
    /**
     * 类静态配置对象
     * 
     * P0-1 修复：从全局常量 INITIALIZATION_CONFIG 改为类静态属性，
     * 避免全局命名空间污染，配置与类内聚。
     * 
     * 配置结构：
     * - sections: 各个初始化区域的配置
     * - styles: 样式配置
     */
    static get CONFIG() {
        // P2-4: 首次访问时创建并缓存，后续直接返回
        if (clsRentCalculationFillinArea._cachedConfig) {
            return clsRentCalculationFillinArea._cachedConfig;
        }
        clsRentCalculationFillinArea._cachedConfig = {
            // 初始化区域配置
            sections: {
                // 价格参数区域配置
                priceParameters: {
                    title: "价格参数填写区域",
                    startCell: "A3",
                    parameters: [
                        "Principal",           // 租赁成本
                        "ActualPayment",       // 实际付款（公式）
                        "InterestRate",        // 利率
                        "RateOption",          // 利率选项（数据有效性）
                        "Deposit",             // 保证金（公式）
                        "DepositMarginRate",   // 押金保证金比例
                        "NominalPrice",        // 名义货价
                        "TotalPeriods",        // 总期数
                        "PaymentInterval",     // 支付间隔（数据有效性）
                        "PaymentsPerYear",     // 每年还款次数（公式）
                        "ProjectDurationYears", // 项目时长/年（公式）
                        "RepaymentMethod",     // 还款方式（数据有效性）
                        "LeaseStartDate",      // 放款日
                        "FirstPaymentDate",     // 起租日（公式）
                        "PreLeaseInterval",    // 租前期间隔（数据有效性）
                        "PreLeaseMonths"       // 租前期月数（公式）
                    ]
                },
                
                // 利率要素区域配置
                interestRateElements: {
                    title: "利率要素填写区域",
                    startCell: "F9",
                    parameters: [
                        "LPRBenchmarkRate",    // LPR基准利率（公式）
                        "FloatingBasisPoints", // 浮动基点（公式）
                        "LPRPeriod",           // LPR期限选择（数据有效性）
                        "LPRDate",             // LPR发布日期（数据有效性）
                        "LPRRateDescription"   // LPR利率描述（公式）
                    ]
                },
                
                // 租赁项目基本信息区域配置
                projectBasicInfo: {
                    title: "租赁项目基本信息区域",
                    startCell: "F3",
                    parameters: [
                        "Lessee",             // 承租人
                        "Guarantor",          // 担保人
                        "GuaranteeMethod",    // 担保方式
                        "LeaseMethod"         // 租赁方式
                    ]
                },
                
                // 经纪人费用参数区域配置
                brokerFeeParameters: {
                    title: "经纪人费用参数区域",
                    startCell: "F13",
                    parameters: [
                        "BrokerPaymentMethod", // 经纪人费用支付方式（数据有效性）
                        "BrokerFeeRate",       // 经纪人费用比例
                        "BrokerTotalFee"       // 经纪人总费用（公式）
                    ]
                }
            },
            
            // 样式配置
            styles: {
                // 主标题样式
                mainTitle: {
                    fontSize: 26,
                    bold: true,
                    fontName: "黑体",
                    horizontalAlignment: -4131, // xlLeft
                    verticalAlignment: -4108    // xlCenter
                },
                
                // 区域标题样式
                sectionTitle: {
                    fontSize: 14,
                    bold: false,
                    fontName: "黑体",
                    horizontalAlignment: -4131, // xlLeft
                    verticalAlignment: -4108    // xlCenter
                },
                
                // 参数单元格样式
                parameterCell: {
                    fontSize: 12,
                    bold: false,
                    fontName: "黑体",
                    horizontalAlignment: -4152, // xlRight
                    verticalAlignment: -4108    // xlCenter
                },
                
                // 参数名称单元格样式
                parameterNameCell: {
                    fontSize: 12,
                    bold: false,
                    fontName: "黑体",
                    horizontalAlignment: -4131, // xlLeft
                    verticalAlignment: -4108    // xlCenter
                }
            }
        };
        return clsRentCalculationFillinArea._cachedConfig;
    }
    
    /**
     * 构造函数
     * 
     * 作用：初始化类实例，设置基本属性
     */
    constructor() {
        this.MODULE_NAME = "clsRentCalculationFillinArea";
        this.ModuleModifyDate = (typeof VERSION_DATE !== 'undefined') ? VERSION_DATE : "20260411";
        this.ws = null;
        this.p = null; // 参数管理器实例
        
        console.log(`[${this.MODULE_NAME}] 类实例创建`);
    }
    
    // ==================== P0-2: 配置辅助方法 ====================
    
    /**
     * 是否为公式参数 - 判断参数是否需要使用公式
     * 
     * P0-2 修复：从参数管理器配置中的 DefaultFormula / IsCalculated 字段判断，
     * 而非维护独立的 formulaParameters 列表，消除双重配置来源。
     * 
     * @param {string} paramName - 参数名称
     * @returns {boolean} 是否为公式参数
     */
    是否为公式参数(paramName) {
        try {
            const paramConfig = this.p.param(paramName);
            const config = paramConfig.config;
            // P1-2 修复：仅当有 DefaultFormula 时才视为公式参数
            // IsCalculated 只是语义标记（表示该值由其他参数计算得到），
            // 但如果没有 DefaultFormula（如 RateOption），则仍需用户输入值
            return !!(config && config.DefaultFormula);
        } catch (error) {
            return false;
        }
    }
    
    /**
     * 获取参数的数据有效性配置
     * 
     * P0-2 修复：从参数管理器配置中的 DataSource 字段解析，
     * 格式为 "工作表名!$范围$"，自动拆分为 sourceSheet 和 sourceRange。
     * 消除了之前在 INITIALIZATION_CONFIG.dataValidation 中重复定义的问题。
     * 
     * @param {string} paramName - 参数名称
     * @returns {Object|null} 数据有效性配置 { sourceSheet, sourceRange } 或 null
     */
    获取数据有效性配置(paramName) {
        try {
            const paramConfig = this.p.param(paramName);
            const config = paramConfig.config;
            const dataSource = config.DataSource;
            
            if (!dataSource) {
                return null;
            }
            
            // P1-3 修复：如果参数有 DefaultFormula，则 DataSource 是公式数据源（如VLOOKUP），
            // 不应用于数据有效性下拉列表
            // 例如 LPRBenchmarkRate 的 DataSource 是 VLOOKUP 的多列范围，不是下拉列表
            if (config.DefaultFormula) {
                return null;
            }
            
            // 解析 "工作表名!范围" 格式
            // 例如 "还款设置!$K$2:$K$3" → sourceSheet="还款设置", sourceRange="$K$2:$K$3"
            const exclamationIdx = dataSource.indexOf("!");
            if (exclamationIdx > 0) {
                return {
                    sourceSheet: dataSource.substring(0, exclamationIdx),
                    sourceRange: dataSource.substring(exclamationIdx + 1)
                };
            }
            
            return null;
        } catch (error) {
            return null;
        }
    }
    
    /**
     * 获取参数的背景颜色
     * 
     * P0-3 修复：根据参数类型自动确定背景颜色
     * - 公式参数 → 浅绿色（LIGHT_GREEN）
     * - 可编辑参数 → 黄色（YELLOW）
     * - 参数名称单元格 → 白色（WHITE）
     * 
     * @param {string} paramName - 参数名称
     * @returns {number} 颜色值
     */
    获取参数背景色(paramName) {
        if (this.是否为公式参数(paramName)) {
            return this.p.m_COLOR_LIGHT_GREEN;
        }
        return this.p.m_COLOR_YELLOW;
    }
    
    // ==================== 初始化方法 ====================
    
    /**
     * Initialize - 初始化方法
     * 
     * 作用：初始化初始化系统，设置参数管理器和工作表引用
     * 
     * @param {Object} parameterManager - 参数管理器实例
     * @param {string} sheetName - 工作表名称（可选）
     * @returns {boolean} 是否初始化成功
     */
    Initialize(parameterManager, sheetName) {
        try {
            this.p = parameterManager;
            
            // 如果提供了工作表名称，则初始化参数管理器
            if (sheetName) {
                this.p.Initialize(sheetName);
            }
            
            this.ws = this.p.m_worksheet;
            
            console.log(`[${this.MODULE_NAME}] 初始化完成 - 工作表: ${this.ws.Name}`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 初始化失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * main - 主初始化函数
     * 
     * 作用：执行所有区域的初始化
     * 设计：配置驱动的主流程，根据配置对象自动处理
     * 
     * @returns {boolean} 是否初始化成功
     */
    main() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始初始化填写区域...`);
            
            // WPS JSA 优化：关闭屏幕刷新，避免150+次工作表操作触发界面重绘
            Application.ScreenUpdating = false;
            
            // 生成主标题
            this.生成主标题("租赁项目预算租金偿还表", "A1");
            
            // 初始化各个区域
            const sectionKeys = Object.keys(clsRentCalculationFillinArea.CONFIG.sections);
            for (const sectionKey of sectionKeys) {
                this.初始化区域(sectionKey);
            }
            
            // 恢复屏幕刷新
            Application.ScreenUpdating = true;
            
            console.log(`[${this.MODULE_NAME}] 填写区域初始化完成`);
            return true;
        } catch (error) {
            // 确保异常时也恢复屏幕刷新
            try { Application.ScreenUpdating = true; } catch (e) {}
            console.log(`[${this.MODULE_NAME}] 填写区域初始化失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 初始化区域 - 通用区域初始化方法
     * 
     * 作用：根据区域配置初始化指定区域
     * 设计：配置驱动，支持灵活的区域管理
     * 
     * @param {string} sectionKey - 区域配置键
     * @returns {boolean} 是否初始化成功
     */
    初始化区域(sectionKey) {
        try {
            const sectionConfig = clsRentCalculationFillinArea.CONFIG.sections[sectionKey];
            
            if (!sectionConfig) {
                throw new Error(`区域配置不存在：${sectionKey}`);
            }
            
            console.log(`[${this.MODULE_NAME}] 开始初始化区域：${sectionConfig.title}`);
            
            // 生成区域标题
            this.生成区域标题(sectionConfig.title, sectionConfig.startCell);
            
            // 初始化区域内的所有参数
            for (const paramName of sectionConfig.parameters) {
                if (this.是否为公式参数(paramName)) {
                    this.设置单元格公式(paramName);
                } else {
                    this.设置单元格值(paramName);
                }
                
                // 后处理：LPRRateDescription 没有默认公式（由 设置LPR利率描述 动态构建），
                // 但需要特殊公式，必须在值设置后单独处理
                if (paramName === "LPRRateDescription") {
                    this.设置LPR利率描述();
                }
            }
            
            console.log(`[${this.MODULE_NAME}] 区域初始化完成：${sectionConfig.title}`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 区域初始化失败：${sectionKey} - ${error.message}`);
            return false;
        }
    }
    
    /**
     * 生成主标题 - 生成主标题
     * 
     * 作用：在工作表中生成主标题
     * 设计：使用配置对象中的样式配置
     * 
     * @param {string} title - 标题文本
     * @param {string} cellAddressA1 - 单元格地址
     * @returns {boolean} 是否生成成功
     */
    生成主标题(title, cellAddressA1) {
        try {
            const titleRange = this.ws.Range(cellAddressA1);
            titleRange.Value2 = title;
            this.应用样式(titleRange, clsRentCalculationFillinArea.CONFIG.styles.mainTitle);
            // WPS JSA 修复：使用 EntireRow 而非 Rows，确保单单元格场景下行高自适应
            titleRange.EntireRow.AutoFit();
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 生成主标题失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 生成区域标题 - 生成区域标题
     * 
     * 作用：在工作表中生成区域标题
     * 设计：使用配置对象中的样式配置
     * 
     * @param {string} title - 标题文本
     * @param {string} cellAddressA1 - 单元格地址
     * @returns {boolean} 是否生成成功
     */
    生成区域标题(title, cellAddressA1) {
        try {
            const titleRange = this.ws.Range(cellAddressA1);
            titleRange.Value2 = title;
            this.应用样式(titleRange, clsRentCalculationFillinArea.CONFIG.styles.sectionTitle);
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 生成区域标题失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 设置单元格值 - 设置参数单元格值
     * 
     * 作用：使用参数管理器设置单个参数单元格的值
     * P0-3 修复：背景色通过 获取参数背景色() 方法自动确定
     * 
     * @param {string} paramName - 参数名称
     * @returns {boolean} 是否设置成功
     */
    设置单元格值(paramName) {
        try {
            // 使用参数管理器获取配置信息
            const cellAddressA1 = this.p.addr(paramName);
            const paramConfig = this.p.param(paramName);
            const displayName = paramConfig.displayName;
            const defaultValue = paramConfig.defaultValue;
            const vbaFormat = paramConfig.config.VbaFormat;
            
            // 设置参数值单元格
            const valueRange = this.ws.Range(cellAddressA1);
            valueRange.Value2 = defaultValue;
            valueRange.NumberFormat = vbaFormat;
            this.应用样式(valueRange, clsRentCalculationFillinArea.CONFIG.styles.parameterCell);
            
            // P0-3: 通过方法获取背景颜色，而非硬编码
            valueRange.Interior.Color = this.获取参数背景色(paramName);
            
            // 设置参数名称单元格（左侧单元格）
            const nameRange = valueRange.Offset(0, -1);
            nameRange.Value2 = displayName;
            this.应用样式(nameRange, clsRentCalculationFillinArea.CONFIG.styles.parameterNameCell);
            nameRange.Interior.Color = this.p.m_COLOR_WHITE;
            
            // P0-2: 从参数管理器 DataSource 读取数据有效性配置
            const validationConfig = this.获取数据有效性配置(paramName);
            if (validationConfig) {
                this.设置参数数据有效性(paramName, validationConfig);
            }
            
            console.log(`[${this.MODULE_NAME}] 参数'${paramName}'设置完成`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 设置单元格值失败：${paramName} - ${error.message}`);
            return false;
        }
    }
    
    /**
     * 设置单元格公式 - 设置参数单元格公式
     * 
     * 作用：使用参数管理器设置单个参数单元格的公式
     * P0-3 修复：背景色通过 获取参数背景色() 方法自动确定
     * 
     * @param {string} paramName - 参数名称
     * @returns {boolean} 是否设置成功
     */
    设置单元格公式(paramName) {
        try {
            // 使用参数管理器获取配置信息
            const cellAddressA1 = this.p.addr(paramName);
            const paramConfig = this.p.param(paramName);
            const displayName = paramConfig.displayName;
            const defaultFormula = paramConfig.config.DefaultFormula;
            const vbaFormat = paramConfig.config.VbaFormat;
            
            // 设置参数公式单元格
            const formulaRange = this.ws.Range(cellAddressA1);
            formulaRange.Formula = defaultFormula;
            formulaRange.NumberFormat = vbaFormat;
            this.应用样式(formulaRange, clsRentCalculationFillinArea.CONFIG.styles.parameterCell);
            
            // P0-3: 通过方法获取背景颜色，而非硬编码
            formulaRange.Interior.Color = this.获取参数背景色(paramName);
            
            // 设置参数名称单元格（左侧单元格）
            const nameRange = formulaRange.Offset(0, -1);
            nameRange.Value2 = displayName;
            this.应用样式(nameRange, clsRentCalculationFillinArea.CONFIG.styles.parameterNameCell);
            nameRange.Interior.Color = this.p.m_COLOR_WHITE;
            
            // P0-2: 从参数管理器 DataSource 读取数据有效性配置
            const validationConfig = this.获取数据有效性配置(paramName);
            if (validationConfig) {
                this.设置参数数据有效性(paramName, validationConfig);
            }
            
            console.log(`[${this.MODULE_NAME}] 参数'${paramName}'公式设置完成`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 设置单元格公式失败：${paramName} - ${error.message}`);
            return false;
        }
    }
    
    /**
     * 设置LPR利率描述 - 设置LPR利率描述的特殊公式
     * 
     * 作用：为LPR利率描述单元格设置动态公式
     * 设计：构建详细的利率描述文本公式
     */
    设置LPR利率描述() {
        try {
            // 获取单元格地址
            const lCell = this.p.addr("LPRDate"); // LPR发布日期单元格地址
            const pCell = this.p.addr("LPRPeriod");
            const rCell = this.p.addr("LPRBenchmarkRate");
            const bpCell = this.p.addr("FloatingBasisPoints");
            
            const descCellAddress = this.p.addr("LPRRateDescription");
            const desc = `= "即 " & TEXT(${lCell},"yyyy年mm月dd日") & "全国银行间同业拆借中心公布的 " & ${pCell} & " 年期人民币贷款基础利率（LPR）" & ${rCell}*100 & "%，加" & ROUND(${bpCell},2) & "BPS（1BP=0.01%）"`;
            
            const descRange = this.ws.Range(descCellAddress);
            descRange.Formula = desc;
            descRange.NumberFormat = "@";
            descRange.Interior.Color = this.p.m_COLOR_LIGHT_GREEN;
            descRange.Font.Color = this.p.m_COLOR_BLACK;
            descRange.HorizontalAlignment = -4131; // xlLeft
            descRange.VerticalAlignment = -4108;   // xlCenter
            descRange.WrapText = false; // 不自动换行
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 设置LPR利率描述失败：${error.message}`);
        }
    }
    
    /**
     * 设置参数数据有效性 - 设置参数的数据有效性
     * 
     * P0-2 修复：参数改为直接接收 validationConfig 对象，
     * 不再从独立的 INITIALIZATION_CONFIG.dataValidation 读取。
     * 配置来源统一为参数管理器的 DataSource 字段。
     * 
     * @param {string} paramName - 参数名称
     * @param {Object} validationConfig - 数据有效性配置 { sourceSheet, sourceRange }
     * @returns {boolean} 是否设置成功
     */
    设置参数数据有效性(paramName, validationConfig) {
        try {
            if (!validationConfig) {
                return true;
            }
            
            const targetCellAddress = this.p.addr(paramName);
            const targetRange = this.ws.Range(targetCellAddress);
            
            // 清除原有数据有效性
            try {
                targetRange.Validation.Delete();
            } catch (e) {
                // 忽略删除错误
            }
            
            // 获取源数据范围 - 与参数管理器使用相同的 Application.Worksheets() 模式
            const sourceSheet = Application.Worksheets(validationConfig.sourceSheet);
            const sourceRange = sourceSheet.Range(validationConfig.sourceRange);
            
            // 添加数据有效性
            // 跨工作表数据有效性必须包含工作表名：='还款设置'!$K$2:$K$3
            // Address(false) 返回本地地址（无工作表名），需手动拼接
            targetRange.Validation.Add({
                Type: 3,              // xlValidateList
                Operator: 1,          // xlBetween
                Formula1: "='" + validationConfig.sourceSheet + "'!" + sourceRange.Address(true, true, 1, false)
            });
            
            // WPS JSA 优化：直接使用参数管理器的 defaultValue，避免 val() 多余的读工作表操作
            const paramConfig = this.p.param(paramName);
            const defaultVal = paramConfig.defaultValue;
            if (defaultVal !== null && defaultVal !== undefined && defaultVal !== "") {
                targetRange.Value2 = defaultVal;
            } else {
                // 设置为数据有效性列表的第一个值
                try {
                    const firstValue = sourceRange.Cells(1, 1).Value2;
                    if (firstValue) {
                        targetRange.Value2 = firstValue;
                    }
                } catch (e) {
                    // 忽略错误
                }
            }
            
            // 设置对齐方式
            targetRange.HorizontalAlignment = -4131; // xlLeft
            targetRange.VerticalAlignment = -4108;   // xlCenter
            
            console.log(`[${this.MODULE_NAME}] 参数'${paramName}'数据有效性设置完成`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 设置参数数据有效性失败：${paramName} - ${error.message}`);
            return false;
        }
    }
    
    /**
     * 应用样式 - 应用样式到单元格范围
     * 
     * 作用：根据样式配置对象应用样式
     * 设计：统一的样式应用方法，减少重复代码
     * 
     * @param {Range} range - 单元格范围对象
     * @param {Object} styleConfig - 样式配置对象
     * @returns {boolean} 是否应用成功
     */
    应用样式(range, styleConfig) {
        try {
            if (!range || !styleConfig) {
                return false;
            }
            
            if (styleConfig.fontSize !== undefined) {
                range.Font.Size = styleConfig.fontSize;
            }
            
            if (styleConfig.bold !== undefined) {
                range.Font.Bold = styleConfig.bold;
            }
            
            if (styleConfig.fontName !== undefined) {
                range.Font.Name = styleConfig.fontName;
            }
            
            if (styleConfig.horizontalAlignment !== undefined) {
                range.HorizontalAlignment = styleConfig.horizontalAlignment;
            }
            
            if (styleConfig.verticalAlignment !== undefined) {
                range.VerticalAlignment = styleConfig.verticalAlignment;
            }
            
            if (styleConfig.color !== undefined) {
                range.Font.Color = styleConfig.color;
            }
            
            if (styleConfig.interiorColor !== undefined) {
                range.Interior.Color = styleConfig.interiorColor;
            }
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 应用样式失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 重置所有区域 - 重置所有初始化区域
     * 
     * 作用：清除所有区域的内容并重新初始化
     * 设计：提供快速重置功能
     * 
     * @returns {boolean} 是否重置成功
     */
    重置所有区域() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始重置所有区域...`);
            
            // 清除工作表内容
            this.ws.UsedRange.Clear();
            
            // 重新初始化
            const result = this.main();
            
            console.log(`[${this.MODULE_NAME}] 所有区域重置完成`);
            return result;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 重置所有区域失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 验证初始化 - 验证初始化是否正确
     * 
     * 作用：验证所有参数是否正确初始化
     * 设计：检查关键参数的值和格式
     * 
     * @returns {Object} 验证结果
     */
    验证初始化() {
        try {
            const results = {
                total: 0,
                valid: 0,
                invalid: 0,
                details: []
            };
            
            // 检查所有区域的参数
            const sections = clsRentCalculationFillinArea.CONFIG.sections;
            for (const sectionKey in sections) {
                const sectionConfig = sections[sectionKey];
                
                for (const paramName of sectionConfig.parameters) {
                    results.total++;
                    
                    const paramConfig = this.p.param(paramName);
                    const value = this.p.val(paramName);
                    const validation = this.p.ValidateParameter(paramName);
                    
                    if (validation.isValid) {
                        results.valid++;
                    } else {
                        results.invalid++;
                    }
                    
                    results.details.push({
                        section: sectionConfig.title,
                        parameter: paramName,
                        displayName: paramConfig.displayName,
                        value: value,
                        isValid: validation.isValid,
                        message: validation.message
                    });
                }
            }
            
            return results;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 验证初始化失败：${error.message}`);
            return { total: 0, valid: 0, invalid: 0, details: [] };
        }
    }
    
    /**
     * 获取初始化报告 - 生成初始化报告
     * 
     * 作用：生成详细的初始化报告
     * 设计：提供清晰的初始化状态信息
     * 
     * @returns {string} 初始化报告文本
     */
    获取初始化报告() {
        try {
            var report = "=== 租金测算表初始化报告 ===\n";
            report += `生成时间: ${new Date().toLocaleString()}\n`;
            report += `工作表: ${this.ws.Name}\n\n`;
            
            const validation = this.验证初始化();
            
            report += `参数概览 (${validation.total}个参数):\n`;
            report += `  有效: ${validation.valid}\n`;
            report += `  无效: ${validation.invalid}\n\n`;
            
            // 按区域分组显示
            const sectionMap = {};
            for (const detail of validation.details) {
                if (!sectionMap[detail.section]) {
                    sectionMap[detail.section] = [];
                }
                sectionMap[detail.section].push(detail);
            }
            
            for (const sectionName in sectionMap) {
                report += `【${sectionName}】\n`;
                for (const detail of sectionMap[sectionName]) {
                    report += `  ${detail.displayName}: ${detail.value}`;
                    report += detail.isValid ? " ✓\n" : " ✗\n";
                    if (!detail.isValid) {
                        report += `    错误: ${detail.message}\n`;
                    }
                }
                report += "\n";
            }
            
            return report;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 获取初始化报告失败：${error.message}`);
            return "获取初始化报告失败";
        }
    }
}

// ============== 快捷调用函数 ==============
/**
 * 测算表填写区域模块调用 - 快捷调用函数
 * 
 * 作用：提供简化的调用接口，无需传入任何参数
 * 设计：内部自动创建参数管理器并初始化
 */
function 测算表填写区域模块调用() {
    try {
        console.log("=== 开始测算表填写区域生成 ===");
        
        // 创建参数管理器实例并初始化
        var p = new clsParameterManager();
        p.Initialize("1租金测算表V1");
        console.log(`[测算表填写区域模块调用] 参数管理器已初始化`);
        
        // 创建初始化器实例
        const initializer = new clsRentCalculationFillinArea();
        initializer.Initialize(p, null); // 不需要重新初始化参数管理器
        
        // 执行初始化
        const result = initializer.main();
        
        // 输出初始化报告
        console.log(initializer.获取初始化报告());
        
        return result;
    } catch (error) {
        console.log(`测算表填写区域生成失败：${error.message}`);
        return false;
    }
}

console.log(`[mInitialization] 模块加载完成 - 版本 2.2026.4.11 (P0修复)`);
// ============== 结束 ==============