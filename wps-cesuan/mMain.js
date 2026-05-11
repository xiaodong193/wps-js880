/**
 * ============== 租金测算系统主文件 ==============
 * 作者：徐晓冬
 * 描述：整合所有租金测算功能的主模块 - 配置驱动架构
 * 
 * 核心改进：
 * - 配置驱动：所有流程配置集中在 WORKFLOW_CONFIG 对象中
 * - 零重复代码：提取所有公共逻辑
 * - 错误处理：统一的错误处理机制
 * - 日志完善：详细的执行日志
 * 
 * 依赖：
 * - mParameterManager.js (参数管理器 V2)
 * - mInitialization.js (初始化模块)
 * - mRentalCalculation.js (租金测算)
 * - mCashFlowGenerator.js (现金流量表)
 * ====================================================
 */

// ============== 工作流配置 ==============
const WORKFLOW_CONFIG = {
    // 工作流步骤配置
    steps: {
        clear: {
            name: "清除数据",
            description: "清除原有表中数据",
            priority: 1
        },
        generateRentTable: {
            name: "生成租金表",
            description: "创建租金测算表表头并计算每期利率",
            priority: 2
        },
        generateCashFlow: {
            name: "生成现金流量表",
            description: "生成现金流量表和综合利率一览",
            priority: 3
        },
        adjustPeriods: {
            name: "调期",
            description: "清除数据后重新生成表并生成月间隔",
            priority: 4
        }
    },
    
    // 系统配置
    // 注意：版本号统一由 mShared_constants.js 中的 VERSION 常量管理
    system: {
        moduleName: "租金测算系统",
        author: "徐晓冬",
        get version() { return typeof VERSION !== 'undefined' ? `V${VERSION}` : "V0.0.0.0"; }
    },
    
    // 错误处理配置
    errorHandling: {
        showAlert: true,
        logToConsole: true,
        continueOnError: false
    }
};

// ============== 主系统类 ==============
class clsRentCalculationSystem {
    constructor() {
        this.MODULE_NAME = WORKFLOW_CONFIG.system.moduleName;
        this.ModuleModifyDate = (typeof VERSION_DATE !== 'undefined') ? VERSION_DATE : "20260130";
        this.p = null; // 参数管理器实例
        this.initializer = null; // 初始化器实例
        this._rentalCalc = null; // 缓存的租金测算实例
        
        console.log(`[${this.MODULE_NAME}] 类实例创建`);
    }
    
    getRentalCalc() {
        if (!this._rentalCalc) {
            if (!this.p || !this.p.IsInitialized) {
                throw new Error(
                    "[clsRentCalculationSystem] 参数管理器未初始化，无法创建租金测算实例。" +
                    "请先调用 Initialize() 初始化系统。"
                );
            }
            // 依赖注入：始终传递已初始化的参数管理器
            this._rentalCalc = new clsRentalCalculation(this.p);
            this._rentalCalc.Initialize(this.p);
        }
        return this._rentalCalc;
    }
    
    Initialize(sheetName) {
        try {
            console.log(`[${this.MODULE_NAME}] 开始初始化系统...`);
            console.log(`版本：${WORKFLOW_CONFIG.system.version}`);
            console.log(`作者：${WORKFLOW_CONFIG.system.author}`);
            
            // 初始化参数管理器
            if (typeof clsParameterManager === 'undefined') {
                throw new Error("clsParameterManager未定义，请检查mParameterManager.js是否正确加载");
            }
            
            this.p = new clsParameterManager();
            const defaultSheetName = sheetName || "1租金测算表V1";
            this.p.Initialize(defaultSheetName);
            
            // 初始化初始化器
            if (typeof clsRentCalculationFillinArea !== 'undefined') {
                this.initializer = new clsRentCalculationFillinArea();
                this.initializer.Initialize(this.p, null);
            }
            
            console.log(`[${this.MODULE_NAME}] 系统初始化完成`);
            return true;
        } catch (error) {
            this.处理错误("系统初始化", error);
            return false;
        }
    }
    
    计算main() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始执行主计算流程...`);
            
            // 确保系统已初始化
            if (!this.p) {
                const initResult = this.Initialize();
                if (!initResult) {
                    return false;
                }
            }
            
            // 执行工作流步骤
            const steps = [
                this.清除,
                this.生成租金表,
                this.生成现金流量表
            ];
            
            for (const step of steps) {
                const stepResult = step.call(this);
                if (!stepResult) {
                    throw new Error(`工作流步骤执行失败`);
                }
            }
            
            console.log(`[${this.MODULE_NAME}] 主计算流程执行完成`);
            return true;
        } catch (error) {
            this.处理错误("主计算", error);
            return false;
        }
    }
    
    清除() {
        try {
            console.log('[' + this.MODULE_NAME + '] 开始清除数据...');

            // 确保系统已初始化
            if (!this.p) {
                var initResult = this.Initialize();
                if (!initResult) {
                    return false;
                }
            }

            // 自动备份
            var totalPeriods = this.p.TotalPeriodsCellValue || 36;
            var lastRow = 5 + totalPeriods - 1;
            backupSheetData('1租金测算表V1', 'A5:M' + lastRow);

            var r = this.getRentalCalc();
            r.清除原有表中数据();

            console.log('[' + this.MODULE_NAME + '] 数据清除完成');
            return true;
        } catch (error) {
            this.处理错误("清除数据", error);
            return false;
        }
    }
    
    调期() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始调期...`);
            
            // 执行清除、生成租金表、生成现金流量表
            const steps = [
                this.清除,
                this.生成租金表,
                this.生成现金流量表
            ];
            
            for (const step of steps) {
                const stepResult = step.call(this);
                if (!stepResult) {
                    throw new Error(`调期步骤执行失败`);
                }
            }
            
            // 生成月间隔
            const r = this.getRentalCalc();
            r.生成月间隔();
            
            console.log(`[${this.MODULE_NAME}] 调期完成`);
            return true;
        } catch (error) {
            this.处理错误("调期", error);
            return false;
        }
    }
    
    初始化系统() {
        try {
            console.log("=== 租金测算系统初始化 ===");
            console.log("系统版本：" + WORKFLOW_CONFIG.system.version);
            console.log("作者：" + WORKFLOW_CONFIG.system.author);
            console.log("系统名称：" + WORKFLOW_CONFIG.system.moduleName);
            console.log("系统初始化完成");
            return true;
        } catch (error) {
            console.log("系统初始化失败：" + error.message);
            return false;
        }
    }
    
    生成租金表() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始生成租金表...`);
            
            // 确保系统已初始化
            if (!this.p) {
                const initResult = this.Initialize();
                if (!initResult) {
                    return false;
                }
            }
            
            const r = this.getRentalCalc();
            // 注意：不再调用清除原有表中数据()，因为工作流中已经有单独的清除步骤
            // 如果单独调用此方法需要先清除，请在调用前先执行清除()
            r.创建租金测算表表头(1, 10);
            r.createDataRange();
            r.创建租金测算表表头(13, 13);
            r.每期适用利率();
            
            console.log(`[${this.MODULE_NAME}] 租金表生成完成`);
            return true;
        } catch (error) {
            this.处理错误("生成租金表", error);
            return false;
        }
    }
    
    生成现金流量表() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始生成现金流量表...`);
            
            // 确保参数管理器已初始化
            if (!this.p || !this.p.IsInitialized) {
                console.log(`[${this.MODULE_NAME}] 参数管理器未初始化，重新初始化...`);
                if (!this.Initialize()) {
                    throw new Error("参数管理器初始化失败");
                }
            }
            
            // 调试：检查关键参数值
            const totalPeriods = this.p.val("TotalPeriods");
            const principal = this.p.val("Principal");
            console.log(`[${this.MODULE_NAME}] 参数检查 - 总期数: ${totalPeriods}, 租赁成本: ${principal}`);
            
            if (!totalPeriods || totalPeriods <= 0) {
                throw new Error(`总期数无效: ${totalPeriods}，请检查B10单元格的值`);
            }
            
            var cashFlowGen;
            try {
                // 依赖注入：传递已有的参数管理器，避免重复创建
                cashFlowGen = new clsCashFlowGenerator(this.p);
                cashFlowGen.Initialize(this.p);
                
                if (cashFlowGen.GenerateCashFlowTable()) {
                    cashFlowGen.generateInterestRateOverview();
                    console.log(`[${cashFlowGen.MODULE_NAME}] 现金流量表生成成功`);
                    return true;
                } else {
                    console.log(`[${cashFlowGen.MODULE_NAME}] 现金流量表生成失败`);
                    return false;
                }
            } catch (error) {
                throw new Error(`现金流量表生成失败：${error.message}`);
            }
        } catch (error) {
            this.处理错误("生成现金流量表", error);
            return false;
        }
    }
    
    生成银承现金流量表() {
        try {
            console.log("开始生成银行承兑汇票现金流量表...");
            
            if (typeof cls银行承兑汇票 === 'undefined') {
                throw new Error("cls银行承兑汇票未定义，请检查m银行承兑汇票模块.js是否正确加载");
            }
            
            // 银承模块自行初始化参数管理器（针对"银行承兑汇票"工作表）
            // 不传 this.p，因为银承需要独立的参数管理器
            const bankModule = new cls银行承兑汇票();
            const result = bankModule.生成银承现金流量表();
            
            console.log("银行承兑汇票现金流量表生成结果: " + result);
            return result;
            
        } catch (error) {
            this.处理错误("生成银承现金流量表", error);
            return false;
        }
    }
    
    复制工作表(sourceSheetName, newSheetName = "", position = "before") {
        try {
            console.log(`[${this.MODULE_NAME}] 开始复制工作表: ${sourceSheetName}`);
            
            // 获取当前活动工作簿
            const workbook = Application.ActiveWorkbook;
            if (!workbook) {
                throw new Error("无法获取活动工作簿");
            }
            
            // 获取源工作表
            const sourceSheet = workbook.Worksheets.Item(sourceSheetName);
            if (!sourceSheet) {
                throw new Error(`找不到源工作表: ${sourceSheetName}`);
            }
            
            // 复制工作表
            // WPS JS API: Copy(Before) 或 Copy(null, After)
            // 注意：WPS中只需要传一个参数
            if (position === "before") {
                // 在源工作表之前复制
                sourceSheet.Copy(sourceSheet);
            } else {
                // 在源工作表之后复制
                sourceSheet.Copy(null, sourceSheet);
            }
            
            // 获取新创建的工作表（复制后成为活动工作表）
            const newSheet = Application.ActiveSheet;
            
            // 设置新工作表名称
            if (newSheetName) {
                newSheet.Name = newSheetName;
            }
            
            console.log(`[${this.MODULE_NAME}] 工作表复制成功: ${newSheet.Name}`);
            return newSheet;
        } catch (error) {
            this.处理错误("复制工作表", error);
            return null;
        }
    }
    
    处理错误(operation, error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        const errorPrefix = `[${this.MODULE_NAME}] ${operation} 失败: `;
        
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(
                error instanceof Error ? error : new Error(String(error)),
                ERROR_CODES.CALCULATION_ERROR,
                { module: this.MODULE_NAME, function: operation }
            );
        } else {
            console.error(`${errorPrefix}${errorMessage}`);
        }
    }
    
    获取系统信息() {
        return {
            moduleName: WORKFLOW_CONFIG.system.moduleName,
            version: WORKFLOW_CONFIG.system.version,
            author: WORKFLOW_CONFIG.system.author,
            isInitialized: !!this.p,
            hasInitializer: !!this.initializer,
            worksheetName: this.p ? this.p.m_worksheet.Name : null
        };
    }
}

// ============== 全局系统实例 ==============
/**
 * 系统全局实例
 * 用于在整个工作簿中共享系统实例
 */
var rentCalculationSystem = null;

// ============== 统一的实例获取方法 ==============
/**
 * getRentSystem - 获取全局系统实例（单例模式）
 * 
 * 作用：确保系统只创建一次，所有快捷函数共享同一个实例
 * 
 * @returns {clsRentCalculationSystem} 系统实例
 */
function getRentSystem() {
    if (!rentCalculationSystem) {
        rentCalculationSystem = new clsRentCalculationSystem();
    }
    return rentCalculationSystem;
}

// ============== 快捷调用函数 ==============
const sys = () => getRentSystem();

function 计算main() { try { return sys().计算main(); } catch (e) { console.log(`失败：${e.message}`); return false; } }
function 清除() { try { return sys().清除(); } catch (e) { console.log(`失败：${e.message}`); return false; } }
function 调期() { try { return sys().调期(); } catch (e) { console.log(`失败：${e.message}`); return false; } }
function 初始化系统() { try { return sys().初始化系统(); } catch (e) { console.log(`失败：${e.message}`); return false; } }
function copySht() { try { return sys().复制工作表("1租金测算表V1", "原合同", "before"); } catch (e) { console.log(`失败：${e.message}`); return null; } }

// 调整第1期和最后一期的自定义支付日/月间隔
function 调1期() {
    try {
        console.log(`[调1期] 开始调整第1期和最后一期的自定义支付日/月间隔...`);
        
        // P0-2修复：先调期（含清除+生成租金表+现金流量表+月间隔），再改支付日，最后重新生成日期
        // 调期（内部流程：清除→生成租金表→生成现金流量表→生成月间隔）
        getRentSystem().调期();
        
        // 改变自定义支付日-自定义的第1期数值为3
        // P1-3修复：复用系统缓存的租金测算实例，避免重复创建
        const r = getRentSystem().getRentalCalc();
        r.改变自定义支付日(1, 3);
        
        // 改变自定义支付日-自定义的最后一期数值为9
        r.改变自定义支付日(-1, 9);
        
        // P0-2修复：改变支付日后，必须重新生成日期公式，否则B列日期不会根据新的K列间隔更新
        const m_COL_DATE = r.p.m_COL_DATE;
        const rentTableStartRow = r.p.RentTableStartRow;
        const totalPeriodsValue = r.p.val("TotalPeriods");
        const leaseStartDateCellA1 = r.p.addr("LeaseStartDate");
        const m_worksheet = r.p.m_worksheet;
        
        const targetRange = m_worksheet.Range(
            `${m_COL_DATE}${rentTableStartRow}:${m_COL_DATE}${rentTableStartRow + totalPeriodsValue - 1}`
        );
        r.自定义月间隔(targetRange, leaseStartDateCellA1);
        
        console.log(`[调1期] 调整完成 - 第1期: 3, 最后一期: 9`);
        return true;
    } catch (error) {
        console.error(`调1期 失败：${error.message}`);
        return false;
    }
}
// 自动填入默认参数并执行完整测算
function 执行Main() {
    // Bug修复：asSheet/asRange 在 WPS JSA 中未定义，改用标准 WPS API
    const ws = Application.Worksheets("1租金测算表V1");
    if (!ws) {
        console.error("找不到工作表：1租金测算表V1");
        return false;
    }
    ws.Activate();
    ws.Range("B4").Value2 = 500000000;   // 租赁成本
    ws.Range("B5").Value2 = 0.035;       // 利率 3.5%
    ws.Range("B6").Value2 = 0;           // 保证金
    ws.Range("B8").Value2 = 1;           // 支付方式
    ws.Range("B10").Value2 = 10;         // 总期数
    ws.Range("D10").Value2 = 6;          // 间隔月数
    清除();
    计算main();
}

// 使用每期适用利率生成测算表
function 调2_每期适用利率() {
    try {
        console.log("========== 开始生成每期适用利率测算表 ==========");
        
        // 1. 清除原有数据
        getRentSystem().清除();
        
        // 2. 获取缓存的租金测算实例（P1-3修复：避免重复创建）
        const rentalCalc = getRentSystem().getRentalCalc();
        
        // 3. 生成表头
        rentalCalc.创建租金测算表表头(1, 10);
        
        // 4. 生成使用每期适用利率的测算表
        if (!rentalCalc.使用每期适用利率生成测算表()) {
            throw new Error("生成测算表失败");
        }
        
        // 5. 生成每期适用利率列（M列）
        rentalCalc.创建租金测算表表头(13, 13);
        rentalCalc.每期适用利率();
        
        // 6. 生成现金流量表
        getRentSystem().生成现金流量表();
        
        // 7. 生成月间隔（调期功能）
        rentalCalc.生成月间隔();
        
        console.log("========== 每期适用利率测算表生成完成 ==========");
        console.log("每期适用利率测算表生成成功！现在可以修改M列的利率值来实现不同的利率设置。");
        
        return true;
    } catch (error) {
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: '调2_每期适用利率', function: '调2_每期适用利率' });
        } else {
            console.error("生成失败：" + error.message);
        }
        return false;
    }
}

// 每期适用利率 + 自定义分段利率
function 调2_每期适用利率_自定义利率(前几期 = 4, 前期利率 = 0.04, 后期利率 = 0.035) {
    try {
        console.log(`========== 开始生成每期适用利率测算表（自定义利率）==========`);
        console.log(`前${前几期}期利率: ${前期利率 * 100}%`);
        console.log(`第${前几期 + 1}期起利率: ${后期利率 * 100}%`);
        
        // 1. 清除原有数据
        getRentSystem().清除();
        
        // 2. 获取缓存的租金测算实例（P1-3修复：避免重复创建）
        const rentalCalc = getRentSystem().getRentalCalc();
        
        // 3. 生成表头
        rentalCalc.创建租金测算表表头(1, 10);
        
        // 4. 生成使用每期适用利率的测算表
        if (!rentalCalc.使用每期适用利率生成测算表()) {
            throw new Error("生成测算表失败");
        }
        
        // 5. 生成每期适用利率列（M列）
        rentalCalc.创建租金测算表表头(13, 13);
        rentalCalc.每期适用利率();
        
        // 6. 设置分段利率
        const ws = Application.Worksheets("1租金测算表V1");
        const rentTableStartRow = rentalCalc.p.RentTableStartRow;
        const totalPeriods = rentalCalc.p.val("TotalPeriods");
        
        // 前N期使用前期利率
        if (前几期 > 0 && 前几期 <= totalPeriods) {
            ws.Range(`M${rentTableStartRow}:M${rentTableStartRow + 前几期 - 1}`).Value2 = 前期利率;
        }
        
        // 后面使用后期利率
        if (前几期 < totalPeriods) {
            ws.Range(`M${rentTableStartRow + 前几期}:M${rentTableStartRow + totalPeriods - 1}`).Value2 = 后期利率;
        }
        
        // 7. 重新计算
        ws.Calculate();
        
        // 8. 生成现金流量表
        getRentSystem().生成现金流量表();
        
        // 9. 生成月间隔（调期功能）
        rentalCalc.生成月间隔();
        
        console.log("========== 每期适用利率测算表生成完成 ==========");
        console.log(`每期适用利率测算表生成成功！利率设置：前${前几期}期: ${前期利率 * 100}% 第${前几期 + 1}期起: ${后期利率 * 100}%`);
        
        return true;
    } catch (error) {
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: '调2_每期适用利率_自定义利率', function: '调2_每期适用利率_自定义利率' });
        } else {
            console.error("生成失败：" + error.message);
        }
        return false;
    }
}

// 复杂融资租赁（租前期+租赁期，不等间隔还款）
function 调3_复杂融资租赁(options = {}) {
    try {
        console.log("========== 开始生成复杂融资租赁测算表 ==========");
        
        // 默认配置
        const config = {
            租前期期数: options.租前期期数 || 2,
            租赁期期数: options.租赁期期数 || 28,
            租前期利率: options.租前期利率 || null,  // null表示使用B5的值
            利率调整计划: options.利率调整计划 || []  // [{起始期: 1, 利率: 0.04}, ...]
        };
        
        // 计算总期数
        const totalPeriods = config.租前期期数 + config.租赁期期数;
        console.log(`配置: 租前期${config.租前期期数}期, 租赁期${config.租赁期期数}期, 总计${totalPeriods}期`);
        
        // 1. 更新总期数
        const ws = Application.Worksheets("1租金测算表V1");
        ws.Range("B10").Value2 = totalPeriods;
        
        // 2. 清除原有数据
        getRentSystem().清除();
        
        // 3. 获取缓存的租金测算实例（P1-3修复：避免重复创建）
        const rentalCalc = getRentSystem().getRentalCalc();
        
        // 4. 生成表头
        rentalCalc.创建租金测算表表头(1, 10);
        
        // 5. 生成使用每期适用利率的测算表
        if (!rentalCalc.使用每期适用利率生成测算表()) {
            throw new Error("生成测算表失败");
        }
        
        // 6. 生成每期适用利率列（M列）
        rentalCalc.创建租金测算表表头(13, 13);
        rentalCalc.每期适用利率();
        
        // 7. 设置浮动利率（根据利率调整计划）
        const rentTableStartRow = rentalCalc.p.RentTableStartRow;
        const defaultRate = rentalCalc.p.val("InterestRate");
        
        if (config.利率调整计划.length > 0) {
            // P0-4修复：按利率调整计划设置，每个条目的结束期为下一个条目的起始期-1
            for (var i = 0; i < config.利率调整计划.length; i++) {
                const plan = config.利率调整计划[i];
                const startRow = rentTableStartRow + plan.起始期 - 1;
                const endPeriod = (i + 1 < config.利率调整计划.length)
                    ? config.利率调整计划[i + 1].起始期 - 1
                    : totalPeriods;
                const endRow = rentTableStartRow + endPeriod - 1;
                ws.Range(`M${startRow}:M${endRow}`).Value2 = plan.利率;
                console.log(`利率设置: 第${plan.起始期}-${endPeriod}期 = ${(plan.利率 * 100).toFixed(2)}%`);
            }
        } else {
            // 默认：每年调整利率（示例：前12期用默认利率）
            console.log(`使用默认利率: ${(defaultRate * 100).toFixed(2)}%`);
        }
        
        // 8. 生成现金流量表
        getRentSystem().生成现金流量表();
        
        // 9. 生成月间隔（调期功能）- 支持不等间隔
        rentalCalc.生成月间隔();
        
        // 10. 设置不等间隔还款计划
        // 间隔说明：
        // - 第1期：放款日后6个月 → 间隔6
        // - 第2期：放款日后12个月 → 间隔6（从第1期算）
        // - 第3期：放款日后18个月 → 间隔6（从第2期算）
        // - 第4期：起租日后3个月 → 间隔3
        // - 第5期：起租日后6个月 → 间隔3
        // - 第6-29期：每6个月 → 间隔6
        // - 第30期：第29期后12个月 → 间隔12
        
        const intervals = [6, 6, 6, 3, 3];  // 第1-5期的间隔
        for (var i = 0; i < intervals.length; i++) {
            rentalCalc.改变自定义支付日(i + 1, intervals[i]);
        }
        // 第6-29期：间隔6
        for (var i = 6; i <= 29; i++) {
            rentalCalc.改变自定义支付日(i, 6);
        }
        // 第30期：间隔12
        rentalCalc.改变自定义支付日(30, 12);
        
        // 11. 重新计算
        ws.Calculate();
        
        console.log("========== 复杂融资租赁测算表生成完成 ==========");
        console.log("还款安排:");
        console.log("  - 租前期(第1-2期): 只付息不还本");
        console.log("  - 租赁期(第3-30期): 等额本金");
        console.log("  - 第1期: 放款日后6个月");
        console.log("  - 第2期: 放款日后12个月");
        console.log("  - 第3期: 起租日(放款日后18个月)");
        console.log("  - 第4期: 起租日后3个月");
        console.log("  - 第5期: 起租日后6个月");
        console.log("  - 第6-29期: 每6个月");
        console.log("  - 第30期: 第29期后12个月");
        
        console.log("复杂融资租赁测算表生成成功！" +
              "\n租前期: 第1-2期（只付息）" +
              "\n租赁期: 第3-30期（等额本金）" +
              "\n如需调整利率，请修改M列的值。");
        
        return true;
    } catch (error) {
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: '调3_复杂融资租赁', function: '调3_复杂融资租赁' });
        } else {
            console.error("生成失败：" + error.message);
        }
        return false;
    }
}

// 带自定义参数的复杂融资租赁
function 调3_复杂融资租赁_自定义(
    租金总额 = 500000000,
    票面利率 = 0.035,
    租前期期数 = 2,
    租赁期期数 = 28,
    偿还方式 = "等额本金（按期计息）"
) {
    try {
        console.log("========== 开始生成复杂融资租赁测算表（自定义参数）==========");
        console.log(`租金总额: ${租金总额}`);
        console.log(`票面利率: ${(票面利率 * 100).toFixed(2)}%`);
        console.log(`租前期: ${租前期期数}期`);
        console.log(`租赁期: ${租赁期期数}期`);
        console.log(`偿还方式: ${偿还方式}`);
        
        // 设置参数
        const ws = Application.Worksheets("1租金测算表V1");
        ws.Activate();
        ws.Range("B4").Value2 = 租金总额;
        ws.Range("B5").Value2 = 票面利率;
        ws.Range("B6").Value2 = 0;  // 保证金
        ws.Range("B8").Value2 = 1;  // 支付方式
        ws.Range("B10").Value2 = 租前期期数 + 租赁期期数;  // 总期数
        ws.Range("D10").Value2 = 6;  // 默认间隔月数
        ws.Range("B12").Value2 = 偿还方式;  // 偿还方式
        
        // 调用主函数
        return 调3_复杂融资租赁({
            租前期期数: 租前期期数,
            租赁期期数: 租赁期期数
        });
        
    } catch (error) {
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: '调3_复杂融资租赁_自定义', function: '调3_复杂融资租赁_自定义' });
        } else {
            console.error("生成失败：" + error.message);
        }
        return false;
    }
}

// 复杂融资租赁 + 浮动利率
function 调3_复杂融资租赁_浮动利率(利率计划 = null) {
    try {
        console.log("========== 开始生成复杂融资租赁测算表（浮动利率）==========");
        
        // 默认利率计划：模拟LPR年度调整
        // 假设：第1年4%，第2年3.5%，第3年3%
        const defaultRatePlan = 利率计划 || [
            { 起始期: 1, 利率: 0.04 },      // 第1-12期（第1年）
            { 起始期: 13, 利率: 0.035 },    // 第13-24期（第2年）
            { 起始期: 25, 利率: 0.03 }      // 第25-30期（第3年）
        ];
        
        console.log("利率调整计划:");
        for (const plan of defaultRatePlan) {
            console.log(`  第${plan.起始期}期起: ${(plan.利率 * 100).toFixed(2)}%`);
        }
        
        // 调用主函数
        return 调3_复杂融资租赁({
            租前期期数: 2,
            租赁期期数: 28,
            利率调整计划: defaultRatePlan
        });
        
    } catch (error) {
        // P1-8: 集成统一错误处理器
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: '调3_复杂融资租赁_浮动利率', function: '调3_复杂融资租赁_浮动利率' });
        } else {
            console.error("生成失败：" + error.message);
        }
        return false;
    }
}

console.log(`[mMain] 加载完成 - ${WORKFLOW_CONFIG.system.version}`);
