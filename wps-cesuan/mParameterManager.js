/**
 * ============== 参数管理器类模块 V3.3 ==============
 * 作者：徐晓冬
 * 版本：V3.3.20260430
 * 描述：参数管理器 - 配置、读写、验证、地址转换
 *
 * V3.3 变更（基于 V3.2 的改进）：
 * - [重构] 引入 _log() 分级日志（debug/info/warn/error），替代无区分的 console.log
 * - [重构] ValidateValueByRule 中 var 统一为 const/var，风格一致
 * - [清理] 移除死代码 ImportConfigFromJson（WPS JSA 不支持文件 I/O）
 * - [清理] 合并静态工厂：移除 AutoCreateFromSheet()，统一使用 Create()
 * - [增强] SetParameterChangeListener 支持范围重叠检测，不再仅精确匹配
 * - [文档] 添加职责分区注释，为后续拆分提供导航
 *
 * 职责分区导航：
 *   §1 日志工具          — _log(), setLogLevel()
 *   §2 初始化与常量      — constructor, Initialize, _initConstants, _initSheetReferences
 *   §3 配置管理          — CreateDefaultConfig, GetConfigValue, GetParameterNames
 *   §4 地址转换          — GetCellAddressA1, ConvertR1C1ToA1, ClearAddressCache
 *   §5 参数值读写        — ReadParameterValue, SetParameterValue, SetParametersBatch
 *   §6 类型转换与工具    — _convertValue, _getDefaultValue, IsValidDate, _formatDate
 *   §7 验证功能          — ValidateValueByRule, ValidateParameter, ValidateAllParameters
 *   §8 工作表管理        — GetOrCreateWorksheet, InitializeWorksheetFormat, CreateParameterInputArea
 *   §9 动态属性与快捷访问 — _createParameterAccessors, val(), addr(), param(), getParam()
 *   §10 公式与重置       — ApplyDefaultFormulas, ResetToDefault
 *   §11 调试与报告       — PrintDebugInfo, GenerateParameterReport, GetAllParametersSummary
 *   §12 事件处理         — SetParameterChangeListener
 *   §13 导出             — ExportConfigToJson
 *   §14 静态工厂         — Create()
 *   §15 参数访问器类     — clsParamAccessor
 * ====================================================
 */

// ============== §15 参数访问器类 ==============
class clsParamAccessor {
    constructor(manager, paramName) {
        this._manager = manager;
        this._name = paramName;
    }

    get value() { return this._manager.ReadParameterValue(this._name); }
    get cellA1() { return this._manager.GetCellAddressA1(this._name); }
    get cellR1C1() {
        const config = this._manager.m_config[this._name];
        return (config && config.CellAddress) ? config.CellAddress : "";
    }
    get config() { return this._manager.m_config[this._name]; }
    get displayName() {
        const config = this._manager.m_config[this._name];
        return (config && config.DisplayName) ? config.DisplayName : "";
    }
    get dataType() {
        const config = this._manager.m_config[this._name];
        return (config && config.DataType) ? config.DataType : "";
    }
    get defaultValue() {
        const config = this._manager.m_config[this._name];
        return (config && config.DefaultValue !== undefined) ? config.DefaultValue : null;
    }
    get isRequired() {
        const config = this._manager.m_config[this._name];
        return (config && config.IsRequired) ? config.IsRequired : false;
    }
    get validationRule() {
        const config = this._manager.m_config[this._name];
        return (config && config.ValidationRule) ? config.ValidationRule : "";
    }
    get description() {
        const config = this._manager.m_config[this._name];
        return (config && config.Description) ? config.Description : "";
    }

    set(val) { return this._manager.SetParameterValue(this._name, val); }
    validate() { return this._manager.ValidateParameter(this._name); }
    toString() { return `${this.displayName || this._name}: ${this.value}`; }
    getInfo() {
        return {
            name: this._name,
            displayName: this.displayName,
            value: this.value,
            cellA1: this.cellA1,
            cellR1C1: this.cellR1C1,
            dataType: this.dataType,
            defaultValue: this.defaultValue,
            isRequired: this.isRequired,
            validation: this.validate()
        };
    }
}

// ============== 参数管理器类 ==============
class clsParameterManager {
    constructor() {
        this.MODULE_NAME = "clsParameterManager";
        this.ModuleModifyDate = (typeof VERSION_DATE !== 'undefined') ? VERSION_DATE : "20260430";

        // 核心状态
        this.m_worksheet = null;
        this._isInitialized = false;
        this.targetSheetName = null;
        this._addressCache = {};

        // §1 日志级别：'debug' | 'info' | 'warn' | 'error'
        // 生产环境建议 'warn'，开发调试用 'debug'
        this._logLevel = 'info';

        // 加载配置（构造时即可用，不依赖 _isInitialized）
        this.m_config = this.CreateDefaultConfig();

        // 初始化常量
        this._initConstants();

        // 动态生成所有参数的属性
        this._createParameterAccessors();

        this._log('info', '类实例创建');
    }

    // ==================== §1 日志工具 ====================

    /**
     * _log - 分级日志输出
     *
     * @param {'debug'|'info'|'warn'|'error'} level - 日志级别
     * @param {string} message - 日志内容
     */
    _log(level, message) {
        const LEVEL_PRIORITY = { debug: 0, info: 1, warn: 2, error: 3 };
        const currentPriority = LEVEL_PRIORITY[this._logLevel] !== undefined
            ? LEVEL_PRIORITY[this._logLevel]
            : LEVEL_PRIORITY.info;
        const msgPriority = LEVEL_PRIORITY[level] !== undefined
            ? LEVEL_PRIORITY[level]
            : LEVEL_PRIORITY.info;

        if (msgPriority < currentPriority) return;

        const prefix = `[${this.MODULE_NAME}]`;
        switch (level) {
            case 'debug':
                console.log(`${prefix}[DEBUG] ${message}`);
                break;
            case 'info':
                console.log(`${prefix} ${message}`);
                break;
            case 'warn':
                console.warn(`${prefix}[WARN] ${message}`);
                break;
            case 'error':
                console.error(`${prefix}[ERROR] ${message}`);
                break;
            default:
                console.log(`${prefix} ${message}`);
        }
    }

    /**
     * setLogLevel - 设置日志级别
     *
     * @param {'debug'|'info'|'warn'|'error'} level - 日志级别
     */
    setLogLevel(level) {
        const validLevels = ['debug', 'info', 'warn', 'error'];
        if (validLevels.indexOf(level) !== -1) {
            this._logLevel = level;
            this._log('info', `日志级别设置为: ${level}`);
        } else {
            this._log('warn', `无效的日志级别: ${level}，有效值: ${validLevels.join(', ')}`);
        }
    }

    // ==================== §2 初始化与常量 ====================

    /**
     * _initSheetReferences - 初始化工作表引用
     *
     * 每个工作表引用获取失败时不影响其他工作表。
     */
    _initSheetReferences() {
        const sheets = [
            { name: "1租金测算表V1", propName: "m_sourceSheet" },
            { name: "银行承兑汇票", propName: "m_billSheet" },
            { name: "还款设置", propName: "m_repaymentSettingSheet" },
            { name: "贷款基础利率", propName: "m_loanRateSheet" }
        ];

        sheets.forEach(sheet => {
            this[`${sheet.propName}Name`] = sheet.name;
            try {
                this[sheet.propName] = Application.Worksheets(sheet.name);
            } catch (e) {
                this._log('warn', `工作表'${sheet.name}'获取失败: ${e.message}`);
                this[sheet.propName] = null;
            }
        });
    }

    _initConstants() {
        // 表格结构常量
        this.m_RowStart = 26;
        this.m_SheetNamerow = 1;
        this.m_MinRowHeight = 15;
        this.m_MinColumnWidth = 20;

        // 列定义（批量初始化）
        const colMap = {
            PERIOD: "A", DATE: "B", RENT: "C", PRINCIPAL: "D", INTEREST: "E",
            RENT_BALANCE: "F", PRINCIPAL_BALANCE: "G", REMAINING_BALANCE: "H",
            PAID_RENT: "I", MONTH_INTERVAL: "J", CUSTOM_INTERVAL: "K",
            PRINCIPAL_RATIO: "L", RatePerPeriod: "M"
        };
        Object.entries(colMap).forEach(([k, v]) => this[`m_COL_${k}`] = v);

        // 颜色常量 — 直接引用 mShared_constants.js 中的 COLORS 统一色板
        // 不再独立计算，确保全项目颜色值一致
        this.m_COLOR_WHITE = COLORS.WHITE;
        this.m_COLOR_BLUE = COLORS.HEADER_BLUE;
        this.m_COLOR_YELLOW = COLORS.YELLOW;
        this.m_COLOR_LIGHT_GREEN = COLORS.LIGHT_GREEN;
        this.m_COLOR_LIGHT_RED = COLORS.LIGHT_RED;
        this.m_COLOR_BLACK = COLORS.BLACK;
        this.m_COLOR_GRAY = COLORS.GRAY;
    }

    // ==================== 核心属性 ====================

    get WsTarget() {
        if (!this._isInitialized) throw new Error("参数管理器未初始化");
        return this.m_worksheet;
    }

    get IsInitialized() { return this._isInitialized; }
    get RowStart() { return this.m_RowStart; }
    get SheetNamerow() { return this.m_SheetNamerow; }
    get RentTableStartRow() {
        return this.m_RowStart + this.m_SheetNamerow * 2;
    }
    get CashFlowTablerowStart() {
        const totalPeriods = this.val("TotalPeriods");
        const rentStart = this.RentTableStartRow;
        const result = rentStart + totalPeriods + 6;
        if (isNaN(result)) {
            this._log('warn', `CashFlowTablerowStart计算结果为NaN, totalPeriods=${totalPeriods}, rentStart=${rentStart}`);
        }
        return result;
    }

    // ==================== §2 初始化方法 ====================

    Initialize(sheetName = "1租金测算表V1") {
        try {
            this._log('info', '开始初始化工作表');

            // 初始化默认日期值
            this.m_config.LeaseStartDate.DefaultValue = this.GetDefaultDate();
            this.m_config.FirstPaymentDate.DefaultValue = this.GetDefaultFirstPaymentDate();

            // 延迟获取工作表引用
            this._initSheetReferences();

            this.m_worksheet = this.GetOrCreateWorksheet(sheetName);
            this.targetSheetName = this.m_worksheet.Name;
            this._isInitialized = true;

            // 初始化时自动清除地址缓存
            this.ClearAddressCache();

            // 再次创建 accessors 以确保 Initialize 后状态一致
            this._createParameterAccessors();

            this._log('info', `初始化成功，工作表: ${this.targetSheetName}`);
            return true;
        } catch (error) {
            this._log('error', `初始化失败: ${error.message}`);
            return false;
        }
    }

    WorksheetExists(sheetName) {
        try { Application.Worksheets(sheetName); return true; } catch (e) { return false; }
    }

    // ==================== §3 配置管理 ====================

    CreateDefaultConfig() {
        const config = {};

        const addParam = (name, props) => {
            config[name] = {
                CellAddress: "", DefaultValue: null, VbaFormat: "",
                DisplayName: "", ValidationRule: "", DataType: "String",
                IsRequired: false, Description: "", ...props
            };
        };

        // ==================== 价格参数配置 ====================
        addParam("Principal", { CellAddress: "R4C2", DefaultValue: 100000000, VbaFormat: "#,##0.00",
            DisplayName: "租赁成本", ValidationRule: ">0", DataType: "Double", IsRequired: true, Description: "租赁资产的总成本" });

        addParam("InterestRate", { CellAddress: "R5C2", DefaultValue: 0.03, VbaFormat: "0.0000%",
            DisplayName: "租赁票面利率", ValidationRule: ">0", DataType: "Double", IsRequired: true, Description: "租赁合同约定的年利率" });

        addParam("Deposit", { CellAddress: "R6C4", DefaultValue: 0, VbaFormat: "#,##0.00",
            DisplayName: "租赁保证金", ValidationRule: ">=0", DataType: "Double", Description: "租赁保证金金额", DefaultFormula: "=(B4*B6)" });

        addParam("DepositMarginRate", { CellAddress: "R6C2", DefaultValue: 0.01, VbaFormat: "0.00%",
            DisplayName: "保证金费率", ValidationRule: ">=0", DataType: "Double", Description: "保证金费率（如0.01表示1%）" });

        addParam("NominalPrice", { CellAddress: "R8C2", DefaultValue: 1, VbaFormat: "#,##0.00",
            DisplayName: "名义货价", ValidationRule: ">=0", DataType: "Double", IsRequired: true, Description: "租赁合同中的名义货价" });

        addParam("TotalPeriods", { CellAddress: "R10C2", DefaultValue: 12, VbaFormat: "0",
            DisplayName: "总期数", ValidationRule: ">0", DataType: "Long", IsRequired: true,
            Description: "租金测算的总期数", MinValue: 1, MaxValue: 1000 });

        addParam("PaymentsPerYear", { CellAddress: "R11C2", DefaultValue: 2, VbaFormat: "0",
            DisplayName: "每年还款次数", ValidationRule: ">0", DataType: "Long", IsRequired: true,
            Description: "每年的还款次数", DefaultFormula: "=12/D10" });

        addParam("RepaymentMethod", { CellAddress: "R12C2", DefaultValue: "等额本息（后付）",
            DisplayName: "偿还方式", ValidationRule: "TRUE", DataType: "String", IsRequired: true, Description: "租金偿还方式",
            DataSource: "还款设置!$A$2:$A$9" });

        addParam("LeaseStartDate", { CellAddress: "R13C2", DefaultValue: null, VbaFormat: "yyyy-mm-dd",
            DisplayName: "放款日", ValidationRule: "TRUE", DataType: "Date", IsRequired: true, Description: "资金实际放出的日期" });

        addParam("FirstPaymentDate", { CellAddress: "R21C2", DefaultValue: null, VbaFormat: "yyyy-mm-dd",
            DisplayName: "起租日（第1期支付日）", ValidationRule: "TRUE", DataType: "Date", IsRequired: true,
            Description: "第1期租金支付日（租期开始日）", DefaultFormula: "=EDATE(B13,6)" });

        addParam("PreLeaseInterval", { CellAddress: "R22C2", DefaultValue: 3, VbaFormat: "0",
            DisplayName: "租前期间隔（月）", ValidationRule: "1,2,3,6,12", DataType: "Long", IsRequired: true,
            Description: "租前期利息支付间隔(月)", DataSource: "还款设置!$P$2:$P$6" });

        addParam("PreLeaseMonths", { CellAddress: "R23C2", DefaultValue: 6, VbaFormat: "0",
            DisplayName: "租前期月数", ValidationRule: ">=0", DataType: "Long", IsRequired: true,
            Description: "从放款日到起租日的月数", DefaultFormula: "=DATEDIF(B13,B21,\"M\")" });

        addParam("PaymentInterval", { CellAddress: "R10C4", DefaultValue: 6, VbaFormat: "0",
            DisplayName: "支付间隔", ValidationRule: ">0", DataType: "Long", IsRequired: true, Description: "支付间隔(月)",
            DataSource: "还款设置!$B$2:$B$5" });

        addParam("ProjectDurationYears", { CellAddress: "R11C4", DefaultValue: 3, VbaFormat: "0",
            DisplayName: "项目时长/年", ValidationRule: ">0", DataType: "Long", Description: "项目时长/年", DefaultFormula: "=D10*B10/12" });

        addParam("ProjectDurationMonths", { CellAddress: "R12C4", DefaultValue: 36, VbaFormat: "0",
            DisplayName: "项目时长/月", ValidationRule: ">0", DataType: "Long", Description: "项目时长/月", DefaultFormula: "=D10*B10" });

        addParam("StaticValueConversion", { CellAddress: "R1C4", DefaultValue: false, VbaFormat: "General",
            DisplayName: "静态值转换", ValidationRule: "", DataType: "Boolean", Description: "是否将公式转换为静态值" });

        // ==================== 利率参数配置 ====================
        addParam("LPRDate", { CellAddress: "R11C9", DefaultValue: "", VbaFormat: "yyyy-mm-dd",
            DisplayName: "LPR发布日期", ValidationRule: "TRUE", DataType: "Date", IsRequired: true, Description: "LPR利率发布日期",
            DataSource: "还款设置!$N$2:$N$13" });

        addParam("LPRBenchmarkRate", { CellAddress: "R10C7", DefaultValue: 0.0425, VbaFormat: "0.00%",
            DisplayName: "LPR基准利率", ValidationRule: ">=0,<=0.2", DataType: "Double", Description: "LPR基准利率",
            FormulaDependency: "LPRPeriod,LPRDate", DataSource: "贷款基础利率!$B$2:$D$14", IsCalculated: true,
            DefaultFormula: "=IFS(G11=5,VLOOKUP(I11,贷款基础利率!$B$2:$D$14,3,FALSE)/100,G11=1,VLOOKUP(I11,贷款基础利率!$B$2:$D$14,2,FALSE)/100)" });

        addParam("LPRPeriod", { CellAddress: "R11C7", DefaultValue: 5, VbaFormat: "0",
            DisplayName: "LPR期限选择", ValidationRule: "1,5", DataType: "Long", IsRequired: true, Description: "LPR期限选择（1=1年期，5=5年期）",
            DataSource: "还款设置!$L$2:$L$3" });

        addParam("FloatingBasisPoints", { CellAddress: "R10C9", DefaultValue: 0, VbaFormat: "0.00",
            DisplayName: "浮动基点(BP)", ValidationRule: "", DataType: "Double", Description: "相对于LPR的浮动基点数", DefaultFormula: "=(B5-G10)*10000" });

        addParam("RateOption", { CellAddress: "R5C4", DefaultValue: "固定利率",
            DisplayName: "利率选择", ValidationRule: "", DataType: "String", Description: "利率选择",
            DataSource: "还款设置!$K$2:$K$3", IsCalculated: true });

        addParam("ActualPayment", { CellAddress: "R4C4", DefaultValue: 0, VbaFormat: "#,##0.00",
            DisplayName: "实际付款", ValidationRule: "", DataType: "Double", Description: "实际付款是租赁本金减去保证金",
            FormulaDependency: "Principal,Deposit", IsCalculated: true, DefaultFormula: "=B4-D6" });

        addParam("LPRRateDescription", { CellAddress: "R12C7", DefaultValue: "",
            DisplayName: "LPR利率描述", ValidationRule: "", DataType: "String", Description: "LPR利率描述",
            FormulaDependency: "LPRPeriod,LPRBenchmarkRate", IsCalculated: true });

        addParam("ActualInterestRate", { CellAddress: "R10C8", DefaultValue: 0.03, VbaFormat: "0.00%",
            DisplayName: "实际利率", ValidationRule: ">0", DataType: "Double", Description: "实际执行利率",
            FormulaDependency: "LPRBenchmarkRate,FloatingBasisPoints", IsCalculated: true, DefaultFormula: "=G10+I10/10000" });

        // ==================== 方案要素配置 ====================
        addParam("Lessee", { CellAddress: "R5C7", DisplayName: "承租人", ValidationRule: "TRUE",
            DataType: "String", Description: "承租人名称" });

        addParam("Guarantor", { CellAddress: "R6C7", DisplayName: "担保人", ValidationRule: "TRUE",
            DataType: "String", Description: "担保人名称" });

        addParam("GuaranteeMethod", { CellAddress: "R7C7", DisplayName: "担保方式", ValidationRule: "TRUE",
            DataType: "String", Description: "担保方式" });

        addParam("LeaseMethod", { CellAddress: "R8C7", DisplayName: "租赁方式", ValidationRule: "TRUE",
            DataType: "String", Description: "租赁方式" });

        // ==================== 经纪人参数配置 ====================
        addParam("BrokerPaymentMethod", { CellAddress: "R14C7", DefaultValue: "一次性支付-放款时",
            DisplayName: "经纪人费用支付方式", ValidationRule: "TRUE", DataType: "String", IsRequired: true, Description: "经纪人费用的支付方式",
            DataSource: "还款设置!$D$2:$D$6" });

        addParam("BrokerFeeRate", { CellAddress: "R16C7", DefaultValue: 0.001, VbaFormat: "0.00%",
            DisplayName: "经纪人费用比例", ValidationRule: ">=0", DataType: "Double", IsRequired: true, Description: "经纪人费用比例（如0.02表示2%）" });

        addParam("BrokerTotalFee", { CellAddress: "R15C7", DefaultValue: 0, VbaFormat: "#,##0.00",
            DisplayName: "经纪人总费用", ValidationRule: ">=0", DataType: "Double", Description: "经纪人总费用(BrokerTotalFee)", DefaultFormula: "=B4*G16" });

        return config;
    }

    // ==================== §6 类型转换与工具 ====================

    /**
     * RGB - 已弃用，请使用 mShared_constants.js 中的全局 RGB() 函数
     * @deprecated 保留此方法仅为向后兼容，新代码请直接调用 RGB(r, g, b)
     */
    RGB(r, g, b) {
        return RGB(r, g, b);
    }

    _formatDate(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        const dateStr = `${year}-${month}-${day}`;

        if (isNaN(new Date(dateStr).getTime())) {
            this._log('warn', `生成的日期格式可能不兼容: ${dateStr}`);
        }
        return dateStr;
    }

    GetDefaultDate() {
        return this._formatDate(new Date(new Date().getFullYear(), new Date().getMonth() + 1, 15));
    }

    GetDefaultFirstPaymentDate() {
        return this._formatDate(new Date(new Date().getFullYear(), new Date().getMonth() + 7, 15));
    }

    // ==================== §8 工作表管理 ====================

    GetOrCreateWorksheet(sheetName) {
        try {
            var worksheet;
            try { worksheet = Application.Sheets(sheetName); return worksheet; }
            catch (e) {
                this._log('info', "工作表不存在，正在创建: " + sheetName);
                worksheet = Application.Sheets.Add();
                worksheet.Name = sheetName;
                Application.Calculate();
                try { this.InitializeWorksheetFormat(worksheet); }
                catch (formatError) { this._log('warn', `格式初始化失败，但工作表已创建: ${formatError.message}`); }
                return worksheet;
            }
        } catch (error) {
            this._log('error', `获取或创建工作表失败: ${error.message}`);
            return null;
        }
    }

    _IsWorksheetWritable(worksheet) {
        try {
            const testCell = worksheet.Range("A1");
            testCell.Value2 = "测试";
            if (testCell.Value2 !== "测试") return false;
            testCell.Value2 = "";
            return true;
        } catch (e) { return false; }
    }

    InitializeWorksheetFormat(worksheet) {
        try {
            this._log('debug', '开始初始化工作表格式');

            if (!this._IsWorksheetWritable(worksheet)) {
                this._log('warn', '工作表只读，跳过格式设置');
                return;
            }

            try { worksheet.Cells.Clear(); } catch (e) { /* ignore */ }

            try {
                const titleCell = worksheet.Range("A1");
                titleCell.Value2 = "参数配置表";
                titleCell.Font.Bold = true;
                titleCell.Font.Size = 16;
                worksheet.Range("A1:C1").Merge();
            } catch (e) { /* ignore */ }

            try {
                const headers = ["参数名称", "参数值", "说明"];
                headers.forEach(function(h, i) {
                    const cell = worksheet.Range("A" + (i + 3));
                    cell.Value2 = h;
                    cell.Font.Bold = true;
                    try { cell.Interior.Color = RGB(200, 200, 200); } catch (e) { /* ignore */ }
                });
            } catch (e) { /* ignore */ }

            try {
                worksheet.Columns("A:A").ColumnWidth = 20;
                worksheet.Columns("B:B").ColumnWidth = 15;
                worksheet.Columns("C:C").ColumnWidth = 30;
            } catch (e) { /* ignore */ }

            this._log('debug', '工作表格式初始化完成');
        } catch (error) {
            this._log('error', `工作表格式初始化失败: ${error.message}`);
        }
    }

    CreateParameterInputArea() {
        try {
            if (!this._isInitialized) {
                this._log('warn', '参数管理器未初始化，无法创建输入区域');
                return false;
            }

            this._log('info', '开始创建参数输入区域');
            var row = 13;

            try {
                const areaTitle = this.m_worksheet.Range("A" + row);
                areaTitle.Value2 = "参数输入区域";
                areaTitle.Font.Bold = true;
                areaTitle.Font.Size = 14;
                try { areaTitle.Interior.Color = RGB(173, 216, 230); } catch (e) { /* ignore */ }
                this.m_worksheet.Range("A" + row + ":C" + row).Merge();
            } catch (e) { /* ignore */ }

            row += 2;
            const paramNames = this.GetParameterNames();
            var createdCount = 0;

            for (const paramName of paramNames) {
                const config = this.m_config[paramName];
                try {
                    const nameCell = this.m_worksheet.Range("A" + row);
                    nameCell.Value2 = config.DisplayName;
                    nameCell.Font.Bold = true;

                    const valueCell = this.m_worksheet.Range("B" + row);
                    valueCell.Value2 = config.DefaultValue;
                    if (config.VbaFormat) try { valueCell.NumberFormat = config.VbaFormat; } catch (e) { /* ignore */ }

                    const descCell = this.m_worksheet.Range("C" + row);
                    descCell.Value2 = config.Description || "";
                    try {
                        descCell.Font.Color = RGB(128, 128, 128);
                        descCell.Font.Italic = true;
                    } catch (e) { /* ignore */ }

                    config.CellAddressA1 = "B" + row;
                    try { this.m_worksheet.Range("A" + row + ":C" + row).Borders.LineStyle = 1; } catch (e) { /* ignore */ }
                    createdCount++;
                } catch (e) { /* ignore */ }
                row++;
            }

            this._log('info', `参数输入区域创建完成，成功创建 ${createdCount}/${paramNames.length} 个参数`);
            return createdCount > 0;
        } catch (error) {
            this._log('error', `创建参数输入区域失败: ${error.message}`);
            return false;
        }
    }

    // ==================== §4 地址转换 ====================

    GetCellAddressA1(parameterName) {
        if (this._addressCache[parameterName]) return this._addressCache[parameterName];

        const r1c1Address = this.GetConfigValue(parameterName, "CellAddress");
        if (!r1c1Address) return "";

        const a1Address = this.ConvertR1C1ToA1(r1c1Address, parameterName);
        if (a1Address) this._addressCache[parameterName] = a1Address;

        return a1Address;
    }

    ClearAddressCache() {
        this._addressCache = {};
        this._log('debug', '地址缓存已清除');
    }

    ConvertR1C1ToA1(r1c1Address, parameterName) {
        try {
            const rPos = r1c1Address.indexOf("R");
            const cPos = r1c1Address.indexOf("C");
            if (rPos === -1 || cPos === -1) return this._convertR1C1ToA1Manual(r1c1Address);

            const rowNumber = parseInt(r1c1Address.substring(rPos + 1, cPos));
            const colNumber = parseInt(r1c1Address.substring(cPos + 1));

            // 如果工作表已初始化，使用工作表的 Cells 方法
            if (this.m_worksheet) {
                return this.m_worksheet.Cells(rowNumber, colNumber).Address(false, false);
            }

            // 否则手动转换 R1C1 到 A1 格式
            return this._convertR1C1ToA1Manual(r1c1Address);
        } catch (error) {
            this._log('warn', `ConvertR1C1ToA1 失败，回退手动转换: ${error.message}`);
            return this._convertR1C1ToA1Manual(r1c1Address);
        }
    }

    /**
     * _convertR1C1ToA1Manual - 手动将 R1C1 格式转换为 A1 格式
     *
     * @param {string} r1c1Address - R1C1 格式的地址（如 "R10C2"）
     * @returns {string} A1 格式的地址（如 "B10"）
     */
    _convertR1C1ToA1Manual(r1c1Address) {
        try {
            const rPos = r1c1Address.indexOf("R");
            const cPos = r1c1Address.indexOf("C");
            if (rPos === -1 || cPos === -1) return "";

            const rowNumber = parseInt(r1c1Address.substring(rPos + 1, cPos));
            const colNumber = parseInt(r1c1Address.substring(cPos + 1));

            // 将列号转换为列字母（1=A, 2=B, ..., 26=Z, 27=AA, ...）
            var colLetter = "";
            var col = colNumber;
            while (col > 0) {
                col--; // 调整为 0 索引
                colLetter = String.fromCharCode(65 + (col % 26)) + colLetter;
                col = Math.floor(col / 26);
            }

            return colLetter + rowNumber;
        } catch (error) {
            this._log('error', `_convertR1C1ToA1Manual 失败: ${error.message}`);
            return "";
        }
    }

    GetDefaultCellAddress(parameterName) {
        if (parameterName && this.m_config[parameterName]) {
            const address = this.GetConfigValue(parameterName, "CellAddress");
            this._log('debug', `GetDefaultCellAddress: 参数'${parameterName}'返回配置地址: ${address}`);
            return address;
        }
        this._log('warn', `GetDefaultCellAddress: 无法获取参数地址，返回空字符串`);
        return "";
    }

    // ==================== §3 配置访问 ====================

    GetConfigValue(parameterName, configKey) {
        try {
            const paramConfig = this.m_config[parameterName];
            if (!paramConfig) return "";
            const value = paramConfig[configKey];
            return (value !== undefined && value !== null) ? value : "";
        } catch (e) {
            this._log('warn', "读取配置值出错: " + parameterName + ", 键: " + configKey);
            return "";
        }
    }

    // ==================== §5 参数值读取 ====================

    ReadParameterValue(parameterName) {
        try {
            if (!this.m_worksheet || !this.m_config[parameterName]) return null;

            const config = this.m_config[parameterName];
            const cellAddressA1 = this.GetCellAddressA1(parameterName);
            if (!cellAddressA1) {
                this._log('warn', `参数'${parameterName}'无单元格地址，使用默认值: ${config.DefaultValue}`);
                return this._getDefaultValue(config.DefaultValue, config.DataType);
            }

            const cellValue = this.m_worksheet.Range(cellAddressA1).Value2;
            if (cellValue === null || cellValue === undefined || cellValue === "") {
                this._log('warn', `参数'${parameterName}'单元格为空，使用默认值: ${config.DefaultValue}`);
                return this._getDefaultValue(config.DefaultValue, config.DataType);
            }

            var resultValue = this._convertValue(cellValue, config.DataType);

            if (!this.ValidateValueByRule(resultValue, config.ValidationRule)) {
                this._log('warn', `参数'${parameterName}'值${resultValue}不符合规则'${config.ValidationRule}'，使用默认值: ${config.DefaultValue}`);
                return this._getDefaultValue(config.DefaultValue, config.DataType);
            }

            return resultValue;
        } catch (error) {
            this._log('error', `读取参数值失败: ${parameterName}, 错误: ${error.message}`);
            const config = this.m_config[parameterName];
            const defaultValue = config ? config.DefaultValue : null;
            const dataType = config ? config.DataType : null;
            return this._getDefaultValue(defaultValue, dataType);
        }
    }

    _convertValue(value, dataType) {
        switch (dataType) {
            case "Long": return !isNaN(value) ? parseInt(value) : 0;
            case "Double":
                const result = parseFloat(value);
                return isFinite(result) ? result : 0.0;
            case "Date":
                // 修复：Value2读取日期返回Excel数值（如44896），必须用fromExcelDate转换
                if (typeof value === 'number') {
                    if (typeof DateUtils !== 'undefined' && typeof DateUtils.fromExcelDate === 'function') {
                        return DateUtils.fromExcelDate(value);
                    }
                    // 降级方案：手动转换Excel日期数值
                    var excelBase = new Date(1899, 11, 30).getTime();
                    return new Date(excelBase + value * 86400000);
                }
                return this.IsValidDate(value) ? new Date(value) : new Date();
            case "Boolean":
                const str = String(value).toUpperCase().trim();
                if (["TRUE", "1", "是", "YES", "Y", "T"].includes(str)) return true;
                if (["FALSE", "0", "否", "NO", "N", "F"].includes(str)) return false;
                return !isNaN(value) ? parseInt(value) === 1 : false;
            case "String": return String(value);
            default: return value;
        }
    }

    _getDefaultValue(defaultValue, dataType) {
        try {
            switch (dataType) {
                case "Long": return !isNaN(defaultValue) ? parseInt(defaultValue) : 0;
                case "Double": return !isNaN(defaultValue) ? parseFloat(defaultValue) : 0.0;
                case "Date": return this.IsValidDate(defaultValue) ? defaultValue : new Date();
                case "Boolean": return defaultValue === true || defaultValue === "TRUE" || defaultValue === "1";
                case "String": return defaultValue !== null && defaultValue !== undefined ? String(defaultValue) : "";
                default: return defaultValue;
            }
        } catch (e) {
            const defaults = { Long: 0, Double: 0.0, Date: new Date(), Boolean: false, String: "" };
            const result = defaults[dataType];
            return (result !== undefined && result !== null) ? result : null;
        }
    }

    IsValidDate(date) {
        if (date instanceof Date) return !isNaN(date.getTime());
        if (typeof date === 'string') {
            if (!isNaN(new Date(date).getTime())) return true;
            const excelDate = parseFloat(date);
            if (!isNaN(excelDate) && excelDate > 0) {
                return !isNaN(new Date(1899, 11, 30).getTime() + excelDate * 86400000);
            }
        }
        return false;
    }

    // ==================== §5 参数值设置 ====================

    SetParameterValue(parameterName, value) {
        try {
            if (!this._isInitialized || !this.m_config[parameterName]) return false;
            if (!this.ValidateParameterValueRange(parameterName, value)) {
                this._log('warn', `参数值验证失败：${parameterName}=${value}`);
                return false;
            }
            const cellAddressA1 = this.GetCellAddressA1(parameterName);
            if (!cellAddressA1) return false;
            this.m_worksheet.Range(cellAddressA1).Value2 = value;
            return true;
        } catch (error) {
            this._log('error', `设置参数值失败: ${error.message}`);
            return false;
        }
    }

    SetParametersBatch(parameters) {
        var successCount = 0, failureCount = 0;
        for (const [paramName, value] of Object.entries(parameters)) {
            this.SetParameterValue(paramName, value) ? successCount++ : failureCount++;
        }
        return { successCount, failureCount };
    }

    // ==================== §7 验证功能 ====================

    /**
     * ValidateValueByRule - 根据规则验证值
     *
     * 支持规则格式：
     * - 比较规则: ">0", ">=0", "<0", "<=0", "!=0"
     * - 范围规则: ">=0,<=0.2" (min,max 格式)
     * - 枚举规则: "1,2,3,6,12" (逗号分隔的允许值列表)
     * - 特殊规则: "TRUE", "Boolean"
     *
     * @param {*} value - 待验证的值
     * @param {string} rule - 验证规则字符串
     * @returns {boolean} 是否通过验证
     */
    ValidateValueByRule(value, rule) {
        if (!rule) return true;

        // 特殊标记规则
        switch (rule) {
            case "TRUE": return true;
            case "Boolean": return typeof value === 'boolean' || ["TRUE", "FALSE", "1", "0"].includes(String(value));
            case ">0": return !isNaN(value) && value > 0;
            case ">=0": return !isNaN(value) && value >= 0;
            case "<0": return !isNaN(value) && value < 0;
            case "<=0": return !isNaN(value) && value <= 0;
        }

        // 范围规则: ">=0,<=0.2" 格式（包含比较运算符的逗号分隔）
        if (rule.indexOf(">=") !== -1 || rule.indexOf("<=") !== -1 || rule.indexOf(">") !== -1 || rule.indexOf("<") !== -1) {
            const parts = rule.split(",");
            var allPassed = true;
            for (var i = 0; i < parts.length; i++) {
                const p = parts[i].trim();
                if (p.indexOf(">=") === 0) {
                    if (isNaN(value) || value < parseFloat(p.substring(2))) { allPassed = false; break; }
                } else if (p.indexOf("<=") === 0) {
                    if (isNaN(value) || value > parseFloat(p.substring(2))) { allPassed = false; break; }
                } else if (p.indexOf(">") === 0) {
                    if (isNaN(value) || value <= parseFloat(p.substring(1))) { allPassed = false; break; }
                } else if (p.indexOf("<") === 0) {
                    if (isNaN(value) || value >= parseFloat(p.substring(1))) { allPassed = false; break; }
                }
            }
            if (allPassed) return true;
        }

        // 枚举规则: "1,5" 或 "1,2,3,6,12" 格式（纯数字逗号分隔）
        const enumValues = rule.split(",").map(function(v) { return v.trim(); });
        const isNumericEnum = enumValues.every(function(v) { return !isNaN(v); });
        if (isNumericEnum) {
            return enumValues.some(function(v) { return Number(value) === Number(v); });
        }

        // 字符串枚举: 逗号分隔的字符串值
        return enumValues.indexOf(String(value)) !== -1;
    }

    ValidateParameterValueRange(parameterName, value) {
        try {
            const config = this.m_config[parameterName];
            if (!config) return false;

            if (config.MinValue !== undefined && value < config.MinValue) {
                this._log('warn', `参数'${parameterName}'值${value}小于最小值${config.MinValue}`);
                return false;
            }
            if (config.MaxValue !== undefined && value > config.MaxValue) {
                this._log('warn', `参数'${parameterName}'值${value}大于最大值${config.MaxValue}`);
                return false;
            }
            if (!this.ValidateValueByRule(value, config.ValidationRule)) {
                this._log('warn', `参数'${parameterName}'值${value}不符合验证规则${config.ValidationRule}`);
                return false;
            }
            return true;
        } catch (error) {
            this._log('error', `参数值范围验证失败：${error.message}`);
            return false;
        }
    }

    ValidateParameter(parameterName) {
        try {
            const config = this.m_config[parameterName];
            if (!config) return { isValid: false, message: `参数不存在: ${parameterName}`, value: null };

            const value = this.ReadParameterValue(parameterName);

            if (config.IsRequired && (value === null || value === undefined || value === "")) {
                return { isValid: false, message: `必需参数未填写: ${config.DisplayName}`, value };
            }
            if (!this.ValidateValueByRule(value, config.ValidationRule)) {
                return { isValid: false, message: `参数值不符合验证规则: ${config.DisplayName}`, value };
            }
            return { isValid: true, message: "参数验证通过", value };
        } catch (error) {
            return { isValid: false, message: `验证过程中发生错误: ${error.message}`, value: null };
        }
    }

    ValidateAllParameters() {
        const results = { total: 0, valid: 0, invalid: 0, details: [] };
        for (const paramName of this.GetParameterNames()) {
            const validation = this.ValidateParameter(paramName);
            results.total++;
            if (validation.isValid) results.valid++; else results.invalid++;
            results.details.push({ parameter: paramName, displayName: this.m_config[paramName].DisplayName, ...validation });
        }
        return results;
    }

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
        return { isValid: missingParams.length === 0, missingParams };
    }

    // ==================== §9 动态属性生成 ====================

    _createParameterAccessors() {
        const paramNames = this.GetParameterNames();
        paramNames.forEach(name => {
            Object.defineProperty(this, `${name}CellR1C1`, {
                get: () => this.GetConfigValue(name, "CellAddress"),
                configurable: true
            });

            Object.defineProperty(this, `${name}CellA1`, {
                get: () => this.GetCellAddressA1(name),
                configurable: true
            });

            Object.defineProperty(this, `${name}CellValue`, {
                get: () => this.ReadParameterValue(name),
                configurable: true
            });

            // 保持向后兼容：PaymentsPerYear 的特殊处理
            if (name === "PaymentsPerYear") {
                Object.defineProperty(this, "PaymentsPerYearValue", {
                    get: () => this.ReadParameterValue(name),
                    configurable: true
                });
            }
        });
    }

    // ==================== §9 快捷访问 ====================

    getParam(name, type = "value") {
        const types = {
            value: () => this.ReadParameterValue(name),
            cellA1: () => this.GetCellAddressA1(name),
            cellR1C1: () => this.GetConfigValue(name, "CellAddress"),
            config: () => this.m_config[name]
        };
        return (types[type] || types.value)();
    }

    val(name) {
        const result = this.getParam(name, "value");
        if (isNaN(result) && result !== null && result !== undefined && result !== "") {
            this._log('debug', `val() 参数'${name}'返回NaN值: ${result}`);
        }
        return result;
    }
    addr(name, format = "A1") { return this.getParam(name, format === "A1" ? "cellA1" : "cellR1C1"); }
    param(name) { return new clsParamAccessor(this, name); }

    /**
     * GetParameterNames - 获取所有参数名称
     * m_config 在构造函数中即已创建，不依赖 _isInitialized。
     */
    GetParameterNames() { return Object.keys(this.m_config); }

    /**
     * GetParameterConfig - 获取参数配置
     */
    GetParameterConfig(parameterName) { return this.m_config[parameterName] ? this.m_config[parameterName] : null; }

    // ==================== §11 调试与报告 ====================

    GetAllParametersSummary() {
        if (!this._isInitialized) return "参数管理器未初始化";

        var summary = "=== 所有参数配置摘要 ===\n";
        const paramNames = this.GetParameterNames();
        summary += `参数总数: ${paramNames.length}\n\n`;

        for (const paramName of paramNames) {
            const config = this.m_config[paramName];
            summary += `【${config.DisplayName}】\n`;
            summary += `  参数名称: ${paramName}\n`;
            summary += `  单元格地址: ${config.CellAddress}\n`;
            summary += `  默认值: ${config.DefaultValue}\n`;
            summary += `  数据类型: ${config.DataType}\n`;
            summary += `  是否必需: ${config.IsRequired ? "是" : "否"}\n`;
            summary += `  验证规则: ${config.ValidationRule}\n`;
            summary += `  描述: ${config.Description}\n`;
            if (config.DefaultFormula) summary += `  默认公式: ${config.DefaultFormula}\n`;
            summary += "\n";
        }
        return summary;
    }

    PrintDebugInfo() {
        this._log('info', '=== clsParameterManager 调试信息 ===');
        this._log('info', `初始化状态: ${this._isInitialized}`);
        const sheetName = (this.m_worksheet && this.m_worksheet.Name) ? this.m_worksheet.Name : "未设置";
        this._log('info', `工作表: ${sheetName}`);
        this._log('info', `参数总数: ${this.GetParameterNames().length}`);

        const validation = this.ValidateAllParameters();
        this._log('info', `参数验证: ${validation.valid}/${validation.total} 通过`);

        if (validation.invalid > 0) {
            this._log('warn', '未通过验证的参数:');
            validation.details.filter(d => !d.isValid).forEach(d => {
                this._log('warn', `  - ${d.displayName}: ${d.message}`);
            });
        }
    }

    GenerateParameterReport() {
        var report = "=== 参数管理器报告 ===\n";
        report += `生成时间: ${new Date().toLocaleString()}\n`;
        const reportSheetName = (this.m_worksheet && this.m_worksheet.Name) ? this.m_worksheet.Name : "未设置";
        report += `工作表: ${reportSheetName}\n\n`;

        const paramNames = this.GetParameterNames();
        report += `参数概览 (${paramNames.length}个参数):\n`;

        for (const paramName of paramNames) {
            const config = this.m_config[paramName];
            const value = this.ReadParameterValue(paramName);
            const validation = this.ValidateParameter(paramName);

            report += `【${config.DisplayName}】\n`;
            report += `  值: ${value}\n`;
            report += `  状态: ${validation.isValid ? "✓ 有效" : "✗ 无效"}\n`;
            if (!validation.isValid) report += `  错误: ${validation.message}\n`;
            report += `  位置: ${this.GetCellAddressA1(paramName)}\n\n`;
        }
        return report;
    }

    // ==================== §10 公式与重置 ====================

    ResetToDefault(parameterName) {
        var successCount = 0, failureCount = 0;
        try {
            const paramsToReset = parameterName ? [parameterName] : this.GetParameterNames();
            for (const paramName of paramsToReset) {
                const config = this.m_config[paramName];
                if (config) {
                    this.SetParameterValue(paramName, config.DefaultValue) ? successCount++ : failureCount++;
                }
            }
        } catch (error) { this._log('error', `重置参数失败: ${error.message}`); }
        return { successCount, failureCount };
    }

    ApplyDefaultFormulas() {
        var successCount = 0, failureCount = 0;
        try {
            for (const paramName of this.GetParameterNames()) {
                const config = this.m_config[paramName];
                if (config.DefaultFormula) {
                    const cellAddressA1 = this.GetCellAddressA1(paramName);
                    cellAddressA1 ? (this.m_worksheet.Range(cellAddressA1).Formula = config.DefaultFormula, successCount++) : failureCount++;
                }
            }
        } catch (error) { this._log('error', `应用默认公式失败: ${error.message}`); }
        return { successCount, failureCount };
    }

    // ==================== §13 导出 ====================

    ExportConfigToJson(filePath) {
        try {
            const configData = { exportTime: new Date().toISOString(), sheetName: this.m_worksheet.Name, parameters: this.m_config };
            this._log('info', `JSON配置内容: ${JSON.stringify(configData, null, 2)}`);
            this._log('info', `文件路径: ${filePath}`);
            return true;
        } catch (error) {
            this._log('error', `导出配置失败: ${error.message}`);
            return false;
        }
    }

    // 注意：ImportConfigFromJson 已在 V3.3 中移除。
    // 原因：WPS JSA 不支持 ActiveXObject 文件 I/O，该方法始终返回 false 属于死代码。
    // 如需导入配置，请手动编辑参数单元格或使用 ExportConfigToJson 查看当前配置。

    // ==================== §12 事件处理 ====================

    /**
     * SetParameterChangeListener - 设置参数变更监听器
     *
     * V3.3 改进：支持范围重叠检测。
     * 当用户编辑一个包含多个参数单元格的范围时，所有受影响的参数都会触发回调。
     *
     * @param {Function} callback - 回调函数 (parameterName, newValue, changedAddress)
     */
    SetParameterChangeListener(callback) {
        if (!this.m_worksheet) return;

        this.m_worksheet.Change = (changeRange) => {
            const changedAddress = changeRange.Address;

            // 预计算所有参数的 A1 地址，避免在循环中重复调用
            const paramNames = this.GetParameterNames();
            const paramAddresses = {};
            for (const paramName of paramNames) {
                const addr = this.GetCellAddressA1(paramName);
                if (addr) paramAddresses[paramName] = addr;
            }

            // 解析变更范围为行列范围
            const changedRange = this._parseRangeAddress(changedAddress);

            for (const paramName of paramNames) {
                const paramAddr = paramAddresses[paramName];
                if (!paramAddr) continue;

                // 精确匹配（最常见场景，性能最优）
                if (paramAddr === changedAddress) {
                    callback(paramName, this.ReadParameterValue(paramName), changedAddress);
                    continue;
                }

                // 范围重叠检测：当编辑范围覆盖多个参数单元格时
                if (changedRange) {
                    const paramRange = this._parseCellAddress(paramAddr);
                    if (paramRange && this._rangesOverlap(changedRange, paramRange)) {
                        callback(paramName, this.ReadParameterValue(paramName), changedAddress);
                    }
                }
            }
        };
    }

    /**
     * _parseRangeAddress - 解析范围地址（如 "A1:B10" 或 "$A$1:$B$10"）
     *
     * @param {string} address - 单元格范围地址
     * @returns {Object|null} { startRow, endRow, startCol, endCol } 或 null
     */
    _parseRangeAddress(address) {
        try {
            // 移除 $ 符号
            const clean = address.replace(/\$/g, '');

            // 检查是否为范围格式（含冒号）
            const colonPos = clean.indexOf(':');
            if (colonPos === -1) {
                // 单个单元格
                return this._parseCellAddress(clean);
            }

            const start = this._parseCellAddress(clean.substring(0, colonPos));
            const end = this._parseCellAddress(clean.substring(colonPos + 1));

            if (!start || !end) return null;

            return {
                startRow: Math.min(start.startRow, end.startRow),
                endRow: Math.max(start.endRow, end.endRow),
                startCol: Math.min(start.startCol, end.startCol),
                endCol: Math.max(start.endCol, end.endCol)
            };
        } catch (e) {
            return null;
        }
    }

    /**
     * _parseCellAddress - 解析单个单元格地址（如 "B4"）
     *
     * @param {string} address - 单元格地址
     * @returns {Object|null} { startRow, endRow, startCol, endCol }
     */
    _parseCellAddress(address) {
        try {
            const match = address.match(/^([A-Z]+)(\d+)$/);
            if (!match) return null;

            const colStr = match[1];
            const row = parseInt(match[2]);

            // 将列字母转为数字
            var col = 0;
            for (var i = 0; i < colStr.length; i++) {
                col = col * 26 + (colStr.charCodeAt(i) - 64);
            }

            return { startRow: row, endRow: row, startCol: col, endCol: col };
        } catch (e) {
            return null;
        }
    }

    /**
     * _rangesOverlap - 检测两个范围是否有重叠
     *
     * @param {Object} a - { startRow, endRow, startCol, endCol }
     * @param {Object} b - { startRow, endRow, startCol, endCol }
     * @returns {boolean}
     */
    _rangesOverlap(a, b) {
        return !(a.endRow < b.startRow || a.startRow > b.endRow ||
                 a.endCol < b.startCol || a.startCol > b.endCol);
    }

    // ==================== 类变量访问器 ====================

    get RowStartValue() { return this.m_RowStart; }
    set RowStartValue(value) { this.m_RowStart = value; }

    // ==================== §14 静态工厂 ====================

    /**
     * Create - 创建并初始化参数管理器实例
     *
     * @param {string} sheetName - 工作表名称（默认 "1租金测算表V1"）
     * @param {Object} [config] - 可选自定义配置，合并到默认配置中
     * @returns {clsParameterManager} 已初始化的参数管理器实例
     */
    static Create(sheetName, config) {
        const manager = new clsParameterManager();
        if (config) {
            Object.keys(config).forEach(key => {
                if (manager.m_config[key]) {
                    Object.assign(manager.m_config[key], config[key]);
                }
            });
        }
        manager.Initialize(sheetName || "1租金测算表V1");
        return manager;
    }
}