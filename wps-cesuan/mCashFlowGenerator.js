/**
 * ============== 现金流量表生成器类模块 ==============
 * 作者：徐晓冬
 * 版本：V3.20260504
 * 描述：重构版现金流量表生成器 - 配置驱动架构
 *
 * V3 重构说明：
 * - [重构] CASHFLOW_CONFIG 移入类内为静态属性，不再污染全局命名空间
 * - [重构] Initialize() 简化为单一职责：接收参数管理器
 * - [重构] 新增 ReInitialize(sheetName) 处理工作表切换
 * - [移除] IsWPS()/SetCompatibilityMode() 移至独立的环境兼容模块
 * - [重构] GenerateRemark1/2 使用 CF_COL 常量替代硬编码索引
 * - [重构] 使用 mShared_constants 中的工具函数
 * - [重构] transposeArray 替换为 JSA880.z转置（如果可用）
 * - [统一] 方法命名统一为英文（内部工具函数保留中文命名）
 * ====================================================
 */

// ============== 静态配置对象 ==============
/**
 * CASHFLOW_CONFIG - 现金流表系统配置对象（类静态属性）
 *
 * 作用：集中管理所有配置数据，实现配置与业务逻辑分离
 * 设计：修改配置无需改动业务逻辑代码
 * 注意：V3版本移入类内，此处保留为向后兼容引用
 */
var CASHFLOW_CONFIG = {
    // 表头数组：所有列的表头文本
    headers: [
        "期次", "日期", "净现金流1（1+2+3+4+5）", "净现金流1-备注",
        "净现金流2（1+2+3+4）", "净现金流2-备注", "（1）电汇放款",
        "（2)租金偿付", "（3）保证金", "（4）名义货价", "（5）经纪人费用"
    ],

    // 经纪人费用支付方式配置
    brokerPaymentMethods: {
        "一次性支付-放款时": {
            paymentPeriod: 0,
            description: "在放款时一次性支付全部经纪人费用"
        },
        "一次性支付-第1期租金": {
            paymentPeriod: 1,
            description: "在第1期租金支付时一次性支付全部经纪人费用"
        },
        "分三次支付（第1\\中\\末期）": {
            paymentPeriod: 3,
            description: "分三次支付经纪人费用：第1期、中期、末期各支付1/3"
        }
    },

    // 行类型配置：定义不同期次的现金流行为
    rowTypes: {
        FIRST_ROW: 0,      // 首行（放款期）
        MIDDLE_ROWS: 1,    // 中间行（租金支付期）
        LAST_ROW: 2        // 末行（最后一期）
    }
};

// ============== 现金流量表生成器类 ==============
if (typeof clsCashFlowGenerator === 'undefined') {
    clsCashFlowGenerator = class {
        /**
         * 构造函数
         *
         * @param {Object} parameterManager - 参数管理器实例（可选）
         */
        constructor(parameterManager) {
            this.MODULE_NAME = "CashFlowGenerator";
            this.ModuleModifyDate = (typeof VERSION_DATE !== 'undefined') ? VERSION_DATE : "20260504";
            console.log(`[${this.MODULE_NAME}] 类实例创建 - 版本 3.${this.ModuleModifyDate}`);

            // 依赖注入
            this.p = parameterManager || null;

            // 状态属性
            this.pIsInitialized = false;

            // 性能计时器
            this.pPerformanceTimer = {
                startTime: 0,
                endTime: 0,
                duration: 0,
                operation: ""
            };

            // 错误码常量（与 mErrorHandler.ERROR_CODES 保持一致）
            this.cfValidationError = 3002;  // DATA_VALIDATION_ERROR
            this.cfCalculationError = 4000;  // CALCULATION_ERROR
            this.cfDataMissingError = 3000;  // DATA_NOT_FOUND
            this.cfFormulaError = 4002;      // INVALID_FORMULA
            this.cfInitializationError = 1001; // INITIALIZATION_ERROR
            this.cfCompatibilityError = 1002;   // CONFIGURATION_ERROR
        }

        // ============== 属性访问器 ==============

        /** 获取初始化状态 */
        get IsInitialized() {
            return this.pIsInitialized;
        }

        /** 获取总期数 */
        get TotalPeriods() {
            return this.p ? this.p.val("TotalPeriods") : null;
        }

        /** 获取现金流表起始行 */
        get CashFlowStartRow() {
            return this.p ? this.p.CashFlowTablerowStart : null;
        }

        // ============== 初始化方法 ==============

        /**
         * Initialize - 初始化方法
         *
         * 作用：注入参数管理器，完成依赖注入
         * 设计：单一职责，只负责注入，不负责工作表切换
         *
         * @param {Object} parameterManager - 参数管理器实例
         * @returns {boolean} 是否初始化成功
         */
        Initialize(parameterManager) {
            try {
                if (!parameterManager || typeof parameterManager !== 'object') {
                    throw new Error(
                        "clsCashFlowGenerator.Initialize() 需要传入参数管理器实例。" +
                        "请通过构造函数或 Initialize(paramManager) 注入。"
                    );
                }

                this.p = parameterManager;

                if (!this.p.IsInitialized) {
                    throw new Error("参数管理器初始化失败");
                }

                this.pIsInitialized = true;

                var totalPeriods = this.p.val("TotalPeriods");
                var principal = this.p.val("Principal");
                console.log(`[${this.MODULE_NAME}] 初始化完成 - 总期数:${totalPeriods}, 租赁成本:${principal}`);
                console.log("------------------------");
                return true;
            } catch (error) {
                this.pIsInitialized = false;
                var errMsg = `初始化失败：${error.message}`;
                console.log("------------------------");
                this._handleError("Initialize", error.message, this.cfInitializationError);
                return false;
            }
        }

        /**
         * ReInitialize - 重新初始化（切换工作表时使用）
         *
         * @param {string} sheetName - 工作表名称
         * @returns {boolean} 是否重新初始化成功
         */
        ReInitialize(sheetName) {
            try {
                if (!this.p) {
                    throw new Error("参数管理器未注入，请先调用 Initialize()");
                }
                this.p.Initialize(sheetName);
                this.pIsInitialized = this.p.IsInitialized;
                console.log(`[${this.MODULE_NAME}] 工作表切换完成: ${sheetName}`);
                return this.pIsInitialized;
            } catch (error) {
                this._handleError("ReInitialize", error.message, this.cfInitializationError);
                return false;
            }
        }

        // ============== 主流程方法 ==============

        /**
         * GenerateCashFlowTable - 生成完整的现金流量表
         *
         * @returns {boolean} 是否生成成功
         */
        GenerateCashFlowTable() {
            try {
                if (!this.pIsInitialized) {
                    throw new Error("生成器未初始化");
                }

                this._startPerformanceTimer("现金流量表生成");

                if (!this.ValidateParameters()) return false;
                if (!this.CreateCashFlowHeaders()) return false;
                if (!this.GenerateCashFlowData()) return false;
                if (!this.ProcessBrokerFees()) return false;
                if (!this.GenerateRemarks()) return false;
                if (!this.ApplyFormatting()) return false;

                this._endPerformanceTimer();
                console.log(`[${this.MODULE_NAME}] 现金流量表生成完成，耗时: ${this.pPerformanceTimer.duration.toFixed(3)} 秒`);
                return true;

            } catch (error) {
                this._handleError("GenerateCashFlowTable", error.message, this.cfCalculationError);
                return false;
            }
        }

        // ============== 参数验证 ==============

        /**
         * ValidateParameters - 验证参数
         *
         * @returns {boolean} 是否验证通过
         */
        ValidateParameters() {
            try {
                if (!this.pIsInitialized) {
                    throw new Error("生成器未初始化");
                }

                var totalPeriods = this.p.val("TotalPeriods");
                if (!totalPeriods || totalPeriods <= 0) {
                    throw new Error(`无效的总期数: ${totalPeriods}`);
                }

                var principal = this.p.val("Principal");
                if (!principal || principal <= 0) {
                    throw new Error("租赁成本必须大于0");
                }

                if (!this.p.m_worksheet) {
                    throw new Error("工作表对象无效");
                }

                var requiredParams = [
                    "LeaseStartDate", "Principal", "Deposit",
                    "NominalPrice", "BrokerPaymentMethod", "BrokerFeeRate"
                ];

                for (var paramName of requiredParams) {
                    var value = this.p.val(paramName);
                    if (value === null || value === undefined || value === "") {
                        throw new Error(`参数 ${paramName} 未设置或无效`);
                    }
                }

                console.log(`[${this.MODULE_NAME}] 参数验证通过`);
                return true;
            } catch (error) {
                this._handleError("ValidateParameters", error.message, this.cfValidationError);
                return false;
            }
        }

        // ============== 表头生成 ==============

        /**
         * CreateCashFlowHeaders - 创建现金流量表表头
         *
         * @returns {boolean} 是否创建成功
         */
        CreateCashFlowHeaders() {
            try {
                console.log(`[${this.MODULE_NAME}] CreateCashFlowHeaders - 开始执行`);

                var worksheet = this.p.m_worksheet;
                console.log(`[${this.MODULE_NAME}] 工作表对象: ${worksheet ? worksheet.Name : 'null'}`);

                var cashFlowTablerowStart = this.p.CashFlowTablerowStart;
                console.log(`[${this.MODULE_NAME}] 现金流表起始行: ${cashFlowTablerowStart}`);

                var headers = this._getHeaders();
                console.log(`[${this.MODULE_NAME}] 表头数量: ${headers.length}`);

                // 设置总标题
                var titleRow = cashFlowTablerowStart - 2;
                console.log(`[${this.MODULE_NAME}] 标题行: ${titleRow}`);
                var titleCell = worksheet.Range(`A${titleRow}`);
                titleCell.Value2 = "现金流及综合利率测算";
                titleCell.Interior.Color = COLORS.WHITE;
                titleCell.Font.Name = FONT_DEFAULT;
                titleCell.Font.Size = FONT_SIZE_TITLE;
                titleCell.Font.Color = COLORS.BLACK;
                titleCell.HorizontalAlignment = XL.HCenter;

                // 设置表头
                var headerRange = worksheet.Range(
                    worksheet.Cells(cashFlowTablerowStart - 1, 1),
                    worksheet.Cells(cashFlowTablerowStart - 1, headers.length)
                );

                var headerArray = [headers];
                headerRange.Value2 = headerArray;

                // 设置表头样式
                headerRange.Interior.Color = COLORS.HEADER_BLUE;
                headerRange.Font.Name = FONT_DEFAULT;
                headerRange.Font.Size = FONT_SIZE_HEADER;
                headerRange.Font.Color = COLORS.BLACK;
                headerRange.HorizontalAlignment = XL.HCenter;
                headerRange.VerticalAlignment = XL.VCenter;
                headerRange.WrapText = true;

                this.addBorder(headerRange);

                console.log(`[${this.MODULE_NAME}] 表头创建完成`);
                return true;
            } catch (error) {
                this._handleError("CreateCashFlowHeaders", error.message, this.cfCalculationError);
                return false;
            }
        }

        /**
         * _getHeaders - 获取表头数组
         *
         * @returns {Array} 表头数组
         */
        _getHeaders() {
            return CASHFLOW_CONFIG.headers;
        }

        // ============== 现金流数据生成 ==============

        /**
         * GenerateCashFlowData - 生成现金流数据
         *
         * @returns {boolean} 是否生成成功
         */
        GenerateCashFlowData() {
            try {
                var params = this._getFormulaParams();
                var totalPeriods = this.p.val("TotalPeriods");

                if (!totalPeriods || totalPeriods <= 0) {
                    throw new Error(`无效的总期数: ${totalPeriods}`);
                }

                var cashFlowArray = this._createCashFlowArray(totalPeriods);

                for (var i = 0; i <= totalPeriods; i++) {
                    this._generateRowData(cashFlowArray, i, params, totalPeriods);
                }

                this._writeCashFlowArray(cashFlowArray, params.cashFlowTablerowStart, totalPeriods);

                console.log(`[${this.MODULE_NAME}] 现金流数据生成完成，共${totalPeriods + 1}期`);
                return true;
            } catch (error) {
                this._handleError("GenerateCashFlowData", error.message, this.cfCalculationError);
                return false;
            }
        }

        /**
         * _getFormulaParams - 获取公式生成所需的参数
         *
         * @returns {Object} 参数对象
         */
        _getFormulaParams() {
            return {
                leaseStartDateCell: this.p.addr("LeaseStartDate"),
                principalCell: this.p.addr("Principal"),
                depositCell: this.p.addr("Deposit"),
                nominalPriceCell: this.p.addr("NominalPrice"),
                cashFlowTablerowStart: this.p.CashFlowTablerowStart,
                rentTableStartRow: this.p.RentTableStartRow
            };
        }

        /**
         * _createCashFlowArray - 创建现金流数据数组
         *
         * @param {number} totalPeriods - 总期数
         * @returns {Array} 空的现金流数组
         */
        _createCashFlowArray(totalPeriods) {
            var cashFlowArray = [];
            for (var i = 0; i <= totalPeriods; i++) {
                cashFlowArray[i] = new Array(11);
            }
            return cashFlowArray;
        }

        /**
         * _generateRowData - 生成单行现金流数据
         *
         * @param {Array} cashFlowArray - 现金流数组
         * @param {number} rowIndex - 行索引
         * @param {Object} params - 参数对象
         * @param {number} totalPeriods - 总期数
         */
        _generateRowData(cashFlowArray, rowIndex, params, totalPeriods) {
            var leaseStartDateCell = params.leaseStartDateCell;
            var principalCell = params.principalCell;
            var depositCell = params.depositCell;
            var nominalPriceCell = params.nominalPriceCell;
            var rentTableStartRow = params.rentTableStartRow;

            var rowType = this._getRowType(rowIndex, totalPeriods);

            // 期次 (CF_COL.PERIOD = 0)
            cashFlowArray[rowIndex][CF_COL.PERIOD] = rowIndex;

            // 日期 (CF_COL.DATE = 1)
            if (rowIndex === 0) {
                cashFlowArray[rowIndex][CF_COL.DATE] = `=${leaseStartDateCell}`;
            } else {
                cashFlowArray[rowIndex][CF_COL.DATE] = `=B${rentTableStartRow + rowIndex - 1}`;
            }

            // 净现金流公式
            // CF_COL.NET_CASHFLOW_1 = 2, CF_COL.NET_CASHFLOW_2 = 4
            cashFlowArray[rowIndex][CF_COL.NET_CASHFLOW_1] = "=SUM(RC[4]:RC[8])";
            // 净现金流2 = 电汇放款+租金偿付+保证金+名义货价，不含备注列(F)
            cashFlowArray[rowIndex][CF_COL.NET_CASHFLOW_2] = "=SUM(RC[2]:RC[5])";

            // 根据行类型生成现金流项目
            switch (rowType) {
                case CASHFLOW_CONFIG.rowTypes.FIRST_ROW:
                    // 首行：放款 + 保证金收取
                    // CF_COL.WIRE_TRANSFER = 6, CF_COL.DEPOSIT = 8
                    cashFlowArray[rowIndex][CF_COL.WIRE_TRANSFER] = `=-${principalCell}`;
                    cashFlowArray[rowIndex][CF_COL.DEPOSIT] = `=${depositCell}`;
                    break;

                case CASHFLOW_CONFIG.rowTypes.MIDDLE_ROWS:
                    // 中间行：租金偿付
                    // CF_COL.RENT_PAYMENT = 7
                    cashFlowArray[rowIndex][CF_COL.RENT_PAYMENT] = `=C${rentTableStartRow + rowIndex - 1}`;
                    break;

                case CASHFLOW_CONFIG.rowTypes.LAST_ROW:
                    // 末行：租金偿付 + 保证金退还 + 名义货价
                    cashFlowArray[rowIndex][CF_COL.RENT_PAYMENT] = `=C${rentTableStartRow + rowIndex - 1}`;
                    cashFlowArray[rowIndex][CF_COL.DEPOSIT] = `=-${depositCell}`;
                    // CF_COL.NOMINAL_PRICE = 9
                    cashFlowArray[rowIndex][CF_COL.NOMINAL_PRICE] = `=${nominalPriceCell}`;
                    break;
            }
        }

        /**
         * _getRowType - 获取行类型
         *
         * @param {number} rowIndex - 行索引
         * @param {number} totalPeriods - 总期数
         * @returns {number} 行类型
         */
        _getRowType(rowIndex, totalPeriods) {
            if (rowIndex === 0) {
                return CASHFLOW_CONFIG.rowTypes.FIRST_ROW;
            } else if (rowIndex === totalPeriods) {
                return CASHFLOW_CONFIG.rowTypes.LAST_ROW;
            } else {
                return CASHFLOW_CONFIG.rowTypes.MIDDLE_ROWS;
            }
        }

        /**
         * _writeCashFlowArray - 写入现金流数组到工作表
         *
         * @param {Array} cashFlowArray - 现金流数组
         * @param {number} startRow - 起始行号
         * @param {number} totalPeriods - 总期数
         * @returns {Range} 写入的数据范围
         */
        _writeCashFlowArray(cashFlowArray, startRow, totalPeriods) {
            var worksheet = this.p.m_worksheet;
            var targetRange = worksheet.Range(
                `A${startRow}`
            ).Resize(totalPeriods + 1, 11);
            targetRange.Value2 = cashFlowArray;
            return targetRange;
        }

        // ============== 经纪人费用处理 ==============

        /**
         * ProcessBrokerFees - 处理经纪人费用
         *
         * @returns {boolean} 是否处理成功
         */
        ProcessBrokerFees() {
            try {
                var brokerPaymentMethod = this.p.val("BrokerPaymentMethod");
                var cashFlowTablerowStart = this.p.CashFlowTablerowStart;
                var totalPeriodsValue = this.p.val("TotalPeriods");
                var worksheet = this.p.m_worksheet;

                var principalCell = this.p.addr("Principal");
                var brokerFeeRateCell = this.p.addr("BrokerFeeRate");

                var brokerFeeRange = worksheet.Range(
                    `K${cashFlowTablerowStart}:K${cashFlowTablerowStart + totalPeriodsValue}`
                );

                switch (brokerPaymentMethod) {
                    case "一次性支付-放款时":
                        worksheet.Range(`K${cashFlowTablerowStart}`).Formula =
                            `=-${principalCell}*${brokerFeeRateCell}`;
                        break;
                    case "一次性支付-第1期租金":
                        worksheet.Range(`K${cashFlowTablerowStart + 1}`).Formula =
                            `=-${principalCell}*${brokerFeeRateCell}`;
                        break;
                    case "分三次支付（第1\\中\\末期）":
                        var midPeriod = cashFlowTablerowStart + 1 + Math.round(totalPeriodsValue / 2);
                        worksheet.Range(`K${cashFlowTablerowStart + 1}`).Formula =
                            `=-${principalCell}*${brokerFeeRateCell}/3`;
                        worksheet.Range(`K${midPeriod}`).Formula =
                            `=-${principalCell}*${brokerFeeRateCell}/3`;
                        worksheet.Range(`K${cashFlowTablerowStart + totalPeriodsValue}`).Formula =
                            `=-${principalCell}*${brokerFeeRateCell}/3`;
                        break;
                }

                应用格式(brokerFeeRange, "Standard");

                console.log(`[${this.MODULE_NAME}] 经纪人费用处理完成，支付方式: ${brokerPaymentMethod}`);
                return true;
            } catch (error) {
                this._handleError("ProcessBrokerFees", error.message, this.cfCalculationError);
                return false;
            }
        }

        // ============== 备注生成 ==============

        /**
         * GenerateRemarks - 生成备注
         *
         * @returns {boolean} 是否生成成功
         */
        GenerateRemarks() {
            try {
                var cashFlowTablerowStart = this.p.CashFlowTablerowStart;
                var totalPeriodsValue = this.p.val("TotalPeriods");
                var worksheet = this.p.m_worksheet;

                // 读取现金流数据
                var cashFlowData = worksheet.Range(
                    `G${cashFlowTablerowStart}:K${cashFlowTablerowStart + totalPeriodsValue}`
                ).Value2;

                // 直接构建二维列数组 [[""],[""],...]，避免 JSA.z转置 把字符串拆成字符
                var remarkCol1 = [];
                var remarkCol2 = [];

                for (var i = 0; i <= totalPeriodsValue; i++) {
                    remarkCol1[i] = [this._generateRemark1(cashFlowData, i + 1)];
                    remarkCol2[i] = [this._generateRemark2(cashFlowData, i + 1)];
                }

                worksheet.Range(`D${cashFlowTablerowStart}`).Resize(totalPeriodsValue + 1, 1).Value2 = remarkCol1;
                worksheet.Range(`F${cashFlowTablerowStart}`).Resize(totalPeriodsValue + 1, 1).Value2 = remarkCol2;

                console.log(`[${this.MODULE_NAME}] 备注生成完成`);
                return true;
            } catch (error) {
                this._handleError("GenerateRemarks", error.message, this.cfCalculationError);
                return false;
            }
        }

        /**
         * _generateRemark1 - 生成净现金流1备注
         *
         * 使用 CF_COL 常量替代硬编码索引
         *
         * @param {Array} cashFlowData - 现金流数据数组
         * @param {number} rowIndex - 行索引（1-based）
         * @returns {string} 备注文本
         */
        _generateRemark1(cashFlowData, rowIndex) {
            var remark = "";
            var row = cashFlowData[rowIndex - 1] || [];

            // CF_COL 索引: WIRE_TRANSFER=6, RENT_PAYMENT=7, DEPOSIT=8,
            //              NOMINAL_PRICE=9, BROKER_FEE=10
            // cashFlowData 是 G:K 即列6-10，对应数组索引 0-4
            var wireTransfer = row[CF_COL.WIRE_TRANSFER - CF_COL.WIRE_TRANSFER]; // 0
            var rentPayment = row[CF_COL.RENT_PAYMENT - CF_COL.WIRE_TRANSFER];   // 1
            var deposit = row[CF_COL.DEPOSIT - CF_COL.WIRE_TRANSFER];           // 2
            var nominalPrice = row[CF_COL.NOMINAL_PRICE - CF_COL.WIRE_TRANSFER]; // 3
            var brokerFee = row[CF_COL.BROKER_FEE - CF_COL.WIRE_TRANSFER];       // 4

            if (wireTransfer !== undefined && wireTransfer !== 0) remark += "电汇放款/";
            if (rentPayment !== undefined && rentPayment !== 0) remark += `第${rowIndex - 1}期租金/`;
            if (deposit > 0) {
                remark += "出租人收取保证金/";
            } else if (deposit < 0) {
                remark += "出租人退还保证金/";
            }
            if (nominalPrice !== undefined && nominalPrice !== 0) remark += "出租人收取名义货价/";
            if (brokerFee !== undefined && brokerFee !== 0) remark += "经纪人费用/";

            return remark;
        }

        /**
         * _generateRemark2 - 生成净现金流2备注
         *
         * @param {Array} cashFlowData - 现金流数据数组
         * @param {number} rowIndex - 行索引（1-based）
         * @returns {string} 备注文本
         */
        _generateRemark2(cashFlowData, rowIndex) {
            var remark = "";
            var row = cashFlowData[rowIndex - 1] || [];

            var wireTransfer = row[CF_COL.WIRE_TRANSFER - CF_COL.WIRE_TRANSFER];
            var rentPayment = row[CF_COL.RENT_PAYMENT - CF_COL.WIRE_TRANSFER];
            var deposit = row[CF_COL.DEPOSIT - CF_COL.WIRE_TRANSFER];
            var nominalPrice = row[CF_COL.NOMINAL_PRICE - CF_COL.WIRE_TRANSFER];

            if (wireTransfer !== undefined && wireTransfer !== 0) remark += "电汇放款/";
            if (rentPayment !== undefined && rentPayment !== 0) remark += `第${rowIndex - 1}期租金/`;
            if (deposit > 0) {
                remark += "承租人支付保证金/";
            } else if (deposit < 0) {
                remark += "承租人收回保证金/";
            }
            if (nominalPrice !== undefined && nominalPrice !== 0) remark += "出租人名义货价收取/";

            return remark;
        }

        /**
         * _transposeArray - 转置数组（内部使用，当 JSA880 不可用时）
         *
         * @param {Array} arr - 要转置的数组
         * @returns {Array} 转置后的数组
         */
        _transposeArray(arr) {
            if (!Array.isArray(arr)) return arr;
            return arr.map(function(item) { return [item]; });
        }

        // ============== 格式应用 ==============

        /**
         * ApplyFormatting - 应用格式
         *
         * @returns {boolean} 是否应用成功
         */
        ApplyFormatting() {
            try {
                var cashFlowTablerowStart = this.p.CashFlowTablerowStart;
                var totalPeriodsValue = this.p.val("TotalPeriods");
                var worksheet = this.p.m_worksheet;

                // 数据区域
                var dataRange = worksheet.Range(
                    `A${cashFlowTablerowStart}:K${cashFlowTablerowStart + totalPeriodsValue}`
                );

                设置表格样式(dataRange);

                // 日期格式
                var dateRange = worksheet.Range(
                    `B${cashFlowTablerowStart}:B${cashFlowTablerowStart + totalPeriodsValue}`
                );
                应用格式(dateRange, "Date");
                设置表格样式(dateRange);

                // 数字格式
                var numberRange = worksheet.Range(
                    `C${cashFlowTablerowStart}:K${cashFlowTablerowStart + totalPeriodsValue}`
                );
                应用格式(numberRange, "Standard");
                设置表格样式(numberRange);

                // 文本格式（备注列）
                var textRange = worksheet.Range(
                    `D${cashFlowTablerowStart}:D${cashFlowTablerowStart + totalPeriodsValue},` +
                    `F${cashFlowTablerowStart}:F${cashFlowTablerowStart + totalPeriodsValue}`
                );
                应用格式(textRange, "Standard");
                设置表格样式(textRange);

                // 全区域边框
                var allRange = worksheet.Range(
                    `A${cashFlowTablerowStart - 1}:K${cashFlowTablerowStart + totalPeriodsValue}`
                );
                this.addBorder(allRange);

                console.log(`[${this.MODULE_NAME}] 格式应用完成`);
                return true;
            } catch (error) {
                this._handleError("ApplyFormatting", error.message, this.cfCalculationError);
                return false;
            }
        }

        // ============== 边框工具 ==============

        /**
         * addBorder - 添加单元格边框
         *
         * @param {Range} rng - 目标范围对象
         * @returns {boolean} 是否添加成功
         */
        addBorder(rng) {
            try {
                rng.Borders.LineStyle = XL.Continuous;
                rng.Borders.Color = COLORS.BLACK;
                rng.Borders.Weight = XL.Thin;
                rng.Borders.TintAndShade = 0;
                return true;
            } catch (error) {
                console.error(`添加框线失败：${error.message}`);
                return false;
            }
        }

        // ============== 综合利率一览 ==============

        /**
         * generateInterestRateOverview - 生成综合利率一览表
         *
         * @returns {boolean} 是否生成成功
         */
        generateInterestRateOverview() {
            try {
                var cashFlowTableRowStart = this.p.CashFlowTablerowStart;
                var totalPeriods = this.p.val("TotalPeriods");
                var paymentsPerYearA1 = this.p.addr("PaymentsPerYear", "A1");
                console.log(`[${this.MODULE_NAME}] 综合利率一览 - paymentsPerYearA1: ${paymentsPerYearA1}`);
                var worksheet = this.p.m_worksheet;

                var titleCell = worksheet.Range("A15");
                titleCell.Value2 = "综合利率一览";
                titleCell.Interior.Color = COLORS.WHITE;
                titleCell.Font.Name = FONT_DEFAULT;
                titleCell.Font.Size = FONT_SIZE_TITLE;
                titleCell.Font.Color = COLORS.BLACK;

                // XIRR 净内含报酬率
                var xirr1Cell = this.p.ConvertR1C1ToA1("R16C4");
                var xirr1Label = worksheet.Range(xirr1Cell).Offset(0, -1);
                xirr1Label.Value2 = "XIRR净内含报酬率";
                xirr1Label.Font.Name = FONT_DEFAULT;
                xirr1Label.Font.Size = FONT_SIZE_HEADER;
                xirr1Label.Font.Color = COLORS.BLACK;

                // 企业看XIRR
                var xirr2Cell = this.p.ConvertR1C1ToA1("R17C4");
                var xirr2Label = worksheet.Range(xirr2Cell).Offset(0, -1);
                xirr2Label.Value2 = "（1）企业看XIRR";
                xirr2Label.Font.Name = FONT_DEFAULT;
                xirr2Label.Font.Size = FONT_SIZE_HEADER;
                xirr2Label.Font.Color = COLORS.BLACK;

                // 经纪人费用影响
                var xirrDiffCell = this.p.ConvertR1C1ToA1("R18C4");
                var xirrDiffLabel = worksheet.Range(xirrDiffCell).Offset(0, -1);
                xirrDiffLabel.Value2 = "（2）经纪人费用影响";
                xirrDiffLabel.Font.Name = FONT_DEFAULT;
                xirrDiffLabel.Font.Size = FONT_SIZE_HEADER;
                xirrDiffLabel.Font.Color = COLORS.BLACK;

                // IRR 内含报酬率
                var irr1Cell = this.p.ConvertR1C1ToA1("R16C2");
                var irr1Label = worksheet.Range(irr1Cell).Offset(0, -1);
                irr1Label.Value2 = "IRR内含报酬率";
                irr1Label.Font.Name = FONT_DEFAULT;
                irr1Label.Font.Size = FONT_SIZE_HEADER;
                irr1Label.Font.Color = COLORS.BLACK;

                // (1)企业看IRR
                var irr2Cell = this.p.ConvertR1C1ToA1("R17C2");
                var irr2Label = worksheet.Range(irr2Cell).Offset(0, -1);
                irr2Label.Value2 = "(1)企业看IRR";
                irr2Label.Font.Name = FONT_DEFAULT;
                irr2Label.Font.Size = FONT_SIZE_HEADER;
                irr2Label.Font.Color = COLORS.BLACK;

                // XIRR 公式 (净现金流1)
                var formula = `=XIRR(C${cashFlowTableRowStart}:C${cashFlowTableRowStart + totalPeriods},` +
                             `B${cashFlowTableRowStart}:B${cashFlowTableRowStart + totalPeriods})`;
                var xirr1Range = worksheet.Range(xirr1Cell);
                xirr1Range.Formula = formula;
                xirr1Range.Font.Name = FONT_DEFAULT;
                xirr1Range.Font.Size = FONT_SIZE_HEADER;
                xirr1Range.Font.Color = COLORS.BLACK;
                xirr1Range.Interior.Color = COLORS.LIGHT_GREEN;

                // XIRR 公式 (净现金流2)
                formula = `=XIRR(E${cashFlowTableRowStart}:E${cashFlowTableRowStart + totalPeriods},` +
                         `B${cashFlowTableRowStart}:B${cashFlowTableRowStart + totalPeriods})`;
                var xirr2Range = worksheet.Range(xirr2Cell);
                xirr2Range.Formula = formula;
                xirr2Range.Font.Name = FONT_DEFAULT;
                xirr2Range.Font.Size = FONT_SIZE_HEADER;
                xirr2Range.Font.Color = COLORS.BLACK;
                xirr2Range.Interior.Color = COLORS.LIGHT_GREEN;

                // XIRR 差异公式
                formula = "=R[-2]C-R[-1]C";
                var xirrDiffRange = worksheet.Range(xirrDiffCell);
                xirrDiffRange.FormulaR1C1 = formula;
                xirrDiffRange.Font.Name = FONT_DEFAULT;
                xirrDiffRange.Font.Size = FONT_SIZE_HEADER;
                xirrDiffRange.Font.Color = COLORS.BLACK;
                xirrDiffRange.Interior.Color = COLORS.LIGHT_RED;

                // IRR 公式 (净现金流1)
                formula = `=IRR(C${cashFlowTableRowStart}:C${cashFlowTableRowStart + totalPeriods})*${paymentsPerYearA1}`;
                var irr1Range = worksheet.Range(irr1Cell);
                irr1Range.Formula = formula;
                irr1Range.Font.Name = FONT_DEFAULT;
                irr1Range.Font.Size = FONT_SIZE_HEADER;
                irr1Range.Font.Color = COLORS.BLACK;
                irr1Range.Interior.Color = COLORS.LIGHT_GREEN;

                // IRR 公式 (净现金流2)
                formula = `=IRR(E${cashFlowTableRowStart}:E${cashFlowTableRowStart + totalPeriods})*${paymentsPerYearA1}`;
                var irr2Range = worksheet.Range(irr2Cell);
                irr2Range.Formula = formula;
                irr2Range.Font.Name = FONT_DEFAULT;
                irr2Range.Font.Size = FONT_SIZE_HEADER;
                irr2Range.Font.Color = COLORS.BLACK;
                irr2Range.Interior.Color = COLORS.LIGHT_GREEN;

                // 设置数字格式
                var formatRange = worksheet.Range(`${irr1Cell}:${irr2Cell},${xirr1Cell}:${xirrDiffCell}`);
                formatRange.NumberFormatLocal = "0.00%";

                console.log(`[${this.MODULE_NAME}] 综合利率一览计算完成`);
                return true;
            } catch (error) {
                console.error(`[${this.MODULE_NAME}] 综合利率一览计算失败：${error.message}`);
                return false;
            }
        }

        // ============== 性能计时 ==============

        /**
         * _startPerformanceTimer - 开始性能计时
         *
         * @param {string} operation - 操作名称
         */
        _startPerformanceTimer(operation) {
            this.pPerformanceTimer.startTime = new Date().getTime();
            this.pPerformanceTimer.operation = operation;
            console.log(`[${this.MODULE_NAME}] 开始: ${operation}`);
        }

        /**
         * _endPerformanceTimer - 结束性能计时
         */
        _endPerformanceTimer() {
            this.pPerformanceTimer.endTime = new Date().getTime();
            this.pPerformanceTimer.duration =
                (this.pPerformanceTimer.endTime - this.pPerformanceTimer.startTime) / 1000;
            console.log(`[${this.MODULE_NAME}] 完成: ${this.pPerformanceTimer.operation}，耗时: ${this.pPerformanceTimer.duration.toFixed(3)} 秒`);
        }

        // ============== 错误处理 ==============

        /**
         * _handleError - 统一错误处理
         *
         * @param {string} errSource - 错误来源
         * @param {string} errDescription - 错误描述
         * @param {number} defaultErrorCode - 默认错误码
         */
        _handleError(errSource, errDescription, defaultErrorCode) {
            var errorMsg = `[${errSource}] ${errDescription}`;
            console.error(`[${this.MODULE_NAME}] 错误: ${errorMsg}`);

            // 优先使用统一错误处理器
            if (typeof g_errorHandler !== 'undefined') {
                var error = new Error(errDescription);
                g_errorHandler.handleError(error, defaultErrorCode, {
                    module: this.MODULE_NAME,
                    function: errSource
                });
            }
        }
    }
}