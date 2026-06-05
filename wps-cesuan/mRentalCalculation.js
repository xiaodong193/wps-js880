/**
 * ============== 租金测算系统 V2.1（配置驱动架构） ==============
 * 
 * 版本更新：
 * - V2.1.20260430: P0-P2审计修复（写入范围截断、单期公式冲突、流水线回滚、死代码清理等）
 * - V2.1.20260409: P0-P2重构（拆分上帝文件、修复Bug、集成错误处理器）
 * - V2.1.20260130: 修复全局变量引用问题，增强错误处理
 * 作者：徐晓冬
 * 版本：V2.2026.1.51
 * 描述：完全重构版租金测算模块 - 配置驱动架构
 * 
 * 核心改进：
 * - 配置驱动：所有配置集中在 clsRentalConfig 中
 * - 单一职责：每个类只负责一个功能域
 * - 消除重复：公式生成逻辑统一，差异参数化
 * - 缓存优先：充分利用参数管理器的 _addressCache
 * - 完善注释：所有复杂逻辑都有设计意图说明
 * 
 * 架构：
 * - clsRentalConfig: 配置管理
 * - clsFormulaGenerator: 公式生成
 * - clsStyleManager: 样式管理
 * - clsRentalCalculation: 主类（协调其他类）
 * - clsRentalTableRowManager: 行管理（继承自主类，未来可改为组合模式避免继承链过深）
 * 
 * 依赖：
 * - mParameterManager.js (参数管理器 V3)
 * - mShared_constants.js (共享常量和工具函数)
 * ====================================================
 */

// ============== Array2D 优化器（延迟初始化） ==============
var _array2DOptimizer = null;

function getArray2DOptimizer(parameterManager) {
    if (!_array2DOptimizer) {
        _array2DOptimizer = new clsArray2DOptimizer(parameterManager);
    } else if (parameterManager && _array2DOptimizer.p !== parameterManager) {
        _array2DOptimizer.p = parameterManager;
    }
    return _array2DOptimizer;
}

// ============== clsRentalConfig（原 mRentalConfig.js） ==============
/**
 * clsRentalConfig - 租金测算配置管理类
 *
 * 作用：集中管理所有配置，实现配置与业务逻辑分离
 * 设计：修改配置无需改动业务逻辑代码
 */
class clsRentalConfig {
    constructor(parameterManager) {
        this.p = parameterManager;
        this.MODULE_NAME = "clsRentalConfig";

        this.columnDefinitions = this._initColumnDefinitions();
        this.repaymentMethods = this._initRepaymentMethods();
        this.totalRowConfig = this._initTotalRowConfig();

        console.log(`[${this.MODULE_NAME}] 配置管理器初始化完成`);
    }

    _initColumnDefinitions() {
        return {
            PERIOD: 'A',
            DATE: 'B',
            RENT: 'C',
            PRINCIPAL: 'D',
            INTEREST: 'E',
            RENT_BALANCE: 'F',
            PRINCIPAL_BALANCE: 'G',
            REMAINING_BALANCE: 'H',
            PAID_RENT: 'I',
            MONTH_INTERVAL: 'J',
            CUSTOM_INTERVAL: 'K',
            PRINCIPAL_RATIO: 'L',
            RATE_PER_PERIOD: 'M',
            ADJUSTMENT_REMARK: 'N'
        };
    }

    _initRepaymentMethods() {
        return {
            "等额本息（后付）": {
                formulaMethod: "generateEqualPaymentFormulas",
                usePrincipalRatio: false,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [],
                headerRange: [1, 10],
                addFrame: ["A:J"],
                extraColumns: []
            },
            "等额本息（先付）": {
                formulaMethod: "generateEqualPaymentAdvanceFormulas",
                usePrincipalRatio: false,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [],
                headerRange: [1, 10],
                addFrame: ["A:J"],
                extraColumns: []
            },
            "等额本金（按天计息）": {
                formulaMethod: "generateEqualPrincipalDailyInterestFormulas",
                usePrincipalRatio: false,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [12],
                headerRange: [1, 10],
                addFrame: ["A:J"],
                extraColumns: []
            },
            "等额本金（按期计息）": {
                formulaMethod: "generateEqualPrincipalPeriodicInterestFormulas",
                usePrincipalRatio: false,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [12],
                headerRange: [1, 10],
                addFrame: ["A:J"],
                extraColumns: []
            },
            "本金比例（按期计息）": {
                formulaMethod: "generatePrincipalRatioPeriodicInterestFormulas",
                usePrincipalRatio: true,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [],
                headerRange: [1, 10],
                addFrame: ["A:J", "L:L"],
                extraColumns: [12]
            },
            "本金比例（按天计息）": {
                formulaMethod: "generatePrincipalRatioDailyInterestFormulas",
                usePrincipalRatio: true,
                needCustomInterval: false,
                convertColumns: [1, 2],
                clearColumns: [],
                headerRange: [1, 10],
                addFrame: ["A:J", "L:L"],
                extraColumns: [12]
            }
        };
    }

    _initTotalRowConfig() {
        return {
            summaryColumns: [3, 4, 5, 10, 11, 12],
            dashColumns: [2, 6, 7, 8, 9],
            textColumns: [1]
        };
    }

    getRepaymentMethodConfig(methodName) {
        return this.repaymentMethods[methodName] || null;
    }

    getColumnLetter(columnName) {
        return this.columnDefinitions[columnName] || null;
    }
}

// ============== 公式生成器类 ==============
/**
 * clsFormulaGenerator - 公式生成器类
 *
 * 作用：根据还款方式生成对应的Excel/WPS公式（R1C1引用格式）
 * 设计：策略模式 — 每种还款方式对应一个生成方法
 */
class clsFormulaGenerator {
    constructor(parameterManager, config) {
        this.p = parameterManager;
        this.config = config || null;
        this.MODULE_NAME = "clsFormulaGenerator";

        // 利率引用模式: 'fixed' = 固定利率单元格(如R5C2), 'column' = M列每期适用利率(RC[8])
        this._rateReferenceMode = 'fixed';
        // 租前期配置: { enabled, preLeaseRowCount, firstPaymentDateRef, preLeaseMonthsRef }
        this._preLeaseConfig = null;

        console.log(`[${this.MODULE_NAME}] 公式生成器初始化完成`);
    }

    setRateReferenceMode(mode) {
        this._rateReferenceMode = mode === 'column' ? 'column' : 'fixed';
        console.log(`[${this.MODULE_NAME}] 利率引用模式: ${this._rateReferenceMode}`);
    }

    setPreLeaseConfig(config) {
        this._preLeaseConfig = config || null;
        console.log(`[${this.MODULE_NAME}] 租前期配置: ${this._preLeaseConfig ? '已设置' : '已清除'}`);
    }

    _getRateRef(params) {
        return this._rateReferenceMode === 'column' ? 'RC[8]' : params.interestRateCell;
    }

    _applyPreLeaseOverrides(arr, params) {
        if (!this._preLeaseConfig || !this._preLeaseConfig.enabled) return;

        const cfg = this._preLeaseConfig;
        // arr[1][2] 日期列: 第1期支付日 = 起租日 + 租前期月数
        arr[1][2] = `=EDATE(${cfg.firstPaymentDateRef},${cfg.preLeaseMonthsRef})`;
        // arr[1][5] 利息列: 日期引用从放款日改为上一行日期（仅在公式含放款日引用时生效）
        if (arr[1][5] && typeof arr[1][5] === 'string' && params.leaseStartDateCell) {
            arr[1][5] = arr[1][5].replace(params.leaseStartDateCell, 'R[-1]C[-3]');
        }
        // arr[1][10] 月间隔列
        arr[1][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
        // arr[3][2] 末期日期列: 从起租日开始计算
        arr[3][2] = `=EDATE(${cfg.firstPaymentDateRef},${params.paymentInterval}*${params.totalPeriodsValue})`;
    }

    _getFormulaParams() {
        return {
            leaseStartDateCell: this.p.addr("LeaseStartDate", "R1C1"),
            paymentInterval: this.p.val("PaymentInterval"),
            interestRateCell: this.p.addr("InterestRate", "R1C1"),
            paymentsPerYearCell: this.p.addr("PaymentsPerYear", "R1C1"),
            totalPeriodsCell: this.p.addr("TotalPeriods", "R1C1"),
            principalCell: this.p.addr("Principal", "R1C1"),
            totalPeriodsValue: this.p.val("TotalPeriods"),
            rentTableStartRow: this.p.RentTableStartRow
        };
    }

    _createFormulaArray() {
        return createFormulaTemplate(5, 13);
    }

    // ============== 公共公式模板 ==============

    _applyCommonFormulas(arr, params) {
        const { leaseStartDateCell, paymentInterval, principalCell,
                totalPeriodsValue, rentTableStartRow } = params;

        arr[1][1] = "1";
        arr[1][2] = `=EDATE(${leaseStartDateCell}, ${paymentInterval})`;
        arr[1][7] = `=${principalCell} - RC[-1]`;
        arr[1][8] = `=SUM(R${rentTableStartRow}C3:R${rentTableStartRow + totalPeriodsValue - 1}C3) - RC[1]`;
        arr[1][9] = "=RC[-6]";
        arr[1][10] = `=DATEDIF(${leaseStartDateCell},RC[-8], "M")`;

        arr[2][1] = "=R[-1]C+1";
        arr[2][2] = "=EDATE(R[-1]C, " + paymentInterval + ")";
        arr[2][7] = arr[1][7];
        arr[2][8] = arr[1][8];
        arr[2][9] = "=RC[-6] + R[-1]C";
        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';

        arr[3][1] = arr[2][1];
        arr[3][2] = "=EDATE(" + leaseStartDateCell + "," + paymentInterval + "*" + totalPeriodsValue + ")";
        arr[3][7] = arr[1][7];
        arr[3][8] = arr[1][8];
        arr[3][9] = "=RC[-6] + R[-1]C";
        arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
    }

    _applyCommonCumulativePrincipal(arr) {
        arr[2][6] = "=RC[-2] + R[-1]C";
        arr[3][6] = "=RC[-2] + R[-1]C";
    }

    _applyLastPeriodPrincipalFix(arr, params) {
        arr[3][4] = `=${params.principalCell} - SUM(R[${-params.totalPeriodsValue + 1}]C:R[-1]C)`;
    }

    // ============== 各还款方式的差异化公式 ==============

    generateEqualPaymentFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);
            this._applyCommonCumulativePrincipal(arr);

            arr[1][3] = `=ROUND(-PMT(${this._getRateRef(params)}/${params.paymentsPerYearCell},${params.totalPeriodsCell},${params.principalCell},0),2)`;
            arr[1][4] = `=ROUND(-PPMT(${this._getRateRef(params)}/${params.paymentsPerYearCell},RC[-3],${params.totalPeriodsCell},${params.principalCell},0),2)`;
            arr[1][5] = "=RC[-2]-RC[-1]";
            arr[1][6] = "=RC[-2]";

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = arr[1][5];

            arr[3][3] = arr[1][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[1][5];

            this._applyLastPeriodPrincipalFix(arr, params);

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '等额本息' });
            } else {
                console.error(`[${this.MODULE_NAME}] 等额本息公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generateEqualPaymentAdvanceFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);

            arr[1][3] = `=ROUND(-PPMT(${this._getRateRef(params)}/${params.paymentsPerYearCell},RC[-2],${params.totalPeriodsCell},${params.principalCell}+RC[2],,1),2)`;
            arr[1][4] = "=RC[-1]-RC[1]";
            arr[1][5] = `=ROUND(${params.principalCell}*${this._getRateRef(params)}*(RC[-3] - ${params.leaseStartDateCell})/360,2)`;
            arr[1][6] = "=RC[-2]";

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=R[-1]C[2]*${this._getRateRef(params)}/${params.paymentsPerYearCell}`;
            arr[2][6] = "=RC[-2] + R[-1]C";

            arr[3][3] = arr[2][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];
            arr[3][6] = arr[2][6];

            this._applyLastPeriodPrincipalFix(arr, params);

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '等额本息先付' });
            } else {
                console.error(`[${this.MODULE_NAME}] 等额本息先付公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generateEqualPrincipalDailyInterestFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);
            this._applyCommonCumulativePrincipal(arr);

            arr[1][3] = "=ROUND(RC[1]+RC[2],2)";
            arr[1][4] = `=ROUND(${params.principalCell}/${params.totalPeriodsCell},2)`;
            arr[1][5] = `=ROUND(${params.principalCell}*${this._getRateRef(params)}/360*(RC[-3]-${params.leaseStartDateCell}),2)`;
            arr[1][6] = "=RC[-2]";

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*${this._getRateRef(params)}/360*(RC[-3]-R[-1]C[-3]),2)`;

            arr[3][3] = arr[1][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];

            this._applyLastPeriodPrincipalFix(arr, params);

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '等额本金按天计息' });
            } else {
                console.error(`[${this.MODULE_NAME}] 等额本金按天计息公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generateEqualPrincipalPeriodicInterestFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);
            this._applyCommonCumulativePrincipal(arr);

            arr[1][3] = "=ROUND(RC[1]+RC[2],2)";
            arr[1][4] = `=ROUND(${params.principalCell}/${params.totalPeriodsCell},2)`;
            arr[1][5] = `=ROUND(${params.principalCell}*${this._getRateRef(params)}/${params.paymentsPerYearCell},2)`;
            arr[1][6] = "=RC[-2]";

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*${this._getRateRef(params)}/${params.paymentsPerYearCell},2)`;

            arr[3][3] = arr[1][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '等额本金按期计息' });
            } else {
                console.error(`[${this.MODULE_NAME}] 等额本金按期计息公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generatePrincipalRatioPeriodicInterestFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);
            this._applyCommonCumulativePrincipal(arr);

            arr[1][3] = "=ROUND(RC[1]+RC[2],2)";
            arr[1][4] = `=ROUND(${params.principalCell}*RC[8]/100,2)`;
            arr[1][5] = `=ROUND(${params.principalCell}*${this._getRateRef(params)}/${params.paymentsPerYearCell},2)`;
            arr[1][6] = "=RC[-2]";
            arr[1][12] = `=round(100/${params.totalPeriodsCell},2)`;

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*${this._getRateRef(params)}/${params.paymentsPerYearCell},2)`;
            arr[2][12] = arr[1][12];

            arr[3][3] = arr[2][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];
            arr[3][12] = `=100-SUM(R[${-params.totalPeriodsValue + 1}]C:R[-1]C)`;

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '本金比例按期计息' });
            } else {
                console.error(`[${this.MODULE_NAME}] 本金比例按期计息公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generatePrincipalRatioDailyInterestFormulas() {
        try {
            const params = this._getFormulaParams();
            const arr = this._createFormulaArray();

            this._applyCommonFormulas(arr, params);
            this._applyCommonCumulativePrincipal(arr);

            arr[1][3] = "=ROUND(RC[1]+RC[2],2)";
            arr[1][4] = `=ROUND(${params.principalCell}*RC[8]/100,2)`;
            arr[1][5] = `=ROUND(${params.principalCell}*${this._getRateRef(params)}/360*(RC[-3]-${params.leaseStartDateCell}),2)`;
            arr[1][6] = "=RC[-2]";
            arr[1][12] = `=round(100/${params.totalPeriodsCell},2)`;

            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*${this._getRateRef(params)}/360*(RC[-3]-R[-1]C[-3]),2)`;
            arr[2][12] = arr[1][12];

            arr[3][3] = arr[2][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];
            arr[3][12] = `=100-SUM(R[${-params.totalPeriodsValue + 1}]C:R[-1]C)`;

            return arr;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '本金比例按天计息' });
            } else {
                console.error(`[${this.MODULE_NAME}] 本金比例按天计息公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    generateFormulas(repaymentMethod) {
        try {
            console.log(`[${this.MODULE_NAME}] 公式分发器，偿还方式: ${repaymentMethod}`);

            const configInstance = this.config || new clsRentalConfig(this.p);
            const methodConfig = configInstance.getRepaymentMethodConfig(repaymentMethod);

            if (!methodConfig) {
                throw new Error(`不支持的还款方式：${repaymentMethod}`);
            }

            const formulaMethod = methodConfig.formulaMethod;
            console.log(`[${this.MODULE_NAME}] 调用公式生成方法: ${formulaMethod}`);

            if (typeof this[formulaMethod] === 'function') {
                const result = this[formulaMethod]();
                if (result) {
                    this._applyPreLeaseOverrides(result, this._getFormulaParams());
                }
                return result;
            } else {
                throw new Error(`公式生成方法不存在: ${formulaMethod}`);
            }
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: 'dispatchFormula' });
            } else {
                console.error(`[${this.MODULE_NAME}] 公式分发失败: ${error.message}`);
            }
            return null;
        }
    }

    适用每期利率(repaymentMethod) {
        try {
            console.log(`[${this.MODULE_NAME}] ========== 开始生成每期适用利率公式 ==========`);
            console.log(`[${this.MODULE_NAME}] 偿还方式: ${repaymentMethod}`);

            const interestRateCellA1 = this.p.addr("InterestRate", "A1");
            const interestRateCellR1C1 = this.p.addr("InterestRate", "R1C1");
            console.log(`[${this.MODULE_NAME}] 利率单元格地址(A1): ${interestRateCellA1}`);
            console.log(`[${this.MODULE_NAME}] 利率单元格地址(R1C1): ${interestRateCellR1C1}`);

            const arrFormula = this.generateFormulas(repaymentMethod);
            if (!arrFormula) {
                throw new Error("原始公式生成失败");
            }

            console.log(`[${this.MODULE_NAME}] 原始公式生成完成，开始替换利率引用...`);

            const ratePatterns = this._buildRatePatterns(interestRateCellA1, interestRateCellR1C1);

            const columnOffsetMap = {
                3: 'RC[10]',
                4: 'RC[9]',
                5: 'RC[8]'
            };

            var totalReplacements = 0;
            for (var type = 1; type <= 3; type++) {
                const typeName = type === 1 ? '首期' : (type === 3 ? '末期' : '中间期');
                console.log(`[${this.MODULE_NAME}] --- 处理${typeName}公式 ---`);

                for (const [colIndex, offsetRef] of Object.entries(columnOffsetMap)) {
                    const formula = arrFormula[type][colIndex];

                    if (formula && typeof formula === 'string') {
                        const colName = colIndex === '3' ? '租金' : (colIndex === '4' ? '本金' : '利息');
                        const originalFormula = formula;
                        const newFormula = this._replaceRateInFormula(formula, ratePatterns, offsetRef);

                        if (originalFormula !== newFormula) {
                            arrFormula[type][colIndex] = newFormula;
                            totalReplacements++;
                            console.log(`[${this.MODULE_NAME}] ${typeName}${colName}公式已替换:`);
                            console.log(`[${this.MODULE_NAME}]   原始: ${originalFormula}`);
                            console.log(`[${this.MODULE_NAME}]   替换: ${newFormula}`);
                        } else if (formula.includes('PMT') || formula.includes('interestRate') ||
                                   this._containsRateReference(formula, ratePatterns)) {
                            console.log(`[${this.MODULE_NAME}] ${typeName}${colName}公式未检测到利率引用: ${formula}`);
                        }
                    }
                }
            }

            console.log(`[${this.MODULE_NAME}] 共完成 ${totalReplacements} 处利率引用替换`);
            console.log(`[${this.MODULE_NAME}] ========== 每期适用利率公式生成完成 ==========`);
            return arrFormula;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '适用每期利率' });
            } else {
                console.error(`[${this.MODULE_NAME}] 适用每期利率公式生成失败: ${error.message}`);
            }
            return null;
        }
    }

    _buildRatePatterns(a1Format, r1c1Format) {
        const patterns = [];

        if (a1Format) {
            const escaped = a1Format.replace(/\$/g, '\\$');
            patterns.push({
                name: 'A1绝对引用',
                regex: new RegExp(escaped, 'g')
            });

            const match = a1Format.match(/^\$?([A-Z]+)\$?(\d+)$/);
            if (match) {
                const col = match[1];
                const row = match[2];

                const mixedColRelative = col + '\\$' + row;
                if (mixedColRelative !== escaped) {
                    patterns.push({
                        name: 'A1混合引用_列相对',
                        regex: new RegExp(mixedColRelative, 'g')
                    });
                }

                const mixedRowRelative = '\\$' + col + row;
                if (mixedRowRelative !== escaped) {
                    patterns.push({
                        name: 'A1混合引用_行相对',
                        regex: new RegExp(mixedRowRelative, 'g')
                    });
                }

                var relative = col + row;
                if (relative !== escaped.replace(/\\\$/g, '')) {
                    patterns.push({
                        name: 'A1相对引用',
                        regex: new RegExp('(\\$)?(' + relative + ')', 'g'),
                        hasLookbehind: true
                    });
                }
            }
        }

        if (r1c1Format) {
            const escaped = r1c1Format.replace(/\$/g, '\\$');
            patterns.push({
                name: 'R1C1格式',
                regex: new RegExp(escaped + '(?!\\d)', 'g')
            });
        }

        return patterns;
    }

    _replaceRateInFormula(formula, patterns, replacement) {
        var result = formula;
        for (var i = 0; i < patterns.length; i++) {
            var pattern = patterns[i];
            if (pattern.hasLookbehind) {
                result = result.replace(pattern.regex, function(match, dollarSign, ref) {
                    return dollarSign ? match : replacement;
                });
            } else {
                result = result.replace(pattern.regex, replacement);
            }
        }
        return result;
    }

    _containsRateReference(formula, patterns) {
        for (var i = 0; i < patterns.length; i++) {
            var pattern = patterns[i];
            if (pattern.hasLookbehind) {
                var m;
                while ((m = pattern.regex.exec(formula)) !== null) {
                    if (!m[1]) return true;
                }
            } else {
                if (pattern.regex.test(formula)) {
                    return true;
                }
            }
        }
        return false;
    }
}

// ============== 样式管理器类 ==============
/**
 * clsStyleManager - 样式管理器类
 *
 * 作用：统一管理所有样式设置
 * 设计：避免样式代码重复，统一样式管理
 */
class clsStyleManager {
    constructor(parameterManager) {
        this.p = parameterManager;
        this.MODULE_NAME = "clsStyleManager";

        console.log(`[${this.MODULE_NAME}] 样式管理器初始化完成`);
    }

    createTableHeaders(startCol, endCol, headers) {
        try {
            const ws = this.p.m_worksheet;
            const rowStart = this.p.RowStart;

            const titleCell = ws.Cells(rowStart, 1);
            titleCell.Value2 = "租金测算表";
            titleCell.Interior.Color = COLORS.WHITE;
            titleCell.Font.Name = FONT_CHINESE;
            titleCell.Font.Size = FONT_SIZE_TITLE;
            titleCell.Font.Color = COLORS.BLACK;
            titleCell.HorizontalAlignment = XL.HCenter;

            const headerRange = ws.Range(
                ws.Cells(rowStart + 1, startCol),
                ws.Cells(rowStart + 1, endCol)
            );

            for (var j = 1; j <= (endCol - startCol + 1); j++) {
                headerRange.Cells(1, j).Value2 = headers[(startCol - 1) + (j - 1)];
                const colRange = headerRange.Columns(j);
                colRange.Interior.Color = COLORS.HEADER_BLUE;
                colRange.Font.Name = FONT_DEFAULT;
                colRange.Font.Size = FONT_SIZE_HEADER;
                colRange.Font.Color = COLORS.BLACK;
                colRange.HorizontalAlignment = XL.HCenter;
                colRange.VerticalAlignment = XL.VCenter;
                colRange.WrapText = true;
                colRange.Borders.LineStyle = XL.Continuous;
                colRange.Borders.Weight = XL.Thin;
                colRange.Borders.Color = COLORS.GRAY;
            }

            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '创建表头' });
            } else {
                console.error(`[${this.MODULE_NAME}] 创建表头失败: ${error.message}`);
            }
            return false;
        }
    }

    applyDataFormat(rng) {
        try {
            const totalCols = rng.Columns.Count;

            应用格式(rng.Columns(1), "Integer");
            应用格式(rng.Columns(2), "Date");

            for (var i = 3; i <= 9; i++) {
                应用格式(rng.Columns(i), "Standard");
            }

            if (totalCols >= 10) {
                应用格式(rng.Columns(10), "Integer");
            }
            if (totalCols >= 11) {
                应用格式(rng.Columns(11), "Integer");
            }

            if (totalCols >= 12) {
                应用格式(rng.Columns(12), "Integer");
            }

            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '应用数据格式' });
            } else {
                console.error(`[${this.MODULE_NAME}] 应用数据格式失败: ${error.message}`);
            }
            return false;
        }
    }

    addBorder(rng) {
        try {
            rng.Borders.LineStyle = XL.Continuous;
            rng.Borders.Color = COLORS.BLACK;
            rng.Borders.Weight = XL.Thin;
            rng.Borders.TintAndShade = 0;
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '添加框线' });
            } else {
                console.error(`[${this.MODULE_NAME}] 添加框线失败: ${error.message}`);
            }
            return false;
        }
    }

    setBackColor(rng, color) {
        try {
            rng.Interior.Color = color;
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '设置背景颜色' });
            } else {
                console.error(`[${this.MODULE_NAME}] 设置背景颜色失败: ${error.message}`);
            }
            return false;
        }
    }
}

// ============== 租金测算主类 ==============
/**
 * clsRentalCalculation - 租金测算主类
 * 
 * 作用：协调配置、公式生成、样式管理等模块
 * 设计：组合模式，各模块各司其职
 */
class clsRentalCalculation {
    constructor(parameterManager, undoManager) {
        this.MODULE_NAME = "clsRentalCalculation";
        this.ModuleModifyDate = (typeof VERSION_DATE !== 'undefined') ? VERSION_DATE : "20260409";
        
        // §1 日志级别：'debug' | 'info' | 'warn' | 'error'
        this._logLevel = 'info';

        this.WsTarget = null;
        this.m_staticValueConversion = false;
        // §2 备份控制：true=启用自动备份（默认），false=取消自动备份
        this._backupEnabled = true;
        this.arrHeaders = ["期次", "支付日", "租金", "本金", "利息", "累积偿还本金额",
                          "租金本金余额", "剩余租金余额", "已还租金", "支付日/月间隔",
                          "支付日/月间隔-自定义", "本金比例", "每期适用利率", "备注"];
        this.lengthHeader = this.arrHeaders.length;
        this.arrData = [];
        
        // 子模块（延迟初始化）
        this.config = null;
        this.formulaGenerator = null;
        this.styleManager = null;
        
        // 参数管理器（支持依赖注入）
        this.p = parameterManager || null;
        
        // 撤销管理器支持
        this.m_undoManager = undoManager || (typeof g_undoManager !== "undefined" ? g_undoManager : null);
        this.m_undoEnabled = this.m_undoManager !== null;
        
        this._log('info', `类实例创建${this.m_undoEnabled ? "（已启用撤销功能）" : ""}`);
    }
    
    // ==================== §1 日志工具 ====================
    
    /**
     * _log - 分级日志输出（与 clsParameterManager 一致）
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
     * @param {'debug'|'info'|'warn'|'error'} level
     */
    setLogLevel(level) {
        const validLevels = ['debug', 'info', 'warn', 'error'];
        if (validLevels.indexOf(level) !== -1) {
            this._logLevel = level;
            this._log('info', `日志级别设置为: ${level}`);
        }
    }
    
    /**
     * Initialize - 初始化方法
     * 
     * 设计原则（依赖注入）：不再自行创建 clsParameterManager，
     * 必须由调用方（如 clsRentCalculationSystem）传入已初始化的参数管理器。
     * 
     * @param {Object} parameterManager - 已初始化的参数管理器实例（必需）
     * @param {string} sheetName - 工作表名称（可选，当需要重新初始化参数管理器时使用）
     */
    Initialize(parameterManager, sheetName) {
        try {
            // 依赖注入：优先使用传入的参数管理器
            if (parameterManager && typeof parameterManager === 'object') {
                this.p = parameterManager;
                // 如果提供了工作表名，则重新初始化参数管理器
                if (sheetName) {
                    this.p.Initialize(sheetName);
                }
            } else if (this.p) {
                // 降级：使用构造函数注入的参数管理器（兼容旧调用方式）
                // 如果 parameterManager 实际是 string（旧式调用），忽略它
                if (typeof parameterManager === 'string' && parameterManager !== "1租金测算表V1") {
                    this.p.Initialize(parameterManager);
                }
            } else {
                throw new Error(
                    "clsRentalCalculation.Initialize() 需要传入参数管理器实例。" +
                    "请通过构造函数或 Initialize(paramManager) 注入，不要自行创建。"
                );
            }
            this.WsTarget = this.p.m_worksheet;
            
            // 初始化子模块
            this.config = new clsRentalConfig(this.p);
            this.formulaGenerator = new clsFormulaGenerator(this.p, this.config);
            this.styleManager = new clsStyleManager(this.p);
            
            const totalPeriods = this.p.val("TotalPeriods");
            const principal = this.p.val("Principal");
            
            // P1修复：从参数管理器读取 StaticValueConversion 配置（D1单元格）
            // 此前硬编码为 false，导致转换分支永远是死代码
            this.m_staticValueConversion = this.p.val("StaticValueConversion") || false;
            
            this._log('info', `初始化完成 - 总期数:${totalPeriods}, 租赁成本:${principal}, 静态值转换:${this.m_staticValueConversion}`);
            
            return true;
        } catch (error) {
            const errMsg = `初始化失败：${error.message}`;
            // P1-8: 集成统一错误处理器
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.INITIALIZATION_ERROR, { module: this.MODULE_NAME, function: 'Initialize' });
            } else {
                console.error(`[${this.MODULE_NAME}] ${errMsg}`);
            }
            return false;
        }
    }
    
    /**
     * 创建租金测算表表头
     */
    创建租金测算表表头(startCol = 1, lastCol = 10) {
        // 防御性检查：确保 styleManager 已初始化
        if (!this.styleManager) {
            if (this.p && this.p.IsInitialized) {
                this._log('warn', 'styleManager 未初始化，使用已有参数管理器初始化');
                this.Initialize(this.p);
            } else {
                throw new Error("styleManager 未初始化，且无可用参数管理器。请先调用 Initialize(paramManager)。");
            }
        }
        return this.styleManager.createTableHeaders(startCol, lastCol, this.arrHeaders);
    }
    
    /**
     * 生成公式（根据还款方式）
     */
    generateFormulas(repaymentMethod) {
        const methodConfig = this.config.getRepaymentMethodConfig(repaymentMethod);
        if (!methodConfig) {
            throw new Error(`不支持的还款方式：${repaymentMethod}`);
        }
        
        const formulaMethod = methodConfig.formulaMethod;
        return this.formulaGenerator[formulaMethod]();
    }
    
    /**
     * createDataRange - 创建租金测算表数据区域
     * 
     * 使用标准公式生成方法（generateFormulas），不额外创建表头（由调用方提前创建）
     */
    createDataRange() {
        return this._createTableData(this.generateFormulas.bind(this), {
            logPrefix: "租金测算表",
            createHeaders: false
        });
    }
    
    /**
     * _createTableData - 通用测算表生成模板
     * 
     * 流程：读取偿还方式 → 生成公式 → 转数据数组 → 写入工作表 → 列处理 → 框线 → 额外列 → 格式 → 合计行 → 末期调整
     * 
     * @param {Function} formulaGenerator - 公式生成函数，接收 repaymentMethod 返回 arrFormula
     * @param {Object} [options] - 可选配置
     * @param {string} [options.logPrefix="测算表"] - 日志前缀
     * @param {boolean} [options.createHeaders=false] - 是否额外创建表头
     */
    _createTableData(formulaGenerator, options = {}) {
        const logPrefix = options.logPrefix || "测算表";
        // P1修复：流水线开始前备份，失败时可回滚避免半成品残留
        var _backup = null;
        if (this._backupEnabled) {
            try { _backup = this.backupWorksheetData(); } catch(_e) { this._log('warn', "备份失败（无回滚保障）: " + _e.message); }
        }
        try {
            this._log('info', `========== 开始生成${logPrefix} ==========`);
            
            // 初始化检查
            if (!this.p || !this.p.IsInitialized) {
                throw new Error("clsRentalCalculation 未初始化。请先调用 Initialize(paramManager)。");
            }
            
            // 从参数管理器读取偿还方式
            const repaymentMethodCellA1 = this.p.addr("RepaymentMethod");
            const repaymentMethod = this.WsTarget.Range(repaymentMethodCellA1).Value2;
            this._log('info', `${logPrefix} - 偿还方式: ${repaymentMethod}`);
            
            // 获取偿还方式配置
            const methodConfig = this.config.getRepaymentMethodConfig(repaymentMethod);
            if (!methodConfig) {
                throw new Error(`不支持的还款方式：${repaymentMethod}`);
            }
            
            // 生成公式（差异点：由调用方决定使用哪个公式生成方法）
            const arrFormula = formulaGenerator(repaymentMethod);
            if (!arrFormula) {
                throw new Error(`${logPrefix}公式生成失败`);
            }
            
            // 保存末期公式（角分调整时在公式末尾追加差值，保持公式可见）
            this.lastPeriodFormulas = arrFormula[FORMULA_ROW.LAST];
            
            // 转换为数据数组
            this.arrData = this.arrToArrData(arrFormula);
            
            // 写入工作表
            const rng = this.arrDataToDataRange(this.arrData);
            if (!rng) {
                throw new Error("数据写入失败");
            }
            
            // 应用列处理（转换为数值和清除指定列）
            this.列转化成数值以及清除(rng, methodConfig.convertColumns, methodConfig.clearColumns);
            
            // 添加框线
            for (const frameRange of methodConfig.addFrame) {
                this.添加框线(rng.Columns(frameRange));
            }
            
            // 处理额外列（如本金比例列）
            for (const colIndex of methodConfig.extraColumns) {
                if (colIndex === 12) {
                    this.创建租金测算表表头(12, 12);
                    this.添加框线(rng.Columns(12));
                    this.设置背景颜色(rng.Columns(12), this.p.m_COLOR_YELLOW);
                    this.租金测算表合计行(12, 12);
                }
            }
            
            // 应用数据格式
            this.styleManager.applyDataFormat(rng);
            设置表格样式(rng);
            
            // 额外创建表头（可选）
            if (options.createHeaders) {
                this.创建租金测算表表头();
            }
            
            // 生成合计行
            this.租金测算表合计行(1, 10);
            
            // 写入末期调整备注（N列）
            this.写入末期调整备注(repaymentMethod);

            // P1修复：末期调整后触发最终重算，确保所有公式值正确
            try { this.WsTarget.Calculate(); } catch(_ce) { this._log('warn', `最终重算警告: ${_ce.message}`); }
            
            this._log('info', `========== ${logPrefix}生成完成 ==========`);
            return true;
        } catch (error) {
            // P1修复：失败时尝试恢复备份，避免半成品残留
            if (_backup) {
                this._log('warn', `${logPrefix}生成失败，尝试恢复备份...`);
                try { this.restoreWorksheetData(_backup); this._log('info', '备份恢复成功'); } catch(_re) { this._log('error', `备份恢复失败: ${_re.message}`); }
            }
            // P1-8: 集成统一错误处理器
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: logPrefix });
            } else {
                console.error(`[${this.MODULE_NAME}] ${logPrefix}生成失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 列转化成数值以及清除
     * 
     * 作用：根据配置将公式列转为数值、清除指定列
     * 
     * 开关：this.m_staticValueConversion（默认 false）
     * - false：仅转换结构列（期次A、日期B），保留数据列公式 → 本金/租金/利息显示公式
     * - true：转换 convertColumns 中的所有列为纯数值 → 全部显示数值
     * 
     * 结构列（col 1=期次, col 2=日期）始终转为数值，不受开关影响
     */
    列转化成数值以及清除(rng, convertColumns = [], clearColumns = []) {
        try {
            const STRUCTURAL_COLUMNS = [1, 2]; // 期次、日期列始终转为数值
            
            // 结构列始终转换（不受开关影响）
            STRUCTURAL_COLUMNS.forEach(colIndex => {
                if (colIndex > 0) {
                    const colRange = rng.Columns.Item(colIndex);
                    colRange.Value2 = colRange.Value2;
                }
            });
            
            // 数据列转换受 m_staticValueConversion 开关控制
            if (this.m_staticValueConversion) {
                convertColumns.forEach(colIndex => {
                    // 跳过已处理的结构列
                    if (colIndex > 0 && !STRUCTURAL_COLUMNS.includes(colIndex)) {
                        const colRange = rng.Columns.Item(colIndex);
                        colRange.Value2 = colRange.Value2;
                    }
                });
            }
            
            // 清除指定列（不受开关影响）
            clearColumns.forEach(colIndex => {
                if (colIndex > 0) {
                    const colRange = rng.Columns.Item(colIndex);
                    colRange.ClearContents();
                }
            });
            
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '列转化成数值以及清除' });
            } else {
                console.error(`处理列数据失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * arrToArrData - 将公式模板展开为逐行数据数组（Array2D 优化版）
     * 
     * arrFormula 结构：[0]=空(废位), [1]=首行公式, [2]=中间行公式, [3]=末期公式
     * 展开规则：
     *   - 单期（length===1）：使用首行公式（FORMULA_ROW.FIRST），首行公式是自含的不引用 R[-1]
     *   - 多期：第0行用 FIRST，最后一行用 LAST，中间行用 MIDDLE
     * 
     * P0修复：单期时 row===0 既是首行又是末行，原来用 FORMULA_ROW.LAST（末期公式含R[-1]引用）
     *        导致引用越界。现在单期统一用 FORMULA_ROW.FIRST（首行公式自含，无R[-1]引用）。
     * 
     * Array2D 优化：使用批量处理替代逐元素赋值，减少循环内条件判断次数
     * 
     * @param {Array} arrFormula - 公式模板数组（4行×N列）
     * @param {number} actualLength - 可选，指定行数（默认从参数管理器读取 TotalPeriods）
     */
    arrToArrData(arrFormula, actualLength) {
        try {
            var length = (actualLength !== undefined && actualLength !== null)
                ? actualLength
                : this.p.val("TotalPeriods");
            
            // Array2D 优化：使用优化器处理大数据集
            if (length > 50 && typeof clsArray2DOptimizer !== 'undefined') {
                var optimizer = getArray2DOptimizer(this.p);
                var result = optimizer.arrToArrDataOptimized(arrFormula, length, FORMULA_ROW);
                if (result) {
                    this._log('info', `数据数组生成完成（优化），长度: ${length}`);
                    return result;
                }
            }
            
            // 降级：原实现
            return this._arrToArrDataOriginal(arrFormula, length);
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: 'arrToArrData' });
            } else {
                console.error(`[${this.MODULE_NAME}] 租金测算表数据数组转换失败：${error.message}`);
            }
            return null;
        }
    }
    
    /**
     * _arrToArrDataOriginal - 原实现（降级使用）
     * @private
     */
    _arrToArrDataOriginal(arrFormula, length) {
        const maxCol = arrFormula[FORMULA_ROW.FIRST].length - 1;
        
        var arrData = [];
        for (var i = 0; i < length; i++) {
            arrData[i] = new Array(arrFormula[FORMULA_ROW.FIRST].length);
        }
        
        for (var row = 0; row < length; row++) {
            var rowIndex;
            if (length === 1) {
                rowIndex = FORMULA_ROW.FIRST;
            } else if (row === 0) {
                rowIndex = FORMULA_ROW.FIRST;
            } else if (row === length - 1) {
                rowIndex = FORMULA_ROW.LAST;
            } else {
                rowIndex = FORMULA_ROW.MIDDLE;
            }
            for (var col = 1; col <= maxCol && col < arrFormula[FORMULA_ROW.FIRST].length; col++) {
                arrData[row][col - 1] = arrFormula[rowIndex][col];
            }
        }
        
        this._log('info', `数据数组生成完成，降级实现，长度: ${length}`);
        return arrData;
    }
    
    /**
     * arrToArrDataWithLength - 根据指定长度生成数据数组（兼容方法，委托给 arrToArrData）
     * 
     * @param {Array} arrFormula - 公式数组
     * @param {number} actualLength - 指定的数据长度
     * @deprecated 请直接使用 arrToArrData(arrFormula, actualLength)
     */
    arrToArrDataWithLength(arrFormula, actualLength) {
        return this.arrToArrData(arrFormula, actualLength);
    }
    
    /**
     * arrDataToDataRange - 数据数组写入工作表
     * 
     * P0修复：写入范围动态匹配数据列宽，而非硬编码 A:L
     * - arrData 每行有13个元素（A-M列），之前只写到 L 列（m_COL_PRINCIPAL_RATIO），
     *   导致 M 列（每期适用利率）数据静默丢弃
     * - 现在根据 arrData[0].length 动态计算结束列，确保所有数据都写入
     */
    arrDataToDataRange(arrData) {
        try {
            const m_COL_PERIOD = this.p.m_COL_PERIOD;
            const rentTableStartRow = this.p.RentTableStartRow;
            const actualLength = arrData.length;
            const m_worksheet = this.p.m_worksheet;
            
            if (!arrData || !Array.isArray(arrData) || arrData.length === 0) {
                throw new Error("数据数组无效");
            }
            
            // P0修复：根据数据实际列数动态计算结束列
            // arrData 每行元素对应 A(0), B(1), ..., L(11), M(12) 共13列
            const dataColCount = arrData[0] ? arrData[0].length : 12;
            // 列号转字母：0=A, 1=B, ..., 11=L, 12=M
            const endColLetter = String.fromCharCode(65 + dataColCount - 1);
            
            const rngData = m_worksheet.Range(
                `${m_COL_PERIOD}${rentTableStartRow}:${endColLetter}${rentTableStartRow + actualLength - 1}`
            );
            
            rngData.Value2 = arrData;
            
            this._log('info', `数据数组写入完成，范围: ${m_COL_PERIOD}${rentTableStartRow}:${endColLetter}${rentTableStartRow + actualLength - 1}，列数: ${dataColCount}`);
            return rngData;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: 'arrDataToDataRange' });
            } else {
                console.error(`[${this.MODULE_NAME}] 租金测算表数据数组写入表格失败：${error.message}`);
            }
            return null;
        }
    }
    
    
    /**
     * 租金测算表应用格式（兼容方法）
     * 
     * 委托给 mStyleManager，子类（如 m直租）大量调用此方法
     */
    租金测算表应用格式(rng) {
        return this.styleManager.applyDataFormat(rng);
    }
    
    /**
     * 添加框线（兼容方法）
     * 
     * 委托给 mStyleManager，子类（如 m直租、mCashFlowGenerator、m调息）大量调用此方法
     */
    添加框线(rng) {
        return this.styleManager.addBorder(rng);
    }
    
    /**
     * 设置背景颜色（兼容方法）
     * 
     * 委托给 mShared_constants 全局函数，子类（如 m直租、m调息）调用此方法
     */
    设置背景颜色(rng, color) {
        return 设置背景颜色(rng, color);
    }
    
    /**
     * 清除原有表中数据
     */
    清除原有表中数据() {
        try {
            const WsTarget = this.p.m_worksheet;

            WsTarget.Range("B16:B17").ClearContents();
            WsTarget.Range("D16:D18").ClearContents();

            // 清除范围：当前期数不能反映历史数据行数（用户可能改过参数），
            // 因此取当前期数与36的最大值作为兜底，+30行余量确保覆盖合计行 + 备注 + 旧数据
            const totalPeriods = Math.max(this.p.val("TotalPeriods") || 0, 36);
            const clearEndRow = this.p.RentTableStartRow + totalPeriods + 10;
            const clearRange = WsTarget.Range(
                `A${this.p.RowStart}:N${clearEndRow}`
            );
            clearRange.Clear();

            this._log('info', `原有数据清除完成，范围: A${this.p.RowStart}:N${clearEndRow}`);
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '清除原有表中数据' });
            } else {
                console.error(`[${this.MODULE_NAME}] 数据清除失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 租金测算表合计行
     */
    租金测算表合计行(startCol = 1, lastCol = 12) {
        try {
            const rentTableStartRow = this.p.RentTableStartRow;
            const totalPeriodsValue = this.p.val("TotalPeriods");
            const lastRow = rentTableStartRow + totalPeriodsValue;
            
            // 动态计算合计行范围：从 startCol 到 lastCol 对应的列字母
            // 列号转字母：1=A, 2=B, ..., 12=L, 14=N
            const startColLetter = String.fromCharCode(64 + startCol);
            const endColLetter = String.fromCharCode(64 + lastCol);
            const range = this.p.m_worksheet.Range(`${startColLetter}${lastRow}:${endColLetter}${lastRow}`);
            const sumFormula = `=SUM(R[${-totalPeriodsValue}]C:R[-1]C)`;
            
            for (var i = 1; i <= lastCol; i++) {
                if (i < startCol || i > lastCol) continue;
                
                if (i === 1) {
                    range.Columns(1).Value2 = "合计";
                    this.styleManager.addBorder(range.Columns(1));
                } else if (i === 2) {
                    range.Columns(2).Value2 = "-";
                    this.styleManager.addBorder(range.Columns(2));
                } else if ((i >= 3 && i <= 5) || (i >= 10 && i <= 12)) {
                    range.Columns(i).FormulaR1C1 = sumFormula;
                    this.styleManager.addBorder(range.Columns(i));
                } else if (i >= 6 && i <= 9) {
                    range.Columns(i).Value2 = "-";
                    this.styleManager.addBorder(range.Columns(i));
                }
            }
            
            range.Font.Bold = true;
            range.NumberFormat = "#,##0.00";
            range.HorizontalAlignment = XL.HCenter;
            range.VerticalAlignment = XL.VCenter;
            
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '租金测算表合计行' });
            } else {
                console.error(`租金测算表合计行生成失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 自定义月间隔
     */
    自定义月间隔(targetRange, startDateCellA1) {
        try {
            this._log('debug', `开始批量生成日期公式，目标范围：${targetRange.Address}`);
            this._log('debug', `起租日单元格：${startDateCellA1}`);
            
            var i = 0;
            var cell = null;
            var formula = "";
            var rowCount = targetRange.Rows.Count;
            
            for (var row = 1; row <= rowCount; row++) {
                cell = targetRange.Cells(row, 1);
                i = i + 1;
                if (i === 1) {
                    formula = `=EDATE(${startDateCellA1}, K${cell.Row})`;
                } else {
                    formula = "=EDATE(B" + (cell.Row - 1) + ",K" + cell.Row + ")";
                }
                cell.Formula = formula;
                cell.NumberFormat = "yyyy-mm-dd";
                this._log('debug', `第${i}期单元格：${cell.Address}，公式设置完成:${formula}`);
            }
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '自定义月间隔' });
            } else {
                console.error(`自定义月间隔失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 生成月间隔
     */
    生成月间隔() {
        try {
            const m_COL_MONTH_INTERVAL = this.p.m_COL_MONTH_INTERVAL;
            const m_COL_CUSTOM_INTERVAL = this.p.m_COL_CUSTOM_INTERVAL;
            const m_COL_DATE = this.p.m_COL_DATE;
            const rentTableStartRow = this.p.RentTableStartRow;
            const totalPeriodsValue = this.p.val("TotalPeriods");
            const leaseStartDateCellA1 = this.p.addr("LeaseStartDate");
            const m_worksheet = this.p.m_worksheet;
            
            this.创建租金测算表表头(11, 11);
            
            const sourceRange = m_worksheet.Range(
                `${m_COL_MONTH_INTERVAL}${rentTableStartRow}:${m_COL_MONTH_INTERVAL}${rentTableStartRow + totalPeriodsValue - 1}`
            );
            const tarRange = m_worksheet.Range(
                `${m_COL_CUSTOM_INTERVAL}${rentTableStartRow}:${m_COL_CUSTOM_INTERVAL}${rentTableStartRow + totalPeriodsValue - 1}`
            );
            const targetRange = m_worksheet.Range(
                `${m_COL_DATE}${rentTableStartRow}:${m_COL_DATE}${rentTableStartRow + totalPeriodsValue - 1}`
            );
            
            tarRange.Value2 = sourceRange.Value2;
            this.自定义月间隔(targetRange, leaseStartDateCellA1);
            this.styleManager.addBorder(tarRange);
            this.styleManager.setBackColor(tarRange, this.p.m_COLOR_YELLOW);
            this.租金测算表合计行(11, 11);
            
            const r1 = m_worksheet.Range(`${m_COL_CUSTOM_INTERVAL}${rentTableStartRow + totalPeriodsValue}`);
            this.styleManager.addBorder(r1);
            
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '生成月间隔' });
            } else {
                console.error(`生成月间隔失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 写入每期利率
     */
    写入每期利率() {
        const rowStart = this.p.RentTableStartRow;
        const totalPeriod = this.p.val("TotalPeriods");
        const ratePerPeriod = this.p.val("InterestRate");
        
        var arr2D = new Array(totalPeriod);
        for(var i = 0; i < totalPeriod; i++){
            arr2D[i] = [ratePerPeriod];
        }
        return arr2D;
    }
    
    /**
     * 调整期利率
     */
    调整期利率(arr2D, rate, periodn) {
        try {
            const totalPeriodsValue = this.p.val("TotalPeriods");
            const m_COL_RatePerPeriod = this.p.m_COL_RatePerPeriod;
            const rentTableStartRow = this.p.RentTableStartRow;
            const m_worksheet = this.p.m_worksheet;
            
            if (!arr2D) {
                arr2D = this.写入每期利率();
            }
            
            for(var i = periodn - 1; i < totalPeriodsValue; i++){
                arr2D[i][0] = rate;
            }
            
            const rngData = m_worksheet.Range(
                `${m_COL_RatePerPeriod}${rentTableStartRow}:${m_COL_RatePerPeriod}${rentTableStartRow + totalPeriodsValue - 1}`
            );
            rngData.Value2 = arr2D;
            
            return arr2D;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '调整期利率' });
            } else {
                console.error(`调整期利率失败：${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 每期适用利率
     */
    每期适用利率() {
        const m_COL_RatePerPeriod = this.p.m_COL_RatePerPeriod;
        const rentTableStartRow = this.p.RentTableStartRow;
        const totalPeriodsValue = this.p.val("TotalPeriods");
        const m_worksheet = this.p.m_worksheet;
        
        const rngData = m_worksheet.Range(
            `${m_COL_RatePerPeriod}${rentTableStartRow}:${m_COL_RatePerPeriod}${rentTableStartRow + totalPeriodsValue - 1}`
        );
        var arrData = this.写入每期利率();
        rngData.Value2 = arrData;
        应用格式(rngData, "Percentage");
        return true;
    }
    
    /**
     * 使用每期适用利率生成测算表
     * 
     * 与 createDataRange 的区别：使用 formulaGenerator.适用每期利率 生成公式（M列利率由用户手动设置），
     * 并额外创建表头
     */
    使用每期适用利率生成测算表() {
        return this._createTableData(
            (repaymentMethod) => this.formulaGenerator.适用每期利率(repaymentMethod),
            { logPrefix: "每期适用利率测算表", createHeaders: true }
        );
    }
    
    /**
     * 写入末期调整备注
     * 
     * 作用：在最后一期进行角分调整，使本金合计精确等于租赁成本，并在N列记录调整金额
     * 
     * 逻辑：
     * 1. 所有期次先用标准公式计算（包括最后一期）
     * 2. 对比本金合计与租赁成本的差额（ROUND导致的角分累积误差）
     * 3. 差额≠0时调整最后一期的值，使本金合计精确等于租赁成本
     * 4. N列记录调整金额
     * 
     * 各还款方式调整策略：
     * - 等额本息（后付）：调整D列本金，E列利息公式自动跟随（利息=租金-本金）
     * - 等额本息（先付）：调整C列租金，D列本金公式自动跟随（本金=租金-利息）
     * - 其他方式：调整D列本金，C列租金公式自动跟随（租金=本金+利息）
     * 
     * @param {string} repaymentMethod - 偿还方式
     */
    写入末期调整备注(repaymentMethod) {
        try {
            this._log('info', `开始角分调整与备注，偿还方式: ${repaymentMethod}`);
            
            // 强制重算以确保所有公式值是最新的
            this.WsTarget.Calculate();
            
            const startRow = this.p.RentTableStartRow;
            const totalPeriods = this.p.val("TotalPeriods");
            const principal = this.p.val("Principal");
            const lastRow = startRow + totalPeriods - 1;
            const totalRow = startRow + totalPeriods;
            
            // 读取合计行的本金合计值（D列）
            const principalSum = this.WsTarget.Range(`D${totalRow}`).Value2;
            
            // 计算本金差异 = 租赁成本 - 本金合计
            const principalDiff = Math.round((principal - principalSum) * 100) / 100;
            
            // 如果差异为0，无需调整
            if (principalDiff === 0) {
                this._log('info', '本金合计与租赁成本一致，无角分差异');
                return true;
            }
            
            // 读取调整前的末期利息
            const interestBefore = this.WsTarget.Range(`E${lastRow}`).Value2;
            
            // 构建调整后的公式：原公式+差值（保持公式可见，而非替换为纯数字）
            // lastPeriodFormulas 是 arr[3]，索引对应列号（3=C租金, 4=D本金, 5=E利息）
            const adjustCol = (repaymentMethod === "等额本息（先付）") ? 3 : 4; // 先付调整租金列，其他调整本金列
            const adjustColLetter = (adjustCol === 3) ? 'C' : 'D';
            const originalFormulaR1C1 = this.lastPeriodFormulas ? this.lastPeriodFormulas[adjustCol] : null;
            
            if (originalFormulaR1C1 && typeof originalFormulaR1C1 === 'string' && originalFormulaR1C1.startsWith('=')) {
                // 在原公式后面追加差值，如 =ROUND(-PPMT(...),2)+0.03
                const adjustedFormula = originalFormulaR1C1 + (principalDiff >= 0 ? '+' : '') + principalDiff;
                this.WsTarget.Range(`${adjustColLetter}${lastRow}`).FormulaR1C1 = adjustedFormula;
                this._log('info', `末期公式调整: ${adjustedFormula}`);
            } else {
                // 降级：无原始公式时直接用数值
                const lastValue = this.WsTarget.Range(`${adjustColLetter}${lastRow}`).Value2;
                this.WsTarget.Range(`${adjustColLetter}${lastRow}`).Value2 = lastValue + principalDiff;
                this._log('info', `末期数值调整（降级）: ${lastValue}→${lastValue + principalDiff}`);
            }
            
            // 重算工作表
            this.WsTarget.Calculate();
            
            // 读取调整后的末期利息，计算利息变化
            const interestAfter = this.WsTarget.Range(`E${lastRow}`).Value2;
            const interestDiff = Math.round((interestAfter - interestBefore) * 100) / 100;
            
            // 创建N列表头并写入备注
            this.创建租金测算表表头(14, 14);
            const note = `最后一期本金调整：${principalDiff.toFixed(2)}，利息调整：${interestDiff.toFixed(2)}`;
            this.WsTarget.Range(`N${lastRow}`).Value2 = note;
            
            this._log('info', `角分调整完成: ${note}（本金合计 ${principalSum}→${principal}）`);
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '写入末期调整备注' });
            } else {
                console.error(`[${this.MODULE_NAME}] 角分调整与备注失败: ${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 改变自定义支付日
     * 
     * 作用：修改指定期的自定义支付日/月间隔值
     * 
     * @param {number} period - 期次（1表示第1期，-1表示最后一期）
     * @param {number} value - 要设置的值
     * @returns {boolean} 是否设置成功
     */
    改变自定义支付日(period, value) {
        try {
            const m_COL_CUSTOM_INTERVAL = this.p.m_COL_CUSTOM_INTERVAL;
            const rentTableStartRow = this.p.RentTableStartRow;
            const totalPeriodsValue = this.p.val("TotalPeriods");
            const m_worksheet = this.p.m_worksheet;
            
            // 计算实际的行号
            var rowIndex;
            if (period === -1) {
                // -1 表示最后一期
                rowIndex = totalPeriodsValue;
            } else {
                rowIndex = period;
            }
            
            // 验证期次范围
            if (rowIndex < 1 || rowIndex > totalPeriodsValue) {
                this._log('warn', `期次 ${period} 超出范围 (1-${totalPeriodsValue})`);
                return false;
            }
            
            // 计算单元格地址
            const cellAddress = `${m_COL_CUSTOM_INTERVAL}${rentTableStartRow + rowIndex - 1}`;
            
            // 设置值
            m_worksheet.Range(cellAddress).Value2 = value;
            
            this._log('info', `第${rowIndex}期自定义支付日/月间隔已设置为: ${value}`);
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '改变自定义支付日' });
            } else {
                console.error(`[${this.MODULE_NAME}] 改变自定义支付日失败: ${error.message}`);
            }
            return false;
        }
    }
    
    // ============== 统一撤销管理器支持 ==============
    
    /**
     * 设置撤销管理器
     * @param {clsUndoManager} undoManager - 撤销管理器实例
     */
    setUndoManager(undoManager) {
        this.m_undoManager = undoManager;
        this.m_undoEnabled = undoManager !== null;
        this._log('info', `撤销管理器${this.m_undoEnabled ? "已设置" : "已清除"}`);
    }
    
    /**
     * 获取撤销管理器
     * @returns {clsUndoManager|null} 撤销管理器实例
     */
    getUndoManager() {
        return this.m_undoManager;
    }
    
    /**
     * 检查是否启用了撤销功能
     * @returns {boolean} 是否启用
     */
    isUndoEnabled() {
        return this.m_undoEnabled;
    }
    
    /**
     * 启用/禁用撤销功能
     * @param {boolean} enabled - 是否启用
     */
    setUndoEnabled(enabled) {
        this.m_undoEnabled = enabled && this.m_undoManager !== null;
        this._log('info', `撤销功能已${this.m_undoEnabled ? "启用" : "禁用"}`);
    }
    
    /**
     * 执行可撤销的操作
     * @param {string} type - 操作类型
     * @param {string} description - 操作描述
     * @param {Function} doFn - 执行函数
     * @param {Function} undoFn - 撤销函数
     * @param {Object} metadata - 元数据
     * @returns {boolean} 是否成功
     */
    executeUndoable(type, description, doFn, undoFn, metadata) {
        if (!this.m_undoEnabled || !this.m_undoManager) {
            // 撤销功能未启用，直接执行
            return doFn.call(this);
        }
        
        // 创建命令
        const self = this;
        const command = {
            execute: function() {
                return doFn.call(self);
            },
            undo: function() {
                return undoFn.call(self);
            },
            getInfo: function() {
                return {
                    type: type,
                    description: description,
                    metadata: metadata
                };
            }
        };
        
        // 包装为标准命令
        const wrappedCommand = new clsUndoableCommand(
            this,
            type,
            description,
            command.execute,
            command.undo,
            metadata
        );
        
        return this.m_undoManager.execute(wrappedCommand);
    }
    
    /**
     * 备份当前工作表数据
     * @returns {Object|null} 备份数据对象
     */
    backupWorksheetData() {
        try {
            if (!this.p || !this.p.m_worksheet) {
                return null;
            }
            
            const ws = this.p.m_worksheet;
            const startRow = this.p.RentTableStartRow;
            const totalPeriods = this.p.val("TotalPeriods");
            const endRow = startRow + totalPeriods + 10; // 多备份一些行
            
            const backupRange = `A${startRow}:N${endRow}`;
            const backupData = ws.Range(backupRange).Value2;
            
            return {
                range: backupRange,
                data: backupData,
                timestamp: new Date()
            };
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: 'backupWorksheetData' });
            } else {
                console.error(`[${this.MODULE_NAME}] 备份工作表数据失败: ${error.message}`);
            }
            return null;
        }
    }
    
    /**
     * 恢复工作表数据
     * @param {Object} backup - 备份数据对象
     * @returns {boolean} 是否成功
     */
    restoreWorksheetData(backup) {
        try {
            if (!backup || !backup.data) {
                return false;
            }
            
            if (!this.p || !this.p.m_worksheet) {
                return false;
            }
            
            const ws = this.p.m_worksheet;
            ws.Range(backup.range).Value2 = backup.data;
            ws.Calculate();
            
            this._log('info', '工作表数据已恢复');
            return true;
        } catch (error) {
            if (typeof g_errorHandler !== 'undefined') {
                g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: 'restoreWorksheetData' });
            } else {
                console.error(`[${this.MODULE_NAME}] 恢复工作表数据失败: ${error.message}`);
            }
            return false;
        }
    }
    
    /**
     * 创建数据区域（带撤销支持）
     * @returns {boolean} 是否成功
     */
    createDataRangeWithUndo() {
        const self = this;
        const backup = this._backupEnabled ? this.backupWorksheetData() : null;
        
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.TABLE_GENERATE,
            "创建租金测算表数据区域",
            function() {
                return self.createDataRange();
            },
            function() {
                if (backup) {
                    return self.restoreWorksheetData(backup);
                }
                return false;
            },
            { hasBackup: backup !== null }
        );
    }
    
    /**
     * 清除数据（带撤销支持）
     * @returns {boolean} 是否成功
     */
    清除原有表中数据WithUndo() {
        const self = this;
        const backup = this._backupEnabled ? this.backupWorksheetData() : null;
        
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY,
            "清除原有表中数据",
            function() {
                return self.清除原有表中数据();
            },
            function() {
                if (backup) {
                    return self.restoreWorksheetData(backup);
                }
                return false;
            },
            { hasBackup: backup !== null }
        );
    }
    
    /**
     * 使用每期适用利率生成测算表（带撤销支持）
     * @returns {boolean} 是否成功
     */
    使用每期适用利率生成测算表WithUndo() {
        const self = this;
        const backup = this._backupEnabled ? this.backupWorksheetData() : null;
        
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.RATE_ADJUST,
            "使用每期适用利率生成测算表",
            function() {
                return self.使用每期适用利率生成测算表();
            },
            function() {
                if (backup) {
                    return self.restoreWorksheetData(backup);
                }
                return false;
            },
            { hasBackup: backup !== null }
        );
    }
    
    /**
     * 改变自定义支付日（带撤销支持）
     * @param {number} period - 期次
     * @param {number} value - 值
     * @returns {boolean} 是否成功
     */
    改变自定义支付日WithUndo(period, value) {
        const self = this;
        
        // 获取旧值
        var oldValue = null;
        try {
            const m_COL_CUSTOM_INTERVAL = this.p.m_COL_CUSTOM_INTERVAL;
            const rentTableStartRow = this.p.RentTableStartRow;
            var rowIndex = period === -1 ? this.p.val("TotalPeriods") : period;
            const cellAddress = `${m_COL_CUSTOM_INTERVAL}${rentTableStartRow + rowIndex - 1}`;
            oldValue = this.p.m_worksheet.Range(cellAddress).Value2;
        } catch (e) {
            // 忽略错误
        }
        
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY,
            `改变第${period}期自定义支付日为${value}`,
            function() {
                return self.改变自定义支付日(period, value);
            },
            function() {
                if (oldValue !== null) {
                    return self.改变自定义支付日(period, oldValue);
                }
                return false;
            },
            { period: period, oldValue: oldValue, newValue: value }
        );
    }
    
    /**
     * 撤销上一步操作
     * @returns {Object} 撤销结果
     */
    undo() {
        if (this.m_undoManager) {
            return this.m_undoManager.undo();
        }
        return { success: false, message: "撤销管理器未初始化" };
    }
    
    /**
     * 重做上一步撤销的操作
     * @returns {Object} 重做结果
     */
    redo() {
        if (this.m_undoManager) {
            return this.m_undoManager.redo();
        }
        return { success: false, message: "撤销管理器未初始化" };
    }
    
    /**
     * 检查是否可以撤销
     * @returns {boolean} 是否可以撤销
     */
    canUndo() {
        return this.m_undoManager ? this.m_undoManager.canUndo() : false;
    }
    
    /**
     * 检查是否可以重做
     * @returns {boolean} 是否可以重做
     */
    canRedo() {
        return this.m_undoManager ? this.m_undoManager.canRedo() : false;
    }
    
    /**
     * 获取撤销历史
     * @returns {Array} 历史记录
     */
    getUndoHistory() {
        return this.m_undoManager ? this.m_undoManager.getUndoHistory() : [];
    }
    
    /**
     * 清空撤销历史
     */
    clearUndoHistory() {
        if (this.m_undoManager) {
            this.m_undoManager.clear();
        }
    }
}

// ========== 租金测算表行管理器类 ==========
/**
 * clsRentalTableRowManager - 租金测算表行管理器
 * 
 * 作用：继承自 clsRentalCalculation，处理行插入和删除
 * 改进（V2.1重构）：
 * - 统一使用 clsUndoManager 进行撤销（移除了独立的 operationHistory）
 * - 抽取 _performRowOperation 模板方法消除 insertRows/deleteRows 重复代码
 * - 删除操作撤销时恢复原始数据快照（而非仅插入空行）
 */
class clsRentalTableRowManager extends clsRentalCalculation {
    constructor(parameterManager, undoManager) {
        super(parameterManager, undoManager);
        this.MODULE_NAME = "clsRentalTableRowManager";
        this._log('info', `类实例创建（继承自clsRentalCalculation${this.m_undoEnabled ? "，已启用撤销功能" : ""}）`);
    }
    
    // ============== 核心操作方法 ==============
    
    /**
     * insertRows - 插入行（委托给模板方法）
     */
    insertRows(arrData, insertPosition, rowCount, arrFormula = null, repaymentMethod = null, options = {}) {
        return this._performRowOperation('insert', arrData, insertPosition, rowCount, arrFormula, repaymentMethod, options);
    }

    /**
     * deleteRows - 删除行（委托给模板方法）
     */
    deleteRows(arrData, deletePosition, rowCount, arrFormula = null, repaymentMethod = null, options = {}) {
        return this._performRowOperation('delete', arrData, deletePosition, rowCount, arrFormula, repaymentMethod, options);
    }
    
    /**
     * _performRowOperation - 行操作模板方法
     * 
     * 作用：统一处理插入和删除操作的公共逻辑，消除重复代码
     * 设计：模板方法模式（Template Method），insert/delete 差异通过 operationType 控制
     * 
     * 公共步骤：验证 → 获取还款方式 → 公式更新 → 本金比例 → 更新总期数 → 同步工作表
     * 差异点：数组变更方式（splice方向）、工作表同步细节（清除/高亮）
     * 
     * @param {string} operationType - 'insert' 或 'delete'
     */
    _performRowOperation(operationType, arrData, position, rowCount, arrFormula, repaymentMethod, options) {
        try {
            const opName = operationType === 'insert' ? '插入' : '删除';
            this._log('info', `========== ${opName}行操作开始 ==========`);

            const mergedOptions = this._mergeOptions(options);

            // 步骤1：验证输入
            const validation = this.validateInput(arrData, position, rowCount, operationType);
            if (!validation.valid) {
                throw new Error(validation.error);
            }

            // 步骤2：获取还款方式
            if (!repaymentMethod) {
                const repaymentMethodCellA1 = this.p.addr("RepaymentMethod");
                repaymentMethod = this.WsTarget.Range(repaymentMethodCellA1).Value2;
            }

            const oldTotalPeriods = arrData.length;
            const newTotalPeriods = operationType === 'insert' 
                ? oldTotalPeriods + rowCount 
                : oldTotalPeriods - rowCount;

            this._log('info', `${opName}位置: ${position} (第${position + 1}行)`);
            this._log('info', `${opName}行数: ${rowCount}`);
            this._log('info', `还款方式: ${repaymentMethod}`);
            this._log('info', `原总期数: ${oldTotalPeriods}`);
            this._log('info', `新总期数: ${newTotalPeriods}`);

            // 步骤3：保存操作前快照（用于撤销）
            const snapshot = this._saveSnapshot(arrData, repaymentMethod);

            // 步骤4：执行数组变更（insert/delete 差异点）
            if (operationType === 'insert') {
                const insertedRows = [];
                for (var i = 0; i < rowCount; i++) {
                    insertedRows.push(new Array(arrFormula ? arrFormula[FORMULA_ROW.FIRST].length : 13));
                }
                arrData.splice(position, 0, ...insertedRows);
            } else {
                arrData.splice(position, rowCount);
            }

            // 步骤5：公式重新生成与更新（公共逻辑）
            const needRegenerate = this.shouldRegenerateFormulas(repaymentMethod);
            if (needRegenerate || !arrFormula) {
                const originalTotalPeriods = this.p.val("TotalPeriods");
                this.p.m_worksheet.Range(this.p.addr("TotalPeriods")).Value2 = newTotalPeriods;
                arrFormula = this.generateFormulas(repaymentMethod);
                if (!arrFormula) {
                    throw new Error("重新生成公式模板失败");
                }
                this.p.m_worksheet.Range(this.p.addr("TotalPeriods")).Value2 = originalTotalPeriods;
            }

            // 步骤6：更新受影响行的公式（公共逻辑）
            const modifiedRows = this.identifyAffectedRows(oldTotalPeriods, newTotalPeriods, position, operationType);
            for (const rowIndex of modifiedRows) {
                const formulaType = this.getFormulaType(rowIndex, newTotalPeriods);
                arrData = this.updateRowFormulas(arrData, rowIndex, formulaType, arrFormula);
            }

            arrData = this.updateRemainingBalanceFormula(arrData);

            // 步骤7：本金比例重分配（公共逻辑）
            if (repaymentMethod.includes("本金比例")) {
                arrData = this.redistributePrincipalRatio(arrData, newTotalPeriods);
            }

            // 步骤8：更新总期数（公共逻辑）
            if (mergedOptions.autoUpdateTotalPeriods) {
                const updateSuccess = this.updateTotalPeriods(newTotalPeriods);
                if (!updateSuccess) {
                    throw new Error("更新TotalPeriods参数失败");
                }
            }

            // 步骤9：同步到工作表（insert/delete 有差异）
            if (mergedOptions.syncToWorksheet) {
                this._syncToWorksheet(arrData, operationType, oldTotalPeriods, newTotalPeriods, position, rowCount, repaymentMethod);
            }

            // 步骤10：重新计算工作表（公共逻辑）
            if (mergedOptions.recalculateSheet && mergedOptions.syncToWorksheet) {
                try {
                    this.p.m_worksheet.Calculate();
                } catch (calcError) {
                    this._log('warn', `工作表重新计算警告: ${calcError.message}`);
                }
            }

            // 步骤11：注册撤销命令到统一管理器
            this._registerUndoCommand(operationType, position, rowCount, repaymentMethod, snapshot);

            this._log('info', `========== ${opName}行操作完成 ==========`);

            return {
                success: true,
                arrData: arrData,
                oldTotalPeriods: oldTotalPeriods,
                newTotalPeriods: newTotalPeriods,
                modifiedRows: modifiedRows,
                operation: { type: operationType, position: position, rowCount: rowCount },
                repaymentMethod: repaymentMethod,
                errors: []
            };

        } catch (error) {
            const opName = operationType === 'insert' ? '插入' : '删除';
            this._log('error', `${opName}行操作失败: ${error.message}`);
            return {
                success: false,
                arrData: arrData,
                oldTotalPeriods: operationType === 'insert' ? arrData.length - rowCount : arrData.length + rowCount,
                newTotalPeriods: arrData.length,
                modifiedRows: [],
                operation: { type: operationType, position: position, rowCount: rowCount },
                repaymentMethod: repaymentMethod,
                errors: [error.message]
            };
        }
    }
    
    // ============== 辅助方法（模板方法的组成部分） ==============
    
    /**
     * _mergeOptions - 合并默认选项
     */
    _mergeOptions(options) {
        const defaultOptions = {
            autoUpdateTotalPeriods: true,
            syncToWorksheet: true,
            recalculateSheet: true,
            allowManualEdit: true
        };
        return { ...defaultOptions, ...options };
    }

    /**
     * _saveSnapshot - 保存操作前快照（用于撤销恢复）
     */
    _saveSnapshot(arrData, repaymentMethod) {
        return {
            arrDataCopy: JSON.parse(JSON.stringify(arrData)),
            repaymentMethod: repaymentMethod,
            totalPeriods: this.p.val("TotalPeriods"),
            worksheetBackup: this._backupEnabled ? this.backupWorksheetData() : null
        };
    }

    /**
     * _syncToWorksheet - 同步到工作表（处理 insert/delete 差异）
     */
    _syncToWorksheet(arrData, operationType, oldTotalPeriods, newTotalPeriods, position, rowCount, repaymentMethod) {
        const rng = this.arrDataToDataRange(arrData);
        if (!rng) {
            throw new Error("同步到工作表失败");
        }

        this.styleManager.applyDataFormat(rng);
        设置表格样式(rng);

        // 本金比例列样式（公共）
        if (repaymentMethod.includes("本金比例")) {
            this.styleManager.addBorder(rng.Columns(12));
            this.styleManager.setBackColor(rng.Columns(12), this.p.m_COLOR_YELLOW);
        }

        // 删除操作特有：清除多余行区域
        if (operationType === 'delete') {
            const rentTableStartRow = this.p.RentTableStartRow;
            const m_COL_PERIOD = this.p.m_COL_PERIOD;
            const m_COL_PRINCIPAL_RATIO = this.p.m_COL_PRINCIPAL_RATIO;
            const m_worksheet = this.p.m_worksheet;
            const newEndRow = rentTableStartRow + newTotalPeriods;
            const oldEndRow = rentTableStartRow + oldTotalPeriods;

            if (newEndRow < oldEndRow) {
                const clearRange = m_worksheet.Range(
                    `${m_COL_PERIOD}${newEndRow}:${m_COL_PRINCIPAL_RATIO}${oldEndRow}`
                );
                clearRange.Clear();
            }
        }

        // 插入操作特有：高亮新增行
        if (operationType === 'insert') {
            const m_COL_PERIOD = this.p.m_COL_PERIOD;
            const rentTableStartRow = this.p.RentTableStartRow;
            for (var i = 0; i < rowCount; i++) {
                const rowIdx = position + i;
                const cell = this.p.m_worksheet.Range(`${m_COL_PERIOD}${rentTableStartRow + rowIdx}`);
                this.styleManager.setBackColor(cell, this.p.m_COLOR_YELLOW);
            }
        }

        // 合计行（公共）
        this.租金测算表合计行(1, 10);
        if (repaymentMethod.includes("本金比例")) {
            this.租金测算表合计行(12, 12);
        }
    }

    /**
     * _registerUndoCommand - 注册撤销命令到统一管理器
     * 
     * 设计：使用 clsUndoableCommand 适配器将行操作封装为标准命令
     * 撤销时优先使用工作表快照恢复，确保数据完整性
     */
    _registerUndoCommand(operationType, position, rowCount, repaymentMethod, snapshot) {
        if (!this.m_undoEnabled || !this.m_undoManager) {
            return;
        }
        
        const self = this;
        const opName = operationType === 'insert' ? '插入' : '删除';
        
        const command = new clsUndoableCommand(
            this,
            operationType === 'insert' ? UNDO_CONFIG.OPERATION_TYPES.ROW_INSERT : UNDO_CONFIG.OPERATION_TYPES.ROW_DELETE,
            `${opName}${rowCount}行（位置${position}）`,
            function() { return true; }, // 操作已执行，直接返回成功
            function() {
                // 撤销：优先使用快照恢复
                console.log(`[${self.MODULE_NAME}] 撤销${opName}操作，恢复快照...`); // 注：闭包内无法用 _log
                if (snapshot.worksheetBackup) {
                    const result = self.restoreWorksheetData(snapshot.worksheetBackup);
                    if (result) {
                        // 恢复 TotalPeriods 参数
                        self.p.m_worksheet.Range(self.p.addr("TotalPeriods")).Value2 = snapshot.totalPeriods;
                    }
                    return result;
                }
                console.warn(`[${self.MODULE_NAME}] 无快照可用，撤销失败`); // 注：闭包内无法用 _log
                return false;
            },
            { operationType: operationType, position: position, rowCount: rowCount, repaymentMethod: repaymentMethod }
        );
        
        this.m_undoManager.execute(command);
        this._log('info', `已注册撤销命令: ${opName}${rowCount}行`);
    }

    /**
     * updateRowFormulas - 更新指定行的公式
     */
    updateRowFormulas(arrData, rowIndex, formulaType, arrFormula) {
        try {
            if (!arrData || !Array.isArray(arrData)) {
                throw new Error("数据数组无效");
            }
            if (rowIndex < 0 || rowIndex >= arrData.length) {
                throw new Error(`行索引越界: ${rowIndex}`);
            }
            if (!arrFormula || !arrFormula[formulaType]) {
                throw new Error("公式模板无效");
            }

            const maxCol = arrFormula[FORMULA_ROW.FIRST].length - 1;

            for (var col = 1; col <= maxCol && col < arrFormula[FORMULA_ROW.FIRST].length; col++) {
                arrData[rowIndex][col - 1] = arrFormula[formulaType][col];
            }

            return arrData;
        } catch (error) {
            this._log('error', `更新行公式失败: ${error.message}`);
            return null;
        }
    }

    /**
     * updateRemainingBalanceFormula - 更新剩余租金余额公式
     */
    updateRemainingBalanceFormula(arrData) {
        try {
            if (!arrData || !Array.isArray(arrData)) {
                throw new Error("数据数组无效");
            }

            const rentTableStartRow = this.p.RentTableStartRow;
            const totalPeriodsValue = arrData.length;
            const newEndRow = rentTableStartRow + totalPeriodsValue - 1;

            this._log('debug', `更新剩余租金余额公式范围: R${rentTableStartRow}C3:R${newEndRow}C3`);

            for (var row = 0; row < arrData.length; row++) {
                if (arrData[row] && arrData[row].length > 7) {
                    const currentFormula = arrData[row][7];
                    
                    if (typeof currentFormula === 'string' && currentFormula.includes('SUM(') && currentFormula.includes('C3')) {
                        arrData[row][7] = currentFormula.replace(
                            /SUM\(R\d+C3:R(\d+)C3\)/,
                            `SUM(R${rentTableStartRow}C3:R${newEndRow}C3)`
                        );
                    }
                }
            }

            return arrData;
        } catch (error) {
            this._log('error', `更新剩余租金余额公式失败: ${error.message}`);
            return null;
        }
    }

    /**
     * redistributePrincipalRatio - 重新分配本金比例
     */
    redistributePrincipalRatio(arrData, newTotalPeriods) {
        try {
            if (!arrData || !Array.isArray(arrData)) {
                throw new Error("数据数组无效");
            }
            if (newTotalPeriods <= 0) {
                throw new Error("总期数必须大于0");
            }

            const avgRatio = 100 / newTotalPeriods;

            for (var row = 0; row < arrData.length; row++) {
                if (arrData[row] && arrData[row].length > 11) {
                    arrData[row][11] = `=round(${avgRatio.toFixed(2)},2)`;
                }
            }

            const lastRowIndex = arrData.length - 1;
            if (arrData[lastRowIndex] && arrData[lastRowIndex].length > 11) {
                arrData[lastRowIndex][11] = `=100-SUM(R[${-newTotalPeriods + 1}]C:R[-1]C)`;
            }

            return arrData;
        } catch (error) {
            this._log('error', `重新分配本金比例失败: ${error.message}`);
            return null;
        }
    }

    /**
     * updateTotalPeriods - 更新TotalPeriods参数
     */
    updateTotalPeriods(newTotal) {
        try {
            if (!this.p || !this.p.m_worksheet) {
                throw new Error("参数管理器或工作表未初始化");
            }

            const totalPeriodsCell = this.p.addr("TotalPeriods");
            this.p.m_worksheet.Range(totalPeriodsCell).Value2 = newTotal;

            this._log('info', `TotalPeriods参数已更新为: ${newTotal}`);
            return true;
        } catch (error) {
            this._log('error', `更新TotalPeriods参数失败: ${error.message}`);
            return false;
        }
    }

    /**
     * shouldRegenerateFormulas - 判断是否需要重新生成公式
     */
    shouldRegenerateFormulas(repaymentMethod) {
        if (repaymentMethod === "等额本息（后付）" || repaymentMethod === "等额本息（先付）") {
            return false;
        }
        return true;
    }

    /**
     * getFormulaType - 确定行的公式类型
     */
    getFormulaType(rowIndex, totalPeriods) {
        if (rowIndex === 0) {
            return FORMULA_ROW.FIRST;
        } else if (rowIndex === totalPeriods - 1) {
            return FORMULA_ROW.LAST;
        } else {
            return FORMULA_ROW.MIDDLE;
        }
    }

    /**
     * identifyAffectedRows - 识别受影响的行
     */
    identifyAffectedRows(oldTotalPeriods, newTotalPeriods, operationPosition, operationType) {
        const affectedRows = new Set();

        if (operationType === 'insert') {
            if (operationPosition === 0) {
                affectedRows.add(0);
                for (var i = operationPosition + 1; i < newTotalPeriods; i++) {
                    affectedRows.add(i);
                }
            } else if (operationPosition >= oldTotalPeriods) {
                affectedRows.add(oldTotalPeriods - 1);
                affectedRows.add(newTotalPeriods - 1);
                for (var i = operationPosition; i < newTotalPeriods; i++) {
                    affectedRows.add(i);
                }
            } else if (operationPosition > 0 && operationPosition < oldTotalPeriods) {
                affectedRows.add(newTotalPeriods - 1);
                for (var i = operationPosition; i < newTotalPeriods; i++) {
                    affectedRows.add(i);
                }
            }
        }
        
        if (operationType === 'delete') {
            if (operationPosition === 0) {
                affectedRows.add(0);
            } else if (operationPosition >= oldTotalPeriods - 1) {
                affectedRows.add(newTotalPeriods - 1);
            } else if (operationPosition > 0 && operationPosition < oldTotalPeriods) {
                affectedRows.add(newTotalPeriods - 1);
            }
        }

        return Array.from(affectedRows).sort((a, b) => a - b);
    }

    /**
     * validateInput - 验证输入参数
     */
    validateInput(arrData, position, rowCount, operationType) {
        if (!arrData || !Array.isArray(arrData)) {
            return { valid: false, error: "数据数组无效或为空" };
        }
        
        if (arrData.length === 0) {
            return { valid: false, error: "数据数组不能为空" };
        }

        if (position < 0) {
            return { valid: false, error: "操作位置不能为负数" };
        }

        if (rowCount <= 0) {
            return { valid: false, error: "操作行数必须大于0" };
        }

        if (operationType === 'insert') {
            if (position > arrData.length) {
                return { valid: false, error: `插入位置 ${position} 超出范围（0-${arrData.length}）` };
            }
        }

        if (operationType === 'delete') {
            if (position >= arrData.length) {
                return { valid: false, error: `删除位置 ${position} 超出范围（0-${arrData.length - 1}）` };
            }
            if (position + rowCount > arrData.length) {
                return { valid: false, error: `删除行数超出范围，最多可删除 ${arrData.length - position} 行` };
            }
            if (position === 0 && rowCount === arrData.length) {
                return { valid: false, error: "不能删除所有行" };
            }
        }

        return { valid: true, error: "" };
    }

    /**
     * logOperation - 记录操作日志（兼容方法，委托给 clsErrorHandler）
     */
    logOperation(operationInfo) {
        try {
            const timestamp = new Date().toLocaleString('zh-CN');
            this._log('info', `操作日志: 类型=${operationInfo.type}, 位置=${operationInfo.position}, 行数=${operationInfo.rowCount}, 成功=${operationInfo.success}`);
        } catch (error) {
            this._log('error', `记录操作日志失败: ${error.message}`);
        }
    }

    /**
     * getOperationHistory - 获取操作历史（委托给统一撤销管理器）
     */
    getOperationHistory() {
        return this.m_undoManager ? this.m_undoManager.getUndoHistory() : [];
    }
    
    /**
     * undoLastOperation - 撤销上一次操作（委托给统一撤销管理器）
     * 
     * 改进：原实现使用独立的 operationHistory，撤销删除操作只能插入空行。
     * 现在委托给 clsUndoManager，使用快照恢复，可完整恢复原始数据。
     */
    undoLastOperation() {
        return this.undo();
    }
    
    /**
     * _getCurrentDataFromWorksheet - 从工作表获取当前数据
     */
    _getCurrentDataFromWorksheet() {
        try {
            const rentTableStartRow = this.p.RentTableStartRow;
            const totalPeriodsValue = this.p.val("TotalPeriods");
            const m_COL_PERIOD = this.p.m_COL_PERIOD;
            const m_COL_PRINCIPAL_RATIO = this.p.m_COL_PRINCIPAL_RATIO;
            const m_worksheet = this.p.m_worksheet;
            
            const dataRange = m_worksheet.Range(
                `${m_COL_PERIOD}${rentTableStartRow}:${m_COL_PRINCIPAL_RATIO}${rentTableStartRow + totalPeriodsValue - 1}`
            );
            
            return dataRange.Value2;
        } catch (error) {
            this._log('error', `获取当前数据失败: ${error.message}`);
            return [];
        }
    }
    
    /**
     * clearOperationHistory - 清空操作历史（委托给统一撤销管理器）
     */
    clearOperationHistory() {
        this.clearUndoHistory();
    }
    
    /**
     * getUndoInfo - 获取可撤销操作的信息（委托给统一撤销管理器）
     */
    getUndoInfo() {
        if (this.m_undoManager) {
            return this.m_undoManager.getNextUndoInfo();
        }
        return null;
    }
}

// ============== 便捷函数 ==============

/**
 * 创建租金测算实例（带撤销支持）
 * @param {string} sheetName - 工作表名称
 * @param {clsUndoManager} undoManager - 撤销管理器（可选）
 * @returns {clsRentalCalculation} 租金测算实例
 */
function createRentalCalculationWithUndo(sheetName, undoManager) {
    // 依赖注入：创建独立的参数管理器供独立调用使用
    const paramManager = new clsParameterManager();
    paramManager.Initialize(sheetName || "1租金测算表V1");
    const calc = new clsRentalCalculation(paramManager, undoManager || getUndoManager());
    calc.Initialize(paramManager);
    return calc;
}

/**
 * 创建数据区域（带撤销支持）
 * @returns {boolean} 是否成功
 */
function generateRentalTableWithUndo(sheetName) {
    const calc = createRentalCalculationWithUndo(sheetName);
    return calc.createDataRangeWithUndo();
}

/**
 * 清除数据（带撤销支持）
 * @returns {boolean} 是否成功
 */
function clearRentalTableWithUndo(sheetName) {
    const calc = createRentalCalculationWithUndo(sheetName);
    return calc.清除原有表中数据WithUndo();
}

/**
 * 使用每期适用利率生成测算表（带撤销支持）
 * @param {string} sheetName - 工作表名称
 * @returns {boolean} 是否成功
 */
function generatePeriodRateTableWithUndo(sheetName) {
    const calc = createRentalCalculationWithUndo(sheetName);
    return calc.使用每期适用利率生成测算表WithUndo();
}

/**
 * 修改自定义支付日（带撤销支持）
 * @param {number} period - 期次
 * @param {number} value - 值
 * @param {string} sheetName - 工作表名称
 * @returns {boolean} 是否成功
 */
function changeCustomPaymentDayWithUndo(period, value, sheetName) {
    const calc = createRentalCalculationWithUndo(sheetName);
    return calc.改变自定义支付日WithUndo(period, value);
}

/**
 * 显示租金测算模块的撤销历史
 */
function showRentalUndoHistory() {
    const calc = createRentalCalculationWithUndo("1租金测算表V1");
    const history = calc.getUndoHistory();
    
    console.log("========== 租金测算撤销历史 ==========");
    
    if (history.length === 0) {
        console.log("暂无操作历史");
    } else {
        history.forEach(function(item) {
            console.log(`${item.index + 1}. [${item.info.type}] ${item.info.description}`);
        });
    }
    
    console.log(`\n可撤销：${calc.canUndo() ? "是" : "否"}`);
    console.log(`可重做：${calc.canRedo() ? "是" : "否"}`);
    console.log("======================================");
}

console.log("[mRentalCalculation.js] 租金测算模块加载完成（已集成统一撤销管理器）");
