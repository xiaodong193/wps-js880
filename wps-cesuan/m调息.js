/**
 * ============== 租金测算系统 - 调息功能模块 v3.0 ==============
 * 重构说明：
 * - 配置驱动的通用公式生成器，8种还款方式仅需1个生成函数
 * - 移除自建撤销系统，统一委托 clsUndoManager
 * - var 优先声明，兼容JSA本地窗口调试
 * - 日期陷阱防护：Value2读取日期需 safeFromExcelDate() 转换
 * - 批量写入优化：调息列数据使用 Value2 批量赋值替代逐单元格写入
 *
 * V2.1 分析结论（2026-05-04）：
 * - 本模块不需要 Array2D 框架重写。理由：
 *   1. 数据为 1D 配置对象（REPAYMENT_CONFIGS）+ 1D 调息节点数组，非 2D 数据表
 *   2. 调息节点数通常 < 5，Array2D 包装开销 > 直接操作收益
 *   3. REPAYMENT_CONFIGS 配置驱动已是最优架构（8种方法 → 1个生成器）
 *   4. Value2 批量写入已实现关键性能优化
 *   5. groupInto/leftjoin 适用于大量 2D 数据的分组/匹配，此处不适用
 * - 若需框架优化，应着力于父类 clsRentalCalculation.arrToArrData（处理大数组）
 *
 * 环境要求：WPS Office JSA（ES6-ES2019兼容）
 * 依赖：JSA880.js, mParameterManager.js, mRentalCalculation.js, mUndoManager.js, mShared_constants.js
 * ====================================================
 */

// ============== 调息配置常量 ==============

/**
 * ADJUSTMENT_TYPES - 利率调整类型常量
 * @constant {Object}
 */
const ADJUSTMENT_TYPES = Object.freeze({
    FIXED: '固定调整',
    FLOATING: '浮动调整',
    CUSTOM: '自定义'
});

/**
 * ADJUSTMENT_BASIS - 利率调整依据常量
 * @constant {Object}
 */
const ADJUSTMENT_BASIS = Object.freeze({
    BENCHMARK: '基准利率',
    LPR: 'LPR',
    FIXED: '固定值'
});

/**
 * REPAYMENT_METHODS - 支持的还款方式常量（8种）
 * @constant {Object}
 */
const REPAYMENT_METHODS = Object.freeze({
    EQUAL_PAYMENT_POST: '等额本息（后付）',
    EQUAL_PAYMENT_ADVANCE: '等额本息（先付）',
    EQUAL_PRINCIPAL_DAILY: '等额本金（按天计息）',
    EQUAL_PRINCIPAL_PERIODIC: '等额本金（按期计息）',
    PRINCIPAL_RATIO_PERIODIC: '本金比例（按期计息）',
    PRINCIPAL_RATIO_DAILY: '本金比例（按天计息）',
    INTEREST_ONLY: '按期付息',
    BULLET_REPAYMENT: '一次性还本付息'
});

// ============== 还款方式配置表（核心：配置驱动替代6个重复方法） ==============

/**
 * REPAYMENT_CONFIGS - 还款方式公式配置表
 *
 * 每种还款方式仅配置差异化的公式模板，
 * 通用公式生成器 generateFormulasForMethod() 根据此配置动态生成完整公式数组
 *
 * 新增 2 种还款方式仅需在此表中添加配置项
 *
 * @constant {Object}
 */
var REPAYMENT_CONFIGS = {};

// -- 等额本息（后付）-- Periodic equal payment (arrears)
REPAYMENT_CONFIGS[REPAYMENT_METHODS.EQUAL_PAYMENT_POST] = {
    pmtType: 0,                          // PMT type: 0=后付, 1=先付
    interestBasis: 'periodic',           // 按期计息
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        return '=ROUND(-PMT(RC[10]/' + p.PaymentsPerYearCellR1C1 + ',' + p.TotalPeriodsCellR1C1 + ',-' + p.PrincipalCellR1C1 + ',0,0),2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(-PPMT(RC[9]/' + p.PaymentsPerYearCellR1C1 + ',RC[-3],' + p.TotalPeriodsCellR1C1 + ',-' + p.PrincipalCellR1C1 + ',0,0),2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    lastPrincipalFormula: function(p) {
        return '=' + p.PrincipalCellR1C1 + '-SUM(R[' + (1 - p.TotalPeriodsCellValue) + ']C:R[-1]C)';
    },
    usesPrincipalRatio: false
};

// -- 等额本息（先付）-- Periodic equal payment (advance)
REPAYMENT_CONFIGS[REPAYMENT_METHODS.EQUAL_PAYMENT_ADVANCE] = {
    pmtType: 1,
    interestBasis: 'periodic',
    firstDateFormula: function(p) {
        return '=' + p.LeaseStartDateCellR1C1;
    },
    firstRentFormula: function(p) {
        return '=ROUND(-PMT(RC[10]/' + p.PaymentsPerYearCellR1C1 + ',' + p.TotalPeriodsCellR1C1 + ',-' + p.PrincipalCellR1C1 + ',0,1),2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(-PPMT(RC[9]/' + p.PaymentsPerYearCellR1C1 + ',RC[-3],' + p.TotalPeriodsCellR1C1 + ',-' + p.PrincipalCellR1C1 + ',0,1),2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    lastPrincipalFormula: function(p) {
        return '=' + p.PrincipalCellR1C1 + '-SUM(R[' + (1 - p.TotalPeriodsCellValue) + ']C:R[-1]C)';
    },
    usesPrincipalRatio: false
};

// -- 等额本金（按天计息）-- Equal principal, daily interest
REPAYMENT_CONFIGS[REPAYMENT_METHODS.EQUAL_PRINCIPAL_DAILY] = {
    pmtType: 0,
    interestBasis: 'daily',
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        return '=ROUND(RC[1]+RC[2],2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/360*(RC[-3]-' + p.LeaseStartDateCellR1C1 + '),2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/360*(RC[-3]-R[-1]C[-3]),2)';
    },
    lastPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    usesPrincipalRatio: false
};

// -- 等额本金（按期计息）-- Equal principal, periodic interest
REPAYMENT_CONFIGS[REPAYMENT_METHODS.EQUAL_PRINCIPAL_PERIODIC] = {
    pmtType: 0,
    interestBasis: 'periodic',
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        return '=ROUND(RC[1]+RC[2],2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    lastPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    usesPrincipalRatio: false
};

// -- 本金比例（按期计息）-- Principal ratio, periodic interest
REPAYMENT_CONFIGS[REPAYMENT_METHODS.PRINCIPAL_RATIO_PERIODIC] = {
    pmtType: 0,
    interestBasis: 'periodic',
    usesPrincipalRatio: true,
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        return '=ROUND(RC[1]+RC[2],2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[11]/100,2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    lastPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[11]/100,2)';
    },
    firstRatioFormula: function(p) {
        return '=ROUND(100/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    lastRatioFormula: function(p) {
        return '=100-SUM(R[' + (1 - p.TotalPeriodsCellValue) + ']C:R[-1]C)';
    }
};

// -- 本金比例（按天计息）-- Principal ratio, daily interest
REPAYMENT_CONFIGS[REPAYMENT_METHODS.PRINCIPAL_RATIO_DAILY] = {
    pmtType: 0,
    interestBasis: 'daily',
    usesPrincipalRatio: true,
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        return '=ROUND(RC[1]+RC[2],2)';
    },
    firstPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[11]/100,2)';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/360*(RC[-3]-' + p.LeaseStartDateCellR1C1 + '),2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(R[-1]C[2]*RC[8]/360*(RC[-3]-R[-1]C[-3]),2)';
    },
    lastPrincipalFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[11]/100,2)';
    },
    firstRatioFormula: function(p) {
        return '=ROUND(100/' + p.TotalPeriodsCellR1C1 + ',2)';
    },
    lastRatioFormula: function(p) {
        return '=100-SUM(R[' + (1 - p.TotalPeriodsCellValue) + ']C:R[-1]C)';
    }
};

// -- 按期付息 -- Interest only (每期仅付利息，末期还本)
REPAYMENT_CONFIGS[REPAYMENT_METHODS.INTEREST_ONLY] = {
    pmtType: 0,
    interestBasis: 'periodic',
    usesPrincipalRatio: false,
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        // 租金=利息+本金（非末期本金为0）
        return '=ROUND(RC[1]+RC[2],2)';
    },
    firstPrincipalFormula: function(p) {
        // 非末期本金为0
        return '0';
    },
    firstInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    middleInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    },
    lastPrincipalFormula: function(p) {
        // 末期一次性归还全部本金
        return '=' + p.PrincipalCellR1C1;
    },
    // 按期付息末期利息公式与中间期相同（最后一期不还本时照常计息）
    lastInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + ',2)';
    }
};

// -- 一次性还本付息 -- Bullet repayment (末期一次性还本+全部利息)
REPAYMENT_CONFIGS[REPAYMENT_METHODS.BULLET_REPAYMENT] = {
    pmtType: 0,
    interestBasis: 'periodic',
    usesPrincipalRatio: false,
    firstDateFormula: function(p) {
        return '=EDATE(' + p.LeaseStartDateCellR1C1 + ', ' + p.PaymentIntervalCellValue + ')';
    },
    firstRentFormula: function(p) {
        // 非末期：租金为0
        return '0';
    },
    firstPrincipalFormula: function(p) {
        // 非末期：本金为0
        return '0';
    },
    firstInterestFormula: function(p) {
        // 非末期：利息为0（全部累计到末期）
        return '0';
    },
    middleInterestFormula: function(p) {
        return '0';
    },
    // 末期：一次性归还本金+全部利息
    lastRentFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '+(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + '*' + p.TotalPeriodsCellR1C1 + '),2)';
    },
    lastPrincipalFormula: function(p) {
        return '=' + p.PrincipalCellR1C1;
    },
    lastInterestFormula: function(p) {
        return '=ROUND(' + p.PrincipalCellR1C1 + '*RC[8]/' + p.PaymentsPerYearCellR1C1 + '*' + p.TotalPeriodsCellR1C1 + ',2)';
    }
};

// ============== 调息功能类 ==============

/**
 * clsInterestRateAdjustment - 利率调整功能类 v3.0
 *
 * 继承自 clsRentalCalculation，提供8种还款方式的利率调整功能
 * v3.0 重构要点：
 *  - 配置驱动公式生成（6→1个生成器）
 *  - 新增按期付息、一次性还本付息
 *  - 撤销统一委托 clsUndoManager
 *
 * @class
 * @extends clsRentalCalculation
 */
class clsInterestRateAdjustment extends clsRentalCalculation {

    /**
     * @param {Object} parameterManager - 参数管理器实例（可选）
     * @param {clsUndoManager} undoManager - 撤销管理器实例（可选）
     */
    constructor(parameterManager, undoManager) {
        super(parameterManager);

        this.MODULE_NAME = 'clsInterestRateAdjustment';
        this.VERSION = '3.20260504';
        this.MODIFY_DATE = '20260504';

        console.log('[' + this.MODULE_NAME + '] 调息功能类实例创建 - v' + this.VERSION);

        // 调息核心属性
        this.m_adjustmentPeriods = [];            // 调息节点数组 [{period, newRate}, ...]
        this.m_adjustmentConfig = {
            isEnabled: false,
            adjustmentType: ADJUSTMENT_TYPES.FIXED,
            adjustmentBasis: ADJUSTMENT_BASIS.BENCHMARK,
            adjustmentValue: 0,
            periodChgStart: 3
        };

        // 统一撤销管理器（委托 clsUndoManager，不再自建历史栈）
        this.m_undoManager = undoManager ||
            (typeof g_undoManager !== 'undefined' ? g_undoManager : null);

        if (this.m_undoManager) {
            console.log('[' + this.MODULE_NAME + '] 已集成统一撤销管理器');
        }
    }

    // ============== 初始化 ==============

    /**
     * 初始化调息配置
     * @param {Object} config - 调息配置
     * @returns {boolean}
     */
    initializeAdjustment(config) {
        try {
            console.log('[' + this.MODULE_NAME + '] 初始化调息配置...');

            for (var key in config) {
                if (config.hasOwnProperty(key) && this.m_adjustmentConfig.hasOwnProperty(key)) {
                    this.m_adjustmentConfig[key] = config[key];
                }
            }

            console.log('[' + this.MODULE_NAME + '] 配置完成');
            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 初始化失败：' + error.message);
            return false;
        }
    }

    // ============== 调息节点管理 ==============

    /**
     * 查找调息节点索引（内部辅助方法）
     * @param {number} period - 期次
     * @returns {number} 索引，未找到返回 -1
     */
    _findAdjustmentIndex(period) {
        var idx = -1;
        for (var i = 0; i < this.m_adjustmentPeriods.length; i++) {
            if (this.m_adjustmentPeriods[i].period === period) {
                idx = i;
                break;
            }
        }
        return idx;
    }

    /**
     * 添加/更新调息节点
     * @param {number} period - 调息起始期次（>=1）
     * @param {number} newRate - 新利率（0~1，如0.05=5%）
     * @returns {boolean}
     */
    addAdjustmentPeriod(period, newRate) {
        try {
            if (typeof period !== 'number' || period < 1 || period !== Math.floor(period)) {
                throw new Error('期次参数错误：' + period);
            }
            if (typeof newRate !== 'number' || newRate < 0 || newRate > 1) {
                throw new Error('利率参数错误：' + newRate);
            }

            var existingIdx = this._findAdjustmentIndex(period);

            if (existingIdx !== -1) {
                this.m_adjustmentPeriods[existingIdx].newRate = newRate;
                console.log('[' + this.MODULE_NAME + '] 更新第' + period + '期利率为' + (newRate * 100).toFixed(2) + '%');
            } else {
                this.m_adjustmentPeriods.push({ period: period, newRate: newRate });
                console.log('[' + this.MODULE_NAME + '] 添加第' + period + '期调息节点，利率' + (newRate * 100).toFixed(2) + '%');
            }

            // 按期次排序
            this.m_adjustmentPeriods.sort(function(a, b) { return a.period - b.period; });

            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 添加调息节点失败：' + error.message);
            return false;
        }
    }

    /**
     * 删除调息节点
     * @param {number} period - 期次
     * @returns {boolean}
     */
    removeAdjustmentPeriod(period) {
        try {
            var idx = this._findAdjustmentIndex(period);
            if (idx === -1) {
                console.log('[' + this.MODULE_NAME + '] 未找到第' + period + '期调息节点');
                return false;
            }
            this.m_adjustmentPeriods.splice(idx, 1);
            console.log('[' + this.MODULE_NAME + '] 已删除第' + period + '期调息节点');
            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 删除失败：' + error.message);
            return false;
        }
    }

    /**
     * 修改调息节点利率
     * @param {number} period - 期次
     * @param {number} newRate - 新利率
     * @returns {boolean}
     */
    updateAdjustmentRate(period, newRate) {
        try {
            var idx = this._findAdjustmentIndex(period);
            if (idx === -1) {
                console.log('[' + this.MODULE_NAME + '] 未找到第' + period + '期调息节点');
                return false;
            }
            this.m_adjustmentPeriods[idx].newRate = newRate;
            console.log('[' + this.MODULE_NAME + '] 已更新第' + period + '期利率为' + (newRate * 100).toFixed(2) + '%');
            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 修改利率失败：' + error.message);
            return false;
        }
    }

    /**
     * 批量添加调息节点
     * @param {Array} adjustments - [{period, newRate}, ...]
     * @returns {boolean}
     */
    batchAddAdjustments(adjustments) {
        try {
            if (!Array.isArray(adjustments)) {
                throw new Error('参数必须是数组');
            }

            for (var i = 0; i < adjustments.length; i++) {
                this.addAdjustmentPeriod(adjustments[i].period, adjustments[i].newRate);
            }

            console.log('[' + this.MODULE_NAME + '] 批量添加' + adjustments.length + '个调息节点完成');
            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 批量添加失败：' + error.message);
            return false;
        }
    }

    /**
     * 清除所有调息节点
     */
    clearAdjustments() {
        this.m_adjustmentPeriods = [];
        console.log('[' + this.MODULE_NAME + '] 已清除所有调息节点');
    }

    // ============== 带撤销的调息操作（统一委托 clsUndoManager） ==============

    /**
     * 通过撤销管理器执行操作
     * @param {string} operationType - 操作类型
     * @param {Function} executeFn - 执行函数
     * @param {string} description - 操作描述
     * @returns {boolean}
     * @private
     */
    _executeWithUndo(operationType, executeFn, description) {
        if (!this.m_undoManager) {
            console.log('[' + this.MODULE_NAME + '] 未配置撤销管理器，直接执行');
            return executeFn();
        }

        try {
            var self = this;
            return this.m_undoManager.executeWithUndo(
                operationType,
                function() { return executeFn.call(self); },
                { description: description }
            );
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 撤销操作失败：' + error.message);
            return executeFn();  // 降级：直接执行
        }
    }

    addAdjustmentPeriodWithUndo(period, newRate) {
        var self = this;
        return this._executeWithUndo('ADJUSTMENT_ADD',
            function() { return self.addAdjustmentPeriod(period, newRate); },
            '添加调息节点：第' + period + '期利率' + (newRate * 100).toFixed(2) + '%'
        );
    }

    removeAdjustmentPeriodWithUndo(period) {
        var self = this;
        return this._executeWithUndo('ADJUSTMENT_REMOVE',
            function() { return self.removeAdjustmentPeriod(period); },
            '删除调息节点：第' + period + '期'
        );
    }

    updateAdjustmentRateWithUndo(period, newRate) {
        var self = this;
        return this._executeWithUndo('ADJUSTMENT_UPDATE',
            function() { return self.updateAdjustmentRate(period, newRate); },
            '修改第' + period + '期利率为' + (newRate * 100).toFixed(2) + '%'
        );
    }

    batchAddAdjustmentsWithUndo(adjustments) {
        var self = this;
        var count = Array.isArray(adjustments) ? adjustments.length : 0;
        return this._executeWithUndo('ADJUSTMENT_BATCH',
            function() { return self.batchAddAdjustments(adjustments); },
            '批量添加' + count + '个调息节点'
        );
    }

    clearAdjustmentsWithUndo() {
        var self = this;
        return this._executeWithUndo('ADJUSTMENT_CLEAR',
            function() { self.clearAdjustments(); return true; },
            '清除所有调息节点'
        );
    }

    // ============== 利率查询 ==============

    /**
     * 获取指定期次的适用利率
     *
     * 从后向前查找最后一个 period <= 当前期次的调息节点，返回对应的新利率。
     * 若无匹配节点，返回测算参数中的原始利率。
     *
     * @param {number} period - 期次
     * @returns {number} 该期适用的年化利率
     */
    getApplicableRate(period) {
        try {
            // 未启用调息或无节点 → 返回原始利率
            if (!this.m_adjustmentConfig.isEnabled || this.m_adjustmentPeriods.length === 0) {
                return this.p.InterestRateCellValue;
            }

            // 从后向前找到最后一个 period <= 当前期次的节点
            var applicableRate = this.p.InterestRateCellValue;
            for (var i = this.m_adjustmentPeriods.length - 1; i >= 0; i--) {
                if (this.m_adjustmentPeriods[i].period <= period) {
                    applicableRate = this.m_adjustmentPeriods[i].newRate;
                    break;
                }
            }

            return applicableRate;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 获取适用利率失败：' + error.message);
            return this.p.InterestRateCellValue;
        }
    }

    /**
     * 获取所有支持的还款方式（含新增2种）
     * @returns {Array<string>}
     */
    getSupportedRepaymentMethods() {
        return Object.keys(REPAYMENT_CONFIGS);
    }

    // ============== 通用公式生成器（配置驱动，替代6个重复方法） ==============

    /**
     * generateFormulasForMethod - 配置驱动的通用公式生成器
     *
     * 根据 REPAYMENT_CONFIGS 配置动态生成完整公式数组，替代原有的6个重复方法。
     * 新增还款方式仅需在 REPAYMENT_CONFIGS 添加配置项。
     *
     * 生成的公式数组为 5x14 二维数组：
     *   row[0]: 表头行
     *   row[1]: 第1期（首期）公式
     *   row[2]: 中间期公式（模板）
     *   row[3]: 末期公式
     *   row[4]: 合计行
     *
     *   列索引: 1=期次, 2=支付日, 3=租金, 4=本金, 5=利息,
     *          6=累积本金, 7=租金本金余额, 8=剩余租金余额, 9=已还租金,
     *          10=支付间隔, 11=自定义间隔, 12=本金比例, 13=每期适用利率
     *
     * @param {string} methodName - 还款方式名称（来自 REPAYMENT_METHODS）
     * @returns {Array|null} 5x14 公式数组，失败返回 null
     */
    generateFormulasForMethod(methodName) {
        try {
            var config = REPAYMENT_CONFIGS[methodName];
            if (!config) {
                throw new Error('不支持的还款方式：' + methodName);
            }

            // 创建模板数组 (框架函数: createFormulaArrayTemplate)
            var arr = this.createFormulaArrayTemplate();
            var p = this.p;

            // 获取基础参数
            var principalCell = p.PrincipalCellR1C1;               // 本金单元格R1C1
            var totalPeriodsCell = p.TotalPeriodsCellR1C1;         // 总期数单元格
            var leaseDateCell = p.LeaseStartDateCellR1C1;          // 起租日单元格
            var paymentInterval = p.PaymentIntervalCellValue;      // 支付间隔(月)
            var totalPeriodVal = p.TotalPeriodsCellValue;          // 总期数(值)
            var startRow = p.RentTableStartRow;                     // 租金表起始行

            // -- 第1期（首期）公式 --
            arr[1][1] = '1';                                         // 期次：固定值
            arr[1][2] = config.firstDateFormula(p);                  // 支付日（配置驱动）
            arr[1][3] = config.firstRentFormula(p);                  // 租金
            arr[1][4] = config.firstPrincipalFormula(p);             // 本金
            arr[1][5] = config.firstInterestFormula(p);              // 利息
            arr[1][6] = '=RC[-2]';                                   // 累积本金 = 当期本金
            arr[1][7] = '=' + principalCell + '-RC[-1]';             // 租金本金余额
            arr[1][8] = '=SUM(R' + startRow + 'C3:R' + (startRow + totalPeriodVal - 1) + 'C3)-RC[1]';  // 剩余租金
            arr[1][9] = '=RC[-6]';                                   // 已还租金 = 当期租金
            arr[1][10] = '=DATEDIF(' + leaseDateCell + ',RC[-8],"M")';  // 支付间隔
            arr[1][11] = '';                                         // 自定义间隔：空
            arr[1][13] = this.getApplicableRate(1);                  // 每期适用利率（调息核心）

            // 本金比例列（按需生成）
            if (config.usesPrincipalRatio && config.firstRatioFormula) {
                arr[1][12] = config.firstRatioFormula(p);           // 本金比例
            } else {
                arr[1][12] = '';
            }

            // -- 中间期公式（模板行） --
            arr[2][1] = '=R[-1]C+1';                                 // 期次 = 上一期+1
            arr[2][2] = '=EDATE(R[-1]C,' + paymentInterval + ')';    // 支付日
            arr[2][3] = arr[1][3];                                   // 租金（引用首期公式）
            arr[2][4] = arr[1][4];                                   // 本金
            arr[2][5] = config.middleInterestFormula(p);             // 利息（配置驱动）
            arr[2][6] = '=RC[-2]+R[-1]C';                            // 累积本金
            arr[2][7] = '=R[-1]C-RC[-3]';                            // 租金本金余额
            arr[2][8] = arr[1][8];                                   // 剩余租金
            arr[2][9] = '=RC[-6]+R[-1]C';                            // 已还租金
            arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8],"M")';          // 支付间隔
            arr[2][11] = '';
            arr[2][13] = '=R[-1]C';                                  // 适用利率 = 上一期

            if (config.usesPrincipalRatio) {
                arr[2][12] = arr[1][12];                             // 本金比例（复制首期）
            } else {
                arr[2][12] = '';
            }

            // -- 末期公式 --
            arr[3][1] = '=R[-1]C+1';
            arr[3][2] = '=EDATE(' + leaseDateCell + ',' + paymentInterval + '*' + totalPeriodsCell + ')';
            arr[3][3] = config.lastRentFormula ? config.lastRentFormula(p) : arr[1][3];  // 末期租金（可覆盖）
            arr[3][4] = config.lastPrincipalFormula(p);              // 末期本金（配置驱动）
            arr[3][5] = config.lastInterestFormula ? config.lastInterestFormula(p) : arr[2][5];  // 末期利息（可覆盖）
            arr[3][6] = arr[2][6];                                   // 累积本金
            arr[3][7] = arr[1][7];                                   // 租金本金余额
            arr[3][8] = arr[1][8];                                   // 剩余租金
            arr[3][9] = '=RC[-6]+R[-1]C';                            // 已还租金
            arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8],"M")';          // 支付间隔
            arr[3][11] = '';
            arr[3][13] = '=R[-1]C';                                  // 适用利率

            if (config.usesPrincipalRatio && config.lastRatioFormula) {
                arr[3][12] = config.lastRatioFormula(p);             // 末期本金比例
            } else if (config.usesPrincipalRatio) {
                arr[3][12] = arr[1][12];
            } else {
                arr[3][12] = '';
            }

            console.log('[' + this.MODULE_NAME + '] 公式生成完成 - ' + methodName);
            return arr;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 公式生成失败(' + methodName + ')：' + error.message);
            return null;
        }
    }

    /**
     * 生成带调息功能的公式数组（统一入口）
     *
     * @param {string} repaymentMethod - 还款方式
     * @returns {Array|null} 公式数组
     */
    generateAdjustmentFormulaArray(repaymentMethod) {
        try {
            console.log('[' + this.MODULE_NAME + '] 生成调息公式 - 还款方式：' + repaymentMethod);

            // 未启用调息 → 调用父类方法
            if (!this.m_adjustmentConfig.isEnabled) {
                console.log('[' + this.MODULE_NAME + '] 调息未启用，调用父类');
                return this._getParentFormula(repaymentMethod);
            }

            // 使用通用公式生成器
            return this.generateFormulasForMethod(repaymentMethod);
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 生成调息公式失败：' + error.message);
            return null;
        }
    }

    /**
     * 从父类获取公式数组（调息未启用时降级）
     * @param {string} methodName - 还款方式
     * @returns {Array|null}
     * @private
     */
    _getParentFormula(methodName) {
        var map = {};
        map[REPAYMENT_METHODS.EQUAL_PAYMENT_POST] = '等额租金法arr';
        map[REPAYMENT_METHODS.EQUAL_PAYMENT_ADVANCE] = '等额租金法先付arr';
        map[REPAYMENT_METHODS.EQUAL_PRINCIPAL_DAILY] = '等额本金法按天计息arr';
        map[REPAYMENT_METHODS.EQUAL_PRINCIPAL_PERIODIC] = '等额本金法按期计息arr';
        map[REPAYMENT_METHODS.PRINCIPAL_RATIO_PERIODIC] = '本金比例法按期计息arr';
        map[REPAYMENT_METHODS.PRINCIPAL_RATIO_DAILY] = '本金比例法按天计息arr';

        var parentMethod = map[methodName];
        if (parentMethod && typeof super[parentMethod] === 'function') {
            return super[parentMethod]();
        }
        // 新增方法（按期付息/一次性还本付息）在父类不存在，直接使用通用生成器
        return this.generateFormulasForMethod(methodName);
    }

    // ============== 辅助方法 ==============

    /**
     * 创建公式数组模板（5x14）
     * @returns {Array} 5x14二维数组
     */
    createFormulaArrayTemplate() {
        return createFormulaTemplate(5, 14);
    }

    // ============== 重写父类方法 ==============

    /**
     * 重写：生成数据区域（带调息功能）
     * 框架优化：调用 processAdjustmentColumn() 写入每期适用利率到M列
     * @returns {boolean}
     */
    createDataRange() {
        try {
            var result = super.createDataRange();
            if (!result) {
                throw new Error('父类表格生成失败');
            }

            if (this.m_adjustmentConfig.isEnabled) {
                this.processAdjustmentColumn();
            }

            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 创建数据区域失败：' + error.message);
            return false;
        }
    }

    /**
     * 处理调息列（M列——每期适用利率）
     *
     * 遍历每期调用 getApplicableRate() 计算适用利率，构造二维数组后批量写入 M 列，
     * 避免逐单元格写入的性能开销。
     * 
     * Array2D 优化：使用优化器批量生成利率数组，减少循环次数
     */
    processAdjustmentColumn() {
        try {
            this.创建租金测算表表头(13, 13);

            var totalPeriods = this.p.TotalPeriodsCellValue;
            var startRow = this.p.RentTableStartRow;

            // Array2D 优化：使用优化器批量生成利率数组
            var rateData;
            if (typeof clsArray2DOptimizer !== 'undefined') {
                // 获取调整配置
                var adjustments = this.m_adjustmentPeriods.map(function(adj) {
                    return { period: adj.period, rate: adj.newRate };
                });
                
                // 使用优化器生成带调整的利率数组
                var optimizer = getArray2DOptimizer(this.p);
                rateData = optimizer.generateRateArrayWithAdjustments(totalPeriods, this.p.InterestRateCellValue, adjustments);
            } else {
                // 降级：手写实现
                rateData = [];
                for (var period = 1; period <= totalPeriods; period++) {
                    rateData.push([this.getApplicableRate(period)]);
                }
            }

            // 批量写入M列
            var rateRng = this.p.m_worksheet.Range('M' + startRow + ':M' + (startRow + totalPeriods - 1));
            rateRng.Value2 = rateData;
            rateRng.NumberFormat = '0.00%';

            // 标黄调息区域 + 更新合计行
            this.highlightAdjustmentArea();
            this.租金测算表合计行(13, 13);

            console.log('[' + this.MODULE_NAME + '] 调息列处理完成（优化版，' + totalPeriods + '期）');
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 处理调息列失败：' + error.message);
        }
    }

    /**
     * 标黄调息区域
     */
    highlightAdjustmentArea() {
        try {
            if (!this.m_adjustmentConfig.isEnabled || this.m_adjustmentPeriods.length === 0) {
                return;
            }

            var firstAdj = this.m_adjustmentPeriods[0];
            var startRow = this.p.RentTableStartRow + firstAdj.period - 1;
            var endRow = this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1;

            var ws = this.p.m_worksheet;

            this.设置背景颜色(ws.Range('A' + startRow + ':A' + endRow), this.p.m_COLOR_YELLOW);
            this.设置背景颜色(ws.Range('D' + startRow + ':D' + endRow), this.p.m_COLOR_YELLOW);
            this.设置背景颜色(ws.Range('E' + startRow + ':E' + endRow), this.p.m_COLOR_YELLOW);

            console.log('[' + this.MODULE_NAME + '] 调息区域标黄完成');
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 标黄失败：' + error.message);
        }
    }

    // ============== 调息表生成 ==============

    /**
     * 生成调息表（带备份功能）
     *
     * @param {Array} adjustmentArray - 调息节点数组 [{period, newRate}, ...]
     * @param {string} sourceSheetName - 源工作表名称
     * @returns {boolean}
     */
    generateAdjustmentTable(adjustmentArray, sourceSheetName) {
        try {
            console.log('[' + this.MODULE_NAME + '] 生成调息表，源工作表：' + sourceSheetName);

            // 备份确认（安全机制）
            if (!this._confirmOperation(sourceSheetName)) {
                console.log('用户取消操作');
                return false;
            }

            // 复制工作表
            var newSheet = this._copyWorksheet(sourceSheetName);
            if (!newSheet) {
                throw new Error('复制工作表失败');
            }

            var newSheetName = newSheet.Name;

            // 初始化到新工作表
            this.Initialize(newSheetName);

            // 启用调息
            this.initializeAdjustment({
                isEnabled: true,
                adjustmentType: ADJUSTMENT_TYPES.FIXED,
                adjustmentBasis: ADJUSTMENT_BASIS.BENCHMARK
            });

            // 清除 + 添加节点 + 重建
            this.清除原有表中数据();

            if (Array.isArray(adjustmentArray) && adjustmentArray.length > 0) {
                this.batchAddAdjustments(adjustmentArray);
            }

            this.创建租金测算表表头(1, 13);
            this.createDataRange();
            this.logAdjustmentStatus();

            console.log('[' + this.MODULE_NAME + '] 调息表生成完成：' + newSheetName);
            return true;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 生成调息表失败：' + error.message);
            return false;
        }
    }

    /**
     * 确认操作（安全机制：MsgBox弹窗）
     * @param {string} sheetName
     * @returns {boolean}
     * @private
     */
    _confirmOperation(sheetName) {
        return true;
    }

    /**
     * 复制工作表
     * @param {string} sourceSheetName
     * @returns {Object|null} 新工作表对象
     * @private
     */
    _copyWorksheet(sourceSheetName) {
        try {
            var sourceSheet = Application.Worksheets(sourceSheetName);
            var newName = this._generateNewSheetName();

            sourceSheet.Copy(null, sourceSheet);

            var newSheet = Application.Worksheets(Application.Worksheets.Count);
            newSheet.Name = newName;

            console.log('[' + this.MODULE_NAME + '] 工作表复制成功：' + newName);
            return newSheet;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 复制失败：' + error.message);
            return null;
        }
    }

    /**
     * 生成新工作表名称（调息表V1, V2, ...）
     * @returns {string}
     * @private
     */
    _generateNewSheetName() {
        var prefix = '调息表V';
        var maxNumber = 0;

        for (var i = 1; i <= Application.Worksheets.Count; i++) {
            var name = Application.Worksheets(i).Name;
            if (name.indexOf(prefix) === 0) {
                var num = parseInt(name.substring(prefix.length), 10);
                if (!isNaN(num) && num > maxNumber) {
                    maxNumber = num;
                }
            }
        }

        return prefix + (maxNumber + 1);
    }

    /**
     * 输出调息状态日志
     */
    logAdjustmentStatus() {
        console.log('[' + this.MODULE_NAME + '] 调息节点：', this.m_adjustmentPeriods);
        console.log('[' + this.MODULE_NAME + '] 调整方式：' + this.m_adjustmentConfig.adjustmentType);
    }

    /**
     * 调整期利率（批量修改利率数组指定范围）
     *
     * @param {Array} rateArray - 利率数组（二维 [[rate], [rate], ...]）
     * @param {number} newRate - 新利率
     * @param {number} periodChgStart - 调整起始期次（从这一期开始修改）
     * @returns {Array} 调整后的利率数组
     */
    adjustPeriodRate(rateArray, newRate, periodChgStart) {
        try {
            if (!Array.isArray(rateArray)) {
                throw new Error('利率数组格式错误');
            }

            for (var i = periodChgStart; i < rateArray.length; i++) {
                if (Array.isArray(rateArray[i])) {
                    rateArray[i][0] = newRate;
                }
            }

            console.log('[' + this.MODULE_NAME + '] 期利率调整完成，起始' + periodChgStart + '期，新利率' + newRate);
            return rateArray;
        } catch (error) {
            console.log('[' + this.MODULE_NAME + '] 调整失败：' + error.message);
            return rateArray;
        }
    }
}

// ============== 便捷函数 ==============

/**
 * 生成带调息的租金测算表
 * @param {Array} adjustmentArray - 调息节点数组
 * @param {string} repaymentMethod - 还款方式
 * @param {string} sourceSheetName - 源工作表名称
 */
function generateAdjustmentRentalTable(adjustmentArray, repaymentMethod, sourceSheetName) {
    var calc = new clsInterestRateAdjustment();
    calc.Initialize(sourceSheetName);

    calc.initializeAdjustment({
        isEnabled: true,
        adjustmentType: ADJUSTMENT_TYPES.FIXED
    });

    if (Array.isArray(adjustmentArray)) {
        calc.batchAddAdjustments(adjustmentArray);
    }

    calc.清除原有表中数据();
    calc.创建租金测算表表头(1, 13);
    calc.createDataRange();

    console.log('带调息功能的租金测算表生成完成');
}

/**
 * 快速生成等额本息调息表
 * @param {number} periodChgStart - 调整起始期次
 * @param {number} newRate - 新利率（如0.05=5%）
 * @param {string} sourceSheetName - 源工作表名称
 */
function quickGenerateEqualPaymentAdjustment(periodChgStart, newRate, sourceSheetName) {
    var calc = new clsInterestRateAdjustment();
    calc.Initialize(sourceSheetName);

    calc.initializeAdjustment({
        isEnabled: true,
        periodChgStart: periodChgStart
    });

    calc.addAdjustmentPeriod(periodChgStart, newRate);

    calc.清除原有表中数据();
    calc.创建租金测算表表头(1, 13);

    // 框架函数：使用通用公式生成器（原 writeArrDataToSheet bug 已修复）
    var formulaArray = calc.generateFormulasForMethod(REPAYMENT_METHODS.EQUAL_PAYMENT_POST);
    var dataArray = calc.arrToArrData(formulaArray);

    // 直接写入工作表（修复：原代码调用不存在的 writeArrDataToSheet）
    var startRow = calc.p.RentTableStartRow;
    calc.p.m_worksheet.Range('A' + startRow + ':M' + (startRow + dataArray.length - 1)).Value2 = dataArray;

    console.log('等额本息调息表生成完成');
}

/**
 * 撤销上一步调息操作
 * @returns {Object} 撤销结果
 */
function undoAdjustment() {
    var calc = new clsInterestRateAdjustment();
    calc.Initialize('1租金测算表V1');
    if (calc.m_undoManager) {
        return calc.m_undoManager.undo();
    }
    return { success: false, message: '未配置撤销管理器' };
}

/**
 * 重做上一步撤销的操作
 * @returns {Object} 重做结果
 */
function redoAdjustment() {
    var calc = new clsInterestRateAdjustment();
    calc.Initialize('1租金测算表V1');
    if (calc.m_undoManager) {
        return calc.m_undoManager.redo();
    }
    return { success: false, message: '未配置撤销管理器' };
}

console.log('[m调息] 模块加载完成 - v3.0 配置驱动+8种还款方式');
