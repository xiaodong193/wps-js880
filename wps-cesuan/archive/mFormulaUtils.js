/**
 * ============== 公式工具模块 ==============
 * 作者：徐晓冬
 * 版本：V2.20260130
 * 描述：提供公共的公式生成工具函数，消除代码重复
 * 
 * 核心功能：
 * - 日期公式生成
 * - 利息计算公式生成
 * - R1C1单元格引用生成
 * - 公式验证和格式化
 * ====================================================
 */

// ============== 公式工具类 ==============
class clsFormulaUtils {
    constructor(parameterManager) {
        this.p = parameterManager;
        this.MODULE_NAME = "clsFormulaUtils";
    }
    
    /**
     * 生成日期递增公式
     * @param {string} baseDateRef - 基准日期单元格引用（R1C1格式）
     * @param {string|number} intervalRef - 间隔月数（单元格引用或数值）
     * @param {boolean} useRelativeRef - 是否使用相对引用（R[-1]C）
     * @returns {string} EDATE公式
     */
    generateDateFormula(baseDateRef, intervalRef, useRelativeRef = false) {
        if (useRelativeRef) {
            return `=EDATE(R[-1]C, ${intervalRef})`;
        }
        return `=EDATE(${baseDateRef}, ${intervalRef})`;
    }
    
    /**
     * 生成利息计算公式（按天计息）
     * @param {string} principalRef - 本金单元格引用
     * @param {string} rateRef - 利率单元格引用
     * @param {string} daysRef - 天数引用（可以是计算公式）
     * @param {number} decimals - 小数位数
     * @returns {string} ROUND公式
     */
    generateInterestFormula(principalRef, rateRef, daysRef, decimals = 2) {
        return `=ROUND(${principalRef}*${rateRef}/360*(${daysRef}),${decimals})`;
    }
    
    /**
     * 生成利息计算公式（按期计息）
     * @param {string} principalRef - 本金单元格引用
     * @param {string} rateRef - 利率单元格引用
     * @param {string} periodsRef - 期数引用
     * @param {number} paymentsPerYear - 每年支付次数
     * @param {number} decimals - 小数位数
     * @returns {string} ROUND公式
     */
    generatePeriodicInterestFormula(principalRef, rateRef, periodsRef, paymentsPerYear = 12, decimals = 2) {
        return `=ROUND(${principalRef}*${rateRef}/${paymentsPerYear}*${periodsRef},${decimals})`;
    }
    
    /**
     * 生成PMT租金计算公式
     * @param {string} rateRef - 利率单元格引用
     * @param {string} periodsRef - 总期数单元格引用
     * @param {string} principalRef - 本金单元格引用
     * @param {boolean} isAdvance - 是否先付
     * @returns {string} PMT公式
     */
    generatePMTFormula(rateRef, periodsRef, principalRef, isAdvance = false) {
        const type = isAdvance ? 1 : 0;
        return `=ROUND(PMT(${rateRef}/RC[8],${periodsRef},-${principalRef},0,${type}),2)`;
    }
    
    /**
     * 生成本金余额公式
     * @param {boolean} isFirstRow - 是否首行
     * @param {string} principalRef - 本金单元格引用（首行使用）
     * @returns {string} 本金余额公式
     */
    generatePrincipalBalanceFormula(isFirstRow = false, principalRef = null) {
        if (isFirstRow) {
            return `=${principalRef}-RC[-3]`;
        }
        return "=R[-1]C-RC[-3]";
    }
    
    /**
     * 生成剩余租金余额公式
     * @param {boolean} isFirstRow - 是否首行
     * @param {string} totalRentRef - 总租金单元格引用（首行使用）
     * @returns {string} 剩余租金余额公式
     */
    generateRentBalanceFormula(isFirstRow = false, totalRentRef = null) {
        if (isFirstRow) {
            return `=${totalRentRef}-RC[-5]`;
        }
        return "=R[-1]C-RC[-5]";
    }
    
    /**
     * 生成已还租金公式
     * @param {boolean} isFirstRow - 是否首行
     * @returns {string} 已还租金公式
     */
    generatePaidRentFormula(isFirstRow = false) {
        if (isFirstRow) {
            return "=RC[-5]";
        }
        return "=R[-1]C+RC[-5]";
    }
    
    /**
     * 生成累积偿还本金公式
     * @param {boolean} isFirstRow - 是否首行
     * @returns {string} 累积偿还本金公式
     */
    generateAccumulatedPrincipalFormula(isFirstRow = false) {
        if (isFirstRow) {
            return "=RC[-3]";
        }
        return "=R[-1]C+RC[-3]";
    }
    
    /**
     * 生成R1C1单元格引用
     * @param {number} rowOffset - 行偏移（相对于当前行）
     * @param {number} colOffset - 列偏移（相对于当前列）
     * @param {boolean} absoluteRow - 是否绝对行引用
     * @param {boolean} absoluteCol - 是否绝对列引用
     * @returns {string} R1C1引用字符串
     */
    generateR1C1Ref(rowOffset = 0, colOffset = 0, absoluteRow = false, absoluteCol = false) {
        const rowRef = absoluteRow ? `R${rowOffset}` : (rowOffset === 0 ? "R" : `R[${rowOffset}]`);
        const colRef = absoluteCol ? `C${colOffset}` : (colOffset === 0 ? "C" : `C[${colOffset}]`);
        
        // 如果都是相对引用且偏移为0，简化为 RC
        if (!absoluteRow && !absoluteCol && rowOffset === 0 && colOffset === 0) {
            return "RC";
        }
        
        return rowRef + colRef;
    }
    
    /**
     * 从参数管理器获取常用单元格引用
     * @returns {Object} 常用单元格引用对象
     */
    getCommonCellRefs() {
        if (!this.p) {
            throw new Error("参数管理器未设置");
        }
        
        return {
            principal: this.p.PrincipalCellR1C1,
            interestRate: this.p.InterestRateCellR1C1,
            totalPeriods: this.p.TotalPeriodsCellR1C1,
            paymentInterval: this.p.PaymentIntervalCellR1C1,
            paymentsPerYear: this.p.PaymentsPerYearCellR1C1,
            leaseStartDate: this.p.LeaseStartDateCellR1C1,
            firstPaymentDate: this.p.FirstPaymentDateCellR1C1
        };
    }
    
    /**
     * 验证公式语法
     * @param {string} formula - 公式字符串
     * @returns {Object} 验证结果
     */
    validateFormula(formula) {
        const result = {
            valid: true,
            errors: []
        };
        
        if (!formula || typeof formula !== 'string') {
            result.valid = false;
            result.errors.push("公式不能为空");
            return result;
        }
        
        // 检查公式是否以=开头
        if (!formula.startsWith("=") && !formula.startsWith("+")) {
            result.errors.push("公式应以 = 或 + 开头");
        }
        
        // 检查括号匹配
        var bracketCount = 0;
        for (const char of formula) {
            if (char === '(') bracketCount++;
            if (char === ')') bracketCount--;
            if (bracketCount < 0) {
                result.errors.push("括号不匹配：多余的右括号");
                break;
            }
        }
        if (bracketCount > 0) {
            result.errors.push("括号不匹配：缺少右括号");
        }
        
        // 检查常见错误模式
        if (formula.includes(",,")) {
            result.errors.push("公式中包含连续的逗号");
        }
        if (formula.includes("..")) {
            result.errors.push("公式中包含连续的点");
        }
        
        if (result.errors.length > 0) {
            result.valid = false;
        }
        
        return result;
    }
    
    /**
     * 创建数组行模板
     * @param {number} colCount - 列数
     * @param {*} defaultValue - 默认值
     * @returns {Array} 初始化后的数组
     */
    createRowTemplate(colCount, defaultValue = null) {
        return new Array(colCount).fill(defaultValue);
    }
    
    /**
     * 计算天数差公式
     * @param {string} endDateRef - 结束日期引用
     * @param {string} startDateRef - 开始日期引用
     * @returns {string} 天数差
     */
    generateDaysDiffFormula(endDateRef, startDateRef) {
        return `${endDateRef}-${startDateRef}`;
    }
    
    /**
     * 生成条件公式
     * @param {string} condition - 条件
     * @param {string} trueValue - 条件为真时的值
     * @param {string} falseValue - 条件为假时的值
     * @returns {string} IF公式
     */
    generateIfFormula(condition, trueValue, falseValue) {
        return `=IF(${condition},${trueValue},${falseValue})`;
    }
    
    /**
     * 生成最大值/最小值公式
     * @param {Array<string>} refs - 单元格引用数组
     * @param {string} type - "MAX" 或 "MIN"
     * @returns {string} MAX/MIN公式
     */
    generateMinMaxFormula(refs, type = "MAX") {
        const args = refs.join(",");
        return `=${type}(${args})`;
    }
}

// ============== 静态工具函数 ==============

/**
 * 创建公式数组模板
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @returns {Array} 二维数组
 */
function createFormulaArrayTemplate(rows, cols) {
    const arr = [];
    for (var i = 0; i < rows; i++) {
        arr[i] = new Array(cols);
    }
    return arr;
}

/**
 * 复制数组行
 * @param {Array} sourceRow - 源行
 * @returns {Array} 复制后的行
 */
function copyRow(sourceRow) {
    return [...sourceRow];
}

/**
 * 安全设置数组元素
 * @param {Array} arr - 数组
 * @param {number} index - 索引
 * @param {*} value - 值
 * @returns {boolean} 是否成功
 */
function safeSetArrayElement(arr, index, value) {
    if (!arr || !Array.isArray(arr)) {
        console.warn("safeSetArrayElement: 数组无效");
        return false;
    }
    if (index < 0 || index >= arr.length) {
        console.warn(`safeSetArrayElement: 索引 ${index} 超出范围 [0, ${arr.length})`);
        return false;
    }
    arr[index] = value;
    return true;
}

/**
 * 检查公式是否包含循环引用
 * @param {string} formula - 公式
 * @param {number} currentRow - 当前行号
 * @param {number} currentCol - 当前列号
 * @returns {boolean} 是否可能包含循环引用
 */
function checkCircularReference(formula, currentRow, currentCol) {
    // 简单检查：如果公式引用了 RC（当前单元格）
    if (formula.includes("RC") && !formula.includes("R[") && !formula.includes("R-")) {
        // 检查是否是纯 RC 或 R[0]C[0]
        const pureRcPattern = /R\[?0?\]?C\[?0?\]?/;
        if (pureRcPattern.test(formula)) {
            return true;
        }
    }
    return false;
}

console.log("[mFormulaUtils.js] 公式工具模块加载完成");
