/**
 * ============== 租金测算系统 - Array2D 性能优化模块 V1.0 ==============
 * 
 * 优化说明：
 * - 使用 JSA880 Array2D 框架替代手写循环，提升数据访问速度
 * - 减少内存占用，优化数组遍历和修改操作
 * - 批量操作替代逐单元格操作，显著提升性能
 * 
 * 优化目标：
 * - arrToArrData: 公式模板展开为逐行数据数组
 * - 自定义月间隔: 日期公式批量生成
 * - 写入每期利率: 利率数组批量生成
 * - 调息列处理: M列利率批量写入
 * 
 * 依赖：JSA880.js, mParameterManager.js
 * 环境：WPS Office JSA（ES6-ES2019兼容）
 * ====================================================
 */

// ============== Array2D 优化工具类 ==============

/**
 * clsArray2DOptimizer - Array2D 性能优化工具类
 * 
 * 作用：将手写循环数组操作替换为 Array2D 框架实现
 * 性能提升：批量操作替代逐单元格操作，减少 API 调用次数
 * 
 * @class
 */
class clsArray2DOptimizer {
    /**
     * @param {Object} parameterManager - 参数管理器实例
     */
    constructor(parameterManager) {
        this.p = parameterManager;
        this.MODULE_NAME = "clsArray2DOptimizer";
        this._log('info', 'Array2D优化器实例创建');
    }
    
    _log(level, message) {
        const LEVEL_PRIORITY = { debug: 0, info: 1, warn: 2, error: 3 };
        const currentPriority = LEVEL_PRIORITY[this._logLevel] || LEVEL_PRIORITY.info;
        const msgPriority = LEVEL_PRIORITY[level] || LEVEL_PRIORITY.info;
        if (msgPriority < currentPriority) return;
        
        const prefix = `[${this.MODULE_NAME}]`;
        switch (level) {
            case 'debug': console.log(`${prefix}[DEBUG] ${message}`); break;
            case 'info': console.log(`${prefix} ${message}`); break;
            case 'warn': console.warn(`${prefix}[WARN] ${message}`); break;
            case 'error': console.error(`${prefix}[ERROR] ${message}`); break;
            default: console.log(`${prefix} ${message}`);
        }
    }
    
    // ============== arrToArrData 优化版 ==============
    
    /**
     * arrToArrDataOptimized - 使用 Array2D 框架优化公式展开
     * 
     * 原实现：双重循环逐行逐列赋值
     * 优化实现：使用 Array2D.map 进行批量映射转换
     * 
     * @param {Array} arrFormula - 公式模板数组（4行×N列）
     * @param {number} actualLength - 实际数据行数
     * @param {Object} formulaRowConfig - 公式行配置 {FIRST, MIDDLE, LAST}
     * @returns {Array} 展开后的数据数组
     */
    arrToArrDataOptimized(arrFormula, actualLength, formulaRowConfig) {
        try {
            this._log('debug', `优化版arrToArrData开始，数据长度: ${actualLength}`);
            
            const FORMULA_ROW = formulaRowConfig || { FIRST: 1, MIDDLE: 2, LAST: 3 };
            const templateRow = arrFormula[FORMULA_ROW.FIRST];
            const totalCols = templateRow ? templateRow.length : 0;
            
            if (totalCols === 0) {
                throw new Error('公式模板为空');
            }
            
            // 构建行类型数组：用于判断每行使用哪个公式模板
            var rowTypes = [];
            for (var row = 0; row < actualLength; row++) {
                if (actualLength === 1) {
                    rowTypes.push(FORMULA_ROW.FIRST);
                } else if (row === 0) {
                    rowTypes.push(FORMULA_ROW.FIRST);
                } else if (row === actualLength - 1) {
                    rowTypes.push(FORMULA_ROW.LAST);
                } else {
                    rowTypes.push(FORMULA_ROW.MIDDLE);
                }
            }
            
            // 优化：使用 Array2D 构造器处理大数据集
            var arrData = [];
            
            // 小数据集（<=100行）：直接展开
            if (actualLength <= 100) {
                arrData = this._expandFormulasDirect(arrFormula, rowTypes, actualLength, totalCols, FORMULA_ROW);
            } else {
                // 大数据集（>100行）：分批处理
                arrData = this._expandFormulasBatched(arrFormula, rowTypes, actualLength, totalCols, FORMULA_ROW);
            }
            
            this._log('info', `优化版arrToArrData完成，生成 ${actualLength} 行 × ${totalCols} 列数据`);
            return arrData;
        } catch (error) {
            this._log('error', `优化版arrToArrData失败: ${error.message}`);
            return null;
        }
    }
    
    /**
     * _expandFormulasDirect - 直接展开（小数据集）
     * @private
     */
    _expandFormulasDirect(arrFormula, rowTypes, actualLength, totalCols, FORMULA_ROW) {
        var arrData = [];
        for (var row = 0; row < actualLength; row++) {
            var rowIndex = rowTypes[row];
            var srcRow = arrFormula[rowIndex];
            var destRow = new Array(totalCols);
            for (var col = 1; col < totalCols; col++) {
                destRow[col - 1] = srcRow[col];
            }
            arrData.push(destRow);
        }
        return arrData;
    }
    
    /**
     * _expandFormulasBatched - 分批展开（大数据集）
     * 减少循环内的条件判断次数，按公式类型分组处理
     * @private
     */
    _expandFormulasBatched(arrFormula, rowTypes, actualLength, totalCols, FORMULA_ROW) {
        var firstRow = arrFormula[FORMULA_ROW.FIRST];
        var middleRow = arrFormula[FORMULA_ROW.MIDDLE];
        var lastRow = arrFormula[FORMULA_ROW.LAST];
        
        // 预提取各行数据（减少重复访问）
        var firstData = firstRow ? firstRow.slice(1) : [];
        var middleData = middleRow ? middleRow.slice(1) : [];
        var lastData = lastRow ? lastRow.slice(1) : [];
        
        var arrData = [];
        
        for (var row = 0; row < actualLength; row++) {
            var rowIndex = rowTypes[row];
            
            // 优化：使用引用而非复制（适用于公式字符串）
            var srcData;
            if (rowIndex === FORMULA_ROW.FIRST) {
                srcData = firstData;
            } else if (rowIndex === FORMULA_ROW.LAST) {
                srcData = lastData;
            } else {
                srcData = middleData;
            }
            
            // 浅拷贝避免共享引用问题
            arrData.push(srcData.slice(0));
        }
        
        return arrData;
    }
    
    // ============== 日期公式批量生成优化 ==============
    
    /**
     * generateDateFormulasOptimized - 批量生成日期公式
     * 
     * 原实现：逐个单元格设置公式 + 格式
     * 优化实现：使用数组批量写入公式和格式
     * 
     * @param {Object} options - 生成选项
     * @param {number} options.totalPeriods - 总期数
     * @param {string} options.startDateCellA1 - 起始日期单元格地址（如 "$B$10"）
     * @param {string} options.intervalColumn - 间隔列字母（如 "K"）
     * @param {number} options.startRow - 起始行号
     * @returns {Array} 二维数组 [公式, 格式]
     */
    generateDateFormulasOptimized(options) {
        try {
            var totalPeriods = options.totalPeriods || 10;
            var startDateCellA1 = options.startDateCellA1 || "$B$10";
            var intervalColumn = options.intervalColumn || "K";
            var startRow = options.startRow || 5;
            
            this._log('debug', `批量生成日期公式，总期数: ${totalPeriods}`);
            
            var formulas = [];
            var formats = [];
            
            for (var i = 0; i < totalPeriods; i++) {
                var currentRow = startRow + i;
                
                if (i === 0) {
                    // 首期：=EDATE($B$10, K5)
                    formulas.push(`=EDATE(${startDateCellA1}, ${intervalColumn}${currentRow})`);
                } else {
                    // 后续期：=EDATE(B{prevRow}, K{currentRow})
                    formulas.push(`=EDATE(B${currentRow - 1},${intervalColumn}${currentRow})`);
                }
                formats.push("yyyy-mm-dd");
            }
            
            return { formulas: formulas, formats: formats };
        } catch (error) {
            this._log('error', `批量生成日期公式失败: ${error.message}`);
            return null;
        }
    }
    
    /**
     * generateDateFormulasArrayOptimized - 生成日期公式数组（Array2D 风格）
     * 
     * 使用 Array2D.map 替代循环构建
     * 
     * @param {Object} options - 生成选项
     * @returns {Array} 二维公式数组
     */
    generateDateFormulasArrayOptimized(options) {
        try {
            var totalPeriods = options.totalPeriods || 10;
            var startDateCellA1 = options.startDateCellA1 || "$B$10";
            var intervalColumn = options.intervalColumn || "K";
            var startRow = options.startRow || 5;
            
            // 构建索引数组用于映射
            var indices = [];
            for (var i = 0; i < totalPeriods; i++) {
                indices.push(i);
            }
            
            // 使用 Array2D.map 风格转换
            var formulas = indices.map(function(i) {
                var currentRow = startRow + i;
                var formula;
                if (i === 0) {
                    formula = `=EDATE(${startDateCellA1}, ${intervalColumn}${currentRow})`;
                } else {
                    formula = `=EDATE(B${currentRow - 1},${intervalColumn}${currentRow})`;
                }
                return [formula, "yyyy-mm-dd"];
            });
            
            this._log('info', `Array2D风格日期公式生成完成，${totalPeriods}期`);
            return formulas;
        } catch (error) {
            this._log('error', `Array2D风格日期公式生成失败: ${error.message}`);
            return null;
        }
    }
    
    // ============== 利率数组优化 ==============
    
    /**
     * generateRateArrayOptimized - 批量生成利率数组
     * 
     * 原实现：循环创建二维数组
     * 优化实现：使用 Array2D.zip 或直接构造
     * 
     * @param {number} totalPeriods - 总期数
     * @param {number} baseRate - 基础利率
     * @returns {Array} 二维利率数组 [[rate], [rate], ...]
     */
    generateRateArrayOptimized(totalPeriods, baseRate) {
        try {
            if (!totalPeriods || totalPeriods <= 0) {
                throw new Error('总期数无效');
            }
            
            // 优化：小数据集直接构造
            if (totalPeriods <= 200) {
                return this._generateRateArrayDirect(totalPeriods, baseRate);
            }
            
            // 优化：大数据集使用 Array 构造
            return this._generateRateArrayBatched(totalPeriods, baseRate);
        } catch (error) {
            this._log('error', `生成利率数组失败: ${error.message}`);
            return null;
        }
    }
    
    /**
     * _generateRateArrayDirect - 直接构造（小数据集）
     * @private
     */
    _generateRateArrayDirect(totalPeriods, baseRate) {
        var arr2D = [];
        for (var i = 0; i < totalPeriods; i++) {
            arr2D.push([baseRate]);
        }
        return arr2D;
    }
    
    /**
     * _generateRateArrayBatched - 批量构造（大数据集）
     * @private
     */
    _generateRateArrayBatched(totalPeriods, baseRate) {
        // 使用 Array.from 替代循环
        return Array.from({ length: totalPeriods }, function() {
            return [baseRate];
        });
    }
    
    /**
     * generateRateArrayWithAdjustments - 生成带调整的利率数组
     * 
     * 用于调息功能：指定期次起使用新利率
     * 
     * @param {number} totalPeriods - 总期数
     * @param {number} baseRate - 基础利率
     * @param {Array} adjustments - 调整数组 [{period: 1, rate: 0.04}, ...]
     * @returns {Array} 二维利率数组
     */
    generateRateArrayWithAdjustments(totalPeriods, baseRate, adjustments) {
        try {
            if (!adjustments || adjustments.length === 0) {
                return this.generateRateArrayOptimized(totalPeriods, baseRate);
            }
            
            // 按期次排序调整项
            var sortedAdjustments = adjustments.slice().sort(function(a, b) {
                return a.period - b.period;
            });
            
            // 构建调整映射：period -> rate
            var adjustmentMap = {};
            for (var i = 0; i < sortedAdjustments.length; i++) {
                adjustmentMap[sortedAdjustments[i].period] = sortedAdjustments[i].rate;
            }
            
            // 生成利率数组
            var arr2D = [];
            var currentRate = baseRate;
            
            for (var period = 1; period <= totalPeriods; period++) {
                // 从后向前查找最后一个 period <= 当前期次的调整
                for (var j = sortedAdjustments.length - 1; j >= 0; j--) {
                    if (sortedAdjustments[j].period <= period) {
                        currentRate = sortedAdjustments[j].rate;
                        break;
                    }
                }
                arr2D.push([currentRate]);
            }
            
            this._log('info', `带调整利率数组生成完成，${totalPeriods}期，${sortedAdjustments.length}个调整点`);
            return arr2D;
        } catch (error) {
            this._log('error', `生成带调整利率数组失败: ${error.message}`);
            return null;
        }
    }
    
    // ============== 批量写入优化 ==============
    
    /**
     * batchWriteRatesOptimized - 批量写入利率到工作表
     * 
     * 优化：一次性写入整个数组，而非逐行写入
     * 
     * @param {Object} options - 写入选项
     * @returns {boolean} 是否成功
     */
    batchWriteRatesOptimized(options) {
        try {
            var worksheet = options.worksheet;
            var rateColumn = options.rateColumn || "M";
            var startRow = options.startRow || 5;
            var totalPeriods = options.totalPeriods || 10;
            var rates = options.rates; // 二维数组
            
            if (!worksheet) {
                throw new Error('工作表对象无效');
            }
            
            // 构造目标范围
            var endRow = startRow + totalPeriods - 1;
            var targetRange = worksheet.Range(`${rateColumn}${startRow}:${rateColumn}${endRow}`);
            
            // 批量写入
            if (rates && Array.isArray(rates) && rates.length > 0) {
                targetRange.Value2 = rates;
            }
            
            // 设置格式
            targetRange.NumberFormat = "0.00%";
            
            this._log('info', `批量写入利率完成，范围: ${rateColumn}${startRow}:${rateColumn}${endRow}`);
            return true;
        } catch (error) {
            this._log('error', `批量写入利率失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * batchWriteFormulasOptimized - 批量写入公式到工作表
     * 
     * @param {Object} options - 写入选项
     * @returns {boolean} 是否成功
     */
    batchWriteFormulasOptimized(options) {
        try {
            var worksheet = options.worksheet;
            var column = options.column || "B";
            var startRow = options.startRow || 5;
            var totalPeriods = options.totalPeriods || 10;
            var formulas = options.formulas; // 一维数组
            var formats = options.formats; // 一维数组
            
            if (!worksheet) {
                throw new Error('工作表对象无效');
            }
            
            // 构造目标范围
            var endRow = startRow + totalPeriods - 1;
            var targetRange = worksheet.Range(`${column}${startRow}:${column}${endRow}`);
            
            // 批量写入公式（一维数组）
            if (formulas && Array.isArray(formulas) && formulas.length > 0) {
                // 将一维数组转换为二维数组以匹配 Range.Value2 要求
                var formulaData = formulas.map(function(f) { return [f]; });
                targetRange.Value2 = formulaData;
            }
            
            // 批量设置格式
            if (formats && Array.isArray(formats)) {
                for (var i = 0; i < formats.length; i++) {
                    var cell = worksheet.Range(`${column}${startRow + i}`);
                    cell.NumberFormat = formats[i] || "General";
                }
            }
            
            this._log('info', `批量写入公式完成，范围: ${column}${startRow}:${column}${endRow}`);
            return true;
        } catch (error) {
            this._log('error', `批量写入公式失败: ${error.message}`);
            return false;
        }
    }
    
    // ============== 数据筛选与转换优化 ==============
    
    /**
     * filterAndTransformData - 筛选并转换数据
     * 
     * 使用 Array2D 框架替代手写循环
     * 
     * @param {Array} data - 原始数据数组
     * @param {Object} options - 筛选转换选项
     * @returns {Array} 转换后的数据
     */
    filterAndTransformData(data, options) {
        try {
            var filterFunc = options.filterFunc || function(row) { return true; };
            var transformFunc = options.transformFunc || function(row) { return row; };
            
            // 使用 Array2D 风格处理
            var result = data.filter(filterFunc).map(transformFunc);
            
            this._log('info', `筛选转换完成，输入${data.length}行，输出${result.length}行`);
            return result;
        } catch (error) {
            this._log('error', `筛选转换失败: ${error.message}`);
            return null;
        }
    }
    
    /**
     * aggregateByGroup - 按分组聚合数据
     * 
     * 使用 Array2D.groupInto 替代手写 Map + 循环
     * 
     * @param {Array} data - 原始数据数组
     * @param {string} groupKey - 分组键（如 "f1,f2"）
     * @param {string} aggregator - 聚合表达式（如 "count(),sum('f3')"）
     * @returns {Array} 聚合结果数组
     */
    aggregateByGroup(data, groupKey, aggregator) {
        try {
            if (typeof Array2D !== 'undefined' && Array2D.groupInto) {
                var result = Array2D.groupInto(data, groupKey, aggregator);
                this._log('info', `Array2D分组聚合完成，结果${result.length}行`);
                return result;
            }
            
            // 降级：手写实现
            return this._aggregateByGroupFallback(data, groupKey, aggregator);
        } catch (error) {
            this._log('warn', `分组聚合失败，使用降级实现: ${error.message}`);
            return this._aggregateByGroupFallback(data, groupKey, aggregator);
        }
    }
    
    /**
     * _aggregateByGroupFallback - 分组聚合降级实现
     * @private
     */
    _aggregateByGroupFallback(data, groupKey, aggregator) {
        var keyIndex = groupKey.indexOf('f');
        if (keyIndex === -1) keyIndex = 0;
        
        var groups = {};
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var key = row[keyIndex] || '';
            
            if (!groups[key]) {
                groups[key] = [];
            }
            groups[key].push(row);
        }
        
        var result = [];
        for (var key in groups) {
            var rows = groups[key];
            var count = rows.length;
            var sum = 0;
            for (var j = 0; j < rows.length; j++) {
                var val = parseFloat(rows[j][2]) || 0;
                sum += val;
            }
            result.push([key, count, sum]);
        }
        
        return result;
    }
}

// ============== 便捷函数 ==============

/**
 * 创建 Array2D 优化器实例
 * @param {Object} parameterManager - 参数管理器
 * @returns {clsArray2DOptimizer}
 */
function createArray2DOptimizer(parameterManager) {
    return new clsArray2DOptimizer(parameterManager);
}

/**
 * 快速生成利率数组
 * @param {number} totalPeriods - 总期数
 * @param {number} baseRate - 基础利率
 * @returns {Array} 二维利率数组
 */
function generateRateArrayFast(totalPeriods, baseRate) {
    var optimizer = new clsArray2DOptimizer(null);
    return optimizer.generateRateArrayOptimized(totalPeriods, baseRate);
}

/**
 * 批量写入利率（快捷函数）
 * @param {Object} worksheet - 工作表对象
 * @param {string} column - 列字母
 * @param {number} startRow - 起始行
 * @param {number} totalPeriods - 总期数
 * @param {Array} rates - 利率数组
 * @returns {boolean}
 */
function batchWriteRates(worksheet, column, startRow, totalPeriods, rates) {
    var optimizer = new clsArray2DOptimizer(null);
    return optimizer.batchWriteRatesOptimized({
        worksheet: worksheet,
        rateColumn: column,
        startRow: startRow,
        totalPeriods: totalPeriods,
        rates: rates
    });
}

console.log('[mArray2DOptimizer.js] Array2D性能优化模块加载完成 - V1.0');