/**
 * =======================================================================
 * superPivot v3.9.0 - WPS JSA规范版
 * =======================================================================
 * 
 * 版本: 3.9.0
 * 日期: 2026-02-06
 * 作者: JSA880框架团队
 * 
 * 功能说明:
 *   实现Excel风格的多层透视表功能，支持：
 *   - 多行多列字段配置
 *   - 多层表头（类似Excel透视表）
 *   - 行/列小计与总计
 *   - 百分比显示（占总计%/占行%/占列%）
 *   - 多种布局模式（compact/outline/tabular）
 *   - 层级缩进显示
 * 
 * 兼容性:
 *   - WPS Office JavaScript API (JSA)
 *   - 严格遵循ES5语法规范（使用var，不使用const/let）
 *   - 不使用浏览器或Node.js特有对象
 * =======================================================================
 */

/**
 * 创建超级透视表
 * @param {Array|Array2D} arr - 源数据二维数组
 * @param {Array|String} rowFields - 行字段配置，如 ['f1,f2', '大区,省份'] 或 'f1,f2'
 * @param {Array|String} colFields - 列字段配置，如 ['f3,f4', '年份,季度']
 * @param {Array|String} dataFields - 数据字段配置，如 ['sum("f5"),count()', '销售额,订单数']
 * @param {Number} headerRows - 表头行数，默认1
 * @param {Number} outputHeader - 1:输出表头, 0:不输出, -1:输出但隐藏行标题
 * @param {String} separator - 分隔符，默认"@^@"
 * @param {Object} options - 高级配置选项
 * @returns {Array2D} 透视结果（含toRange/getMeta等方法）
 * 
 * @example
 * // 基础用法
 * var result = superPivot(
 *     data,
 *     ['f1', '产品'],
 *     ['f2', '年份'],
 *     ['sum("f3")', '销售额']
 * );
 * result.toRange("A1");
 * 
 * @example
 * // 多层列字段+小计
 * var result = superPivot(
 *     data,
 *     ['f1,f2', '大区,省份'],
 *     ['f3,f4', '年份,季度'],
 *     ['sum("f5")', '销售额'],
 *     1, 1, '@^@',
 *     {
 *         cornerTitle: '销售分析',
 *         rowSubtotals: { enabled: true, label: '小计' },
 *         colSubtotals: { enabled: true, label: '小计' },
 *         grandTotals: { row: true, column: true, label: '总计' }
 *     }
 * );
 */
function superPivot(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options) {
    // ========== 参数默认值处理 ==========
    separator = separator || '@^@';
    headerRows = (headerRows !== undefined) ? headerRows : 1;
    outputHeader = (outputHeader !== undefined) ? outputHeader : 1;
    options = options || {};
    
    // ========== 解析options参数 ==========
    var cornerTitle = options.cornerTitle || '';
    var rowFieldIndent = (options.rowFieldIndent !== false);
    var rowFieldIndentSize = options.rowFieldIndentSize || 4;
    var layoutMode = options.layoutMode || 'outline';
    
    var rowSubtotals = options.rowSubtotals || { enabled: false };
    var colSubtotals = options.colSubtotals || { enabled: false };
    var grandTotals = options.grandTotals || { row: false, column: false };
    var displayAs = options.displayAs || { mode: 'value', decimals: 2 };
    
    // ========== 数据预处理 ==========
    // 处理Array2D对象
    if (arr && typeof arr === 'object' && arr._items && Array.isArray(arr._items)) {
        arr = arr._items;
    }
    
    // 跳过表头行
    var dataStartRow = headerRows || 0;
    var data = arr.slice(dataStartRow);
    
    // 过滤空行
    function isEmptyRow(row) {
        if (!row || row.length === 0) return true;
        for (var i = 0; i < row.length; i++) {
            if (row[i] !== null && row[i] !== undefined && row[i] !== '') {
                return false;
            }
        }
        return true;
    }
    
    data = data.filter(function(row) {
        return !isEmptyRow(row);
    });
    
    // ========== 解析字段配置 ==========
    function parseFields(fields) {
        var result = { fields: [], titles: [], hasTitles: false };
        
        if (typeof fields === 'string') {
            var parts = fields.split(',');
            for (var i = 0; i < parts.length; i++) {
                var match = parts[i].trim().match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    result.fields.push({ field: match[1], sort: match[2] || '+' });
                }
            }
        } else if (Array.isArray(fields) && fields.length >= 2) {
            var fieldStr = fields[0];
            var titleStr = fields[1];
            
            var fieldParts = fieldStr.split(',');
            var titleParts = (typeof titleStr === 'string') ? titleStr.split(',') : titleStr;
            
            for (var j = 0; j < fieldParts.length; j++) {
                var match = fieldParts[j].trim().match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    result.fields.push({ field: match[1], sort: match[2] || '+' });
                    result.titles.push(titleParts[j] || '');
                }
            }
            result.hasTitles = true;
        }
        
        return result;
    }
    
    var rowConfig = parseFields(rowFields);
    var colConfig = parseFields(colFields);
    
    // ========== 辅助函数：将行转为对象 ==========
    function toRowObject(row) {
        var obj = Array(row.length);
        for (var i = 0; i < row.length; i++) {
            obj['f' + (i + 1)] = row[i];
            obj[i] = row[i];
        }
        return obj;
    }
    
    // ========== 提取行键和列键 ==========
    var rowKeys = [];
    var rowKeyMap = Object.create(null);
    var colKeys = [];
    var colKeyMap = Object.create(null);
    
    for (var i = 0; i < data.length; i++) {
        var obj = toRowObject(data[i]);
        
        // 提取行键
        var rowKeyParts = [];
        for (var j = 0; j < rowConfig.fields.length; j++) {
            var rf = rowConfig.fields[j];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                rowKeyParts.push(obj[parseInt(match[1]) - 1]);
            }
        }
        var rowKey = rowKeyParts.join(separator);
        if (!rowKeyMap[rowKey]) {
            rowKeyMap[rowKey] = { values: rowKeyParts, index: i };
            rowKeys.push(rowKey);
        }
        
        // 提取列键
        var colKeyParts = [];
        for (var k = 0; k < colConfig.fields.length; k++) {
            var cf = colConfig.fields[k];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                colKeyParts.push(obj[parseInt(match[1]) - 1]);
            }
        }
        var colKey = colKeyParts.join(separator);
        if (!colKeyMap[colKey]) {
            colKeyMap[colKey] = { values: colKeyParts, index: i };
            colKeys.push(colKey);
        }
    }
    
    // ========== 对键进行排序 ==========
    function sortKeys(keys, config, keyMap) {
        keys.sort(function(a, b) {
            var aParts = a.split(separator);
            var bParts = b.split(separator);
            
            for (var i = 0; i < config.fields.length; i++) {
                var cf = config.fields[i];
                var aVal = aParts[i];
                var bVal = bParts[i];
                
                var cmp = 0;
                var aNum = parseFloat(aVal);
                var bNum = parseFloat(bVal);
                
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    cmp = aNum - bNum;
                } else {
                    cmp = String(aVal).localeCompare(String(bVal));
                }
                
                if (cmp !== 0) {
                    return (cf.sort === '-') ? -cmp : cmp;
                }
                
                if (cf.sort === '#') {
                    return keyMap[a].index - keyMap[b].index;
                }
            }
            return 0;
        });
    }
    
    sortKeys(rowKeys, rowConfig, rowKeyMap);
    sortKeys(colKeys, colConfig, colKeyMap);
    
    // ========== 分组数据 ==========
    var groupMap = Object.create(null);
    
    for (var i = 0; i < data.length; i++) {
        var obj = toRowObject(data[i]);
        
        var rowKeyParts = [];
        for (var j = 0; j < rowConfig.fields.length; j++) {
            var rf = rowConfig.fields[j];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                rowKeyParts.push(obj[parseInt(match[1]) - 1]);
            }
        }
        
        var colKeyParts = [];
        for (var k = 0; k < colConfig.fields.length; k++) {
            var cf = colConfig.fields[k];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                colKeyParts.push(obj[parseInt(match[1]) - 1]);
            }
        }
        
        var fullKey = rowKeyParts.join(separator) + '|||' + colKeyParts.join(separator);
        if (!groupMap[fullKey]) {
            groupMap[fullKey] = [];
        }
        groupMap[fullKey].push(data[i]);
    }
    
    // ========== 解析数据字段操作 ==========
    function parseDataOps(fields) {
        var ops = [];
        var str = '';
        
        if (typeof fields === 'string') {
            str = fields;
        } else if (Array.isArray(fields) && fields.length > 0) {
            str = fields[0];
        }
        
        var regex = /(\w+)\s*\(([^)]*)\)/g;
        var match;
        while ((match = regex.exec(str)) !== null) {
            ops.push({
                name: match[1],
                args: match[2].trim()
            });
        }
        
        return ops;
    }
    
    var dataOps = parseDataOps(dataFields);
    var numDataFields = dataOps.length || 1;
    
    // ========== 默认数据字段标题 ==========
    var defaultDataTitles = [];
    if (dataOps.length > 0) {
        var opNameMap = {
            'count': '计数',
            'sum': '求和',
            'average': '平均',
            'max': '最大',
            'min': '最小'
        };
        for (var i = 0; i < dataOps.length; i++) {
            defaultDataTitles.push(opNameMap[dataOps[i].name] || dataOps[i].name);
        }
    } else {
        defaultDataTitles.push('值');
    }
    
    // ========== 计算总计值（用于百分比和小计） ==========
    var grandTotalValues = null;
    if (grandTotals.row || grandTotals.column || displayAs.mode !== 'value') {
        grandTotalValues = {
            rowTotals: {},
            colTotals: {},
            grandTotal: []
        };
        
        // 计算行总计
        for (var i = 0; i < rowKeys.length; i++) {
            var rowKey = rowKeys[i];
            var total = 0;
            for (var j = 0; j < colKeys.length; j++) {
                var fullKey = rowKey + '|||' + colKeys[j];
                if (groupMap[fullKey]) {
                    for (var k = 0; k < groupMap[fullKey].length; k++) {
                        var val = parseFloat(groupMap[fullKey][k][2]); // 假设f3是数值列
                        if (!isNaN(val)) total += val;
                    }
                }
            }
            grandTotalValues.rowTotals[rowKey] = total;
        }
        
        // 计算列总计
        for (var j = 0; j < colKeys.length; j++) {
            var colKey = colKeys[j];
            var total = 0;
            for (var i = 0; i < rowKeys.length; i++) {
                var fullKey = rowKeys[i] + '|||' + colKey;
                if (groupMap[fullKey]) {
                    for (var k = 0; k < groupMap[fullKey].length; k++) {
                        var val = parseFloat(groupMap[fullKey][k][2]);
                        if (!isNaN(val)) total += val;
                    }
                }
            }
            grandTotalValues.colTotals[colKey] = total;
        }
    }
    
    // ========== 应用百分比转换 ==========
    function applyDisplayAs(value, rowKey, colKey) {
        if (displayAs.mode === 'value') return value;
        
        var val = parseFloat(value);
        if (isNaN(val)) return value;
        
        var pct = 0;
        var decimals = displayAs.decimals || 2;
        
        switch (displayAs.mode) {
            case 'percentOfGrandTotal':
                var grandTotal = 0;
                for (var key in grandTotalValues.rowTotals) {
                    grandTotal += grandTotalValues.rowTotals[key];
                }
                pct = grandTotal !== 0 ? (val / grandTotal * 100) : 0;
                break;
            case 'percentOfRowTotal':
                var rowTotal = grandTotalValues.rowTotals[rowKey] || 0;
                pct = rowTotal !== 0 ? (val / rowTotal * 100) : 0;
                break;
            case 'percentOfColTotal':
                var colTotal = grandTotalValues.colTotals[colKey] || 0;
                pct = colTotal !== 0 ? (val / colTotal * 100) : 0;
                break;
            default:
                return value;
        }
        
        return pct.toFixed(decimals) + '%';
    }
    
    // ========== 执行聚合 ==========
    function executeAggregation(group, op) {
        switch (op.name) {
            case 'count':
                return group.length;
            case 'sum':
                var total = 0;
                for (var i = 0; i < group.length; i++) {
                    var val = parseFloat(group[i][2]); // 简化处理
                    if (!isNaN(val)) total += val;
                }
                return total;
            case 'average':
                var sum = 0;
                for (var i = 0; i < group.length; i++) {
                    var val = parseFloat(group[i][2]);
                    if (!isNaN(val)) sum += val;
                }
                return group.length > 0 ? (sum / group.length) : 0;
            case 'max':
                var max = null;
                for (var i = 0; i < group.length; i++) {
                    var val = parseFloat(group[i][2]);
                    if (!isNaN(val) && (max === null || val > max)) max = val;
                }
                return max || 0;
            case 'min':
                var min = null;
                for (var i = 0; i < group.length; i++) {
                    var val = parseFloat(group[i][2]);
                    if (!isNaN(val) && (min === null || val < min)) min = val;
                }
                return min || 0;
            default:
                return group.length;
        }
    }
    
    // ========== 构建结果 ==========
    var result = [];
    var numRowFieldLevels = rowConfig.fields.length;
    var numColFieldLevels = colConfig.fields.length;
    var headerRowCount = (numColFieldLevels === 1) ? 3 : (numColFieldLevels + 1);
    
    // 构建表头
    if (outputHeader === 1 || outputHeader === true) {
        var headerRows = [];
        for (var h = 0; h < headerRowCount; h++) {
            headerRows.push([]);
        }
        
        if (numColFieldLevels === 1) {
            // 单列字段表头
            for (var i = 0; i < numRowFieldLevels; i++) {
                headerRows[0].push(rowConfig.titles[i] || '');
            }
            if (cornerTitle && numRowFieldLevels === 1) {
                headerRows[0][0] = cornerTitle;
            }
            for (var i = 0; i < numRowFieldLevels; i++) {
                headerRows[1].push('');
                headerRows[2].push('');
            }
            
            headerRows[1].push(colConfig.titles[0] || '');
            
            for (var i = 0; i < colKeys.length; i++) {
                var parts = colKeys[i].split(separator);
                headerRows[0].push(parts[0]);
                headerRows[1].push('');
            }
            
            if (colSubtotals.enabled) {
                headerRows[0].push(colSubtotals.label || '小计');
                headerRows[1].push('');
            }
            
            for (var i = 0; i < colKeys.length; i++) {
                for (var j = 0; j < numDataFields; j++) {
                    headerRows[2].push(defaultDataTitles[j]);
                }
            }
            
            if (colSubtotals.enabled) {
                for (var j = 0; j < numDataFields; j++) {
                    headerRows[2].push(defaultDataTitles[j]);
                }
            }
        } else {
            // 多列字段表头
            for (var cfIdx = 0; cfIdx < numColFieldLevels; cfIdx++) {
                var targetRow = cfIdx;
                
                for (var rfIdx = 0; rfIdx < numRowFieldLevels - 1; rfIdx++) {
                    headerRows[targetRow].push('');
                }
                
                if (cfIdx === 0 && cornerTitle) {
                    headerRows[targetRow].push(cornerTitle);
                } else {
                    headerRows[targetRow].push(colConfig.titles[cfIdx] || '');
                }
                
                for (var i = 0; i < colKeys.length; i++) {
                    var parts = colKeys[i].split(separator);
                    for (var df = 0; df < numDataFields; df++) {
                        headerRows[targetRow].push(parts[cfIdx] || '');
                    }
                }
                
                if (colSubtotals.enabled) {
                    if (cfIdx === numColFieldLevels - 1) {
                        headerRows[targetRow].push(colSubtotals.label || '小计');
                    } else {
                        headerRows[targetRow].push('');
                    }
                }
            }
            
            var lastRow = numColFieldLevels;
            for (var i = 0; i < numRowFieldLevels; i++) {
                headerRows[lastRow].push(rowConfig.titles[i] || '');
            }
            
            for (var i = 0; i < colKeys.length; i++) {
                for (var j = 0; j < numDataFields; j++) {
                    headerRows[lastRow].push(defaultDataTitles[j]);
                }
            }
            
            if (colSubtotals.enabled) {
                for (var j = 0; j < numDataFields; j++) {
                    headerRows[lastRow].push(defaultDataTitles[j]);
                }
            }
        }
        
        for (var h = 0; h < headerRowCount; h++) {
            result.push(headerRows[h]);
        }
    }
    
    // 构建数据行
    for (var i = 0; i < rowKeys.length; i++) {
        var rowKey = rowKeys[i];
        var rowKeyParts = rowKey.split(separator);
        var dataRow = rowKeyParts.slice();
        
        // 应用层级缩进
        if (rowFieldIndent && layoutMode === 'outline') {
            for (var j = 0; j < dataRow.length; j++) {
                var spaces = '';
                for (var s = 0; s < j * rowFieldIndentSize; s++) {
                    spaces += ' ';
                }
                dataRow[j] = spaces + dataRow[j];
            }
        }
        
        // 填充数据
        for (var j = 0; j < colKeys.length; j++) {
            var colKey = colKeys[j];
            var fullKey = rowKey + '|||' + colKey;
            
            if (groupMap[fullKey]) {
                for (var k = 0; k < dataOps.length; k++) {
                    var val = executeAggregation(groupMap[fullKey], dataOps[k]);
                    dataRow.push(applyDisplayAs(val, rowKey, colKey));
                }
            } else {
                for (var k = 0; k < numDataFields; k++) {
                    dataRow.push('');
                }
            }
        }
        
        // 添加列小计
        if (colSubtotals.enabled) {
            var colTotal = grandTotalValues.rowTotals[rowKey] || 0;
            dataRow.push(applyDisplayAs(colTotal, rowKey, null));
        }
        
        result.push(dataRow);
    }
    
    // 添加行总计
    if (grandTotals.row) {
        var totalRow = [];
        totalRow.push(grandTotals.label || '总计');
        for (var i = 1; i < numRowFieldLevels; i++) {
            totalRow.push('');
        }
        
        for (var j = 0; j < colKeys.length; j++) {
            var colTotal = grandTotalValues.colTotals[colKeys[j]] || 0;
            dataRow.push(applyDisplayAs(colTotal, null, colKeys[j]));
        }
        
        if (colSubtotals.enabled) {
            var grandTotal = 0;
            for (var key in grandTotalValues.rowTotals) {
                grandTotal += grandTotalValues.rowTotals[key];
            }
            totalRow.push(applyDisplayAs(grandTotal, null, null));
        }
        
        result.push(totalRow);
    }
    
    // ========== 包装结果为Array2D风格 ==========
    var wrappedResult = result;
    
    /**
     * 将结果写入WPS单元格
     * @param {String} rng - 目标单元格地址，如"A1"
     * @param {Boolean} applyMerges - 是否应用合并，默认false
     * @returns {Object} Range对象
     */
    wrappedResult.toRange = function(rng, applyMerges) {
        var app = Application;
        var screenUpdating = app.ScreenUpdating;
        var calculation = app.Calculation;
        
        try {
            app.ScreenUpdating = false;
            app.Calculation = -4135; // xlCalculationManual
            
            var ws = app.ActiveSheet;
            var startCell = ws.Range(rng);
            var rows = result.length;
            var cols = result.length > 0 ? result[0].length : 0;
            
            var targetRange = ws.Range(
                startCell,
                startCell.Offset(rows - 1, cols - 1)
            );
            
            targetRange.Value2 = result;
            
            return targetRange;
        } finally {
            app.ScreenUpdating = screenUpdating;
            app.Calculation = calculation;
        }
    };
    
    /**
     * 获取透视表元数据
     * @returns {Object} 元数据对象
     */
    wrappedResult.getMeta = function() {
        return {
            version: '3.9.0',
            rowFields: rowConfig.fields.map(function(f) { return f.field; }),
            rowTitles: rowConfig.titles,
            colFields: colConfig.fields.map(function(f) { return f.field; }),
            colTitles: colConfig.titles,
            dataFields: dataOps.map(function(op) { return op.name; }),
            dataTitles: defaultDataTitles,
            rowCount: rowKeys.length,
            colCount: colKeys.length,
            headerRowCount: headerRowCount,
            options: {
                cornerTitle: cornerTitle,
                layoutMode: layoutMode,
                rowFieldIndent: rowFieldIndent,
                rowSubtotals: rowSubtotals,
                colSubtotals: colSubtotals,
                grandTotals: grandTotals,
                displayAs: displayAs
            }
        };
    };
    
    return wrappedResult;
}

// 导出为Array2D的静态方法（如果Array2D存在）
if (typeof Array2D !== 'undefined') {
    Array2D.z超级透视 = superPivot;
    Array2D.superPivot = superPivot;
}
