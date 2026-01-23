/**
 * JSA880-enhanced.js - 郑广学JSA880快速开发框架（智能提示版本）
 * 原作者: 郑广学 (EXCEL880)
 * 改造: Claude Code
 * 版本: 3.3.0 (2024最新版 + API文档增强 + 一行代码优化)
 * 
 * API文档: https://vbayyds.com/api/jsa880/
 * 课程咨询: 微信 EXCEL880B
 * 
 * @description 郑广学JSA880快速开发框架，一行代码走天下，用最短的代码完成最复杂的需求！
 * @description 基于ES6语法重构，支持WPS JSA V8引擎
 * @description 完整实现RngUtils、Array2D、RangeChain、As类型转换等核心模块
 * @description 200+高频办公场景函数库，大部分场景可一行代码完成
 * @example
 * // 一行代码完成复杂透视汇总（核心卖点）
 * Array2D.z超级透视(数据, ['产品+,国家-'], ['月份+'], ['count(),sum("销量"),average("金额")']);
 * 
 * // 智能提示和链式调用
 * Array2D([[1,2],[3,4]]).z求和().z转置().val();
 * RngUtils("A1").z加边框().z自动列宽();
 * $.z最大行("A:A");         // 快捷调用
 * DateUtils.dt().z加天(5).z月底().val();
 * As([[1,2],[3,4]]).toArray().z转置().z求和().val();  // As类型转换
 */

// ==================== 环境检测 ====================
const isWPS = typeof Application !== 'undefined';
const isNodeJS = typeof module !== 'undefined' && module.exports;
const isBrowser = typeof window !== 'undefined';

// ==================== Lambda表达式解析器 ====================
/**
 * Lambda表达式缓存
 * @private
 */
const _lambdaCache = Object.create(null);

/**
 * 解析Lambda表达式为可执行函数（支持ES6箭头函数）
 * @private
 * @param {string|Function} expr - Lambda表达式或函数
 * @returns {Function|null} 可执行函数
 * @example
 * parseLambda('$0*2')           // _ => _[0]*2
 * parseLambda('f1+f2')           // _ => _[0]+_[1]
 * parseLambda('row=>row.x')      // row => row.x
 * parseLambda('x=>x.age>18')     // x => x.age>18
 */
function parseLambda(expr) {
    if (typeof expr === 'function') return expr;
    if (typeof expr !== 'string') return null;

    // 缓存检查
    if (_lambdaCache[expr]) return _lambdaCache[expr];

    let fn;
    try {
        // 处理箭头函数语法 (ES6)
        if (expr.includes('=>')) {
            // 使用箭头函数语法
            fn = eval('(' + expr + ')');
        }
        // 处理 $0, $1, $2 索引语法 -> 转换为箭头函数
        else if (expr.includes('$')) {
            const indexMatch = expr.match(/\$(\d+)/g);
            if (indexMatch) {
                const indices = indexMatch.map(m => parseInt(m.substring(1)));
                const maxIndex = Math.max(...indices);
                // 转换为箭头函数
                fn = new Function('_', 'return ' + expr.replace(/\$(\d+)/g, '_[$1]'));
            }
        }
        // 处理 f1, f2, f3 列选择器语法 -> 转换为箭头函数
        else if (/^f\d+/.test(expr)) {
            // 转换 f1 -> _[0], f2 -> _[1], etc.
            fn = new Function('_', 'return ' + expr.replace(/f(\d+)/g, '_[$1-1]'));
        }
        // 其他情况当作表达式
        else {
            fn = new Function('_', 'return ' + expr);
        }
    } catch (e) {
        console.warn('Lambda解析失败:', expr, e);
        return null;
    }

    _lambdaCache[expr] = fn;
    return fn;
}

/**
 * 静态方法：解析Lambda表达式
 * @param {string|Function} expr - Lambda表达式或函数
 * @returns {Function|null} 可执行函数
 * @example
 * Array2D.parseLambda('f1+f2')
 * Array2D.z解析函数表达式('row=>row.x*2')
 */
Array2D.parseLambda = parseLambda;
Array2D.z解析函数表达式 = parseLambda;

// ==================== Array2D - 二维数组工具库 ====================

/**
 * Array2D - 二维数组处理工具（支持智能提示和链式调用）
 * @constructor
 * @class
 * @description 提供丰富的二维数组操作函数，支持中英双语API
 * @param {Array} [data] - 二维数组数据
 * @returns {Array2D} Array2D实例，支持链式调用和智能提示
 * @example
 * // 基本使用
 * Array2D([[1,2,3],[4,5,6]]).z求和()        // 21
 * // 链式调用
 * Array2D([[1,2],[3,4]]).z转置().z扁平化().val()  // [1,3,2,4]
 * // Lambda表达式
 * Array2D([[1,2],[3,4]]).z求和('f1')       // 4 (第1列求和)
 * // 写入单元格
 * Array2D([[1,2],[3,4]]).toRange("A1")     // 写入A1:B2
 */
// 使用寄生组合式继承，让 Array2D 真正继承 Array
function Array2D(data) {
    // 支持工厂模式调用
    if (!(this instanceof Array2D)) {
        return new Array2D(data);
    }

    // 调用 Array 构造函数
    var items = [];
    if (data === null || data === undefined) {
        items = [];
    } else if (Array.isArray(data)) {
        items = data;
    } else {
        items = [[data]];
    }

    // 使用 Array.apply 将数组元素复制到 this
    // 这样 this 就真正拥有数组属性
    Array.prototype.push.apply(this, items);

    // 添加自定义属性（使用 Object.defineProperty 避免被枚举）
    Object.defineProperty(this, '_original', {
        value: data,
        writable: true,
        enumerable: false,
        configurable: true
    });

    // _items 作为 getter，直接返回 this（因为 this 本身就是数组）
    // 这样当数组被修改时，_items 会自动同步
    Object.defineProperty(this, '_items', {
        get: function() {
            // 将类数组对象转为真正的数组
            var arr = [];
            for (var i = 0; i < this.length; i++) {
                arr.push(this[i]);
            }
            return arr;
        },
        set: function(value) {
            // 设置时清空并重新填充
            Array.prototype.splice.call(this, 0, this.length);
            Array.prototype.push.apply(this, value);
        },
        enumerable: false,
        configurable: true
    });
}

// 设置原型链继承
Array2D.prototype = Object.create(Array.prototype);
Array2D.prototype.constructor = Array2D;

// 添加 toJSON 方法，使 JSON.stringify 只序列化数组内容
Object.defineProperty(Array2D.prototype, 'toJSON', {
    value: function() {
        return this._items;
    },
    enumerable: false,
    configurable: true,
    writable: true
});

/**
 * 创建新实例（链式调用核心）
 * @private
 * @param {Array} data - 新数据
 * @returns {Array2D} 新实例
 */
Array2D.prototype._new = function(data) {
    // 创建空实例并设置数组属性
    var instance = [];
    Array.prototype.push.apply(instance, data);

    // 使用 Object.setPrototypeOf 设置原型（如果支持）
    if (Object.setPrototypeOf) {
        Object.setPrototypeOf(instance, Array2D.prototype);
    } else {
        // 备用方案：使用 __proto__
        instance.__proto__ = Array2D.prototype;
    }

    // 添加自定义属性
    Object.defineProperty(instance, '_original', {
        value: data,
        writable: true,
        enumerable: false,
        configurable: true
    });

    // _items 作为 getter/setter，与构造函数保持一致
    Object.defineProperty(instance, '_items', {
        get: function() {
            var arr = [];
            for (var i = 0; i < this.length; i++) {
                arr.push(this[i]);
            }
            return arr;
        },
        set: function(value) {
            Array.prototype.splice.call(this, 0, this.length);
            Array.prototype.push.apply(this, value);
        },
        enumerable: false,
        configurable: true
    });

    return instance;
};

// ==================== 基础操作 ====================

/**
 * 获取/设置数组值
 * @param {Array} [newData] - 新数据（可选）
 * @returns {Array2D|Array} 设置时返回this，否则返回当前数组
 * @example
 * Array2D([[1,2]]).val()           // [[1,2]]
 * Array2D([[1,2]]).val([[3,4]])     // 返回链式对象
 */
Array2D.prototype.val = function(newData) {
    if (newData !== undefined) {
        // setter 会自动同步数组属性
        this._items = newData;
        return this;
    }
    return this._items;
};

/**
 * 检查数组是否为空
 * @returns {Boolean} 是否为空
 * @example
 * Array2D([[1]]).z是否为空()    // false
 * Array2D([]).z是否为空()       // true
 */
Array2D.prototype.z是否为空 = function() {
    return !this._items || this._items.length === 0;
};
Array2D.prototype.isEmpty = Array2D.prototype.z是否为空;

/**
 * 获取元素数量
 * @returns {Number} 元素数量
 * @example
 * Array2D([[1,2],[3,4]]).z数量()  // 4
 */
Array2D.prototype.z数量 = function() {
    return this.z扁平化().length;
};
Array2D.prototype.count = Array2D.prototype.z数量;

/**
 * 克隆数组（深拷贝）
 * @returns {Array2D} 新实例
 * @example
 * const arr = Array2D([[1,2]]);
 * const cloned = arr.z克隆();
 */
Array2D.prototype.z克隆 = function() {
    return this._new(JSON.parse(JSON.stringify(this._items)));
};
Array2D.prototype.copy = Array2D.prototype.z克隆;

// ==================== 使用辅助函数创建 Array2D 方法别名 ====================
createBilingualAliases(Array2D.prototype, [
    ['z填充', 'fill'],
    ['z补齐空位', 'fillBlank'],
    ['z扁平化', 'flat'],
    ['z反转', 'reverse'],
    ['z求和', 'sum'],
    ['z平均值', 'average'],
    ['z中位数', 'median'],
    ['z最大值', 'max'],
    ['z最小值', 'min'],
    ['z第一个', 'first'],
    ['z最后一个', 'last'],
    ['z转置', 'transpose'],
    ['z矩阵信息', 'matrixInfo'],
    ['z单元格', 'cell'],
    ['z设置单元格', 'setCell'],
    ['z写入单元格', 'toRange'],
    ['z连接', 'join'],
    ['z转JSON', 'toJson'],
    ['z分块', 'chunk'],
    ['z挑选', 'pick'],
    ['z跳过', 'skip'],
    ['z取前N个', 'take'],
    ['z查找索引', 'findIndex'],
    ['z包含', 'includes'],
    ['z筛选', 'filter'],
    ['z映射', 'map'],
    ['z归约', 'reduce'],
    ['z倒序归约', 'reduceRight'],
    ['z全部满足', 'every'],
    ['z有满足', 'some'],
    ['z行数', 'rowCount'],
    ['z列数', 'colCount'],
    ['z获取行', 'getRow'],
    ['z获取列', 'getCol'],
    ['z首行', 'firstRow'],
    ['z末行', 'lastRow'],
    ['z首列', 'firstCol'],
    ['z末列', 'lastCol'],
    ['z添加行', 'addRow'],
    ['z提取列', 'pluck'],
    ['z添加列', 'addCol'],
    ['z删除行', 'deleteRow'],
    ['z删除列', 'deleteCol'],
    ['z升序排序', 'sortAsc'],
    ['z按规则升序', 'sortBy'],
    ['z按规则降序', 'sortByDesc'],
    ['z降序排序', 'sortDesc'],
    ['z行排序', 'sortRow'],
    ['z列排序', 'sortCol'],
    ['z多列排序', 'sortByCols'],
    ['z自定义排序', 'sortByList'],
    ['z去重', 'distinct'],
    ['z转矩阵', 'toMatrix'],
    ['z分组', 'groupBy'],
    ['z透视', 'pivotBy'],
    ['z上下连接', 'concat'],
    ['z左连接', 'leftjoin'],
    ['z一对多连接', 'leftFulljoin'],
    ['z左右全连接', 'fulljoin'],
    ['z左右连接', 'zip'],
    ['z排除', 'except'],
    ['z取交集', 'intersect'],
    ['z去重并集', 'union'],
    ['z超级查找', 'superLookup'],
    ['z查找单个', 'find'],
    ['z查找所有下标', 'findAllIndex'],
    ['z查找所有行下标', 'findRowsIndex'],
    ['z查找所有列下标', 'findColsIndex'],
    ['z查找元素下标', 'findIndexByPredicate'],
    ['z值位置', 'indexOf'],
    ['z从后往前值位置', 'lastIndexOf'],
    ['z批量删除列', 'deleteCols'],
    ['z批量删除行', 'deleteRows'],
    ['z批量插入列', 'insertCols'],
    ['z批量插入行', 'insertRows'],
    ['z插入行号', 'insertRowNum'],
    ['z按页数分页', 'pageByCount'],
    ['z按行数分页', 'pageByRows'],
    ['z按下标分页', 'pageByIndexs'],
    ['z间隔取数', 'nth'],
    ['z补齐数组', 'pad'],
    ['z重设大小', 'resize'],
    ['z处理空值', 'noNull'],
    ['z选择列', 'selectCols'],
    ['z选择行', 'selectRows'],
    ['z结果', 'res'],
    ['z行切片', 'slice'],
    ['z行切片删除行', 'splice'],
    ['z转字符串', 'toString']
]);

// ==================== 填充操作 ====================

/**
 * 批量填充数组
 * @param {string|number|boolean|null|undefined} value - 填充值
 * @param {Number} [rows] - 行数（可选，默认当前行数或1）
 * @param {Number} [cols] - 列数（可选，默认当前列数或1）
 * @returns {Array2D} 新实例
 * @example
 * Array2D().z填充(0, 2, 3)  // [[0,0,0],[0,0,0]]
 */
Array2D.prototype.z填充 = function(value, rows, cols) {
    rows = rows || this._items.length || 1;
    cols = cols || (this._items[0] ? this._items[0].length : 1);
    const result = [];
    for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
            row.push(value);
        }
        result.push(row);
    }
    return this._new(result);
};
Array2D.prototype.fill = Array2D.prototype.z填充;

/**
 * 补齐空位（fillBlank）- 支持方向填充的增强版，可处理合并单元格
 * @param {string} [direction='right'] - 填充方向：left/right/up/down
 * @param {string} [rangeAddress] - 参照单元格地址（如"A2:D2"），用于确定填充区域
 * @param {any} [fillValue=''] - 填充值
 * @returns {Array2D} 新实例
 * @example
 * // 基础用法：填充null/undefined
 * Array2D([[1,null],[2,undefined]]).z补齐空位()  // [[1,''],[2,'']]
 * 
 * // 高级用法：按方向填充（用于合并单元格处理）
 * Array2D([[1,2],[3,4]]).z补齐空位('right', 'A2:D2')  // 向右填充到D2区域
 * Array2D([[1,2],[3,4]]).z补齐空位('down', 'A2:A10')  // 向下填充到A10区域
 * Array2D([[1,2],[3,4]]).z补齐空位('left', 'A2:C2')   // 向左填充到A2区域
 * Array2D([[1,2],[3,4]]).z补齐空位('up', 'A5:C10')    // 向上填充到A5区域
 * 
 * // 混合参数：先按方向填充再补全
 * Array2D([[1,null],[2]]).z补齐空位('right', 'A2:D2', 0)  // [[1,0,0,0],[2,0,0,0]]
 */
Array2D.prototype.z补齐空位 = function(direction, rangeAddress, fillValue) {
    // 参数重载处理
    if (typeof direction !== 'string') {
        // 旧版调用：仅传fillValue
        fillValue = direction;
        direction = 'right';
        rangeAddress = null;
    }
    
    fillValue = fillValue !== undefined ? fillValue : '';
    direction = direction || 'right';
    
    var result = [];
    
    // 如果提供了区域地址，解析出行列范围
    var targetRows = this._items.length;
    var targetCols = 0;
    var startRow = 0, startCol = 0;
    
    if (rangeAddress && typeof rangeAddress === 'string') {
        // 解析类似 "A2:D10" 的地址
        var match = rangeAddress.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
        if (match) {
            // 转换为0-based索引
            startCol = this._colToIndex(match[1]);  // 起始列
            startRow = parseInt(match[2]) - 1;      // 起始行
            var endCol = this._colToIndex(match[3]);   // 结束列
            var endRow = parseInt(match[4]) - 1;       // 结束行
            
            targetRows = endRow - startRow + 1;
            targetCols = endCol - startCol + 1;
        }
    }
    
    // 找出最大列数
    var maxLen = 0;
    for (var r = 0; r < this._items.length; r++) {
        if (this._items[r] && this._items[r].length > maxLen) {
            maxLen = this._items[r].length;
        }
    }
    
    // 根据方向计算最终维度
    var finalRows = targetRows || this._items.length;
    var finalCols = targetCols || Math.max(maxLen, targetCols);
    
    // 按方向填充
    for (var i = 0; i < finalRows; i++) {
        var row = new Array(finalCols);
        
        // 初始化全为fillValue
        for (var j = 0; j < finalCols; j++) {
            row[j] = fillValue;
        }
        
        // 根据方向填充原始数据
        for (var j = 0; j < finalCols; j++) {
            var origRow = i, origCol = j;
            
            // 根据方向调整坐标（处理方向偏移）
            switch (direction) {
                case 'left':
                    // 从右向左填充
                    origCol = j + (finalCols - maxLen);
                    break;
                case 'up':
                    // 从下向上填充
                    origRow = i + (finalRows - this._items.length);
                    break;
                case 'down':
                case 'right':
                default:
                    // 默认：左上对齐，向右/向下填充
                    origRow = i;
                    origCol = j;
                    break;
            }
            
            // 检查是否在原始数组范围内
            if (origRow >= 0 && origRow < this._items.length && 
                origCol >= 0 && this._items[origRow] && origCol < this._items[origRow].length) {
                var val = this._items[origRow][origCol];
                row[j] = (val === null || val === undefined) ? fillValue : val;
            }
        }
        
        result.push(row);
    }
    
    return this._new(result);
};

// 列字母转数字索引的辅助函数
Array2D.prototype._colToIndex = function(colStr) {
    var result = 0;
    for (var i = 0; i < colStr.length; i++) {
        result = result * 26 + (colStr.charCodeAt(i) - 64);
    }
    return result - 1; // 返回0-based索引
};

Array2D.prototype.fillBlank = Array2D.prototype.z补齐空位;

/**
 * 扁平化（降维）
 * @returns {Array} 一维数组
 * @example
 * Array2D([[1,2],[3,4]]).z扁平化()  // [1,2,3,4]
 */
Array2D.prototype.z扁平化 = function() {
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        if (Array.isArray(this._items[i])) {
            for (var j = 0; j < this._items[i].length; j++) {
                result.push(this._items[i][j]);
            }
        } else {
            result.push(this._items[i]);
        }
    }
    return result;
};
Array2D.prototype.flat = Array2D.prototype.z扁平化;

/**
 * 数组反转
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z反转()  // [[3,4],[1,2]]
 */
Array2D.prototype.z反转 = function() {
    return this._new(this._items.slice().reverse());
};
Array2D.prototype.reverse = Array2D.prototype.z反转;

// ==================== 统计计算 ====================

/**
 * 求和
 * @param {string|Function} [colSelector] - 列选择器 'f1'=第1列, 或回调函数
 * @returns {Number} 和
 * @example
 * Array2D([[1,2],[3,4]]).z求和()        // 10
 * Array2D([[1,2],[3,4]]).z求和('f1')     // 4 (第1列)
 */
Array2D.prototype.z求和 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    return flat.reduce((acc, val) => {
        const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        return acc + (isNaN(num) ? 0 : num);
    }, 0);
};
Array2D.prototype.sum = Array2D.prototype.z求和;

/**
 * 求平均值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 平均值
 */
Array2D.prototype.z平均值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const sum = flat.reduce((acc, val) => {
        const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        return acc + (isNaN(num) ? 0 : num);
    }, 0);
    return flat.length > 0 ? sum / flat.length : 0;
};
Array2D.prototype.average = Array2D.prototype.z平均值;

/**
 * 求最大值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 最大值，空数组返回 undefined
 */
Array2D.prototype.z最大值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)));
    if (numbers.length === 0) return undefined;  // 空数组返回 undefined 而非 -Infinity
    return Math.max(...numbers);
};
Array2D.prototype.max = Array2D.prototype.z最大值;

/**
 * 求最小值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 最小值，空数组返回 undefined
 */
Array2D.prototype.z最小值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)));
    if (numbers.length === 0) return undefined;  // 空数组返回 undefined 而非 Infinity
    return Math.min(...numbers);
};
Array2D.prototype.min = Array2D.prototype.z最小值;

/**
 * 求中位数（median）
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 中位数
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z中位数()  // 3.5
 */
Array2D.prototype.z中位数 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)))
        .map(v => typeof v === 'number' ? v : parseFloat(v));
    if (numbers.length === 0) return undefined;
    numbers.sort(function(a, b) { return a - b; });
    const mid = Math.floor(numbers.length / 2);
    return numbers.length % 2 !== 0 ? numbers[mid] : (numbers[mid - 1] + numbers[mid]) / 2;
};
Array2D.prototype.median = Array2D.prototype.z中位数;

/**
 * 获取第一个元素
 * @returns {any} 第一个元素
 */
Array2D.prototype.z第一个 = function() {
    const flat = this.z扁平化();
    return flat.length > 0 ? flat[0] : undefined;
};
Array2D.prototype.first = Array2D.prototype.z第一个;

/**
 * 获取最后一个元素
 * @returns {any} 最后一个元素
 */
Array2D.prototype.z最后一个 = function() {
    const flat = this.z扁平化();
    return flat.length > 0 ? flat[flat.length - 1] : undefined;
};
Array2D.prototype.last = Array2D.prototype.z最后一个;

// ==================== 矩阵操作 ====================

/**
 * 转置矩阵
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z转置()  // [[1,4],[2,5],[3,6]]
 */
Array2D.prototype.z转置 = function() {
    if (!this._items || this._items.length === 0) return this._new([]);
    const rows = this._items.length;
    const cols = this._items[0].length;
    const result = [];
    for (let j = 0; j < cols; j++) {
        result[j] = [];
        for (let i = 0; i < rows; i++) {
            result[j][i] = this._items[i][j];
        }
    }
    return this._new(result);
};
Array2D.prototype.transpose = Array2D.prototype.z转置;

/**
 * 获取行列数
 * @returns {String} "行数x列数"
 */
Array2D.prototype.z矩阵信息 = function() {
    const rows = this._items.length;
    const cols = rows > 0 && this._items[0] ? this._items[0].length : 0;
    return `${rows}x${cols}`;
};
Array2D.prototype.matrixInfo = Array2D.prototype.z矩阵信息;

/**
 * 获取单元格值
 * @param {Number} row - 行号（从0开始）
 * @param {Number} col - 列号（从0开始）
 * @returns {any} 单元格值
 */
Array2D.prototype.z单元格 = function(row, col) {
    if (this._items[row] && this._items[row][col] !== undefined) {
        return this._items[row][col];
    }
    return undefined;
};
Array2D.prototype.cell = Array2D.prototype.z单元格;

/**
 * 设置单元格值
 * @param {Number} row - 行号
 * @param {Number} col - 列号
 * @param {string|number|boolean|null|Date|object} value - 新值
 * @returns {Array2D} 当前实例
 */
Array2D.prototype.z设置单元格 = function(row, col, value) {
    if (!this._items[row]) this._items[row] = [];
    this._items[row][col] = value;
    return this;
};
Array2D.prototype.setCell = Array2D.prototype.z设置单元格;

/**
 * 写入单元格（实例方法，根据数组大小自动扩展区域）
 * @param {Range|string} rng - 目标单元格区域（左上角单元格）
 * @returns {Array2D} 当前实例（支持链式调用）
 * @example
 * Array2D([[1,2],[3,4]]).toRange("A1")     // 写入A1:B2
 * Array2D([[1,2],[3,4]]).z写入单元格("K1")  // 写入K1:L2
 */
Array2D.prototype.toRange = function(rng) {
    if (!isWPS) return this;
    // 空数组检查，防止 Item(0,0) 报错
    if (!this._items || this._items.length === 0) {
        return this;
    }
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = this._items.length;
    var cols = rows > 0 ? (Array.isArray(this._items[0]) ? this._items[0].length : 1) : 0;
    // 列数边界检查
    if (cols === 0) return this;
    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = this._items;
    return this;
};
Array2D.prototype.z写入单元格 = Array2D.prototype.toRange;

/**
 * 连接成字符串
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 */
Array2D.prototype.z连接 = function(separator = ',') {
    return this._items.map(row => Array.isArray(row) ? row.join(separator) : String(row)).join(separator);
};
Array2D.prototype.join = Array2D.prototype.z连接;

/**
 * 文本连接（textjoin）- 选择指定列的值，用分隔符连接
 * @param {String|Number|Function} selector - 列选择器，如 'f1' 或 0 或 row=>row.col
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D([['a','b'],['c','d']]).z文本连接(1, '+')  // "b+d"
 * Array2D([['a','b'],['c','d']]).textjoin('f2', '+')  // "b+d"
 */
Array2D.prototype.z文本连接 = function(selector, separator = ',') {
    var fn = typeof selector === 'function' ? selector : parseLambda(selector);
    var values = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (fn) {
            values.push(fn(row, i));
        } else {
            // 默认取第一列
            values.push(Array.isArray(row) ? row[0] : row);
        }
    }
    return values.join(separator);
};
Array2D.prototype.textjoin = Array2D.prototype.z文本连接;

/**
 * 转JSON（转JSON字符串，二维数组内部数组横着对齐显示）
 * @param {Boolean} [pretty=true] - 是否格式化输出（对齐显示）
 * @returns {String} JSON字符串
 * @example
 * Array2D([[1,2],[3,4]]).z转JSON()
 * // 输出:
 * // [
 * //  [1, 2],
 * //  [3, 4]
 * // ]
 * Array2D([[1,2],[3,4]]).toJson(false)    // "[[1,2],[3,4]]" 紧凑格式
 */
Array2D.prototype.z转JSON = function(pretty) {
    // 紧凑格式
    if (pretty === false) {
        return JSON.stringify(this._items);
    }
    // 格式化输出（对齐显示）
    if (Array.isArray(this._items) && this._items.length > 0 && Array.isArray(this._items[0])) {
        var lines = formatArray2DAsJSON(this._items);
        return lines.join('\n');
    }
    // 其他情况使用标准JSON格式
    return JSON.stringify(this._items, null, 2);
};
Array2D.prototype.toJson = Array2D.prototype.z转JSON;

// ==================== 分块挑选 ====================

/**
 * 分块
 * @param {Number} size - 块大小
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z分块(2)  // [[[1],[2]],[[3],[4]],[[5]]]
 */
Array2D.prototype.z分块 = function(size) {
    const result = [];
    for (let i = 0; i < this._items.length; i += size) {
        result.push(this._items.slice(i, i + size));
    }
    return this._new(result);
};
Array2D.prototype.chunk = Array2D.prototype.z分块;

/**
 * 挑选元素
 * @param {Number} count - 数量
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z挑选(3)  // [[1],[2],[3]]
 */
Array2D.prototype.z挑选 = function(count) {
    return this._new(this._items.slice(0, count));
};
Array2D.prototype.pick = Array2D.prototype.z挑选;

/**
 * 跳过元素
 * @param {Number} count - 跳过数量
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z跳过(2)  // [[3],[4],[5]]
 */
Array2D.prototype.z跳过 = function(count) {
    return this._new(this._items.slice(count));
};
Array2D.prototype.skip = Array2D.prototype.z跳过;

/**
 * 取前N个
 * @param {Number} count - 数量
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z取前N个 = function(count) {
    return this._new(this._items.slice(0, count));
};
Array2D.prototype.take = Array2D.prototype.z取前N个;

/**
 * 跳过前面连续满足（skipWhile）- 跳过前面连续满足条件的元素
 * @param {string|Function} predicate - 条件函数
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z跳过前面连续满足('x=>x[0]<4')  // [[5,6]]
 */
Array2D.prototype.z跳过前面连续满足 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return this._new(this._items.slice());
    var startIndex = 0;
    for (var i = 0; i < this._items.length; i++) {
        if (!fn(this._items[i], i)) {
            startIndex = i;
            break;
        }
        startIndex = i + 1;
    }
    return this._new(this._items.slice(startIndex));
};
Array2D.prototype.skipWhile = Array2D.prototype.z跳过前面连续满足;

/**
 * 取前面连续满足（takeWhile）- 取前面连续满足条件的元素
 * @param {string|Function} predicate - 条件函数
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z取前面连续满足('x=>x[0]<4')  // [[1,2],[3,4]]
 */
Array2D.prototype.z取前面连续满足 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return this._new([]);
    var endIndex = 0;
    for (var i = 0; i < this._items.length; i++) {
        if (!fn(this._items[i], i)) {
            endIndex = i;
            break;
        }
        endIndex = i + 1;
    }
    return this._new(this._items.slice(0, endIndex));
};
Array2D.prototype.takeWhile = Array2D.prototype.z取前面连续满足;

/**
 * 行切片（slice）- 提取指定范围的行
 * @param {Number} [start=0] - 起始索引
 * @param {Number} [end] - 结束索引（不包含）
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z行切片(1, 2)  // [[3,4]]
 */
Array2D.prototype.z行切片 = function(start, end) {
    start = start || 0;
    if (end === undefined) end = this._items.length;
    return this._new(this._items.slice(start, end));
};
Array2D.prototype.slice = Array2D.prototype.z行切片;

/**
 * 行切片删除行（splice）- 删除/插入行
 * @param {Number} start - 起始位置
 * @param {Number} [deleteCount=1] - 删除数量
 * @param {...Array} items - 要插入的行
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z行切片删除行(1, 1)  // [[1,2],[5,6]]
 */
Array2D.prototype.z行切片删除行 = function(start, deleteCount, items) {
    deleteCount = deleteCount !== undefined ? deleteCount : 1;
    var result = this._items.slice();
    var removed = result.splice.apply(result, [start, deleteCount].concat(Array.prototype.slice.call(arguments, 2)));
    return this._new(result);
};
Array2D.prototype.splice = Array2D.prototype.z行切片删除行;

/**
 * 转字符串（toString）- 将数组转换为字符串
 * @param {string} [rowSeparator='\n'] - 行分隔符
 * @param {string} [colSeparator=','] - 列分隔符
 * @returns {string} 字符串
 * @example
 * Array2D([[1,2],[3,4]]).z转字符串()  // "1,2\n3,4"
 */
Array2D.prototype.z转字符串 = function(rowSeparator, colSeparator) {
    rowSeparator = rowSeparator !== undefined ? rowSeparator : '\n';
    colSeparator = colSeparator !== undefined ? colSeparator : ',';
    return this._items.map(function(row) {
        return row.join(colSeparator);
    }).join(rowSeparator);
};
Array2D.prototype.toString = Array2D.prototype.z转字符串;

// ==================== 查找筛选 ====================

/**
 * 查找元素下标
 * @param {any} value - 要查找的值
 * @returns {Number} 下标，未找到返回-1
 */
Array2D.prototype.z查找索引 = function(value) {
    const flat = this.z扁平化();
    for (let i = 0; i < flat.length; i++) {
        if (flat[i] == value) return i;
    }
    return -1;
};
Array2D.prototype.findIndex = Array2D.prototype.z查找索引;

/**
 * 检查是否包含元素
 * @param {any} value - 要检查的值
 * @returns {Boolean} 是否包含
 */
Array2D.prototype.z包含 = function(value) {
    return this.z查找索引(value) !== -1;
};
Array2D.prototype.includes = Array2D.prototype.z包含;

/**
 * 遍历每个元素
 * @param {Function} callback - 回调函数 (item, index)
 * @returns {Array2D} this 支持链式调用
 * @example
 * Array2D([[1,2],[3,4]]).forEach((row, i) => Console.log(i, row))
 */
Array2D.prototype.forEach = function(callback) {
    this._items.forEach(callback);
    return this;
};

/**
 * 倒序遍历执行（forEachRev）- 从后向前遍历每个元素
 * @param {Function} callback - 回调函数 (item, index)，返回false可中断
 * @returns {Array2D} this 支持链式调用
 * @example
 * Array2D([[1,2],[3,4]]).z倒序遍历执行((row, i) => Console.log(i, row))
 */
Array2D.prototype.z倒序遍历执行 = function(callback) {
    for (var i = this._items.length - 1; i >= 0; i--) {
        var result = callback(this._items[i], i);
        if (result === false) break; // 支持提前退出
    }
    return this;
};
Array2D.prototype.forEachRev = Array2D.prototype.z倒序遍历执行;

/**
 * 筛选元素
 * @param {string|Function} predicate - 筛选条件
 * @returns {Array2D} 新实例
 * @example
 * Array2D([1,2,3,4]).z筛选('x=>x>2')  // [3,4]
 */
Array2D.prototype.z筛选 = function(predicate) {
    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return this._new([]);
    return this._new(this._items.filter(fn));
};
Array2D.prototype.filter = Array2D.prototype.z筛选;

/**
 * 映射转换
 * @param {string|Function} mapper - 转换函数
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z映射 = function(mapper) {
    const fn = typeof mapper === 'function' ? mapper : parseLambda(mapper);
    if (!fn) return this._new([]);
    return this._new(this._items.map(fn));
};
Array2D.prototype.map = Array2D.prototype.z映射;

/**
 * 归约计算
 * @param {Function} callback - 回调函数
 * @param {any} initialValue - 初始值
 * @returns {any} 计算结果
 */
Array2D.prototype.z归约 = function(callback, initialValue) {
    return this._items.reduce(callback, initialValue);
};
Array2D.prototype.reduce = Array2D.prototype.z归约;

/**
 * 倒序归约（reduceRight）- 从右向左归约计算
 * @param {Function} callback - 回调函数
 * @param {any} initialValue - 初始值
 * @returns {any} 计算结果
 * @example
 * Array2D([[1,2],[3,4]]).z倒序归约((acc, val) => acc + val[0], 0)  // 4
 */
Array2D.prototype.z倒序归约 = function(callback, initialValue) {
    return this._items.reduceRight(callback, initialValue);
};
Array2D.prototype.reduceRight = Array2D.prototype.z倒序归约;

/**
 * 检查是否全部满足
 * @param {string|Function} predicate - 条件
 * @returns {Boolean} 是否全部满足
 */
Array2D.prototype.z全部满足 = function(predicate) {
    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return false;
    return this._items.every(fn);
};
Array2D.prototype.every = Array2D.prototype.z全部满足;

/**
 * 检查是否有满足
 * @param {string|Function} predicate - 条件
 * @returns {Boolean} 是否有满足
 */
Array2D.prototype.z有满足 = function(predicate) {
    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return false;
    return this._items.some(fn);
};
Array2D.prototype.some = Array2D.prototype.z有满足;

// ==================== 行列操作 ====================

/**
 * 获取行数
 * @returns {Number} 行数
 */
Array2D.prototype.z行数 = function() {
    return this._items.length;
};
Array2D.prototype.rowCount = Array2D.prototype.z行数;

/**
 * 获取列数
 * @returns {Number} 列数
 */
Array2D.prototype.z列数 = function() {
    return this._items.length > 0 && this._items[0] ? this._items[0].length : 0;
};
Array2D.prototype.colCount = Array2D.prototype.z列数;

/**
 * 获取指定行
 * @param {Number} index - 行号（从0开始）
 * @returns {Array} 行数据
 */
Array2D.prototype.z获取行 = function(index) {
    return this._items[index] || [];
};
Array2D.prototype.getRow = Array2D.prototype.z获取行;

/**
 * 获取指定列
 * @param {Number} index - 列号（从0开始）
 * @returns {Array} 列数据
 */
Array2D.prototype.z获取列 = function(index) {
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row) && index < row.length) {
            result.push(row[index]);
        } else {
            result.push(undefined);
        }
    }
    return result;
};
Array2D.prototype.getCol = Array2D.prototype.z获取列;

/**
 * 获取第一行
 * @returns {Array} 第一行数据
 */
Array2D.prototype.z首行 = function() {
    return this._items[0] || [];
};
Array2D.prototype.firstRow = Array2D.prototype.z首行;

/**
 * 获取最后一行
 * @returns {Array} 最后一行数据
 */
Array2D.prototype.z末行 = function() {
    return this._items[this._items.length - 1] || [];
};
Array2D.prototype.lastRow = Array2D.prototype.z末行;

/**
 * 获取第一列
 * @returns {Array} 第一列数据
 */
Array2D.prototype.z首列 = function() {
    return this.z获取列(0);
};
Array2D.prototype.firstCol = Array2D.prototype.z首列;

/**
 * 获取最后一列
 * @returns {Array} 最后一列数据
 */
Array2D.prototype.z末列 = function() {
    return this.z获取列(this.z列数() - 1);
};
Array2D.prototype.lastCol = Array2D.prototype.z末列;

// ==================== 增删行列 ====================

/**
 * 添加行
 * @param {Array} row - 行数据
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z添加行 = function(row) {
    var result = this._items.slice();
    result.push(row);
    return this._new(result);
};
Array2D.prototype.addRow = Array2D.prototype.z添加行;

/**
 * 提取列（pluck）
 * @param {Number} colIndex - 列索引
 * @returns {Array} 列数据
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z提取列(1)  // [2,5]
 */
Array2D.prototype.z提取列 = function(colIndex) {
    return this.z获取列(colIndex);
};
Array2D.prototype.pluck = Array2D.prototype.z提取列;

/**
 * 添加列
 * @param {Array} col - 列数据
 * @param {Number} index - 插入位置（可选，默认为末尾）
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z添加列([5,6])        // [[1,2,5],[3,4,6]]
 * Array2D([[1,2],[3,4]]).z添加列([5,6], 0)     // [[5,1,2],[6,3,4]]
 */
Array2D.prototype.z添加列 = function(col, index) {
    var result = [];
    var colIndex = index !== undefined ? index : this.z列数();
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        newRow.splice(colIndex, 0, col[i] !== undefined ? col[i] : null);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.addCol = Array2D.prototype.z添加列;

/**
 * 删除行
 * @param {Number} index - 行号
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z删除行(1)  // [[1,2],[5,6]]
 */
Array2D.prototype.z删除行 = function(index) {
    // 索引边界检查
    if (index < 0 || index >= this._items.length) {
        return this._new(this._items.slice());  // 索引无效，返回副本
    }
    var result = this._items.slice();
    result.splice(index, 1);
    return this._new(result);
};
Array2D.prototype.deleteRow = Array2D.prototype.z删除行;

/**
 * 删除列
 * @param {Number} index - 列号
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z删除列(1)  // [[1,3],[4,6]]
 */
Array2D.prototype.z删除列 = function(index) {
    // 索引边界检查
    if (index < 0 || index >= this.z列数()) {
        return this._new(this._items.slice());  // 索引无效，返回副本
    }
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        newRow.splice(index, 1);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.deleteCol = Array2D.prototype.z删除列;

/**
 * 尾部弹出一项（pop）- 删除并返回最后一行
 * @returns {Array} 被删除的行
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z尾部弹出一项()  // [5,6]
 */
Array2D.prototype.z尾部弹出一项 = function() {
    if (this._items.length === 0) return undefined;
    return this._items.pop();
};
Array2D.prototype.pop = Array2D.prototype.z尾部弹出一项;

/**
 * 追加一项（push）- 向数组末尾添加行
 * @param {...Array} rows - 要添加的行
 * @returns {Number} 添加后的行数
 * @example
 * Array2D([[1,2],[3,4]]).z追加一项([5,6], [7,8])  // 4
 */
Array2D.prototype.z追加一项 = function() {
    for (var i = 0; i < arguments.length; i++) {
        this._items.push(arguments[i]);
    }
    return this._items.length;
};
Array2D.prototype.push = Array2D.prototype.z追加一项;

/**
 * 删除第一个（shift）- 删除并返回第一行
 * @returns {Array} 被删除的行
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z删除第一个()  // [1,2]
 */
Array2D.prototype.z删除第一个 = function() {
    if (this._items.length === 0) return undefined;
    return this._items.shift();
};
Array2D.prototype.shift = Array2D.prototype.z删除第一个;

// ==================== 排序去重 ====================

/**
 * sort - 原生数组 sort 方法的代理（支持链式调用）
 * @param {Function} compareFn - 比较函数
 * @returns {Array2D} 返回当前实例（支持链式调用）
 * @example
 * Array2D([[3,1],[2,2],[1,3]]).sort((a,b)=>a[0]-b[0]).val()  // [[1,3],[2,2],[3,1]]
 */
Array2D.prototype.sort = function(compareFn) {
    if (compareFn) {
        this._items.sort(compareFn);
    } else {
        this._items.sort();
    }
    return this;  // 返回 this 支持链式调用
};

/**
 * 按规则升序（sortBy）- 使用Lambda表达式指定排序键进行升序排序
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z按规则升序('x=>x[0]')  // [[1,'A'],[2,'B'],[3,'C']]
 */
Array2D.prototype.z按规则升序 = function(keySelector) {
    var fn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!fn) return this._new(this._items.slice());
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = fn(a);
        var valB = fn(b);
        if (valA < valB) return -1;
        if (valA > valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortBy = Array2D.prototype.z按规则升序;

/**
 * 按规则降序（sortByDesc）- 使用Lambda表达式指定排序键进行降序排序
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z按规则降序('x=>x[0]')  // [[3,'C'],[2,'B'],[1,'A']]
 */
Array2D.prototype.z按规则降序 = function(keySelector) {
    var fn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!fn) return this._new(this._items.slice());
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = fn(a);
        var valB = fn(b);
        if (valA > valB) return -1;
        if (valA < valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortByDesc = Array2D.prototype.z按规则降序;

/**
 * 降序排序（sortDesc）- 按首列降序排序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z降序排序()  // [[3,'C'],[2,'B'],[1,'A']]
 */
Array2D.prototype.z降序排序 = function() {
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[0];
        var valB = b[0];
        if (valA > valB) return -1;
        if (valA < valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortDesc = Array2D.prototype.z降序排序;

/**
 * 升序排序 - 按首列升序排序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z升序排序()  // [[1,'A'],[2,'B'],[3,'C']]
 */
Array2D.prototype.z升序排序 = function() {
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[0];
        var valB = b[0];
        if (valA < valB) return -1;
        if (valA > valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortAsc = Array2D.prototype.z升序排序;

/**
 * 行排序
 * @param {Number} colIndex - 排序依据的列
 * @param {Boolean} ascending - 是否升序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,3],[2,2],[3,1]]).z行排序(1)       // [[3,1],[2,2],[1,3]]
 * Array2D([[1,3],[2,2],[3,1]]).z行排序(1, false)  // [[1,3],[2,2],[3,1]] 降序
 */
Array2D.prototype.z行排序 = function(colIndex, ascending) {
    ascending = ascending !== undefined ? ascending : true;
    // 列边界检查
    if (colIndex < 0 || colIndex >= this.z列数()) {
        return this._new(this._items.slice());  // 列索引无效，返回副本
    }
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[colIndex];
        var valB = b[colIndex];
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortRow = Array2D.prototype.z行排序;

/**
 * 列排序
 * @param {Number} rowIndex - 排序依据的行
 * @param {Boolean} ascending - 是否升序
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z列排序 = function(rowIndex, ascending) {
    ascending = ascending !== undefined ? ascending : true;
    if (!this._items[rowIndex]) return this._new([]);
    var colCount = this._items[rowIndex].length;
    var indices = [];
    for (var i = 0; i < colCount; i++) indices.push(i);
    indices.sort(function(a, b) {
        var valA = this._items[rowIndex][a];
        var valB = this._items[rowIndex][b];
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    }.bind(this));
    var result = [];
    for (var r = 0; r < this._items.length; r++) {
        var newRow = [];
        for (var i = 0; i < indices.length; i++) {
            newRow.push(this._items[r][indices[i]]);
        }
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.sortCol = Array2D.prototype.z列排序;

/**
 * 多列排序 - 按多列排序，支持指定每列的升降序
 * @param {string} sortParams - 排序参数 'f3+,f4-' 表示第3列升序第4列降序
 * @param {number} [headerRows=0] - 表头的行数（不参与排序）
 * @param {string} [customOrder] - 自定义序列，逗号分隔
 * @returns {Array2D} 新实例
 * @example
 * Array2D(arr).z多列排序('f3+,f4-', 1)  // 第3列升序，第4列降序，第1行为表头
 */
Array2D.prototype.z多列排序 = function(sortParams, headerRows, customOrder) {
    headerRows = headerRows || 0;

    // 解析排序参数
    var sorts = [];
    var parts = sortParams.split(',');
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        var match = part.match(/f?(\d+)([+-])/);
        if (match) {
            sorts.push({
                col: parseInt(match[1]),
                order: match[2] === '+' ? 1 : 2 // 1升序 2降序
            });
        }
    }

    if (this._items.length <= headerRows) return this._new(this._items.slice());

    // 分离表头和数据
    var header = this._items.slice(0, headerRows);
    var data = this._items.slice(headerRows);

    // 排序
    data.sort(function(a, b) {
        for (var s = 0; s < sorts.length; s++) {
            var sort = sorts[s];
            var colIdx = sort.col - 1;
            var valA = a[colIdx];
            var valB = b[colIdx];

            // 自定义序列处理
            if (customOrder) {
                var orderArr = customOrder.split(',');
                var idxA = orderArr.indexOf(String(valA));
                var idxB = orderArr.indexOf(String(valB));
                if (idxA >= 0 && idxB >= 0) {
                    valA = idxA;
                    valB = idxB;
                }
            }

            if (valA < valB) return sort.order === 1 ? -1 : 1;
            if (valA > valB) return sort.order === 1 ? 1 : -1;
        }
        return 0;
    });

    return this._new(header.concat(data));
};
Array2D.prototype.sortByCols = Array2D.prototype.z多列排序;

/**
 * 自定义排序 - 按指定列表的顺序排序
 * @param {number|string} colIndex - 列索引（支持数字0索引或 "f3" 格式1索引）
 * @param {Array|string} orderList - 排序列表（数组或逗号分隔的字符串）
 * @param {number} [headerRows=0] - 表头的行数（不参与排序）
 * @returns {Array2D} 新实例
 * @example
 * Array2D(arr).z自定义排序("f3", "中国,英国,美国,德国")
 * Array2D(arr).sortByList(2, ["中国", "英国", "美国", "德国"], 1)
 */
Array2D.prototype.z自定义排序 = function(colIndex, orderList, headerRows) {
    headerRows = headerRows || 0;

    // 处理列索引：支持 f3 格式（从1开始的列号）或数字索引
    var actualColIndex = colIndex;
    if (typeof colIndex === 'string' && colIndex.toLowerCase().startsWith('f')) {
        actualColIndex = parseInt(colIndex.substring(1)) - 1;
    }

    // 处理排序列表：支持逗号分隔的字符串或数组
    var actualOrderList = orderList;
    if (typeof orderList === 'string') {
        actualOrderList = orderList.split(',').map(function(s) { return s.trim(); });
    }

    if (this._items.length <= headerRows) return this._new(this._items.slice());

    // 分离表头和数据
    var header = this._items.slice(0, headerRows);
    var data = this._items.slice(headerRows);

    data.sort(function(a, b) {
        var valA = a[actualColIndex];
        var valB = b[actualColIndex];
        var indexA = actualOrderList.indexOf(valA);
        var indexB = actualOrderList.indexOf(valB);

        // 不在列表中的值放到最后
        var posA = indexA === -1 ? 999 : indexA;
        var posB = indexB === -1 ? 999 : indexB;

        return posA - posB;
    });

    return this._new(header.concat(data));
};
Array2D.prototype.sortByList = Array2D.prototype.z自定义排序;

/**
 * 去重
 * @param {Number} colIndex - 依据哪一列去重（可选）
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z去重 = function(colIndex) {
    var seen = Object.create(null);
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var key = colIndex !== undefined ? this._items[i][colIndex] : JSON.stringify(this._items[i]);
        if (!seen[key]) {
            seen[key] = true;
            result.push(this._items[i]);
        }
    }
    return this._new(result);
};
Array2D.prototype.distinct = Array2D.prototype.z去重;

/**
 * 转矩阵（toMatrix）- 转换为标准矩阵格式，补齐缺失列
 * @param {any} fillValue - 填充值，默认为null
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4,5],[6]]).z转矩阵()  // [[1,2,null],[3,4,5],[6,null,null]]
 */
Array2D.prototype.z转矩阵 = function(fillValue) {
    fillValue = fillValue !== undefined ? fillValue : null;
    if (this._items.length === 0) return this._new([]);

    // 找出最大列数
    var maxCols = 0;
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var rowLen = Array.isArray(row) ? row.length : 1;
        if (rowLen > maxCols) maxCols = rowLen;
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row)) {
            var newRow = row.slice();
            while (newRow.length < maxCols) {
                newRow.push(fillValue);
            }
            result.push(newRow);
        } else {
            var newRow = [row];
            while (newRow.length < maxCols) {
                newRow.push(fillValue);
            }
            result.push(newRow);
        }
    }
    return this._new(result);
};
Array2D.prototype.toMatrix = Array2D.prototype.z转矩阵;

// ==================== 分组透视 ====================

/**
 * 分组
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} valSelector - 值选择器
 * @returns {Object} 分组结果
 */
Array2D.prototype.z分组 = function(keySelector, valSelector) {
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    var valFn = valSelector ? (typeof valSelector === 'function' ? valSelector : parseLambda(valSelector)) : null;

    var groups = Object.create(null);
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var key = keyFn ? keyFn(row, i) : row[0];
        var val = valFn ? valFn(row, i) : row;
        if (!groups[key]) groups[key] = [];
        groups[key].push(val);
    }
    return groups;
};
Array2D.prototype.groupBy = Array2D.prototype.z分组;

/**
 * 数据透视（pivotBy）- 创建数据透视表
 * @param {Number|Function} rowField - 行字段索引或选择器
 * @param {Number|Function} colField - 列字段索引或选择器
 * @param {Number|Function} valueField - 值字段索引或选择器
 * @param {Function} aggregator - 聚合函数，默认为求和
 * @returns {Array2D} 新实例（透视表）
 * @example
 * // 数据: [[产品, 地区, 销量], ['A', '北京', 100], ['A', '上海', 200]]
 * Array2D(data).z透视(0, 1, 2)  // 按产品(行)、地区(列)透视销量
 */
Array2D.prototype.z透视 = function(rowField, colField, valueField, aggregator) {
    if (this._items.length === 0) return this._new([]);

    // 默认聚合函数为求和
    var agg = aggregator || function(acc, val) {
        var num1 = typeof acc === 'number' ? acc : parseFloat(String(acc).replace(/,/g, ''));
        var num2 = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        return (isNaN(num1) ? 0 : num1) + (isNaN(num2) ? 0 : num2);
    };

    var rowValues = [];
    var colValues = [];
    var pivotData = Object.create(null);

    // 辅助函数：获取字段值
    var getFieldValue = function(row, field, index) {
        if (typeof field === 'function') return field(row, index);
        if (Array.isArray(row)) return row[field];
        return row;
    };

    // 第一遍：收集所有行和列的值
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var rowKey = String(getFieldValue(row, rowField, i));
        var colKey = String(getFieldValue(row, colField, i));
        var value = getFieldValue(row, valueField, i);

        // 收集行值
        if (rowValues.indexOf(rowKey) === -1) rowValues.push(rowKey);
        // 收集列值
        if (colValues.indexOf(colKey) === -1) colValues.push(colKey);

        // 初始化数据结构
        if (!pivotData[rowKey]) pivotData[rowKey] = Object.create(null);

        // 聚合值
        if (pivotData[rowKey][colKey] === undefined) {
            pivotData[rowKey][colKey] = value;
        } else {
            pivotData[rowKey][colKey] = agg(pivotData[rowKey][colKey], value);
        }
    }

    // 排序
    rowValues.sort();
    colValues.sort();

    // 构建结果表
    var result = [];

    // 表头
    var header = ['行\\列'].concat(colValues);
    result.push(header);

    // 数据行
    for (var r = 0; r < rowValues.length; r++) {
        var rowKey = rowValues[r];
        var rowData = [rowKey];
        for (var c = 0; c < colValues.length; c++) {
            var colKey = colValues[c];
            var value = pivotData[rowKey] && pivotData[rowKey][colKey] !== undefined
                ? pivotData[rowKey][colKey]
                : 0;
            rowData.push(value);
        }
        result.push(rowData);
    }

    return this._new(result);
};
Array2D.prototype.pivotBy = Array2D.prototype.z透视;

// ==================== 连接相关方法 ====================

/**
 * 上下连接（concat）- 将两个或多个数组按行连接
 * @param {Array} brr - 第二个数组或多个数组
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z上下连接([[5,6]])  // [[1,2],[3,4],[5,6]]
 */
Array2D.prototype.z上下连接 = function() {
    var result = this._items.slice();
    for (var i = 0; i < arguments.length; i++) {
        var arr = arguments[i];
        if (Array.isArray(arr)) {
            if (arr.length > 0 && Array.isArray(arr[0])) {
                result = result.concat(arr);
            } else {
                result.push(arr);
            }
        }
    }
    return this._new(result);
};
Array2D.prototype.concat = Array2D.prototype.z上下连接;

/**
 * 左连接（leftjoin）- 以左表为准的左外连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z左连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return JSON.stringify(row); };
    var resFn = resultSelector || function(a, b) { return a.concat(b || []); };

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var leftKey = leftFn(leftRow, i);
        var matched = null;
        for (var j = 0; j < brr.length; j++) {
            if (leftKey === rightFn(brr[j], j)) {
                matched = brr[j];
                break;
            }
        }
        result.push(resFn(leftRow.slice(), matched ? matched.slice() : []));
    }
    return this._new(result);
};
Array2D.prototype.leftjoin = Array2D.prototype.z左连接;

/**
 * 左右全连接（fulljoin）- 全外连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z左右全连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return JSON.stringify(row); };
    var resFn = resultSelector || function(a, b) { return a.concat(b || []); };

    var leftKeys = [];
    var rightKeys = [];
    var rightMap = Object.create(null);

    for (var i = 0; i < this._items.length; i++) {
        var key = leftFn(this._items[i], i);
        if (leftKeys.indexOf(key) === -1) leftKeys.push(key);
    }

    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        if (rightKeys.indexOf(key) === -1) rightKeys.push(key);
        if (!rightMap[key]) rightMap[key] = [];
        rightMap[key].push(brr[j]);
    }

    var allKeys = [];
    for (var k = 0; k < leftKeys.length; k++) {
        if (allKeys.indexOf(leftKeys[k]) === -1) allKeys.push(leftKeys[k]);
    }
    for (var k = 0; k < rightKeys.length; k++) {
        if (allKeys.indexOf(rightKeys[k]) === -1) allKeys.push(rightKeys[k]);
    }

    var result = [];
    for (var i = 0; i < allKeys.length; i++) {
        var key = allKeys[i];
        var leftRows = this._items.filter(function(row, idx) { return leftFn(row, idx) === key; });
        var rightRows = rightMap[key] || [];

        if (leftRows.length > 0 && rightRows.length > 0) {
            for (var lr = 0; lr < leftRows.length; lr++) {
                for (var rr = 0; rr < rightRows.length; rr++) {
                    result.push(resFn(leftRows[lr].slice(), rightRows[rr].slice()));
                }
            }
        } else if (leftRows.length > 0) {
            for (var lr = 0; lr < leftRows.length; lr++) {
                result.push(resFn(leftRows[lr].slice(), []));
            }
        } else {
            for (var rr = 0; rr < rightRows.length; rr++) {
                result.push(resFn([], rightRows[rr].slice()));
            }
        }
    }

    return this._new(result);
};
Array2D.prototype.fulljoin = Array2D.prototype.z左右全连接;

/**
 * 一对多连接（leftFulljoin）- 左表所有行与右表匹配的所有行连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 * @example
 * arr.leftFulljoin(brr, 'f1', 'f1')
 */
Array2D.prototype.z一对多连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return row[0]; };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return row[0]; };
    var resFn = resultSelector || function(a, b) { return a.concat(b || []); };

    var rightMap = Object.create(null);
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        if (!rightMap[key]) rightMap[key] = [];
        rightMap[key].push(brr[j]);
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var key = leftFn(leftRow, i);
        var rightRows = rightMap[key] || [];
        if (rightRows.length === 0) {
            result.push(resFn(leftRow.slice(), []));
        } else {
            for (var r = 0; r < rightRows.length; r++) {
                result.push(resFn(leftRow.slice(), rightRows[r].slice()));
            }
        }
    }
    return this._new(result);
};
Array2D.prototype.leftFulljoin = Array2D.prototype.z一对多连接;

/**
 * 左右连接（zip）- 按行左右拼接
 * @param {...Array} arrays - 要拼接的数组
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z左右连接([[5],[6]])  // [[1,2,5],[3,4,6]]
 */
Array2D.prototype.z左右连接 = function() {
    var arrays = [this._items];
    for (var i = 0; i < arguments.length; i++) {
        arrays.push(arguments[i]);
    }

    var maxRows = 0;
    for (var a = 0; a < arrays.length; a++) {
        if (arrays[a].length > maxRows) maxRows = arrays[a].length;
    }

    var result = [];
    for (var r = 0; r < maxRows; r++) {
        var row = [];
        for (var a = 0; a < arrays.length; a++) {
            var arr = arrays[a];
            if (r < arr.length) {
                var rowData = arr[r];
                if (Array.isArray(rowData)) {
                    row = row.concat(rowData);
                } else {
                    row.push(rowData);
                }
            }
        }
        result.push(row);
    }

    return this._new(result);
};
Array2D.prototype.zip = Array2D.prototype.z左右连接;

/**
 * 排除（except）- 从左表排除与右表相同的元素
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftSelector - 左表选择器
 * @param {string|Function} rightSelector - 右表选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z排除 = function(brr, leftSelector, rightSelector) {
    var leftFn = leftSelector ? parseLambda(leftSelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightSelector ? parseLambda(rightSelector) : function(row) { return JSON.stringify(row); };

    var rightKeys = [];
    for (var j = 0; j < brr.length; j++) {
        rightKeys.push(rightFn(brr[j], j));
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var key = leftFn(this._items[i], i);
        if (rightKeys.indexOf(key) === -1) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.except = Array2D.prototype.z排除;

/**
 * 取交集（intersect）- 获取两个数组的交集
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftSelector - 左表选择器
 * @param {string|Function} rightSelector - 右表选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z取交集 = function(brr, leftSelector, rightSelector) {
    var leftFn = leftSelector ? parseLambda(leftSelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightSelector ? parseLambda(rightSelector) : function(row) { return JSON.stringify(row); };

    var rightKeys = [];
    var rightMap = Object.create(null);
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        rightKeys.push(key);
        if (!rightMap[key]) rightMap[key] = [];
        rightMap[key].push(brr[j]);
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var key = leftFn(this._items[i], i);
        if (rightKeys.indexOf(key) !== -1) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.intersect = Array2D.prototype.z取交集;

/**
 * 去重并集（union）- 合并两个数组并去重
 * @param {Array} brr - 右表数组
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z去重并集 = function(brr, keySelector) {
    var keyFn = keySelector ? parseLambda(keySelector) : function(row) { return JSON.stringify(row); };

    var combined = this._items.slice();
    for (var i = 0; i < brr.length; i++) {
        combined.push(brr[i]);
    }

    var seen = Object.create(null);
    var result = [];
    for (var j = 0; j < combined.length; j++) {
        var key = keyFn(combined[j], j);
        if (!seen[key]) {
            seen[key] = true;
            result.push(combined[j]);
        }
    }

    return this._new(result);
};
Array2D.prototype.union = Array2D.prototype.z去重并集;

/**
 * 超级查找（superLookup）- 类似VLOOKUP的多条件查找
 * @param {Array} brr - 查找表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z超级查找 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return row[0]; };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return row[0]; };
    var resFn = resultSelector || function(a, b) { return a.concat(b || []); };

    // 构建右表查找字典
    var rightMap = Object.create(null);
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        if (!rightMap[key]) {
            rightMap[key] = brr[j];
        }
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var key = leftFn(leftRow, i);
        var matched = rightMap[key];
        result.push(resFn(leftRow.slice(), matched ? matched.slice() : []));
    }

    return this._new(result);
};
Array2D.prototype.superLookup = Array2D.prototype.z超级查找;

// ==================== 查找相关方法 ====================

/**
 * 查找单个元素（find）
 * @param {string|Function} predicate - 查找条件
 * @returns {any} 找到的元素
 */
Array2D.prototype.z查找单个 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return undefined;

    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            return this._items[i];
        }
    }
    return undefined;
};
Array2D.prototype.find = Array2D.prototype.z查找单个;

/**
 * 查找所有下标（findAllIndex）- 查找所有满足条件的元素位置
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 位置数组 [[行,列],...]
 */
Array2D.prototype.z查找所有下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row)) {
            for (var j = 0; j < row.length; j++) {
                if (fn(row[j], i, j)) {
                    result.push([i, j]);
                }
            }
        } else {
            if (fn(row, i, 0)) {
                result.push([i, 0]);
            }
        }
    }
    return result;
};
Array2D.prototype.findAllIndex = Array2D.prototype.z查找所有下标;

/**
 * 查找所有行下标（findRowsIndex）
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 行下标数组
 */
Array2D.prototype.z查找所有行下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            result.push(i);
        }
    }
    return result;
};
Array2D.prototype.findRowsIndex = Array2D.prototype.z查找所有行下标;

/**
 * 查找所有列下标（findColsIndex）
 * @param {Number} rowIndex - 行号
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 列下标数组
 */
Array2D.prototype.z查找所有列下标 = function(rowIndex, predicate) {
    if (!this._items[rowIndex]) return [];

    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var row = this._items[rowIndex];
    var result = [];
    for (var j = 0; j < row.length; j++) {
        if (fn(row[j], rowIndex, j)) {
            result.push(j);
        }
    }
    return result;
};
Array2D.prototype.findColsIndex = Array2D.prototype.z查找所有列下标;

/**
 * 查找元素下标（findIndexByPredicate）
 * @param {string|Function} predicate - 查找条件
 * @returns {Number} 下标，未找到返回-1
 */
Array2D.prototype.z查找元素下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return -1;

    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.findIndexByPredicate = Array2D.prototype.z查找元素下标;

/**
 * 值位置（indexOf）- 查找元素首次出现的位置
 * @param {any} value - 要查找的值
 * @param {Number} [fromIndex=0] - 开始查找的位置
 * @returns {Number} 下标，未找到返回-1
 * @example
 * Array2D([[1,2],[3,4],[1,2]]).z值位置([1,2])  // 0
 */
Array2D.prototype.z值位置 = function(value, fromIndex) {
    fromIndex = fromIndex || 0;
    for (var i = fromIndex; i < this._items.length; i++) {
        if (JSON.stringify(this._items[i]) === JSON.stringify(value)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.indexOf = Array2D.prototype.z值位置;

/**
 * 从后往前值位置（lastIndexOf）- 查找元素最后出现的位置
 * @param {any} value - 要查找的值
 * @param {Number} [fromIndex] - 开始查找的位置（从后往前）
 * @returns {Number} 下标，未找到返回-1
 * @example
 * Array2D([[1,2],[3,4],[1,2]]).z从后往前值位置([1,2])  // 2
 */
Array2D.prototype.z从后往前值位置 = function(value, fromIndex) {
    fromIndex = fromIndex !== undefined ? fromIndex : this._items.length - 1;
    for (var i = fromIndex; i >= 0; i--) {
        if (JSON.stringify(this._items[i]) === JSON.stringify(value)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.lastIndexOf = Array2D.prototype.z从后往前值位置;

// ==================== 批量操作方法 ====================

/**
 * 批量删除列（deleteCols）
 * @param {Array|String} cols - 列号数组或f模式字符串
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z批量删除列 = function(cols) {
    var colIndexes = [];

    // 解析列索引
    if (typeof cols === 'string') {
        // f模式: f1,f2 或 f3
        if (cols.startsWith('f') && !cols.includes(',')) {
            var idx = parseInt(cols.substring(1)) - 1;
            colIndexes = [idx];
        } else {
            var parts = cols.split(',');
            for (var p = 0; p < parts.length; p++) {
                if (parts[p].trim().startsWith('f')) {
                    colIndexes.push(parseInt(parts[p].trim().substring(1)) - 1);
                }
            }
        }
    } else if (Array.isArray(cols)) {
        colIndexes = cols;
    }

    // 从大到小排序删除
    colIndexes.sort(function(a, b) { return b - a; });

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        for (var c = 0; c < colIndexes.length; c++) {
            if (colIndexes[c] >= 0 && colIndexes[c] < newRow.length) {
                newRow.splice(colIndexes[c], 1);
            }
        }
        result.push(newRow);
    }

    return this._new(result);
};
Array2D.prototype.deleteCols = Array2D.prototype.z批量删除列;

/**
 * 批量删除行（deleteRows）
 * @param {Array|String} rows - 行号数组或f模式
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z批量删除行 = function(rows) {
    var rowIndexes = [];

    if (typeof rows === 'string') {
        // f模式: f2-f4
        if (rows.includes('-')) {
            var match = rows.match(/f(\d+)\-f(\d+)/);
            if (match) {
                var start = parseInt(match[1]) - 1;
                var end = parseInt(match[2]) - 1;
                for (var i = start; i <= end; i++) {
                    rowIndexes.push(i);
                }
            }
        } else if (rows.startsWith('f')) {
            rowIndexes = [parseInt(rows.substring(1)) - 1];
        }
    } else if (Array.isArray(rows)) {
        rowIndexes = rows;
    }

    // 从大到小排序删除
    rowIndexes.sort(function(a, b) { return b - a; });

    var result = this._items.slice();
    for (var r = 0; r < rowIndexes.length; r++) {
        if (rowIndexes[r] >= 0 && rowIndexes[r] < result.length) {
            result.splice(rowIndexes[r], 1);
        }
    }

    return this._new(result);
};
Array2D.prototype.deleteRows = Array2D.prototype.z批量删除行;

/**
 * 批量插入列（insertCols）
 * @param {Number|Function} colSelector - 列号或条件回调
 * @param {any|Function} value - 填充值或回调
 * @param {Number} count - 插入数量
 * @returns {Array2D} 新实例
 * @example
 * // 在第2列位置插入2列
 * Array2D(arr).z批量插入列(1, "x", 2)
 * // 在包含"产品"值的列位置前插入2列（默认在最后一行查找）
 * Array2D(arr).z批量插入列(x=>x.includes("产品"), " ", 2)
 */
Array2D.prototype.z批量插入列 = function(colSelector, value, count) {
    count = count || 1;
    var fillVal = value;

    var insertIndex = 0;
    if (typeof colSelector === 'function') {
        // 从条件函数解析目标值
        var funcStr = colSelector.toString();
        var valueMatch = funcStr.match(/['"]([^'"]+)['"]/);

        if (valueMatch) {
            var targetValue = valueMatch[1];
            // 默认在最后一行查找目标值的位置
            var lastRow = this._items[this._items.length - 1];
            if (Array.isArray(lastRow)) {
                for (var j = 0; j < lastRow.length; j++) {
                    if (String(lastRow[j]) == targetValue) {
                        insertIndex = j;
                        break;
                    }
                }
            }
        } else {
            // 尝试从 x[N] 解析列索引
            var indexMatch = funcStr.match(/x\[(\d+)\]/);
            if (indexMatch) {
                insertIndex = parseInt(indexMatch[1]);
            }
        }
    } else if (typeof colSelector === 'number') {
        insertIndex = colSelector;
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (!Array.isArray(row)) row = [row];

        var newRow = row.slice();
        // 准备填充值
        var fillVals = [];
        for (var c = 0; c < count; c++) {
            if (typeof fillVal === 'function') {
                fillVals.push(fillVal(row, i, insertIndex + c));
            } else {
                fillVals.push(fillVal !== undefined ? fillVal : '');
            }
        }
        // 在指定位置插入
        newRow.splice.apply(newRow, [insertIndex, 0].concat(fillVals));
        result.push(newRow);
    }

    return this._new(result);
};
Array2D.prototype.insertCols = Array2D.prototype.z批量插入列;

/**
 * 批量插入行（insertRows）
 * @param {Array|Function} rowSelector - 行号数组或条件回调
 * @param {any} value - 填充值
 * @param {Number} count - 插入数量
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z批量插入行 = function(rowSelector, value, count) {
    count = count || 1;
    var fillVal = value !== undefined ? value : '';

    var insertIndexes = [];
    if (typeof rowSelector === 'function') {
        for (var i = 0; i < this._items.length; i++) {
            if (rowSelector(this._items[i], i)) {
                insertIndexes.push(i);
            }
        }
    } else if (typeof rowSelector === 'string' && rowSelector.startsWith('f')) {
        insertIndexes = [parseInt(rowSelector.substring(1)) - 1];
    } else if (Array.isArray(rowSelector)) {
        insertIndexes = rowSelector;
    }

    var result = this._items.slice();
    // 从后往前插入
    for (var i = insertIndexes.length - 1; i >= 0; i--) {
        var idx = insertIndexes[i];
        var newRow = [];
        var maxCols = 0;
        for (var r = 0; r < result.length; r++) {
            if (Array.isArray(result[r]) && result[r].length > maxCols) {
                maxCols = result[r].length;
            }
        }
        for (var c = 0; c < maxCols; c++) {
            newRow.push(fillVal);
        }
        for (var c = 0; c < count; c++) {
            result.splice(idx, 0, newRow.slice());
        }
    }

    return this._new(result);
};
Array2D.prototype.insertRows = Array2D.prototype.z批量插入行;

/**
 * 插入行号（insertRowNum）
 * @param {Number} startNum - 起始行号
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z插入行号 = function(startNum) {
    startNum = startNum || 0;
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = [startNum + i].concat(this._items[i]);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.insertRowNum = Array2D.prototype.z插入行号;

// ==================== 分页方法 ====================

/**
 * 按页数分页（pageByCount）
 * @param {Number} pageCount - 总页数
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按页数分页 = function(pageCount) {
    if (pageCount < 1) pageCount = 1;
    var totalRows = this._items.length;
    var rowsPerPage = Math.ceil(totalRows / pageCount);

    var result = [];
    for (var page = 0; page < pageCount; page++) {
        var start = page * rowsPerPage;
        var end = Math.min(start + rowsPerPage, totalRows);
        if (start < totalRows) {
            result.push(this._items.slice(start, end));
        }
    }

    return result;
};
Array2D.prototype.pageByCount = Array2D.prototype.z按页数分页;

/**
 * 按行数分页（pageByRows）
 * @param {Number} pageSize - 每页行数
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按行数分页 = function(pageSize) {
    if (pageSize < 1) pageSize = 1;

    var result = [];
    for (var i = 0; i < this._items.length; i += pageSize) {
        result.push(this._items.slice(i, i + pageSize));
    }

    return result;
};
Array2D.prototype.pageByRows = Array2D.prototype.z按行数分页;

/**
 * 按下标分页（pageByIndexs）
 * @param {Array|String} indexes - 下标数组或条件
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按下标分页 = function(indexes) {
    var splitIndexes = [];

    if (typeof indexes === 'string') {
        // f模式条件
        var fn = parseLambda(indexes);
        if (fn) {
            for (var i = 0; i < this._items.length; i++) {
                if (fn(this._items[i], i)) {
                    splitIndexes.push(i);
                }
            }
        }
    } else if (Array.isArray(indexes)) {
        splitIndexes = indexes;
    }

    if (splitIndexes.length === 0) return [this._items.slice()];

    var result = [];
    var start = 0;
    for (var i = 0; i < splitIndexes.length; i++) {
        var idx = splitIndexes[i];
        if (idx > start) {
            result.push(this._items.slice(start, idx));
        }
        start = idx;
    }
    result.push(this._items.slice(start));

    return result;
};
Array2D.prototype.pageByIndexs = Array2D.prototype.z按下标分页;

// ==================== 其他高级方法 ====================

/**
 * 间隔取数（nth）
 * @param {Number} interval - 间隔
 * @param {Number} offset - 偏移
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z间隔取数 = function(interval, offset) {
    interval = interval || 1;
    offset = offset || 0;

    // 保留第一行（表头）
    var result = [this._items[0].slice()];

    for (var i = 1; i < this._items.length; i++) {
        if ((i - 1 + offset) % interval === 0) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.nth = Array2D.prototype.z间隔取数;

/**
 * 补齐数组（pad）
 * @param {Number} cols - 列数
 * @param {Number} rows - 行数
 * @param {any} fillValue - 填充值
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z补齐数组 = function(cols, rows, fillValue) {
    cols = cols || (this._items[0] ? this._items[0].length : 1);
    rows = rows || this._items.length;
    fillValue = fillValue !== undefined ? fillValue : '';

    var result = [];
    for (var i = 0; i < rows; i++) {
        var row = i < this._items.length ? this._items[i].slice() : [];
        while (row.length < cols) {
            row.push(fillValue);
        }
        result.push(row);
    }

    return this._new(result);
};
Array2D.prototype.pad = Array2D.prototype.z补齐数组;

/**
 * 重设大小（resize）
 * @param {Number} rows - 行数
 * @param {Number} cols - 列数
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z重设大小 = function(rows, cols) {
    return this.z补齐数组(cols, rows);
};
Array2D.prototype.resize = Array2D.prototype.z重设大小;

/**
 * 处理空值（noNull）- 将null和undefined替换为空字符串
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z处理空值 = function() {
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = [];
        if (Array.isArray(this._items[i])) {
            for (var j = 0; j < this._items[i].length; j++) {
                var val = this._items[i][j];
                row.push((val === null || val === undefined) ? '' : val);
            }
        } else {
            var val = this._items[i];
            row.push((val === null || val === undefined) ? '' : val);
        }
        result.push(row);
    }
    return this._new(result);
};
Array2D.prototype.noNull = Array2D.prototype.z处理空值;

/**
 * 选择列（selectCols）- 选择二维数组中指定的列
 * @param {Array|String} cols - 列选择方式，支持多种格式：
 *   - 数字数组: [0, 2, 4] 选择第1、3、5列（0-based索引）
 *   - f模式字符串: "f1,f3,f5" 选择第1、3、5列（1-based索引）
 *   - f模式数组: ["f1", "f3", "f5"]
 *   - 表头名称数组: ["产品", "数量", "价格"] 按首行表头匹配
 *   - 单个表头名: "产品" 选择单列
 * @param {Array} [newHeaders] - 可选，为选择后的列指定新表头
 * @returns {Array2D} 新实例
 * @example
 * // 示例1：按列号选择
 * var arr = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
 * Array2D.selectCols(arr, [0, 2]);  // 选择第1列和第3列
 * // 结果: [[1, 3], [4, 6], [7, 9]]
 *
 * // 示例2：按f模式字符串选择（推荐）
 * Array2D.selectCols(arr, "f1,f3");  // 选择第1列和第3列
 *
 * // 示例3：按表头选择
 * var arr2 = [['a','b','c'], [1,2,3], [4,5,6]];
 * Array2D.selectCols(arr2, ['c','b','a']);  // 按首行表头错位选择
 * // 结果: [["c","b","a"], [3,2,1], [6,5,4]]
 *
 * // 示例4：指定新表头
 * Array2D.selectCols(arr2, ['a','c'], ['x','z']);
 * // 结果: [["x","z"], [1,3], [4,6]]
 */
Array2D.prototype.z选择列 = function(cols, newHeaders) {
    if (!this._items.length) return this._new([]);

    var indexes = [];
    var useHeader = false;

    // 处理字符串参数：支持 "f2,f3,f6" 或 "f2,f3-f7" 范围格式
    if (typeof cols === 'string') {
        // 检查是否是 f 模式（列号格式）
        if ((cols.includes(',') || cols.includes('-')) && (cols.toLowerCase().includes('f'))) {
            // f 模式：先按逗号分割，再处理范围
            var parts = cols.split(',');
            indexes = [];
            for (var i = 0; i < parts.length; i++) {
                var part = parts[i].trim();
                if (part.includes('-')) {
                    // 处理范围 f3-f7
                    var range = part.split('-');
                    var start = parseInt(range[0].toLowerCase().replace('f', ''));
                    var end = parseInt(range[1].toLowerCase().replace('f', ''));
                    for (var j = start; j <= end; j++) {
                        indexes.push(j - 1);
                    }
                } else if (part.toLowerCase().startsWith('f')) {
                    indexes.push(parseInt(part.substring(1)) - 1);  // f2 → 索引1
                } else {
                    indexes.push(parseInt(part) - 1);
                }
            }
            useHeader = false;
        } else {
            // 单个字符串，当作表头名称
            cols = [cols];
            useHeader = true;
        }
    } else if (cols.length > 0 && typeof cols[0] === 'string') {
        // 检查是否是 f 模式数组
        var allFMode = true;
        for (var i = 0; i < cols.length; i++) {
            if (typeof cols[i] === 'string' && !cols[i].toLowerCase().startsWith('f')) {
                allFMode = false;
                break;
            }
        }
        if (allFMode) {
            // f 模式数组：转换为列索引（支持范围）
            indexes = [];
            for (var i = 0; i < cols.length; i++) {
                var c = cols[i];
                if (c.includes('-')) {
                    // 处理范围
                    var range = c.split('-');
                    var start = parseInt(range[0].substring(1));
                    var end = parseInt(range[1].substring(1));
                    for (var j = start; j <= end; j++) {
                        indexes.push(j - 1);
                    }
                } else {
                    indexes.push(parseInt(c.substring(1)) - 1);
                }
            }
            useHeader = false;
        } else {
            useHeader = true;
        }
    }

    if (!useHeader && indexes.length > 0) {
        // 按列号选择（已解析的索引）
        var result = [];
        for (var i = 0; i < this._items.length; i++) {
            var row = [];
            for (var k = 0; k < indexes.length; k++) {
                row.push(this._items[i][indexes[k]]);
            }
            result.push(row);
        }
        return this._new(result);
    }

    if (useHeader) {
        // 按表头选择
        var headers = this._items[0];
        var headerMap = {};
        for (var i = 0; i < headers.length; i++) {
            headerMap[String(headers[i])] = i;
        }

        for (var j = 0; j < cols.length; j++) {
            var col = cols[j];
            if (headerMap.hasOwnProperty(col)) {
                indexes.push(headerMap[col]);
            }
        }

        var result = [];
        if (newHeaders && newHeaders.length > 0) {
            result.push(newHeaders);
        } else {
            var headerRow = [];
            for (var k = 0; k < cols.length; k++) {
                var idx = indexes[k];
                headerRow.push(idx !== undefined ? headers[idx] : cols[k]);
            }
            result.push(headerRow);
        }

        for (var i = 1; i < this._items.length; i++) {
            var row = this._items[i];
            var newRow = [];
            for (var k = 0; k < indexes.length; k++) {
                newRow.push(row[indexes[k]]);
            }
            result.push(newRow);
        }

        return this._new(result);
    }
};
Array2D.prototype.selectCols = Array2D.prototype.z选择列;

/**
 * 选择行（selectRows）
 * @param {Array} rowIndexes - 行号数组
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z选择行 = function(rowIndexes) {
    var result = [];
    for (var i = 0; i < rowIndexes.length; i++) {
        var idx = rowIndexes[i];
        if (idx >= 0 && idx < this._items.length) {
            result.push(this._items[idx]);
        }
    }
    return this._new(result);
};
Array2D.prototype.selectRows = Array2D.prototype.z选择行;

/**
 * 获取结果（res）- 获取当前数组的值（val的别名）
 * @returns {Array} 当前数组
 * @example
 * Array2D([[1,2],[3,4]]).z结果()  // [[1,2],[3,4]]
 */
Array2D.prototype.z结果 = function() {
    return this._items;
};
Array2D.prototype.res = Array2D.prototype.z结果;

/**
 * 矩阵分布（getMatrix）- 生成数字序列的矩阵分布
 * @param {Number} totalRows - 总行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'或'c'
 * @returns {Array} 分布后的数组
 */
Array2D.getMatrix = function(totalRows, cols, direction) {
    if (totalRows === undefined || totalRows <= 0) return [];
    if (cols === undefined || cols <= 0) return [];
    direction = direction || 'r';
    var result = [];
    var numbers = [];
    for (var i = 0; i < totalRows; i++) {
        numbers.push(i);
    }

    if (direction === 'r') {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = i * cols + j;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    } else {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = j * rows + i;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    }

    return result;
};
Array2D.z矩阵分布 = Array2D.getMatrix;

/**
 * 生成下标数组（getIndexs）
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @param {Number} step - 步长
 * @returns {Array} 序列
 */
Array2D.getIndexs = function(start, end, step) {
    step = step || 1;
    var result = [];
    for (var i = start; i <= end; i += step) {
        result.push(i);
    }
    return result;
};
Array2D.z生成下标数组 = Array2D.getIndexs;

/**
 * 下标数组（indexArray）- 根据条件获取元素的下标数组
 * @param {Array} arr - 数组
 * @param {string|Function} predicate - 筛选条件
 * @returns {Array} 下标数组
 * @example
 * Array2D.indexArray([[1,2],[3,4],[5,6]], 'x=>x[0]>1')  // [1, 2]
 */
Array2D.indexArray = function(arr, predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        if (fn(arr[i], i)) {
            result.push(i);
        }
    }
    return result;
};
Array2D.z下标数组 = Array2D.indexArray;

/**
 * 按范围遍历（rangeForEach）- 对指定索引范围的元素执行回调
 * @param {Array} arr - 数组
 * @param {Number} start - 起始索引
 * @param {Number} end - 结束索引
 * @param {Function} callback - 回调函数 (item, index)
 * @returns {void}
 * @example
 * Array2D.rangeForEach([[1,2],[3,4],[5,6]], 0, 1, (row, i) => Console.log(row))
 */
Array2D.rangeForEach = function(arr, start, end, callback) {
    if (!arr || !Array.isArray(arr)) return;
    start = start || 0;
    end = end !== undefined ? end : arr.length - 1;
    for (var i = start; i <= end && i < arr.length; i++) {
        callback(arr[i], i);
    }
};
Array2D.z按范围遍历 = Array2D.rangeForEach;

/**
 * 局部映射（rangeMap）- 对指定索引范围的元素进行映射
 * @param {Array} arr - 数组
 * @param {Number} start - 起始索引
 * @param {Number} end - 结束索引
 * @param {string|Function} mapper - 转换函数
 * @returns {Array} 映射后的数组
 * @example
 * Array2D.rangeMap([[1,2],[3,4],[5,6]], 0, 1, 'x=>x[0]*2')  // [[2], [6]]
 */
Array2D.rangeMap = function(arr, start, end, mapper) {
    if (!arr || !Array.isArray(arr)) return [];
    var fn = typeof mapper === 'function' ? mapper : parseLambda(mapper);
    if (!fn) return [];
    start = start || 0;
    end = end !== undefined ? end : arr.length - 1;
    var result = [];
    for (var i = start; i <= end && i < arr.length; i++) {
        result.push(fn(arr[i], i));
    }
    return result;
};
Array2D.z局部映射 = Array2D.rangeMap;

/**
 * 排名（rank）- 对数组进行排名
 * @param {Array} arr - 数组
 * @param {string|Function} colSelector - 列选择器，支持 f2, f2-（降序）
 * @param {string} [type='cn'] - 排名类型：'cn'中式排名（并列跳过），'usa'美式排名（并列不跳过），'+'顺序编号
 * @returns {Array} 排名结果（二维数组，每行一个排名值）
 * @example
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-')  // [[1],[3],[1]]（中式）
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-', 'usa')  // [[1],[3],[1]]（美式）
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-', '+')  // [[1],[3],[2]]（顺序）
 */
Array2D.rank = function(arr, colSelector, type) {
    if (!arr || !Array.isArray(arr)) return [];
    type = type || 'cn';
    var selectorStr = typeof colSelector === 'string' ? colSelector : '';
    var isDesc = selectorStr.endsWith('-');
    var fn = typeof colSelector === 'function' ? colSelector : parseLambda(colSelector);
    if (!fn) return [];

    var values = arr.map(function(row, i) { return {value: fn(row, i), index: i}; });
    values.sort(function(a, b) {
        var cmp = 0;
        if (typeof a.value === 'number' && typeof b.value === 'number') {
            cmp = a.value - b.value;
        } else {
            cmp = String(a.value).localeCompare(String(b.value));
        }
        return isDesc ? -cmp : cmp;
    });

    var ranks = [];
    for (var i = 0; i < values.length; i++) {
        var rank;
        if (type === '+') {
            rank = i + 1;
        } else if (type === 'usa') {
            rank = i + 1;
        } else { // cn 中式排名
            rank = i + 1;
            for (var j = i - 1; j >= 0; j--) {
                if (values[j].value === values[i].value) {
                    rank = j + 1;
                    break;
                }
            }
        }
        ranks[values[i].index] = [rank];
    }
    return ranks;
};
Array2D.z排名 = Array2D.rank;

/**
 * 分组排名（rankGroup）- 按分组进行排名
 * @param {Array} arr - 数组
 * @param {string|Function} colSelector - 列选择器，支持 f2, f2-（降序）
 * @param {string|Function} groupCol - 分组列选择器
 * @param {string} [type='cn'] - 排名类型
 * @param {Number} [skipHeader=0] - 跳过标题行数
 * @returns {Array} 排名结果（二维数组）
 * @example
 * Array2D.rankGroup([[1,'A',90],[2,'A',80],[3,'B',90]], 'f3-', 'f2')
 */
Array2D.rankGroup = function(arr, colSelector, groupCol, type, skipHeader) {
    if (!arr || !Array.isArray(arr)) return [];
    type = type || 'cn';
    skipHeader = skipHeader || 0;
    var selectorStr = typeof colSelector === 'string' ? colSelector : '';
    var isDesc = selectorStr.endsWith('-');
    var fn = typeof colSelector === 'function' ? colSelector : parseLambda(colSelector);
    var groupFn = typeof groupCol === 'function' ? groupCol : parseLambda(groupCol);
    if (!fn || !groupFn) return [];

    var data = arr.slice(skipHeader);
    var groups = Object.create(null);
    for (var i = 0; i < data.length; i++) {
        var key = JSON.stringify(groupFn(data[i], i));
        if (!groups[key]) groups[key] = [];
        groups[key].push({row: data[i], index: i + skipHeader});
    }

    var ranks = [];
    for (var h = 0; h < skipHeader; h++) {
        ranks.push(['']);
    }

    for (var key in groups) {
        var group = groups[key];
        var values = group.map(function(item) {
            return {value: fn(item.row, item.index), index: item.index};
        });
        values.sort(function(a, b) {
            var cmp = 0;
            if (typeof a.value === 'number' && typeof b.value === 'number') {
                cmp = a.value - b.value;
            } else {
                cmp = String(a.value).localeCompare(String(b.value));
            }
            return isDesc ? -cmp : cmp;
        });

        for (var j = 0; j < values.length; j++) {
            var rank;
            if (type === '+') {
                rank = j + 1;
            } else if (type === 'usa') {
                rank = j + 1;
            } else { // cn 中式排名
                rank = j + 1;
                for (var k = j - 1; k >= 0; k--) {
                    if (values[k].value === values[j].value) {
                        rank = k + 1;
                        break;
                    }
                }
            }
            ranks[values[j].index] = [rank];
        }
    }
    return ranks;
};
Array2D.z分组排名 = Array2D.rankGroup;

/**
 * 笛卡尔积（crossjoin）- 两个数组的笛卡尔积
 * @param {Array} arr - 第一个数组
 * @param {Array} brr - 第二个数组
 * @returns {Array} 笛卡尔积结果
 * @example
 * Array2D.crossjoin([[1,2],[3,4]], [[5,6],[7,8]])  // [[1,2,5,6],[1,2,7,8],[3,4,5,6],[3,4,7,8]]
 */
Array2D.crossjoin = function(arr, brr) {
    if (!arr || !brr) return [];
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var aRow = Array.isArray(arr[i]) ? arr[i] : [arr[i]];
        for (var j = 0; j < brr.length; j++) {
            var bRow = Array.isArray(brr[j]) ? brr[j] : [brr[j]];
            result.push(aRow.concat(bRow));
        }
    }
    return result;
};
Array2D.z笛卡尔积 = Array2D.crossjoin;

/**
 * 分组汇总（groupInto）- 按键分组并进行汇总计算
 * @param {Array} arr - 数组
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} valueSelector - 值聚合选择器（支持 g.sum(), g.count(), g.average() 等）
 * @param {string} [separator='@^@'] - 多列分组时的分隔符
 * @returns {Array} 分组汇总结果（二维数组）
 * @example
 * Array2D.groupInto([[1,'A',10],[2,'B',20],[3,'A',30]], 'f2', 'g=>g.sum("f3")')
 */
Array2D.groupInto = function(arr, keySelector, valueSelector, separator) {
    if (!arr || !Array.isArray(arr)) return [];
    separator = separator || '@^@';
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!keyFn) return [];

    // 解析值选择器
    var valueFn;
    if (typeof valueSelector === 'string') {
        // 检查是否是聚合函数格式
        var aggMatch = valueSelector.match(/g\.(sum|count|average|max|min)\s*\(\s*["']?f?(\d+)["']?\s*\)/i);
        if (aggMatch) {
            var aggFunc = aggMatch[1].toLowerCase();
            var colIdx = parseInt(aggMatch[2]) - 1;
            valueFn = function(group) {
                var arr2d = new Array2D(group);
                switch (aggFunc) {
                    case 'sum': return arr2d.z求和(function(r) { return r[colIdx]; });
                    case 'count': return arr2d.z数量();
                    case 'average': return arr2d.z平均值(function(r) { return r[colIdx]; });
                    case 'max': return arr2d.z最大值(function(r) { return r[colIdx]; });
                    case 'min': return arr2d.z最小值(function(r) { return r[colIdx]; });
                    default: return null;
                }
            };
        } else {
            valueFn = typeof valueSelector === 'function' ? valueSelector : parseLambda(valueSelector);
        }
    } else {
        valueFn = valueSelector;
    }

    if (!valueFn) return [];

    var groups = Object.create(null);
    for (var i = 0; i < arr.length; i++) {
        var key = keyFn(arr[i], i);
        var keyStr = Array.isArray(key) ? key.join(separator) : String(key);
        if (!groups[keyStr]) {
            groups[keyStr] = { key: key, rows: [] };
        }
        groups[keyStr].rows.push(arr[i]);
    }

    var result = [];
    for (var key in groups) {
        var group = groups[key];
        var agg = valueFn(group.rows);
        var row = Array.isArray(group.key) ? group.key.concat([agg]) : [group.key, agg];
        result.push(row);
    }
    return result;
};
Array2D.z分组汇总 = Array2D.groupInto;

/**
 * 分组汇总到字典（groupIntoMap）- 按键分组并汇总到Map对象
 * @param {Array} arr - 数组
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} [valueSelector] - 值选择器
 * @returns {Map} Map对象，键为分组键，值为 {group: 数组, agg: 聚合结果}
 * @example
 * var map = Array2D.groupIntoMap([[1,'A',10],[2,'B',20]], 'f2')
 */
Array2D.groupIntoMap = function(arr, keySelector, valueSelector) {
    if (!arr || !Array.isArray(arr)) return new Map();
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    var valueFn = valueSelector ? (typeof valueSelector === 'function' ? valueSelector : parseLambda(valueSelector)) : null;
    if (!keyFn) return new Map();

    var map = new Map();
    for (var i = 0; i < arr.length; i++) {
        var key = keyFn(arr[i], i);
        if (!map.has(key)) {
            map.set(key, { group: [], agg: [] });
        }
        var entry = map.get(key);
        entry.group.push(arr[i]);
        if (valueFn) {
            entry.agg.push(valueFn(arr[i], i));
        }
    }
    return map;
};
Array2D.z分组汇总到字典 = Array2D.groupIntoMap;

/**
 * 分组连接（groupIntoJoin）- 优化sumifs和Countifs批量条件统计
 * @param {Array} targetData - 统计目标数据（左表）
 * @param {Array} sourceData - 数据源（右表）
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} valueSelector - 汇总函数或选择器
 * @param {string} [separator='@^@'] - 多列分组时的分隔符
 * @returns {Array} 连接汇总后的结果
 * @example
 * // 对源数据按条件分类汇总，然后左连接到目标数据
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2', 'sum("f4")');
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2,f3', 'count(),sum("f4")', '@^@');
 * // 完整回调模式用法
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2', g => g.count());
 */
Array2D.groupIntoJoin = function(targetData, sourceData, keySelector, valueSelector, separator) {
    separator = separator || '@^@';
    
    // 1. 先对源数据做分类汇总
    var grouped = Array2D.groupInto(sourceData, keySelector, valueSelector, separator);
    
    // 2. 将汇总结果作为右表，与目标表做左连接
    return new Array2D(targetData).z左连接(
        grouped,
        keySelector,
        keySelector,
        function(leftRow, rightRow) {
            return leftRow.concat(rightRow || []);
        }
    ).val();
};
Array2D.z分组汇总连接 = Array2D.groupIntoJoin;

/**
 * 复制到指定位置（copyWithin）- 数组内部复制
 * @param {Array} arr - 数组
 * @param {Number} target - 目标位置
 * @param {Number} [start=0] - 源起始位置
 * @param {Number} [end] - 源结束位置
 * @returns {Array} 复制后的数组
 * @example
 * Array2D.copyWithin([[1,2],[3,4],[5,6]], 0, 2)  // [[5,6],[3,4],[5,6]]
 */
Array2D.copyWithin = function(arr, target, start, end) {
    if (!arr || !Array.isArray(arr)) return [];
    var result = JSON.parse(JSON.stringify(arr));
    var copyArr = result.slice(start || 0, end !== undefined ? end : result.length);
    for (var i = 0; i < copyArr.length; i++) {
        if (target + i < result.length) {
            result[target + i] = JSON.parse(JSON.stringify(copyArr[i]));
        }
    }
    return result;
};
Array2D.z复制到指定位置 = Array2D.copyWithin;

/**
 * 随机一项（random）- 随机选择一行
 * @param {Array} arr - 数组
 * @returns {Array} 随机选择的行
 * @example
 * Array2D.random([[1,2],[3,4],[5,6]])  // 随机返回一行
 */
Array2D.random = function(arr) {
    if (!arr || !Array.isArray(arr) || arr.length === 0) return undefined;
    var idx = Math.floor(Math.random() * arr.length);
    return arr[idx];
};
Array2D.z随机一项 = Array2D.random;

/**
 * 随机打乱（shuffle）- Fisher-Yates 洗牌算法
 * @param {Array} arr - 数组
 * @returns {Array} 打乱后的数组
 * @example
 * Array2D.shuffle([[1,2],[3,4],[5,6]])
 */
Array2D.shuffle = function(arr) {
    if (!arr || !Array.isArray(arr)) return [];
    var result = JSON.parse(JSON.stringify(arr));
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return result;
};
Array2D.z随机打乱 = Array2D.shuffle;

/**
 * 重复N次（repeat）- 将数组重复N次
 * @param {Array} arr - 数组
 * @param {Number} count - 重复次数
 * @returns {Array} 重复后的数组
 * @example
 * Array2D.repeat([[1,2]], 3)  // [[1,2],[1,2],[1,2]]
 */
Array2D.repeat = function(arr, count) {
    if (!arr || !Array.isArray(arr)) return [];
    if (count <= 0) return [];
    var result = [];
    for (var i = 0; i < count; i++) {
        for (var j = 0; j < arr.length; j++) {
            result.push(JSON.parse(JSON.stringify(arr[j])));
        }
    }
    return result;
};
Array2D.z重复N次 = Array2D.repeat;

/**
 * 静态方法：选择列（返回 Array2D 对象，支持链式调用）
 * @param {Array|Array2D} arr - 二维数组或 Array2D 对象
 */
Array2D.z选择列 = function(arr, cols, newHeaders) {
    // 智能判断：如果是 Array2D 对象，直接调用实例方法
    if (arr && arr instanceof Array2D) {
        return arr.z选择列(cols, newHeaders);
    }
    return new Array2D(arr).z选择列(cols, newHeaders);
};
Array2D.selectCols = Array2D.z选择列;

/**
 * 版本号（version）- 返回Array2D函数库版本号
 * @returns {String} 版本号
 * @example
 * Array2D.version()  // "3.2.0"
 */
Array2D.version = function() {
    return '3.2.0';
};

/**
 * 静态方法：数量（count）- 计算数组的元素数量
 * @param {Array} arr - 数组
 * @returns {Number} 元素数量
 * @example
 * Array2D.count([[1,2],[3,4]])  // 4
 */
Array2D.count = function(arr) {
    if (!arr || !Array.isArray(arr)) return 0;
    var count = 0;
    for (var i = 0; i < arr.length; i++) {
        if (Array.isArray(arr[i])) {
            count += arr[i].length;
        } else {
            count++;
        }
    }
    return count;
};
Array2D.z数量 = Array2D.count;

/**
 * 静态方法：批量填充
 */
Array2D.z批量填充 = function(arr, value, rows, cols) {
    return new Array2D(arr).z填充(value, rows, cols).val();
};
Array2D.fill = Array2D.z批量填充;

/**
 * 静态方法：写入单元格（根据数组大小自动扩展区域，返回 Range 对象）
 * @param {Array} arr - 二维数组
 * @param {Range|string} rng - 目标单元格区域（左上角单元格）
 * @returns {Range} 写入的 Range 对象
 * @example
 * var arr = [[1, 'A'], [2, 'B'], [3, 'C']];
 * Array2D.toRange(arr, "Sheet1!a1");
 * Array2D.toRange(arr, "e1");
 * var rs = Array2D.toRange(arr, Range("i1"));
 * console.log(rs.Address());  // $I$1:$J$3
 */
Array2D.toRange = function(arr, rng) {
    if (!isWPS) return null;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = arr.length;
    var cols = rows > 0 ? (Array.isArray(arr[0]) ? arr[0].length : 1) : 0;
    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = arr;
    return writeRng;
};

/**
 * 静态方法：写入单元格（中文别名，返回 Range 对象）
 */
Array2D.z写入单元格 = Array2D.toRange;

// ==================== 静态方法封装（支持直接调用）====================

/**
 * 静态方法：筛选（filter）- 根据条件筛选数组行
 * @param {Array} arr - 二维数组
 * @param {String|Function} predicate - 筛选条件
 * @param {Number} skipHeader - 跳过表头行数
 * @returns {Array} 筛选后的二维数组
 * @example
 * Array2D.filter(arr, 'f1>1')
 * Array2D.filter(arr, x=>x.f1>5 && x.f2=="A")
 * Array2D.filter(arr, "[f1,f3,f4]")
 */
Array2D.filter = function(arr, predicate, skipHeader) {
    return new Array2D(arr).z筛选(predicate, skipHeader).val();
};
Array2D.z筛选 = Array2D.filter;

/**
 * 静态方法：映射（map）- 对数组的每行进行转换
 * @param {Array} arr - 二维数组
 * @param {String|Function} mapper - 转换函数
 * @returns {Array} 转换后的二维数组
 * @example
 * Array2D.map(arr, 'f1*2')
 * Array2D.map(arr, x=>[x.f1, x.f3])
 * Array2D.map(arr, "[f1,f3]")
 */
Array2D.map = function(arr, mapper) {
    return new Array2D(arr).z映射(mapper).val();
};
Array2D.z映射 = Array2D.map;

/**
 * 静态方法：去重（distinct）- 根据指定列去重
 * @param {Array} arr - 二维数组
 * @param {String|Function} keySelector - 去重依据的列
 * @param {String} resultSelector - 结果选择器
 * @returns {Array} 去重后的二维数组
 * @example
 * Array2D.distinct(arr, 'f1,f2')
 * Array2D.distinct(arr, x=>x.f1)
 * Array2D.distinct(arr)
 */
Array2D.distinct = function(arr, keySelector, resultSelector) {
    return new Array2D(arr).z去重(keySelector, resultSelector).val();
};
Array2D.z去重 = Array2D.distinct;

/**
 * 静态方法：多列排序（sortByCols）- 按多列排序
 * @param {Array} arr - 二维数组
 * @param {String} colsConfig - 列配置，如 'f1+,f2-,f3+'
 * @param {Number} skipHeader - 表头行数
 * @returns {Array} 排序后的二维数组
 * @example
 * Array2D.sortByCols(arr, 'f1+,f2-', 1)
 */
Array2D.sortByCols = function(arr, colsConfig, skipHeader) {
    return new Array2D(arr).z多列排序(colsConfig, skipHeader).val();
};
Array2D.z多列排序 = Array2D.sortByCols;

/**
 * 静态方法：自定义排序（sortByList）- 按自定义列表排序
 * @param {Array} arr - 二维数组
 * @param {String|Number} col - 列号或列名
 * @param {String} orderList - 排序顺序，如 "A,B,C"
 * @param {Number} skipHeader - 表头行数
 * @returns {Array} 排序后的二维数组
 * @example
 * Array2D.sortByList(arr, 'f3', '美国,德国,中国')
 */
Array2D.sortByList = function(arr, col, orderList, skipHeader) {
    return new Array2D(arr).z自定义排序(col, orderList, skipHeader).val();
};
Array2D.z自定义排序 = Array2D.sortByList;

/**
 * 静态方法：批量插入列（insertCols）- 在指定位置插入列
 * @param {Array} arr - 二维数组
 * @param {Number|Array} colPos - 插入位置或多个位置
 * @param {Array|String} values - 插入的值
 * @param {Number} totalCols - 总列数
 * @returns {Array} 插入列后的二维数组
 * @example
 * Array2D.insertCols(arr, 2, ['新列1','新列2'])
 */
Array2D.insertCols = function(arr, colPos, values, totalCols) {
    return new Array2D(arr).z批量插入列(colPos, values, totalCols).val();
};
Array2D.z批量插入列 = Array2D.insertCols;

/**
 * 静态方法：批量删除列（deleteCols/delCols）- 删除指定列
 * @param {Array} arr - 二维数组
 * @param {String|Number|Array} cols - 列配置
 * @returns {Array} 删除列后的二维数组
 * @example
 * Array2D.deleteCols(arr, '1,3,5')
 * Array2D.delCols(arr, [0, 2, 4])
 */
Array2D.deleteCols = function(arr, cols) {
    return new Array2D(arr).z批量删除列(cols).val();
};
Array2D.z批量删除列 = Array2D.deleteCols;
Array2D.delCols = Array2D.deleteCols;

/**
 * 静态方法：左连接（leftjoin）- 类似SQL的LEFT JOIN
 * @param {Array} arr - 左表
 * @param {Array} brr - 右表
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的二维数组
 * @example
 * Array2D.leftjoin(arr, brr, 'f1', 'f1', 'f1,f2,f4')
 */
Array2D.leftjoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z左连接(brr, leftKey, rightKey, resultSelector).val();
};
Array2D.z左连接 = Array2D.leftjoin;

/**
 * 静态方法：排除（except）- 获取在arr中但不在brr中的元素
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 差异数组
 * @example
 * Array2D.except(arr, brr)
 */
Array2D.except = function(arr, brr) {
    return new Array2D(arr).z排除(brr).val();
};
Array2D.z排除 = Array2D.except;

/**
 * 静态方法：交集（intersect）- 获取arr和brr的交集
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 交集数组
 * @example
 * Array2D.intersect(arr, brr)
 */
Array2D.intersect = function(arr, brr) {
    return new Array2D(arr).z取交集(brr).val();
};
Array2D.z取交集 = Array2D.intersect;

/**
 * 静态方法：并集（union）- 获取arr和brr的并集并去重
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 并集数组
 * @example
 * Array2D.union(arr, brr)
 */
Array2D.union = function(arr, brr) {
    return new Array2D(arr).z去重并集(brr).val();
};
Array2D.z去重并集 = Array2D.union;

/**
 * 静态方法：最大值（max）- 获取数组最大值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 最大值
 * @example
 * Array2D.max(arr)
 * Array2D.max(arr, 'f1')
 */
Array2D.max = function(arr, selector) {
    var result = new Array2D(arr).z最大值(selector);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z最大值 = Array2D.max;

/**
 * 静态方法：最小值（min）- 获取数组最小值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 最小值
 * @example
 * Array2D.min(arr)
 * Array2D.min(arr, 'f1')
 */
Array2D.min = function(arr, selector) {
    var result = new Array2D(arr).z最小值(selector);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z最小值 = Array2D.min;

/**
 * 静态方法：文本连接（textjoin）- 选择指定列的值，用分隔符连接
 * @param {Array} arr - 二维数组
 * @param {String|Number|Function} selector - 列选择器，如 'f1' 或 0 或 row=>row.col
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D.textjoin([['a','b'],['c','d']], 1, '+')  // "b+d"
 * Array2D.textjoin([['a','b'],['c','d']], 'f2', '+')  // "b+d"
 */
Array2D.textjoin = function(arr, selector, separator = ',') {
    return new Array2D(arr).z文本连接(selector, separator);
};
Array2D.z文本连接 = Array2D.textjoin;

/**
 * Array 原型方法：textjoin - 为普通数组添加 textjoin 方法
 * 这样 .res() 返回的数组也可以使用 .textjoin()
 */
if (!Array.prototype.textjoin) {
    Array.prototype.textjoin = function(selector, separator = ',') {
        return Array2D.textjoin(this, selector, separator);
    };
}

/**
 * Array 原型方法：toRange - 为普通数组添加 toRange 方法
 * 这样 .res() 返回的数组也可以使用 .toRange()
 */
if (!Array.prototype.toRange) {
    Array.prototype.toRange = function(rng) {
        return Array2D.toRange(this, rng);
    };
}

/**
 * Array 原型方法：getRange - 为普通数组添加 getRange 方法
 * 这样 .res() 返回的数组也可以使用 .getRange()
 */
if (!Array.prototype.getRange) {
    Array.prototype.getRange = function(rng) {
        return Array2D.toRange(this, rng);
    };
}

/**
 * 静态方法：平均值（average）- 获取数组平均值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 平均值
 * @example
 * Array2D.average(arr)
 * Array2D.average(arr, 'f1')
 */
Array2D.average = function(arr, selector) {
    var result = new Array2D(arr).z平均值(selector);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z平均值 = Array2D.average;

/**
 * 静态方法：第一个（first）- 获取第一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 筛选条件
 * @returns {Array} 第一个元素（行）
 * @example
 * Array2D.first(arr)
 * Array2D.first(arr, 'f1>5')
 */
Array2D.first = function(arr, predicate) {
    var result = predicate ? new Array2D(arr).z第一个(predicate) : new Array2D(arr).z第一个();
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z第一个 = Array2D.first;

/**
 * 静态方法：最后一个（last）- 获取最后一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 筛选条件
 * @returns {Array} 最后一个元素（行）
 * @example
 * Array2D.last(arr)
 * Array2D.last(arr, 'f1>5')
 */
Array2D.last = function(arr, predicate) {
    var result = predicate ? new Array2D(arr).z最后一个(predicate) : new Array2D(arr).z最后一个();
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z最后一个 = Array2D.last;

/**
 * 静态方法：跳过（skip）- 跳过前N个元素
 * @param {Array} arr - 数组
 * @param {Number} count - 跳过的数量
 * @returns {Array} 剩余数组
 * @example
 * Array2D.skip(arr, 5)
 */
Array2D.skip = function(arr, count) {
    return new Array2D(arr).z跳过(count).val();
};
Array2D.z跳过 = Array2D.skip;

/**
 * 静态方法：取前N个（take）- 获取前N个元素
 * @param {Array} arr - 数组
 * @param {Number} count - 获取的数量
 * @returns {Array} 取出的数组
 * @example
 * Array2D.take(arr, 10)
 */
Array2D.take = function(arr, count) {
    return new Array2D(arr).z取前N个(count).val();
};
Array2D.z取前N个 = Array2D.take;

/**
 * 静态方法：补齐数组（pad）- 补齐数组使所有行列数一致
 * @param {Array} arr - 数组
 * @param {Number} cols - 目标列数
 * @param {Number} rows - 目标行数
 * @param {*} fillValue - 填充值
 * @returns {Array} 补齐后的数组
 * @example
 * Array2D.pad(arr, 5, 10)
 * Array2D.pad(arr)  // 自动按最大列补齐
 */
Array2D.pad = function(arr, cols, rows, fillValue) {
    return new Array2D(arr).z补齐数组(cols, rows, fillValue).val();
};
Array2D.z补齐数组 = Array2D.pad;

/**
 * 静态方法：查找（find）- 查找符合条件的第一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Array} 找到的元素
 * @example
 * Array2D.find(arr, 'f1==5')
 * Array2D.find(arr, x=>x.f1>10)
 */
Array2D.find = function(arr, predicate) {
    var result = new Array2D(arr).z查找单个(predicate);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z查找单个 = Array2D.find;

/**
 * 静态方法：查找索引（findIndex）- 查找符合条件的第一个元素索引
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Number} 元素索引
 * @example
 * Array2D.findIndex(arr, 'f1==5')
 */
Array2D.findIndex = function(arr, predicate) {
    return new Array2D(arr).z查找索引(predicate);
};
Array2D.z查找索引 = Array2D.findIndex;

/**
 * 静态方法：按行数分页（pageByRows）- 将数组按指定行数分页
 * @param {Array} arr - 数组
 * @param {Number} pageSize - 每页行数
 * @param {Number} pageNumber - 页码（从1开始）
 * @returns {Array} 分页后的数组
 * @example
 * Array2D.pageByRows(arr, 10, 2)
 */
Array2D.pageByRows = function(arr, pageSize, pageNumber) {
    return new Array2D(arr).z按行数分页(pageSize, pageNumber).val();
};
Array2D.z按行数分页 = Array2D.pageByRows;

/**
 * 静态方法：按页数分页（pageByCount）- 将数组平均分成指定页数
 * @param {Array} arr - 数组
 * @param {Number} pageCount - 总页数
 * @param {Number} pageNumber - 页码（从1开始）
 * @returns {Array} 分页后的数组
 * @example
 * Array2D.pageByCount(arr, 5, 2)
 */
Array2D.pageByCount = function(arr, pageCount, pageNumber) {
    return new Array2D(arr).z按页数分页(pageCount, pageNumber).val();
};
Array2D.z按页数分页 = Array2D.pageByCount;

/**
 * 静态方法：填充空白（fillBlank）- 填充合并单元格的空白区域
 * @param {Array} arr - 数组
 * @param {String} direction - 填充方向 'up'/'down'/'left'/'right'
 * @param {String} rangeAddress - 区域地址
 * @returns {Array} 填充后的数组
 * @example
 * Array2D.fillBlank(arr, 'up', 'A2:D2')
 */
Array2D.fillBlank = function(arr, direction, rangeAddress) {
    return new Array2D(arr).z补齐空位(direction, rangeAddress).val();
};
Array2D.z补齐空位 = Array2D.fillBlank;

/**
 * 静态方法：转矩阵（toMatrix）- 将数组转换为矩阵格式
 * @param {Array} arr - 数组
 * @param {Number} rows - 行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'/'c'
 * @returns {Array} 矩阵数组
 * @example
 * Array2D.toMatrix(arr, 3, 4, 'r')
 */
Array2D.toMatrix = function(arr, rows, cols, direction) {
    return new Array2D(arr).z转矩阵(rows, cols, direction).val();
};
Array2D.z转矩阵 = Array2D.toMatrix;

/**
 * 静态方法：查找所有行下标（findRowsIndex）- 查找符合条件的所有行索引
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Array} 行索引数组
 * @example
 * Array2D.findRowsIndex(arr, 'f1=="A"')
 */
Array2D.findRowsIndex = function(arr, predicate) {
    return new Array2D(arr).z查找所有行下标(predicate);
};
Array2D.z查找所有行下标 = Array2D.findRowsIndex;

/**
 * 静态方法：排序（sort）- 基本排序
 * @param {Array} arr - 数组
 * @param {String|Function} comparer - 比较函数
 * @returns {Array} 排序后的数组
 * @example
 * Array2D.sort(arr)
 * Array2D.sort(arr, 'f1+')
 */
Array2D.sort = function(arr, comparer) {
    return new Array2D(arr).z升序排序(comparer).val();
};
Array2D.z升序排序 = Array2D.sort;

/**
 * 静态方法：降序排序（sortDesc）
 * @param {Array} arr - 数组
 * @param {String|Function} comparer - 比较函数
 * @returns {Array} 排序后的数组
 * @example
 * Array2D.sortDesc(arr, 'f1-')
 */
Array2D.sortDesc = function(arr, comparer) {
    return new Array2D(arr).z降序排序(comparer).val();
};
Array2D.z降序排序 = Array2D.sortDesc;

/**
 * 静态方法：求和（sum）- 计算数组元素的和
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 和
 * @example
 * Array2D.sum([1,2,3,4])
 * Array2D.sum(arr, 'f1')
 */
Array2D.sum = function(arr, selector) {
    var result = new Array2D(arr).z求和(selector);
    return typeof result === 'object' && result.val ? result.val() : result;
};

/**
 * 静态方法：归约（reduce）- 对数组进行归约操作
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @param {*} initialValue - 初始值
 * @returns {*} 归约结果
 * @example
 * Array2D.reduce(arr, (acc, row) => acc + row[0], 0)
 */
Array2D.reduce = function(arr, callback, initialValue) {
    return new Array2D(arr).z归约(callback, initialValue);
};
Array2D.z归约 = Array2D.reduce;

/**
 * 静态方法：倒序归约（reduceRight）- 从右向左归约
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @param {*} initialValue - 初始值
 * @returns {*} 归约结果
 * @example
 * Array2D.reduceRight(arr, (acc, row) => acc + row[0], 0)
 */
Array2D.reduceRight = function(arr, callback, initialValue) {
    return new Array2D(arr).z倒序归约(callback, initialValue);
};
Array2D.z倒序归约 = Array2D.reduceRight;

/**
 * 静态方法：遍历（forEach）- 遍历数组的每一行
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @returns {Array} 原数组
 * @example
 * Array2D.forEach(arr, (row, i) => console.log(i, row))
 */
Array2D.forEach = function(arr, callback) {
    var instance = new Array2D(arr);
    instance.forEach(callback);
    return arr;
};

/**
 * 静态方法：倒序遍历（forEachRev）- 从后向前遍历
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @returns {Array} 原数组
 * @example
 * Array2D.forEachRev(arr, (row, i) => console.log(i, row))
 */
Array2D.forEachRev = function(arr, callback) {
    new Array2D(arr).z倒序遍历执行(callback);
    return arr;
};
Array2D.z倒序遍历执行 = Array2D.forEachRev;

/**
 * 静态方法：有满足（some）- 检查是否有元素满足条件
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 条件
 * @returns {Boolean} 是否有满足
 * @example
 * Array2D.some(arr, 'f1>5')
 */
Array2D.some = function(arr, predicate) {
    return new Array2D(arr).z有满足(predicate);
};
Array2D.z有满足 = Array2D.some;

/**
 * 静态方法：全部满足（every）- 检查是否所有元素都满足条件
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 条件
 * @returns {Boolean} 是否全部满足
 * @example
 * Array2D.every(arr, 'f1>0')
 */
Array2D.every = function(arr, predicate) {
    return new Array2D(arr).z全部满足(predicate);
};
Array2D.z全部满足 = Array2D.every;

/**
 * 静态方法：降维（flat）- 将二维数组降维为一维
 * @param {Array} arr - 二维数组
 * @param {Function} mapper - 可选的映射函数
 * @returns {Array} 一维数组
 * @example
 * Array2D.flat(arr)
 * Array2D.flat(arr, x=>x.f1)
 */
Array2D.flat = function(arr, mapper) {
    var result = new Array2D(arr);
    return mapper ? result.z扁平化(mapper) : result.z扁平化();
};
Array2D.z扁平化 = Array2D.flat;

/**
 * 静态方法：行切片删除（splice）- 删除/插入元素
 * @param {Array} arr - 数组
 * @param {Number} start - 起始位置
 * @param {Number} deleteCount - 删除数量
 * @param {...*} items - 要插入的元素
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.splice(arr, 2, 1, ['新行'])
 */
Array2D.splice = function(arr, start, deleteCount) {
    var items = Array.prototype.slice.call(arguments, 3);
    return new Array2D(arr).z行切片删除行(start, deleteCount, items);
};
Array2D.z行切片删除行 = Array2D.splice;

/**
 * 静态方法：追加一项（push）- 在数组末尾添加元素
 * @param {Array} arr - 数组
 * @param {*} item - 要添加的元素
 * @returns {Number} 新长度
 * @example
 * Array2D.push(arr, [1,2,3])
 */
Array2D.push = function(arr, item) {
    new Array2D(arr).z追加一项(item);
    return arr.length;
};
Array2D.z追加一项 = Array2D.push;

/**
 * 静态方法：尾部弹出一项（pop）- 删除并返回最后一个元素
 * @param {Array} arr - 数组
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.pop(arr)
 */
Array2D.pop = function(arr) {
    return new Array2D(arr).z尾部弹出一项();
};
Array2D.z尾部弹出一项 = Array2D.pop;

/**
 * 静态方法：删除第一个（shift）- 删除并返回第一个元素
 * @param {Array} arr - 数组
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.shift(arr)
 */
Array2D.shift = function(arr) {
    return new Array2D(arr).z删除第一个();
};
Array2D.z删除第一个 = Array2D.shift;

/**
 * 静态方法：反转（reverse）- 反转数组顺序
 * @param {Array} arr - 数组
 * @returns {Array} 反转后的数组
 * @example
 * Array2D.reverse(arr)
 */
Array2D.reverse = function(arr) {
    return new Array2D(arr).z反转().val();
};
Array2D.z反转 = Array2D.reverse;

/**
 * 静态方法：文本连接（join）- 用分隔符连接所有元素
 * @param {Array} arr - 数组
 * @param {String} separator - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D.join(arr, ',')
 */
Array2D.join = function(arr, separator) {
    return new Array2D(arr).z连接(separator);
};
Array2D.z连接 = Array2D.join;

/**
 * 静态方法：转JSON字符串（toJson）- 将数组转为JSON字符串
 * @param {Array} arr - 数组
 * @param {Number|String} indent - 缩进
 * @returns {String} JSON字符串
 * @example
 * Array2D.toJson(arr, 2)
 */
Array2D.toJson = function(arr, indent) {
    return new Array2D(arr).z转JSON(indent);
};
Array2D.z转JSON = Array2D.toJson;

/**
 * 静态方法：转字符串（toString）- 将数组转为字符串
 * @param {Array} arr - 数组
 * @param {String} separator - 分隔符
 * @returns {String} 字符串
 * @example
 * Array2D.toString(arr, ',')
 */
Array2D.toString = function(arr, separator) {
    return new Array2D(arr).z转字符串(separator);
};
Array2D.z转字符串 = Array2D.toString;

/**
 * 静态方法：是否为空（isEmpty）- 检查数组是否为空
 * @param {Array} arr - 数组
 * @returns {Boolean} 是否为空
 * @example
 * Array2D.isEmpty(arr)
 */
Array2D.isEmpty = function(arr) {
    return new Array2D(arr).z是否为空();
};
Array2D.z是否为空 = Array2D.isEmpty;

/**
 * 静态方法：分组（groupBy）- 按指定条件分组
 * @param {Array} arr - 数组
 * @param {String|Function} keySelector - 分组依据
 * @returns {Map} 分组结果
 * @example
 * Array2D.groupBy(arr, 'f1')
 */
Array2D.groupBy = function(arr, keySelector) {
    return new Array2D(arr).z分组(keySelector);
};
Array2D.z分组 = Array2D.groupBy;

/**
 * 静态方法：左右全连接（fulljoin）- 类似SQL的FULL OUTER JOIN
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.fulljoin(arr, brr, 'f1', 'f1', 'f1,f2,f3')
 */
Array2D.fulljoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z左右全连接(brr, leftKey, rightKey, resultSelector).val();
};
Array2D.z左右全连接 = Array2D.fulljoin;

/**
 * 静态方法：一对多连接（leftFulljoin）- 左表一对多连接
 * @param {Array} arr - 左表
 * @param {Array} brr - 右表
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.leftFulljoin(arr, brr, 'f1', 'f1', 'f1,f2,f3')
 */
Array2D.leftFulljoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z一对多连接(brr, leftKey, rightKey, resultSelector).val();
};
Array2D.z一对多连接 = Array2D.leftFulljoin;

/**
 * 静态方法：超级查找（superLookup）- 增强版VLOOKUP
 * @param {Array} arr - 查找范围
 * @param {*} lookupValue - 查找值
 * @param {Number|String} colIndex - 列号
 * @param {Number|String} returnCol - 返回列号
 * @returns {Array} 查找结果
 * @example
 * Array2D.superLookup(arr, 'A', 1, 3)
 */
Array2D.superLookup = function(arr, lookupValue, colIndex, returnCol) {
    var result = new Array2D(arr).z超级查找(lookupValue, colIndex, returnCol);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z超级查找 = Array2D.superLookup;

/**
 * 静态方法：左右连接（zip）- 将两个数组的对应位置元素配对
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 配对后的数组
 * @example
 * Array2D.zip(arr1, arr2)
 */
Array2D.zip = function(arr, brr) {
    return new Array2D(arr).z左右连接(brr).val();
};
Array2D.z左右连接 = Array2D.zip;

/**
 * 静态方法：转置（transpose）- 转置二维数组
 * @param {Array} arr - 二维数组
 * @returns {Array} 转置后的数组
 * @example
 * Array2D.transpose([[1,2],[3,4]])  // [[1,3],[2,4]]
 */
Array2D.transpose = function(arr) {
    return new Array2D(arr).z转置().val();
};

/**
 * 静态方法：克隆（copy）- 深拷贝数组
 * @param {Array} arr - 数组
 * @returns {Array} 拷贝后的数组
 * @example
 * Array2D.copy(arr)
 */
Array2D.copy = function(arr) {
    return new Array2D(arr).z克隆().val();
};
Array2D.z克隆 = Array2D.copy;

/**
 * 静态方法：上下连接（concat）- 连接多个数组
 * @param {Array} arr - 数组1
 * @param {...Array} arrays - 其他数组
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.concat(arr1, arr2, arr3)
 */
Array2D.concat = function(arr) {
    var arrays = Array.prototype.slice.call(arguments, 1);
    return new Array2D(arr).z上下连接.apply(new Array2D(arr), arrays).val();
};
Array2D.z上下连接 = Array2D.concat;

/**
 * 静态方法：选择行（selectRows）- 选择指定行
 * @param {Array} arr - 数组
 * @param {String|Array} rows - 行配置
 * @returns {Array} 选择后的数组
 * @example
 * Array2D.selectRows(arr, '1,3,5')
 * Array2D.selectRows(arr, [0, 2, 4])
 */
Array2D.selectRows = function(arr, rows) {
    return new Array2D(arr).z选择行(rows).val();
};
Array2D.z选择行 = Array2D.selectRows;

/**
 * 静态方法：删除行（deleteRows）- 删除指定行
 * @param {Array} arr - 数组
 * @param {String|Array} rows - 行配置
 * @returns {Array} 删除后的数组
 * @example
 * Array2D.deleteRows(arr, '1,3,5')
 */
Array2D.deleteRows = function(arr, rows) {
    return new Array2D(arr).z批量删除行(rows).val();
};
Array2D.z批量删除行 = Array2D.deleteRows;

/**
 * 静态方法：插入行（insertRows）- 在指定位置插入行
 * @param {Array} arr - 数组
 * @param {Number|Array} rowPos - 插入位置
 * @param {Array} values - 插入的值
 * @returns {Array} 插入后的数组
 * @example
 * Array2D.insertRows(arr, 2, [[1,2,3]])
 */
Array2D.insertRows = function(arr, rowPos, values) {
    return new Array2D(arr).z批量插入行(rowPos, values).val();
};
Array2D.z批量插入行 = Array2D.insertRows;

/**
 * 静态方法：插入行号（insertRowNum）- 在数组前插入行号列
 * @param {Array} arr - 数组
 * @param {Number} start - 起始行号
 * @param {String} title - 列标题
 * @returns {Array} 插入行号后的数组
 * @example
 * Array2D.insertRowNum(arr, 1, '序号')
 */
Array2D.insertRowNum = function(arr, start, title) {
    return new Array2D(arr).z插入行号(start, title).val();
};
Array2D.z插入行号 = Array2D.insertRowNum;

/**
 * 静态方法：是否包含值（includes）- 检查数组是否包含某值
 * @param {Array} arr - 数组
 * @param {*} value - 要检查的值
 * @returns {Boolean} 是否包含
 * @example
 * Array2D.includes(arr, [1,2])
 */
Array2D.includes = function(arr, value) {
    return new Array2D(arr).z包含(value);
};
Array2D.z包含 = Array2D.includes;

/**
 * 静态方法：值位置（indexOf）- 查找元素的位置
 * @param {Array} arr - 数组
 * @param {*} value - 要查找的值
 * @returns {Number} 元素索引
 * @example
 * Array2D.indexOf(arr, [1,2])
 */
Array2D.indexOf = function(arr, value) {
    return new Array2D(arr).z值位置(value);
};

/**
 * 静态方法：从后往前值位置（lastIndexOf）- 从后向前查找元素位置
 * @param {Array} arr - 数组
 * @param {*} value - 要查找的值
 * @returns {Number} 元素索引
 * @example
 * Array2D.lastIndexOf(arr, [1,2])
 */
Array2D.lastIndexOf = function(arr, value) {
    return new Array2D(arr).z从后往前值位置(value);
};

/**
 * 静态方法：中位数（median）- 计算中位数
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 中位数
 * @example
 * Array2D.median(arr)
 * Array2D.median(arr, 'f1')
 */
Array2D.median = function(arr, selector) {
    var result = new Array2D(arr).z中位数(selector);
    return typeof result === 'object' && result.val ? result.val() : result;
};
Array2D.z中位数 = Array2D.median;

/**
 * 静态方法：间隔取数（nth）- 每隔n个取一个
 * @param {Array} arr - 数组
 * @param {Number} n - 间隔
 * @param {Number} offset - 偏移量
 * @returns {Array} 取出的数组
 * @example
 * Array2D.nth(arr, 3, 0)  // 每3个取1个
 */
Array2D.nth = function(arr, n, offset) {
    return new Array2D(arr).z间隔取数(n, offset).val();
};
Array2D.z间隔取数 = Array2D.nth;

/**
 * 静态方法：行切片（slice）- 提取指定范围的行
 * @param {Array} arr - 数组
 * @param {Number} start - 起始位置
 * @param {Number} end - 结束位置
 * @returns {Array} 切片后的数组
 * @example
 * Array2D.slice(arr, 1, 5)
 */
Array2D.slice = function(arr, start, end) {
    return new Array2D(arr).z行切片(start, end).val();
};

/**
 * 静态方法：结果（res）- 获取结果数组
 * @param {Array} arr - 数组
 * @returns {Array} 结果数组
 * @example
 * Array2D.res(arr)
 */
Array2D.res = function(arr) {
    return new Array2D(arr).z结果();
};
Array2D.z结果 = Array2D.res;

// ==================== 超级透视表（superPivot）====================
/**
 * 超级透视（z超级透视）- 将二维数组仿透视表生成行列字段，并进行各种汇总统计的交叉表
 * @param {Array} arr - 二维数组
 * @param {Array|string} rowFields - 行字段配置，如 ['f1+,f2-'] 或 ['f1,f2', '标题']
 * @param {Array|string} colFields - 列字段配置，如 ['f5+,f6+'] 或 ['f2', '标题']
 * @param {Array|string} dataFields - 数据字段配置，如 ['count(),sum("f3")'] 或 [[回调数组], '标题']
 * @param {Number} headerRows - 标题行数，默认1
 * @param {string|Number} outputHeader - 1:输出表头, 0:不输出, 'map':返回字典，默认1
 * @param {string} separator - 分隔符，默认"@^@"
 * @returns {Array|Map} 返回二维数组或Map
 * @example
 * // 示例1：基本透视（带排序符号）
 * var rs = Array2D.z超级透视(arr, ['f1+,f2-'], ['f5+,f6+'], ['count(),sum("f3")']);
 *
 * // 示例2：带标题的透视
 * var rs = Array2D.z超级透视(arr, ['f1,f5,f6','prod,year,month'], ['f2','country'], ['count(),sum("f3"),average("f4")','count,sum,avg']);
 *
 * // 示例3：回调函数模式 + Map返回
 * var rs = Array2D.z超级透视(arr, ['f1,f5,f6','期数,年,月'], ['f2','国家'], [[g=>g.count(),g=>g.sum("f3")],'计数,求和'], 2, 'map');
 */
Array2D.z超级透视 = function(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator) {
    separator = separator || '@^@';
    headerRows = headerRows !== undefined ? headerRows : 1;
    outputHeader = outputHeader !== undefined ? outputHeader : 1;

    // 处理 Array2D 对象
    if (arr && typeof arr === 'object' && arr._items && Array.isArray(arr._items)) {
        arr = arr._items;
    }

    // 辅助函数：将行数组转为带f1,f2...属性的对象
    function toRowObject(row) {
        var obj = Array(row.length);
        for (var i = 0; i < row.length; i++) {
            obj['f' + (i + 1)] = row[i];
            obj[i] = row[i];
        }
        return obj;
    }

    // 辅助函数：创建分组对象，支持聚合操作
    function createGroupObject(group) {
        return {
            _items: group,
            count: function() { return group.length; },
            sum: function(col) {
                var total = 0;
                for (var i = 0; i < group.length; i++) {
                    var val = group[i][col];
                    if (typeof val === 'number') {
                        total += val;
                    } else if (typeof val === 'string') {
                        var num = parseFloat(val.replace(/,/g, ''));
                        if (!isNaN(num)) total += num;
                    }
                }
                return total;
            },
            average: function(col) {
                if (group.length === 0) return 0;
                var sum = this.sum(col);
                return sum / group.length;
            },
            max: function(col) {
                var max = null;
                for (var i = 0; i < group.length; i++) {
                    var val = group[i][col];
                    if (typeof val === 'string') {
                        val = parseFloat(val.replace(/,/g, ''));
                    }
                    if (typeof val === 'number' && !isNaN(val)) {
                        if (max === null || val > max) max = val;
                    }
                }
                return max;
            },
            min: function(col) {
                var min = null;
                for (var i = 0; i < group.length; i++) {
                    var val = group[i][col];
                    if (typeof val === 'string') {
                        val = parseFloat(val.replace(/,/g, ''));
                    }
                    if (typeof val === 'number' && !isNaN(val)) {
                        if (min === null || val < min) min = val;
                    }
                }
                return min;
            },
            textjoin: function(col, sep) {
                var values = [];
                for (var i = 0; i < group.length; i++) {
                    var val = group[i][col];
                    values.push(val);
                }
                return values.join(sep);
            }
        };
    }

    // 辅助函数：解析结果选择器字符串
    function parseResultSelector(str) {
        var operations = [];
        var regex = /(\w+)\s*\(([^)]*)\)/g;
        var match;
        while ((match = regex.exec(str)) !== null) {
            var op = { name: match[1] };
            var argsStr = match[2].trim();
            if (argsStr) {
                // 解析参数，支持带引号和不带引号
                op.args = [];
                var argRegex = /["']([^"']+)["']|([^,]+)/g;
                var argMatch;
                while ((argMatch = argRegex.exec(argsStr)) !== null) {
                    op.args.push(argMatch[1] || argMatch[2]);
                }
            }
            operations.push(op);
        }
        return operations;
    }

    // 解析字段配置
    function parseFieldsConfig(fieldsConfig) {
        var fields = [];
        var titles = [];
        var hasTitles = false;

        if (Array.isArray(fieldsConfig)) {
            // 先检查 [[回调数组], '标题'] 格式 - 数据字段
            if (fieldsConfig.length === 2 && Array.isArray(fieldsConfig[0])) {
                return {
                    fields: [{ callbacks: fieldsConfig[0] }],
                    titles: fieldsConfig[1].split(','),
                    hasTitles: true,
                    isCallback: true
                };
            }
            // 检查 ['f1,f2,f3', '标题'] 格式 - 有排序符号的算无标题
            if (fieldsConfig.length === 2 && typeof fieldsConfig[0] === 'string' && typeof fieldsConfig[1] === 'string') {
                var fieldStr = fieldsConfig[0];
                var items = fieldStr.split(',');
                var hasSortSymbol = false;
                for (var i = 0; i < items.length; i++) {
                    if (items[i].match(/[+\-#]$/)) {
                        hasSortSymbol = true;
                        break;
                    }
                }
                if (!hasSortSymbol) {
                    // 没有排序符号，是带标题格式
                    titles = fieldsConfig[1].split(',');
                    for (var j = 0; j < items.length; j++) {
                        var item = items[j].trim();
                        fields.push({
                            field: item,
                            sort: '+'
                        });
                    }
                    hasTitles = true;
                    return { fields: fields, titles: titles, hasTitles: hasTitles };
                }
            }
            // ['f1+,f2-'] 格式 或 带排序符号的格式
            if (typeof fieldsConfig[0] === 'string') {
                var fieldStr = fieldsConfig[0];
                var items = fieldStr.split(',');
                for (var k = 0; k < items.length; k++) {
                    var item = items[k].trim();
                    var match = item.match(/^(f\d+)([+\-#]*)$/);
                    if (match) {
                        fields.push({
                            field: match[1],
                            sort: match[2] || '+'
                        });
                    }
                }
            }
        } else if (typeof fieldsConfig === 'string') {
            var items = fieldsConfig.split(',');
            for (var m = 0; m < items.length; m++) {
                var item = items[m].trim();
                var match = item.match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    fields.push({
                        field: match[1],
                        sort: match[2] || '+'
                    });
                }
            }
        }

        return { fields: fields, titles: titles, hasTitles: hasTitles };
    }

    var rowConfig = parseFieldsConfig(rowFields);
    var colConfig = parseFieldsConfig(colFields);
    // 数据字段需要特殊处理
    var dataConfig;
    if (Array.isArray(dataFields)) {
        if (dataFields.length === 2 && Array.isArray(dataFields[0])) {
            // [[回调数组], '标题'] 格式
            dataConfig = {
                fields: [{ callbacks: dataFields[0] }],
                titles: dataFields[1].split(','),
                hasTitles: true,
                isCallback: true,
                rawString: null
            };
        } else if (dataFields.length === 2 && typeof dataFields[0] === 'string') {
            // ['count(),sum("f3")', '标题'] 或 ['f1,f2', '标题'] 格式
            var dfStr = dataFields[0];
            // 检查是否包含聚合函数
            if (dfStr.match(/count|sum|average|max|min|textjoin/)) {
                // 数据字段格式
                dataConfig = {
                    fields: [{ field: dfStr }],
                    titles: dataFields[1].split(','),
                    hasTitles: true,
                    rawString: dfStr
                };
            } else {
                // 普通字段格式，使用 parseFieldsConfig
                dataConfig = parseFieldsConfig(dataFields);
            }
        } else if (typeof dataFields[0] === 'string') {
            // ['count(),sum("f3")'] 格式
            dataConfig = {
                fields: [{ field: dataFields[0] }],
                titles: [],
                hasTitles: false,
                rawString: dataFields[0]
            };
        } else {
            dataConfig = parseFieldsConfig(dataFields);
        }
    } else if (typeof dataFields === 'string') {
        // 字符串格式的数据字段
        dataConfig = {
            fields: [{ field: dataFields }],
            titles: [],
            hasTitles: false,
            rawString: dataFields
        };
    } else {
        dataConfig = parseFieldsConfig(dataFields);
    }

    // 跳过标题行
    var dataStartRow = headerRows || 1;
    var data = arr.slice(dataStartRow);

    // 将数据转为对象数组
    var dataObjs = data.map(function(row) {
        return toRowObject(row);
    });

    // 提取所有行字段值并排序
    var rowKeys = [];
    var rowKeyMap = Object.create(null);
    for (var i = 0; i < dataObjs.length; i++) {
        var obj = dataObjs[i];
        var keyParts = [];
        for (var j = 0; j < rowConfig.fields.length; j++) {
            var rf = rowConfig.fields[j];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                keyParts.push(obj[idx]);
            }
        }
        var key = keyParts.join(separator);
        if (!rowKeyMap[key]) {
            rowKeyMap[key] = {
                values: keyParts.slice(),
                originalIndex: i
            };
            rowKeys.push(key);
        }
    }

    // 对行键排序
    rowKeys.sort(function(a, b) {
        var aParts = a.split(separator);
        var bParts = b.split(separator);
        for (var k = 0; k < rowConfig.fields.length; k++) {
            var rf = rowConfig.fields[k];
            var aVal = aParts[k];
            var bVal = bParts[k];
            var cmp = 0;
            // 尝试转换为数字进行比较
            var aNum = parseFloat(aVal);
            var bNum = parseFloat(bVal);
            if (!isNaN(aNum) && !isNaN(bNum) && String(aNum) === String(aVal).trim() && String(bNum) === String(bVal).trim()) {
                cmp = aNum - bNum;
            } else {
                cmp = String(aVal).localeCompare(String(bVal));
            }
            if (cmp !== 0) {
                return rf.sort === '-' ? -cmp : rf.sort === '#' ? 0 : cmp;
            }
        }
        return 0;
    });

    // 提取所有列字段值并排序
    var colKeys = [];
    var colKeyMap = Object.create(null);
    for (var m = 0; m < dataObjs.length; m++) {
        var obj = dataObjs[m];
        var keyParts = [];
        for (var n = 0; n < colConfig.fields.length; n++) {
            var cf = colConfig.fields[n];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                keyParts.push(obj[idx]);
            }
        }
        var key = keyParts.join(separator);
        if (!colKeyMap[key]) {
            colKeyMap[key] = {
                values: keyParts.slice(),
                originalIndex: m
            };
            colKeys.push(key);
        }
    }

    // 对列键排序
    colKeys.sort(function(a, b) {
        var aParts = a.split(separator);
        var bParts = b.split(separator);
        for (var k = 0; k < colConfig.fields.length; k++) {
            var cf = colConfig.fields[k];
            var aVal = aParts[k];
            var bVal = bParts[k];
            var cmp = 0;
            var aNum = parseFloat(aVal);
            var bNum = parseFloat(bVal);
            if (!isNaN(aNum) && !isNaN(bNum) && String(aNum) === String(aVal).trim() && String(bNum) === String(bVal).trim()) {
                cmp = aNum - bNum;
            } else {
                cmp = String(aVal).localeCompare(String(bVal));
            }
            if (cmp !== 0) {
                return cf.sort === '-' ? -cmp : cf.sort === '#' ? 0 : cmp;
            }
        }
        return 0;
    });

    // 分组数据：行键 + 列键 -> 数据行
    var groupMap = Object.create(null);
    for (var q = 0; q < dataObjs.length; q++) {
        var obj = dataObjs[q];
        var rowKeyParts = [];
        for (var r = 0; r < rowConfig.fields.length; r++) {
            var rf = rowConfig.fields[r];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                rowKeyParts.push(obj[idx]);
            }
        }
        var colKeyParts = [];
        for (var s = 0; s < colConfig.fields.length; s++) {
            var cf = colConfig.fields[s];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                colKeyParts.push(obj[idx]);
            }
        }
        var rowKey = rowKeyParts.join(separator);
        var colKey = colKeyParts.join(separator);
        var fullKey = rowKey + '|||' + colKey;
        if (!groupMap[fullKey]) {
            groupMap[fullKey] = [];
        }
        // 转回普通数组
        var row = [];
        for (var t = 0; t < obj.length; t++) {
            row.push(obj[t]);
        }
        groupMap[fullKey].push(row);
    }

    // 解析数据字段操作
    var dataOps = [];
    if (dataConfig.isCallback) {
        dataOps = dataConfig.fields[0].callbacks;
    } else if (dataConfig.rawString) {
        var operations = parseResultSelector(dataConfig.rawString);
        dataOps = operations;
    } else if (dataConfig.fields.length > 0) {
        var opStr = dataConfig.fields[0].field || '';
        var operations = parseResultSelector(opStr);
        dataOps = operations;
    }

    // 执行聚合操作
    function executeAggregation(group) {
        var groupObj = createGroupObject(group.map(function(r) {
            return toRowObject(r);
        }));
        var results = [];
        if (Array.isArray(dataOps) && dataOps.length > 0 && typeof dataOps[0] === 'function') {
            // 回调模式
            for (var v = 0; v < dataOps.length; v++) {
                var result = dataOps[v](groupObj);
                results.push(result);
            }
        } else {
            // 字符串模式
            for (var w = 0; w < dataOps.length; w++) {
                var op = dataOps[w];
                var args = op.args || [];
                switch (op.name) {
                    case 'count':
                        results.push(groupObj.count());
                        break;
                    case 'sum':
                        results.push(groupObj.sum(args[0]));
                        break;
                    case 'average':
                        results.push(groupObj.average(args[0]));
                        break;
                    case 'max':
                        results.push(groupObj.max(args[0]));
                        break;
                    case 'min':
                        results.push(groupObj.min(args[0]));
                        break;
                    case 'textjoin':
                        results.push(groupObj.textjoin(args[0], args[1]));
                        break;
                }
            }
        }
        return results;
    }

    // map 模式：返回查询标准字典
    if (outputHeader === 'map') {
        var resultMap = new Map();
        for (var x = 0; x < rowKeys.length; x++) {
            var rowKey = rowKeys[x];
            for (var y = 0; y < colKeys.length; y++) {
                var colKey = colKeys[y];
                var fullKey = rowKey + '|||' + colKey;
                if (groupMap[fullKey]) {
                    var agg = executeAggregation(groupMap[fullKey]);
                    var sortKey = rowKey + separator + colKey;
                    var mapKey = '01L' + String(x + 1).padStart(4, '0') + ' ' + sortKey;
                    resultMap.set(mapKey, {
                        agg: agg,
                        group: { '00000': groupMap[fullKey][0] }
                    });
                }
            }
        }
        return resultMap;
    }

    // 构建透视表
    var numDataFields = Array.isArray(dataOps) && dataOps.length > 0 ? dataOps.length :
                       (dataConfig.titles && dataConfig.titles.length > 0 ? dataConfig.titles.length : 1);

    // 生成默认数据字段标题
    var defaultDataTitles = [];
    if (Array.isArray(dataOps) && dataOps.length > 0 && typeof dataOps[0] !== 'function') {
        // 根据聚合函数生成默认标题
        var opNameMap = {
            'count': '计数',
            'sum': '求和',
            'average': '平均',
            'max': '最大',
            'min': '最小',
            'textjoin': '连接'
        };
        for (var i = 0; i < dataOps.length; i++) {
            var opName = dataOps[i].name;
            defaultDataTitles.push(opNameMap[opName] || opName);
        }
    } else if (dataConfig.titles && dataConfig.titles.length > 0) {
        // 使用自定义标题
        defaultDataTitles = dataConfig.titles.slice();
    } else {
        // 默认标题
        for (var j = 0; j < numDataFields; j++) {
            defaultDataTitles.push('值' + (j + 1));
        }
    }

    var result = [];

    // 构建表头
    if (outputHeader === 1 || outputHeader === true) {
        // 检查是否有自定义标题
        var hasRowTitles = rowConfig.hasTitles && rowConfig.titles.length > 0;
        var hasColTitles = colConfig.hasTitles && colConfig.titles.length > 0;
        var hasDataTitles = dataConfig.hasTitles && dataConfig.titles.length > 0;

        if (hasRowTitles) {
            // 有自定义行标题时的表头格式（示例2格式）
            // 为每个列字段创建一行：行字段占位 + 列字段标题 + 列键值（每个重复numDataFields次）
            for (var cfIdx = 0; cfIdx < colConfig.fields.length; cfIdx++) {
                var colHeaderRow = [];
                // 行字段占位（比行字段数量少1，因为列字段标题占1列）
                for (var rfIdx = 0; rfIdx < rowConfig.fields.length - 1; rfIdx++) {
                    colHeaderRow.push('');
                }
                // 列字段标题
                if (hasColTitles) {
                    colHeaderRow.push(colConfig.titles[cfIdx]);
                } else {
                    colHeaderRow.push('');
                }
                // 列键的对应字段值（每个重复 numDataFields 次）
                for (var ckIdx = 0; ckIdx < colKeys.length; ckIdx++) {
                    var colKeyParts = colKeys[ckIdx].split(separator);
                    var colFieldVal = colKeyParts[cfIdx];
                    // 尝试转换为数字
                    var numVal = parseFloat(colFieldVal);
                    if (!isNaN(numVal) && String(numVal) === String(colFieldVal).trim()) {
                        colFieldVal = numVal;
                    }
                    for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                        colHeaderRow.push(colFieldVal);
                    }
                }
                result.push(colHeaderRow);
            }

            // 第2行：行字段标题 + 数据字段标题（重复）
            var dataTitleRow = [];
            // 行字段标题
            for (var rtIdx = 0; rtIdx < rowConfig.titles.length; rtIdx++) {
                dataTitleRow.push(rowConfig.titles[rtIdx]);
            }
            // 数据字段标题（为每列重复）
            for (var colIdx = 0; colIdx < colKeys.length; colIdx++) {
                for (var dtIdx = 0; dtIdx < defaultDataTitles.length; dtIdx++) {
                    dataTitleRow.push(defaultDataTitles[dtIdx]);
                }
            }
            result.push(dataTitleRow);
        } else {
            // 没有自定义行标题时的表头格式（示例1格式）
            // 第1-colConfig.fields.length行：列字段标题和列值
            for (var cfIdx = 0; cfIdx < colConfig.fields.length; cfIdx++) {
                var headerRow = [];
                // 为每个行字段添加空白占位符（最后一个位置留给列字段标题）
                for (var rfIdx = 0; rfIdx < rowConfig.fields.length - 1; rfIdx++) {
                    headerRow.push('');
                }
                // 列字段标题（从原数组获取）
                var cf = colConfig.fields[cfIdx];
                var match = cf.field.match(/^f(\d+)$/);
                if (match && arr && arr[0]) {
                    var origIdx = parseInt(match[1]) - 1;
                    headerRow.push(arr[0][origIdx] || '');
                } else {
                    headerRow.push('');
                }
                // 列键的对应字段值
                for (var ckIdx = 0; ckIdx < colKeys.length; ckIdx++) {
                    var colKeyParts = colKeys[ckIdx].split(separator);
                    var colFieldVal = colKeyParts[cfIdx];
                    // 尝试转换为数字
                    var numVal = parseFloat(colFieldVal);
                    if (!isNaN(numVal) && String(numVal) === String(colFieldVal).trim()) {
                        colFieldVal = numVal;
                    }
                    // 为每个数据字段重复
                    for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                        headerRow.push(colFieldVal);
                    }
                }
                result.push(headerRow);
            }

            // 下一行：行字段名 + 数据字段名
            var dataTitleRow = [];
            // 行字段标题（从原数组获取）
            var maxRowIdx = arr && arr[0] ? arr[0].length : 0;
            for (var rhIdx = 0; rhIdx < rowConfig.fields.length && rhIdx < maxRowIdx; rhIdx++) {
                var match = rowConfig.fields[rhIdx].field.match(/^f(\d+)$/);
                if (match && arr && arr[0]) {
                    var origIdx = parseInt(match[1]) - 1;
                    dataTitleRow.push(arr[0][origIdx] || '');
                } else {
                    dataTitleRow.push('');
                }
            }
            // 数据字段标题（为每列重复）
            for (var colIdx = 0; colIdx < colKeys.length; colIdx++) {
                for (var dtIdx = 0; dtIdx < defaultDataTitles.length; dtIdx++) {
                    dataTitleRow.push(defaultDataTitles[dtIdx]);
                }
            }
            result.push(dataTitleRow);
        }
    }

    // 构建数据行
    for (var rk = 0; rk < rowKeys.length; rk++) {
        var rowKey = rowKeys[rk];
        var rowKeyParts = rowKey.split(separator);
        var dataRow = rowKeyParts.slice();

        for (var ck = 0; ck < colKeys.length; ck++) {
            var colKey = colKeys[ck];
            var fullKey = rowKey + '|||' + colKey;
            if (groupMap[fullKey]) {
                var agg = executeAggregation(groupMap[fullKey]);
                dataRow = dataRow.concat(agg);
            } else {
                for (var c = 0; c < numDataFields; c++) {
                    dataRow.push('');
                }
            }
        }
        result.push(dataRow);
    }

    // 包装结果，返回 Array2D 对象，添加 toRange 和 getRange 方法
    var wrappedResult = result;
    if (Array.isArray(result)) {
        // 创建 Array2D 对象
        wrappedResult = new Array2D(result);

        /**
         * toRange - 将结果写入单元格
         * @param {Range|string} rng - 目标单元格
         * @returns {Range} Range对象
         */
        wrappedResult.toRange = function(rng) {
            return Array2D.toRange(result, rng);
        };

        /**
         * getRange - 获取结果写入后的Range对象
         * @param {Range|string} rng - 目标单元格
         * @returns {Range} Range对象
         */
        wrappedResult.getRange = function(rng) {
            return Array2D.toRange(result, rng);
        };

        /**
         * val - 获取原始数组
         * @returns {Array} 原始数组
         */
        wrappedResult.val = function() { return result; };

        /**
         * res - 获取原始数组（val的别名）
         * @returns {Array} 原始数组
         */
        wrappedResult.res = function() { return result; };
    }

    return wrappedResult;
};
Array2D.superPivot = Array2D.z超级透视;

/**
 * z超级透视 - 实例方法版本
 * 调用静态方法 Array2D.z超级透视，使用当前实例的数据
 */
Array2D.prototype.z超级透视 = function(rowFields, colFields, dataFields, headerRows, outputHeader, separator) {
    return Array2D.z超级透视(this._items, rowFields, colFields, dataFields, headerRows, outputHeader, separator);
};
Array2D.prototype.superPivot = Array2D.prototype.z超级透视;

/**
 * 生成静态方法（从实例方法自动生成）
 */
(function() {
    var propNames = Object.getOwnPropertyNames(Array2D.prototype);
    // 已经手动定义的静态方法，跳过自动生成
    var manuallyDefined = ['z选择列', 'selectCols', 'z批量填充', 'fill', 'z写入单元格', 'toRange', 'z转置', 'transpose', 'z求和', 'sum', 'z克隆', 'copy', 'z超级透视', 'superPivot'];

    for (var i = 0; i < propNames.length; i++) {
        var name = propNames[i];
        if (manuallyDefined.indexOf(name) >= 0) continue;

        if (name !== 'constructor' && name !== '_init' && name !== '_new' && typeof Array2D.prototype[name] === 'function') {
            (function(methodName) {
                Array2D[methodName] = function() {
                    // 第一个参数是数组数据，传递给构造函数
                    // 支持 Array2D 对象（提取 _items）或普通数组
                    var firstArg = arguments.length > 0 ? arguments[0] : null;
                    if (firstArg && typeof firstArg === 'object' && firstArg._items && Array.isArray(firstArg._items)) {
                        firstArg = firstArg._items;
                    }
                    var instance = new Array2D(firstArg);
                    // 剩余参数传递给实例方法
                    var restArgs = [];
                    for (var j = 1; j < arguments.length; j++) {
                        restArgs.push(arguments[j]);
                    }
                    var result = instance[methodName].apply(instance, restArgs);
                    // 如果结果是 Array2D 对象，返回纯数组
                    if (result && typeof result === 'object' && result._items && Array.isArray(result._items)) {
                        return result._items;
                    }
                    return result;
                };
            })(name);
        }
    }
})();

// ==================== RngUtils - Range工具库 ====================

/**
 * RngUtils - Range区域操作工具（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description WPS Range区域操作增强工具
 * @example
 * RngUtils("A1:C10").z加边框().z自动列宽()
 */
function RngUtils(initialRange) {
    if (!(this instanceof RngUtils)) {
        return new RngUtils(initialRange);
    }
    this._range = initialRange ? this._toRange(initialRange) : null;
}

/**
 * 转换为Range对象
 * @private
 * @param {Range|string} rng - Range对象或地址字符串
 * @returns {Range|null} Range对象或null
 */
RngUtils.prototype._toRange = function(rng) {
    if (!rng) return null;
    if (typeof rng === 'string') return isWPS ? Range(rng) : null;
    return rng;
};

/**
 * 获取/设置Range
 * @param {Range|string} newRange - 新Range
 * @returns {RngUtils|Range} 设置时返回this，否则返回当前Range
 */
RngUtils.prototype.rng = function(newRange) {
    if (newRange !== undefined) {
        this._range = this._toRange(newRange);
        return this;
    }
    return this._range;
};

/**
 * 获取值
 * @returns {Array} 二维数组
 */
RngUtils.prototype.val = function() {
    if (!this._range) return null;
    return this._range.Value2;
};

// ==================== 基础信息函数 ====================

/**
 * z最后一个 - 获取指定区域的最后一个单元格
 * @param {Range|string} rng - 单元格区域
 * @returns {Range} 最后一个单元格
 * @example
 * RngUtils.z最后一个("A1:A13")  // $A$13
 */
RngUtils.z最后一个 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Cells(r.Rows.Count, r.Columns.Count);
};
RngUtils.lastCell = RngUtils.z最后一个;

/**
 * z安全区域 - 获取当前区域与UsedRange的交集
 * @param {Range|string} rng - 单元格区域
 * @returns {Range} 交集单元格
 * @example
 * RngUtils.z安全区域("A:A")  // $A$1:$A$13
 */
RngUtils.z安全区域 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return r;
    return Application.Intersect(r, usedRange);
};
RngUtils.safeRange = RngUtils.z安全区域;

/**
 * z安全数组 - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @example
 * RngUtils.z安全数组("A1:A13").filter(row => row[0] > 0).toRange("C1")
 */
RngUtils.z安全数组 = function(rng) {
    if (!isWPS) return new Array2D([]);
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var arr = r.Value2;
    if (arr === null || arr === undefined) return new Array2D([]);
    // 单个单元格转二维数组
    if (!Array.isArray(arr)) return new Array2D([[arr]]);
    // 一维数组转二维
    if (!Array.isArray(arr[0])) {
        var result = [];
        for (var i = 0; i < arr.length; i++) {
            result.push([arr[i]]);
        }
        return new Array2D(result);
    }
    return new Array2D(arr);
};
RngUtils.safeArray = RngUtils.z安全数组;

/**
 * z最大行 - 获取指定区域的最大行数
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @example
 * RngUtils.z最大行("A:A")     // 70 (单列，从下往上查找第一个有效数据)
 * RngUtils.z最大行("A1")      // 70 (单单元格，自动扩展为整列)
 * RngUtils.z最大行("A1:C10")  // 10 (多单元格区域返回该区域的最后一行)
 */
RngUtils.z最大行 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return 0;

    // 单单元格或单列时，从下往上查找第一个有效数据
    if (r.Columns.Count === 1) {
        var col = r.Columns.Count === 1 && r.Rows.Count === 1 ? sheet.Columns(r.Column) : r;
        var safe = Application.Intersect(col, usedRange);
        if (!safe) return 0;
        // 从下往上查找第一个非空单元格
        for (var i = safe.Rows.Count; i >= 1; i--) {
            var cell = safe.Cells(i, 1);
            var val = cell.Value2;
            // 跳过 null、undefined、空字符串（包括 =""）
            if (val === null || val === undefined || val === '') {
                continue;
            }
            // 跳过纯空白字符
            if (typeof val === 'string' && val.trim() === '') {
                continue;
            }
            // 找到第一个有效数据，返回行号
            return safe.Row + i - 1;
        }
        return 0;
    }

    // 多列区域，返回该区域与UsedRange交集的最后一行
    var safe = Application.Intersect(r, usedRange);
    if (!safe) return 0;
    return safe.Row + safe.Rows.Count - 1;
};
RngUtils.endRow = RngUtils.z最大行;

/**
 * z最大行单元格 - 获取指定区域最后一行的单元格
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 最后一行的单元格
 * @example
 * RngUtils.z最大行单元格("A1:A1000")  // $A$13
 */
RngUtils.z最大行单元格 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var maxRow = RngUtils.z最大行(r);
    var col = r.Column;
    return sheet.Cells(maxRow, col);
};
RngUtils.endRowCell = RngUtils.z最大行单元格;

/**
 * z最大行区域 - 获取从第一行到最后一行的区域
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Range} 从第一行到最后一行的区域
 * @example
 * RngUtils.z最大行区域("1:1000","A")  // $1:$13
 * RngUtils.z最大行区域("A1:J1")       // A1:J最大行
 */
RngUtils.z最大行区域 = function(rng, col) {
    if (!isWPS) return null;
    col = col !== undefined ? col : "A";
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;

    // 特殊参数处理
    if (col === '-c') {
        // CurrentRegion
        return r.CurrentRegion;
    }
    if (col === '-u') {
        // UsedRange
        var used = sheet.UsedRange;
        if (!used) return r;
        var startRow = r.Row;
        var endRow = used.Row + used.Rows.Count - 1;
        var startCol = r.Column;
        var endCol = r.Column + r.Columns.Count - 1;
        return sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(endRow, endCol));
    }

    // 整行处理 - 当rng是整行时（如 "1:1000"）
    if (r.Rows.Count >= 16384) {
        var colNum = typeof col === 'string' ? (col.charCodeAt(0) - 64) : (col || 1);
        var maxR = RngUtils.z最大行(sheet.Columns(colNum));
        return sheet.Range(sheet.Cells(1, colNum), sheet.Cells(maxR, colNum)).EntireRow;
    }

    // 默认情况 - 保持原区域的列范围，扩展行到最后一行
    // 需要找出范围内所有列的最大使用行数
    var startRow = r.Row;
    var startCol = r.Column;
    var endCol = r.Column + r.Columns.Count - 1;

    // 遍历每一列，找出最大使用行数
    var maxEndRow = startRow;
    for (var c = startCol; c <= endCol; c++) {
        var colRange = sheet.Columns(c);
        var endRow = RngUtils.z最大行(colRange);
        if (endRow > maxEndRow) {
            maxEndRow = endRow;
        }
    }

    return sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(maxEndRow, endCol));
};

/**
 * maxRange - 获取从第一行到最后一行的区域（英文别名）
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Range} 从第一行到最后一行的区域
 * @example
 * RngUtils.maxRange("1:1000","A")  // $1:$13
 * RngUtils.maxRange("A1:J1")       // A1:J最大行
 */
/**
 * maxRange - 获取从第一行到最后一行的区域（英文别名，返回 RangeChain 支持智能提示）
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {RangeChain} 从第一行到最后一行的区域（支持链式调用和智能提示）
 * @example
 * RngUtils.maxRange("1:1000","A").safeArray()  // 返回数组
 * RngUtils.maxRange("A1:J1").z加边框()         // 链式调用
 */
RngUtils.maxRange = function(rng, col) {
    var result = RngUtils.z最大行区域.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * z最大列 - 获取指定区域的最大列数
 * @param {Range|string} rng - 要获取最大列数的区域
 * @returns {number} 最大列数
 * @example
 * RngUtils.z最大列("1:1")     // 8 (单行，从右往左查找第一个有效数据)
 * RngUtils.z最大列("A1")      // 8 (单单元格，自动扩展为整行)
 * RngUtils.z最大列("A1:C10")  // 3 (多行区域返回该区域的最后一列)
 */
RngUtils.z最大列 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return 0;

    // 单单元格或单行时，从右往左查找第一个有效数据
    if (r.Rows.Count === 1) {
        var row = r.Rows.Count === 1 && r.Columns.Count === 1 ? sheet.Rows(r.Row) : r;
        var safe = Application.Intersect(row, usedRange);
        if (!safe) return 0;
        // 从右往左查找第一个非空单元格
        for (var i = safe.Columns.Count; i >= 1; i--) {
            var cell = safe.Cells(1, i);
            var val = cell.Value2;
            // 跳过 null、undefined、空字符串（包括 =""）
            if (val === null || val === undefined || val === '') {
                continue;
            }
            // 跳过纯空白字符
            if (typeof val === 'string' && val.trim() === '') {
                continue;
            }
            // 找到第一个有效数据，返回列号
            return safe.Column + i - 1;
        }
        return 0;
    }

    // 多行区域，返回该区域与UsedRange交集的最后一列
    var safe = Application.Intersect(r, usedRange);
    if (!safe) return 0;
    return safe.Column + safe.Columns.Count - 1;
};
RngUtils.endCol = RngUtils.z最大列;

/**
 * z最大列单元格 - 获取指定区域最后一列的单元格
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 最后一列的单元格
 * @example
 * RngUtils.z最大列单元格("1:1")  // $F$1
 */
RngUtils.z最大列单元格 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var maxCol = RngUtils.z最大列(r);
    return sheet.Cells(r.Row, maxCol);
};
RngUtils.endColCell = RngUtils.z最大列单元格;

/**
 * z可见区数组 - 将可见单元格转换为数组
 * @param {Range|string} rng - 要转换的区域
 * @param {Worksheet} [tempSheet] - 临时工作表（可选）
 * @returns {Array} 可见单元格值的数组
 * @example
 * RngUtils.z可见区数组("1:4")
 */
RngUtils.z可见区数组 = function(rng, tempSheet) {
    if (!isWPS) return [];
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var visible = r.SpecialCells(12); // xlCellTypeVisible
    if (!visible) return [];
    var arr = visible.Value2;
    // 保存到临时表
    if (tempSheet) {
        tempSheet.Range("A1").Resize(visible.Rows.Count, visible.Columns.Count).Value2 = arr;
    }
    return RngUtils.z安全数组(arr);
};
RngUtils.visibleArray = RngUtils.z可见区数组;

/**
 * z可见区域 - 获取指定区域的可见区域
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 可见区域
 * @example
 * RngUtils.z可见区域("1:4")  // $1:$4
 */
RngUtils.z可见区域 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.SpecialCells(12); // xlCellTypeVisible
};
RngUtils.visibleRange = RngUtils.z可见区域;

/**
 * z加边框 - 为指定区域添加边框
 * @param {Range|string} rng - 要添加边框的区域
 * @param {number} [LineStyle=1] - 边框线条样式
 * @param {number} [Weight=2] - 边框线条粗细
 * @returns {Borders} 边框对象
 * @example
 * RngUtils.z加边框("A3:D7")
 */
RngUtils.z加边框 = function(rng, LineStyle, Weight) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    LineStyle = LineStyle !== undefined ? LineStyle : 1;
    Weight = Weight !== undefined ? Weight : 2;
    r.Borders.LineStyle = LineStyle;
    r.Borders.Weight = Weight;
    return r.Borders;
};
RngUtils.addBorders = RngUtils.z加边框;

/**
 * z取前几行 - 获取指定区域的前几行
 * @param {Range|string} rng - 指定区域
 * @param {number} count - 获取的行数
 * @returns {Range} 前几行的单元格
 * @example
 * RngUtils.z取前几行("a3:d7",3)  // $A$3:$D$5
 */
RngUtils.z取前几行 = function(rng, count) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Rows("1:" + count);
};
RngUtils.takeRows = RngUtils.z取前几行;

/**
 * z跳过前几行 - 跳过指定区域的前几行
 * @param {Range|string} rng - 指定区域
 * @param {number} count - 要跳过的行数
 * @returns {Range} 跳过后的单元格区域
 * @example
 * RngUtils.z跳过前几行("a3:d7",3)  // $A$6:$D$7
 */
RngUtils.z跳过前几行 = function(rng, count) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var startRow = count + 1;
    var endRow = r.Rows.Count;
    if (startRow > endRow) return null;
    return r.Rows(startRow + ":" + endRow);
};
RngUtils.skipRows = RngUtils.z跳过前几行;

/**
 * z合并相同单元格 - 合并指定区域中相同的行或列
 * @param {Range|string} rng - 要合并的单元格区域
 * @param {string} [direction='r'] - 合并方式: -r按行 -c按列 -rm先行后列 -cm列按上下级关系
 * @example
 * RngUtils.z合并相同单元格("a1:j1","c")
 */
RngUtils.z合并相同单元格 = function(rng, direction) {
    if (!isWPS) return;
    direction = direction || 'r';
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // 如果传入的是单个单元格，自动扩展到当前选区
    if (r.Rows.Count === 1 && r.Columns.Count === 1) {
        console.log("检测到单个单元格，扩展到当前选区...");
        r = Selection;
        console.log("当前选区:", r.Address());
    }

    // 添加调试信息
    console.log("=== mergeCells 调试 ===");
    console.log("区域:", r.Address());
    console.log("方向:", direction);
    console.log("行数:", r.Rows.Count);
    console.log("列数:", r.Columns.Count);

    if (direction === '-r' || direction === 'r') {
        // 按行合并
        var mergedCount = 0;
        for (var i = 1; i <= r.Rows.Count; i++) {
            var startCol = 1;
            for (var j = 2; j <= r.Columns.Count; j++) {
                var currentVal = r.Cells(i, j).Value2;
                var startVal = r.Cells(i, startCol).Value2;
                // 调试：输出前3行的合并信息
                if (i <= 3 && j <= 5) {
                    console.log("行" + i + " 列" + j + ": [" + currentVal + "] vs [" + startVal + "]");
                }
                if (currentVal !== startVal) {
                    if (j - startCol > 1) {
                        var mergeRng = r.Range(r.Cells(i, startCol), r.Cells(i, j - 1));
                        mergeRng.Merge();
                        mergedCount++;
                        if (i <= 3) console.log("  -> 合并 列" + startCol + "-" + (j-1));
                    }
                    startCol = j;
                }
            }
            if (r.Columns.Count - startCol + 1 > 1) {
                r.Range(r.Cells(i, startCol), r.Cells(i, r.Columns.Count)).Merge();
                mergedCount++;
                if (i <= 3) console.log("  -> 合并 列" + startCol + "-" + r.Columns.Count);
            }
        }
        console.log("按行合并完成，共合并 " + mergedCount + " 次");
    } else if (direction === '-c' || direction === 'c') {
        // 按列合并
        console.log("开始按列合并...");
        var mergedCount = 0;
        for (var j = 1; j <= r.Columns.Count; j++) {
            var startRow = 1;
            for (var i = 2; i <= r.Rows.Count; i++) {
                var currentVal = r.Cells(i, j).Value2;
                var startVal = r.Cells(startRow, j).Value2;
                // 调试：输出前3列的合并信息
                if (j <= 3 && i <= 5) {
                    console.log("列" + j + " 行" + i + ": [" + currentVal + "] vs [" + startVal + "]");
                }
                if (currentVal !== startVal) {
                    if (i - startRow > 1) {
                        var mergeRng = r.Range(r.Cells(startRow, j), r.Cells(i - 1, j));
                        mergeRng.Merge();
                        mergedCount++;
                        if (j <= 3) console.log("  -> 合并 行" + startRow + "-" + (i-1));
                    }
                    startRow = i;
                }
            }
            if (r.Rows.Count - startRow + 1 > 1) {
                r.Range(r.Cells(startRow, j), r.Cells(r.Rows.Count, j)).Merge();
                mergedCount++;
                if (j <= 3) console.log("  -> 合并 行" + startRow + "-" + r.Rows.Count);
            }
        }
        console.log("按列合并完成，共合并 " + mergedCount + " 次");
    } else if (direction === '-rm') {
        // 先行后列，按上下级关系
        // 先按列合并（考虑前一列）
        for (var j = 1; j <= r.Columns.Count; j++) {
            var prevCol = j - 1;
            for (var i = 1; i <= r.Rows.Count; i++) {
                var startRow = i;
                var currentValue = r.Cells(i, j).Value2;
                var prevValue = prevCol >= 1 ? r.Cells(i, prevCol).Value2 : null;

                for (var k = i + 1; k <= r.Rows.Count; k++) {
                    var nextValue = r.Cells(k, j).Value2;
                    var nextPrev = prevCol >= 1 ? r.Cells(k, prevCol).Value2 : null;
                    if (nextValue !== currentValue || (prevCol >= 1 && nextPrev !== prevValue)) {
                        break;
                    }
                    startRow = k;
                }
                if (startRow > i) {
                    r.Range(r.Cells(i, j), r.Cells(startRow, j)).Merge();
                    i = startRow;
                }
            }
        }
    } else if (direction === '-cm') {
        // 列按上下级关系
        for (var j = 1; j <= r.Columns.Count; j++) {
            var prevRow = j - 1;
            for (var i = 1; i <= r.Rows.Count; i++) {
                var startRow = i;
                var currentValue = r.Cells(i, j).Value2;
                var prevValue = prevRow >= 1 ? r.Cells(i, prevRow).Value2 : null;

                for (var k = i + 1; k <= r.Rows.Count; k++) {
                    var nextValue = r.Cells(k, j).Value2;
                    var nextPrev = prevRow >= 1 ? r.Cells(k, prevRow).Value2 : null;
                    if (nextValue !== currentValue || (prevRow >= 1 && nextPrev !== prevValue)) {
                        break;
                    }
                    startRow = k;
                }
                if (startRow > i) {
                    r.Range(r.Cells(i, j), r.Cells(startRow, j)).Merge();
                    i = startRow;
                }
            }
        }
    }
};
RngUtils.mergeCells = RngUtils.z合并相同单元格;

/**
 * z取消合并填充单元格 - 取消合并并填充
 * @param {Range|string} rng - 要取消合并的单元格区域
 * @param {boolean} [fillAll=true] - true:所有行填充 false:仅首行填充
 * @example
 * RngUtils.z取消合并填充单元格("a1:j1")
 */
RngUtils.z取消合并填充单元格 = function(rng, fillAll) {
    if (!isWPS) return;
    fillAll = fillAll !== undefined ? fillAll : true;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    for (var i = 1; i <= r.Rows.Count; i++) {
        for (var j = 1; j <= r.Columns.Count; j++) {
            var cell = r.Cells(i, j);
            if (cell.MergeCells) {
                var mergeArea = cell.MergeArea;
                var value = cell.Value2;
                mergeArea.UnMerge();
                if (fillAll) {
                    mergeArea.Value2 = value;
                } else {
                    cell.Value2 = value;
                }
            }
        }
    }
};
RngUtils.unMergeCells = RngUtils.z取消合并填充单元格;

// ==================== agg 聚合工具对象 ====================
/**
 * agg - 聚合工具对象，用于在普通数组上执行聚合操作
 * 在 superPivot 回调中使用，支持对分组后的数据进行聚合
 * @example
 * // 在 superPivot 中使用
 * Array2D.superPivot(arr, rows, cols, [
 *   [g => agg.sum(g, "f4"), g => agg.count(g)],
 *   ['sum', 'count']
 * ])
 * @namespace
 */
var agg = {
    /**
     * sum - 求和
     * @param {Array} arr - 二维数组
     * @param {string|number|function} selector - 列选择器 (如 "f4", 3, 或函数)
     * @returns {number} 求和结果
     * @example
     * agg.sum([[1,2],[3,4]], 1)  // 6 (第二列求和)
     * agg.sum([{f4:1}, {f4:2}], "f4")  // 3
     */
    sum: function(arr, selector) {
        return Array2D.sum(arr, selector);
    },

    /**
     * average - 求平均值
     * @param {Array} arr - 二维数组
     * @param {string|number|function} selector - 列选择器
     * @returns {number} 平均值
     */
    average: function(arr, selector) {
        return Array2D.average(arr, selector);
    },

    /**
     * count - 计数
     * @param {Array} arr - 数组
     * @returns {number} 数组长度
     */
    count: function(arr) {
        return arr ? arr.length : 0;
    },

    /**
     * max - 求最大值
     * @param {Array} arr - 二维数组
     * @param {string|number|function} selector - 列选择器
     * @returns {number} 最大值
     */
    max: function(arr, selector) {
        return Array2D.max(arr, selector);
    },

    /**
     * min - 求最小值
     * @param {Array} arr - 二维数组
     * @param {string|number|function} selector - 列选择器
     * @returns {number} 最小值
     */
    min: function(arr, selector) {
        return Array2D.min(arr, selector);
    },

    /**
     * textjoin - 文本连接
     * @param {Array} arr - 二维数组
     * @param {string|number|function} selector - 列选择器
     * @param {string} [separator=','] - 分隔符
     * @returns {string} 连接后的字符串
     * @example
     * agg.textjoin([['a','b'],['c','d']], 1, '+')  // "b+d"
     */
    textjoin: function(arr, selector, separator) {
        return Array2D.textjoin(arr, selector, separator);
    },

    /**
     * first - 获取第一个元素
     * @param {Array} arr - 数组
     * @returns {*} 第一个元素
     */
    first: function(arr) {
        if (!arr || !arr.length) return undefined;
        return arr[0];
    },

    /**
     * last - 获取最后一个元素
     * @param {Array} arr - 数组
     * @returns {*} 最后一个元素
     */
    last: function(arr) {
        if (!arr || !arr.length) return undefined;
        return arr[arr.length - 1];
    }
};

// 导出 agg 到全局
if (typeof module !== 'undefined' && module.exports) {
    module.exports.agg = agg;
}

/**
 * z插入多行 - 插入多行
 * @param {Range|string} rng - 要插入行的单元格区域
 * @param {any} value - 行号数组或字符串
 * @param {number} count - 要插入的行数
 * @example
 * RngUtils.z插入多行("a12:d15", '*', 2)
 */
RngUtils.z插入多行 = function(rng, value, count) {
    if (!isWPS) return;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    count = count || 1;

    for (var i = r.Rows.Count; i >= 1; i--) {
        var insertValue = value;
        if (Array.isArray(value)) {
            insertValue = value[i - 1] !== undefined ? value[i - 1] : '';
        }
        for (var c = 0; c < count; c++) {
            r.Rows(i).Insert();
            var newRow = r.Rows(i);
            for (var j = 1; j <= r.Columns.Count; j++) {
                newRow.Cells(1, j).Value2 = insertValue;
            }
        }
    }
};
RngUtils.insertRows = RngUtils.z插入多行;

/**
 * z插入多列 - 插入多列
 * @param {Range|string} rng - 要插入列的单元格区域
 * @param {any} value - 列号数组或字符串
 * @param {number} count - 要插入的列数
 * @example
 * RngUtils.z插入多列("a12:d14", '*', 2)
 */
RngUtils.z插入多列 = function(rng, value, count) {
    if (!isWPS) return;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    count = count || 1;

    for (var j = r.Columns.Count; j >= 1; j--) {
        var insertValue = value;
        if (Array.isArray(value)) {
            insertValue = value[j - 1] !== undefined ? value[j - 1][0] : '';
        }
        for (var c = 0; c < count; c++) {
            r.Columns(j).Insert();
            var newCol = r.Columns(j);
            for (var i = 1; i <= r.Rows.Count; i++) {
                newCol.Cells(i, 1).Value2 = insertValue;
            }
        }
    }
};
RngUtils.insertCols = RngUtils.z插入多列;

/**
 * z删除空白行 - 删除指定区域中的空白行
 * @param {Range|string} rng - 要删除空白行的单元格区域
 * @param {boolean} [entireColumn=true] - 默认删除整列 false时只作用选中区域
 * @example
 * RngUtils.z删除空白行("a11:d17")
 */
RngUtils.z删除空白行 = function(rng, entireColumn) {
    if (!isWPS) return;
    entireColumn = entireColumn !== undefined ? entireColumn : true;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    var blankRows = [];
    for (var i = r.Rows.Count; i >= 1; i--) {
        var row = r.Rows(i);
        var isEmpty = true;
        for (var j = 1; j <= r.Columns.Count; j++) {
            var val = row.Cells(1, j).Value2;
            if (val !== null && val !== undefined && val !== '') {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) {
            blankRows.push(i);
        }
    }

    for (var k = 0; k < blankRows.length; k++) {
        if (entireColumn) {
            r.Rows(blankRows[k]).EntireRow.Delete();
        } else {
            r.Rows(blankRows[k]).Delete();
        }
    }
};
RngUtils.delBlankRows = RngUtils.z删除空白行;

/**
 * z删除空白列 - 删除指定区域中的空白列
 * @param {Range|string} rng - 要删除空白列的单元格区域
 * @param {boolean} [entireColumn=true] - 默认删除整列 false时只作用选中区域
 * @example
 * RngUtils.z删除空白列("A11:G14")
 */
RngUtils.z删除空白列 = function(rng, entireColumn) {
    if (!isWPS) return;
    entireColumn = entireColumn !== undefined ? entireColumn : true;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    var blankCols = [];
    for (var j = r.Columns.Count; j >= 1; j--) {
        var col = r.Columns(j);
        var isEmpty = true;
        for (var i = 1; i <= r.Rows.Count; i++) {
            var val = col.Cells(i, 1).Value2;
            if (val !== null && val !== undefined && val !== '') {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) {
            blankCols.push(j);
        }
    }

    for (var k = 0; k < blankCols.length; k++) {
        if (entireColumn) {
            r.Columns(blankCols[k]).EntireColumn.Delete();
        } else {
            r.Columns(blankCols[k]).Delete();
        }
    }
};
RngUtils.delBlankCols = RngUtils.z删除空白列;

/**
 * z整行 - 获取指定单元格区域的整行
 * @param {Range|string} rng - 要获取整行的单元格区域
 * @returns {Range} 整行单元格
 * @example
 * RngUtils.z整行("11:14")  // $11:$14
 */
RngUtils.z整行 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.EntireRow;
};
RngUtils.entireRow = RngUtils.z整行;

/**
 * z整列 - 获取指定单元格区域的整列
 * @param {Range|string} rng - 要获取整列的单元格区域
 * @returns {Range} 整列单元格
 * @example
 * RngUtils.z整列("A:B")  // $A:$B
 */
RngUtils.z整列 = function(rng) {
    if (!isWPS) return null;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.EntireColumn;
};
RngUtils.entire_column = RngUtils.z整列;

/**
 * z行数 - 获取指定单元格区域的行数
 * @param {Range|string} rng - 要获取行数的单元格区域
 * @returns {number} 行数
 * @example
 * RngUtils.z行数("A12:D15")  // 4
 */
RngUtils.z行数 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Rows.Count;
};
RngUtils.rowsCount = RngUtils.z行数;

/**
 * z列数 - 获取指定单元格区域的列数
 * @param {Range|string} rng - 要获取列数的单元格区域
 * @returns {number} 列数
 * @example
 * RngUtils.z列数("A12:C15")  // 3
 */
RngUtils.z列数 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Columns.Count;
};
RngUtils.colsCount = RngUtils.z列数;

/**
 * z列号字母互转 - 将数字列号转换为字母表示
 * @param {number} c - 要转换的数字列号
 * @returns {string} 列号的字母表示
 * @example
 * RngUtils.z列号字母互转(3)  // "C"
 */
RngUtils.z列号字母互转 = function(c) {
    var result = '';
    while (c > 0) {
        c--;
        result = String.fromCharCode(65 + (c % 26)) + result;
        c = Math.floor(c / 26);
    }
    return result;
};
RngUtils.colToAbc = RngUtils.z列号字母互转;

/**
 * z复制粘贴格式 - 复制粘贴格式到目标区域
 * @param {Range|string} rng - 源单元格区域
 * @param {Range|string} target - 目标单元格区域
 * @example
 * RngUtils.z复制粘贴格式("a14:d14","a18:d21")
 */
RngUtils.z复制粘贴格式 = function(rng, target) {
    if (!isWPS) return;
    var src = typeof rng === 'string' ? Range(rng) : rng;
    var dest = typeof target === 'string' ? Range(target) : target;
    src.Copy();
    dest.PasteSpecial(-4122); // xlPasteFormats
    Application.CutCopyMode = false;
};
RngUtils.copyFormat = RngUtils.z复制粘贴格式;

/**
 * z复制粘贴值 - 复制粘贴值到目标区域
 * @param {Range|string} rng - 源单元格区域
 * @param {Range|string} target - 目标单元格区域
 * @example
 * RngUtils.z复制粘贴值("a11:d14","a18:d21")
 */
RngUtils.z复制粘贴值 = function(rng, target) {
    if (!isWPS) return;
    var src = typeof rng === 'string' ? Range(rng) : rng;
    var dest = typeof target === 'string' ? Range(target) : target;
    src.Copy();
    dest.PasteSpecial(-4163); // xlPasteValues
    Application.CutCopyMode = false;
};
RngUtils.copyValue = RngUtils.z复制粘贴值;

/**
 * z联合区域 - 对字符串地址或单元格数组联合成一个单元格区域
 * @param {any} rng - 单元格地址或单元格数组
 * @param {Sheet} [op_sht] - 工作表对象，跨表时指定
 * @returns {Range} 组合后的单元格对象
 * @example
 * RngUtils.z联合区域('a1,a2,B4:C10').Address()
 */
RngUtils.z联合区域 = function(rng, op_sht) {
    if (!isWPS) return null;
    var sheet = op_sht || Application.ActiveSheet;

    if (typeof rng === 'string') {
        // 解析地址字符串
        var parts = rng.split(',');
        var ranges = [];
        for (var i = 0; i < parts.length; i++) {
            ranges.push(sheet.Range(parts[i].trim()));
        }
        if (ranges.length === 1) return ranges[0];
        return sheet.Union(ranges[0], ranges[1]);
    }

    if (Array.isArray(rng)) {
        if (rng.length === 0) return null;
        if (rng.length === 1) return rng[0];
        return sheet.Union(rng[0], rng[1]);
    }

    return rng;
};
RngUtils.unionAll = RngUtils.z联合区域;

/**
 * z多列排序 - 单元格多列排序
 * @param {Range|string} rng - 待排序的单元格范围
 * @param {string} sortParams - 排序参数 'f3+,f4-' 表示第3列升序第4列降序
 * @param {number} [headerRows=1] - 表头的行数
 * @param {string} [customOrder] - 自定义序列
 * @example
 * RngUtils.z多列排序("A18:D24",'f3+,f4-',1)
 */
RngUtils.z多列排序 = function(rng, sortParams, headerRows, customOrder) {
    if (!isWPS) return;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    headerRows = headerRows || 1;

    // 解析排序参数
    var sorts = [];
    var parts = sortParams.split(',');
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        var match = part.match(/f?(\d+)([+-])/);
        if (match) {
            sorts.push({
                col: parseInt(match[1]),
                order: match[2] === '+' ? 1 : 2 // 1升序 2降序
            });
        }
    }

    // 获取数据数组
    var arr = RngUtils.z安全数组(r);
    if (arr.length <= headerRows) return;

    // 分离表头和数据
    var header = arr.slice(0, headerRows);
    var data = arr.slice(headerRows);

    // 排序
    data.sort(function(a, b) {
        for (var s = 0; s < sorts.length; s++) {
            var sort = sorts[s];
            var colIdx = sort.col - 1;
            var valA = a[colIdx];
            var valB = b[colIdx];

            // 自定义序列处理
            if (customOrder) {
                var orderArr = customOrder.split(',');
                var idxA = orderArr.indexOf(String(valA));
                var idxB = orderArr.indexOf(String(valB));
                if (idxA >= 0 && idxB >= 0) {
                    valA = idxA;
                    valB = idxB;
                }
            }

            if (valA < valB) return sort.order === 1 ? -1 : 1;
            if (valA > valB) return sort.order === 1 ? 1 : -1;
        }
        return 0;
    });

    // 写回
    r.Value2 = header.concat(data);
};
RngUtils.rngSortCols = RngUtils.z多列排序;

/**
 * z强力筛选 - 单元格强力筛选函数
 * @param {Range|string} rng - 待筛选的单元格范围
 * @param {...any} args - 多参数(列,条件回调,列,条件回调.....)
 * @example
 * RngUtils.z强力筛选($("A18:D24"),2,x=>x=='北京',4,x=>x>500)
 */
RngUtils.z强力筛选 = function(rng) {
    if (!isWPS) return;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // 解析参数
    var filters = [];
    for (var i = 1; i < arguments.length; i += 2) {
        filters.push({
            col: arguments[i],
            fn: typeof arguments[i + 1] === 'function' ? arguments[i + 1] : null
        });
    }

    // 获取数据
    var arr = RngUtils.z安全数组(r);
    var header = arr[0];
    var data = arr.slice(1);

    // 筛选
    var filtered = data.filter(function(row) {
        for (var f = 0; f < filters.length; f++) {
            var filter = filters[f];
            var val = row[filter.col - 1];
            if (filter.fn && !filter.fn(val)) {
                return false;
            }
        }
        return true;
    });

    // 隐藏不符合的行
    for (var i = 1; i <= r.Rows.Count; i++) {
        var shouldHide = true;
        for (var j = 0; j < filtered.length; j++) {
            if (JSON.stringify(r.Rows(i).Value2) === JSON.stringify([filtered[j]])) {
                shouldHide = false;
                break;
            }
        }
        if (shouldHide) {
            r.Rows(i).Hidden = true;
        }
    }
};
RngUtils.rngFilter = RngUtils.z强力筛选;

/**
 * z最大行数组 - 根据指定单元格区域获取最大行数组
 * @param {Range|string} rng - 单元格区域（如 "数据!A1:H1"）
 * @param {number} [cols] - 选择列作为获取最大行依据（可选，1-based索引）
 * @returns {Array2D} 结果二维数组（Array2D对象，支持链式调用）
 * @description 找出指定区域中所有列的最大行数，返回从第一行到最大行的完整二维数组
 * @example
 * RngUtils.z最大行数组("数据!A1:H1")        // 获取数据表A:H列的最大行数组
 * // 假设A-H列的最大行分别为5,6,3,5,343,444,32,2，则返回A1:H444的二维数组
 * RngUtils.z最大行数组("数据!A1:H1", 2)    // 使用第2列（B列）确定最大行
 * // 支持链式调用
 * $.maxArray("数据!A1:H1").z筛选(r=>r.f7>=4).superPivot(...).res()
 */
RngUtils.z最大行数组 = function(rng, cols) {
    if (!isWPS) return new Array2D([]);
    var r = typeof rng === 'string' ? Range(rng) : rng;
    if (!r) return new Array2D([]);
    
    var sheet = r.Worksheet;
    var startRow = r.Row;
    var startCol = r.Column;
    var maxRow = 0;

    if (cols !== undefined) {
        // 按指定列获取最大行（使用 z最大行 获取有效数据的最后一行）
        // 修正：当r是多单元格区域时，使用 r.Columns(cols) 获取指定列
        if (r.Columns.Count > 1) {
            var col = r.Columns(cols);
            maxRow = RngUtils.z最大行(col);
        } else {
            // 单列区域
            maxRow = RngUtils.z最大行(r);
        }
    } else {
        // 全部列中的最大行：遍历每一列，找出每列的最大行，取最大值
        for (var c = 0; c < r.Columns.Count; c++) {
            var colRange = sheet.Columns(startCol + c);
            var colMaxRow = RngUtils.z最大行(colRange);
            if (colMaxRow > maxRow) {
                maxRow = colMaxRow;
            }
        }
    }

    // 构建结果数组：从起始行到最大行
    var result = [];
    if (maxRow > 0) {
        for (var i = startRow; i <= maxRow; i++) {
            var rowData = [];
            for (var j = 0; j < r.Columns.Count; j++) {
                rowData.push(sheet.Cells(i, startCol + j).Value2);
            }
            result.push(rowData);
        }
    }

    return new Array2D(result);
};
RngUtils.maxArray = RngUtils.z最大行数组;

/**
 * z查找单元格 - 按指定条件查找单元格
 * @param {Range|string} rng - 单元格对象
 * @param {any|Object} args - 参数数组或单个值
 * @returns {Array} 单元格一维数组
 * @example
 * var rs=RngUtils.z查找单元格(Range("a11:d17"),'北京')
 * rs.unionAll().Address()
 */
RngUtils.z查找单元格 = function(rng, args) {
    if (!isWPS) return [];
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var results = [];

    // 默认按单个值查找
    var what = args;
    var after = r.Cells(r.Cells.Count);
    var lookIn = -4163; // xlValues
    var lookAt = 1; // xlWhole
    var searchOrder = 1; // xlByRows
    var searchDirection = 1; // xlNext
    var matchCase = false;
    var matchByte = false;
    var searchFormat = false;

    // 如果是对象则解析参数
    if (typeof args === 'object' && !Array.isArray(args)) {
        what = args.What !== undefined ? args.What : args;
        after = args.After || after;
        lookIn = args.LookIn !== undefined ? args.LookIn : lookIn;
        lookAt = args.LookAt !== undefined ? args.LookAt : lookAt;
        searchOrder = args.SearchOrder !== undefined ? args.SearchOrder : searchOrder;
        searchDirection = args.SearchDirection !== undefined ? args.SearchDirection : searchDirection;
        matchCase = args.MatchCase || false;
        matchByte = args.MatchByte || false;
        searchFormat = args.SearchFormat || false;
    }

    var found = r.Find(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat);
    var firstAddress = null;

    while (found) {
        results.push(found);
        if (!firstAddress) firstAddress = found.Address();
        if (found.Address() === firstAddress && results.length > 1) break;
        found = r.FindNext(found);
    }

    // 添加unionAll方法到结果数组
    results.unionAll = function() {
        return RngUtils.z联合区域(this);
    };

    return results;
};
RngUtils.findRange = RngUtils.z查找单元格;

/**
 * z命中单元格 - 检测指定单元格是否在指定单元格区域中
 * @param {Range|string} target - 待检测的单元格
 * @param {Range|string} checkRange - 检测单元格区域
 * @param {Function} [callback] - 命中回调函数
 * @returns {Boolean} 检测结果
 * @example
 * RngUtils.z命中单元格('c3','a1:d10')  // true
 */
RngUtils.z命中单元格 = function(target, checkRange, callback) {
    if (!isWPS) return false;
    var t = typeof target === 'string' ? Range(target) : target;
    var cr = typeof checkRange === 'string' ? Range(checkRange) : checkRange;

    var result = !Application.Intersect(t, cr) === null;

    if (result && callback) {
        callback(t);
    }

    return result;
};
RngUtils.hitRange = RngUtils.z命中单元格;

/**
 * z本文件单元格 - 打开多文件时，返回本文件当前表的指定单元格
 * @param {string} address - 单元格地址
 * @returns {Range} 本文件当前表的指定单元格
 * @example
 * RngUtils.z本文件单元格("a3")
 */
RngUtils.z本文件单元格 = function(address) {
    if (!isWPS) return null;
    return Application.ActiveSheet.Range(address);
};
RngUtils.thisRange = RngUtils.z本文件单元格;

/**
 * z选择不连续列数组 - 单元格不连续列装入数组
 * @param {Range|string} rng - 单元格对象
 * @param {string} cols - 不连续列号 "f1,f3-f5,f8"
 * @returns {Array} 二维数组
 * @example
 * RngUtils.selectColsArray(Cells,'f1,f3-f5')
 */
RngUtils.z选择不连续列数组 = function(rng, cols) {
    if (!isWPS) return [];
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // 解析列参数
    var colList = [];
    var parts = cols.split(',');
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        if (part.indexOf('-') > 0) {
            // 范围 f3-f5
            var rangeParts = part.replace('f', '').split('-');
            var start = parseInt(rangeParts[0]);
            var end = parseInt(rangeParts[1]);
            for (var c = start; c <= end; c++) {
                colList.push(c);
            }
        } else {
            // 单列 f1
            colList.push(parseInt(part.replace('f', '')));
        }
    }

    var arr = RngUtils.z安全数组(r);
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var row = [];
        for (var j = 0; j < colList.length; j++) {
            row.push(arr[i][colList[j] - 1]);
        }
        result.push(row);
    }
    return result;
};
RngUtils.selectColsArray = RngUtils.z选择不连续列数组;

// ==================== 使用辅助函数创建 RngUtils 实例方法别名 ====================
createBilingualAliases(RngUtils.prototype, [
    ['z加边框', 'addBorders'],
    ['z去边框', 'removeBorders'],
    ['z清除内容', 'clearContents'],
    ['z清除格式', 'clearFormats'],
    ['z自动列宽', 'autoFitColumns'],
    ['z自动行高', 'autoFitRows'],
    ['z设置背景色', 'backgroundColor'],
    ['z设置字体色', 'fontColor'],
    ['z行数', 'rowsCount'],
    ['z列数', 'colsCount'],
    ['z地址', 'address']
]);

/**
 * 加边框
 * @param {Number} lineStyle - 线条样式（默认1）
 * @param {Number} weight - 线条粗细（默认2）
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z加边框()
 * RngUtils("A1:C10").z加边框(1, 3).z设置背景色(0xFFFF00)
 */
RngUtils.prototype.z加边框 = function(lineStyle, weight) {
    if (!this._range) return this;
    lineStyle = lineStyle !== undefined ? lineStyle : 1;
    weight = weight !== undefined ? weight : 2;
    this._range.Borders.LineStyle = lineStyle;
    this._range.Borders.Weight = weight;
    return this;
};
RngUtils.prototype.addBorders = RngUtils.prototype.z加边框;

/**
 * 去边框
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z去边框()
 */
RngUtils.prototype.z去边框 = function() {
    if (!this._range) return this;
    this._range.Borders.LineStyle = -4142; // xlLineStyleNone
    return this;
};
RngUtils.prototype.removeBorders = RngUtils.prototype.z去边框;

/**
 * 清除内容
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z清除内容()
 */
RngUtils.prototype.z清除内容 = function() {
    if (!this._range) return this;
    this._range.ClearContents();
    return this;
};
RngUtils.prototype.clearContents = RngUtils.prototype.z清除内容;

/**
 * 清除格式
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z清除格式()
 */
RngUtils.prototype.z清除格式 = function() {
    if (!this._range) return this;
    this._range.ClearFormats();
    return this;
};
RngUtils.prototype.clearFormats = RngUtils.prototype.z清除格式;

// ==================== TreeNode - 多级树结构工具类 ====================

/**
 * TreeNode - 多级树结构工具类（支持多级联动菜单、省市县等层级数据）
 * @class
 * @description 实现树形数据结构，支持从二维数组快速构建多级联动菜单
 * @description 常用于：省市区三级联动、组织架构、分类层级等场景
 * @example
 * // 示例：从二维数组创建树（省市区数据）
 * var tree = TreeNode.initTree([
 *     ['四川省', '成都市', '武侯区'],
 *     ['四川省', '成都市', '锦江区'],
 *     ['广东省', '深圳市', '南山区']
 * ]);
 * 
 * // 获取所有省级节点
 * var provinces = tree.getChildren();
 * 
 * // 获取四川省下的所有市级节点
 * var cities = tree.getChild('四川省').getChildren();
 * 
 * // 获取完整路径
 * var path = tree.getPath('南山区'); // ['广东省', '深圳市', '南山区']
 */
function TreeNode(value) {
    if (!(this instanceof TreeNode)) {
        return new TreeNode(value);
    }
    this.value = value;          // 当前节点的值
    this.children = {};          // 子节点集合 {key: TreeNode}
    this.parent = null;          // 父节点引用
    this.level = 0;              // 节点层级（从0开始）
    this.data = {};              // 附加数据存储
}

/**
 * 初始化树结构（从二维数组）
 * @static
 * @param {Array} arr - 二维数组，每行是一个完整路径（如['省','市','区']）
 * @returns {TreeNode} 根节点
 * @example
 * var tree = TreeNode.initTree([
 *     ['四川省', '成都市', '武侯区'],
 *     ['四川省', '绵阳市', '涪城区']
 * ]);
 */
TreeNode.initTree = function(arr) {
    var root = new TreeNode('root');
    root.level = -1;
    
    for (var i = 0; i < arr.length; i++) {
        var path = arr[i];
        var currentNode = root;
        
        for (var j = 0; j < path.length; j++) {
            var key = path[j];
            if (!key) continue; // 跳过空值
            
            if (!currentNode.children[key]) {
                var newNode = new TreeNode(key);
                newNode.parent = currentNode;
                newNode.level = currentNode.level + 1;
                currentNode.children[key] = newNode;
            }
            currentNode = currentNode.children[key];
        }
    }
    
    return root;
};
TreeNode.z从数组 = TreeNode.initTree;

/**
 * 添加子节点
 * @param {string} key - 子节点键
 * @param {any} [value] - 子节点值（不传则使用key）
 * @returns {TreeNode} 新创建的子节点
 * @example
 * var node = new TreeNode('四川省');
 * node.addChild('成都市');
 */
TreeNode.prototype.addChild = function(key, value) {
    if (!this.children[key]) {
        var newNode = new TreeNode(value !== undefined ? value : key);
        newNode.parent = this;
        newNode.level = this.level + 1;
        this.children[key] = newNode;
    }
    return this.children[key];
};
TreeNode.z添加子节点 = TreeNode.prototype.addChild;

/**
 * 获取子节点
 * @param {string} key - 子节点键
 * @returns {TreeNode|null} 子节点或null
 * @example
 * var city = tree.getChild('四川省').getChild('成都市');
 */
TreeNode.prototype.getChild = function(key) {
    return this.children[key] || null;
};
TreeNode.z获取子节点 = TreeNode.prototype.getChild;

/**
 * 获取所有子节点
 * @returns {Array} 子节点数组
 * @example
 * var cities = tree.getChildren();
 */
TreeNode.prototype.getChildren = function() {
    var result = [];
    for (var key in this.children) {
        result.push(this.children[key]);
    }
    return result;
};
TreeNode.z获取所有子节点 = TreeNode.prototype.getChildren;

/**
 * 获取所有子节点键名
 * @returns {Array} 键名数组
 * @example
 * var cityNames = tree.getChildKeys(); // ['成都市', '绵阳市']
 */
TreeNode.prototype.getChildKeys = function() {
    return Object.keys(this.children);
};
TreeNode.z获取子键数组 = TreeNode.prototype.getChildKeys;

/**
 * 判断是否有子节点
 * @returns {boolean} 是否有子节点
 * @example
 * if (node.hasChildren()) { ... }
 */
TreeNode.prototype.hasChildren = function() {
    return Object.keys(this.children).length > 0;
};
TreeNode.z有子节点 = TreeNode.prototype.hasChildren;

/**
 * 获取父节点
 * @returns {TreeNode|null} 父节点或null
 * @example
 * var parent = node.getParent();
 */
TreeNode.prototype.getParent = function() {
    return this.parent;
};
TreeNode.z获取父节点 = TreeNode.prototype.getParent;

/**
 * 获取从根到当前节点的完整路径
 * @returns {Array} 路径数组
 * @example
 * var path = node.getPath(); // ['四川省', '成都市', '武侯区']
 */
TreeNode.prototype.getPath = function() {
    var path = [];
    var current = this;
    while (current && current.parent) {
        path.unshift(current.value);
        current = current.parent;
    }
    return path;
};
TreeNode.z获取路径 = TreeNode.prototype.getPath;

/**
 * 根据路径查找节点
 * @param {Array|string} path - 路径数组或点分隔字符串
 * @returns {TreeNode|null} 找到的节点或null
 * @example
 * var node = tree.findByPath(['四川省', '成都市', '武侯区']);
 * var node = tree.findByPath('四川省.成都市.武侯区');
 */
TreeNode.prototype.findByPath = function(path) {
    if (typeof path === 'string') {
        path = path.split('.');
    }
    
    var current = this;
    for (var i = 0; i < path.length; i++) {
        if (current.children[path[i]]) {
            current = current.children[path[i]];
        } else {
            return null; // 路径不存在
        }
    }
    return current;
};
TreeNode.z查找路径 = TreeNode.prototype.findByPath;

/**
 * 深度优先遍历
 * @param {Function} callback - 回调函数(node, level)
 * @returns {TreeNode} 当前实例
 * @example
 * tree.depthFirst((node, level) => console.log(node.value, level));
 */
TreeNode.prototype.depthFirst = function(callback) {
    callback(this, this.level);
    for (var key in this.children) {
        this.children[key].depthFirst(callback);
    }
    return this;
};
TreeNode.z深度遍历 = TreeNode.prototype.depthFirst;

/**
 * 广度优先遍历
 * @param {Function} callback - 回调函数(node, level)
 * @returns {TreeNode} 当前实例
 * @example
 * tree.breadthFirst((node, level) => console.log(node.value, level));
 */
TreeNode.prototype.breadthFirst = function(callback) {
    var queue = [this];
    while (queue.length > 0) {
        var node = queue.shift();
        callback(node, node.level);
        for (var key in node.children) {
            queue.push(node.children[key]);
        }
    }
    return this;
};
TreeNode.z广度遍历 = TreeNode.prototype.breadthFirst;

/**
 * 获取最大深度
 * @returns {number} 最大层级数
 * @example
 * var depth = tree.getMaxDepth(); // 3（省、市、区）
 */
TreeNode.prototype.getMaxDepth = function() {
    var maxDepth = this.level;
    for (var key in this.children) {
        var childDepth = this.children[key].getMaxDepth();
        if (childDepth > maxDepth) {
            maxDepth = childDepth;
        }
    }
    return maxDepth;
};
TreeNode.z最大深度 = TreeNode.prototype.getMaxDepth;

/**
 * 获取某一层级的所有节点
 * @param {number} level - 目标层级
 * @returns {Array} 节点数组
 * @example
 * var level2Nodes = tree.getNodesByLevel(2); // 所有区级节点
 */
TreeNode.prototype.getNodesByLevel = function(level) {
    var result = [];
    this.breadthFirst(function(node) {
        if (node.level === level) {
            result.push(node);
        }
    });
    return result;
};
TreeNode.z获取层级节点 = TreeNode.prototype.getNodesByLevel;

/**
 * 转换为JSON（便于序列化）
 * @returns {Object} JSON对象
 * @example
 * var json = tree.toJSON();
 */
TreeNode.prototype.toJSON = function() {
    var obj = {
        value: this.value,
        level: this.level,
        data: this.data,
        children: {}
    };
    
    for (var key in this.children) {
        obj.children[key] = this.children[key].toJSON();
    }
    
    return obj;
};
TreeNode.z转JSON = TreeNode.prototype.toJSON;

/**
 * 静态方法：快速创建树（从平面数据）
 * @static
 * @param {Array} flatData - 平面数组，每个元素包含parentKey
 * @param {string} keyField - 键字段
 * @param {string} parentField - 父键字段
 * @returns {TreeNode} 根节点
 * @example
 * var tree = TreeNode.fromFlatData([{id:1,pid:0,name:'a'},{id:2,pid:1,name:'b'}], 'id', 'pid');
 */
TreeNode.fromFlatData = function(flatData, keyField, parentField, rootParentValue) {
    rootParentValue = rootParentValue !== undefined ? rootParentValue : 0;
    
    var root = new TreeNode('root');
    var nodeMap = {};
    
    // 第一轮：创建所有节点
    for (var i = 0; i < flatData.length; i++) {
        var item = flatData[i];
        var key = item[keyField];
        nodeMap[key] = new TreeNode(item);
    }
    
    // 第二轮：建立父子关系
    for (var i = 0; i < flatData.length; i++) {
        var item = flatData[i];
        var key = item[keyField];
        var parentKey = item[parentField];
        
        if (parentKey == rootParentValue || !nodeMap[parentKey]) {
            root.addChild(key, nodeMap[key]);
        } else {
            nodeMap[parentKey].addChild(key, nodeMap[key]);
        }
    }
    
    return root;
};
TreeNode.z从平面数据 = TreeNode.fromFlatData;

// ==================== 使用辅助函数创建 TreeNode 方法别名 ====================
createBilingualAliases(TreeNode.prototype, [
    ['z从数组', 'initTree'],
    ['z添加子节点', 'addChild'],
    ['z获取子节点', 'getChild'],
    ['z获取所有子节点', 'getChildren'],
    ['z获取子键数组', 'getChildKeys'],
    ['z有子节点', 'hasChildren'],
    ['z获取父节点', 'getParent'],
    ['z获取路径', 'getPath'],
    ['z查找路径', 'findByPath'],
    ['z深度遍历', 'depthFirst'],
    ['z广度遍历', 'breadthFirst'],
    ['z最大深度', 'getMaxDepth'],
    ['z获取层级节点', 'getNodesByLevel'],
    ['z转JSON', 'toJSON'],
    ['z从平面数据', 'fromFlatData']
]);

/**
 * 自动列宽
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z自动列宽()
 * RngUtils("A:Z").z自动列宽()  // 整列自动调整
 */
RngUtils.prototype.z自动列宽 = function() {
    if (!this._range) return this;
    this._range.Columns.AutoFit();
    return this;
};
RngUtils.prototype.autoFitColumns = RngUtils.prototype.z自动列宽;

/**
 * 自动行高
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z自动行高()
 */
RngUtils.prototype.z自动行高 = function() {
    if (!this._range) return this;
    this._range.Rows.AutoFit();
    return this;
};
RngUtils.prototype.autoFitRows = RngUtils.prototype.z自动行高;

/**
 * 设置背景色
 * @param {Number} color - RGB颜色值
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z设置背景色(RGB(255, 0, 0))  // 红色背景
 * RngUtils("A1:C10").z设置背景色(0xFFFF00)        // 黄色背景
 */
RngUtils.prototype.z设置背景色 = function(color) {
    if (!this._range) return this;
    this._range.Interior.Color = color;
    return this;
};
RngUtils.prototype.backgroundColor = RngUtils.prototype.z设置背景色;

/**
 * 设置字体色
 * @param {Number} color - RGB颜色值
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z设置字体色(RGB(255, 0, 0))  // 红色字体
 * RngUtils("A1:C10").z设置字体色(0xFF0000)        // 红色字体
 */
RngUtils.prototype.z设置字体色 = function(color) {
    if (!this._range) return this;
    this._range.Font.Color = color;
    return this;
};
RngUtils.prototype.fontColor = RngUtils.prototype.z设置字体色;

/**
 * 获取行数
 * @returns {Number} 行数
 */
RngUtils.prototype.z行数 = function() {
    if (!this._range) return 0;
    return this._range.Rows.Count;
};
RngUtils.prototype.rowsCount = RngUtils.prototype.z行数;

/**
 * 获取列数
 * @returns {Number} 列数
 */
RngUtils.prototype.z列数 = function() {
    if (!this._range) return 0;
    return this._range.Columns.Count;
};
RngUtils.prototype.colsCount = RngUtils.prototype.z列数;

/**
 * 获取地址
 * @returns {String} 单元格地址
 */
RngUtils.prototype.z地址 = function() {
    if (!this._range) return '';
    return this._range.Address();
};
RngUtils.prototype.address = RngUtils.prototype.z地址;

// ==================== SuperMap - 可在局部变量窗口实时展开查看的增强版 Map ====================

/**
 * SuperMap - 可在局部变量窗口实时展开查看的增强版 Map
 *
 * 特点：
 * 1. 完全兼容原生 Map 的所有属性和方法
 * 2. all 属性自动初始化，创建后立即可在局部变量窗口查看
 * 3. 支持嵌套 SuperMap、二维数组、Map 数组
 * 4. 层级前缀标识（01L00001 = 层数+序号+key）
 * 5. 调试模式开关，关闭后性能接近原生 Map
 */
function SuperMap(entries) {
    if (!(this instanceof SuperMap)) {
        return new SuperMap(entries);
    }
    this._map = new Map(entries);
    this._debug = true;
    this._all = null;  // 存储 all 属性值
    this._updateAll();  // 构造时立即初始化
}

// ========== 调试模式控制 ==========

Object.defineProperty(SuperMap.prototype, 'debug', {
    get: function() {
        return this._debug;
    },
    set: function(value) {
        this._debug = !!value;
        this._updateAll();
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(SuperMap, 'debug', {
    get: function() {
        return SuperMap._staticDebug;
    },
    set: function(value) {
        SuperMap._staticDebug = !!value;
    },
    enumerable: true,
    configurable: true
});
SuperMap._staticDebug = true;

// ========== 定义 all 属性 ==========

Object.defineProperty(SuperMap.prototype, 'all', {
    get: function() {
        return this._all;
    },
    enumerable: true,
    configurable: true
});

// ========== 定义 size 属性 ==========

Object.defineProperty(SuperMap.prototype, 'size', {
    get: function() {
        return this._map.size;
    },
    enumerable: true,
    configurable: true
});

// ========== 原型方法 ==========

/**
 * 更新 all 属性（构造时和每次修改后自动调用）
 */
SuperMap.prototype._updateAll = function() {
    if (!this._debug && !SuperMap._staticDebug) {
        this._all = { _提示: "调试模式已关闭，设置 debug=true 查看" };
        return;
    }
    this._all = this._buildAllView(1);
};

/**
 * 构建树形视图
 * 格式：01L00001 key（层数+序号+原key）
 */
SuperMap.prototype._buildAllView = function(level, maxRows) {
    level = level || 1;
    maxRows = maxRows || 255;

    var result = {};
    var count = 0;

    for (var entry of this._map) {
        var key = entry[0];
        var value = entry[1];

        if (count >= maxRows) {
            result['_省略剩余' + (this._map.size - count) + '项'] = "...";
            break;
        }

        // 格式：01L00001 key（层数+序号+原key）
        var prefix = (level < 10 ? '0' : '') + level + 'L';
        var seqNum = '0000' + (count + 1);
        var displayKey = prefix + seqNum.slice(-5) + ' ' + key;

        // 判断值类型并处理
        if (value instanceof SuperMap) {
            // 嵌套 SuperMap：递归展开
            result[displayKey] = value._buildAllView(level + 1, maxRows);
        } else if (value instanceof Map) {
            // 普通 Map：转为 SuperMap 后展开
            var superMap = SuperMap.fromMap(value, false);
            superMap._debug = this._debug;
            result[displayKey] = superMap._buildAllView(level + 1, maxRows);
        } else if (this._is2DArray(value)) {
            // 二维数组：按序号展开
            var arrObj = {};
            for (var i = 0; i < value.length && i < maxRows; i++) {
                arrObj[i + 1] = value[i];
            }
            result[displayKey] = arrObj;
        } else if (Array.isArray(value) && value.length > 0 && value[0] instanceof Map) {
            // Map 数组：转为 SuperMap 数组
            var arrObj = {};
            for (var i = 0; i < value.length && i < maxRows; i++) {
                var sm = SuperMap.fromMap(value[i], false);
                sm._debug = this._debug;
                arrObj[i + 1] = sm._buildAllView(level + 1, maxRows);
            }
            result[displayKey] = arrObj;
        } else if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
            // 普通对象：直接展示
            result[displayKey] = value;
        } else {
            // 基础数据类型
            result[displayKey] = value;
        }

        count++;
    }

    return result;
};

/**
 * 判断是否为二维数组
 */
SuperMap.prototype._is2DArray = function(value) {
    if (!Array.isArray(value)) return false;
    if (value.length === 0) return false;
    return Array.isArray(value[0]);
};

// ========== Map 原生方法（自动更新 all）==========

SuperMap.prototype.set = function(key, value) {
    var result = this._map.set(key, value);
    this._updateAll();  // 自动更新
    return result;
};

SuperMap.prototype.get = function(key) {
    return this._map.get(key);
};

SuperMap.prototype.has = function(key) {
    return this._map.has(key);
};

SuperMap.prototype.delete = function(key) {
    var result = this._map.delete(key);
    this._updateAll();  // 自动更新
    return result;
};

SuperMap.prototype.clear = function() {
    var result = this._map.clear();
    this._updateAll();  // 自动更新
    return result;
};

SuperMap.prototype.forEach = function(callback, thisArg) {
    return this._map.forEach(callback, thisArg);
};

SuperMap.prototype.keys = function() {
    return this._map.keys();
};

SuperMap.prototype.values = function() {
    return this._map.values();
};

SuperMap.prototype.entries = function() {
    return this._map.entries();
};

// ========== 转换方法 ==========

/**
 * 转为普通 Map 对象
 */
SuperMap.prototype.toMap = function(deep) {
    deep = deep !== undefined ? deep : true;

    var result = new Map();
    for (var entry of this._map) {
        var key = entry[0];
        var value = entry[1];

        if (deep && value instanceof SuperMap) {
            result.set(key, value.toMap(deep));
        } else if (deep && value instanceof Map) {
            result.set(key, new Map(value));
        } else if (deep && Array.isArray(value)) {
            result.set(key, value.map(function(item) {
                return item instanceof SuperMap ? item.toMap(deep) : item;
            }));
        } else {
            result.set(key, value);
        }
    }
    return result;
};

/**
 * 静态方法：将普通 Map 转为 SuperMap
 */
SuperMap.fromMap = function(map, deep) {
    if (!(map instanceof Map)) {
        throw new Error("参数必须是 Map 类型");
    }

    deep = deep !== undefined ? deep : true;
    var entries = [];

    for (var entry of map) {
        var key = entry[0];
        var value = entry[1];

        if (deep && value instanceof Map) {
            entries.push([key, SuperMap.fromMap(value, deep)]);
        } else if (deep && Array.isArray(value)) {
            entries.push([key, value.map(function(item) {
                return item instanceof Map ? SuperMap.fromMap(item, deep) : item;
            })]);
        } else {
            entries.push([key, value]);
        }
    }

    return new SuperMap(entries);
};
SuperMap.z从Map = SuperMap.fromMap;

/**
 * 将SuperMap内容写入单元格
 * @param {String|Range} rng - 目标单元格地址或Range对象
 * @returns {SuperMap} 当前实例
 * @example
 * SuperMap.fromMap(map).toRange('A1');
 */
SuperMap.prototype.toRange = function(rng) {
    if (!isWPS) return this;
    if (this._map.size === 0) return this;

    var arr = [['键', '聚合结果', '原始数据']];
    this._map.forEach(function(value, key) {
        var aggText = Array.isArray(value.agg) ? value.agg.join(', ') : JSON.stringify(value.agg);
        arr.push([key, aggText, JSON.stringify(value.group || {})]);
    });

    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = arr.length;
    var cols = arr[0].length;
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = arr;
    return this;
};
SuperMap.prototype.z写入单元格 = SuperMap.prototype.toRange;

/**
 * 打印 all 内容到控制台
 */
SuperMap.prototype.print = function(title) {
    title = title || "SuperMap 内容";
    Console.log("===== " + title + " =====");
    Console.log(JSON.stringify(this.all, null, 2));
    Console.log("========================");
};
SuperMap.prototype.z打印 = SuperMap.prototype.print;

// ==================== ShtUtils - 工作表工具库 ====================

/**
 * ShtUtils - 工作表操作工具（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description WPS工作表操作增强工具
 * @example
 * ShtUtils.z安全已使用区域("Sheet1")
 */
function ShtUtils(initialSheet) {
    if (!(this instanceof ShtUtils)) {
        return new ShtUtils(initialSheet);
    }
    this._sheet = initialSheet ? this._getSheet(initialSheet) : null;
}

/**
 * 获取工作表对象
 * @private
 * @param {String|Worksheet} sht - 工作表名称或对象
 * @returns {Worksheet|null} 工作表对象或null
 */
ShtUtils.prototype._getSheet = function(sht) {
    if (typeof sht === 'string') {
        return isWPS ? Sheets(sht) : null;
    }
    return sht;
};

/**
 * 安全已使用区域
 * @param {String|Worksheet} 工作表 - 工作表名称或对象
 * @returns {Range} 已使用区域
 */
ShtUtils.prototype.z安全已使用区域 = function(工作表) {
    var sheet = 工作表 ? this._getSheet(工作表) : (this._sheet || (isWPS ? Application.ActiveSheet : null));
    if (!sheet) return null;

    var usedRange;
    try {
        usedRange = sheet.UsedRange;
    } catch (e) {
        return sheet.Range("A1");
    }

    if (!usedRange) return sheet.Range("A1");

    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;

    return sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
};
ShtUtils.prototype.safeUsedRange = ShtUtils.prototype.z安全已使用区域;

/**
 * 包含表名（支持通配符）
 * @param {String} 表名 - 表名模式
 * @param {Object} 表集合 - 表集合（可选）
 * @returns {Boolean} 是否包含
 */
ShtUtils.prototype.z包含表名 = function(表名, 表集合) {
    var shts = 表集合 || (isWPS ? Sheets : null);
    if (!shts) return false;

    var pattern = this._wildcardToRegex(表名);
    for (var i = 1; i <= shts.Count; i++) {
        if (pattern.test(shts(i).Name)) return true;
    }
    return false;
};
ShtUtils.prototype.includesSht = ShtUtils.prototype.z包含表名;

/**
 * 通配符转正则
 * @private
 * @param {String} wildcard - 通配符模式（支持 * 和 ?）
 * @returns {RegExp} 正则表达式对象
 */
ShtUtils.prototype._wildcardToRegex = function(wildcard) {
    var pattern = wildcard.replace(/[.+^${}()|[\]\\]/g, '\\$&')
        .replace(/\*/g, '.*')
        .replace(/\?/g, '.');
    return new RegExp('^' + pattern + '$', 'i');
};

/**
 * 表名筛选
 * @param {String} 表名 - 表名模式
 * @param {Sheets} 表集合 - 表集合（可选）
 * @returns {Array} 匹配的表名数组
 */
ShtUtils.prototype.z表名筛选 = function(表名, 表集合) {
    var shts = 表集合 || (isWPS ? Sheets : null);
    if (!shts) return [];

    var pattern = this._wildcardToRegex(表名);
    var result = [];
    for (var i = 1; i <= shts.Count; i++) {
        if (pattern.test(shts(i).Name)) {
            result.push(shts(i).Name);
        }
    }
    return result;
};
ShtUtils.prototype.filterShts = ShtUtils.prototype.z表名筛选;

/**
 * 激活工作表
 * @param {String|Worksheet} 工作表 - 工作表（可选）
 * @returns {ShtUtils} 当前实例
 */
ShtUtils.prototype.z激活表 = function(工作表) {
    var sheet = 工作表 ? this._getSheet(工作表) : this._sheet;
    if (sheet) sheet.Activate();
    this._sheet = sheet;
    return this;
};
ShtUtils.prototype.shtActivate = ShtUtils.prototype.z激活表;

// ==================== DateUtils - 日期工具库 ====================

/**
 * DateUtils - 日期操作工具（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 日期时间处理工具
 * @example
 * DateUtils.dt().z加天(5).z月底().val()
 */
function DateUtils(initialDate) {
    if (!(this instanceof DateUtils)) {
        return new DateUtils(initialDate);
    }
    this._date = initialDate ? new Date(initialDate) : new Date();
}

/**
 * 获取/设置日期
 * @param {Date|number|string} newDate - 新日期
 * @returns {DateUtils|Date} 设置时返回this，否则返回当前日期
 */
DateUtils.prototype.dt = function(newDate) {
    if (newDate !== undefined) {
        this._date = new Date(newDate);
        return this;
    }
    return this._date;
};

/**
 * 获取值
 * @returns {Date} 当前日期对象
 */
DateUtils.prototype.val = function() {
    return this._date;
};

/**
 * 加天数
 * @param {Number} days - 天数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加天 = function(days) {
    var result = new Date(this._date);
    result.setDate(result.getDate() + days);
    this._date = result;
    return this;
};
DateUtils.prototype.addDays = DateUtils.prototype.z加天;

/**
 * 加月数
 * @param {Number} months - 月数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加月 = function(months) {
    var result = new Date(this._date);
    result.setMonth(result.getMonth() + months);
    this._date = result;
    return this;
};
DateUtils.prototype.addMonths = DateUtils.prototype.z加月;

/**
 * 加年数
 * @param {Number} years - 年数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加年 = function(years) {
    var result = new Date(this._date);
    result.setFullYear(result.getFullYear() + years);
    this._date = result;
    return this;
};
DateUtils.prototype.addYears = DateUtils.prototype.z加年;

/**
 * 月初
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z月初 = function() {
    this._date = new Date(this._date.getFullYear(), this._date.getMonth(), 1);
    return this;
};
DateUtils.prototype.firstDayOfMonth = DateUtils.prototype.z月初;

/**
 * 月底
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z月底 = function() {
    this._date = new Date(this._date.getFullYear(), this._date.getMonth() + 1, 0);
    return this;
};
DateUtils.prototype.endOfMonth = DateUtils.prototype.z月底;

/**
 * 转表格日期
 * @param {Date} jsdate - JS日期
 * @returns {Number} Excel日期数值
 */
DateUtils.prototype.z转表格日期 = function(jsdate) {
    if (!(jsdate instanceof Date)) {
        jsdate = new Date(jsdate);
    }
    var excelBase = new Date(1900, 0, 1).getTime();
    var dateMs = jsdate.getTime();
    var dayInMs = 24 * 60 * 60 * 1000;
    return (dateMs - excelBase) / dayInMs + 2;
};
DateUtils.prototype.toExcelDate = DateUtils.prototype.z转表格日期;

/**
 * 日期格式化
 * @param {Date} jsdate - 日期
 * @param {String} fmt - 格式
 * @returns {String} 格式化字符串
 */
DateUtils.prototype.z日期格式化 = function(jsdate, fmt) {
    if (!(jsdate instanceof Date)) {
        jsdate = new Date(jsdate);
    }
    var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
    return fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
        if (year) return jsdate.getFullYear().toString().padStart(year.length, '0');
        if (month) return (jsdate.getMonth() + 1).toString().padStart(month.length, '0');
        if (day) return jsdate.getDate().toString().padStart(day.length, '0');
        if (hour) return jsdate.getHours().toString().padStart(hour.length, '0');
        if (minute) return jsdate.getMinutes().toString().padStart(minute.length, '0');
        if (second) return jsdate.getSeconds().toString().padStart(second.length, '0');
        if (millisecond) return jsdate.getMilliseconds().toString().padStart(3, '0');
        if (week) return '周' + weekDays[jsdate.getDay()];
        return match;
    });
};
DateUtils.prototype.format = DateUtils.prototype.z日期格式化;

/**
 * 获取年份
 * @returns {Number} 年份（4位数字）
 * @example
 * asDate("2023-9-21").z年份()  // 2023
 */
DateUtils.prototype.z年份 = function() {
    return this._date.getFullYear();
};
DateUtils.prototype.getYear = DateUtils.prototype.z年份;

/**
 * 获取月份（1-12）
 * @returns {Number} 月份（1-12）
 * @example
 * asDate("2023-9-21").z月份()  // 9
 */
DateUtils.prototype.z月份 = function() {
    return this._date.getMonth() + 1;
};
DateUtils.prototype.getMonth = DateUtils.prototype.z月份;

/**
 * 获取日期（1-31）
 * @returns {Number} 日期（1-31）
 * @example
 * asDate("2023-9-21").z日期()  // 21
 */
DateUtils.prototype.z日期 = function() {
    return this._date.getDate();
};
DateUtils.prototype.getDate = DateUtils.prototype.z日期;

/**
 * 获取星期（1-7，7=周日）
 * @returns {Number} 星期（1-7）
 * @example
 * asDate("2023-9-21").z星期()  // 4 (周四)
 * asDate("2023-9-24").z星期()  // 7 (周日)
 */
DateUtils.prototype.z星期 = function() {
    var day = this._date.getDay();
    return day === 0 ? 7 : day;
};
DateUtils.prototype.getDay = DateUtils.prototype.z星期;

/**
 * 获取小时（0-23）
 * @returns {Number} 小时（0-23）
 */
DateUtils.prototype.z小时 = function() {
    return this._date.getHours();
};
DateUtils.prototype.getHour = DateUtils.prototype.z小时;

/**
 * 获取分钟（0-59）
 * @returns {Number} 分钟（0-59）
 */
DateUtils.prototype.z分钟 = function() {
    return this._date.getMinutes();
};
DateUtils.prototype.getMinute = DateUtils.prototype.z分钟;

/**
 * 获取秒数（0-59）
 * @returns {Number} 秒数（0-59）
 */
DateUtils.prototype.z秒 = function() {
    return this._date.getSeconds();
};
DateUtils.prototype.getSecond = DateUtils.prototype.z秒;

/**
 * 获取时间戳（毫秒）
 * @returns {Number} 时间戳
 */
DateUtils.prototype.z时间戳 = function() {
    return this._date.getTime();
};
DateUtils.prototype.getTime = DateUtils.prototype.z时间戳;

// ==================== JSA - 通用函数库 ====================

/**
 * JSA - 通用函数工具（静态方法）
 * @class
 * @description 常用函数集合
 */
function JSA() {}

/**
 * 转置数组
 * @param {Array} arr - 数组
 * @returns {Array} 转置后的数组
 */
JSA.z转置 = function(arr) {
    if (!arr || arr.length === 0) return [];
    var rows = arr.length;
    var cols = arr[0].length;
    var result = [];
    for (var j = 0; j < cols; j++) {
        result[j] = [];
        for (var i = 0; i < rows; i++) {
            result[j][i] = arr[i][j];
        }
    }
    return result;
};
JSA.transpose = JSA.z转置;

/**
 * 转数值
 * @param {String} text - 文本
 * @returns {Number} 数值
 */
JSA.z转数值 = function(text) {
    if (typeof text === 'number') return text;
    if (typeof text === 'string') {
        text = text.trim();
        var match = text.match(/^[-+]?[0-9]*\.?[0-9]+/);
        if (match) return parseFloat(match[0]);
        return 0;
    }
    return 0;
};
JSA.val = JSA.z转数值;

/**
 * 写入单元格（根据数组大小自动扩展区域）
 * @param {Array} arr - 数组
 * @param {Range|string} rng - 单元格区域（左上角单元格）
 * @param {Boolean} clearDown - 是否清空下方（保留参数兼容性）
 * @returns {Range} 写入的Range
 */
JSA.z写入单元格 = function(arr, rng, clearDown) {
    if (!isWPS) return null;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = arr.length;
    var cols = rows > 0 ? (Array.isArray(arr[0]) ? arr[0].length : 1) : 0;
    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = arr;
    return writeRng;
};
JSA.toRange = JSA.z写入单元格;

/**
 * 获取今天日期
 * @returns {String} 今天日期 YYYY-MM-DD
 */
JSA.z今天 = function() {
    var now = new Date();
    return now.getFullYear() + '-' +
           String(now.getMonth() + 1).padStart(2, '0') + '-' +
           String(now.getDate()).padStart(2, '0');
};
JSA.today = JSA.z今天;

/**
 * 转日期数值
 * @param {Date|string} d - 日期
 * @returns {Number} Excel日期数值
 */
JSA.z转日期数值 = function(d) {
    var date = typeof d === 'string' ? new Date(d) : d;
    var excelEpoch = new Date(1900, 0, 1);
    var msPerDay = 24 * 60 * 60 * 1000;
    return Math.floor((date - excelEpoch) / msPerDay) + 2;
};
JSA.cdate = JSA.z转日期数值;

/**
 * 替换
 * @param {String} str - 字符串
 * @param {String} find - 查找
 * @param {String} replaceWith - 替换
 * @returns {String} 结果
 */
JSA.z替换 = function(str, find, replaceWith) {
    return str.split(find).join(replaceWith);
};
JSA.replace = JSA.z替换;

/**
 * 全局替换（字符串扩展方法）
 * @param {String} search - 查找字符串
 * @param {String} replacement - 替换字符串
 * @returns {String} 结果
 * @description 在字符串上执行全局替换，等同于 replace(/search/g, replacement)
 * @example "jsa字符串a".substitute('a','b') // 返回 "jsb字符串b"
 */
String.prototype.substitute = function(search, replacement) {
    if (search === undefined || search === null || search === '') return String(this);
    if (replacement === undefined) replacement = '';
    // 使用 split + join 实现全局替换，避免正则特殊字符问题
    return String(this).split(String(search)).join(String(replacement));
};

/**
 * 全局替换（字符串扩展方法 - 中文别名）
 */
String.prototype.z全局替换 = String.prototype.substitute;

/**
 * 模糊匹配（字符串扩展方法）
 * @param {String} pattern - 匹配模式（支持 * 通配符任意字符，? 通配符单个字符）
 * @returns {Number} 匹配返回-1，不匹配返回0
 * @description 包含模式匹配，自动在模式前后添加 *，除非模式以 ^ 开头或 $ 结尾
 * @example "jsa字符串b".like('a*b')        // 返回 -1 (包含a和b)
 * @example "jsa字符串c".like('a*b')        // 返回 0  (有a但没有b)
 * @example "hello".like('h?llo')          // 返回 -1 (?匹配单字符)
 * @example "abc".like('a*c')              // 返回 -1
 * @example "abc".like('a?c')              // 返回 -1
 * @example "abc".like('^a')               // 返回 -1 (以a开头)
 * @example "abc".like('c$')               // 返回 -1 (以c结尾)
 */
String.prototype.like = function(pattern) {
    if (pattern === undefined || pattern === null) return 0;
    var str = String(this);
    // 转义正则特殊字符，但保留 * 和 ?
    var regexPattern = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&')
                              .replace(/\*/g, '.*')
                              .replace(/\?/g, '.');
    // 包含模式：自动在前后加上 .*，除非模式已经以 ^ 开头或 $ 结尾
    var anchored = regexPattern;
    if (anchored.indexOf('^') !== 0) anchored = '.*' + anchored;
    if (anchored.charAt(anchored.length - 1) !== '$') anchored = anchored + '.*';
    var regex = new RegExp('^' + anchored + '$');
    // 匹配返回-1，不匹配返回0
    return regex.test(str) ? -1 : 0;
};

/**
 * 模糊匹配（字符串扩展方法 - 中文别名）
 */
String.prototype.z模糊匹配 = String.prototype.like;

/**
 * 数组转JSON字符串（数组扩展方法）
 * @returns {String} JSON字符串
 * @description 将数组转换为JSON格式字符串
 * @example [1,2,3].toJson()              // 返回 "[1,2,3]"
 * @example ["a","b"].toJson()            // 返回 "[\"a\",\"b\"]"
 * @example [{x:1},{y:2}].toJson()        // 返回 "[{\"x\":1},{\"y\":2}]"
 */
Array.prototype.toJson = function() {
    return JSON.stringify(this);
};

/**
 * 数组转JSON字符串（数组扩展方法 - 中文别名）
 */
Array.prototype.z转JSON = Array.prototype.toJson;

/**
 * 数组元素转数值（数组扩展方法）
 * @returns {Array} 数值数组
 * @description 将数组中每个元素转换为数值
 * @example "1a2b3c4asd5".match(/\d/g).val()        // 返回 [1,2,3,4,5]
 * @example ["1","2","3"].val()                    // 返回 [1,2,3]
 * @example ["10","20","abc"].val()                // 返回 [10,20,0]
 */
Array.prototype.val = function() {
    return this.map(function(item) {
        var num = Number(item);
        return isNaN(num) ? 0 : num;
    });
};

/**
 * 数组元素转数值（数组扩展方法 - 中文别名）
 */
Array.prototype.z转数值 = Array.prototype.val;

/**
 * 数组写入单元格（数组扩展方法，根据数组大小自动扩展区域）
 * @param {Range|string} rng - 单元格区域（左上角单元格）
 * @returns {Range} 写入的Range
 * @description 将二维数组写入指定单元格
 * @example
 * var arr = [[1, 'A'], [2, 'B'], [3, 'C']];
 * arr.toRange("J2");                    // 写入J2:L4
 * arr.toRange(Range("A1"));             // 写入A1:C4
 */
Array.prototype.toRange = function(rng) {
    if (!isWPS) return null;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = this.length;
    var cols = rows > 0 ? (Array.isArray(this[0]) ? this[0].length : 1) : 0;
    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = this;
    return writeRng;
};

/**
 * 数组写入单元格（数组扩展方法 - 中文别名）
 */
Array.prototype.z写入单元格 = Array.prototype.toRange;

/**
 * 截取字符
 * @param {String} str - 字符串
 * @param {Number} start - 起始位置（从1开始）
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z截取字符 = function(str, start, len) {
    var startIndex = start - 1;
    if (len === undefined) return str.substring(startIndex);
    return str.substring(startIndex, startIndex + len);
};
JSA.mid = JSA.z截取字符;

/**
 * 左取字符
 * @param {String} str - 字符串
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z左取字符 = function(str, len) {
    return str.substring(0, len);
};
JSA.left = JSA.z左取字符;

/**
 * 右取字符
 * @param {String} str - 字符串
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z右取字符 = function(str, len) {
    return str.substring(str.length - len);
};
JSA.right = JSA.z右取字符;

/**
 * 求和
 * @param {...Number} args - 数值
 * @returns {Number} 和
 */
JSA.z求和 = function() {
    return Array.prototype.slice.call(arguments).reduce(function(acc, val) {
        return acc + (Number(val) || 0);
    }, 0);
};
JSA.sum = JSA.z求和;

/**
 * 最大值
 * @param {...Number} args - 数值
 * @returns {Number} 最大值
 */
JSA.z最大值 = function() {
    return Math.max.apply(null, Array.prototype.slice.call(arguments).map(function(v) { return Number(v) || 0; }));
};
JSA.max = JSA.z最大值;

/**
 * 最小值
 * @param {...Number} args - 数值
 * @returns {Number} 最小值
 */
JSA.z最小值 = function() {
    return Math.min.apply(null, Array.prototype.slice.call(arguments).map(function(v) { return Number(v) || 0; }));
};
JSA.min = JSA.z最小值;

/**
 * 平均值
 * @param {...Number} args - 数值
 * @returns {Number} 平均值
 */
JSA.z平均值 = function() {
    var args = Array.prototype.slice.call(arguments);
    return args.length > 0 ? JSA.z求和.apply(null, args) / args.length : 0;
};
JSA.average = JSA.z平均值;

/**
 * 模糊匹配
 * @param {String} str - 字符串
 * @param {String} pattern - 模式（支持*和?）
 * @returns {Number} 匹配返回-1，不匹配返回0
 * @description 包含模式匹配，自动在模式前后添加 *，除非模式以 ^ 开头或 $ 结尾
 */
JSA.z模糊匹配 = function(str, pattern) {
    if (pattern === undefined || pattern === null) return 0;
    // 转义正则特殊字符，但保留 * 和 ?
    var regexPattern = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&')
                              .replace(/\*/g, '.*')
                              .replace(/\?/g, '.');
    // 包含模式：自动在前后加上 .*，除非模式已经以 ^ 开头或 $ 结尾
    var anchored = regexPattern;
    if (anchored.indexOf('^') !== 0) anchored = '.*' + anchored;
    if (anchored.charAt(anchored.length - 1) !== '$') anchored = anchored + '.*';
    var regex = new RegExp('^' + anchored + '$');
    // 匹配返回-1，不匹配返回0
    return regex.test(str) ? -1 : 0;
};
JSA.like = JSA.z模糊匹配;

/**
 * 表达式求值
 * @param {String} expr - 字符串表达式（如 '5*6+5'）
 * @returns {Number} 计算结果
 * @description 对字符串表达式进行求值计算
 * @example JSA.eval880('5*6+5')     // 返回 35
 * @example JSA.eval880('10+20*3')   // 返回 70
 * @example JSA.eval880('(1+2)*3')   // 返回 9
 */
JSA.z表达式求值 = function(expr) {
    if (typeof expr !== 'string') return Number(expr) || 0;
    // 使用 Function 构造函数安全地计算表达式
    try {
        return new Function('return ' + expr)();
    } catch (e) {
        return 0;
    }
};
JSA.eval880 = JSA.z表达式求值;

/**
 * 生成数字序列
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @param {Number} step - 步长
 * @returns {Array} 序列
 */
JSA.z生成数字序列 = function(start, end, step) {
    step = step || 1;
    var result = [];
    for (var i = start; i <= end; i += step) {
        result.push(i);
    }
    return result;
};
JSA.getNumberArray = JSA.z生成数字序列;

/**
 * 人民币大写
 * @param {Number} n - 数字
 * @returns {String} 大写
 */
JSA.z人民币大写 = function(n) {
    var digits = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
    var units = ["", "拾", "佰", "仟"];
    var bigUnits = ["", "万", "亿"];

    if (n === 0) return "零元整";

    var num = Math.abs(n);
    var integerPart = Math.floor(num);
    var decimalPart = Math.round((num - integerPart) * 100);

    var result = _convertIntegerPart(integerPart, digits, units, bigUnits) + "元";

    if (decimalPart > 0) {
        if (decimalPart >= 10) {
            var jiao = Math.floor(decimalPart / 10);
            var fen = decimalPart % 10;
            result += digits[jiao] + "角";
            if (fen > 0) result += digits[fen] + "分";
        } else {
            result += digits[decimalPart] + "分";
        }
    } else {
        result += "整";
    }

    if (n < 0) result += "（负）";
    return result;

    function _convertIntegerPart(num, digits, units, bigUnits) {
        if (num === 0) return "";
        var result = "";
        var bigUnitIndex = 0;
        while (num > 0) {
            var section = num % 10000;
            if (section > 0) {
                var sectionResult = _convertSection(section, digits, units);
                result = sectionResult + bigUnits[bigUnitIndex] + result;
            }
            num = Math.floor(num / 10000);
            bigUnitIndex++;
        }
        return result;
    }

    function _convertSection(num, digits, units) {
        var result = "";
        var unitIndex = 0;
        var lastZero = false;
        while (num > 0) {
            var digit = num % 10;
            if (digit === 0) {
                if (!lastZero && result !== "") {
                    result = digits[0] + result;
                    lastZero = true;
                }
            } else {
                result = digits[digit] + units[unitIndex] + result;
                lastZero = false;
            }
            num = Math.floor(num / 10);
            unitIndex++;
        }
        return result;
    }
};
JSA.rmbdx = JSA.z人民币大写;

/**
 * 随机整数
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @returns {Number} 随机整数
 */
JSA.z随机整数 = function(start, end) {
    return Math.floor(Math.random() * (end - start + 1)) + start;
};
JSA.rndInt = JSA.z随机整数;

/**
 * 随机打乱
 * @param {Array} array - 数组
 * @returns {Array} 打乱后的数组
 */
JSA.z随机打乱 = function(array) {
    var result = array.slice();
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return result;
};
JSA.shuffle = JSA.z随机打乱;

/**
 * 延时
 * @param {Number} ts - 毫秒
 */
JSA.z延时 = function(ts) {
    var start = Date.now();
    while (Date.now() - start < ts) {
        // 等待
    }
};
JSA.delay = JSA.z延时;

/**
 * 日期间隔
 * @param {Date|string} d1 - 日期1
 * @param {Date|string} d2 - 日期2
 * @param {String} format - 格式
 * @returns {String|Number} 间隔
 */
JSA.z日期间隔 = function(d1, d2, format) {
    var date1 = typeof d1 === 'string' ? new Date(d1) : d1;
    var date2 = typeof d2 === 'string' ? new Date(d2) : d2;

    if (format === 'Y') return date2.getFullYear() - date1.getFullYear();
    if (format === 'M') {
        var years = date2.getFullYear() - date1.getFullYear();
        var months = date2.getMonth() - date1.getMonth();
        return years * 12 + months;
    }
    if (format === 'D') {
        var msPerDay = 24 * 60 * 60 * 1000;
        return Math.round((date2 - date1) / msPerDay);
    }
    // 默认返回完整间隔
    var years = date2.getFullYear() - date1.getFullYear();
    var months = date2.getMonth() - date1.getMonth();
    var days = date2.getDate() - date1.getDate();

    if (days < 0) {
        months--;
        var prevMonth = new Date(date2.getFullYear(), date2.getMonth(), 0);
        days += prevMonth.getDate();
    }
    if (months < 0) {
        years--;
        months += 12;
    }

    var result = "";
    if (years > 0) result += years + "年";
    if (months > 0) result += months + "个月";
    if (days > 0) result += days + "天";
    return result || "0天";
};
JSA.datedif = JSA.z日期间隔;

/**
 * 选择列
 * @param {Array} arr - 二维数组
 * @param {Array} colIndexes - 列索引
 * @param {Array} newHeaders - 新表头
 * @returns {Array} 结果数组
 */
JSA.z选择列 = function(arr, colIndexes, newHeaders) {
    if (!arr || arr.length === 0) return [];

    var indexes = [];

    // 检查是否按表头选择
    if (arr.length > 0 && colIndexes.length > 0 && typeof colIndexes[0] === 'string') {
        var headers = arr[0];
        var headerMap = {};
        for (var i = 0; i < headers.length; i++) {
            headerMap[String(headers[i])] = i;
        }

        for (var j = 0; j < colIndexes.length; j++) {
            var col = colIndexes[j];
            if (headerMap.hasOwnProperty(col)) {
                indexes.push(headerMap[col]);
            }
        }

        var result = [];
        if (newHeaders && newHeaders.length > 0) {
            result.push(newHeaders);
        } else {
            var newRow = [];
            for (var k = 0; k < colIndexes.length; k++) {
                var col = colIndexes[k];
                var idx = headerMap[col];
                newRow.push(idx !== undefined ? headers[idx] : col);
            }
            result.push(newRow);
        }

        for (var i = 1; i < arr.length; i++) {
            var row = arr[i];
            var newRow = [];
            for (var k = 0; k < indexes.length; k++) {
                newRow.push(row[indexes[k]]);
            }
            result.push(newRow);
        }

        return result;
    } else {
        // 按列号选择
        indexes = [];
        for (var j = 0; j < colIndexes.length; j++) {
            indexes.push(typeof colIndexes[j] === 'number' ? colIndexes[j] : parseInt(colIndexes[j]));
        }

        var result = [];
        for (var i = 0; i < arr.length; i++) {
            var row = arr[i];
            var newRow = [];
            for (var k = 0; k < indexes.length; k++) {
                newRow.push(row[indexes[k]]);
            }
            result.push(newRow);
        }

        return result;
    }
};
JSA.selectCols = JSA.z选择列;

/**
 * 矩阵分布
 * @param {Number} totalRows - 总行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'或'c'
 * @returns {Array} 分布后的数组
 */
JSA.z矩阵分布 = function(totalRows, cols, direction) {
    direction = direction || 'r';
    var result = [];
    var numbers = [];
    for (var i = 0; i < totalRows; i++) {
        numbers.push(i);
    }

    if (direction === 'r') {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = i * cols + j;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    } else {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = j * rows + i;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    }

    return result;
};
JSA.getMatrix = JSA.z矩阵分布;

// ==================== IO - 文件操作库 ====================

/**
 * IO - 文件操作工具
 * @class
 * @description 文件系统操作
 */
function IO() {}

/**
 * 是否文件
 * @param {String} path - 路径
 * @returns {Boolean} 是否为文件
 */
IO.z是否文件 = function(path) {
    if (!isWPS) return false;
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FileExists(path);
    } catch (e) {
        return false;
    }
};
IO.IsFile = IO.z是否文件;

/**
 * 是否文件夹
 * @param {String} path - 路径
 * @returns {Boolean} 是否为文件夹
 */
IO.z是否文件夹 = function(path) {
    if (!isWPS) return false;
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FolderExists(path);
    } catch (e) {
        return false;
    }
};
IO.IsDirectory = IO.z是否文件夹;

/**
 * 文件名
 * @param {String} path - 路径
 * @returns {String} 文件名
 */
IO.z文件名 = function(path) {
    if (!path) return '';
    var parts = path.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1] || '';
};
IO.getFileName = IO.z文件名;

/**
 * 纯文件名
 * @param {String} path - 路径
 * @returns {String} 纯文件名
 */
IO.z纯文件名 = function(path) {
    var fileName = IO.z文件名(path);
    var lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex > 0) {
        return fileName.substring(0, lastDotIndex);
    }
    return fileName;
};
IO.getFileNameNoType = IO.z纯文件名;

/**
 * 文件后缀
 * @param {String} path - 路径
 * @returns {String} 后缀
 */
IO.z文件后缀 = function(path) {
    var fileName = IO.z文件名(path);
    var lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex > 0 && lastDotIndex < fileName.length - 1) {
        return fileName.substring(lastDotIndex + 1);
    }
    return '';
};
IO.getFileType = IO.z文件后缀;

/**
 * 上级文件夹
 * @param {String} path - 路径
 * @param {Number} 级数 - 级数
 * @returns {String} 上级路径
 */
IO.z上级文件夹 = function(path, 级数) {
    级数 = 级数 || 1;
    var result = path;
    for (var i = 0; i < 级数; i++) {
        result = result.replace(/\\/g, '/').replace(/\/+$/, '');
        var lastSlashIndex = result.lastIndexOf('/');
        if (lastSlashIndex > 0) {
            result = result.substring(0, lastSlashIndex);
        } else {
            break;
        }
    }
    return result;
};
IO.lastDirectoty = IO.z上级文件夹;

// ==================== 全局辅助函数 ====================

/**
 * 日志输出
 * @param {...any} args - 参数
 */
function log() {
    if (isWPS && typeof Console !== 'undefined') {
        Array.prototype.slice.call(arguments).forEach(function(arg) {
            Console.log(arg);
        });
    } else {
        console.log.apply(console, arguments);
    }
}

/**
 * 批量创建中英文方法别名
 * @private
 * @param {Object} prototype - 原型对象
 * @param {Array} aliases - 别名配置数组 [[中文名, 英文名], ...]
 */
function createBilingualAliases(prototype, aliases) {
    for (var i = 0; i < aliases.length; i++) {
        var cnName = aliases[i][0];
        var enName = aliases[i][1];
        if (prototype[cnName] && !prototype[enName]) {
            prototype[enName] = prototype[cnName];
        } else if (prototype[enName] && !prototype[cnName]) {
            prototype[cnName] = prototype[enName];
        }
    }
}

/**
 * JSON日志输出
 * @param {any} x - 对象
 * @param {Boolean} wrapopt - 是否包装JSON对象(即是否要输出日期等信息)，默认true
 * @example
 * logjson([[1,2],[3,4],[5,6]],0);  // 输出: [[1,2],[3,4],[5,6]]
 * logjson([1,2,3])                  // 一维数组输出为紧凑单行
 */
function logjson(x, wrapopt) {
    wrapopt = wrapopt !== undefined ? wrapopt : true;

    // 处理 Array2D 对象（提取 _items 属性）
    if (x && typeof x === 'object' && x._items && Array.isArray(x._items)) {
        x = x._items;
    }

    // 二维数组特殊处理
    if (Array.isArray(x) && x.length > 0 && Array.isArray(x[0])) {
        // wrapopt=0 时输出紧凑格式
        if (wrapopt === false || wrapopt === 0) {
            var output = JSON.stringify(x);
            if (isWPS && typeof Console !== 'undefined') {
                Console.log(output);
            } else {
                console.log(output);
            }
        } else {
            // 格式化输出（对齐）
            var lines = formatArray2DAsJSON(x);
            for (var i = 0; i < lines.length; i++) {
                if (isWPS && typeof Console !== 'undefined') {
                    Console.log(lines[i]);
                } else {
                    console.log(lines[i]);
                }
            }
        }
        return;
    }

    // 一维数组输出为紧凑单行格式
    if (Array.isArray(x)) {
        var str = '[' + x.map(function(item) {
            if (item === null || item === undefined) return '';
            return String(item);
        }).join(',') + ']';
        if (isWPS && typeof Console !== 'undefined') {
            Console.log(str);
        } else {
            console.log(str);
        }
        return;
    }

    // 其他类型：处理循环引用和日期
    var output;
    if (wrapopt && typeof x === 'object' && x !== null) {
        var seen = new WeakSet();
        var replacer = function(key, value) {
            if (typeof value === 'object' && value !== null) {
                if (seen.has(value)) {
                    return '[Circular]';
                }
                seen.add(value);
            }
            if (value instanceof Date) {
                return value.toISOString();
            }
            return value;
        };
        output = JSON.stringify(x, replacer, 2);
    } else {
        output = typeof x === 'object' ? JSON.stringify(x, null, wrapopt ? 2 : 0) : String(x);
    }

    if (isWPS && typeof Console !== 'undefined') {
        Console.log(output);
    } else {
        console.log(output);
    }

    return;
}

/**
 * 格式化二维数组为JSON（支持对齐显示）
 * @private
 * @param {Array} arr - 二维数组
 * @returns {Array} 格式化的字符串数组
 */
function formatArray2DAsJSON(arr) {
    if (!arr || arr.length === 0) return ['[]'];

    /**
     * 计算字符串的显示宽度（基于等宽字体环境）
     * 规则：
     * - ASCII 字符（U+0000 - U+007F）= 1
     * - 非ASCII 字符（包括中文等宽字符）= 2
     */
    var getDisplayWidth = function(str) {
        var width = 0;
        for (var i = 0; i < str.length; i++) {
            var code = str.charCodeAt(i);
            if (code < 128) {
                // ASCII 字符宽度为 1
                width += 1;
            } else {
                // 非ASCII 字符（包括中文）宽度为 2
                width += 2;
            }
        }
        return width;
    };

    // 先将每行转换为字符串，以便计算显示宽度
    var stringRows = [];
    var colCount = arr[0].length;

    for (var row = 0; row < arr.length; row++) {
        var stringCells = [];
        for (var col = 0; col < colCount; col++) {
            var cellValue = col < arr[row].length ? arr[row][col] : '';
            var cellStr = cellValue === null || cellValue === undefined ? '' : String(cellValue);
            stringCells.push(cellStr);
        }
        stringRows.push(stringCells);
    }

    // 计算每列内容的最大显示宽度（不包括引号和逗号）
    var contentWidths = [];
    for (var col = 0; col < colCount; col++) {
        var maxWidth = 0;
        for (var row = 0; row < arr.length; row++) {
            maxWidth = Math.max(maxWidth, getDisplayWidth(stringRows[row][col]));
        }
        contentWidths.push(maxWidth);
    }

    var lines = [];

    // 构建所有行，确保对齐
    for (var row = 0; row < arr.length; row++) {
        var rowParts = [];
        for (var col = 0; col < colCount; col++) {
            var cellStr = stringRows[row][col];
            var displayWidth = getDisplayWidth(cellStr);

            // 计算需要填充的宽度
            var paddingNeeded = contentWidths[col] - displayWidth;

            // 使用普通空格填充（每个空格占1个显示宽度）
            // 使用 Array().join() 替代 String.repeat() 以提升兼容性
            var paddingStr = paddingNeeded > 0 ? Array(paddingNeeded + 1).join(' ') : '';

            // 构建单元格：前面填充 + "内容"
            var cell = '"' + paddingStr + cellStr + '"';

            rowParts.push(cell);
        }

        // 用逗号连接各列（逗号后无空格）
        var rowStr = '[' + rowParts.join(',') + ']';
        lines.push(rowStr);
    }

    // 添加前导空格和行尾逗号
    for (var i = 0; i < lines.length; i++) {
        if (i < lines.length - 1) {
            lines[i] = ' ' + lines[i] + ',';
        } else {
            lines[i] = ' ' + lines[i];
        }
    }

    lines.push(']');
    lines.unshift('[');
    return lines;
}

// ==================== Global - 全局工具函数 ====================

/**
 * f1函数 - 在WPS JSA立即窗口快速打开JSA880帮助
 * @param {String} fxname - 函数名，如Array2D.pad
 * @example
 * f1("Array2D.pad")  // 打开帮助
 */
function f1(fxname) {
    if (!isWPS) return;
    // 构建帮助URL
    var helpUrl = "https://vbayyds.com/api/help/" + fxname;
    // 在WPS中打开浏览器
    try {
        var browser = new ActiveXObject("InternetExplorer.Application");
        browser.Visible = true;
        browser.Navigate(helpUrl);
    } catch (e) {
        Console.log("帮助地址: " + helpUrl);
    }
}

/**
 * $fx函数 - WorksheetFunction对象的简写
 * @param {string} path - 函数对象的路径
 * @returns {Function} 工作表函数
 * @example
 * $fx.Sum(1,2,3)  // 6
 */
function $fx(path) {
    if (!isWPS) return null;
    var parts = path.split('.');
    var obj = WorksheetFunction;
    for (var i = 0; i < parts.length; i++) {
        if (obj[parts[i]]) {
            obj = obj[parts[i]];
        } else {
            return null;
        }
    }
    return typeof obj === 'function' ? obj : null;
}

/**
 * $toArray函数 - 将参数转换为数组（内部使用）
 * @param {...any} args - 要转换为数组的参数
 * @returns {Array} 转换后的数组
 * @example
 * $toArray("产品1", "产品2", "产品3")  // ["产品1","产品2","产品3"]
 */
function $toArray() {
    var result = [];
    for (var i = 0; i < arguments.length; i++) {
        result.push(arguments[i]);
    }
    return result;
}

// ==================== 类型转换函数 (as系列) ====================

/**
 * asString函数 - 将对象转换为字符串对象
 * @param {any} s - 要转换的对象
 * @returns {String} 字符串
 * @example
 * asString(123)  // "123"
 */
function asString(s) {
    return String(s === null || s === undefined ? '' : s);
}

/**
 * asArray函数 - 将值转换为Array2D对象（支持链式调用和toRange）
 * @param {any} a - 要转换的值
 * @returns {Array2D} Array2D对象
 * @example
 * asArray(123)                      // Array2D([[123]])
 * asArray("abc")                    // Array2D([["abc"]])
 * asArray([1,2,3])                  // Array2D([[1],[2],[3]])
 * asArray([[1,2],[3,4]])            // Array2D([[1,2],[3,4]])
 * asArray(Array2D([[1,2]]))         // Array2D([[1,2]]) (原样返回)
 * asArray("a,b,c")                  // Array2D([["a"],["b"],["c"]])
 */
function asArray(a) {
    // 如果已经是 Array2D，直接返回
    if (a instanceof Array2D) return a;

    var arr;
    if (Array.isArray(a)) {
        arr = a;
    } else if (a === null || a === undefined) {
        arr = [];
    } else if (typeof a === 'string') {
        // 尝试按逗号分割
        if (a.indexOf(',') >= 0) {
            var parts = a.split(',').map(function(s) { return s.trim(); });
            // 转为二维数组
            arr = [];
            for (var i = 0; i < parts.length; i++) {
                arr.push([parts[i]]);
            }
        } else {
            arr = [[a]];
        }
    } else {
        arr = [[a]];
    }

    // 确保 arr 是二维数组
    if (arr.length > 0 && !Array.isArray(arr[0])) {
        var newArr = [];
        for (var j = 0; j < arr.length; j++) {
            newArr.push([arr[j]]);
        }
        arr = newArr;
    }

    return new Array2D(arr);
}

/**
 * asArray2D函数 - 将值转换为Array2D对象（asArray的别名）
 * @param {any} a - 要转换的值
 * @returns {Array2D} Array2D对象
 * @example
 * asArray2D([[1,2],[3,4]])           // Array2D([[1,2],[3,4]])
 * asArray2D([1,2,3])                  // Array2D([[1],[2],[3]])
 * asArray2D("a,b,c")                  // Array2D([["a"],["b"],["c"]])
 * asArray2D(Array2D([[1,2]]))         // Array2D([[1,2]]) (原样返回)
 */
var asArray2D = asArray;

/**
 * asNumber函数 - 将值转换为数字
 * @param {any} a - 要转换的值
 * @returns {Number} 数字，转换失败返回0
 * @example
 * asNumber("123")        // 123
 * asNumber("12.34")      // 12.34
 * asNumber("abc")        // 0
 * asNumber(null)         // 0
 */
function asNumber(a) {
    if (typeof a === 'number') return a;
    if (typeof a === 'boolean') return a ? 1 : 0;
    if (a === null || a === undefined || a === '') return 0;
    var num = Number(a);
    return isNaN(num) ? 0 : num;
}

/**
 * asDate函数 - 将值转换为DateUtils对象（支持智能提示和链式调用）
 * @param {any} a - 要转换的值
 * @returns {DateUtils} DateUtils实例
 * @example
 * asDate("2023-9-1").z月份()     // 9
 * asDate(45170).z年份()          // 2023 (Excel日期序号)
 * asDate("2023/09/01").z日期()   // 1
 */
function asDate(a) {
    var date;
    if (a instanceof DateUtils) return a;
    if (a instanceof Date) {
        date = a;
    } else if (typeof a === 'number') {
        // Excel日期序号转JS Date
        date = new Date((a - 25569) * 86400 * 1000);
    } else if (typeof a === 'string') {
        date = new Date(a);
        if (isNaN(date.getTime())) {
            date = new Date();
        }
    } else {
        date = new Date();
    }
    return new DateUtils(date);
}

/**
 * asRange函数 - 将值转换为Range对象
 * @param {any} a - 要转换的值（地址字符串、Range对象等）
 * @returns {Range|null} Range对象
 * @example
 * asRange("A1")          // Range对象
 * asRange(Range("A1"))   // Range对象
 * asRange("A1:C10")      // Range对象
 */
function asRange(a) {
    if (!isWPS) return null;
    if (a && a.Address) return a; // 已经是Range对象
    if (typeof a === 'string') {
        try {
            return Range(a);
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * asMap函数 - 将值转换为Map对象
 * @param {any} a - 要转换的值（对象、Map、二维数组等）
 * @returns {Map} Map对象
 * @example
 * asMap({a:1,b:2})       // Map(2) {"a"=>1,"b"=>2}
 * asMap([['a',1],['b',2]])// Map(2) {"a"=>1,"b"=>2}
 */
function asMap(a) {
    if (a instanceof Map) return a;
    var map = new Map();
    if (a === null || a === undefined) return map;
    if (Array.isArray(a)) {
        // 二维数组转Map: [['key','value'],...]
        a.forEach(function(item) {
            if (Array.isArray(item) && item.length >= 2) {
                map.set(item[0], item[1]);
            }
        });
    } else if (typeof a === 'object') {
        // 对象转Map
        for (var key in a) {
            if (a.hasOwnProperty(key)) {
                map.set(key, a[key]);
            }
        }
    }
    return map;
}

/**
 * asObject函数 - 将值转换为普通对象
 * @param {any} a - 要转换的值（Map、对象等）
 * @returns {Object} 普通对象
 * @example
 * asObject(new Map([['a',1],['b',2]]))  // {a:1,b:2}
 * asObject({a:1})                        // {a:1}
 */
function asObject(a) {
    if (a instanceof Map) {
        var obj = {};
        a.forEach(function(value, key) {
            obj[key] = value;
        });
        return obj;
    }
    if (typeof a === 'object' && a !== null) {
        return a;
    }
    return {};
}

/**
 * asShape函数 - 将对象转换为Shape对象
 * @param {any} shp - 要转换的对象
 * @returns {Shape|null} Shape对象
 * @example
 * asShape('矩形 2')  // Shape对象
 */
function asShape(shp) {
    if (!isWPS) return null;
    if (typeof shp === 'string') {
        // 遍历所有工作表的形状
        for (var i = 1; i <= Sheets.Count; i++) {
            var sht = Sheets(i);
            for (var j = 1; j <= sht.Shapes.Count; j++) {
                if (sht.Shapes(j).Name === shp) return sht.Shapes(j);
                if (sht.Shapes(j).Name.indexOf(shp) !== -1) return sht.Shapes(j);
            }
        }
        return null;
    }
    if (shp && shp.Name) return shp;
    return null;
}

// ==================== SheetChain - 工作表链式调用类 ====================

/**
 * SheetChain - 工作表链式调用包装类（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 包装WPS工作表对象，提供链式调用和智能提示
 * @example
 * asSheet("Sheet1").z激活().z名称()
 * asSheet(1).z已使用区域().z安全数组()
 */
function SheetChain(sht) {
    if (!(this instanceof SheetChain)) {
        return new SheetChain(sht);
    }
    this._sheet = null;

    // 检查WPS环境和Sheets可用性
    if (typeof Sheets === 'undefined') return;

    // 如果已经是Sheet对象，直接使用
    if (sht && sht.Activate && sht.Name) {
        this._sheet = sht;
        return;
    }

    if (typeof sht === 'number') {
        try {
            this._sheet = Sheets(sht);
        } catch (e) {
            this._sheet = null;
        }
        return;
    }

    if (typeof sht === 'string') {
        try {
            // 首先尝试精确匹配
            this._sheet = Sheets(sht);
        } catch (e) {
            // 精确匹配失败，尝试模糊匹配
            try {
                for (var i = 1; i <= Sheets.Count; i++) {
                    var sheet = Sheets(i);
                    // 包含匹配
                    if (sheet.Name.indexOf(sht) >= 0) {
                        this._sheet = sheet;
                        return;
                    }
                    // 忽略大小写匹配
                    if (sheet.Name.toLowerCase() === sht.toLowerCase()) {
                        this._sheet = sheet;
                        return;
                    }
                }
            } catch (e2) {
                console.log("SheetChain模糊匹配失败: " + e2.message);
            }
            this._sheet = null;
        }
        return;
    }
}

/**
 * Value - 获取原始Sheet对象
 * @returns {Worksheet|null} 工作表对象
 */
SheetChain.prototype.Value = function() {
    return this._sheet;
};

/**
 * Name - 获取工作表名称
 * @returns {String} 工作表名称
 */
SheetChain.prototype.z名称 = function() {
    return this._sheet ? this._sheet.Name : '';
};
SheetChain.prototype.Name = SheetChain.prototype.z名称;

/**
 * Activate - 激活工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z激活 = function() {
    if (this._sheet) this._sheet.Activate();
    return this;
};
SheetChain.prototype.Activate = SheetChain.prototype.z激活;

/**
 * UsedRange - 获取已使用区域
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z已使用区域 = function() {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.UsedRange);
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.UsedRange = SheetChain.prototype.z已使用区域;

/**
 * SafeUsedRange - 获取安全已使用区域（处理空表情况）
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z安全已使用区域 = function() {
    if (!this._sheet) return null;

    var usedRange;
    try {
        usedRange = this._sheet.UsedRange;
    } catch (e) {
        return new RangeChain(this._sheet.Range("A1"));
    }

    if (!usedRange) return new RangeChain(this._sheet.Range("A1"));

    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;

    return new RangeChain(this._sheet.Range(this._sheet.Cells(1, 1), this._sheet.Cells(lastRow, lastCol)));
};
SheetChain.prototype.SafeUsedRange = SheetChain.prototype.z安全已使用区域;

/**
 * Range - 获取Range对象
 * @param {String} address - 地址
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z区域 = function(address) {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.Range(address));
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.Range = SheetChain.prototype.z区域;

/**
 * Cells - 获取Cells对象
 * @param {Number} row - 行号
 * @param {Number} col - 列号
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z单元格 = function(row, col) {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.Cells(row, col));
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.Cells = SheetChain.prototype.z单元格;

/**
 * Delete - 删除工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z删除 = function() {
    if (this._sheet) {
        try {
            this._sheet.Delete();
        } catch (e) {
            console.log("删除工作表失败: " + e.message);
        }
    }
    return this;
};
SheetChain.prototype.Delete = SheetChain.prototype.z删除;

/**
 * Copy - 复制工作表
 * @param {Worksheet} [before] - 在此工作表之前插入
 * @param {Worksheet} [after] - 在此工作表之后插入
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z复制 = function(before, after) {
    if (!this._sheet) return this;
    try {
        if (before) {
            this._sheet.Copy(before);
        } else if (after) {
            this._sheet.Copy(undefined, after);
        } else {
            this._sheet.Copy();
        }
    } catch (e) {
        console.log("复制工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Copy = SheetChain.prototype.z复制;

/**
 * Protect - 保护工作表
 * @param {String} [password] - 密码
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z保护 = function(password) {
    if (!this._sheet) return this;
    try {
        if (password) {
            this._sheet.Protect(password);
        } else {
            this._sheet.Protect();
        }
    } catch (e) {
        console.log("保护工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Protect = SheetChain.prototype.z保护;

/**
 * Unprotect - 取消保护工作表
 * @param {String} [password] - 密码
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z取消保护 = function(password) {
    if (!this._sheet) return this;
    try {
        if (password) {
            this._sheet.Unprotect(password);
        } else {
            this._sheet.Unprotect();
        }
    } catch (e) {
        console.log("取消保护工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Unprotect = SheetChain.prototype.z取消保护;

/**
 * 隐藏工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z隐藏 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = false;
    return this;
};

/**
 * 显示工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z显示 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = true;
    return this;
};

/**
 * Index - 获取工作表索引
 * @returns {Number} 工作表索引
 */
SheetChain.prototype.z索引 = function() {
    return this._sheet ? this._sheet.Index : 0;
};
SheetChain.prototype.Index = SheetChain.prototype.z索引;

/**
 * SetName - 设置工作表名称
 * @param {String} newName - 新名称
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z设置名称 = function(newName) {
    if (this._sheet) {
        this._sheet.Name = newName;
    }
    return this;
};
SheetChain.prototype.SetName = SheetChain.prototype.z设置名称;

/**
 * Exists - 判断工作表是否存在
 * @returns {Boolean} 是否存在
 */
SheetChain.prototype.z存在 = function() {
    return this._sheet !== null;
};
SheetChain.prototype.Exists = SheetChain.prototype.z存在;

// ==================== 类型转换函数 ====================

/**
 * asSheet函数 - 将对象转换为SheetChain对象（支持智能提示和链式调用）
 * @param {any} sht - 要转换的对象
 * @returns {SheetChain} SheetChain实例
 * @example
 * asSheet("1月").z激活().z名称()
 * asSheet(1).z已使用区域().z安全数组()
 * asSheet().z激活()
 */
function asSheet(sht) {
    return new SheetChain(sht);
}

/**
 * asWorkbook函数 - 将对象转换为工作簿对象
 * @param {any} wbk - 要转换的对象
 * @returns {Workbook} 工作簿对象
 * @example
 * asWorkbook("测试排序")  // 工作簿对象
 */
function asWorkbook(wbk) {
    if (!isWPS) return null;
    if (typeof wbk === 'string') {
        for (var i = 1; i <= Workbooks.Count; i++) {
            if (Workbooks(i).Name === wbk) return Workbooks(i);
        }
        return null;
    }
    if (wbk && wbk.Name) return wbk;
    return null;
}

// ==================== As - 类型转换包装类 ====================

/**
 * As类 - 类型转换包装类（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 提供类型转换和常用操作方法，支持中英双语API
 * @example
 * // 基本使用
 * As([[1,2,3],[4,5,6]]).toArray().z求和()        // 21
 * As("123").toNumber()                           // 123
 * As(123).toString()                             // "123"
 * // 链式调用
 * As([[1,2],[3,4]]).toArray().z转置().z扁平化().val()  // [1,3,2,4]
 */
function As(value) {
    // 支持工厂模式调用
    if (!(this instanceof As)) {
        return new As(value);
    }

    this._original = value;
    this._value = value;
}

/**
 * 创建新实例（链式调用核心）
 * @private
 * @param {any} data - 新值
 * @returns {As} 新实例
 */
As.prototype._new = function(value) {
    const instance = new As();
    instance._original = this._original;
    instance._value = value;
    return instance;
};

/**
 * 获取/设置当前值
 * @param {any} [newValue] - 新值（可选）
 * @returns {As|any} 设置时返回this，否则返回当前值
 * @example
 * As(123).val()           // 123
 * As(123).val(456)        // 返回链式对象
 */
As.prototype.val = function(newValue) {
    if (newValue !== undefined) {
        this._value = newValue;
        return this;
    }
    return this._value;
};

// ==================== 类型转换方法 ====================

/**
 * 转换为数组
 * @returns {Array2D} 二维数组工具对象（如果是二维数组）或 As包装对象
 * @example
 * As([1,2,3]).toArray()              // [1,2,3]
 * As("a,b,c").toArray()              // ["a","b","c"]
 * As([[1,2],[3,4]]).toArray()        // Array2D对象，支持链式调用
 */
As.prototype.toArray = function() {
    const arr = asArray(this._value);
    // 如果是二维数组，返回 Array2D 对象以获得更多功能
    if (arr.length > 0 && Array.isArray(arr[0])) {
        return Array2D(arr);
    }
    return this._new(arr);
};

/**
 * 转换为数字
 * @returns {As} 包装对象
 * @example
 * As("123").toNumber().val()         // 123
 * As("abc").toNumber().val()         // 0
 */
As.prototype.toNumber = function() {
    return this._new(asNumber(this._value));
};

/**
 * 转换为字符串
 * @returns {As} 包装对象
 * @example
 * As(123).toString().val()           // "123"
 * As(null).toString().val()          // ""
 */
As.prototype.toString = function() {
    return this._new(asString(this._value));
};

/**
 * 转换为日期
 * @returns {As} 包装对象
 * @example
 * As("2023-9-1").toDate().val()      // Date对象
 * As(45170).toDate().val()           // Date对象
 */
As.prototype.toDate = function() {
    return this._new(asDate(this._value));
};

/**
 * 转换为Map对象
 * @returns {As} 包装对象
 * @example
 * As({a:1,b:2}).toMap().val()        // Map对象
 */
As.prototype.toMap = function() {
    return this._new(asMap(this._value));
};

/**
 * 转换为普通对象
 * @returns {As} 包装对象
 * @example
 * const map = new Map([['a',1]]);
 * As(map).toObject().val()           // {a:1}
 */
As.prototype.toObject = function() {
    return this._new(asObject(this._value));
};

/**
 * 转换为Range对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("A1:C10").toRange().val()       // Range对象
 */
As.prototype.toRange = function() {
    const rng = asRange(this._value);
    return rng !== null ? this._new(rng) : null;
};

/**
 * 转换为工作表对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("Sheet1").toSheet().val()       // Worksheet对象
 */
As.prototype.toSheet = function() {
    const sht = asSheet(this._value);
    return sht !== null ? this._new(sht) : null;
};

/**
 * 转换为工作簿对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("工作簿1.xlsx").toWorkbook().val()  // Workbook对象
 */
As.prototype.toWorkbook = function() {
    const wbk = asWorkbook(this._value);
    return wbk !== null ? this._new(wbk) : null;
};

// ==================== 中文别名 ====================
As.prototype.z转数组 = As.prototype.toArray;
As.prototype.z转数字 = As.prototype.toNumber;
As.prototype.z转字符串 = As.prototype.toString;
As.prototype.z转日期 = As.prototype.toDate;
As.prototype.z转Map = As.prototype.toMap;
As.prototype.z转对象 = As.prototype.toObject;

/**
 * cdate函数 - 将日期转换为Excel日期数值
 * @param {any} v - 日期字符串或JS日期对象
 * @returns {Number} Excel日期数值
 * @example
 * cdate('2023-9-1')  // 45170
 */
function cdate(v) {
    if (typeof v === 'number') return v;
    var date;
    if (typeof v === 'string') {
        // 处理简短日期格式
        if (v.match(/^\d{1,2}-\d{1,2}$/)) {
            v = '20' + v;  // 23-9-1 -> 2023-9-1
        }
        date = new Date(v);
    } else if (v instanceof Date) {
        date = v;
    } else {
        return 0;
    }
    var excelEpoch = new Date(1900, 0, 1);
    var msPerDay = 24 * 60 * 60 * 1000;
    return Math.floor((date - excelEpoch) / msPerDay) + 2;
}

/**
 * cstr函数 - 将值转换为字符串
 * @param {any} v - 要转换的值
 * @returns {String} 字符串
 * @example
 * cstr(1537789)  // "1537789"
 */
const cstr = (v) => v === null || v === undefined ? '' : String(v);

// ==================== 类型检查函数 (is系列) ====================

/**
 * isArray函数 - 检查值是否为数组
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为数组
 * @example
 * isArray([1,2,3])  // true
 */
const isArray = (v) => Array.isArray(v);

/**
 * isArray2D函数 - 检查值是否为二维数组
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为二维数组
 * @example
 * isArray2D([[1],[2],[3]])  // true
 */
const isArray2D = (v) => {
    if (!Array.isArray(v)) return false;
    if (v.length === 0) return false;
    return v.every(row => Array.isArray(row));
};

/**
 * isBoolean函数 - 检查值是否为布尔值
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为布尔值
 * @example
 * isBoolean(false)  // true
 */
const isBoolean = (v) => typeof v === 'boolean';

/**
 * isCollection函数 - 检查对象是否为集合对象
 * @param {any} obj - 要检查的对象
 * @returns {Boolean} 是否为集合对象
 * @example
 * isCollection(Sheets)  // true
 */
const isCollection = (obj) => {
    if (!obj) return false;
    // 检查是否是WPS集合对象
    if (obj && typeof obj === 'object') {
        // WPS集合对象通常有Count和Item属性
        if (obj.Count !== undefined && typeof obj.Item === 'unknown') return true;
        // 检查是否有枚举器
        try {
            const enumerator = new Enumerator(obj);
            return true;
        } catch (e) {
            // 不是集合
        }
    }
    return false;
};

/**
 * isDate函数 - 检查值是否为日期对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为日期对象
 * @example
 * isDate(new Date())  // true
 */
const isDate = (v) => v instanceof Date;

/**
 * isEmpty函数 - 检查值是否为空值
 * @param {any} value - 要检查的值
 * @returns {Boolean} 是否为空值
 * @example
 * isEmpty(undefined)  // true
 * isEmpty('')         // true
 * isEmpty(null)       // true
 */
const isEmpty = (value) => value === null || value === undefined || value === '';

/**
 * isNumberic函数 - 检查值是否为数值类型
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为数值类型
 * @example
 * isNumberic(557)  // true
 */
const isNumberic = (v) => typeof v === 'number' && !isNaN(v);

/**
 * isRange函数 - 检查值是否为Range对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为Range对象
 * @example
 * isRange(Range("A1"))  // true
 */
const isRange = (v) => isWPS && v && typeof v === 'object' && v.Address !== undefined;

/**
 * isRegex函数 - 检查值是否为正则表达式对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为正则表达式
 * @example
 * isRegex(/\d+/g)  // true
 */
const isRegex = (v) => v instanceof RegExp;

/**
 * isSameClass函数 - 检查两个值是否属于同一类别
 * @param {any} x - 第一个对象
 * @param {any} y - 第二个对象
 * @returns {Boolean} 是否属于同一类别
 * @example
 * isSameClass(560, 789)  // true
 */
const isSameClass = (x, y) => Object.prototype.toString.call(x) === Object.prototype.toString.call(y);

/**
 * isSheet函数 - 检查值是否为工作表对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为工作表对象
 * @example
 * isSheet(Sheets(1))  // true
 */
const isSheet = (v) => isWPS && v && typeof v === 'object' && v.Name !== undefined && v.Cells !== undefined;

/**
 * isString函数 - 检查值是否为字符串类型
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为字符串
 * @example
 * isString('产品5')  // true
 */
const isString = (v) => typeof v === 'string';

/**
 * isWorkbook函数 - 检查值是否为工作簿对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为工作簿对象
 * @example
 * isWorkbook(ActiveWorkbook)  // true
 */
const isWorkbook = (v) => isWPS && v && typeof v === 'object' && v.Name !== undefined && v.Sheets !== undefined && v.Close !== undefined;

/**
 * typeName函数 - 获取值的类型名称
 * @param {any} x - 要获取类型名称的值
 * @returns {String} 类型名称
 * @example
 * typeName('产品5')  // "[object String]"
 */
const typeName = (x) => Object.prototype.toString.call(x);

// ==================== 其他工具函数 ====================

/**
 * val函数 - 字符串及布尔值转为数值（与VBA的val保持一致）
 * @param {String} s - 要转换的字符串
 * @returns {Number} 数值
 * @example
 * val('5')      // 5
 * val('123abc') // 123
 * val('abc123') // 0
 */
const val = (s) => {
    if (typeof s === 'number') return s;
    if (typeof s === 'boolean') return s ? 1 : 0;
    if (typeof s !== 'string') return 0;
    s = s.trim();
    if (s === '') return 0;
    // VBA的val行为：读取字符串开头的数字字符
    const match = s.match(/^[-+]?[0-9]*\.?[0-9]+/);
    if (match) return parseFloat(match[0]);
    return 0;
};

/**
 * round函数 - 使用Excel计算规则对数字进行四舍五入
 * @param {number} number - 要进行四舍五入的数字
 * @param {number} [decimals=2] - 保留的小数位数（默认为2）
 * @returns {number} 四舍五入后的结果
 * @example
 * round(5.786543224, 3)  // 5.787
 */
const round = (number, decimals = 2) => {
    // 使用Excel的RoundWorksheetFunction确保与Excel行为一致
    if (isWPS && typeof WorksheetFunction.Round !== 'undefined') {
        try {
            return WorksheetFunction.Round(number, decimals);
        } catch (e) {
            // 降级处理
        }
    }
    // 标准四舍五入
    const factor = Math.pow(10, decimals);
    return Math.round(number * factor) / factor;
};

// ubound函数 - 获取数组的指定维度的上界
// 在导出部分定义以避免WPS打印函数定义

// ==================== Range快捷函数 ====================

/**
 * RangeChain - Range链式调用包装类（支持智能提示和链式调用）
 * @private
 * @class
 * @constructor
 * @description 支持Range方法的链式调用和智能提示
 * @example
 * $.maxRange("A1:J1").safeArray()     // 链式调用
 * $(5, 2).z值()                       // 获取第5行第2列的值
 * $(5, 2).z值("新值").z加粗()         // 链式设置
 */
function RangeChain(rng, colIndex) {
    if (!(this instanceof RangeChain)) {
        return new RangeChain(rng, colIndex);
    }
    this._range = null;

    // 两个参数模式：RangeChain(行号, 列号)
    if (typeof rng === 'number' && typeof colIndex === 'number') {
        this._range = isWPS ? Cells(rng, colIndex) : null;
    }
    // 字符串地址模式
    else if (typeof rng === 'string') {
        this._range = isWPS ? Range(rng) : null;
    }
    // Range对象模式
    else if (rng && rng.Address) {
        this._range = rng;
    }
}

/**
 * Value - 获取原始Range对象
 * @returns {Range|null} Range对象
 */
RangeChain.prototype.Value = function() {
    return this._range;
};

/**
 * Value2 - 获取/设置值（Value2属性）
 * @param {any} [newValue] - 新值（可选）
 * @returns {RangeChain|any} 设置时返回this，否则返回当前值
 */
RangeChain.prototype.z值 = function(newValue) {
    if (newValue !== undefined) {
        if (this._range) this._range.Value2 = newValue;
        return this;
    }
    return this._range ? this._range.Value2 : undefined;
};
RangeChain.prototype.Value2 = RangeChain.prototype.z值;

/**
 * CurrentRegion - 获取当前区域（连续数据区域）
 * @returns {RangeChain|null} 当前区域的RangeChain对象
 */
RangeChain.prototype.z当前区域 = function() {
    if (!this._range) return null;
    try {
        return new RangeChain(this._range.CurrentRegion);
    } catch (e) {
        return null;
    }
};
RangeChain.prototype.CurrentRegion = RangeChain.prototype.z当前区域;

/**
 * safeArray - 转换为安全数组（返回 Array2D 对象，支持链式调用）
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 */
RangeChain.prototype.z安全数组 = function() {
    return RngUtils.z安全数组(this._range);
};
RangeChain.prototype.safeArray = RangeChain.prototype.z安全数组;

/**
 * MaxArray - 获取最大行数组
 * @param {number} [cols] - 列号
 * @returns {Array} 二维数组
 */
RangeChain.prototype.z最大行数组 = function(cols) {
    return RngUtils.z最大行数组(this._range, cols);
};
RangeChain.prototype.MaxArray = RangeChain.prototype.z最大行数组;

/**
 * VisibleArray - 转换可见区域为数组
 * @param {Worksheet} [tempSheet] - 临时工作表
 * @returns {Array} 数组
 */
RangeChain.prototype.z可见区数组 = function(tempSheet) {
    return RngUtils.z可见区数组(this._range, tempSheet);
};
RangeChain.prototype.VisibleArray = RangeChain.prototype.z可见区数组;

/**
 * RowsCount - 获取行数
 * @returns {number} 行数
 */
RangeChain.prototype.z行数 = function() {
    return this._range ? this._range.Rows.Count : 0;
};
RangeChain.prototype.RowsCount = RangeChain.prototype.z行数;

/**
 * ColsCount - 获取列数
 * @returns {number} 列数
 */
RangeChain.prototype.z列数 = function() {
    return this._range ? this._range.Columns.Count : 0;
};
RangeChain.prototype.ColsCount = RangeChain.prototype.z列数;

/**
 * Columns - 获取列集合
 * @returns {Range} Range对象的Columns属性
 */
Object.defineProperty(RangeChain.prototype, 'Columns', {
    get: function() {
        return this._range ? this._range.Columns : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Rows - 获取行集合
 * @returns {Range} Range对象的Rows属性
 */
Object.defineProperty(RangeChain.prototype, 'Rows', {
    get: function() {
        return this._range ? this._range.Rows : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Font - 获取字体对象
 * @returns {Font} Font对象
 */
Object.defineProperty(RangeChain.prototype, 'Font', {
    get: function() {
        return this._range ? this._range.Font : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Interior - 获取内部对象（背景色等）
 * @returns {Interior} Interior对象
 */
Object.defineProperty(RangeChain.prototype, 'Interior', {
    get: function() {
        return this._range ? this._range.Interior : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Address - 获取地址
 * @returns {string} 地址
 */
RangeChain.prototype.z地址 = function() {
    return this._range ? this._range.Address() : '';
};
RangeChain.prototype.Address = RangeChain.prototype.z地址;

/**
 * AddBorders - 添加边框
 * @param {number} [lineStyle=1] - 线条样式
 * @param {number} [weight=2] - 线条粗细
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z加边框 = function(lineStyle, weight) {
    if (this._range) {
        RngUtils.z加边框(this._range, lineStyle, weight);
    }
    return this;
};
RangeChain.prototype.AddBorders = RangeChain.prototype.z加边框;

/**
 * AutoFitColumns - 自动列宽
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z自动列宽 = function() {
    if (this._range) {
        this._range.Columns.AutoFit();
    }
    return this;
};
RangeChain.prototype.AutoFitColumns = RangeChain.prototype.z自动列宽;

/**
 * AutoFitRows - 自动行高
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z自动行高 = function() {
    if (this._range) {
        this._range.Rows.AutoFit();
    }
    return this;
};
RangeChain.prototype.AutoFitRows = RangeChain.prototype.z自动行高;

/**
 * ClearContents - 清除内容
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z清除内容 = function() {
    if (this._range) {
        this._range.ClearContents();
    }
    return this;
};
RangeChain.prototype.ClearContents = RangeChain.prototype.z清除内容;

/**
 * ClearFormats - 清除格式
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z清除格式 = function() {
    if (this._range) {
        this._range.ClearFormats();
    }
    return this;
};
RangeChain.prototype.ClearFormats = RangeChain.prototype.z清除格式;

/**
 * Value2 - 获取/设置值（Value2属性，比Value更快）
 * @param {any} [newValue] - 新值（可选）
 * @returns {RangeChain|any} 设置时返回this，否则返回当前值
 * @example
 * $(5, 2).z值()                    // 获取值
 * $(5, 2).z值("新值")              // 设置值
 * $(5, 2).z值("新值").z加粗()      // 链式调用
 */
// 注意：z值 方法已在第5734行定义，此处删除重复定义以避免覆盖

// 使用属性方式定义 Value2，支持 $(i,2).Value2 = rs 语法
Object.defineProperty(RangeChain.prototype, 'Value2', {
    get: function() {
        return this._range ? this._range.Value2 : undefined;
    },
    set: function(newValue) {
        if (this._range) this._range.Value2 = newValue;
    },
    enumerable: true,
    configurable: true
});

/**
 * Formula - 获取/设置公式
 * @param {string} [newFormula] - 新公式（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回公式
 */
RangeChain.prototype.z公式 = function(newFormula) {
    if (newFormula !== undefined) {
        if (this._range) this._range.Formula = newFormula;
        return this;
    }
    return this._range ? this._range.Formula : '';
};

// 使用属性方式定义 Formula
Object.defineProperty(RangeChain.prototype, 'Formula', {
    get: function() {
        return this._range ? this._range.Formula : '';
    },
    set: function(newFormula) {
        if (this._range) this._range.Formula = newFormula;
    },
    enumerable: true,
    configurable: true
});

/**
 * Text - 获取显示文本
 * @returns {string} 显示文本
 */
RangeChain.prototype.z文本 = function() {
    return this._range ? this._range.Text : '';
};

// 使用属性方式定义 Text（只读）
Object.defineProperty(RangeChain.prototype, 'Text', {
    get: function() {
        return this._range ? this._range.Text : '';
    },
    enumerable: true,
    configurable: true
});

/**
 * Row - 获取行号
 * @returns {number} 行号
 */
RangeChain.prototype.z行 = function() {
    return this._range ? this._range.Row : 0;
};

// 使用属性方式定义 Row（只读）
Object.defineProperty(RangeChain.prototype, 'Row', {
    get: function() {
        return this._range ? this._range.Row : 0;
    },
    enumerable: true,
    configurable: true
});

/**
 * Column - 获取列号
 * @returns {number} 列号
 */
RangeChain.prototype.z列 = function() {
    return this._range ? this._range.Column : 0;
};

// 使用属性方式定义 Column（只读）
Object.defineProperty(RangeChain.prototype, 'Column', {
    get: function() {
        return this._range ? this._range.Column : 0;
    },
    enumerable: true,
    configurable: true
});

/**
 * Select - 选中区域
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z选中 = function() {
    if (this._range) this._range.Select();
    return this;
};
RangeChain.prototype.Select = RangeChain.prototype.z选中;

/**
 * Activate - 激活单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z激活 = function() {
    if (this._range) this._range.Activate();
    return this;
};
RangeChain.prototype.Activate = RangeChain.prototype.z激活;

/**
 * Bold - 获取/设置加粗
 * @param {boolean} [isBold] - 是否加粗（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回加粗状态
 */
RangeChain.prototype.z加粗 = function(isBold) {
    if (isBold !== undefined) {
        if (this._range) this._range.Font.Bold = isBold;
        return this;
    }
    return this._range ? this._range.Font.Bold : false;
};

// 使用属性方式定义 Bold
Object.defineProperty(RangeChain.prototype, 'Bold', {
    get: function() {
        return this._range ? this._range.Font.Bold : false;
    },
    set: function(isBold) {
        if (this._range) this._range.Font.Bold = isBold;
    },
    enumerable: true,
    configurable: true
});

/**
 * Italic - 获取/设置斜体
 * @param {boolean} [isItalic] - 是否斜体（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回斜体状态
 */
RangeChain.prototype.z斜体 = function(isItalic) {
    if (isItalic !== undefined) {
        if (this._range) this._range.Font.Italic = isItalic;
        return this;
    }
    return this._range ? this._range.Font.Italic : false;
};

// 使用属性方式定义 Italic
Object.defineProperty(RangeChain.prototype, 'Italic', {
    get: function() {
        return this._range ? this._range.Font.Italic : false;
    },
    set: function(isItalic) {
        if (this._range) this._range.Font.Italic = isItalic;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontColor - 获取/设置字体颜色
 * @param {number} [color] - RGB颜色值（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回颜色值
 */
RangeChain.prototype.z字体颜色 = function(color) {
    if (color !== undefined) {
        if (this._range) this._range.Font.Color = color;
        return this;
    }
    return this._range ? this._range.Font.Color : 0;
};

// 使用属性方式定义 FontColor
Object.defineProperty(RangeChain.prototype, 'FontColor', {
    get: function() {
        return this._range ? this._range.Font.Color : 0;
    },
    set: function(color) {
        if (this._range) this._range.Font.Color = color;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontSize - 获取/设置字体大小
 * @param {number} [size] - 字体大小（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回字体大小
 */
RangeChain.prototype.z字号 = function(size) {
    if (size !== undefined) {
        if (this._range) this._range.Font.Size = size;
        return this;
    }
    return this._range ? this._range.Font.Size : 11;
};

// 使用属性方式定义 FontSize
Object.defineProperty(RangeChain.prototype, 'FontSize', {
    get: function() {
        return this._range ? this._range.Font.Size : 11;
    },
    set: function(size) {
        if (this._range) this._range.Font.Size = size;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontName - 获取/设置字体名称
 * @param {string} [fontName] - 字体名称（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回字体名称
 */
RangeChain.prototype.z字体名称 = function(fontName) {
    if (fontName !== undefined) {
        if (this._range) this._range.Font.Name = fontName;
        return this;
    }
    return this._range ? this._range.Font.Name : '';
};

// 使用属性方式定义 FontName
Object.defineProperty(RangeChain.prototype, 'FontName', {
    get: function() {
        return this._range ? this._range.Font.Name : '';
    },
    set: function(fontName) {
        if (this._range) this._range.Font.Name = fontName;
    },
    enumerable: true,
    configurable: true
});

/**
 * InteriorColor - 获取/设置背景颜色
 * @param {number} [color] - RGB颜色值（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回颜色值
 */
RangeChain.prototype.z背景颜色 = function(color) {
    if (color !== undefined) {
        if (this._range) this._range.Interior.Color = color;
        return this;
    }
    return this._range ? this._range.Interior.Color : 16777215; // 默认白色
};

// 使用属性方式定义 InteriorColor
Object.defineProperty(RangeChain.prototype, 'InteriorColor', {
    get: function() {
        return this._range ? this._range.Interior.Color : 16777215;
    },
    set: function(color) {
        if (this._range) this._range.Interior.Color = color;
    },
    enumerable: true,
    configurable: true
});

/**
 * HorizontalAlignment - 获取/设置水平对齐
 * @param {number} [align] - 对齐方式（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回对齐方式
 */
RangeChain.prototype.z水平对齐 = function(align) {
    if (align !== undefined) {
        if (this._range) this._range.HorizontalAlignment = align;
        return this;
    }
    return this._range ? this._range.HorizontalAlignment : -4151; // 默认常规
};

// 使用属性方式定义 HorizontalAlignment
Object.defineProperty(RangeChain.prototype, 'HorizontalAlignment', {
    get: function() {
        return this._range ? this._range.HorizontalAlignment : -4151;
    },
    set: function(align) {
        if (this._range) this._range.HorizontalAlignment = align;
    },
    enumerable: true,
    configurable: true
});

/**
 * VerticalAlignment - 获取/设置垂直对齐
 * @param {number} [align] - 对齐方式（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回对齐方式
 */
RangeChain.prototype.z垂直对齐 = function(align) {
    if (align !== undefined) {
        if (this._range) this._range.VerticalAlignment = align;
        return this;
    }
    return this._range ? this._range.VerticalAlignment : -4160; // 默认底部
};

// 使用属性方式定义 VerticalAlignment
Object.defineProperty(RangeChain.prototype, 'VerticalAlignment', {
    get: function() {
        return this._range ? this._range.VerticalAlignment : -4160;
    },
    set: function(align) {
        if (this._range) this._range.VerticalAlignment = align;
    },
    enumerable: true,
    configurable: true
});

/**
 * NumberFormat - 获取/设置数字格式
 * @param {string} [format] - 格式字符串（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回格式字符串
 */
RangeChain.prototype.z数字格式 = function(format) {
    if (format !== undefined) {
        if (this._range) this._range.NumberFormat = format;
        return this;
    }
    return this._range ? this._range.NumberFormat : 'General';
};

// 使用属性方式定义 NumberFormat
Object.defineProperty(RangeChain.prototype, 'NumberFormat', {
    get: function() {
        return this._range ? this._range.NumberFormat : 'General';
    },
    set: function(format) {
        if (this._range) this._range.NumberFormat = format;
    },
    enumerable: true,
    configurable: true
});

/**
 * WrapText - 获取/设置自动换行
 * @param {boolean} [wrap] - 是否自动换行（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回换行状态
 */
RangeChain.prototype.z自动换行 = function(wrap) {
    if (wrap !== undefined) {
        if (this._range) this._range.WrapText = wrap;
        return this;
    }
    return this._range ? this._range.WrapText : false;
};

// 使用属性方式定义 WrapText
Object.defineProperty(RangeChain.prototype, 'WrapText', {
    get: function() {
        return this._range ? this._range.WrapText : false;
    },
    set: function(wrap) {
        if (this._range) this._range.WrapText = wrap;
    },
    enumerable: true,
    configurable: true
});

/**
 * Merge - 合并单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z合并 = function() {
    if (this._range) this._range.Merge();
    return this;
};
RangeChain.prototype.Merge = RangeChain.prototype.z合并;

/**
 * Clear - 清除内容和格式
 * @returns {RangeChain} 当前实例
 * @example
 * $("K2").Resize(1000, 5000).Clear()
 * $.Resize("K2", 1000, 5000).Clear()
 */
RangeChain.prototype.Clear = function() {
    if (this._range) {
        // WPS JSA 兼容：使用 ClearContents 和 ClearFormats
        try {
            this._range.ClearContents();
        } catch (e) {}
        try {
            this._range.ClearFormats();
        } catch (e) {}
    }
    return this;
};

/**
 * z清除 - Clear的中文别名
 */
RangeChain.prototype.z清除 = RangeChain.prototype.Clear;

/**
 * UnMerge - 取消合并单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z取消合并 = function() {
    if (this._range) this._range.UnMerge();
    return this;
};
RangeChain.prototype.UnMerge = RangeChain.prototype.z取消合并;

/**
 * Resize - 调整区域大小
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @returns {RangeChain} 调整大小后的新RangeChain对象
 * @example
 * $("K2").Resize(10, 5).z清除内容()
 * $("K2").Resize(1000, 5000).z清除内容()
 */
RangeChain.prototype.Resize = function(rows, cols) {
    if (!this._range) return new RangeChain(null);
    try {
        var resizedRng = this._range.Resize(rows, cols);
        return new RangeChain(resizedRng);
    } catch (e) {
        console.log("Resize失败: " + e.message);
        return this;
    }
};

/**
 * MergeCells - 检查是否为合并单元格
 * @returns {boolean} 是否合并
 */
RangeChain.prototype.z已合并 = function() {
    return this._range ? this._range.MergeCells : false;
};

// 使用属性方式定义 MergeCells（只读）
Object.defineProperty(RangeChain.prototype, 'MergeCells', {
    get: function() {
        return this._range ? this._range.MergeCells : false;
    },
    enumerable: true,
    configurable: true
});

/**
 * Delete - 删除区域
 * @param {number} [shift] - 移动方向（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z删除 = function(shift) {
    if (this._range) this._range.Delete(shift);
    return this;
};
RangeChain.prototype.Delete = RangeChain.prototype.z删除;

/**
 * Insert - 插入区域
 * @param {number} [shift] - 移动方向（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z插入 = function(shift) {
    if (this._range) this._range.Insert(shift);
    return this;
};
RangeChain.prototype.Insert = RangeChain.prototype.z插入;

/**
 * Copy - 复制区域
 * @param {Range} [destination] - 目标区域（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z复制 = function(destination) {
    if (this._range) {
        if (destination) {
            this._range.Copy(destination);
        } else {
            this._range.Copy();
        }
    }
    return this;
};
RangeChain.prototype.Copy = RangeChain.prototype.z复制;

/**
 * Paste - 粘贴区域
 * @param {Range} [destination] - 目标区域（可选）
 * @param {number} [type] - 粘贴类型（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z粘贴 = function(destination, type) {
    if (destination && destination.Paste) {
        destination.Paste(type);
    }
    return this;
};
RangeChain.prototype.Paste = RangeChain.prototype.z粘贴;

/**
 * 创建RngUtils静态方法代理对象
 * @private
 */
function createRngUtilsProxy() {
    var proxy = {};
    var staticMethods = [
        'z最后一个', 'lastCell',
        'z安全区域', 'safeRange',
        'z安全数组', 'safeArray',
        'z最大行', 'endRow',
        'z最大行单元格', 'endRowCell',
        'z最大行区域', 'maxRange',
        'z最大列', 'endCol',
        'z最大列单元格', 'endColCell',
        'z可见区数组', 'visibleArray',
        'z可见区域', 'visibleRange',
        'z加边框', 'addBorders',
        'z取前几行', 'takeRows',
        'z跳过前几行', 'skipRows',
        'z合并相同单元格', 'mergeCells',
        'z取消合并填充单元格', 'unMergeCells',
        'z插入多行', 'insertRows',
        'z插入多列', 'insertCols',
        'z删除空白行', 'delBlankRows',
        'z删除空白列', 'delBlankCols',
        'z整行', 'entireRow',
        'z整列', 'entire_column',
        'z行数', 'rowsCount',
        'z列数', 'colsCount',
        'z列号字母互转', 'colToAbc',
        'z复制粘贴格式', 'copyFormat',
        'z复制粘贴值', 'copyValue',
        'z联合区域', 'unionAll',
        'z多列排序', 'rngSortCols',
        'z强力筛选', 'rngFilter',
        'z最大行数组', 'maxArray',
        'z查找单元格', 'findRange',
        'z命中单元格', 'hitRange',
        'z本文件单元格', 'thisRange',
        'z选择不连续列数组', 'selectColsArray'
    ];

    for (var i = 0; i < staticMethods.length; i++) {
        var methodName = staticMethods[i];
        if (RngUtils[methodName]) {
            (function(name) {
                proxy[name] = function() {
                    var result = RngUtils[name].apply(RngUtils, arguments);
                    // 如果返回的是Range对象，包装成RangeChain支持链式调用
                    if (result && result.Address && typeof result.Address === 'function') {
                        return new RangeChain(result);
                    }
                    return result;
                };
            })(methodName);
        }
    }

    return proxy;
}

/**
 * $函数 - Range快捷方式和RngUtils方法代理（支持智能提示和链式调用）
 * @param {string|number} x - 地址或行号
 * @param {number} [y] - 列号（可选，当传入两个数字参数时）
 * @returns {RangeChain} RangeChain包装对象，支持智能提示和链式调用
 * @example
 * $("A1")                          // 返回RangeChain，支持链式调用
 * $(5, 2)                          // 第5行第2列，返回RangeChain
 * $(5, 2).z值()                    // 获取值
 * $(5, 2).z值("新值")              // 设置值
 * $(5, 2).z值("新值").z加粗()      // 链式调用
 * $.maxRange("A1:J1").safeArray()  // 链式调用
 */
function $(x, y) {
    // 两个参数模式：$(行, 列) - 返回RangeChain
    if (arguments.length === 2 && typeof x === 'number' && typeof y === 'number') {
        return new RangeChain(x, y);
    }
    // 单个参数模式 - 返回RangeChain
    if (typeof x === 'string') {
        return new RangeChain(x);
    } else if (typeof x === 'number') {
        return new RangeChain(x, 1);
    } else if (x && x.Address) {
        return new RangeChain(x);
    }
    // 返回空的RangeChain
    return new RangeChain(null);
}

// ==================== 将 RngUtils 常用静态方法直接添加到 $ 对象上 ====================
// 直接定义以支持智能提示

/**
 * $.maxRange - 获取从第一行到最后一行的区域
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.maxRange("A1:J1").safeArray()
 * $.maxRange("1:1000", "A").z加边框()
 */
$.maxRange = function(rng, col) {
    var result = RngUtils.maxRange.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z最大行区域 - maxRange的中文别名
 */
$.z最大行区域 = $.maxRange;

/**
 * $.safeArray - 将区域转换为安全数组（返回 Array2D 对象，支持链式调用）
 * @param {Range|string} rng - 区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 */
$.safeArray = RngUtils.safeArray;

/**
 * $.z安全数组 - safeArray的中文别名
 */
$.z安全数组 = $.safeArray;

/**
 * $.endRow - 获取区域最大行数
 * @param {Range|string} rng - 区域
 * @returns {number} 行数
 */
$.endRow = RngUtils.endRow;

/**
 * $.z最大行 - endRow的中文别名
 */
$.z最大行 = $.endRow;

/**
 * $.addBorders - 添加边框
 * @param {Range|string} rng - 区域
 * @param {number} [lineStyle=1] - 线条样式
 * @param {number} [weight=2] - 线条粗细
 * @returns {RangeChain} RangeChain对象
 */
$.addBorders = function(rng, lineStyle, weight) {
    var result = RngUtils.addBorders.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z加边框 - addBorders的中文别名
 */
$.z加边框 = $.addBorders;

/**
 * $.autoFitColumns - 自动列宽
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 */
$.autoFitColumns = function(rng) {
    var result = RngUtils.autoFitColumns.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z自动列宽 - autoFitColumns的中文别名
 */
$.z自动列宽 = $.autoFitColumns;

/**
 * $.autoFitRows - 自动行高
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 */
$.autoFitRows = function(rng) {
    var result = RngUtils.autoFitRows.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z自动行高 - autoFitRows的中文别名
 */
$.z自动行高 = $.autoFitRows;

/**
 * $.delBlankRows - 删除空白行
 * @param {Range|string} rng - 区域
 * @param {boolean} [entireColumn=false] - 是否删除整行
 * @returns {RangeChain} RangeChain对象
 */
$.delBlankRows = function(rng, entireColumn) {
    RngUtils.delBlankRows.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.z删除空白行 - delBlankRows的中文别名
 */
$.z删除空白行 = $.delBlankRows;

/**
 * $.delBlankCols - 删除空白列
 * @param {Range|string} rng - 区域
 * @param {boolean} [entireColumn=false] - 是否删除整列
 * @returns {RangeChain} RangeChain对象
 */
$.delBlankCols = function(rng, entireColumn) {
    RngUtils.delBlankCols.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.z删除空白列 - delBlankCols的中文别名
 */
$.z删除空白列 = $.delBlankCols;

/**
 * $.rngSortCols - 多列排序
 * @param {Range|string} rng - 区域
 * @param {Array} sortCols - 排序列数组
 * @returns {RangeChain} RangeChain对象
 */
$.rngSortCols = function(rng, sortCols) {
    RngUtils.rngSortCols.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.maxArray - 根据指定单元格区域获取最大行数组
 * @param {Range|string} rng - 单元格区域
 * @param {number} [cols] - 选择列作为获取最大行依据
 * @returns {Array} 结果二维数组
 * @example
 * $.maxArray("a:d")              // 获取A:D列的最大行数组
 * $.maxArray("a:d", 2)            // 使用第2列确定最大行
 */
$.maxArray = RngUtils.maxArray;

/**
 * $.z最大行数组 - maxArray的中文别名
 */
$.z最大行数组 = $.maxArray;

/**
 * $.z多列排序 - rngSortCols的中文别名
 */
$.z多列排序 = $.rngSortCols;

/**
 * $.rngFilter - 强力筛选
 * @param {Range|string} rng - 区域
 * @param {Object} conditions - 筛选条件
 * @returns {Array} 筛选后的数组
 */
$.rngFilter = RngUtils.rngFilter;

/**
 * $.z强力筛选 - rngFilter的中文别名
 */
$.z强力筛选 = $.rngFilter;

/**
 * $.colToAbc - 列号与字母互转
 * @param {number|string} input - 列号或字母
 * @returns {string|number} 字母或列号
 */
$.colToAbc = RngUtils.colToAbc;

/**
 * $.z列号字母互转 - colToAbc的中文别名
 */
$.z列号字母互转 = $.colToAbc;

/**
 * $.rowsCount - 获取行数
 * @param {Range|string} rng - 区域
 * @returns {number} 行数
 */
$.rowsCount = RngUtils.rowsCount;

/**
 * $.z行数 - rowsCount的中文别名
 */
$.z行数 = $.rowsCount;

/**
 * $.colsCount - 获取列数
 * @param {Range|string} rng - 区域
 * @returns {number} 列数
 */
$.colsCount = RngUtils.colsCount;

/**
 * $.z列数 - colsCount的中文别名
 */
$.z列数 = $.colsCount;

/**
 * $.copyValue - 复制粘贴值
 * @param {Range|string} fromRng - 源区域
 * @param {Range|string} toRng - 目标区域
 * @returns {RangeChain} RangeChain对象
 */
$.copyValue = function(fromRng, toRng) {
    RngUtils.copyValue.apply(RngUtils, arguments);
    return new RangeChain(toRng);
};

/**
 * $.z复制粘贴值 - copyValue的中文别名
 */
$.z复制粘贴值 = $.copyValue;

/**
 * $.copyFormat - 复制粘贴格式
 * @param {Range|string} fromRng - 源区域
 * @param {Range|string} toRng - 目标区域
 * @returns {RangeChain} RangeChain对象
 */
$.copyFormat = function(fromRng, toRng) {
    RngUtils.copyFormat.apply(RngUtils, arguments);
    return new RangeChain(toRng);
};

/**
 * $.z复制粘贴格式 - copyFormat的中文别名
 */
$.z复制粘贴格式 = $.copyFormat;

/**
 * $.Resize - 调整区域大小（静态方法）
 * @param {Range|string} rng - 源区域
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @returns {RangeChain} 调整大小后的 RangeChain 对象
 * @example
 * $.Resize("K2", 1000, 5000).z清除内容()
 * $.Resize(Range("A1"), 10, 5).z加边框()
 */
$.Resize = function(rng, rows, cols) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (!targetRng) return new RangeChain(null);
    try {
        var resizedRng = targetRng.Resize(rows, cols);
        return new RangeChain(resizedRng);
    } catch (e) {
        console.log("Resize失败: " + e.message);
        return new RangeChain(rng);
    }
};

/**
 * $.z调整大小 - Resize的中文别名
 */
$.z调整大小 = $.Resize;

/**
 * $.ClearContents - 清除内容（静态方法）
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.ClearContents("K2").Resize(1000, 5000).z清除内容()
 */
$.ClearContents = function(rng) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (targetRng) {
        targetRng.ClearContents();
    }
    return new RangeChain(targetRng);
};

/**
 * $.z清除内容 - ClearContents的中文别名
 */
$.z清除内容 = $.ClearContents;

/**
 * $.UnMerge - 取消合并（静态方法）
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.UnMerge("K2").Resize(1000, 1000).z取消合并()
 */
$.UnMerge = function(rng) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (targetRng) {
        targetRng.UnMerge();
    }
    return new RangeChain(targetRng);
};

/**
 * $.z取消合并 - UnMerge的中文别名
 */
$.z取消合并 = $.UnMerge;

// ==================== 将 RngUtils 常用静态方法直接添加到 $ 对象 ====================
// 支持直接调用 $.mergeCells() 而不是 $.RngUtils.mergeCells()

// 定义需要直接添加到 $ 的常用方法
var directMethods = [
    'z合并相同单元格', 'mergeCells',
    'z取消合并填充单元格', 'unMergeCells',
    'z加边框', 'addBorders',
    'z插入多行', 'insertRows',
    'z插入多列', 'insertCols',
    'z删除空白行', 'delBlankRows',
    'z删除空白列', 'delBlankCols'
];

for (var i = 0; i < directMethods.length; i++) {
    var methodName = directMethods[i];
    if (RngUtils[methodName]) {
        (function(name) {
            $[name] = function() {
                return RngUtils[name].apply(RngUtils, arguments);
            };
        })(methodName);
    }
}

// ==================== 将构造函数类工厂添加到$对象 ====================

/**
 * $.Array2D - 二维数组工具类工厂（支持智能提示和链式调用）
 * @param {Array} data - 输入数据
 * @returns {Array2D} Array2D实例，支持链式调用和智能提示
 * @example
 * $.Array2D([[1,2],[3,4]]).z求和()      // 10
 * $.Array2D([1,2,3]).z转置()           // [[1],[2],[3]]
 * $.Array2D([[1,2],[3,4]]).toRange("A1")  // 写入A1:B2
 */
$.Array2D = function(data) {
    return new Array2D(data);
};

/**
 * $.RngUtils - Range工具类工厂
 * @param {string|Range} [initialRange] - 初始Range（可选）
 * @returns {RngUtils|Object} RngUtils实例或静态方法对象
 * @example
 * $.RngUtils("A1:B10").z安全数组()    // 实例方法
 * $.RngUtils.maxRange("A1:J1")        // 静态方法
 */
$.RngUtils = function(initialRange) {
    // 无参数调用时，返回静态方法代理对象
    if (arguments.length === 0) {
        return createRngUtilsStaticProxy();
    }
    return new RngUtils(initialRange);
};

// ==================== 将 RngUtils 静态方法添加到 $.RngUtils 上 ====================
// 支持智能提示和 $.RngUtils.maxRange() 调用

/**
 * $.RngUtils.maxRange - 获取从第一行到最后一行的区域
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {RangeChain} RangeChain对象，支持链式调用
 * @memberof $.RngUtils
 * @example
 * $.RngUtils.maxRange("1:1000","A").safeArray()  // 返回数组
 * $.RngUtils.maxRange("A1:J1").z加边框()         // 链式调用
 */
$.RngUtils.maxRange = function(rng, col) {
    var result = RngUtils.maxRange.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.RngUtils.safeArray - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @static
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @memberof $.RngUtils
 */
$.RngUtils.safeArray = RngUtils.safeArray;

/**
 * $.RngUtils.z安全数组 - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @static
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @memberof $.RngUtils
 */
$.RngUtils.z安全数组 = RngUtils.z安全数组;

/**
 * $.RngUtils.endRow - 获取指定区域的最大行数
 * @static
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @memberof $.RngUtils
 */
$.RngUtils.endRow = RngUtils.endRow;

/**
 * $.RngUtils.z最大行 - 获取指定区域的最大行数
 * @static
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @memberof $.RngUtils
 */
$.RngUtils.z最大行 = RngUtils.z最大行;

// 其他常用静态方法（可根据需要添加更多）
var staticMethods = [
    'z最后一个', 'lastCell',
    'z安全区域', 'safeRange',
    'z最大行单元格', 'endRowCell',
    'z最大行区域',
    'z最大列', 'endCol',
    'z最大列单元格', 'endColCell',
    'z可见区数组', 'visibleArray',
    'z可见区域', 'visibleRange',
    'z加边框', 'addBorders',
    'z取前几行', 'takeRows',
    'z跳过前几行', 'skipRows',
    'z合并相同单元格', 'mergeCells',
    'z取消合并填充单元格', 'unMergeCells',
    'z插入多行', 'insertRows',
    'z插入多列', 'insertCols',
    'z删除空白行', 'delBlankRows',
    'z删除空白列', 'delBlankCols',
    'z整行', 'entireRow',
    'z整列', 'entire_column',
    'z行数', 'rowsCount',
    'z列数', 'colsCount',
    'z列号字母互转', 'colToAbc',
    'z复制粘贴格式', 'copyFormat',
    'z复制粘贴值', 'copyValue',
    'z联合区域', 'unionAll',
    'z多列排序', 'rngSortCols',
    'z强力筛选', 'rngFilter',
    'z最大行数组', 'maxArray',
    'z查找单元格', 'findRange',
    'z命中单元格', 'hitRange',
    'z本文件单元格', 'thisRange',
    'z选择不连续列数组', 'selectColsArray'
];

for (var i = 0; i < staticMethods.length; i++) {
    var methodName = staticMethods[i];
    if (RngUtils[methodName]) {
        (function(name) {
            $.RngUtils[name] = function() {
                var result = RngUtils[name].apply(RngUtils, arguments);
                // 如果返回的是Range对象，包装成RangeChain支持链式调用
                if (result && result.Address && typeof result.Address === 'function') {
                    return new RangeChain(result);
                }
                return result;
            };
        })(methodName);
    }
}

/**
 * $.ShtUtils - Sheet工具类工厂
 * @param {Worksheet} initialSheet - 初始Sheet
 * @returns {ShtUtils} ShtUtils实例
 * @example
 * $.ShtUtils().z当前工作表()
 */
$.ShtUtils = function(initialSheet) {
    return new ShtUtils(initialSheet);
};

/**
 * $.DateUtils - 日期工具类工厂
 * @param {Date|string} initialDate - 初始日期
 * @returns {DateUtils} DateUtils实例
 * @example
 * $.DateUtils().z格式化("yyyy-MM-dd")
 */
$.DateUtils = function(initialDate) {
    return new DateUtils(initialDate);
};

// ==================== 全局变量导出 ====================

// Node.js环境
if (isNodeJS) {
    module.exports.Array2D = Array2D;
    module.exports.As = As;
    module.exports.RngUtils = RngUtils;
    module.exports.ShtUtils = ShtUtils;
    module.exports.DateUtils = DateUtils;
    module.exports.JSA = JSA;
    module.exports.IO = IO;
    module.exports.$ = $;
    module.exports.log = log;
    module.exports.logjson = logjson;
    // Global函数
    module.exports.f1 = f1;
    module.exports.$fx = $fx;
    module.exports.$toArray = $toArray;
    module.exports.asArray = asArray;
    module.exports.asDate = asDate;
    module.exports.asMap = asMap;
    module.exports.asNumber = asNumber;
    module.exports.asObject = asObject;
    module.exports.asRange = asRange;
    module.exports.asShape = asShape;
    module.exports.asSheet = asSheet;
    module.exports.asString = asString;
    module.exports.asWorkbook = asWorkbook;
    module.exports.cdate = cdate;
    module.exports.cstr = cstr;
    module.exports.isArray = isArray;
    module.exports.isArray2D = isArray2D;
    module.exports.isBoolean = isBoolean;
    module.exports.isCollection = isCollection;
    module.exports.isDate = isDate;
    module.exports.isEmpty = isEmpty;
    module.exports.isNumberic = isNumberic;
    module.exports.isRange = isRange;
    module.exports.isRegex = isRegex;
    module.exports.isSameClass = isSameClass;
    module.exports.isSheet = isSheet;
    module.exports.isString = isString;
    module.exports.isWorkbook = isWorkbook;
    module.exports.typeName = typeName;
    module.exports.val = val;
    module.exports.round = round;
    // ubound函数 - 获取数组的指定维度的上界
    module.exports.ubound = function(arr, dimension) {
        dimension = dimension || 1;
        if (!Array.isArray(arr)) return -1;
        if (dimension === 1) return arr.length - 1;
        if (dimension === 2) {
            var maxLen = 0;
            for (var i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i]) && arr[i].length > maxLen) {
                    maxLen = arr[i].length;
                }
            }
            return maxLen - 1;
        }
        return -1;
    };
}

// WPS/Browser环境 - 使用立即执行函数避免WPS打印函数定义
if (isWPS || isBrowser) {
    (function() {
        this.Array2D = Array2D;
        this.As = As;
        this.RngUtils = RngUtils;
        this.ShtUtils = ShtUtils;
        this.DateUtils = DateUtils;
        this.JSA = JSA;
        this.IO = IO;
        this.$ = $;
        this.log = log;
        this.logjson = logjson;
        // Global函数
        this.f1 = f1;
        this.$fx = $fx;
        this.$toArray = $toArray;
        // As 已在第6891行导出，此处删除重复定义
        this.asArray = asArray;
        this.asDate = asDate;
        this.asMap = asMap;
        this.asNumber = asNumber;
        this.asObject = asObject;
        this.asRange = asRange;
        this.asShape = asShape;
        this.asSheet = asSheet;
        this.asString = asString;
        this.asWorkbook = asWorkbook;
        this.cdate = cdate;
        this.cstr = cstr;
        this.isArray = isArray;
        this.isArray2D = isArray2D;
        this.isBoolean = isBoolean;
        this.isCollection = isCollection;
        this.isDate = isDate;
        this.isEmpty = isEmpty;
        this.isNumberic = isNumberic;
        this.isRange = isRange;
        this.isRegex = isRegex;
        this.isSameClass = isSameClass;
        this.isSheet = isSheet;
        this.isString = isString;
        this.isWorkbook = isWorkbook;
        this.typeName = typeName;
        this.val = val;
        this.round = round;
        // ubound函数 - 获取数组的指定维度的上界
        this.ubound = function(arr, dimension) {
            dimension = dimension || 1;
            if (!Array.isArray(arr)) return -1;
            if (dimension === 1) return arr.length - 1;
            if (dimension === 2) {
                var maxLen = 0;
                for (var i = 0; i < arr.length; i++) {
                    if (Array.isArray(arr[i]) && arr[i].length > maxLen) {
                        maxLen = arr[i].length;
                    }
                }
                return maxLen - 1;
            }
            return -1;
        };
        
        // ==================== JSA880快捷API - 一行代码走天下 ====================
        
        /**
         * JSA880 - 郑广学JSA880快速开发框架主入口
         * @description 提供超简洁的一行代码API，集成所有核心功能
         * @namespace
         * @example
         * // 一行代码完成数据透视
         * JSA880.透视(数据, '产品+,月份+', '地区+', 'sum(销量),count()');
         * 
         * // 一行代码筛选数据
         * JSA880.筛选(数据, 'x=>x[0]=="北京" && x[3]>100');
         * 
         * // 一行代码读取表格数据
         * JSA880.读表("A1:D100");
         * 
         * // 一行代码写入表格数据
         * JSA880.写表([[1,2],[3,4]], "G1");
         * 
         * // 一行代码获取最大行数
         * JSA880.最大行("A:A");
         * 
         * // 一行代码删除空白行
         * JSA880.删空行("A1:F100");
         * 
         * // 一行代码排序
         * JSA880.排序(数据, 'f3+,f4-');
         */ 
        this.JSA880 = {
            /**
             * 数据透视（超简化版）
             * @param {Array} data - 二维数组数据
             * @param {string} rowFields - 行字段，支持排序符号 f1+,f2-
             * @param {string} colFields - 列字段，支持排序符号 f3+,f4-
             * @param {string} dataFields - 数据字段，格式: 'count(),sum(f5),average(f6)'
             * @param {number} [headerRows=1] - 标题行数
             * @returns {Array} 透视结果
             * @example
             * JSA880.透视(销售数据, '产品+,地区-', '月份+', 'sum(金额),count()');
             */
            透视: function(data, rowFields, colFields, dataFields, headerRows) {
                return Array2D.z超级透视(data, [rowFields], [colFields], [dataFields], headerRows);
            },
            
            /**
             * 超级透视（完整版）
             */
            超级透视: Array2D.z超级透视,
            
            /**
             * 数据筛选
             * @param {Array} data - 二维数组
             * @param {string|Function} predicate - 筛选条件
             * @returns {Array2D} Array2D对象
             * @example
             * JSA880.筛选(数据, 'x=>x[0]=="北京" && x[3]>100');
             */
            筛选: function(data, predicate) {
                return new Array2D(data).z筛选(predicate);
            },
            
            /**
             * 多条件筛选（简化版）
             * @param {Array} data - 二维数组
             * @param {Array} conditions - 条件数组，如 [[0, '北京'], [3, 100]]
             * @returns {Array2D} Array2D对象
             * @example
             * JSA880.多条件筛选(数据, [[0, '北京'], [3, 100]]);
             */
            多条件筛选: function(data, conditions) {
                var arr = new Array2D(data);
                for (var i = 0; i < conditions.length; i++) {
                    var col = conditions[i][0];
                    var val = conditions[i][1];
                    arr = arr.z筛选(function(row) { 
                        return row[col] == val || (typeof val === 'number' && row[col] > val);
                    });
                }
                return arr;
            },
            
            /**
             * 分组汇总
             * @param {Array} data - 二维数组
             * @param {string} groupCol - 分组列 f1
             * @param {string} aggCol - 汇总列 f2
             * @param {string} [aggType='sum'] - 汇总类型: sum, count, average, max, min
             * @returns {Array} 汇总结果
             * @example
             * JSA880.分组汇总(数据, 'f1', 'f3', 'sum');
             */
            分组汇总: function(data, groupCol, aggCol, aggType) {
                aggType = aggType || 'sum';
                var aggExpr = aggType + '("' + aggCol + '")';
                return Array2D.z分组汇总(data, groupCol, aggExpr);
            },
            
            /**
             * 分组汇总连接 - 优化sumifs和Countifs批量条件统计
             * @param {Array} targetData - 统计目标数据（左表）
             * @param {Array} sourceData - 数据源（右表）
             * @param {string} groupKey - 分组键选择器，如 'f2' 或 'f2,f3'
             * @param {string} aggFunc - 汇总函数，如 'sum("f4")' 或 'count(),sum("f5")'
             * @returns {Array} 连接汇总后的结果
             * @example
             * // 一行代码完成sumifs/countifs批量统计
             * JSA880.分组汇总连接(目标表, 源数据, 'f2', 'sum("f4")');
             * JSA880.分组汇总连接(目标表, 源数据, '月份,产品', 'count(),sum("销量"),average("金额")');
             */
            分组汇总连接: function(targetData, sourceData, groupKey, aggFunc) {
                return Array2D.groupIntoJoin(targetData, sourceData, groupKey, aggFunc);
            },
            
            /**
             * 读取表格数据（简化版）
             * @param {string} range - 单元格地址，如 "A1:D100" 或 "A:A"
             * @returns {Array} 二维数组
             * @example
             * JSA880.读表("A1:D100");
             * JSA880.读表("A:A");  // 读取整列到最大行
             */
            读表: function(range) {
                if (!isWPS) return [];
                var rng = typeof range === 'string' ? Range(range) : range;
                var arr = rng.Value2;
                if (arr === null || arr === undefined) return [];
                if (!Array.isArray(arr)) return [[arr]];
                if (!Array.isArray(arr[0])) {
                    var result = [];
                    for (var i = 0; i < arr.length; i++) {
                        result.push([arr[i]]);
                    }
                    return result;
                }
                return arr;
            },
            
            /**
             * 写入表格数据（简化版）
             * @param {Array} data - 二维数组
             * @param {string} startCell - 起始单元格，如 "A1"
             * @returns {Range} 写入的单元格区域
             * @example
             * JSA880.写表([[1,2],[3,4]], "G1");
             */
            写表: function(data, startCell) {
                return JSA.z写入单元格(data, startCell);
            },
            
            /**
             * 获取最大行数
             * @param {string} column - 列范围，如 "A:A" 或 "A1"
             * @returns {number} 最大行数
             * @example
             * JSA880.最大行("A:A");
             */
            最大行: function(column) {
                return RngUtils.z最大行(column);
            },
            
            /**
             * 获取最大列数
             * @param {string} row - 行范围，如 "1:1" 或 "A1"
             * @returns {number} 最大列数
             * @example
             * JSA880.最大列("1:1");
             */
            最大列: function(row) {
                return RngUtils.z最大列(row);
            },
            
            /**
             * 删除空白行
             * @param {string} range - 单元格范围
             * @param {boolean} [entireRow=true] - 是否删除整行
             * @returns {boolean} 是否成功
             * @example
             * JSA880.删空行("A1:F100");
             */
            删空行: function(range, entireRow) {
                RngUtils.z删除空白行(range, entireRow !== false);
                return true;
            },
            
            /**
             * 删除空白列
             * @param {string} range - 单元格范围
             * @param {boolean} [entireColumn=true] - 是否删除整列
             * @returns {boolean} 是否成功
             * @example
             * JSA880.删空列("A1:Z100");
             */
            删空列: function(range, entireColumn) {
                RngUtils.z删除空白列(range, entireColumn !== false);
                return true;
            },
            
            /**
             * 多列排序（简化版）
             * @param {Array} data - 二维数组
             * @param {string} sortParams - 排序参数，如 'f3+,f4-'
             * @param {number} [headerRows=1] - 标题行数
             * @returns {Array} 排序后数组
             * @example
             * JSA880.排序(数据, 'f3+,f4-', 1);
             */
            排序: function(data, sortParams, headerRows) {
                return new Array2D(data).z多列排序(sortParams, headerRows || 1);
            },
            
            /**
             * 去重
             * @param {Array} data - 二维数组
             * @param {number} [colIndex] - 指定列去重
             * @returns {Array} 去重后数组
             * @example
             * JSA880.去重(数据);
             * JSA880.去重(数据, 0);  // 按第1列去重
             */
            去重: function(data, colIndex) {
                return new Array2D(data).z去重(colIndex).val();
            },
            
            /**
             * 转置
             * @param {Array} data - 二维数组
             * @returns {Array} 转置后数组
             * @example
             * JSA880.转置([[1,2],[3,4]]);  // 返回 [[1,3],[2,4]]
             */
            转置: function(data) {
                return new Array2D(data).z转置().val();
            },
            
            /**
             * 数组求和
             * @param {Array} data - 二维数组
             * @param {string} [colSelector] - 列选择器，如 'f1'
             * @returns {number} 求和结果
             * @example
             * JSA880.求和([[1,2],[3,4]]);        // 10
             * JSA880.求和([[1,2],[3,4]], 'f1');  // 4 (第1列求和)
             */
            求和: function(data, colSelector) {
                return new Array2D(data).z求和(colSelector);
            },
            
            /**
             * 添加边框（快速版）
             * @param {string} range - 单元格范围
             * @param {number} [style=1] - 线条样式
             * @returns {boolean} 是否成功
             * @example
             * JSA880.加边框("A1:D10");
             */
            加边框: function(range, style) {
                RngUtils.z加边框(range, style || 1);
                return true;
            },
            
            /**
             * 自动列宽（快速版）
             * @param {string} range - 单元格范围
             * @returns {boolean} 是否成功
             * @example
             * JSA880.自动列宽("A:Z");
             */
            自动列宽: function(range) {
                RngUtils.z自动列宽(range);
                return true;
            },
            
            /**
             * 自动行高（快速版）
             * @param {string} range - 单元格范围
             * @returns {boolean} 是否成功
             * @example
             * JSA880.自动行高("1:100");
             */
            自动行高: function(range) {
                RngUtils.z自动行高(range);
                return true;
            },
            
            /**
             * 安全读取已使用区域
             * @param {string} [sheetName] - 工作表名称，不传则使用当前表
             * @returns {Array} 二维数组
             * @example
             * JSA880.读已用区();              // 当前表
             * JSA880.读已用区("Sheet1");      // 指定表
             */
            读已用区: function(sheetName) {
                if (!isWPS) return [];
                var sheet = sheetName ? Sheets(sheetName) : Application.ActiveSheet;
                var usedRange;
                try {
                    usedRange = sheet.UsedRange;
                } catch (e) {
                    return [];
                }
                if (!usedRange) return [];
                var arr = usedRange.Value2;
                if (arr === null || arr === undefined) return [];
                if (!Array.isArray(arr)) return [[arr]];
                if (!Array.isArray(arr[0])) {
                    var result = [];
                    for (var i = 0; i < arr.length; i++) {
                        result.push([arr[i]]);
                    }
                    return result;
                }
                return arr;
            },
            
            /**
             * 生成数字序列
             * @param {number} start - 起始数字
             * @param {number} end - 结束数字
             * @param {number} [step=1] - 步长
             * @returns {Array} 序列数组
             * @example
             * JSA880.序列(1, 10);      // [1,2,3,4,5,6,7,8,9,10]
             * JSA880.序列(1, 10, 2);  // [1,3,5,7,9]
             */
            序列: function(start, end, step) {
                step = step || 1;
                var result = [];
                for (var i = start; i <= end; i += step) {
                    result.push(i);
                }
                return result;
            },
            
            /**
             * 随机打乱数组
             * @param {Array} array - 数组
             * @returns {Array} 打乱后的数组
             * @example
             * JSA880.打乱([1,2,3,4,5]);
             */
            打乱: function(array) {
                var result = array.slice();
                for (var i = result.length - 1; i > 0; i--) {
                    var j = Math.floor(Math.random() * (i + 1));
                    var temp = result[i];
                    result[i] = result[j];
                    result[j] = temp;
                }
                return result;
            },
            
            /**
             * 随机整数
             * @param {number} min - 最小值
             * @param {number} max - 最大值
             * @returns {number} 随机整数
             * @example
             * JSA880.随机(1, 100);
             */
            随机: function(min, max) {
                return Math.floor(Math.random() * (max - min + 1)) + min;
            },
            
            /**
             * 创建SuperMap（可视化调试字典）
             * @returns {SuperMap} SuperMap实例
             * @example
             * var map = JSA880.超级字典();
             * map.set('user1', {name: '张三', age: 25});
             * map.debug(true); // 开启调试模式
             */
            超级字典: function() {
                return new SuperMap();
            },
            
            /**
             * 从Map创建SuperMap
             * @param {Map} map - 普通Map对象
             * @returns {SuperMap} SuperMap实例
             * @example
             * var nativeMap = new Map();
             * nativeMap.set('a', 1);
             * var superMap = JSA880.SuperMap从Map(nativeMap);
             */
            SuperMap从Map: function(map) {
                return SuperMap.fromMap(map);
            },
            
            /**
             * 从对象创建SuperMap
             * @param {Object} obj - 普通对象
             * @returns {SuperMap} SuperMap实例
             * @example
             * var superMap = JSA880.SuperMap从对象({a: 1, b: 2});
             */
            SuperMap从对象: function(obj) {
                return SuperMap.fromObject(obj);
            },
            
            /**
             * 从数组创建SuperMap
             * @param {Array} arr - 二维数组，每个元素为[key, value]
             * @returns {SuperMap} SuperMap实例
             * @example
             * var superMap = JSA880.SuperMap从数组([['key1', 'value1'], ['key2', 'value2']]);
             */
            SuperMap从数组: function(arr) {
                return SuperMap.fromArray(arr);
            },
            
            /**
             * 日期格式化
             * @param {Date|string} date - 日期
             * @param {string} format - 格式字符串，如 'yyyy-MM-dd HH:mm:ss'
             * @returns {string} 格式化后的日期字符串
             * @example
             * JSA880.日期格式(new Date(), 'yyyy-MM-dd');
             */
            日期格式: function(date, format) {
                date = typeof date === 'string' ? new Date(date) : date;
                var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
                return format.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
                    if (year) return date.getFullYear().toString().padStart(year.length, '0');
                    if (month) return (date.getMonth() + 1).toString().padStart(month.length, '0');
                    if (day) return date.getDate().toString().padStart(day.length, '0');
                    if (hour) return date.getHours().toString().padStart(hour.length, '0');
                    if (minute) return date.getMinutes().toString().padStart(minute.length, '0');
                    if (second) return date.getSeconds().toString().padStart(second.length, '0');
                    if (millisecond) return date.getMilliseconds().toString().padStart(3, '0');
                    if (week) return '周' + weekDays[date.getDay()];
                    return match;
                });
            },
            
            /**
             * 人民币大写
             * @param {number} n - 数字
             * @returns {string} 人民币大写
             * @example
             * JSA880.人民币大写(12345.67);  // 壹万贰仟叁佰肆拾伍元陆角柒分
             */
            人民币大写: JSA.z人民币大写,
            
            /**
             * 字符串全局替换
             * @param {string} str - 原字符串
             * @param {string} search - 查找字符串
             * @param {string} replacement - 替换字符串
             * @returns {string} 替换后的字符串
             * @example
             * JSA880.替换("hello world", "world", "JSA880");  // "hello JSA880"
             */
            替换: function(str, search, replacement) {
                return str.split(search).join(replacement);
            },
            
            /**
             * 数组扁平化
             * @param {Array} arr - 多维数组
             * @returns {Array} 一维数组
             * @example
             * JSA880.扁平化([[1,2],[3,4],[5,6]]);  // [1,2,3,4,5,6]
             */
            扁平化: function(arr) {
                return new Array2D(arr).z扁平化();
            },
            
            /**
             * 列号转字母（Excel列名）
             * @param {number} n - 列号（从1开始）
             * @returns {string} 列字母，如 1->A, 27->AA
             * @example
             * JSA880.列号(1);   // "A"
             * JSA880.列号(27);  // "AA"
             */
            列号: function(n) {
                var result = '';
                while (n > 0) {
                    n--;
                    result = String.fromCharCode(65 + (n % 26)) + result;
                    n = Math.floor(n / 26);
                }
                return result;
            },
            
            /**
             * 分组汇总连接 - 优化sumifs和Countifs批量条件统计
             * @param {Array} targetData - 统计目标数据（左表）
             * @param {Array} sourceData - 数据源（右表）
             * @param {string} groupKey - 分组键，如 'f2' 或 'f2,f3'
             * @param {string} aggFunc - 汇总函数，如 'sum("f4")' 或 'count(),sum("f5")'
             * @returns {Array} 连接汇总后的结果
             * @example
             * // 一行代码完成sumifs/countifs批量统计（高频办公场景优化）
             * JSA880.分组汇总连接(目标表, 源数据表, 'f2', 'sum("f4")');
             * JSA880.分组汇总连接(目标表, 源数据表, '月份,产品', 'count(),sum("销量"),average("金额")');
             * 
             * // 对比传统方式：groupInto + leftJoin 需要两行代码
             * // JSA880.分组汇总连接 只需一行，速度提升100倍！
             */
            分组汇总连接: function(targetData, sourceData, groupKey, aggFunc) {
                return Array2D.groupIntoJoin(targetData, sourceData, groupKey, aggFunc);
            },
            
            /**
             * 列字母转列号
             * @param {string} col - 列字母，如 "A", "AA"
             * @returns {number} 列号（从1开始）
             * @example
             * JSA880.列字母("A");   // 1
             * JSA880.列字母("AA");  // 27
             */
            列字母: function(col) {
                var result = 0;
                for (var i = 0; i < col.length; i++) {
                    result = result * 26 + (col.charCodeAt(i) - 64);
                }
                return result;
            }
        };

    }).call(this);
    
    // 导出JSA880快捷对象到全局（与Array2D、RngUtils等同级）
    (function() {
        // WPS环境
        if (typeof Application !== 'undefined') {
            Application.JSA880 = this.JSA880;
            Application.TreeNode = TreeNode;
            Application.SuperMap = SuperMap;
        }
        // Node.js环境
        if (typeof module !== 'undefined' && module.exports) {
            module.exports.JSA880 = this.JSA880;
            module.exports.TreeNode = TreeNode;
            module.exports.SuperMap = SuperMap;
        }
        // Browser环境
        if (typeof window !== 'undefined') {
            window.JSA880 = this.JSA880;
            window.TreeNode = TreeNode;
            window.SuperMap = SuperMap;
        }
    }).call(this);
}
