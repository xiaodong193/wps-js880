/**
 * JSA880-enhanced.js - 郑广学JSA880快速开发框架（智能提示版本）
 * 原作者: 郑广学 (EXCEL880)
 * 改造: Claude Code
 * 版本: 3.2.0 (ES6增强版 + RngUtils完整版 + As类型转换类)
 *
 * @description 完整的JSA880框架，支持智能提示、链式调用、中英双语API
 * @description 基于ES6语法重构，支持WPS JSA V8引擎
 * @description RngUtils完整实现，支持静态方法和实例方法
 * @description As类型转换包装类，支持智能提示和链式调用
 * @example
 * // 支持智能提示和链式调用
 * Array2D([[1,2],[3,4]]).z求和().z转置().val();
 * RngUtils("A1").z加边框().z自动列宽();
 * RngUtils.z最大行("A:A");  // 静态方法
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

// ==================== Array2D - 二维数组工具库 ====================

/**
 * Array2D - 二维数组处理工具（支持智能提示和链式调用）
 * @class
 * @description 提供丰富的二维数组操作函数，支持中英双语API
 * @example
 * // 基本使用
 * Array2D([[1,2,3],[4,5,6]]).z求和()        // 21
 * // 链式调用
 * Array2D([[1,2],[3,4]]).z转置().z扁平化().val()  // [1,3,2,4]
 * // Lambda表达式
 * Array2D([[1,2],[3,4]]).z求和('f1')       // 4 (第1列求和)
 */
function Array2D(data) {
    // 支持工厂模式调用
    if (!(this instanceof Array2D)) {
        return new Array2D(data);
    }

    this._original = data;
    this._items = null;
    this._init(data);
}

/**
 * 初始化数组
 * @private
 */
Array2D.prototype._init = function(data) {
    if (data === null || data === undefined) {
        this._items = [];
    } else if (Array.isArray(data)) {
        this._items = data;
    } else {
        this._items = [[data]];
    }
};

/**
 * 创建新实例（链式调用核心）
 * @private
 * @param {Array} data - 新数据
 * @returns {Array2D} 新实例
 */
Array2D.prototype._new = function(data) {
    const instance = new Array2D();
    instance._items = data;
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

// ==================== 填充操作 ====================

/**
 * 批量填充数组
 * @param {any} value - 填充值
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
 * 补齐空位（用指定值填充null/undefined）
 * @param {any} fillValue - 填充值，默认为空字符串
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z补齐空位 = function(fillValue) {
    fillValue = fillValue !== undefined ? fillValue : '';
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = [];
        var maxLen = 0;
        // 找出最大列数
        for (var r = 0; r < this._items.length; r++) {
            if (this._items[r].length > maxLen) {
                maxLen = this._items[r].length;
            }
        }
        for (var j = 0; j < maxLen; j++) {
            var val = this._items[i][j];
            row.push((val === null || val === undefined) ? fillValue : val);
        }
        result.push(row);
    }
    return this._new(result);
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
 * @returns {Number} 最大值
 */
Array2D.prototype.z最大值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)));
    return Math.max(...numbers);
};
Array2D.prototype.max = Array2D.prototype.z最大值;

/**
 * 求最小值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 最小值
 */
Array2D.prototype.z最小值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)));
    return Math.min(...numbers);
};
Array2D.prototype.min = Array2D.prototype.z最小值;

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
 * @param {any} value - 新值
 * @returns {Array2D} 当前实例
 */
Array2D.prototype.z设置单元格 = function(row, col, value) {
    if (!this._items[row]) this._items[row] = [];
    this._items[row][col] = value;
    return this;
};
Array2D.prototype.setCell = Array2D.prototype.z设置单元格;

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
 */
Array2D.prototype.z挑选 = function(count) {
    return this._new(this._items.slice(0, count));
};
Array2D.prototype.pick = Array2D.prototype.z挑选;

/**
 * 跳过元素
 * @param {Number} count - 跳过数量
 * @returns {Array2D} 新实例
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
        result.push(this._items[i] ? this._items[i][index] : undefined);
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
 */
Array2D.prototype.z删除行 = function(index) {
    var result = this._items.slice();
    result.splice(index, 1);
    return this._new(result);
};
Array2D.prototype.deleteRow = Array2D.prototype.z删除行;

/**
 * 删除列
 * @param {Number} index - 列号
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z删除列 = function(index) {
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        newRow.splice(index, 1);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.deleteCol = Array2D.prototype.z删除列;

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
 */
Array2D.prototype.z行排序 = function(colIndex, ascending) {
    ascending = ascending !== undefined ? ascending : true;
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

    // 处理字符串参数：支持 "f2,f3,f6" 或 "col1,col2,col3" 格式
    if (typeof cols === 'string') {
        // 检查是否是 f 模式（列号格式）
        if (cols.includes(',') && (cols.toLowerCase().includes('f'))) {
            // f 模式：按逗号分割，转换为列索引数组
            var parts = cols.split(',');
            indexes = [];
            for (var i = 0; i < parts.length; i++) {
                var part = parts[i].trim();
                if (part.toLowerCase().startsWith('f')) {
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
            // f 模式数组：转换为列索引
            indexes = cols.map(function(c) { return parseInt(c.substring(1)) - 1; });
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
 * 矩阵分布（getMatrix）- 生成数字序列的矩阵分布
 * @param {Number} totalRows - 总行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'或'c'
 * @returns {Array} 分布后的数组
 */
Array2D.getMatrix = function(totalRows, cols, direction) {
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
 * 静态方法：转置
 */
Array2D.z转置 = function(arr) {
    return new Array2D(arr).z转置().val();
};
Array2D.transpose = Array2D.z转置;

/**
 * 静态方法：求和
 */
Array2D.z求和 = function(arr, colSelector) {
    return new Array2D(arr).z求和(colSelector);
};
Array2D.sum = Array2D.z求和;

/**
 * 静态方法：克隆
 */
Array2D.z克隆 = function(arr) {
    return new Array2D(arr).z克隆().val();
};
Array2D.copy = Array2D.z克隆;

/**
 * 静态方法：选择列
 */
Array2D.z选择列 = function(arr, cols, newHeaders) {
    return new Array2D(arr).z选择列(cols, newHeaders).val();
};
Array2D.selectCols = Array2D.z选择列;

/**
 * 静态方法：批量填充
 */
Array2D.z批量填充 = function(arr, value, rows, cols) {
    return new Array2D(arr).z填充(value, rows, cols).val();
};
Array2D.fill = Array2D.z批量填充;

/**
 * 静态方法：写入单元格
 */
Array2D.z写入单元格 = function(arr, rng) {
    if (!isWPS) return arr;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    targetRng.Value2 = arr;
    return arr;
};
Array2D.toRange = Array2D.z写入单元格;

/**
 * 生成静态方法（从实例方法自动生成）
 */
(function() {
    var propNames = Object.getOwnPropertyNames(Array2D.prototype);
    // 已经手动定义的静态方法，跳过自动生成
    var manuallyDefined = ['z选择列', 'selectCols', 'z批量填充', 'fill', 'z写入单元格', 'toRange', 'z转置', 'transpose', 'z求和', 'sum', 'z克隆', 'copy'];

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
 * z安全数组 - 将指定区域转换为安全二维数组
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array} 二维数组
 * @example
 * RngUtils.z安全数组("A1:A13")
 */
RngUtils.z安全数组 = function(rng) {
    if (!isWPS) return [];
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var arr = r.Value2;
    if (arr === null || arr === undefined) return [];
    // 单个单元格转二维数组
    if (!Array.isArray(arr)) return [[arr]];
    // 一维数组转二维
    if (!Array.isArray(arr[0])) {
        var result = [];
        for (var i = 0; i < arr.length; i++) {
            result.push([arr[i]]);
        }
        return result;
    }
    return arr;
};
RngUtils.safeArray = RngUtils.z安全数组;

/**
 * z最大行 - 获取指定区域的最大行数
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @example
 * RngUtils.z最大行("A:A")  // 13
 */
RngUtils.z最大行 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var safe = RngUtils.z安全区域(r);
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
RngUtils.maxRange = RngUtils.z最大行区域;

/**
 * z最大列 - 获取指定区域的最大列数
 * @param {Range|string} rng - 要获取最大列数的区域
 * @returns {number} 最大列数
 * @example
 * RngUtils.z最大列("3:3")  // 3
 */
RngUtils.z最大列 = function(rng) {
    if (!isWPS) return 0;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var lastCol = 0;
    for (var col = 1; col <= 16384; col++) {
        var c = sheet.Cells(r.Row, col);
        if (c.Value2 !== null && c.Value2 !== undefined && c.Value2 !== '') {
            lastCol = col;
        }
    }
    return lastCol;
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

    if (direction === '-r' || direction === 'r') {
        // 按行合并
        for (var i = 1; i <= r.Rows.Count; i++) {
            var startCol = 1;
            for (var j = 2; j <= r.Columns.Count; j++) {
                if (r.Cells(i, j).Value2 !== r.Cells(i, startCol).Value2) {
                    if (j - startCol > 1) {
                        r.Range(r.Cells(i, startCol), r.Cells(i, j - 1)).Merge();
                    }
                    startCol = j;
                }
            }
            if (r.Columns.Count - startCol + 1 > 1) {
                r.Range(r.Cells(i, startCol), r.Cells(i, r.Columns.Count)).Merge();
            }
        }
    } else if (direction === '-c' || direction === 'c') {
        // 按列合并
        for (var j = 1; j <= r.Columns.Count; j++) {
            var startRow = 1;
            for (var i = 2; i <= r.Rows.Count; i++) {
                if (r.Cells(i, j).Value2 !== r.Cells(startRow, j).Value2) {
                    if (i - startRow > 1) {
                        r.Range(r.Cells(startRow, j), r.Cells(i - 1, j)).Merge();
                    }
                    startRow = i;
                }
            }
            if (r.Rows.Count - startRow + 1 > 1) {
                r.Range(r.Cells(startRow, j), r.Cells(r.Rows.Count, j)).Merge();
            }
        }
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
 * @param {Range|string} rng - 单元格区域
 * @param {number} [cols] - 选择列作为获取最大行依据
 * @returns {Array} 结果二维数组
 * @example
 * RngUtils.z最大行数组("a:d")
 */
RngUtils.z最大行数组 = function(rng, cols) {
    if (!isWPS) return [];
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var maxRow = 0;

    if (cols !== undefined) {
        // 按指定列获取最大行
        var col = typeof cols === 'number' ? r.Columns(cols) : r;
        var safe = RngUtils.z安全区域(col);
        if (safe) {
            maxRow = safe.Row + safe.Rows.Count - 1;
        }
    } else {
        // 全部列中的最大行
        maxRow = RngUtils.z最大行(r);
    }

    var result = [];
    for (var i = r.Row; i <= maxRow; i++) {
        var rowData = [];
        for (var j = 1; j <= r.Columns.Count; j++) {
            rowData.push(sheet.Cells(i, r.Column + j - 1).Value2);
        }
        result.push(rowData);
    }
    return result;
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

// ==================== 实例方法 - 支持链式调用 ====================

/**
 * 加边框
 * @param {Number} lineStyle - 线条样式（默认1）
 * @param {Number} weight - 线条粗细（默认2）
 * @returns {RngUtils} 当前实例
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
 */
RngUtils.prototype.z清除格式 = function() {
    if (!this._range) return this;
    this._range.ClearFormats();
    return this;
};
RngUtils.prototype.clearFormats = RngUtils.prototype.z清除格式;

/**
 * 自动列宽
 * @returns {RngUtils} 当前实例
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

// ==================== ShtUtils - 工作表工具库 ====================

/**
 * ShtUtils - 工作表操作工具（支持智能提示和链式调用）
 * @class
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
 * 获取星期（0-6，0=周日）
 * @returns {Number} 星期（0-6）
 * @example
 * asDate("2023-9-21").z星期()  // 4 (周四)
 */
DateUtils.prototype.z星期 = function() {
    return this._date.getDay();
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
 * 写入单元格
 * @param {Array} arr - 数组
 * @param {Range|string} rng - 单元格区域
 * @param {Boolean} clearDown - 是否清空下方
 * @returns {Range} 写入的Range
 */
JSA.z写入单元格 = function(arr, rng, clearDown) {
    if (!isWPS) return null;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    targetRng.Value2 = arr;
    return targetRng;
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
            var paddingStr = paddingNeeded > 0 ? ' '.repeat(paddingNeeded) : '';

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
 * asArray函数 - 将值转换为数组
 * @param {any} a - 要转换的值
 * @returns {Array} 数组
 * @example
 * asArray(123)           // [123]
 * asArray("abc")         // ["abc"]
 * asArray([1,2,3])       // [1,2,3]
 * asArray("a,b,c")       // ["a","b","c"] (按逗号分割)
 */
function asArray(a) {
    if (Array.isArray(a)) return a;
    if (a === null || a === undefined) return [];
    if (typeof a === 'string') {
        // 尝试按逗号分割
        if (a.indexOf(',') >= 0) {
            return a.split(',').map(function(s) { return s.trim(); });
        }
        return [a];
    }
    return [a];
}

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
 * 获取原始Sheet对象
 * @returns {Worksheet|null} 工作表对象
 */
SheetChain.prototype.value = function() {
    return this._sheet;
};
SheetChain.prototype.val = SheetChain.prototype.value;

/**
 * 获取工作表名称
 * @returns {String} 工作表名称
 */
SheetChain.prototype.z名称 = function() {
    return this._sheet ? this._sheet.Name : '';
};
SheetChain.prototype.name = SheetChain.prototype.z名称;

/**
 * 激活工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z激活 = function() {
    if (this._sheet) this._sheet.Activate();
    return this;
};
SheetChain.prototype.Activate = SheetChain.prototype.z激活;

/**
 * 获取已使用区域
 * @returns {Range|null} 已使用区域
 */
SheetChain.prototype.z已使用区域 = function() {
    if (!this._sheet) return null;
    try {
        return this._sheet.UsedRange;
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.usedRange = SheetChain.prototype.z已使用区域;

/**
 * 获取安全已使用区域（处理空表情况）
 * @returns {Range|null} 安全区域
 */
SheetChain.prototype.z安全已使用区域 = function() {
    if (!this._sheet) return null;

    var usedRange;
    try {
        usedRange = this._sheet.UsedRange;
    } catch (e) {
        return this._sheet.Range("A1");
    }

    if (!usedRange) return this._sheet.Range("A1");

    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;

    return this._sheet.Range(this._sheet.Cells(1, 1), this._sheet.Cells(lastRow, lastCol));
};
SheetChain.prototype.safeUsedRange = SheetChain.prototype.z安全已使用区域;

/**
 * 获取Range对象
 * @param {String} address - 地址
 * @returns {Range|null} Range对象
 */
SheetChain.prototype.z区域 = function(address) {
    if (!this._sheet) return null;
    try {
        return this._sheet.Range(address);
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.range = SheetChain.prototype.z区域;

/**
 * 获取Cells对象
 * @param {Number} row - 行号
 * @param {Number} col - 列号
 * @returns {Range|null} Cell对象
 */
SheetChain.prototype.z单元格 = function(row, col) {
    if (!this._sheet) return null;
    try {
        return this._sheet.Cells(row, col);
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.cell = SheetChain.prototype.z单元格;

/**
 * 删除工作表
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
SheetChain.prototype.delete = SheetChain.prototype.z删除;

/**
 * 复制工作表
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
SheetChain.prototype.copy = SheetChain.prototype.z复制;

/**
 * 保护工作表
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
SheetChain.prototype.protect = SheetChain.prototype.z保护;

/**
 * 取消保护工作表
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
SheetChain.prototype.unprotect = SheetChain.prototype.z取消保护;

/**
 * 隐藏工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z隐藏 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = false;
    return this;
};
SheetChain.prototype.hide = SheetChain.prototype.z隐藏;

/**
 * 显示工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z显示 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = true;
    return this;
};
SheetChain.prototype.show = SheetChain.prototype.z显示;

/**
 * 获取工作表索引
 * @returns {Number} 工作表索引
 */
SheetChain.prototype.z索引 = function() {
    return this._sheet ? this._sheet.Index : 0;
};
SheetChain.prototype.index = SheetChain.prototype.z索引;

/**
 * 设置工作表名称
 * @param {String} newName - 新名称
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z设置名称 = function(newName) {
    if (this._sheet) {
        this._sheet.Name = newName;
    }
    return this;
};
SheetChain.prototype.setName = SheetChain.prototype.z设置名称;

/**
 * 判断工作表是否存在
 * @returns {Boolean} 是否存在
 */
SheetChain.prototype.z存在 = function() {
    return this._sheet !== null;
};
SheetChain.prototype.exists = SheetChain.prototype.z存在;

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
 * RangeChain - Range链式调用包装类
 * @private
 * @class
 * @description 支持Range方法的链式调用
 * @example
 * $.maxRange("A1:J1").safeArray()  // 链式调用
 */
function RangeChain(rng) {
    if (!(this instanceof RangeChain)) {
        return new RangeChain(rng);
    }
    this._range = null;

    if (typeof rng === 'string') {
        this._range = isWPS ? Range(rng) : null;
    } else if (rng && rng.Address) {
        this._range = rng;
    }
}

/**
 * 获取原始Range对象
 * @returns {Range|null} Range对象
 */
RangeChain.prototype.value = function() {
    return this._range;
};

/**
 * safeArray - 转换为安全数组
 * @returns {Array} 二维数组
 */
RangeChain.prototype.safeArray = function() {
    return RngUtils.z安全数组(this._range);
};
RangeChain.prototype.z安全数组 = RangeChain.prototype.safeArray;

/**
 * maxArray - 获取最大行数组
 * @param {number} [cols] - 列号
 * @returns {Array} 二维数组
 */
RangeChain.prototype.maxArray = function(cols) {
    return RngUtils.z最大行数组(this._range, cols);
};
RangeChain.prototype.z最大行数组 = RangeChain.prototype.maxArray;

/**
 * visibleArray - 转换可见区域为数组
 * @param {Worksheet} [tempSheet] - 临时工作表
 * @returns {Array} 数组
 */
RangeChain.prototype.visibleArray = function(tempSheet) {
    return RngUtils.z可见区数组(this._range, tempSheet);
};
RangeChain.prototype.z可见区数组 = RangeChain.prototype.visibleArray;

/**
 * rowsCount - 获取行数
 * @returns {number} 行数
 */
RangeChain.prototype.rowsCount = function() {
    return this._range ? this._range.Rows.Count : 0;
};
RangeChain.prototype.z行数 = RangeChain.prototype.rowsCount;

/**
 * colsCount - 获取列数
 * @returns {number} 列数
 */
RangeChain.prototype.colsCount = function() {
    return this._range ? this._range.Columns.Count : 0;
};
RangeChain.prototype.z列数 = RangeChain.prototype.colsCount;

/**
 * address - 获取地址
 * @returns {string} 地址
 */
RangeChain.prototype.address = function() {
    return this._range ? this._range.Address() : '';
};
RangeChain.prototype.z地址 = RangeChain.prototype.address;

/**
 * addBorders - 添加边框
 * @param {number} [lineStyle=1] - 线条样式
 * @param {number} [weight=2] - 线条粗细
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.addBorders = function(lineStyle, weight) {
    if (this._range) {
        RngUtils.z加边框(this._range, lineStyle, weight);
    }
    return this;
};
RangeChain.prototype.z加边框 = RangeChain.prototype.addBorders;

/**
 * autoFitColumns - 自动列宽
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.autoFitColumns = function() {
    if (this._range) {
        this._range.Columns.AutoFit();
    }
    return this;
};
RangeChain.prototype.z自动列宽 = RangeChain.prototype.autoFitColumns;

/**
 * autoFitRows - 自动行高
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.autoFitRows = function() {
    if (this._range) {
        this._range.Rows.AutoFit();
    }
    return this;
};
RangeChain.prototype.z自动行高 = RangeChain.prototype.autoFitRows;

/**
 * clearContents - 清除内容
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.clearContents = function() {
    if (this._range) {
        this._range.ClearContents();
    }
    return this;
};
RangeChain.prototype.z清除内容 = RangeChain.prototype.clearContents;

/**
 * clearFormats - 清除格式
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.clearFormats = function() {
    if (this._range) {
        this._range.ClearFormats();
    }
    return this;
};
RangeChain.prototype.z清除格式 = RangeChain.prototype.clearFormats;

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
 * $函数 - Range快捷方式和RngUtils方法代理
 * @param {string|number} x - 地址或行列号
 * @param {number} [y] - 列号（可选，当传入两个数字参数时）
 * @returns {Range|null|RangeChain} Range对象或RangeChain包装对象
 * @example
 * $("A1")               // A1单元格（Range对象）
 * $(5, 1)              // 第5行第1列
 * $.maxRange("A1:J1")   // 返回RangeChain，支持链式调用
 * $.maxRange("A1:J1").safeArray()  // 链式调用
 */
function $(x, y) {
    // 两个参数模式：$(行, 列) - 直接返回Range
    if (arguments.length === 2 && typeof x === 'number' && typeof y === 'number') {
        return isWPS ? Cells(x, y) : null;
    }
    // 单个参数模式 - 返回Range对象（不包装，保持向后兼容）
    if (typeof x === 'string') {
        return isWPS ? Range(x) : null;
    } else if (typeof x === 'number') {
        return isWPS ? Cells(x, 1) : null;
    } else if (x && x.Address) {
        return x;
    }
    return null;
}

// 将RngUtils静态方法添加到$函数对象上，支持$.method()调用
// 这些方法返回RangeChain，支持链式调用
var proxy = createRngUtilsProxy();
for (var key in proxy) {
    if (proxy.hasOwnProperty(key)) {
        $[key] = proxy[key];
    }
}

// ==================== 将构造函数类工厂添加到$对象 ====================

/**
 * $.Array2D - 二维数组工具类工厂
 * @param {any} data - 输入数据
 * @returns {Array2D} Array2D实例
 * @example
 * $.Array2D([[1,2],[3,4]]).z求和()  // 10
 * $.Array2D([1,2,3]).z转置()        // [[1],[2],[3]]
 */
$.Array2D = function(data) {
    return new Array2D(data);
};

/**
 * $.RngUtils - Range工具类工厂
 * @param {string|Range} initialRange - 初始Range
 * @returns {RngUtils} RngUtils实例
 * @example
 * $.RngUtils("A1:B10").z安全数组()
 */
$.RngUtils = function(initialRange) {
    return new RngUtils(initialRange);
};

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
        this.As = As;
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
    }).call(this);
}

/**
 * @fileoverview JSA880-enhanced.js - 郑广学JSA880快速开发框架（智能提示版本）
 * @author 郑广学 (EXCEL880)
 * @version 3.2.0
 * @description 完整的JSA880框架，支持智能提示、链式调用、中英双语API
 *
 * 主要特性：
 * 1. 函数式构造函数 - 支持智能提示
 * 2. 原型方法定义 - WPS编辑器自动补全
 * 3. 链式调用 - 通过 .val() 获取结果
 * 4. 中英双语API - 中文方法和英文别名
 * 5. Lambda表达式 - 支持 $0, $1, f1, f2 语法
 * 6. WPS环境集成 - 完整支持WPS JS宏
 * 7. RngUtils增强 - 完整的Range操作函数库
 * 8. $快捷函数 - 支持$.z最大行("A1")快捷调用
 * 9. As类型转换类 - 支持智能提示的类型转换包装
 *
 * @example
 * // Array2D 示例
 * Array2D([[1,2],[3,4]]).z求和()           // 10
 * Array2D([[1,2],[3,4]]).z求和('f1')        // 4
 * Array2D([[1,2],[3,4]]).z转置().val()     // [[1,3],[2,4]]
 *
 * // RngUtils 静态方法示例
 * RngUtils.z最大行("A:A")                   // 13
 * RngUtils.z安全数组("A1:C10")              // [[...],[...]]
 * RngUtils.z加边框("A1:D10")                // 添加边框
 *
 * // RngUtils 实例方法示例（链式调用）
 * RngUtils("A1:C10").z加边框().z自动列宽()
 *
 * // $快捷函数示例
 * $.z最大行("A:A")                          // 13
 * $.endRow("A:A")                           // 13
 * $("A1")                                   // Range对象
 *
 * // DateUtils 示例
 * DateUtils.dt().z加天(5).z月底().val()
 *
 * // As 类型转换类示例（支持智能提示）
 * As([[1,2],[3,4]]).toArray().z转置().z求和().val()  // 10
 * As("123").toNumber().val()                          // 123
 * As(123).toString().val()                            // "123"
 * As("2023-9-1").toDate().val()                       // Date对象
 *
 * // JSA 示例
 * JSA.z转置([[1,2,3],[4,5,6]])           // [[1,4],[2,5],[3,6]]
 * JSA.z今天()                              // "2025-01-15"
 */
