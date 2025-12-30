/**
 * Array2D.js - 郑广学JSA880快速开发框架中的二维数组处理工具
 * 根据 https://vbayyds.com/api/jsa880/Array2D.html 文档编写
 * 作者: Claude Code (基于郑广学JSA880框架)
 * 版本: 1.0.0
 *
 * @description 提供丰富的二维数组操作函数，支持中文和英文函数名
 * @example
 * // 基本使用
 * const arr = new Array2D();
 * const data = [[1, 2, 3], [4, 5, 6]];
 * console.log(arr.sum(data));      // 21
 * console.log(arr.average(data));  // 3.5
 * console.log(arr.max(data));      // 6
 * console.log(arr.min(data));      // 1
 */

// 模块标识
const MODULE_NAME = "Array2D";
const VERSION = "1.0.0";
const AUTHOR = "郑广学JSA880框架";

/**
 * @typedef {Object} Array2D
 * @property {function(): string} version 获取版本信息
 * @property {function(Array<any>): boolean} isEmpty 检查数组是否为空
 * @property {function(Array<any>): number} count 获取数组元素数量
 * @property {function(Array<any>): Array<any>} copy 克隆数组
 * @property {function(Array<any>, any, number=, number=): Array<any>} fill 批量填充数组
 * @property {function(Array<any>, any): Array<any>} fillBlank 补齐空位
 * @property {function(Array<any>, number=): Array<any>} flat 降维（扁平化）
 * @property {function(Array<any>): Array<any>} reverse 数组反转
 * @property {function(Array<any>): number} sum 求和
 * @property {function(Array<any>): number} average 求平均值
 * @property {function(Array<any>): number} max 求最大值
 * @property {function(Array<any>): number} min 求最小值
 * @property {function(Array<any>): any} first 获取第一个元素
 * @property {function(Array<any>): any} last 获取最后一个元素
 * @property {function(Array<any>): Array<any>} transpose 转置矩阵
 * @property {function(Array<any>): Array<any>} toMatrix 转换为矩阵格式
 * @property {function(Array<any>, number, number): any} cell 获取单元格值
 * @property {function(Array<any>, number, number, any): void} setCell 设置单元格值
 * @property {function(Array<any>, string): string} join 连接成字符串
 * @property {function(Array<any>, number): Array<any>} chunk 分块
 * @property {function(Array<any>, number): Array<any>} pick 挑选元素
 * @property {function(Array<any>, any): Array<any>} pluck 提取列
 * @property {function(Array<any>, number): Array<any>} skip 跳过元素
 * @property {function(Array<any>, number): Array<any>} take 取前N个
 * @property {function(Array<any>, any): number} findIndex 查找元素下标
 * @property {function(Array<any>, any): boolean} includes 检查是否包含元素
 * @property {function(Array<any>, Function): Array<any>} filter 筛选元素
 * @property {function(Array<any>, Function): Array<any>} map 映射转换
 * @property {function(Array<any>, Function, any): any} reduce 归约计算
 * @property {function(Array<any>, Function): boolean} every 检查是否全部满足
 * @property {function(Array<any>, Function): boolean} some 检查是否有满足
 * @property {function(Array<any>): number} rowCount 获取行数
 * @property {function(Array<any>): number} colCount 获取列数
 * @property {function(Array<any>, number): Array<any>} getRow 获取指定行
 * @property {function(Array<any>, number): Array<any>} getCol 获取指定列
 * @property {function(Array<any>, number): Array<any>} firstRow 获取第一行
 * @property {function(Array<any>): Array<any>} lastRow 获取最后一行
 * @property {function(Array<any>): Array<any>} firstCol 获取第一列
 * @property {function(Array<any>): Array<any>} lastCol 获取最后一列
 * @property {function(Array<any>, Array<any>): Array<any>} addRow 添加行
 * @property {function(Array<any>, Array<any>, number=): Array<any>} addCol 添加列
 * @property {function(Array<any>, number): Array<any>} deleteRow 删除行
 * @property {function(Array<any>, number): Array<any>} deleteCol 删除列
 * @property {function(Array<any>, number, boolean=): Array<any>} sortRow 行排序
 * @property {function(Array<any>, number, boolean=): Array<any>} sortCol 列排序
 * @property {function(Array<any>, number): Array<any>} distinct 去重
 * @property {function(Array<any>, any, string=, boolean=): Array<any>} groupBy 分组
 * @property {function(Array<any>, any, Function): Array<any>} pivotBy 数据透视
 */

/**
 * Array2D类 - 二维数组处理工具
 * 提供丰富的二维数组操作函数，支持中文和英文函数名
 *
 * @class
 * @description 郑广学JSA880快速开发框架中的二维数组处理工具
 * @example
 * // 创建实例
 * const arr = new Array2D();
 * // 使用方法
 * const data = [[1, 2, 3], [4, 5, 6]];
 * console.log(Array2D.sum(data)); // 21
 */
class clsArray2D {

    /**
     * 构造函数 - 创建Array2D实例
     * @constructor
     * @param {Array<any>} initialData - 初始数据（可选）
     * @description 初始化Array2D对象，支持传入初始数据
     * @example
     * const arr = new Array2D([[1, 2], [3, 4]]);
     */
    constructor(initialData = null) {
        this.MODULE_NAME = MODULE_NAME;
        this.VERSION = VERSION;
        this.AUTHOR = AUTHOR;
        this.data = initialData || [];  // 存储内部状态
        console.log(`[${MODULE_NAME}] 初始化完成 - 版本 ${VERSION}`);
    }

    // ==================== 基本信息函数 ====================

    /**
     * 获取版本信息
     * @returns {string} 版本信息
     * @example
     * const arr = new Array2D();
     * console.log(arr.version()); // "1.0.0"
     */
    version() {
        return this.VERSION;
    }

    /**
     * 获取当前数组数据
     * @returns {Array<any>} 当前数组数据
     * @example
     * const arr = new Array2D([[1, 2], [3, 4]]);
     * arr.transpose();
     * console.log(arr.val()); // [[1, 3], [2, 4]]
     */
    val() {
        return this.data;
    }

    /**
     * 检查数组是否为空
     * @param {Array<any>} array2D - 二维数组
     * @returns {boolean} 是否为空
     * @example
     * const arr = new Array2D();
     * console.log(arr.isEmpty([])); // true
     * console.log(arr.isEmpty([[1, 2]])); // false
     */
    isEmpty(array2D) {
        if (!Array.isArray(array2D) || array2D.length === 0) {
            return true;
        }
        return false;
    }

    /**
     * 解析列选择器 - 内部辅助方法
     * @param {any} selector - 列选择器 ('f1'=第1列, 数字=直接索引, 函数=回调)
     * @returns {number} 列索引（从0开始）
     * @private
     */
    _parseColSelector(selector) {
        if (typeof selector === 'number') {
            return selector;
        }

        if (typeof selector === 'string') {
            if (selector.startsWith('f')) {
                // 'f1' -> 0, 'f2' -> 1, etc.
                return parseInt(selector.substring(1)) - 1;
            }
            // 处理 'f1,f2,f3' 格式，返回第一个列索引
            const parts = selector.split(',');
            if (parts.length > 0 && parts[0].startsWith('f')) {
                return parseInt(parts[0].substring(1)) - 1;
            }
        }

        if (typeof selector === 'function') {
            // 函数选择器，返回0（默认第一列）
            return 0;
        }

        return 0;
    }

    /**
     * 获取数组元素数量 - 返回行数
     * @param {Array<any>} array2D - 二维数组
     * @returns {number} 行数
     * @example
     * const arr = new Array2D();
     * console.log(arr.count([[1, 2], [3, 4]])); // 2 (行数)
     */
    count(array2D) {
        if (this.isEmpty(array2D)) {
            return 0;
        }
        return array2D.length;
    }

    // ==================== 数组操作函数 ====================

    /**
     * 克隆数组（深拷贝）- 支持传入数组参数或链式调用
     * @param {Array<any>} array2D - 要克隆的数组（可选）
     * @returns {Array<any>|clsArray2D} 如果传入数组返回克隆的数组，否则返回实例
     * @example
     * const arr = new Array2D();
     * const cloned = arr.copy([[1, 2], [3, 4]]); // 返回 [[1, 2], [3, 4]]
     * // 或者链式调用
     * const arr2 = new Array2D([[1, 2], [3, 4]]);
     * arr2.copy(); // arr2.data 是原数据的副本
     */
    copy(array2D = null) {
        // 如果传入了数组参数，返回克隆的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }
            const result = [];
            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    result.push([...array2D[i]]);
                } else {
                    result.push(array2D[i]);
                }
            }
            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }
        this.data = result;
        return this;
    }

    /**
     * 克隆指定数组（静态方法风格，保留向后兼容）
     * @param {Array<any>} array2D - 原始数组
     * @returns {Array<any>} 克隆后的数组
     * @example
     * const arr = new Array2D();
     * const original = [[1, 2], [3, 4]];
     * const cloned = arr.copyStatic(original);
     * // cloned 是 original 的副本
     */
    copyStatic(array2D) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = [];
        for (let i = 0; i < array2D.length; i++) {
            if (Array.isArray(array2D[i])) {
                result.push([...array2D[i]]);
            } else {
                result.push(array2D[i]);
            }
        }
        return result;
    }

    /**
     * 批量填充数组 - 支持链式调用
     * @param {any} value - 填充的值
     * @param {number} start - 开始位置（可选）
     * @param {number} end - 结束位置（可选）
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    fill(value, start = 0, end = null) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        const rows = result.length;

        if (end === null) {
            end = rows;
        }

        for (let i = start; i < end && i < rows; i++) {
            if (Array.isArray(result[i])) {
                const cols = result[i].length;
                for (let j = 0; j < cols; j++) {
                    result[i][j] = value;
                }
            } else {
                result[i] = value;
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 补齐空位 - 支持链式调用
     * @param {any} defaultValue - 默认值
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    fillBlank(defaultValue = null) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        let maxCols = 0;

        // 找到最大列数
        for (let i = 0; i < result.length; i++) {
            if (Array.isArray(result[i]) && result[i].length > maxCols) {
                maxCols = result[i].length;
            }
        }

        // 补齐每行
        for (let i = 0; i < result.length; i++) {
            if (!Array.isArray(result[i])) {
                result[i] = [result[i]];
            }
            while (result[i].length < maxCols) {
                result[i].push(defaultValue);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 降维（扁平化）- 支持链式调用和静态调用
     * @param {Array} array2D - 要扁平化的数组（可选）
     * @param {number} depth - 深度（默认1）
     * @returns {Array|clsArray2D} 如果传入数组返回扁平化后的数组，否则返回实例
     */
    flat(array2D = null, depth = 1) {
        // 处理参数：如果第一个参数是数字，则是 depth
        if (typeof array2D === 'number') {
            depth = array2D;
            array2D = null;
        }

        // 如果传入了数组参数，返回扁平化后的新数组
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];

            const flatten = (arr, currentDepth) => {
                for (let i = 0; i < arr.length; i++) {
                    if (Array.isArray(arr[i]) && currentDepth < depth) {
                        flatten(arr[i], currentDepth + 1);
                    } else {
                        result.push(arr[i]);
                    }
                }
            };

            flatten(array2D, 0);
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];

        const flatten = (arr, currentDepth) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i]) && currentDepth < depth) {
                    flatten(arr[i], currentDepth + 1);
                } else {
                    result.push(arr[i]);
                }
            }
        };

        flatten(this.data, 0);
        this.data = result;
        return this;
    }

    /**
     * 获取行数
     * @param {Array<any>} array2D - 二维数组
     * @returns {number} 行数
     * @example
     * const arr = new Array2D();
     * console.log(arr.rowCount([[1, 2], [3, 4]])); // 2
     */
    rowCount(array2D) {
        if (this.isEmpty(array2D)) {
            return 0;
        }
        return array2D.length;
    }

    /**
     * 获取列数
     * @param {Array<any>} array2D - 二维数组
     * @returns {number} 列数
     * @example
     * const arr = new Array2D();
     * console.log(arr.colCount([[1, 2, 3], [4, 5, 6]])); // 3
     */
    colCount(array2D) {
        if (this.isEmpty(array2D)) {
            return 0;
        }
        // 返回第一行的列数
        return array2D[0] ? array2D[0].length : 0;
    }

    /**
     * 获取指定行
     * @param {Array<any>} array2D - 二维数组
     * @param {number} rowIndex - 行索引
     * @returns {Array<any>} 指定行的数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.getRow([[1, 2, 3], [4, 5, 6]], 1)); // [4, 5, 6]
     */
    getRow(array2D, rowIndex) {
        if (this.isEmpty(array2D)) {
            return [];
        }
        if (rowIndex >= 0 && rowIndex < array2D.length) {
            return [...array2D[rowIndex]];
        }
        return [];
    }

    /**
     * 获取指定列
     * @param {Array<any>} array2D - 二维数组
     * @param {number} colIndex - 列索引
     * @returns {Array<any>} 指定列的数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.getCol([[1, 2, 3], [4, 5, 6]], 1)); // [2, 5]
     */
    getCol(array2D, colIndex) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = [];
        for (let i = 0; i < array2D.length; i++) {
            if (Array.isArray(array2D[i]) && colIndex >= 0 && colIndex < array2D[i].length) {
                result.push(array2D[i][colIndex]);
            }
        }
        return result;
    }

    /**
     * 数组反转 - 支持传入数组参数或链式调用
     * @param {Array} array2D - 要反转的数组（可选）
     * @returns {Array|clsArray2D} 如果传入数组返回反转后的数组，否则返回实例
     */
    reverse(array2D = null) {
        // 如果传入了数组参数，返回反转后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = this.copy(array2D);
            result.reverse();
            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }
        result.reverse();
        this.data = result;
        return this;
    }

    /**
     * 获取第一行
     * @param {Array<any>} array2D - 二维数组
     * @param {boolean} skipHeader - 是否跳过表头（默认false）
     * @returns {Array<any>} 第一行数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.firstRow([[1, 2, 3], [4, 5, 6]])); // [1, 2, 3]
     */
    firstRow(array2D, skipHeader = false) {
        if (this.isEmpty(array2D)) {
            return [];
        }
        if (skipHeader && array2D.length > 1) {
            return [...array2D[1]];
        }
        return [...array2D[0]];
    }

    /**
     * 获取最后一行
     * @param {Array<any>} array2D - 二维数组
     * @returns {Array<any>} 最后一行数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.lastRow([[1, 2, 3], [4, 5, 6]])); // [4, 5, 6]
     */
    lastRow(array2D) {
        if (this.isEmpty(array2D)) {
            return [];
        }
        const lastIndex = array2D.length - 1;
        return [...array2D[lastIndex]];
    }

    /**
     * 获取第一列
     * @param {Array<any>} array2D - 二维数组
     * @param {boolean} skipHeader - 是否跳过表头（默认false）
     * @returns {Array<any>} 第一列数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.firstCol([[1, 2, 3], [4, 5, 6]])); // [1, 4]
     */
    firstCol(array2D, skipHeader = false) {
        return this.getCol(array2D, 0).slice(skipHeader ? 1 : 0);
    }

    /**
     * 获取最后一列
     * @param {Array<any>} array2D - 二维数组
     * @returns {Array<any>} 最后一列数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.lastCol([[1, 2, 3], [4, 5, 6]])); // [3, 6]
     */
    lastCol(array2D) {
        if (this.isEmpty(array2D)) {
            return [];
        }
        const maxCol = array2D[0] ? array2D[0].length - 1 : 0;
        return this.getCol(array2D, maxCol);
    }

    /**
     * 提取指定列的所有值
     * @param {Array<any>} array2D - 二维数组
     * @param {number} colIndex - 列索引
     * @returns {Array<any>} 提取的列数据
     * @example
     * const arr = new Array2D();
     * console.log(arr.pluck([[1, 2, 3], [4, 5, 6]], 1)); // [2, 5]
     */
    pluck(array2D, colIndex) {
        return this.getCol(array2D, colIndex);
    }

    // ==================== 查找和筛选函数 ====================

    /**
     * 查找单个元素/行 - 支持按行查找和逐元素查找
     * @param {Array} array2D - 二维数组
     * @param {Function} predicate - 判断函数
     *   - 按行查找: (row, rowIndex) => boolean
     *   - 逐元素查找: (element, rowIndex, colIndex) => boolean
     * @param {boolean} byRow - 是否按行查找（默认true，推荐使用）
     * @returns {any} 找到的元素/行，未找到返回undefined
     */
    find(array2D, predicate, byRow = true) {
        if (this.isEmpty(array2D)) {
            return undefined;
        }

        if (byRow) {
            // 按行查找模式（推荐）
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (predicate(row, i)) {
                    return row;
                }
            }
        } else {
            // 逐元素查找模式（向后兼容）
            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        if (predicate(array2D[i][j], i, j)) {
                            return array2D[i][j];
                        }
                    }
                } else {
                    if (predicate(array2D[i], i, -1)) {
                        return array2d[i];
                    }
                }
            }
        }

        return undefined;
    }

    /**
     * 查找元素下标
     * @param {Array} array2D - 二维数组
     * @param {any} value - 要查找的值
     * @param {number} fromIndex - 开始查找的位置（可选）
     * @returns {Object} 包含row和col的对象，未找到返回{row: -1, col: -1}
     */
    findIndex(array2D, value, fromIndex = 0) {
        if (this.isEmpty(array2D)) {
            return { row: -1, col: -1 };
        }

        for (let i = fromIndex; i < array2D.length; i++) {
            if (Array.isArray(array2D[i])) {
                const colIndex = array2D[i].indexOf(value);
                if (colIndex !== -1) {
                    return { row: i, col: colIndex };
                }
            } else if (array2D[i] === value) {
                return { row: i, col: -1 };
            }
        }

        return { row: -1, col: -1 };
    }

    /**
     * 查找所有下标
     * @param {Array} array2D - 二维数组
     * @param {any} value - 要查找的值
     * @returns {Array} 所有找到的下标数组，每个元素为{row, col}
     */
    findAllIndex(array2D, value) {
        const result = [];

        if (this.isEmpty(array2D)) {
            return result;
        }

        for (let i = 0; i < array2D.length; i++) {
            if (Array.isArray(array2D[i])) {
                for (let j = 0; j < array2D[i].length; j++) {
                    if (array2D[i][j] === value) {
                        result.push({ row: i, col: j });
                    }
                }
            } else if (array2D[i] === value) {
                result.push({ row: i, col: -1 });
            }
        }

        return result;
    }

    /**
     * 筛选数组 - 支持传入数组参数或链式调用
     * @param {Array} array2D - 要筛选的数组（可选）
     * @param {Function} predicate - 筛选函数 (row, rowIndex) => boolean
     * @returns {Array|clsArray2D} 如果传入数组返回筛选后的数组，否则返回实例
     */
    filter(array2D = null, predicate = null) {
        // 处理参数：如果第一个参数是函数，则是 predicate
        if (typeof array2D === 'function') {
            predicate = array2D;
            array2D = null;
        }

        // 如果传入了数组参数，返回筛选后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }
            if (!predicate) {
                return array2D;
            }

            const result = [];
            for (let i = 0; i < array2D.length; i++) {
                if (predicate(array2D[i], i, array2D)) {
                    result.push(Array.isArray(array2D[i]) ? [...array2D[i]] : [array2D[i]]);
                }
            }
            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (!predicate) {
            return this;
        }

        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            if (predicate(row, i)) {
                result.push(row);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 是否包含值
     * @param {Array} array2D - 二维数组
     * @param {any} value - 要查找的值
     * @returns {boolean} 是否包含
     */
    includes(array2D, value) {
        return this.findIndex(array2D, value).row !== -1;
    }

    // ==================== 统计计算函数 ====================

    /**
     * 计算平均值 - 计算二维数组中所有数字的平均值
     * @param {Array<any>} array2D - 二维数组
     * @param {any} colSelector - 列选择器（可选）'f1'表示第1列，或回调函数
     * @returns {number} 平均值
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.average(data)); // 3.5 (所有元素)
     * console.log(arr.average(data, 'f1')); // 2.5 (第一列: (1+4)/2)
     */
    average(array2D, colSelector = null) {
        if (this.isEmpty(array2D)) {
            return 0;
        }

        let sum = 0;
        let count = 0;
        const isCallback = typeof colSelector === 'function';
        const colIndex = colSelector !== null && !isCallback ? this._parseColSelector(colSelector) : null;

        for (let i = 0; i < array2D.length; i++) {
            if (isCallback) {
                // 回调函数模式
                const val = colSelector(array2D[i], i);
                if (typeof val === 'number') {
                    sum += val;
                    count++;
                }
            } else if (colIndex !== null) {
                // 指定列求平均
                if (array2D[i] && array2D[i][colIndex] !== undefined) {
                    const val = array2D[i][colIndex];
                    if (typeof val === 'number') {
                        sum += val;
                        count++;
                    }
                }
            } else {
                // 所有元素求平均
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        if (typeof array2D[i][j] === 'number') {
                            sum += array2D[i][j];
                            count++;
                        }
                    }
                } else if (typeof array2D[i] === 'number') {
                    sum += array2D[i];
                    count++;
                }
            }
        }

        return count > 0 ? sum / count : 0;
    }

    /**
     * 查找最大值 - 找出二维数组中的最大值
     * @param {Array<any>} array2D - 二维数组
     * @param {any} colSelector - 列选择器（可选）'f1'表示第1列，或回调函数
     * @returns {number} 最大值
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.max(data)); // 6 (所有元素)
     * console.log(arr.max(data, 'f1')); // 4 (第一列)
     */
    max(array2D, colSelector = null) {
        if (this.isEmpty(array2D)) {
            return undefined;
        }

        const isCallback = typeof colSelector === 'function';
        const colIndex = colSelector !== null && !isCallback ? this._parseColSelector(colSelector) : null;
        let maxValue = -Infinity;
        let found = false;

        for (let i = 0; i < array2D.length; i++) {
            if (isCallback) {
                // 回调函数模式
                const val = colSelector(array2D[i], i);
                if (typeof val === 'number' && val > maxValue) {
                    maxValue = val;
                    found = true;
                }
            } else if (colIndex !== null) {
                // 指定列求最大值
                if (array2D[i] && array2D[i][colIndex] !== undefined) {
                    const val = array2D[i][colIndex];
                    if (typeof val === 'number' && val > maxValue) {
                        maxValue = val;
                        found = true;
                    }
                }
            } else {
                // 所有元素求最大值
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        if (typeof array2D[i][j] === 'number') {
                            if (array2D[i][j] > maxValue) {
                                maxValue = array2D[i][j];
                                found = true;
                            }
                        }
                    }
                } else if (typeof array2D[i] === 'number') {
                    if (array2D[i] > maxValue) {
                        maxValue = array2D[i];
                        found = true;
                    }
                }
            }
        }

        return found ? maxValue : undefined;
    }

    /**
     * 查找最小值 - 找出二维数组中的最小值
     * @param {Array<any>} array2D - 二维数组
     * @param {any} colSelector - 列选择器（可选）'f1'表示第1列，或回调函数
     * @returns {number} 最小值
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.min(data)); // 1 (所有元素)
     * console.log(arr.min(data, 'f1')); // 1 (第一列)
     */
    min(array2D, colSelector = null) {
        if (this.isEmpty(array2D)) {
            return undefined;
        }

        const isCallback = typeof colSelector === 'function';
        const colIndex = colSelector !== null && !isCallback ? this._parseColSelector(colSelector) : null;
        let minValue = Infinity;
        let found = false;

        for (let i = 0; i < array2D.length; i++) {
            if (isCallback) {
                // 回调函数模式
                const val = colSelector(array2D[i], i);
                if (typeof val === 'number' && val < minValue) {
                    minValue = val;
                    found = true;
                }
            } else if (colIndex !== null) {
                // 指定列求最小值
                if (array2D[i] && array2D[i][colIndex] !== undefined) {
                    const val = array2D[i][colIndex];
                    if (typeof val === 'number' && val < minValue) {
                        minValue = val;
                        found = true;
                    }
                }
            } else {
                // 所有元素求最小值
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        if (typeof array2D[i][j] === 'number') {
                            if (array2D[i][j] < minValue) {
                                minValue = array2D[i][j];
                                found = true;
                            }
                        }
                    }
                } else if (typeof array2D[i] === 'number') {
                    if (array2D[i] < minValue) {
                        minValue = array2D[i];
                        found = true;
                    }
                }
            }
        }

        return found ? minValue : undefined;
    }

    /**
     * 计算中位数
     * @param {Array} array2D - 二维数组
     * @param {any} colSelector - 列选择器（可选）'f1'表示第1列，或回调函数
     * @returns {number} 中位数
     */
    median(array2D, colSelector = null) {
        const isCallback = typeof colSelector === 'function';
        const colIndex = colSelector !== null && !isCallback ? this._parseColSelector(colSelector) : null;
        const values = [];

        for (let i = 0; i < array2D.length; i++) {
            if (isCallback) {
                // 回调函数模式
                const val = colSelector(array2D[i], i);
                if (typeof val === 'number') {
                    values.push(val);
                }
            } else if (colIndex !== null) {
                // 指定列求中位数
                if (array2D[i] && array2D[i][colIndex] !== undefined) {
                    const val = array2D[i][colIndex];
                    if (typeof val === 'number') {
                        values.push(val);
                    }
                }
            } else {
                // 所有元素求中位数
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        if (typeof array2D[i][j] === 'number') {
                            values.push(array2D[i][j]);
                        }
                    }
                } else if (typeof array2D[i] === 'number') {
                    values.push(array2D[i]);
                }
            }
        }

        if (values.length === 0) {
            return undefined;
        }

        values.sort((a, b) => a - b);
        const mid = Math.floor(values.length / 2);

        if (values.length % 2 === 0) {
            return (values[mid - 1] + values[mid]) / 2;
        } else {
            return values[mid];
        }
    }

    // ==================== 数组变换函数 ====================

    /**
     * 映射生成新数组 - 支持按行映射和逐元素映射，支持链式调用和静态调用
     * @param {Array} array2D - 要映射的数组（可选）
     * @param {Function} callback - 映射函数
     *   - 按行映射: (row, rowIndex) => any
     *   - 逐元素映射: (element, rowIndex, colIndex) => any
     * @param {boolean} byRow - 是否按行映射（默认true，推荐使用）
     * @returns {Array|clsArray2D} 如果传入数组返回映射后的数组，否则返回实例
     */
    map(array2D = null, callback = null, byRow = true) {
        // 处理参数：如果第一个参数是函数，则是 callback
        if (typeof array2D === 'function') {
            byRow = callback || true;
            callback = array2D;
            array2D = null;
        } else if (callback === null && typeof array2D !== 'object') {
            // 第二个参数可能是 byRow
            byRow = array2D;
            array2D = null;
            callback = null;
        }

        // 如果传入了数组参数，返回映射后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }
            if (!callback) {
                return array2D.map(row => Array.isArray(row) ? [...row] : [row]);
            }

            const result = [];
            if (byRow) {
                // 按行映射模式（推荐）
                for (let i = 0; i < array2D.length; i++) {
                    const mappedValue = callback(array2D[i], i);
                    // 如果返回值不是数组，包装成数组
                    if (Array.isArray(mappedValue)) {
                        result.push(mappedValue);
                    } else {
                        result.push([mappedValue]);
                    }
                }
            } else {
                // 逐元素映射模式（向后兼容）
                for (let i = 0; i < array2D.length; i++) {
                    if (Array.isArray(array2D[i])) {
                        const mappedRow = [];
                        for (let j = 0; j < array2D[i].length; j++) {
                            mappedRow.push(callback(array2D[i][j], i, j));
                        }
                        result.push(mappedRow);
                    } else {
                        result.push([callback(array2D[i], i, -1)]);
                    }
                }
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];

        if (byRow) {
            // 按行映射模式（推荐）
            for (let i = 0; i < this.data.length; i++) {
                result.push(callback(this.data[i], i));
            }
        } else {
            // 逐元素映射模式（向后兼容）
            for (let i = 0; i < this.data.length; i++) {
                if (Array.isArray(this.data[i])) {
                    const mappedRow = [];
                    for (let j = 0; j < this.data[i].length; j++) {
                        mappedRow.push(callback(this.data[i][j], i, j));
                    }
                    result.push(mappedRow);
                } else {
                    result.push(callback(this.data[i], i, -1));
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 遍历执行 - 支持按行遍历和逐元素遍历
     * @param {Array} array2D - 二维数组
     * @param {Function} callback - 执行函数
     *   - 按行遍历: (row, rowIndex) => void
     *   - 逐元素遍历: (element, rowIndex, colIndex) => void
     * @param {boolean} byRow - 是否按行遍历（默认true，推荐使用）
     */
    forEach(array2D, callback, byRow = true) {
        if (this.isEmpty(array2D)) {
            return;
        }

        if (byRow) {
            // 按行遍历模式（推荐）
            for (let i = 0; i < array2D.length; i++) {
                callback(array2D[i], i);
            }
        } else {
            // 逐元素遍历模式（向后兼容）
            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    for (let j = 0; j < array2D[i].length; j++) {
                        callback(array2D[i][j], i, j);
                    }
                } else {
                    callback(array2D[i], i, -1);
                }
            }
        }
    }

    /**
     * 倒序遍历执行 - 支持按行遍历和逐元素遍历
     * @param {Array} array2D - 二维数组
     * @param {Function} callback - 执行函数
     *   - 按行遍历: (row, rowIndex) => void
     *   - 逐元素遍历: (element, rowIndex, colIndex) => void
     * @param {boolean} byRow - 是否按行遍历（默认true，推荐使用）
     */
    forEachRev(array2D, callback, byRow = true) {
        if (this.isEmpty(array2D)) {
            return;
        }

        if (byRow) {
            // 按行遍历模式（推荐）
            for (let i = array2D.length - 1; i >= 0; i--) {
                callback(array2D[i], i);
            }
        } else {
            // 逐元素遍历模式（向后兼容）
            for (let i = array2D.length - 1; i >= 0; i--) {
                if (Array.isArray(array2D[i])) {
                    for (let j = array2D[i].length - 1; j >= 0; j--) {
                        callback(array2D[i][j], i, j);
                    }
                } else {
                    callback(array2D[i], i, -1);
                }
            }
        }
    }

    /**
     * 数组聚合
     * @param {Array} array2D - 二维数组
     * @param {Function} reducer - 聚合函数
     * @param {any} initialValue - 初始值
     * @returns {any} 聚合结果
     */
    reduce(array2D, reducer, initialValue) {
        if (this.isEmpty(array2D)) {
            return initialValue;
        }

        let accumulator = initialValue;
        let startIndex = 0;

        if (initialValue === undefined) {
            if (Array.isArray(array2D[0])) {
                accumulator = array2D[0][0];
                startIndex = 0;
                // 跳过第一个元素
                let firstElementProcessed = false;
                for (let i = 0; i < array2D.length; i++) {
                    if (Array.isArray(array2D[i])) {
                        for (let j = 0; j < array2D[i].length; j++) {
                            if (!firstElementProcessed) {
                                firstElementProcessed = true;
                                continue;
                            }
                            accumulator = reducer(accumulator, array2D[i][j], i, j);
                        }
                    } else {
                        if (!firstElementProcessed) {
                            firstElementProcessed = true;
                            continue;
                        }
                        accumulator = reducer(accumulator, array2D[i], i, -1);
                    }
                }
                return accumulator;
            } else {
                accumulator = array2D[0];
                startIndex = 1;
            }
        }

        for (let i = startIndex; i < array2D.length; i++) {
            if (Array.isArray(array2D[i])) {
                for (let j = (i === startIndex ? 1 : 0); j < array2D[i].length; j++) {
                    accumulator = reducer(accumulator, array2D[i][j], i, j);
                }
            } else {
                accumulator = reducer(accumulator, array2D[i], i, -1);
            }
        }

        return accumulator;
    }

    // ==================== 数组连接函数 ====================

    /**
     * 上下连接数组 - 支持传入数组参数或链式调用
     * @param {Array} array1 - 第一个数组（可选）
     * @param {...Array} arrays - 要连接的其他数组
     * @returns {Array|clsArray2D} 如果传入数组返回连接后的数组，否则返回实例
     */
    concat(array1 = null, ...arrays) {
        // 如果传入了数组参数，返回连接后的新数组（符合官方文档规范）
        if (array1 !== null) {
            const result = [];

            // 添加第一个数组
            for (let i = 0; i < array1.length; i++) {
                if (Array.isArray(array1[i])) {
                    result.push([...array1[i]]);
                } else {
                    result.push([array1[i]]);
                }
            }

            // 添加其他数组
            for (let a = 0; a < arrays.length; a++) {
                const arr = arrays[a];
                if (arr) {
                    for (let i = 0; i < arr.length; i++) {
                        if (Array.isArray(arr[i])) {
                            result.push([...arr[i]]);
                        } else {
                            result.push([arr[i]]);
                        }
                    }
                }
            }

            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            if (arrays.length > 0 && arrays[0]) {
                this.data = this.copyStatic(arrays[0]);
                return this;
            }
            return this;
        }

        const result = this.copyStatic(this.data);

        if (arrays.length > 0 && arrays[0]) {
            const array2 = arrays[0];
            for (let i = 0; i < array2.length; i++) {
                if (Array.isArray(array2[i])) {
                    result.push([...array2[i]]);
                } else {
                    result.push(array2[i]);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 文本连接
     * @param {Array} array2D - 二维数组
     * @param {string} separator - 分隔符
     * @returns {string} 连接后的文本
     */
    join(array2D, separator = ',') {
        if (this.isEmpty(array2D)) {
            return '';
        }

        const result = [];

        for (let i = 0; i < array2D.length; i++) {
            if (Array.isArray(array2D[i])) {
                result.push(array2D[i].join(separator));
            } else {
                result.push(String(array2D[i]));
            }
        }

        return result.join(separator);
    }

    // ==================== 数组操作函数 ====================

    /**
     * 获取第一个元素/行 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {Function} predicate - 筛选条件（可选）
     * @param {any} defaultValue - 默认值（可选）
     * @returns {Array|any} 静态模式返回包含第一行的数组，链式模式返回实例
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.first(data)); // [[1, 2, 3]]
     */
    first(array2D = null, predicate = null, defaultValue = null) {
        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return defaultValue !== null ? defaultValue : [];
            }

            // 如果有谓词条件，查找第一个匹配的行
            if (predicate && typeof predicate === 'function') {
                for (let i = 0; i < array2D.length; i++) {
                    if (predicate(array2D[i], i, array2D)) {
                        return [Array.isArray(array2D[i]) ? [...array2D[i]] : [array2D[i]]];
                    }
                }
                return defaultValue !== null ? defaultValue : [];
            }

            // 没有谓词，返回第一行
            return [Array.isArray(array2D[0]) ? [...array2D[0]] : [array2D[0]]];
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = defaultValue !== null ? defaultValue : [];
            return this;
        }

        // 如果有谓词条件，查找第一个匹配的行
        if (predicate && typeof predicate === 'function') {
            for (let i = 0; i < this.data.length; i++) {
                if (predicate(this.data[i], i, this.data)) {
                    this.data = [Array.isArray(this.data[i]) ? [...this.data[i]] : [this.data[i]]];
                    return this;
                }
            }
            this.data = defaultValue !== null ? defaultValue : [];
            return this;
        }

        // 没有谓词，保留第一行
        this.data = [Array.isArray(this.data[0]) ? [...this.data[0]] : [this.data[0]]];
        return this;
    }

    /**
     * 获取最后一个元素/行 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {Function} predicate - 筛选条件（可选）
     * @param {any} defaultValue - 默认值（可选）
     * @returns {Array|any} 静态模式返回包含最后一行的数组，链式模式返回实例
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.last(data)); // [[4, 5, 6]]
     */
    last(array2D = null, predicate = null, defaultValue = null) {
        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return defaultValue !== null ? defaultValue : [];
            }

            // 如果有谓词条件，从后往前查找第一个匹配的行
            if (predicate && typeof predicate === 'function') {
                for (let i = array2D.length - 1; i >= 0; i--) {
                    if (predicate(array2D[i], i, array2D)) {
                        return [Array.isArray(array2D[i]) ? [...array2D[i]] : [array2D[i]]];
                    }
                }
                return defaultValue !== null ? defaultValue : [];
            }

            // 没有谓词，返回最后一行
            const lastItem = array2D[array2D.length - 1];
            return [Array.isArray(lastItem) ? [...lastItem] : [lastItem]];
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = defaultValue !== null ? defaultValue : [];
            return this;
        }

        // 如果有谓词条件，从后往前查找第一个匹配的行
        if (predicate && typeof predicate === 'function') {
            for (let i = this.data.length - 1; i >= 0; i--) {
                if (predicate(this.data[i], i, this.data)) {
                    this.data = [Array.isArray(this.data[i]) ? [...this.data[i]] : [this.data[i]]];
                    return this;
                }
            }
            this.data = defaultValue !== null ? defaultValue : [];
            return this;
        }

        // 没有谓词，保留最后一行
        const lastItem = this.data[this.data.length - 1];
        this.data = [Array.isArray(lastItem) ? [...lastItem] : [lastItem]];
        return this;
    }

    /**
     * 数组切片 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number} start - 开始位置
     * @param {number} end - 结束位置
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    slice(array2D = null, start = null, end = null) {
        // 参数重载处理：如果第一个参数是数字，则是链式模式
        if (typeof array2D === 'number') {
            start = array2D;
            end = start !== null && start !== undefined ? end : null;
            array2D = null;
        } else if (start === null && end === null && array2D === null) {
            // 完全没有参数
            start = 0;
            end = this.data.length;
            array2D = null;
        } else if (array2D !== null && typeof array2D !== 'number' && start === null) {
            // 只有数组参数
            start = 0;
            end = array2D.length;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];
            const actualStart = start >= 0 ? start : array2D.length + start;
            const actualEnd = end !== undefined && end !== null ? (end >= 0 ? end : array2D.length + end) : array2D.length;

            for (let i = actualStart; i < actualEnd && i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    result.push([...array2D[i]]);
                } else {
                    result.push(array2D[i]);
                }
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        const actualStart = start !== null && start !== undefined ? (start >= 0 ? start : this.data.length + start) : 0;
        const actualEnd = end !== null && end !== undefined ? (end >= 0 ? end : this.data.length + end) : this.data.length;

        for (let i = actualStart; i < actualEnd && i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 数组去重 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {any} colSelector - 列选择器（可选）
     * @param {any} resultSelector - 结果选择器（可选）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    distinct(array2D = null, colSelector = null, resultSelector = null) {
        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const seen = new Set();
            const result = [];

            for (let i = 0; i < array2D.length; i++) {
                let key;

                if (colSelector) {
                    // 按指定列去重
                    const colIndex = this._parseColSelector(colSelector);
                    if (typeof colSelector === 'function') {
                        key = JSON.stringify(colSelector(array2D[i], i));
                    } else {
                        key = array2D[i] && array2D[i][colIndex];
                    }
                } else {
                    // 全部列去重
                    key = JSON.stringify(array2D[i]);
                }

                if (!seen.has(key)) {
                    seen.add(key);
                    if (resultSelector && typeof resultSelector === 'function') {
                        result.push(resultSelector(array2D[i], i));
                    } else if (resultSelector === '') {
                        result.push(Array.isArray(array2D[i]) ? [...array2D[i]] : [array2D[i]]);
                    } else if (colSelector) {
                        const colIndex = this._parseColSelector(colSelector);
                        result.push([array2D[i] && array2D[i][colIndex]]);
                    } else {
                        result.push(Array.isArray(array2D[i]) ? [...array2D[i]] : [array2D[i]]);
                    }
                }
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const seen = new Set();
        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            let key;

            if (colSelector) {
                const colIndex = this._parseColSelector(colSelector);
                if (typeof colSelector === 'function') {
                    key = JSON.stringify(colSelector(this.data[i], i));
                } else {
                    key = this.data[i] && this.data[i][colIndex];
                }
            } else {
                key = JSON.stringify(this.data[i]);
            }

            if (!seen.has(key)) {
                seen.add(key);
                if (resultSelector && typeof resultSelector === 'function') {
                    result.push(resultSelector(this.data[i], i));
                } else if (resultSelector === '') {
                    result.push(Array.isArray(this.data[i]) ? [...this.data[i]] : [this.data[i]]);
                } else if (colSelector) {
                    const colIndex = this._parseColSelector(colSelector);
                    result.push([this.data[i] && this.data[i][colIndex]]);
                } else {
                    result.push(Array.isArray(this.data[i]) ? [...this.data[i]] : [this.data[i]]);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 数组排序 - 支持传入数组参数或链式调用
     * @param {Array} array2D - 要排序的数组（可选）
     * @param {Function} compareFunction - 比较函数（可选）
     * @returns {Array|clsArray2D} 如果传入数组返回排序后的数组，否则返回实例
     */
    sort(array2D = null, compareFunction = null) {
        // 如果传入了数组参数，返回排序后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = this.copy(array2D);

            if (compareFunction) {
                result.sort(compareFunction);
            } else {
                // 默认按首列升序排序
                result.sort((a, b) => {
                    const aVal = Array.isArray(a) ? a[0] : a;
                    const bVal = Array.isArray(b) ? b[0] : b;
                    if (aVal < bVal) return -1;
                    if (aVal > bVal) return 1;
                    return 0;
                });
            }

            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = this.copyStatic(this.data);

        if (compareFunction) {
            result.sort(compareFunction);
        } else {
            // 默认按首列升序排序
            result.sort((a, b) => {
                const aVal = Array.isArray(a) ? a[0] : a;
                const bVal = Array.isArray(b) ? b[0] : b;
                if (aVal < bVal) return -1;
                if (aVal > bVal) return 1;
                return 0;
            });
        }

        this.data = result;
        return this;
    }

    /**
     * 随机打乱数组 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    shuffle(array2D = null) {
        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];
            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    result.push([...array2D[i]]);
                } else {
                    result.push(array2D[i]);
                }
            }

            for (let i = result.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [result[i], result[j]] = [result[j], result[i]];
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        for (let i = result.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [result[i], result[j]] = [result[j], result[i]];
        }

        this.data = result;
        return this;
    }

    /**
     * 随机获取一项 - 支持链式调用
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    random() {
        if (this.isEmpty(this.data)) {
            this.data = undefined;
            return this;
        }

        const flatArray = [];
        const flatten = (arr) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i])) {
                    flatten(arr[i]);
                } else {
                    flatArray.push(arr[i]);
                }
            }
        };
        flatten(this.data);

        if (flatArray.length === 0) {
            this.data = undefined;
            return this;
        }

        const randomIndex = Math.floor(Math.random() * flatArray.length);
        this.data = flatArray[randomIndex];
        return this;
    }

    // ==================== 新增：笛卡尔积函数 ====================

    /**
     * 笛卡尔积 - 支持链式调用
     * @param {Array} array2 - 第二个数组
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    crossjoin(array2) {
        if (this.isEmpty(this.data) || this.isEmpty(array2)) {
            this.data = [];
            return this;
        }

        const result = [];

        // 确保数组是二维的
        const arr1 = Array.isArray(this.data[0]) ? this.data : [this.data];
        const arr2 = Array.isArray(array2[0]) ? array2 : [array2];

        for (let i = 0; i < arr1.length; i++) {
            const row1 = Array.isArray(arr1[i]) ? arr1[i] : [arr1[i]];
            for (let j = 0; j < arr2.length; j++) {
                const row2 = Array.isArray(arr2[j]) ? arr2[j] : [arr2[j]];
                result.push([...row1, ...row2]);
            }
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：分组函数 ====================

    /**
     * 分组 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {Function|number|string} keySelector - 分组键选择器、列索引或列选择器（如'f2'）
     * @param {string} separator - 分组键的分隔符，默认为"@^@"
     * @returns {Map|clsArray2D} 静态模式返回Map，链式模式返回实例
     */
    groupBy(array2D = null, keySelector = null, separator = '@^@') {
        // 参数重载处理
        if (keySelector === null && typeof array2D !== 'object') {
            keySelector = array2D;
            array2D = null;
        }

        // 处理列选择器字符串（如 'f2' 表示第2列）
        const getKeyValue = (row, index) => {
            let key;
            if (typeof keySelector === 'function') {
                key = keySelector(row, index);
            } else if (typeof keySelector === 'number') {
                key = Array.isArray(row) && row.length > keySelector ? row[keySelector] : row;
            } else if (typeof keySelector === 'string') {
                // 处理列选择器如 'f2'
                const colMatch = keySelector.match(/^f(\d+)$/i);
                if (colMatch) {
                    const colIndex = parseInt(colMatch[1]) - 1; // 'f2' -> 索引1
                    key = Array.isArray(row) && row.length > colIndex ? row[colIndex] : row;
                } else {
                    key = JSON.stringify(row);
                }
            } else {
                key = JSON.stringify(row);
            }
            return key;
        };

        // 静态模式：如果传入了数组参数，返回Map（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return new Map();
            }

            const result = new Map();

            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                const key = getKeyValue(row, i);
                const keyStr = String(key);

                if (!result.has(keyStr)) {
                    result.set(keyStr, []);
                }

                result.get(keyStr).push([...row]); // 存储行的副本
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = {};
            return this;
        }

        const result = {};

        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            const key = getKeyValue(row, i);
            const keyStr = String(key);

            if (!result[keyStr]) {
                result[keyStr] = [];
            }

            result[keyStr].push(row);
        }

        this.data = result;
        return this;
    }

    /**
     * 分组汇总 - 支持链式调用
     * @param {Function|number} keySelector - 分组键选择器或列索引
     * @param {Function|number} valueSelector - 值选择器或列索引
     * @param {Function} aggregator - 聚合函数，默认为求和
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    groupInto(keySelector, valueSelector, aggregator = (acc, val) => acc + val) {
        // 调用 groupBy 静态方法获取 Map
        const groupsMap = this.groupBy(this.data, keySelector);
        const result = [];

        // 遍历 Map
        for (const [key, groupRows] of groupsMap.entries()) {
            let aggregatedValue;

            if (typeof valueSelector === 'function') {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    const val = valueSelector(row, index);
                    return aggregator(acc, val, row, index);
                }, 0);
            } else if (typeof valueSelector === 'number') {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    const val = Array.isArray(row) && row.length > valueSelector ? row[valueSelector] : row;
                    return aggregator(acc, val, row, index);
                }, 0);
            } else {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    return aggregator(acc, row, row, index);
                }, 0);
            }

            result.push([key, aggregatedValue]);
        }

        this.data = result;
        return this;
    }

    /**
     * 分组汇总到字典 - 支持链式调用
     * @param {Function|number} keySelector - 分组键选择器或列索引
     * @param {Function|number} valueSelector - 值选择器或列索引
     * @param {Function} aggregator - 聚合函数，默认为求和
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    groupIntoMap(keySelector, valueSelector, aggregator = (acc, val) => acc + val) {
        // 调用 groupBy 静态方法获取 Map
        const groupsMap = this.groupBy(this.data, keySelector);
        const result = {};

        // 遍历 Map
        for (const [key, groupRows] of groupsMap.entries()) {
            let aggregatedValue;

            if (typeof valueSelector === 'function') {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    const val = valueSelector(row, index);
                    return aggregator(acc, val, row, index);
                }, 0);
            } else if (typeof valueSelector === 'number') {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    const val = Array.isArray(row) && row.length > valueSelector ? row[valueSelector] : row;
                    return aggregator(acc, val, row, index);
                }, 0);
            } else {
                aggregatedValue = groupRows.reduce((acc, row, index) => {
                    return aggregator(acc, row, row, index);
                }, 0);
            }

            result[key] = aggregatedValue;
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：连接函数 ====================

    /**
     * 左连接 - 支持链式调用和静态调用
     * @param {Array} leftArray - 左表（静态模式）或右表（链式模式）
     * @param {Array} rightArray - 右表（静态模式）或左键（链式模式）
     * @param {Function|number|string} leftKey - 左表连接键选择器或列索引
     * @param {Function|number|string} rightKey - 右表连接键选择器或列索引
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    leftjoin(leftArray = null, rightArray = null, leftKey = null, rightKey = null) {
        // 参数重载处理
        if (leftArray === null || (typeof leftArray !== 'object' || !Array.isArray(leftArray))) {
            // 链式模式: leftjoin(rightArray, leftKey, rightKey)
            if (leftArray !== null) {
                rightKey = leftKey;
                leftKey = rightArray;
                rightArray = leftArray;
            }
            leftArray = this.data;
        }

        // 静态模式: leftjoin(leftArray, rightArray, leftKey, rightKey)

        if (this.isEmpty(leftArray)) {
            if (leftArray === this.data) {
                this.data = [];
                return this;
            }
            return [];
        }

        const result = [];
        const rightMap = {};

        // 构建右表映射
        if (!this.isEmpty(rightArray)) {
            for (let i = 0; i < rightArray.length; i++) {
                const row = rightArray[i];
                let key;

                if (typeof rightKey === 'function') {
                    key = rightKey(row, i);
                } else if (typeof rightKey === 'number') {
                    key = Array.isArray(row) && row.length > rightKey ? row[rightKey] : row;
                } else if (typeof rightKey === 'string' && rightKey.startsWith('f')) {
                    // 'f1' 格式 -> 0, 'f2' -> 1
                    const colIndex = parseInt(rightKey.substring(1)) - 1;
                    key = Array.isArray(row) && row.length > colIndex ? row[colIndex] : row;
                } else {
                    key = JSON.stringify(row);
                }

                rightMap[String(key)] = row;
            }
        }

        // 执行左连接
        for (let i = 0; i < leftArray.length; i++) {
            const leftRow = leftArray[i];
            let key;

            if (typeof leftKey === 'function') {
                key = leftKey(leftRow, i);
            } else if (typeof leftKey === 'number') {
                key = Array.isArray(leftRow) && leftRow.length > leftKey ? leftRow[leftKey] : leftRow;
            } else if (typeof leftKey === 'string' && leftKey.startsWith('f')) {
                // 'f1' 格式 -> 0, 'f2' -> 1
                const colIndex = parseInt(leftKey.substring(1)) - 1;
                key = Array.isArray(leftRow) && leftRow.length > colIndex ? leftRow[colIndex] : leftRow;
            } else {
                key = JSON.stringify(leftRow);
            }

            const keyStr = String(key);
            const rightRow = rightMap[keyStr];

            if (rightRow) {
                const leftRowArr = Array.isArray(leftRow) ? leftRow : [leftRow];
                const rightRowArr = Array.isArray(rightRow) ? rightRow : [rightRow];
                result.push([...leftRowArr, ...rightRowArr]);
            } else {
                const leftRowArr = Array.isArray(leftRow) ? leftRow : [leftRow];
                const rightPlaceholder = (rightArray && rightArray[0] && Array.isArray(rightArray[0]))
                    ? new Array(rightArray[0].length).fill(null)
                    : [null];
                result.push([...leftRowArr, ...rightPlaceholder]);
            }
        }

        // 静态模式返回新数组，链式模式更新实例数据
        if (leftArray === this.data) {
            this.data = result;
            return this;
        }
        return result;
    }

    /**
     * 全连接 - 支持链式调用和静态调用
     * @param {Array} leftArray - 左表（静态模式）或右表（链式模式）
     * @param {Array} rightArray - 右表（静态模式）或左键（链式模式）
     * @param {Function|number|string} leftKey - 左表连接键选择器或列索引
     * @param {Function|number|string} rightKey - 右表连接键选择器或列索引
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    fulljoin(leftArray = null, rightArray = null, leftKey = null, rightKey = null) {
        // 参数重载处理
        let isStatic = false;
        if (leftArray === null || (typeof leftArray !== 'object' || !Array.isArray(leftArray))) {
            // 链式模式: fulljoin(rightArray, leftKey, rightKey)
            if (leftArray !== null) {
                rightKey = leftKey;
                leftKey = rightArray;
                rightArray = leftArray;
            }
            leftArray = this.data;
        } else {
            isStatic = true;
        }

        // 执行左连接（left -> right）
        const leftResult = this.leftjoin(leftArray, rightArray, leftKey, rightKey);

        // 执行左连接（right -> left）
        const rightResult = this.leftjoin(rightArray, leftArray, rightKey, leftKey);

        // 合并结果，去重
        const resultMap = new Map();

        for (const row of leftResult) {
            resultMap.set(JSON.stringify(row), row);
        }

        for (const row of rightResult) {
            const key = JSON.stringify(row);
            if (!resultMap.has(key)) {
                resultMap.set(key, row);
            }
        }

        const result = Array.from(resultMap.values());

        // 静态模式返回新数组，链式模式更新实例数据
        if (!isStatic) {
            this.data = result;
            return this;
        }
        return result;
    }

    // ==================== 新增：集合操作函数 ====================

    /**
     * 排除 - 支持链式调用
     * @param {Array} excludeArray - 要排除的数组
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    except(excludeArray) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        if (this.isEmpty(excludeArray)) {
            return this;
        }

        const excludeSet = new Set();
        const flatExclude = [];
        const flatten = (arr) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i])) {
                    flatten(arr[i]);
                } else {
                    flatExclude.push(arr[i]);
                }
            }
        };
        flatten(excludeArray);
        for (const item of flatExclude) {
            excludeSet.add(JSON.stringify(item));
        }

        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                const filteredRow = [];
                for (let j = 0; j < this.data[i].length; j++) {
                    const key = JSON.stringify(this.data[i][j]);
                    if (!excludeSet.has(key)) {
                        filteredRow.push(this.data[i][j]);
                    }
                }
                if (filteredRow.length > 0) {
                    result.push(filteredRow);
                }
            } else {
                const key = JSON.stringify(this.data[i]);
                if (!excludeSet.has(key)) {
                    result.push(this.data[i]);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 取交集 - 支持链式调用和静态调用
     * @param {Array} array1 - 第一个数组（静态模式）或第二个数组（链式模式）
     * @param {Array} array2 - 第二个数组（静态模式）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    intersect(array1 = null, array2 = null) {
        // 参数重载处理
        if (array2 === null) {
            // 链式模式: intersect(array2)
            array2 = array1;
            array1 = this.data;
        }

        // 静态模式: intersect(array1, array2)

        if (this.isEmpty(array1) || this.isEmpty(array2)) {
            if (array1 === this.data) {
                this.data = [];
                return this;
            }
            return [];
        }

        const set1 = new Set();

        // 构建第一个数组的集合（按行）
        for (let i = 0; i < array1.length; i++) {
            const key = JSON.stringify(array1[i]);
            set1.add(key);
        }

        const set2 = new Set();

        // 构建第二个数组的集合（按行）
        for (let i = 0; i < array2.length; i++) {
            const key = JSON.stringify(array2[i]);
            set2.add(key);
        }

        const result = [];
        // 找出交集
        for (const key of set1) {
            if (set2.has(key)) {
                result.push(JSON.parse(key));
            }
        }

        // 静态模式返回新数组，链式模式更新实例数据
        if (array1 === this.data) {
            this.data = result;
            return this;
        }
        return result;
    }

    // ==================== 新增：排名函数 ====================

    /**
     * 排名 - 支持链式调用
     * @param {Function|number} valueSelector - 值选择器或列索引
     * @param {boolean} ascending - 是否升序，默认为false（降序排名）
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    rank(valueSelector, ascending = false) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        // 提取值和索引
        const items = [];
        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            let value;

            if (typeof valueSelector === 'function') {
                value = valueSelector(row, i);
            } else if (typeof valueSelector === 'number') {
                value = Array.isArray(row) && row.length > valueSelector ? row[valueSelector] : row;
            } else {
                value = row;
            }

            items.push({
                original: row,
                value: value,
                index: i
            });
        }

        // 排序
        items.sort((a, b) => {
            if (ascending) {
                return a.value - b.value;
            } else {
                return b.value - a.value;
            }
        });

        // 计算排名
        let currentRank = 1;
        let previousValue = null;
        let skipCount = 0;

        for (let i = 0; i < items.length; i++) {
            const item = items[i];

            if (previousValue !== null && item.value !== previousValue) {
                currentRank += skipCount;
                skipCount = 0;
            }

            item.rank = currentRank;
            previousValue = item.value;
            skipCount++;
        }

        // 按原始顺序恢复并添加排名
        items.sort((a, b) => a.index - b.index);
        const result = [];

        for (const item of items) {
            const row = Array.isArray(item.original) ? [...item.original] : [item.original];
            row.push(item.rank);
            result.push(row);
        }

        this.data = result;
        return this;
    }

    /**
     * 分组排名 - 支持链式调用
     * @param {Function|number} groupSelector - 分组键选择器或列索引
     * @param {Function|number} valueSelector - 值选择器或列索引
     * @param {boolean} ascending - 是否升序，默认为false（降序排名）
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    rankGroup(groupSelector, valueSelector, ascending = false) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        // 调用 groupBy 静态方法获取 Map
        const groupsMap = this.groupBy(this.data, groupSelector);
        const result = [];

        // 遍历 Map，对每个组进行排名
        for (const [groupKey, groupRows] of groupsMap.entries()) {
            const temp2 = new clsArray2D(groupRows);
            const rankedGroup = temp2.rank(groupRows, valueSelector, ascending);

            // 添加回结果
            for (const row of rankedGroup) {
                result.push(row);
            }
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：批量操作函数 ====================

    /**
     * 批量插入列 - 支持链式调用
     * @param {number} index - 插入位置
     * @param {any} value - 插入的值，可以是单个值或数组
     * @param {number} count - 插入列数，默认为1
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    insertCols(index, value = null, count = 1) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = this.copyStatic(this.data);

        for (let i = 0; i < result.length; i++) {
            if (Array.isArray(result[i])) {
                const insertIndex = Math.min(Math.max(index, 0), result[i].length);
                const valuesToInsert = Array.isArray(value) ? value : new Array(count).fill(value);

                // 在指定位置插入值
                result[i].splice(insertIndex, 0, ...valuesToInsert);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 批量插入行 - 支持链式调用
     * @param {number} index - 插入位置
     * @param {any} value - 插入的值，可以是单个值或数组
     * @param {number} count - 插入行数，默认为1
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    insertRows(index, value = null, count = 1) {
        if (this.isEmpty(this.data)) {
            if (Array.isArray(value)) {
                this.data = [value];
            } else {
                this.data = new Array(count).fill([value]);
            }
            return this;
        }

        const result = this.copyStatic(this.data);
        const insertIndex = Math.min(Math.max(index, 0), result.length);

        const rowsToInsert = [];
        for (let i = 0; i < count; i++) {
            if (Array.isArray(value)) {
                rowsToInsert.push([...value]);
            } else {
                // 创建与现有行相同列数的行
                const colCount = result.length > 0 && Array.isArray(result[0]) ? result[0].length : 1;
                rowsToInsert.push(new Array(colCount).fill(value));
            }
        }

        // 在指定位置插入行
        result.splice(insertIndex, 0, ...rowsToInsert);

        this.data = result;
        return this;
    }

    /**
     * 批量删除列 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number|Array} start - 开始位置或列索引数组
     * @param {number} count - 删除列数，默认为1（当start为数字时）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    deleteCols(array2D = null, start = null, count = 1) {
        // 参数重载处理
        let indices = null;
        if (Array.isArray(start)) {
            indices = start;
            start = null;
        } else if (Array.isArray(array2D) && start === null) {
            indices = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];

            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    const row = [...array2D[i]];

                    if (indices) {
                        // 删除指定的列（从后往前删除，避免索引变化）
                        const sortedIndices = [...indices].sort((a, b) => b - a);
                        for (const idx of sortedIndices) {
                            if (idx >= 0 && idx < row.length) {
                                row.splice(idx, 1);
                            }
                        }
                    } else {
                        // 删除从start开始的count列
                        const deleteStart = Math.min(Math.max(start, 0), row.length);
                        const deleteCount = Math.min(count, row.length - deleteStart);
                        if (deleteCount > 0) {
                            row.splice(deleteStart, deleteCount);
                        }
                    }
                    result.push(row);
                } else {
                    result.push(array2D[i]);
                }
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                const row = [...this.data[i]];

                if (indices) {
                    // 删除指定的列（从后往前删除，避免索引变化）
                    const sortedIndices = [...indices].sort((a, b) => b - a);
                    for (const idx of sortedIndices) {
                        if (idx >= 0 && idx < row.length) {
                            row.splice(idx, 1);
                        }
                    }
                } else {
                    // 删除从start开始的count列
                    const deleteStart = Math.min(Math.max(start, 0), row.length);
                    const deleteCount = Math.min(count, row.length - deleteStart);
                    if (deleteCount > 0) {
                        row.splice(deleteStart, deleteCount);
                    }
                }
                result.push(row);
            } else {
                result.push(this.data[i]);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 批量删除行 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number|Array} start - 开始位置或行索引数组
     * @param {number} count - 删除行数，默认为1（当start为数字时）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    deleteRows(array2D = null, start = null, count = 1) {
        // 参数重载处理
        let indices = null;
        if (Array.isArray(start)) {
            indices = start;
            start = null;
        } else if (Array.isArray(array2D) && start === null) {
            indices = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];

            if (indices) {
                // 删除指定的行（从后往前删除，避免索引变化）
                const indicesToDelete = new Set(indices);
                for (let i = 0; i < array2D.length; i++) {
                    if (!indicesToDelete.has(i)) {
                        result.push(Array.isArray(array2D[i]) ? [...array2D[i]] : array2D[i]);
                    }
                }
            } else {
                // 删除从start开始的count行
                const resultCopy = [];
                for (let i = 0; i < array2D.length; i++) {
                    if (Array.isArray(array2D[i])) {
                        resultCopy.push([...array2D[i]]);
                    } else {
                        resultCopy.push(array2D[i]);
                    }
                }
                const deleteStart = Math.min(Math.max(start, 0), resultCopy.length);
                const deleteCount = Math.min(count, resultCopy.length - deleteStart);
                if (deleteCount > 0) {
                    resultCopy.splice(deleteStart, deleteCount);
                }
                return resultCopy;
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];

        if (indices) {
            // 删除指定的行（从后往前删除，避免索引变化）
            const indicesToDelete = new Set(indices);
            for (let i = 0; i < this.data.length; i++) {
                if (!indicesToDelete.has(i)) {
                    result.push(Array.isArray(this.data[i]) ? [...this.data[i]] : this.data[i]);
                }
            }
        } else {
            // 删除从start开始的count行
            for (let i = 0; i < this.data.length; i++) {
                if (Array.isArray(this.data[i])) {
                    result.push([...this.data[i]]);
                } else {
                    result.push(this.data[i]);
                }
            }
            const deleteStart = Math.min(Math.max(start, 0), result.length);
            const deleteCount = Math.min(count, result.length - deleteStart);
            if (deleteCount > 0) {
                result.splice(deleteStart, deleteCount);
            }
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：选择函数 ====================

    /**
     * 选择列 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number|Array} colIndices - 列索引或索引数组
     * @returns {clsArray2D|Array} 静态模式返回新数组，链式模式返回当前实例
     */
    selectCols(array2D = null, colIndices) {
        const indices = Array.isArray(colIndices) ? colIndices : [colIndices];

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (Array.isArray(row)) {
                    const selectedRow = [];
                    for (const index of indices) {
                        if (index >= 0 && index < row.length) {
                            selectedRow.push(row[index]);
                        } else {
                            selectedRow.push(null);
                        }
                    }
                    result.push(selectedRow);
                } else {
                    // 如果行不是数组，只选择第一个索引
                    result.push(indices.includes(0) ? row : null);
                }
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            if (Array.isArray(row)) {
                const selectedRow = [];
                for (const index of indices) {
                    if (index >= 0 && index < row.length) {
                        selectedRow.push(row[index]);
                    } else {
                        selectedRow.push(null);
                    }
                }
                result.push(selectedRow);
            } else {
                // 如果行不是数组，只选择第一个索引
                result.push(indices.includes(0) ? row : null);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 选择行 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number|Array} rowIndices - 行索引或索引数组
     * @returns {clsArray2D|Array} 静态模式返回新数组，链式模式返回当前实例
     */
    selectRows(array2D = null, rowIndices) {
        const indices = Array.isArray(rowIndices) ? rowIndices : [rowIndices];

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = [];
            for (const index of indices) {
                if (index >= 0 && index < array2D.length) {
                    const row = array2D[index];
                    if (Array.isArray(row)) {
                        result.push([...row]);
                    } else {
                        result.push(row);
                    }
                }
            }
            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        for (const index of indices) {
            if (index >= 0 && index < this.data.length) {
                const row = this.data[index];
                if (Array.isArray(row)) {
                    result.push([...row]);
                } else {
                    result.push(row);
                }
            }
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：按范围操作函数 ====================

    /**
     * 按范围遍历
     * @param {Array} array2D - 二维数组
     * @param {number} startRow - 开始行
     * @param {number} endRow - 结束行
     * @param {number} startCol - 开始列
     * @param {number} endCol - 结束列
     * @param {Function} callback - 回调函数
     */
    rangeForEach(array2D, startRow, endRow, startCol, endCol, callback) {
        if (this.isEmpty(array2D)) {
            return;
        }

        const actualStartRow = Math.max(startRow, 0);
        const actualEndRow = Math.min(endRow, array2D.length - 1);
        const actualStartCol = Math.max(startCol, 0);

        for (let i = actualStartRow; i <= actualEndRow; i++) {
            const row = array2D[i];
            if (Array.isArray(row)) {
                const actualEndCol = Math.min(endCol, row.length - 1);
                for (let j = actualStartCol; j <= actualEndCol; j++) {
                    callback(row[j], i, j);
                }
            } else if (i === actualStartRow && actualStartCol === 0) {
                callback(row, i, 0);
            }
        }
    }

    /**
     * 局部映射
     * @param {Array} array2D - 二维数组
     * @param {number} startRow - 开始行
     * @param {number} endRow - 结束行
     * @param {number} startCol - 开始列
     * @param {number} endCol - 结束列
     * @param {Function} callback - 映射函数
     * @returns {Array} 映射后的数组
     */
    rangeMap(array2D, startRow, endRow, startCol, endCol, callback) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = this.copy(array2D);
        const actualStartRow = Math.max(startRow, 0);
        const actualEndRow = Math.min(endRow, result.length - 1);
        const actualStartCol = Math.max(startCol, 0);

        for (let i = actualStartRow; i <= actualEndRow; i++) {
            const row = result[i];
            if (Array.isArray(row)) {
                const actualEndCol = Math.min(endCol, row.length - 1);
                for (let j = actualStartCol; j <= actualEndCol; j++) {
                    row[j] = callback(row[j], i, j);
                }
            } else if (i === actualStartRow && actualStartCol === 0) {
                result[i] = callback(row, i, 0);
            }
        }

        return result;
    }

    /**
     * 按范围选择 - 选择指定范围的元素
     * @param {Array} array2D - 二维数组
     * @param {Array|string} address - 范围数组 [行起, 列起, 行数, 列数] 或 Excel 地址格式 (如 'a1:b2')
     * @returns {Array} 选择的范围
     */
    rangeSelect(array2D, address) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        // 解析地址参数
        let startRow, startCol, rowCount, colCount;

        if (typeof address === 'string') {
            // Excel 地址格式: 'a1:b2' -> [0, 0, 2, 2]
            const parsed = this._parseExcelAddress(address);
            startRow = parsed.startRow;
            startCol = parsed.startCol;
            rowCount = parsed.rowCount;
            colCount = parsed.colCount;
        } else if (Array.isArray(address)) {
            // 数组格式: [行起, 列起, 行数, 列数]
            startRow = address[0] || 0;
            startCol = address[1] || 0;
            rowCount = address[2] !== undefined ? address[2] : Infinity;
            colCount = address[3] !== undefined ? address[3] : Infinity;
        } else {
            return [];
        }

        const result = [];
        const actualStartRow = Math.max(startRow, 0);
        const actualStartCol = Math.max(startCol, 0);

        // 计算实际结束行和列
        const maxRow = Math.min(actualStartRow + rowCount - 1, array2D.length - 1);

        for (let i = actualStartRow; i <= maxRow; i++) {
            const row = array2D[i];
            if (Array.isArray(row)) {
                const maxCol = Math.min(actualStartCol + colCount - 1, row.length - 1);
                const selectedRow = [];

                for (let j = actualStartCol; j <= maxCol; j++) {
                    selectedRow.push(row[j]);
                }

                result.push(selectedRow);
            } else if (i === actualStartRow && actualStartCol === 0) {
                result.push([row]);
            }
        }

        return result;
    }

    /**
     * 解析 Excel 地址格式 (如 'a1:b2')
     * @param {string} address - Excel 地址
     * @returns {Object} {startRow, startCol, rowCount, colCount}
     */
    _parseExcelAddress(address) {
        if (!address || typeof address !== 'string') {
            return { startRow: 0, startCol: 0, rowCount: Infinity, colCount: Infinity };
        }

        // 转小写并移除空格
        const addr = address.toLowerCase().replace(/\s/g, '');

        // 检查是否包含冒号（范围格式）
        const colonIndex = addr.indexOf(':');

        if (colonIndex === -1) {
            // 单个单元格格式 'a1'
            const col = this._colLetterToIndex(addr.replace(/[0-9]/g, ''));
            const row = parseInt(addr.replace(/[a-z]/g, '')) - 1;
            return { startRow: Math.max(row, 0), startCol: Math.max(col, 0), rowCount: 1, colCount: 1 };
        }

        // 范围格式 'a1:b2'
        const startPart = addr.substring(0, colonIndex);
        const endPart = addr.substring(colonIndex + 1);

        const startCol = this._colLetterToIndex(startPart.replace(/[0-9]/g, ''));
        const startRow = parseInt(startPart.replace(/[a-z]/g, '')) - 1;
        const endCol = this._colLetterToIndex(endPart.replace(/[0-9]/g, ''));
        const endRow = parseInt(endPart.replace(/[a-z]/g, '')) - 1;

        const rowCount = endRow - startRow + 1;
        const colCount = endCol - startCol + 1;

        return {
            startRow: Math.max(startRow, 0),
            startCol: Math.max(startCol, 0),
            rowCount: Math.max(rowCount, 1),
            colCount: Math.max(colCount, 1)
        };
    }

    /**
     * 将列字母转换为索引 (a->0, b->1, ..., z->25, aa->26, ...)
     * @param {string} letter - 列字母
     * @returns {number} 列索引
     */
    _colLetterToIndex(letter) {
        if (!letter) return 0;

        let result = 0;
        for (let i = 0; i < letter.length; i++) {
            result = result * 26 + (letter.charCodeAt(i) - 96); // 'a' = 97
        }
        return result - 1;
    }

    // ==================== 新增：分页函数 ====================

    /**
     * 按页数分页 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number} totalPages - 总页数（将数组分成多少页）
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    pageByCount(array2D = null, totalPages = null) {
        // 参数重载处理
        if (totalPages === null && typeof array2D === 'number') {
            totalPages = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D) || totalPages <= 0) {
                return [];
            }

            const result = [];
            const totalRows = array2D.length;
            // 计算每页应该有多少行
            const rowsPerPage = Math.ceil(totalRows / totalPages);

            for (let i = 0; i < totalRows; i += rowsPerPage) {
                const page = [];
                const end = Math.min(i + rowsPerPage, totalRows);

                for (let j = i; j < end; j++) {
                    const row = array2D[j];
                    if (Array.isArray(row)) {
                        page.push([...row]);
                    } else {
                        page.push(row);
                    }
                }

                result.push(page);
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data) || totalPages <= 0) {
            this.data = [];
            return this;
        }

        const result = [];
        const totalRows = this.data.length;
        // 计算每页应该有多少行
        const rowsPerPage = Math.ceil(totalRows / totalPages);

        for (let i = 0; i < totalRows; i += rowsPerPage) {
            const page = [];
            const end = Math.min(i + rowsPerPage, totalRows);

            for (let j = i; j < end; j++) {
                const row = this.data[j];
                if (Array.isArray(row)) {
                    page.push([...row]);
                } else {
                    page.push(row);
                }
            }

            result.push(page);
        }

        this.data = result;
        return this;
    }

    /**
     * 按行数分页 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number} rowsPerPage - 每页行数
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    pageByRows(array2D = null, rowsPerPage = null) {
        // 参数重载处理
        if (rowsPerPage === null && typeof array2D === 'number') {
            rowsPerPage = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D) || rowsPerPage <= 0) {
                return [];
            }

            const result = [];
            const totalRows = array2D.length;

            for (let i = 0; i < totalRows; i += rowsPerPage) {
                const page = [];
                const end = Math.min(i + rowsPerPage, totalRows);

                for (let j = i; j < end; j++) {
                    const row = array2D[j];
                    if (Array.isArray(row)) {
                        page.push([...row]);
                    } else {
                        page.push(row);
                    }
                }

                result.push(page);
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data) || rowsPerPage <= 0) {
            this.data = [];
            return this;
        }

        const result = [];
        const totalRows = this.data.length;

        for (let i = 0; i < totalRows; i += rowsPerPage) {
            const page = [];
            const end = Math.min(i + rowsPerPage, totalRows);

            for (let j = i; j < end; j++) {
                const row = this.data[j];
                if (Array.isArray(row)) {
                    page.push([...row]);
                } else {
                    page.push(row);
                }
            }

            result.push(page);
        }

        this.data = result;
        return this;
    }

    /**
     * 按下标分页 - 支持链式调用
     * @param {Array} pageIndices - 每页的下标数组
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    pageByIndexs(pageIndices) {
        if (this.isEmpty(this.data) || !Array.isArray(pageIndices)) {
            this.data = [];
            return this;
        }

        const result = [];

        for (const indices of pageIndices) {
            if (!Array.isArray(indices)) {
                continue;
            }

            const page = [];
            for (const index of indices) {
                if (index >= 0 && index < this.data.length) {
                    const row = this.data[index];
                    if (Array.isArray(row)) {
                        page.push([...row]);
                    } else {
                        page.push(row);
                    }
                }
            }

            if (page.length > 0) {
                result.push(page);
            }
        }

        this.data = result;
        return this;
    }

    // ==================== 新增：其他实用函数 ====================

    /**
     * 间隔取数 - 支持链式调用
     * @param {number} n - 间隔，每n个取一个
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    nth(n) {
        if (this.isEmpty(this.data) || n <= 0) {
            this.data = [];
            return this;
        }

        const result = [];

        for (let i = 0; i < this.data.length; i += n) {
            const row = this.data[i];
            if (Array.isArray(row)) {
                result.push([...row]);
            } else {
                result.push(row);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 补齐数组 - 支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number} colLength - 目标列长度
     * @param {number} rowLength - 目标行长度
     * @param {any} padValue - 补齐的值
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    pad(array2D = null, colLength = null, rowLength = null, padValue = null) {
        // 参数重载处理
        if (typeof array2D === 'number') {
            padValue = rowLength;
            rowLength = colLength;
            colLength = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                const result = [];
                for (let i = 0; i < rowLength; i++) {
                    result.push(new Array(colLength).fill(padValue));
                }
                return result;
            }

            const result = [];

            for (let i = 0; i < array2D.length; i++) {
                if (Array.isArray(array2D[i])) {
                    const row = [...array2D[i]];
                    // 补齐列
                    while (row.length < colLength) {
                        row.push(padValue);
                    }
                    result.push(row);
                } else {
                    result.push([array2D[i]]);
                }
            }

            // 补齐行
            while (result.length < rowLength) {
                result.push(new Array(colLength).fill(padValue));
            }

            return result;
        }

        // 链式模式：操作实例数据
        if (this.isEmpty(this.data)) {
            const result = [];
            for (let i = 0; i < rowLength; i++) {
                result.push(new Array(colLength).fill(padValue));
            }
            this.data = result;
            return this;
        }

        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                const row = [...this.data[i]];
                // 补齐列
                while (row.length < colLength) {
                    row.push(padValue);
                }
                result.push(row);
            } else {
                result.push([this.data[i]]);
            }
        }

        // 补齐行
        while (result.length < rowLength) {
            result.push(new Array(colLength).fill(padValue));
        }

        this.data = result;
        return this;
    }

    /**
     * 重复N次 - 支持链式调用
     * @param {number} count - 重复次数
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    repeat(count) {
        if (this.isEmpty(this.data) || count <= 0) {
            this.data = [];
            return this;
        }

        const result = [];

        for (let i = 0; i < count; i++) {
            for (const row of this.data) {
                if (Array.isArray(row)) {
                    result.push([...row]);
                } else {
                    result.push(row);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 跳过前N个 - 支持链式调用
     * @param {number} count - 跳过的数量
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    skip(count) {
        if (this.isEmpty(this.data) || count <= 0) {
            return this;
        }

        const result = [];
        for (let i = count; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 跳过前面连续满足条件的元素 - 支持链式调用
     * @param {Function} predicate - 条件函数
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    skipWhile(predicate) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        let skipCount = 0;

        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            let shouldSkip = false;

            if (Array.isArray(row)) {
                shouldSkip = row.every((val, j) => predicate(val, i, j));
            } else {
                shouldSkip = predicate(row, i, -1);
            }

            if (shouldSkip) {
                skipCount++;
            } else {
                break;
            }
        }

        return this.skip(skipCount);
    }

    /**
     * 超级透视（简化版）- 支持链式调用
     * @param {number|Function} rowField - 行字段索引或选择器
     * @param {number|Function} colField - 列字段索引或选择器
     * @param {number|Function} valueField - 值字段索引或选择器
     * @param {Function} aggregator - 聚合函数，默认为求和
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    superPivot(rowField, colField, valueField, aggregator = (acc, val) => acc + val) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        // 构建行和列的映射
        const rowValues = new Set();
        const colValues = new Set();
        const pivotData = {};

        // 第一遍：收集所有行和列的值
        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];

            let rowKey, colKey, value;

            // 获取行键
            if (typeof rowField === 'function') {
                rowKey = String(rowField(row, i));
            } else {
                rowKey = String(Array.isArray(row) && row.length > rowField ? row[rowField] : row);
            }

            // 获取列键
            if (typeof colField === 'function') {
                colKey = String(colField(row, i));
            } else {
                colKey = String(Array.isArray(row) && row.length > colField ? row[colField] : row);
            }

            // 获取值
            if (typeof valueField === 'function') {
                value = valueField(row, i);
            } else {
                value = Array.isArray(row) && row.length > valueField ? row[valueField] : row;
            }

            rowValues.add(rowKey);
            colValues.add(colKey);

            // 初始化数据结构
            if (!pivotData[rowKey]) {
                pivotData[rowKey] = {};
            }

            // 聚合值
            if (pivotData[rowKey][colKey] === undefined) {
                pivotData[rowKey][colKey] = value;
            } else {
                pivotData[rowKey][colKey] = aggregator(pivotData[rowKey][colKey], value);
            }
        }

        // 构建结果表
        const result = [];
        const colArray = Array.from(colValues).sort();

        // 表头
        const header = ['行\\列', ...colArray];
        result.push(header);

        // 数据行
        const rowArray = Array.from(rowValues).sort();
        for (const rowKey of rowArray) {
            const rowData = [rowKey];
            for (const colKey of colArray) {
                const value = pivotData[rowKey] && pivotData[rowKey][colKey] !== undefined ? pivotData[rowKey][colKey] : 0;
                rowData.push(value);
            }
            result.push(rowData);
        }

        this.data = result;
        return this;
    }

    // ==================== 补充：查找所有行下标 ====================

    /**
     * 查找所有符合条件的行下标
     * @param {Array} array2D - 二维数组
     * @param {Function|any} predicate - 判断函数或要查找的值
     * @returns {Array} 所有符合条件的行下标
     */
    findRowsIndex(array2D, predicate) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = [];
        const isFunction = typeof predicate === 'function';

        for (let i = 0; i < array2D.length; i++) {
            const row = array2D[i];
            let match = false;

            if (isFunction) {
                // 如果是函数，检查整行是否满足条件
                match = predicate(row, i);
            } else {
                // 如果是值，检查行中是否包含该值
                if (Array.isArray(row)) {
                    match = row.includes(predicate);
                } else {
                    match = row === predicate;
                }
            }

            if (match) {
                result.push(i);
            }
        }

        return result;
    }

    /**
     * 查找所有符合条件的列下标
     * @param {Array} array2D - 二维数组
     * @param {Function|any} predicate - 判断函数或要查找的值
     * @returns {Array} 所有符合条件的列下标（适用于每行的列位置）
     */
    findColsIndex(array2D, predicate) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        // 假设所有行具有相同的列数
        const firstRow = array2D[0];
        if (!Array.isArray(firstRow)) {
            return [];
        }

        const numCols = firstRow.length;
        const result = [];
        const isFunction = typeof predicate === 'function';

        for (let j = 0; j < numCols; j++) {
            let match = false;

            if (isFunction) {
                // 检查该列是否所有行都满足条件
                match = array2D.every((row, i) => {
                    if (!Array.isArray(row)) {
                        return false;
                    }
                    return j < row.length && predicate(row[j], i, j);
                });
            } else {
                // 检查该列在所有行中是否包含该值
                match = array2D.some((row) => {
                    if (!Array.isArray(row)) {
                        return false;
                    }
                    return j < row.length && row[j] === predicate;
                });
            }

            if (match) {
                result.push(j);
            }
        }

        return result;
    }

    // ==================== 补充：检查函数 ====================

    /**
     * 检查是否有元素/行满足条件 - 支持按行和逐元素检查
     * @param {Array} array2D - 二维数组
     * @param {Function} predicate - 判断函数
     *   - 按行检查: (row, rowIndex) => boolean
     *   - 逐元素检查: (element, rowIndex, colIndex) => boolean
     * @param {boolean} byRow - 是否按行检查（默认true，推荐使用）
     * @returns {boolean} 是否有满足条件的元素/行
     */
    some(array2D, predicate, byRow = true) {
        if (this.isEmpty(array2D)) {
            return false;
        }

        if (byRow) {
            // 按行检查模式（推荐）
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (predicate(row, i)) {
                    return true;
                }
            }
        } else {
            // 逐元素检查模式（向后兼容）
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (Array.isArray(row)) {
                    for (let j = 0; j < row.length; j++) {
                        if (predicate(row[j], i, j)) {
                            return true;
                        }
                    }
                } else {
                    if (predicate(row, i, -1)) {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /**
     * 检查是否所有元素/行都满足条件 - 支持按行和逐元素检查
     * @param {Array} array2D - 二维数组
     * @param {Function} predicate - 判断函数
     *   - 按行检查: (row, rowIndex) => boolean
     *   - 逐元素检查: (element, rowIndex, colIndex) => boolean
     * @param {boolean} byRow - 是否按行检查（默认true，推荐使用）
     * @returns {boolean} 是否所有元素/行都满足条件
     */
    every(array2D, predicate, byRow = true) {
        if (this.isEmpty(array2D)) {
            return false;
        }

        if (byRow) {
            // 按行检查模式（推荐）
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (!predicate(row, i)) {
                    return false;
                }
            }
        } else {
            // 逐元素检查模式（向后兼容）
            for (let i = 0; i < array2D.length; i++) {
                const row = array2D[i];
                if (Array.isArray(row)) {
                    for (let j = 0; j < row.length; j++) {
                        if (!predicate(row[j], i, j)) {
                            return false;
                        }
                    }
                } else {
                    if (!predicate(row, i, -1)) {
                        return false;
                    }
                }
            }
        }

        return true;
    }

    /**
     * 倒序聚合
     * @param {Array} array2D - 二维数组
     * @param {Function} reducer - 聚合函数
     * @param {any} initialValue - 初始值
     * @returns {any} 聚合结果
     */
    reduceRight(array2D, reducer, initialValue) {
        if (this.isEmpty(array2D)) {
            return initialValue;
        }

        // 扁平化数组并反转
        const flatArray = [];
        for (let i = array2D.length - 1; i >= 0; i--) {
            const row = array2D[i];
            if (Array.isArray(row)) {
                for (let j = row.length - 1; j >= 0; j--) {
                    flatArray.push(row[j]);
                }
            } else {
                flatArray.push(row);
            }
        }

        // 从右到左聚合
        let accumulator = initialValue;
        let startIndex = 0;

        if (initialValue === undefined) {
            accumulator = flatArray[0];
            startIndex = 1;
        }

        for (let i = startIndex; i < flatArray.length; i++) {
            accumulator = reducer(accumulator, flatArray[i], i);
        }

        return accumulator;
    }

    // ==================== 补充：数组栈操作函数 ====================

    /**
     * 删除第一个元素
     * @param {Array} array2D - 二维数组
     * @returns {any} 被删除的元素
     */
    shift(array2D) {
        if (this.isEmpty(array2D)) {
            return undefined;
        }

        const removed = array2D.shift();
        return removed;
    }

    /**
     * 在开头添加元素
     * @param {Array} array2D - 二维数组
     * @param {...any} items - 要添加的元素
     * @returns {number} 新的数组长度
     */
    unshift(array2D, ...items) {
        if (!Array.isArray(array2D)) {
            return 0;
        }

        return array2D.unshift(...items);
    }

    /**
     * 尾部弹出一项
     * @param {Array} array2D - 二维数组
     * @returns {any} 被弹出的元素
     */
    pop(array2D) {
        if (!Array.isArray(array2D) || array2D.length === 0) {
            return undefined;
        }

        return array2D.pop();
    }

    /**
     * 追加一项
     * @param {Array} array2D - 二维数组
     * @param {...any} items - 要追加的元素
     * @returns {number} 新的数组长度
     */
    push(array2D, ...items) {
        if (!Array.isArray(array2D)) {
            return 0;
        }

        return array2D.push(...items);
    }

    // ==================== 补充：位置查找函数 ====================

    /**
     * 值位置（indexOf的别名）
     * @param {Array} array2D - 二维数组
     * @param {any} value - 要查找的值
     * @param {number} fromIndex - 开始查找的位置
     * @returns {number} 第一次出现的位置，未找到返回-1
     */
    indexOf(array2D, value, fromIndex = 0) {
        const index = this.findIndex(array2D, value, fromIndex);
        return index.row !== -1 ? index.row : -1;
    }

    /**
     * 从后往前值位置
     * @param {Array} array2D - 二维数组
     * @param {any} value - 要查找的值
     * @param {number} fromIndex - 开始查找的位置（从后往前）
     * @returns {number} 最后一次出现的位置，未找到返回-1
     */
    lastIndexOf(array2D, value, fromIndex = undefined) {
        if (this.isEmpty(array2D)) {
            return -1;
        }

        const startIndex = fromIndex !== undefined ? Math.min(fromIndex, array2D.length - 1) : array2D.length - 1;

        for (let i = startIndex; i >= 0; i--) {
            const row = array2D[i];
            let found = false;

            if (Array.isArray(row)) {
                const colIndex = row.lastIndexOf(value);
                if (colIndex !== -1) {
                    return i;
                }
            } else if (row === value) {
                return i;
            }
        }

        return -1;
    }

    // ==================== 补充：特殊工具函数 ====================

    /**
     * 复制到指定位置 - 支持链式调用
     * @param {number} targetRow - 目标行
     * @param {number} targetCol - 目标列
     * @param {number} startRow - 源开始行
     * @param {number} startCol - 源开始列
     * @param {number} endRow - 源结束行（可选）
     * @param {number} endCol - 源结束列（可选）
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    copyWithin(targetRow, targetCol, startRow, startCol, endRow = undefined, endCol = undefined) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = this.copyStatic(this.data);
        const actualEndRow = endRow !== undefined ? endRow : result.length;
        const actualEndCol = endCol !== undefined ? endCol : (Array.isArray(result[0]) ? result[0].length : 1);

        let sourceRow = startRow;
        let sourceCol = startCol;
        let targetRowCurrent = targetRow;
        let targetColCurrent = targetCol;

        while (sourceRow < actualEndRow && sourceRow < result.length) {
            while (sourceCol < actualEndCol && sourceCol < (Array.isArray(result[sourceRow]) ? result[sourceRow].length : 1)) {
                if (targetRowCurrent < result.length && targetColCurrent < (Array.isArray(result[targetRowCurrent]) ? result[targetRowCurrent].length : 1)) {
                    if (Array.isArray(result[sourceRow]) && Array.isArray(result[targetRowCurrent])) {
                        result[targetRowCurrent][targetColCurrent] = result[sourceRow][sourceCol];
                    }
                }

                sourceCol++;
                targetColCurrent++;
            }

            sourceRow++;
            targetRowCurrent++;
            sourceCol = startCol;
            targetColCurrent = targetCol;
        }

        this.data = result;
        return this;
    }

    /**
     * 生成下标数组
     * @param {number} length - 数组长度
     * @param {number} start - 起始值，默认为0
     * @returns {Array} 下标数组
     */
    getIndexs(length, start = 0) {
        if (length <= 0) {
            return [];
        }

        const result = [];
        for (let i = 0; i < length; i++) {
            result.push(start + i);
        }

        return result;
    }

    /**
     * 下标数组（生成从0开始的下标）- 支持链式调用
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    indexArray() {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        this.data = this.getIndexs(this.data.length, 0);
        return this;
    }

    /**
     * 插入行号 - 支持链式调用
     * @param {number} colIndex - 插入位置，默认为0（最前面）
     * @param {number} startNumber - 起始行号，默认为1
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    insertRowNum(colIndex = 0, startNumber = 1) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = this.copyStatic(this.data);

        for (let i = 0; i < result.length; i++) {
            if (Array.isArray(result[i])) {
                result[i].splice(colIndex, 0, startNumber + i);
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 一对多连接（左全连接）- 支持链式调用
     * @param {Array} rightArray - 右表
     * @param {Function|number} leftKey - 左表连接键选择器或列索引
     * @param {Function|number} rightKey - 右表连接键选择器或列索引
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    leftFulljoin(rightArray, leftKey, rightKey) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        const rightMap = {};

        // 构建右表映射（允许多个值）
        if (!this.isEmpty(rightArray)) {
            for (let i = 0; i < rightArray.length; i++) {
                const row = rightArray[i];
                let key;

                if (typeof rightKey === 'function') {
                    key = rightKey(row, i);
                } else if (typeof rightKey === 'number') {
                    key = Array.isArray(row) && row.length > rightKey ? row[rightKey] : row;
                } else {
                    key = JSON.stringify(row);
                }

                const keyStr = String(key);
                if (!rightMap[keyStr]) {
                    rightMap[keyStr] = [];
                }
                rightMap[keyStr].push(row);
            }
        }

        // 执行一对多连接
        for (let i = 0; i < this.data.length; i++) {
            const leftRow = this.data[i];
            let key;

            if (typeof leftKey === 'function') {
                key = leftKey(leftRow, i);
            } else if (typeof leftKey === 'number') {
                key = Array.isArray(leftRow) && leftRow.length > leftKey ? leftRow[leftKey] : leftRow;
            } else {
                key = JSON.stringify(leftRow);
            }

            const keyStr = String(key);
            const rightRows = rightMap[keyStr];

            if (rightRows && rightRows.length > 0) {
                // 为每个匹配的右表行创建一行结果
                for (const rightRow of rightRows) {
                    const leftRowArr = Array.isArray(leftRow) ? leftRow : [leftRow];
                    const rightRowArr = Array.isArray(rightRow) ? rightRow : [rightRow];
                    result.push([...leftRowArr, ...rightRowArr]);
                }
            } else {
                // 没有匹配，创建带null的行
                const leftRowArr = Array.isArray(leftRow) ? leftRow : [leftRow];
                if (!this.isEmpty(rightArray)) {
                    const rightPlaceholder = Array.isArray(rightArray[0]) ? new Array(rightArray[0].length).fill(null) : [null];
                    result.push([...leftRowArr, ...rightPlaceholder]);
                } else {
                    result.push(leftRowArr);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 返回结果（辅助函数）
     * @param {Array} array2D - 二维数组
     * @returns {Array} 数组本身
     */
    res(array2D) {
        return array2D;
    }

    // ==================== 新增缺失函数 ====================

    /**
     * 转置矩阵 - 行列互换，支持传入数组参数或链式调用
     * @param {Array} array2D - 要转置的数组（可选）
     * @returns {Array|clsArray2D} 如果传入数组返回转置后的数组，否则返回实例
     * @example
     * const arr = new Array2D();
     * const transposed = arr.transpose([[1, 2, 3], [4, 5, 6]]); // [[1, 4], [2, 5], [3, 6]]
     * // 或者链式调用
     * const arr2 = new Array2D([[1, 2, 3], [4, 5, 6]]);
     * arr2.transpose();
     * console.log(arr2.val()); // [[1, 4], [2, 5], [3, 6]]
     */
    transpose(array2D = null) {
        // 如果传入了数组参数，返回转置后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const maxCols = Math.max(...array2D.map(row => Array.isArray(row) ? row.length : 1));
            const result = [];

            for (let j = 0; j < maxCols; j++) {
                const newRow = [];
                for (let i = 0; i < array2D.length; i++) {
                    newRow.push(Array.isArray(array2D[i]) ? array2D[i][j] : array2D[i]);
                }
                result.push(newRow);
            }

            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const rows = this.data.length;
        const cols = this.data[0].length;
        const result = [];

        for (let j = 0; j < cols; j++) {
            result[j] = [];
            for (let i = 0; i < rows; i++) {
                result[j][i] = this.data[i][j];
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 转换为矩阵格式 - 确保所有行长度一致，支持链式调用
     * @param {*} fillValue - 填充值
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    toMatrix(fillValue = null) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const maxCols = Math.max(...this.data.map(row => Array.isArray(row) ? row.length : 1));
        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            if (Array.isArray(row)) {
                result[i] = [...row];
                while (result[i].length < maxCols) {
                    result[i].push(fillValue);
                }
            } else {
                result[i] = [row];
                while (result[i].length < maxCols) {
                    result[i].push(fillValue);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 矩阵分布 - 按指定行列数分布数据，支持链式调用
     * @param {number} rows - 行数
     * @param {number} cols - 列数
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    getMatrix(rows, cols) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        // 先扁平化数组
        const flatArray = [];
        const flatten = (arr, depth) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i]) && depth < Infinity) {
                    flatten(arr[i], depth + 1);
                } else {
                    flatArray.push(arr[i]);
                }
            }
        };
        flatten(this.data, 0);

        const result = [];
        let index = 0;

        for (let i = 0; i < rows; i++) {
            result[i] = [];
            for (let j = 0; j < cols; j++) {
                result[i][j] = index < flatArray.length ? flatArray[index] : null;
                index++;
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 矩阵运算 - 对指定区域应用函数，支持链式调用
     * @param {number} startRow - 起始行
     * @param {number} startCol - 起始列
     * @param {number} endRow - 结束行
     * @param {number} endCol - 结束列
     * @param {Function} operation - 运算函数
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    rangeMatrix(startRow, startCol, endRow, endCol, operation) {
        if (this.isEmpty(this.data)) {
            return this;
        }

        const result = [];
        for (let i = 0; i < this.data.length; i++) {
            if (Array.isArray(this.data[i])) {
                result.push([...this.data[i]]);
            } else {
                result.push(this.data[i]);
            }
        }

        for (let i = startRow; i <= endRow && i < result.length; i++) {
            if (result[i] && Array.isArray(result[i])) {
                for (let j = startCol; j <= endCol && j < result[i].length; j++) {
                    result[i][j] = operation(result[i][j], i, j);
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 重设大小 - 调整数组大小，支持链式调用和静态调用
     * @param {Array} array2D - 二维数组（静态模式）或不传（链式模式）
     * @param {number} newRows - 新行数
     * @param {number} newCols - 新列数
     * @param {*} fillValue - 填充值
     * @returns {Array|clsArray2D} 静态模式返回新数组，链式模式返回实例
     */
    resize(array2D = null, newRows = null, newCols = null, fillValue = null) {
        // 参数重载处理
        if (typeof array2D === 'number') {
            fillValue = newCols;
            newCols = newRows;
            newRows = array2D;
            array2D = null;
        }

        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (array2D !== null) {
            const result = [];

            for (let i = 0; i < newRows; i++) {
                result[i] = [];
                for (let j = 0; j < newCols; j++) {
                    if (i < array2D.length && array2D[i] && j < array2D[i].length) {
                        result[i][j] = array2D[i][j];
                    } else {
                        result[i][j] = fillValue;
                    }
                }
            }

            return result;
        }

        // 链式模式：操作实例数据
        const result = [];

        for (let i = 0; i < newRows; i++) {
            result[i] = [];
            for (let j = 0; j < newCols; j++) {
                if (i < this.data.length && this.data[i] && j < this.data[i].length) {
                    result[i][j] = this.data[i][j];
                } else {
                    result[i][j] = fillValue;
                }
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 按规则升序排序 - 支持链式调用
     * @param {Function|number} keySelector - 键选择器或列索引
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    sortBy(keySelector) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [...this.data];

        result.sort((a, b) => {
            let keyA, keyB;

            if (typeof keySelector === 'function') {
                keyA = keySelector(a);
                keyB = keySelector(b);
            } else if (typeof keySelector === 'number') {
                keyA = a[keySelector];
                keyB = b[keySelector];
            } else {
                keyA = a;
                keyB = b;
            }

            if (keyA < keyB) return -1;
            if (keyA > keyB) return 1;
            return 0;
        });

        this.data = result;
        return this;
    }

    /**
     * 多列排序 - 按多列排序 - 支持链式调用
     * @param {Array} colIndices - 列索引数组
     * @param {Array} orders - 排序顺序数组 (true升序/false降序)
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    sortByCols(colIndices, orders = []) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [...this.data];

        result.sort((a, b) => {
            for (let i = 0; i < colIndices.length; i++) {
                const colIndex = colIndices[i];
                const ascending = orders[i] !== false;

                const valA = a[colIndex];
                const valB = b[colIndex];

                if (valA < valB) return ascending ? -1 : 1;
                if (valA > valB) return ascending ? 1 : -1;
            }
            return 0;
        });

        this.data = result;
        return this;
    }

    /**
     * 按规则降序排序 - 支持链式调用
     * @param {Function|number} keySelector - 键选择器或列索引
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    sortByDesc(keySelector) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [...this.data];

        result.sort((a, b) => {
            let keyA, keyB;

            if (typeof keySelector === 'function') {
                keyA = keySelector(a);
                keyB = keySelector(b);
            } else if (typeof keySelector === 'number') {
                keyA = a[keySelector];
                keyB = b[keySelector];
            } else {
                keyA = a;
                keyB = b;
            }

            if (keyA > keyB) return -1;
            if (keyA < keyB) return 1;
            return 0;
        });

        this.data = result;
        return this;
    }

    /**
     * 自定义排序 - 按指定列表的顺序排序 - 支持链式调用
     * @param {number} colIndex - 列索引
     * @param {Array} orderList - 排序列表
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    sortByList(colIndex, orderList) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [...this.data];

        result.sort((a, b) => {
            const indexA = orderList.indexOf(a[colIndex]);
            const indexB = orderList.indexOf(b[colIndex]);

            const posA = indexA === -1 ? 999 : indexA;
            const posB = indexB === -1 ? 999 : indexB;

            return posA - posB;
        });

        this.data = result;
        return this;
    }

    /**
     * 降序排序 - 支持传入数组参数或链式调用
     * @param {Array} array2D - 要排序的数组（可选）
     * @returns {Array|clsArray2D} 如果传入数组返回排序后的数组，否则返回实例
     */
    sortDesc(array2D = null) {
        // 如果传入了数组参数，返回排序后的新数组（符合官方文档规范）
        if (array2D !== null) {
            if (this.isEmpty(array2D)) {
                return [];
            }

            const result = this.copy(array2D);
            result.sort((a, b) => {
                const aVal = Array.isArray(a) ? a[0] : a;
                const bVal = Array.isArray(b) ? b[0] : b;
                if (aVal > bVal) return -1;
                if (aVal < bVal) return 1;
                return 0;
            });
            return result;
        }

        // 否则使用实例的 data（保持向后兼容）
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = this.copyStatic(this.data);
        result.sort((a, b) => {
            const aVal = Array.isArray(a) ? a[0] : a;
            const bVal = Array.isArray(b) ? b[0] : b;
            if (aVal > bVal) return -1;
            if (aVal < bVal) return 1;
            return 0;
        });

        this.data = result;
        return this;
    }

    /**
     * 行切片删除行 - 删除指定位置的行 - 支持链式调用
     * @param {number} start - 起始位置
     * @param {number} deleteCount - 删除数量
     * @param {Array} items - 要插入的项
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    splice(start, deleteCount = 0, ...items) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [...this.data];

        if (deleteCount > 0) {
            result.splice(start, deleteCount);
        }

        if (items.length > 0) {
            result.splice(start, 0, ...items);
        }

        this.data = result;
        return this;
    }

    /**
     * 求和 - 计算二维数组中所有数字的总和
     * @param {Array<any>} array2D - 二维数组
     * @param {any} colSelector - 列选择器（可选）'f1'表示第1列，或回调函数
     * @returns {number} 所有数字的总和
     * @example
     * const arr = new Array2D();
     * const data = [[1, 2, 3], [4, 5, 6]];
     * console.log(arr.sum(data)); // 21 (所有元素)
     * console.log(arr.sum(data, 'f1')); // 5 (第一列: 1+4)
     */
    sum(array2D, colSelector = null) {
        if (this.isEmpty(array2D)) {
            return 0;
        }

        // 如果没有列选择器，求所有元素的和
        if (!colSelector) {
            const flatArray = this.flat(array2D, Infinity);
            return flatArray.reduce((total, val) => {
                const num = Number(val);
                return total + (isNaN(num) ? 0 : num);
            }, 0);
        }

        // 有列选择器时，按指定列求和
        const colIndex = this._parseColSelector(colSelector);
        let total = 0;

        for (let i = 0; i < array2D.length; i++) {
            if (array2D[i] && array2D[i][colIndex] !== undefined) {
                const val = Number(array2D[i][colIndex]);
                if (!isNaN(val)) {
                    total += val;
                }
            }
        }

        return total;
    }

    /**
     * 超级查找 - 多条件查找
     * @param {Array} array2D - 二维数组
     * @param {Object} conditions - 查找条件对象
     * @returns {Array} 匹配的行
     */
    superLookup(array2D, conditions) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = [];

        for (let i = 0; i < array2D.length; i++) {
            const row = array2D[i];
            let match = true;

            for (const key in conditions) {
                const colIndex = parseInt(key);
                if (row[colIndex] !== conditions[key]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                result.push(row);
            }
        }

        return result;
    }

    /**
     * 取前N个 - 支持链式调用
     * @param {number} n - 取的数量
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    take(n) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const flatArray = [];
        const flatten = (arr) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i])) {
                    flatten(arr[i]);
                } else {
                    flatArray.push(arr[i]);
                }
            }
        };
        flatten(this.data);

        this.data = flatArray.slice(0, n);
        return this;
    }

    /**
     * 取前面连续满足条件的元素 - 支持链式调用
     * @param {Function} predicate - 判断函数
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    takeWhile(predicate) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const flatArray = [];
        const flatten = (arr) => {
            for (let i = 0; i < arr.length; i++) {
                if (Array.isArray(arr[i])) {
                    flatten(arr[i]);
                } else {
                    flatArray.push(arr[i]);
                }
            }
        };
        flatten(this.data);

        const result = [];

        for (let i = 0; i < flatArray.length; i++) {
            if (predicate(flatArray[i], i)) {
                result.push(flatArray[i]);
            } else {
                break;
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 写入单元格 (JSA环境)
     * @param {Array} array2D - 二维数组
     * @param {string} rangeAddress - 单元格地址
     * @returns {Array} 原数组
     */
    toRange(array2D, rangeAddress) {
        // 在JSA环境中，这里会将数组写入到指定的单元格范围
        // 在普通JS环境中，仅返回原数组
        if (typeof Application !== 'undefined' && Application.ActiveSheet) {
            try {
                const range = Application.ActiveSheet.Range(rangeAddress);
                if (range) {
                    range.Value2 = array2D;
                }
            } catch (e) {
                console.warn("写入单元格失败:", e.message);
            }
        }
        return array2D;
    }

    /**
     * 转字符串
     * @param {Array} array2D - 二维数组
     * @param {string} rowSeparator - 行分隔符
     * @param {string} colSeparator - 列分隔符
     * @returns {string} 字符串
     */
    toString(array2D, rowSeparator = '\n', colSeparator = ',') {
        if (this.isEmpty(array2D)) {
            return '';
        }

        return array2D.map(row => {
            if (Array.isArray(row)) {
                return row.join(colSeparator);
            }
            return String(row);
        }).join(rowSeparator);
    }

    /**
     * 去重并集 - 合并数组并去重 - 支持链式调用和静态调用
     * @param {Array} array2D1 - 第一个二维数组（静态模式）或不传（链式模式）
     * @param {Array} array2D2 - 第二个二维数组（静态模式）或传入数组（链式模式）
     * @param {Function} conditionCallback - 可选的去重条件回调函数
     * @returns {clsArray2D|Array} 静态模式返回新数组，链式模式返回当前实例
     */
    union(array2D1 = null, array2D2 = null, conditionCallback = null) {
        // 参数重载处理 - 链式模式: union(array2D2, conditionCallback)
        // 静态模式: union(array2D1, array2D2, conditionCallback)
        let arr1, arr2, callback;

        if (array2D1 === null) {
            // 链式模式，无参数
            this.data = [];
            return this;
        }

        if (Array.isArray(array2D1)) {
            // 静态模式
            arr1 = array2D1;
            arr2 = array2D2;
            callback = conditionCallback;
        } else {
            // 链式模式: array2D1 是 array2D2, array2D2 是 callback
            arr1 = this.data;
            arr2 = array2D1;
            callback = array2D2;
        }

        // 如果没有第二个数组，返回第一个数组的副本
        if (!arr2 || !Array.isArray(arr2)) {
            if (Array.isArray(arr1)) {
                const result = arr1.map(row => Array.isArray(row) ? [...row] : row);
                if (callback) {
                    // 按条件去重
                    const seen = new Set();
                    const filtered = [];
                    for (const row of result) {
                        const key = String(callback(row));
                        if (!seen.has(key)) {
                            seen.add(key);
                            filtered.push(row);
                        }
                    }
                    if (Array.isArray(arr1)) {
                        return filtered;
                    }
                    this.data = filtered;
                    return this;
                }
                if (Array.isArray(arr1)) {
                    return result;
                }
                this.data = result;
                return this;
            }
            if (Array.isArray(arr1)) {
                return [];
            }
            this.data = [];
            return this;
        }

        // 合并两个数组
        const combined = [...arr1, ...arr2];

        if (callback) {
            // 按条件去重
            const seen = new Set();
            const result = [];
            for (const row of combined) {
                const key = String(callback(row));
                if (!seen.has(key)) {
                    seen.add(key);
                    result.push(Array.isArray(row) ? [...row] : row);
                }
            }
            if (Array.isArray(arr1)) {
                return result;
            }
            this.data = result;
            return this;
        } else {
            // 按整行去重
            const seen = new Set();
            const result = [];
            for (const row of combined) {
                const key = JSON.stringify(row);
                if (!seen.has(key)) {
                    seen.add(key);
                    result.push(Array.isArray(row) ? [...row] : row);
                }
            }
            if (Array.isArray(arr1)) {
                return result;
            }
            this.data = result;
            return this;
        }
    }

    /**
     * 左右连接 - zip操作 - 支持链式调用和静态调用
     * @param {...Array} arrays - 多个数组（静态模式）或不传（链式模式）
     * @returns {clsArray2D|Array} 静态模式返回新数组，链式模式返回当前实例
     */
    zip(...arrays) {
        // 静态模式：如果传入了数组参数，返回新数组（符合官方文档规范）
        if (arrays.length > 0) {
            // 检查第一个参数是否是数组（用于区分静态模式和链式模式）
            const firstIsArray = arrays.length > 0 && Array.isArray(arrays[0]);

            if (firstIsArray) {
                // 静态模式：直接合并传入的数组
                const maxLength = Math.max(...arrays.map(arr => arr ? arr.length : 0));
                const result = [];

                for (let i = 0; i < maxLength; i++) {
                    const newRow = [];
                    for (let a = 0; a < arrays.length; a++) {
                        const arr = arrays[a];
                        if (arr && arr[i]) {
                            if (Array.isArray(arr[i])) {
                                newRow.push(...arr[i]);
                            } else {
                                newRow.push(arr[i]);
                            }
                        }
                    }
                    result.push(newRow);
                }
                return result;
            }
        }

        // 链式模式：合并实例数据和传入的数组
        if (arrays.length === 0) {
            this.data = [];
            return this;
        }

        const allArrays = [this.data, ...arrays];
        const maxLength = Math.max(...allArrays.map(arr => arr.length));
        const result = [];

        for (let i = 0; i < maxLength; i++) {
            result[i] = [];
            for (let j = 0; j < allArrays.length; j++) {
                result[i][j] = allArrays[j][i];
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 检查是否为错误值
     * @param {*} value - 要检查的值
     * @returns {boolean} 是否为错误值
     */
    isError(value) {
        return value instanceof Error ||
               (typeof value === 'object' && value !== null && value.message) ||
               ['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NUM!', '#NAME?', '#NULL!'].includes(value);
    }

    /**
     * 处理空值 - 替换null和undefined - 支持链式调用
     * @param {*} replacement - 替换值
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    noNull(replacement = '') {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];

        for (let i = 0; i < this.data.length; i++) {
            const row = this.data[i];
            if (Array.isArray(row)) {
                result[i] = row.map(val =>
                    (val === null || val === undefined) ? replacement : val
                );
            } else {
                result[i] = (row === null || row === undefined) ? replacement : row;
            }
        }

        this.data = result;
        return this;
    }

    /**
     * 解析函数表达式 - 简单的Lambda表达式解析
     * @param {string} expression - 表达式字符串
     * @returns {Function} 解析后的函数
     */
    parseLambda(expression) {
        // 支持格式: "x => x > 5" 或 "x,y => x + y"
        const match = expression.match(/^([a-zA-Z_][a-zA-Z0-9_,\s]*)\s*=>\s*(.+)$/);

        if (match) {
            const params = match[1].split(',').map(p => p.trim());
            const body = match[2];

            try {
                return new Function(...params, `return ${body};`);
            } catch (e) {
                console.error("解析Lambda表达式失败:", e.message);
                return null;
            }
        }

        // 如果不是Lambda表达式，尝试直接创建函数
        try {
            return new Function('x', `return ${expression};`);
        } catch (e) {
            console.error("解析表达式失败:", e.message);
            return null;
        }
    }

    // ==================== 中文函数别名 ====================

    /**
     * 中文函数别名
     * 为了保持与JSA880框架的兼容性
     */

    // 基本信息函数
    z版本 = () => this.version();
    z空结果 = (arr) => this.isEmpty(arr);
    z数量 = (arr) => this.count(arr);

    // 数组操作函数
    z克隆 = (arr) => this.copy(arr);
    z批量填充 = (arr, val, start, end) => this.fill(arr, val, start, end);
    z补齐空位 = (arr, defVal) => this.fillBlank(arr, defVal);
    z降维 = (arr, depth) => this.flat(arr, depth);
    z反转 = (arr) => this.reverse(arr);

    // 查找和筛选函数
    z查找单个 = (arr, predicate) => this.find(arr, predicate);
    z查找元素下标 = (arr, value, fromIndex) => this.findIndex(arr, value, fromIndex);
    z查找所有下标 = (arr, value) => this.findAllIndex(arr, value);
    z筛选 = (arr, predicate) => this.filter(arr, predicate);
    z是否包含值 = (arr, value) => this.includes(arr, value);

    // 统计计算函数
    z平均值 = (arr, colSelector) => this.average(arr, colSelector);
    z最大值 = (arr, colSelector) => this.max(arr, colSelector);
    z最小值 = (arr, colSelector) => this.min(arr, colSelector);
    z中位数 = (arr, colSelector) => this.median(arr, colSelector);

    // 数组变换函数
    z映射生成 = (arr, callback) => this.map(arr, callback);
    z遍历执行 = (arr, callback) => this.forEach(arr, callback);
    z倒序遍历执行 = (arr, callback) => this.forEachRev(arr, callback);
    z聚合 = (arr, reducer, initialValue) => this.reduce(arr, reducer, initialValue);

    // 数组连接函数
    z上下连接 = (arr1, arr2) => this.concat(arr1, arr2);
    z文本连接 = (arr, separator) => this.join(arr, separator);

    // 数组操作函数
    z第一个 = (arr) => this.first(arr);
    z最后一个 = (arr) => this.last(arr);
    z行切片 = (arr, start, end) => this.slice(arr, start, end);
    z去重 = (arr) => this.distinct(arr);
    z升序排序 = (arr, compareFunction) => this.sort(arr, compareFunction);
    z随机打乱 = (arr) => this.shuffle(arr);
    z随机一项 = (arr) => this.random(arr);

    // ==================== 新增函数的中文别名 ====================

    // 笛卡尔积函数
    z笛卡尔积 = (arr1, arr2) => this.crossjoin(arr1, arr2);

    // 分组函数
    z分组 = (arr, keySelector) => this.groupBy(arr, keySelector);
    z分组汇总 = (arr, keySelector, valueSelector, aggregator) => this.groupInto(arr, keySelector, valueSelector, aggregator);
    z分组汇总到字典 = (arr, keySelector, valueSelector, aggregator) => this.groupIntoMap(arr, keySelector, valueSelector, aggregator);

    // 连接函数
    z左连接 = (leftArr, rightArr, leftKey, rightKey) => this.leftjoin(leftArr, rightArr, leftKey, rightKey);
    z左右全连接 = (leftArr, rightArr, leftKey, rightKey) => this.fulljoin(leftArr, rightArr, leftKey, rightKey);

    // 集合操作函数
    z排除 = (arr, excludeArr) => this.except(arr, excludeArr);
    z取交集 = (arr1, arr2) => this.intersect(arr1, arr2);

    // 排名函数
    z排名 = (arr, valueSelector, ascending) => this.rank(arr, valueSelector, ascending);
    z分组排名 = (arr, groupSelector, valueSelector, ascending) => this.rankGroup(arr, groupSelector, valueSelector, ascending);

    // 批量操作函数
    z批量插入列 = (arr, index, value, count) => this.insertCols(arr, index, value, count);
    z批量插入行 = (arr, index, value, count) => this.insertRows(arr, index, value, count);
    z批量删除列 = (arr, start, count) => this.deleteCols(arr, start, count);
    z批量删除行 = (arr, start, count) => this.deleteRows(arr, start, count);

    // 选择函数
    z选择列 = (arr, colIndices) => this.selectCols(arr, colIndices);
    z选择行 = (arr, rowIndices) => this.selectRows(arr, rowIndices);

    // 按范围操作函数
    z按范围遍历 = (arr, startRow, endRow, startCol, endCol, callback) => this.rangeForEach(arr, startRow, endRow, startCol, endCol, callback);
    z局部映射 = (arr, startRow, endRow, startCol, endCol, callback) => this.rangeMap(arr, startRow, endRow, startCol, endCol, callback);
    z按范围选择 = (arr, address) => this.rangeSelect(arr, address);

    // 分页函数
    z按页数分页 = (arr, pageSize) => this.pageByCount(arr, pageSize);
    z按行数分页 = (arr, rowsPerPage) => this.pageByRows(arr, rowsPerPage);
    z按下标分页 = (arr, pageIndices) => this.pageByIndexs(arr, pageIndices);

    // 其他实用函数
    z间隔取数 = (arr, n) => this.nth(arr, n);
    z补齐数组 = (arr, length, padValue) => this.pad(arr, length, padValue);
    z重复N次 = (arr, count) => this.repeat(arr, count);
    z跳过前N个 = (arr, count) => this.skip(arr, count);
    z跳过前面连续满足 = (arr, predicate) => this.skipWhile(arr, predicate);
    z超级透视 = (arr, rowField, colField, valueField, aggregator) => this.superPivot(arr, rowField, colField, valueField, aggregator);

    // ==================== 新增补充函数的中文别名 ====================

    // 查找相关函数
    z查找所有行下标 = (arr, predicate) => this.findRowsIndex(arr, predicate);
    z查找所有列下标 = (arr, predicate) => this.findColsIndex(arr, predicate);

    // 检查函数
    z满足条件 = (arr, predicate) => this.some(arr, predicate);
    z全部满足 = (arr, predicate) => this.every(arr, predicate);
    z倒序聚合 = (arr, reducer, initialValue) => this.reduceRight(arr, reducer, initialValue);

    // 栈操作函数
    z删除第一个 = (arr) => this.shift(arr);
    z尾部弹出一项 = (arr) => this.pop(arr);
    z追加一项 = (arr, ...items) => this.push(arr, ...items);

    // 位置查找函数
    z值位置 = (arr, value, fromIndex) => this.indexOf(arr, value, fromIndex);
    z从后往前值位置 = (arr, value, fromIndex) => this.lastIndexOf(arr, value, fromIndex);

    // 特殊工具函数
    z复制到指定位置 = (arr, targetRow, targetCol, startRow, startCol, endRow, endCol) => this.copyWithin(arr, targetRow, targetCol, startRow, startCol, endRow, endCol);
    z生成下标数组 = (length, start) => this.getIndexs(length, start);
    z下标数组 = (arr) => this.indexArray(arr);
    z插入行号 = (arr, colIndex, startNumber) => this.insertRowNum(arr, colIndex, startNumber);
    z一对多连接 = (leftArr, rightArr, leftKey, rightKey) => this.leftFulljoin(leftArr, rightArr, leftKey, rightKey);
    z结果 = (arr) => this.res(arr);

    // ==================== 新增函数的中文别名 ====================

    // 矩阵操作函数
    z转置 = (arr) => this.transpose(arr);
    z矩阵排版 = (arr, fillValue) => this.toMatrix(arr, fillValue);
    z矩阵分布 = (arr, rows, cols) => this.getMatrix(arr, rows, cols);
    z矩阵运算 = (arr, startRow, startCol, endRow, endCol, operation) => this.rangeMatrix(arr, startRow, startCol, endRow, endCol, operation);
    z重设大小 = (arr, newRows, newCols, fillValue) => this.resize(arr, newRows, newCols, fillValue);

    // 排序扩展函数
    z按规则升序 = (arr, keySelector) => this.sortBy(arr, keySelector);
    z多列排序 = (arr, colIndices, orders) => this.sortByCols(arr, colIndices, orders);
    z按规则降序 = (arr, keySelector) => this.sortByDesc(arr, keySelector);
    z自定义排序 = (arr, colIndex, orderList) => this.sortByList(arr, colIndex, orderList);
    z降序排序 = (arr) => this.sortDesc(arr);

    // 数据取用函数
    z行切片删除行 = (arr, start, deleteCount, ...items) => this.splice(arr, start, deleteCount, ...items);
    z取前N个 = (arr, n) => this.take(arr, n);
    z取前面连续满足 = (arr, predicate) => this.takeWhile(arr, predicate);

    // 数据处理函数
    z求和 = (arr, colSelector) => this.sum(arr, colSelector);
    z超级查找 = (arr, conditions) => this.superLookup(arr, conditions);

    // 输出转换函数
    z写入单元格 = (arr, rangeAddress) => this.toRange(arr, rangeAddress);
    z转字符串 = (arr, rowSeparator, colSeparator) => this.toString(arr, rowSeparator, colSeparator);
    z去重并集 = (arr1, arr2) => this.union(arr1, arr2);
    z左右连接 = (...arrays) => this.zip(...arrays);

    // 工具函数
    z错误值 = (value) => this.isError(value);
    z处理空值 = (arr, replacement) => this.noNull(arr, replacement);
    z解析函数表达式 = (expression) => this.parseLambda(expression);

    // ==================== 新增改进功能 ====================

    /**
     * ==================== WPS Range对象集成功能 ====================
     * 这些功能需要WPS JSA环境支持
     */

    /**
     * 从WPS Range对象读取数据为二维数组 - 支持链式调用
     * @param {Range} rng - WPS Range对象
     * @param {boolean} hasHeader - 是否包含表头,默认true
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    fromRange(rng, hasHeader = true) {
        if (!rng || !rng.Value2) {
            this.data = [];
            return this;
        }

        const data = rng.Value2;
        // 如果是单个单元格
        if (!Array.isArray(data)) {
            this.data = [[data]];
            return this;
        }

        this.data = data;
        return this;
    }

    /**
     * 将二维数组写入WPS Range对象
     * @param {Array} array2D - 二维数组
     * @param {Range} rng - WPS Range对象
     * @returns {boolean} 成功返回true
     */
    toRangeObject(array2D, rng) {
        if (!rng || this.isEmpty(array2D)) {
            return false;
        }

        try {
            rng.Value2 = array2D;
            return true;
        } catch (e) {
            console.log("Array2D.toRangeObject 写入失败:", e.message);
            return false;
        }
    }

    /**
     * 从单元格地址读取数据(需要在WPS环境中使用)
     * @param {string} address - 单元格地址,如"A1:C10"
     * @param {Worksheet} sheet - 工作表对象,默认为当前活动表
     * @returns {Array} 二维数组
     */
    fromAddress(address, sheet = null) {
        try {
            // 检查是否在WPS环境中
            if (typeof Application === 'undefined') {
                console.log("Array2D.fromAddress: 需要在WPS JSA环境中使用");
                return [];
            }

            const ws = sheet || Application.ActiveSheet;
            const rng = ws.Range(address);
            return this.fromRange(rng);
        } catch (e) {
            console.log("Array2D.fromAddress 读取失败:", e.message);
            return [];
        }
    }

    /**
     * 将数据写入指定单元格地址(需要在WPS环境中使用) - 支持链式调用
     * @param {string} address - 单元格地址,如"A1"
     * @param {Worksheet} sheet - 工作表对象
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    toAddress(address, sheet = null) {
        try {
            // 检查是否在WPS环境中
            if (typeof Application === 'undefined') {
                console.log("Array2D.toAddress: 需要在WPS JSA环境中使用");
                return this;
            }

            const ws = sheet || Application.ActiveSheet;
            const rng = ws.Range(address);

            // 自动调整范围大小
            const rows = this.data.length;
            const cols = this.data[0] ? this.data[0].length : 1;
            const targetRng = rng.Resize(rows, cols);

            this.toRangeObject(this.data, targetRng);
        } catch (e) {
            console.log("Array2D.toAddress 写入失败:", e.message);
        }

        return this;
    }

    // 中文版本
    z从区域读取 = (rng, hasHeader = true) => this.fromRange(rng, hasHeader);
    z写入区域 = (array2D, rng) => this.toRangeObject(array2D, rng);
    z从地址读取 = (address, sheet = null) => this.fromAddress(address, sheet);
    z写入地址 = (array2D, address, sheet = null) => this.toAddress(array2D, address, sheet);

    /**
     * ==================== 快速统计聚合函数 ====================
     */

    /**
     * 按列执行聚合计算(类似Excel的AGGREGATE函数)
     * @param {Array} array2D - 二维数组
     * @param {number|string} func - 聚合函数类型 (1-19) 或函数名
     * @param {number} colIndex - 列索引
     * @param {boolean} skipHeader - 是否跳过表头,默认true
     * @returns {number|Array} 聚合结果
     */
    agg(array2D, func, colIndex, skipHeader = true) {
        if (this.isEmpty(array2D)) {
            return 0;
        }

        // 跳过表头
        const data = skipHeader && array2D.length > 1 ? array2D.slice(1) : array2D;

        // 提取指定列
        const column = data.map(row => Array.isArray(row) ? row[colIndex] : row);

        // 函数类型映射
        const funcMap = {
            1: 'average',   // AVERAGE
            2: 'count',     // COUNT
            3: 'countA',    // COUNTA
            4: 'max',       // MAX
            5: 'min',       // MIN
            6: 'product',   // PRODUCT
            7: 'stdevS',    // STDEV.S
            8: 'stdevP',    // STDEV.P
            9: 'sum',       // SUM
            10: 'varS',     // VAR.S
            11: 'varP',     // VAR.P
            12: 'median',   // MEDIAN
            13: 'modeSn',   // MODE.SNGL
            14: 'large',    // LARGE
            15: 'small',    // SMALL
            16: 'percentileInc', // PERCENTILE.INC
            17: 'quartileInc',   // QUARTILE.INC
            18: 'percentileExc', // PERCENTILE.EXC
            19: 'quartileExc'    // QUARTILE.EXC
        };

        const funcType = typeof func === 'number' ? funcMap[func] : func;

        switch (funcType) {
            case 'average':
                return this.average(column);
            case 'count':
                return column.filter(v => typeof v === 'number' && !isNaN(v)).length;
            case 'countA':
                return column.filter(v => v !== null && v !== undefined && v !== '').length;
            case 'max':
                return this.max(column);
            case 'min':
                return this.min(column);
            case 'product':
                return column.reduce((acc, val) => acc * (Number(val) || 1), 1);
            case 'sum':
                return this.sum(column);
            case 'median':
                const sorted = column.filter(v => typeof v === 'number').sort((a, b) => Number(a) - Number(b));
                if (sorted.length === 0) return 0;
                const mid = Math.floor(sorted.length / 2);
                return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
            default:
                console.log(`Array2D.agg: 不支持的函数类型 ${funcType}`);
                return 0;
        }
    }

    /**
     * 多列同时聚合
     * @param {Array} array2D - 二维数组
     * @param {string} func - 聚合函数名
     * @param {Array} colIndexes - 列索引数组
     * @param {boolean} skipHeader - 是否跳过表头
     * @returns {Array} 聚合结果数组
     */
    aggCols(array2D, func, colIndexes, skipHeader = true) {
        const arr2d = this;
        return colIndexes.map(colIdx => arr2d.agg(array2D, func, colIdx, skipHeader));
    }

    // 中文版本
    z聚合 = (array2D, func, colIndex, skipHeader = true) => this.agg(array2D, func, colIndex, skipHeader);
    z多列聚合 = (array2D, func, colIndexes, skipHeader = true) => this.aggCols(array2D, func, colIndexes, skipHeader);

    /**
     * ==================== 数据验证和清洗功能 ====================
     */

    /**
     * 检测并标记空值
     * @param {Array} array2D - 二维数组
     * @param {string} fillValue - 填充值
     * @returns {Array} 填充后的数组
     */
    fillBlank(array2D, fillValue = '') {
        if (this.isEmpty(array2D)) {
            return [];
        }

        return array2D.map(row => {
            if (!Array.isArray(row)) {
                return row;
            }
            return row.map(cell =>
                cell === null || cell === undefined || cell === '' ? fillValue : cell
            );
        });
    }

    /**
     * 检测重复行
     * @param {Array} array2D - 二维数组
     * @param {Array} colIndexes - 用于判断重复的列索引数组
     * @returns {Array} 重复行的索引数组
     */
    findDuplicates(array2D, colIndexes = null) {
        const seen = new Set();
        const duplicates = [];

        const data = array2D.length > 0 ? array2D.slice(1) : array2D;  // 跳过表头

        data.forEach((row, idx) => {
            const key = colIndexes
                ? colIndexes.map(i => row[i]).join('|')
                : row.join('|');

            if (seen.has(key)) {
                duplicates.push(idx + 1);  // +1 因为跳过了表头
            } else {
                seen.add(key);
            }
        });

        return duplicates;
    }

    /**
     * 数据类型转换 - 支持链式调用
     * @param {Object} typeMap - 列索引到类型的映射 {0: 'number', 1: 'date'}
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    convertTypes(typeMap) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = this.data.map((row, i) => {
            if (i === 0) return row;  // 保留表头

            return row.map((cell, colIdx) => {
                if (!typeMap[colIdx]) return cell;

                switch (typeMap[colIdx]) {
                    case 'number':
                        return Number(cell);
                    case 'string':
                        return String(cell);
                    case 'date':
                        return new Date(cell);
                    case 'boolean':
                        return Boolean(cell);
                    default:
                        return cell;
                }
            });
        });

        this.data = result;
        return this;
    }

    // 中文版本
    z填充空值 = (array2D, fillValue = '') => this.fillBlank(array2D, fillValue);
    z查找重复 = (array2D, colIndexes = null) => this.findDuplicates(array2D, colIndexes);
    z转换类型 = (array2D, typeMap) => this.convertTypes(array2D, typeMap);

    /**
     * ==================== 性能监控功能 ====================
     */

    /**
     * 性能计时器存储
     * @private
     */
    _performanceTimers = {};

    /**
     * 开始性能计时
     * @param {string} label - 计时标签
     */
    timeStart(label = 'default') {
        this._performanceTimers[label] = Date.now();
    }

    /**
     * 结束性能计时并输出
     * @param {string} label - 计时标签
     * @returns {number} 耗时(毫秒)
     */
    timeEnd(label = 'default') {
        if (!this._performanceTimers[label]) {
            console.log(`Array2D.timeEnd: 计时器 '${label}' 不存在`);
            return 0;
        }

        const elapsed = Date.now() - this._performanceTimers[label];
        console.log(`Array2D性能计时 [${label}]: ${elapsed}ms`);
        delete this._performanceTimers[label];
        return elapsed;
    }

    /**
     * 获取数组统计信息
     * @param {Array} array2D - 二维数组
     * @returns {Object} 统计信息
     */
    getInfo(array2D) {
        if (this.isEmpty(array2D)) {
            return { rows: 0, cols: 0, cells: 0, hasHeader: false };
        }

        const info = {
            rows: array2D.length,
            cols: array2D[0] ? array2D[0].length : 0,
            cells: 0,
            hasHeader: true,
            columnTypes: []
        };

        info.cells = info.rows * info.cols;

        // 检测每列数据类型
        for (let col = 0; col < info.cols; col++) {
            const types = new Set();
            for (let row = 1; row < info.rows; row++) {
                const val = array2D[row][col];
                if (val === null || val === undefined) {
                    types.add('empty');
                } else if (typeof val === 'number') {
                    types.add('number');
                } else if (typeof val === 'string') {
                    types.add('string');
                } else if (typeof val === 'boolean') {
                    types.add('boolean');
                } else {
                    types.add('other');
                }
            }
            info.columnTypes.push([...types]);
        }

        return info;
    }

    // 中文版本
    z计时开始 = (label = 'default') => this.timeStart(label);
    z计时结束 = (label = 'default') => this.timeEnd(label);
    z获取信息 = (array2D) => this.getInfo(array2D);

    /**
     * ==================== 快速辅助函数 ====================
     */

    /**
     * 快速查找并返回值(VLOOKUP简化版)
     * @param {Array} array2D - 二维数组
     * @param {*} lookupValue - 查找值
     * @param {number} lookupCol - 查找列索引
     * @param {number} returnCol - 返回列索引
     * @returns {*} 查找结果
     */
    vlookup(array2D, lookupValue, lookupCol, returnCol) {
        const result = this.find(array2D, row => row[lookupCol] == lookupValue, true);
        return result ? result[returnCol] : null;
    }

    /**
     * 条件求和(SUMIF简化版)
     * @param {Array} array2D - 二维数组
     * @param {number} conditionCol - 条件列索引
     * @param {*} conditionValue - 条件值
     * @param {number} sumCol - 求和列索引
     * @returns {number} 求和结果
     */
    sumIf(array2D, conditionCol, conditionValue, sumCol) {
        const filtered = this.filter(array2D, row => row[conditionCol] == conditionValue, true);
        return this.sum(filtered.map(row => row[sumCol]));
    }

    /**
     * 条件计数(COUNTIF简化版)
     * @param {Array} array2D - 二维数组
     * @param {number} conditionCol - 条件列索引
     * @param {*} conditionValue - 条件值
     * @returns {number} 计数结果
     */
    countIf(array2D, conditionCol, conditionValue) {
        return this.filter(array2D, row => row[conditionCol] == conditionValue, true).length;
    }

    // 中文版本
    z查找 = (array2D, lookupValue, lookupCol, returnCol) => this.vlookup(array2D, lookupValue, lookupCol, returnCol);
    z条件求和 = (array2D, conditionCol, conditionValue, sumCol) => this.sumIf(array2D, conditionCol, conditionValue, sumCol);
    z条件计数 = (array2D, conditionCol, conditionValue) => this.countIf(array2D, conditionCol, conditionValue);

    /**
     * ==================== 导出和格式化功能 ====================
     */

    /**
     * 导出为CSV格式字符串
     * @param {Array} array2D - 二维数组
     * @param {string} delimiter - 分隔符,默认逗号
     * @returns {string} CSV字符串
     */
    toCSV(array2D, delimiter = ',') {
        if (this.isEmpty(array2D)) {
            return '';
        }

        return array2D.map(row =>
            row.map(cell => {
                const cellStr = String(cell ?? '');
                // 如果包含分隔符或引号,需要用引号包裹
                if (cellStr.includes(delimiter) || cellStr.includes('"') || cellStr.includes('\n')) {
                    return `"${cellStr.replace(/"/g, '""')}"`;
                }
                return cellStr;
            }).join(delimiter)
        ).join('\n');
    }

    /**
     * 导出为HTML表格
     * @param {Array} array2D - 二维数组
     * @param {boolean} hasHeader - 是否有表头
     * @returns {string} HTML表格字符串
     */
    toHTML(array2D, hasHeader = true) {
        if (this.isEmpty(array2D)) {
            return '';
        }

        let html = '<table>\n';

        array2D.forEach((row, idx) => {
            const tag = (hasHeader && idx === 0) ? 'th' : 'td';
            html += '  <tr>\n';
            row.forEach(cell => {
                html += `    <${tag}>${cell ?? ''}</${tag}>\n`;
            });
            html += '  </tr>\n';
        });

        html += '</table>';
        return html;
    }

    /**
     * 导出为Markdown表格
     * @param {Array} array2D - 二维数组
     * @returns {string} Markdown表格字符串
     */
    toMarkdown(array2D) {
        if (this.isEmpty(array2D)) {
            return '';
        }

        let md = '';

        // 表头
        const header = array2D[0].map(cell => String(cell ?? '')).join(' | ');
        md += `| ${header} |\n`;

        // 分隔线
        const separator = array2D[0].map(() => '---').join(' | ');
        md += `| ${separator} |\n`;

        // 数据行
        for (let i = 1; i < array2D.length; i++) {
            const row = array2D[i].map(cell => String(cell ?? '')).join(' | ');
            md += `| ${row} |\n`;
        }

        return md;
    }

    // 中文版本
    z转CSV = (array2D, delimiter = ',') => this.toCSV(array2D, delimiter);
    z转HTML = (array2D, hasHeader = true) => this.toHTML(array2D, hasHeader);
    z转Markdown = (array2D) => this.toMarkdown(array2D);

    /**
     * ==================== 列累加函数 ====================
     */

    /**
     * 列累加求和（运行总计）
     * 计算指定列的累加和，类似Excel中的累计求和
     * @param {Array} array2D - 二维数组
     * @param {number} colIndex - 列索引
     * @param {boolean} skipHeader - 是否跳过表头，默认true
     * @returns {Array} 累加结果数组
     */
    cumulativeSum(array2D, colIndex, skipHeader = true) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = JSON.parse(JSON.stringify(array2D)); // 深拷贝
        let sum = 0;
        const startRow = skipHeader ? 1 : 0;

        for (let i = startRow; i < result.length; i++) {
            const value = Number(result[i][colIndex]) || 0;
            sum += value;
            result[i][colIndex] = sum;
        }

        return result;
    }

    /**
     * 分组列累加求和
     * 按指定列分组后，计算累加和
     * @param {Array} array2D - 二维数组
     * @param {number} colIndex - 要累加的列索引
     * @param {number} groupCol - 分组列索引
     * @param {boolean} skipHeader - 是否跳过表头
     * @returns {Array} 累加结果数组
     */
    cumulativeSumBy(array2D, colIndex, groupCol, skipHeader = true) {
        if (this.isEmpty(array2D)) {
            return [];
        }

        const result = JSON.parse(JSON.stringify(array2D)); // 深拷贝
        const startRow = skipHeader ? 1 : 0;

        // 按分组列索引
        const groups = {};
        for (let i = startRow; i < result.length; i++) {
            const groupKey = result[i][groupCol];
            if (!groups[groupKey]) {
                groups[groupKey] = 0;
            }
        }

        // 计算每组累加和
        for (let i = startRow; i < result.length; i++) {
            const groupKey = result[i][groupCol];
            const value = Number(result[i][colIndex]) || 0;
            groups[groupKey] += value;
            result[i][colIndex] = groups[groupKey];
        }

        return result;
    }

    /**
     * 添加累加列 - 支持链式调用
     * 不修改原列，而是添加新的累加列
     * @param {number} sourceCol - 源列索引
     * @param {number} targetCol - 目标列索引（可选，默认添加到末尾）
     * @param {boolean} skipHeader - 是否跳过表头
     * @returns {clsArray2D} 返回当前实例以支持链式调用
     */
    addCumulativeCol(sourceCol, targetCol = null, skipHeader = true) {
        if (this.isEmpty(this.data)) {
            this.data = [];
            return this;
        }

        const result = [];
        const startRow = skipHeader ? 1 : 0;

        // 确定目标列位置
        const colCount = this.data[0].length;
        const insertCol = targetCol !== null ? targetCol : colCount;

        let sum = 0;

        for (let i = 0; i < this.data.length; i++) {
            const row = [...this.data[i]];

            if (i < startRow) {
                // 表头行：添加列名
                row.splice(insertCol, 0, this.data[i][sourceCol] + '(累加)');
            } else {
                // 数据行：计算累加和
                const value = Number(this.data[i][sourceCol]) || 0;
                sum += value;
                row.splice(insertCol, 0, sum);
            }

            result.push(row);
        }

        this.data = result;
        return this;
    }

    // 中文版本
    z列累加 = (array2D, colIndex, skipHeader = true) => this.cumulativeSum(array2D, colIndex, skipHeader);
    z分组列累加 = (array2D, colIndex, groupCol, skipHeader = true) => this.cumulativeSumBy(array2D, colIndex, groupCol, skipHeader);
    z添加累加列 = (array2D, sourceCol, targetCol, skipHeader) => this.addCumulativeCol(array2D, sourceCol, targetCol, skipHeader);
}

/**
 * 创建Array2D全局实例 - 可直接使用 Array2D.方法名() 调用
 * @example
 * Array2D.sum([[1,2,3],[4,5,6]]); // 21
 */
const Array2D = new clsArray2D();

/**
 * 导出Array2D类 - 支持WPS JSA环境
 */
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { Array2D };
}
// WPS JSA环境：导出为全局变量
if (typeof window !== 'undefined' || typeof Application !== 'undefined') {
    this.Array2D = Array2D;
}

/**
 * 示例使用函数
 */
function 示例使用Array2D() {
    console.log("=== Array2D JSA880 示例使用 ===");

    const array2d = new Array2D();

    // 1. 创建测试数组
    const testArray = [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
    ];

    console.log("1. 测试数组:");
    console.log(testArray);

    // 2. 测试基本信息函数
    console.log("\n2. 基本信息:");
    console.log("版本:", array2d.version());
    console.log("是否为空:", array2d.isEmpty(testArray));
    console.log("元素数量:", array2d.count(testArray));

    // 3. 测试数组操作
    console.log("\n3. 数组操作:");
    const cloned = array2d.copy(testArray);
    console.log("克隆数组:", cloned);

    const filled = array2d.fill(testArray, 0, 1, 2);
    console.log("填充后:", filled);

    // 4. 测试查找和筛选
    console.log("\n4. 查找和筛选:");
    console.log("查找元素5的下标:", array2d.findIndex(testArray, 5));
    console.log("是否包含10:", array2d.includes(testArray, 10));
    console.log("是否包含5:", array2d.includes(testArray, 5));

    // 5. 测试统计计算
    console.log("\n5. 统计计算:");
    console.log("平均值:", array2d.average(testArray));
    console.log("最大值:", array2d.max(testArray));
    console.log("最小值:", array2d.min(testArray));

    // 6. 测试中文函数
    console.log("\n6. 中文函数测试:");
    console.log("版本:", array2d.z版本());
    console.log("数量:", array2d.z数量(testArray));
    console.log("第一个:", array2d.z第一个(testArray));
    console.log("最后一个:", array2d.z最后一个(testArray));

    console.log("\n=== 示例结束 ===");
}

/**
 * ==================== RngUtils 单元格区域辅助函数库 ====================
 * 根据 https://vbayyds.com/api/jsa880/RngUtils.html 文档编写
 * 用于处理表格单元格区域
 *
 * @class _RngUtilsClass
 * @description 单元格对象辅助函数库，用于处理表格单元格区域
 * @example
 * const lastCell = RngUtils.z最后一个(Range("A1:A13"));
 * console.log(lastCell.Address());
 */
/**
 * Range包装类 - 支持链式调用
 * @private
 * @class
 */
class _RangeWrapper {
    /**
     * 构造函数
     * @param {Range} range - WPS Range对象
     * @param {clsRngUtils} utils - RngUtils实例
     */
    constructor(range, utils) {
        this._range = range;
        this._utils = utils;
    }

    /**
     * 获取原始Range对象
     * @returns {Range} WPS Range对象
     */
    get value() {
        return this._range;
    }

    /**
     * 转换为安全数组
     * @returns {Array} 二维数组
     */
    safeArray() {
        return this._utils.z安全数组(this._range);
    }

    /**
     * 获取安全区域
     * @returns {_RangeWrapper} 包装后的安全区域
     */
    safeRange() {
        return new _RangeWrapper(this._utils.z安全区域(this._range), this._utils);
    }

    /**
     * 获取最大行区域
     * @param {string} col - 列参数
     * @returns {_RangeWrapper} 包装后的最大行区域
     */
    maxRange(col = "A") {
        return new _RangeWrapper(this._utils.z最大行区域(this._range, col), this._utils);
    }

    /**
     * 转为字符串
     * @returns {string} Range地址
     */
    toString() {
        return this._range.Address ? this._range.Address(false, false) : String(this._range);
    }

    /**
     * 代理Range的所有属性和方法
     */
    get Address() { return this._range.Address.bind(this._range); }
    get Value() { return this._range.Value; }
    set Value(v) { this._range.Value = v; }
    get Value2() { return this._range.Value2; }
    set Value2(v) { this._range.Value2 = v; }
    get Row() { return this._range.Row; }
    get Column() { return this._range.Column; }
    get Rows() { return this._range.Rows; }
    get Columns() { return this._range.Columns; }
    get Count() { return this._range.Count; }
    get Worksheet() { return this._range.Worksheet; }
    get Cells() { return this._range.Cells; }
    get Item() { return this._range.Item; }
}

class clsRngUtils {
    /**
     * 构造函数 - 创建RngUtils实例
     * @constructor
     * @description 初始化RngUtils对象
     * @example
     * const rng = new RngUtils();
     */
    constructor(initialRange = null) {
        this.MODULE_NAME = "RngUtils";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学JSA880框架";
        this.range = initialRange || null;  // 存储内部Range状态
    }

    // ==================== 辅助函数 ====================

    /**
     * 将参数转换为Range对象
     * @private
     * @param {Range|string} rng - Range对象或地址字符串
     * @returns {Range} Range对象
     */
    _toRange(rng) {
        if (typeof rng === 'string') {
            return Range(rng);
        }
        // 处理 clsChainableRange 对象
        if (rng && typeof rng.unwrap === 'function') {
            return rng.unwrap();
        }
        return rng;
    }

    /**
     * 获取/设置当前Range（支持链式调用）
     * @param {Range|string} newRange - 新的Range（可选）
     * @returns {clsRngUtils|Range} 传入参数时返回this，否则返回当前Range
     * @example
     * const rng = new clsRngUtils("A1");
     * rng.rng("B1").z加边框();  // 链式调用
     * const current = rng.rng();  // 获取当前Range
     */
    rng(newRange) {
        if (newRange !== undefined) {
            this.range = this._toRange(newRange);
            return this;
        }
        return this.range;
    }

    /**
     * 获取当前Range的值
     * @returns {Array} Range的二维数组值
     * @example
     * const rng = new clsRngUtils("A1:C3");
     * const values = rng.val();  // 获取值
     */
    val() {
        if (!this.range) return null;
        return this.z安全数组(this.range);
    }

    // ==================== 基本区域操作函数 ====================

    /**
     * 获取区域最后一个单元格（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A1:A13");
     * rng.z最后一个();
     * console.log(rng.rng().Address()); // $A$13
     */
    z最后一个() {
        const range = this._toRange(this.range);
        const rows = range.Rows.Count;
        const cols = range.Columns.Count;
        this.range = range.Cells(rows, cols);
        return this;
    }

    /**
     * 获取区域最后一个单元格（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    lastCell() {
        return this.z最后一个();
    }

    /**
     * 获取安全区域（与UsedRange的交集，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A:A");
     * rng.z安全区域();
     * console.log(rng.rng().Address()); // $A$1:$A$13
     */
    z安全区域() {
        const range = this._toRange(this.range);
        const usedRange = Application.ActiveSheet.UsedRange;
        this.range = Application.Intersect(range, usedRange);
        return this;
    }

    /**
     * 获取安全区域（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    safeRange() {
        return this.z安全区域();
    }

    /**
     * 将区域转换为安全数组
     * @param {Range|string} rng - 要转换为安全数组的区域
     * @returns {Array} 结果二维数组
     * @example
     * const rng = new RngUtils();
     * const safeArray = rng.z安全数组("A1:A13");
     * console.log(safeArray); // 二维数组
     */
    z安全数组(rng) {
        const range = this._toRange(rng);
        const arr = range.Value2;

        // 如果是单个单元格
        if (!Array.isArray(arr)) {
            return [[arr]];
        }

        // 如果是二维数组但只有一行
        if (arr.length === 1 || (arr.length === 1 && !Array.isArray(arr[0]))) {
            return [arr];
        }

        return arr;
    }

    /**
     * 将区域转换为安全数组（英文别名）
     * @param {Range|string} rng - 要转换为安全数组的区域
     * @returns {Array} 结果二维数组
     */
    safeArray(rng) {
        return this.z安全数组(rng);
    }

    // ==================== 行列操作函数 ====================

    /**
     * 获取指定区域的最大行数
     * @param {Range|string} rng - 要获取最大行数的区域
     * @returns {number} 最大行数
     * @example
     * const rng = new RngUtils();
     * console.log(rng.z最大行("A:A")); // 13
     */
    z最大行(rng) {
        const range = this._toRange(rng);
        const safeRng = this.z安全区域(range);

        // 如果安全区域为空，返回原区域的行
        if (!safeRng) {
            return range.Row + range.Rows.Count - 1;
        }

        return safeRng.Row + safeRng.Rows.Count - 1;
    }

    /**
     * 获取指定区域的最大行数（英文别名）
     * @param {Range|string} rng - 要获取最大行数的区域
     * @returns {number} 最大行数
     */
    endRow(rng) {
        return this.z最大行(rng);
    }

    /**
     * 获取指定区域的最后一行的单元格（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A1:A1000");
     * rng.z最大行单元格();
     * console.log(rng.rng().Address()); // $A$13
     */
    z最大行单元格() {
        const range = this._toRange(this.range);
        const safeRng = this.z安全区域(range);

        // 如果安全区域为空，使用原区域
        const lastRow = safeRng ? (safeRng.Row + safeRng.Rows.Count - 1) : (range.Row + range.Rows.Count - 1);
        const col = range.Column;
        const sheet = range.Worksheet;
        this.range = sheet.Cells(lastRow, col);
        return this;
    }

    /**
     * 获取指定区域的最后一行的单元格（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    endRowCell() {
        return this.z最大行单元格();
    }

    /**
     * 获取指定区域从第一行到最后一行的单元格区域（支持链式调用）
     * @param {string} col - 如果第一参数rng没有列号(整行), 则使用本参数指定列 "A","B","C"..... 默认为"A", 另有'-c','-u'参数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A1:A1000");
     * rng.z最大行区域();
     * console.log(rng.rng().Address()); // $A$1:$A$13
     */
    z最大行区域(col = "A") {
        const range = this._toRange(this.range);

        // 处理特殊参数
        if (col === "-c") {
            // 连续区域(CurrentRegion)
            this.range = range.CurrentRegion;
        } else if (col === "-u") {
            // 使用区域(UsedRange)
            const usedRange = range.Worksheet.UsedRange;
            const startRow = range.Row;
            const endRow = usedRange.Row + usedRange.Rows.Count - 1;
            const colIndex = range.Column;
            this.range = range.Worksheet.Range(
                range.Worksheet.Cells(startRow, colIndex),
                range.Worksheet.Cells(endRow, colIndex)
            );
        } else if (range.Address && range.Address().toString().match(/^\d+:\d+$/)) {
            // 处理整行的情况
            const endRow = this.z最大行(range.Worksheet.Columns(col));
            const rows = range.Address().toString().split(":");
            const startRow = parseInt(rows[0]);
            this.range = range.Worksheet.Rows(startRow + ":" + endRow);
        } else {
            // 默认情况 - 保持原区域的列范围，扩展行到最后一行
            const safeRng = this.z安全区域(range);
            const startRow = range.Row;
            const endRow = safeRng.Row + safeRng.Rows.Count - 1;
            const startCol = range.Column;
            const endCol = range.Column + range.Columns.Count - 1;
            this.range = range.Worksheet.Range(
                range.Worksheet.Cells(startRow, startCol),
                range.Worksheet.Cells(endRow, endCol)
            );
        }
        return this;
    }

    /**
     * 获取指定区域从第一行到最后一行的单元格区域（英文别名，支持链式调用）
     * @param {string} col - 列参数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    maxRange(col = "A") {
        return this.z最大行区域(col);
    }

    /**
     * 获取指定区域的最大列数
     * @param {Range|string} rng - 要获取最大列数的区域
     * @returns {number} 最大列数
     * @example
     * const rng = new RngUtils();
     * console.log(rng.z最大列("3:3")); // 3
     */
    z最大列(rng) {
        const range = this._toRange(rng);
        const safeRng = this.z安全区域(range);

        // 如果安全区域为空，返回原区域的列
        if (!safeRng) {
            return range.Column + range.Columns.Count - 1;
        }

        return safeRng.Column + safeRng.Columns.Count - 1;
    }

    /**
     * 获取指定区域的最大列数（英文别名）
     * @param {Range|string} rng - 要获取最大列数的区域
     * @returns {number} 最大列数
     */
    endCol(rng) {
        return this.z最大列(rng);
    }

    /**
     * 获取指定区域的最后一列的单元格（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("1:1");
     * rng.z最大列单元格();
     * console.log(rng.rng().Address()); // $F$1
     */
    z最大列单元格() {
        const range = this._toRange(this.range);
        const safeRng = this.z安全区域(range);

        // 如果安全区域为空，使用原区域
        const lastCol = safeRng ? (safeRng.Column + safeRng.Columns.Count - 1) : (range.Column + range.Columns.Count - 1);
        const row = range.Row;
        const sheet = range.Worksheet;
        this.range = sheet.Cells(row, lastCol);
        return this;
    }

    /**
     * 获取指定区域的最后一列的单元格（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    endColCell() {
        return this.z最大列单元格();
    }

    // ==================== 可见区域操作函数 ====================

    /**
     * 将指定区域的可见单元格(不包括隐藏行)转换为数组
     * @param {Range|string} rng - 要转换为数组的区域
     * @param {Worksheet} tempSheet - 临时工作表（可选）
     * @returns {Array} 包含可见单元格值的数组
     * @example
     * const rng = new RngUtils();
     * const visibleArr = rng.z可见区数组("1:4");
     * console.log(visibleArr);
     */
    z可见区数组(rng, tempSheet = null) {
        const range = this._toRange(rng);
        const visibleRange = this.z可见区域(range);

        if (tempSheet) {
            // 如果提供了临时表，复制到临时表
            visibleRange.Copy(tempSheet.Range("A1"));
            Application.CutCopyMode = false;
            return tempSheet.Range("A1").Resize(visibleRange.Rows.Count, visibleRange.Columns.Count).Value2;
        } else {
            // 直接返回可见区域的值
            return this.z安全数组(visibleRange);
        }
    }

    /**
     * 将指定区域的可见单元格转换为数组（英文别名）
     * @param {Range|string} rng - 要转换为数组的区域
     * @param {Worksheet} tempSheet - 临时工作表（可选）
     * @returns {Array} 包含可见单元格值的数组
     */
    visibleArray(rng, tempSheet = null) {
        return this.z可见区数组(rng, tempSheet);
    }

    /**
     * 获取指定区域的可见区域单元格（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils(Range("A1:D15"));
     * rng.z可见区域();
     * console.log(rng.rng().Address()); // $A$1:$D$11
     */
    z可见区域() {
        const range = this._toRange(this.range);

        // 获取特殊单元格（可见单元格）
        try {
            const visibleCells = range.SpecialCells(12); // xlCellTypeVisible = 12
            this.range = visibleCells;
        } catch (e) {
            // 如果没有可见单元格或出错，返回原区域
            this.range = range;
        }
        return this;
    }

    /**
     * 获取指定区域的可见区域单元格（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    visibleRange() {
        return this.z可见区域();
    }

    // ==================== 格式操作函数 ====================

    /**
     * 为指定区域添加边框（支持链式调用）
     * @param {number} LineStyle - 边框线条样式（默认1）
     * @param {number} Weight - 边框线条粗细（默认2）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A3:D7");
     * rng.z加边框();
     */
    z加边框(LineStyle = 1, Weight = 2) {
        const range = this._toRange(this.range);
        const borders = range.Borders;
        borders.LineStyle = LineStyle;
        borders.Weight = Weight;
        return this;
    }

    /**
     * 为指定区域添加边框（英文别名，支持链式调用）
     * @param {number} LineStyle - 边框线条样式
     * @param {number} Weight - 边框线条粗细
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    addBorders(LineStyle = 1, Weight = 2) {
        return this.z加边框(LineStyle, Weight);
    }

    /**
     * 从源单元格区域复制粘贴格式到目标单元格区域（支持链式调用）
     * @param {Range|string} target - 要粘贴格式的单元格区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a14:d14");
     * rng.z复制粘贴格式("a18:d21");
     */
    z复制粘贴格式(target) {
        const sourceRange = this._toRange(this.range);
        const targetRange = this._toRange(target);

        try {
            // 先尝试使用 PasteSpecial 方法
            sourceRange.Copy();

            // 尝试不同的 PasteSpecial 调用方式
            try {
                // 方式1: 直接调用 PasteSpecial
                targetRange.PasteSpecial(-4122); // xlPasteFormats
                Application.CutCopyMode = false;
                return this;
            } catch (e1) {
                try {
                    // 方式2: 使用对象参数
                    targetRange.PasteSpecial(null, null, null, null, -4122);
                    Application.CutCopyMode = false;
                    return this;
                } catch (e2) {
                    try {
                        // 方式3: 通过 Selection 粘贴
                        targetRange.Select();
                        Application.Selection.PasteSpecial(-4122);
                        Application.CutCopyMode = false;
                        return this;
                    } catch (e3) {
                        // 所有 PasteSpecial 方法都失败，使用手动复制格式的方式
                        Application.CutCopyMode = false;
                        console.log("PasteSpecial 失败，使用手动复制格式");
                    }
                }
            }

            // 手动复制格式（当 PasteSpecial 不可用时）
            const src = sourceRange.Cells(1, 1);

            // 字体
            try {
                targetRange.Font.Name = src.Font.Name;
                targetRange.Font.Size = src.Font.Size;
                targetRange.Font.Bold = src.Font.Bold;
                targetRange.Font.Italic = src.Font.Italic;
                targetRange.Font.Color = src.Font.Color;
            } catch (e) { /* 忽略字体错误 */ }

            // 背景色
            try {
                targetRange.Interior.Color = src.Interior.Color;
            } catch (e) { /* 忽略背景色错误 */ }

            // 对齐
            try {
                targetRange.HorizontalAlignment = src.HorizontalAlignment;
                targetRange.VerticalAlignment = src.VerticalAlignment;
            } catch (e) { /* 忽略对齐错误 */ }

            // 数字格式
            try {
                targetRange.NumberFormat = src.NumberFormat;
            } catch (e) { /* 忽略数字格式错误 */ }

        } catch (e) {
            throw new Error("复制格式失败: " + e.message);
        }
        return this;
    }

    /**
     * 从源单元格区域复制粘贴格式（英文别名，支持链式调用）
     * @param {Range|string} target - 要粘贴格式的单元格区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    copyFormat(target) {
        return this.z复制粘贴格式(target);
    }

    /**
     * 从源单元格区域复制粘贴值到目标单元格区域（支持链式调用）
     * @param {Range|string} target - 要粘贴值的单元格区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a11:d14");
     * rng.z复制粘贴值("a18:d21");
     */
    z复制粘贴值(target) {
        const sourceRange = this._toRange(this.range);
        const targetRange = this._toRange(target);
        targetRange.Value2 = sourceRange.Value2;
        return this;
    }

    /**
     * 从源单元格区域复制粘贴值（英文别名，支持链式调用）
     * @param {Range|string} target - 要粘贴值的单元格区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    copyValue(target) {
        return this.z复制粘贴值(target);
    }

    // ==================== 行列选择函数 ====================

    /**
     * 获取指定区域的指定前几行（支持链式调用）
     * @param {number} count - 获取的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a3:d7");
     * rng.z取前几行(3);
     * console.log(rng.rng().Address()); // $A$3:$D$5
     */
    z取前几行(count) {
        const range = this._toRange(this.range);
        const startRow = range.Row;
        const endRow = startRow + count - 1;
        const startCol = range.Column;
        const endCol = range.Column + range.Columns.Count - 1;
        this.range = range.Worksheet.Range(
            range.Worksheet.Cells(startRow, startCol),
            range.Worksheet.Cells(endRow, endCol)
        );
        return this;
    }

    /**
     * 获取指定区域的指定前几行（英文别名，支持链式调用）
     * @param {number} count - 获取的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    takeRows(count) {
        return this.z取前几行(count);
    }

    /**
     * 跳过指定区域的指定前几行（支持链式调用）
     * @param {number} count - 要跳过的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a3:d7");
     * rng.z跳过前几行(3);
     * console.log(rng.rng().Address()); // $A$6:$D$7
     */
    z跳过前几行(count) {
        const range = this._toRange(this.range);
        const startRow = range.Row + count;
        const endRow = range.Row + range.Rows.Count - 1;
        const startCol = range.Column;
        const endCol = range.Column + range.Columns.Count - 1;

        if (startRow > endRow) {
            this.range = null;
            return this;
        }

        this.range = range.Worksheet.Range(
            range.Worksheet.Cells(startRow, startCol),
            range.Worksheet.Cells(endRow, endCol)
        );
        return this;
    }

    /**
     * 跳过指定区域的指定前几行（英文别名，支持链式调用）
     * @param {number} count - 要跳过的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    skipRows(count) {
        return this.z跳过前几行(count);
    }

    /**
     * 获取指定单元格区域的整行（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("11:14");
     * rng.z整行();
     * console.log(rng.rng().Address()); // $11:$14
     */
    z整行() {
        const range = this._toRange(this.range);
        this.range = range.EntireRow;
        return this;
    }

    /**
     * 获取指定单元格区域的整行（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    entireRow() {
        return this.z整行();
    }

    /**
     * 获取指定单元格区域的整列（支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A:B");
     * rng.z整列();
     * console.log(rng.rng().Address()); // $A:$B
     */
    z整列() {
        const range = this._toRange(this.range);
        this.range = range.EntireColumn;
        return this;
    }

    /**
     * 获取指定单元格区域的整列（英文别名，支持链式调用）
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    entireColumn() {
        return this.z整列();
    }

    /**
     * 获取指定单元格区域的行数
     * @param {Range|string} rng - 要获取行数的单元格区域
     * @returns {number} 行数
     * @example
     * const rng = new RngUtils();
     * console.log(rng.z行数("A12:D15")); // 4
     */
    z行数(rng) {
        const range = this._toRange(rng);
        return range.Rows.Count;
    }

    /**
     * 获取指定单元格区域的行数（英文别名）
     * @param {Range|string} rng - 要获取行数的单元格区域
     * @returns {number} 行数
     */
    rowsCount(rng) {
        return this.z行数(rng);
    }

    /**
     * 获取指定单元格区域的列数
     * @param {Range|string} rng - 要获取列数的单元格区域
     * @returns {number} 列数
     * @example
     * const rng = new RngUtils();
     * console.log(rng.z列数("A12:C15")); // 3
     */
    z列数(rng) {
        const range = this._toRange(rng);
        return range.Columns.Count;
    }

    /**
     * 获取指定单元格区域的列数（英文别名）
     * @param {Range|string} rng - 要获取列数的单元格区域
     * @returns {number} 列数
     */
    colsCount(rng) {
        return this.z列数(rng);
    }

    // ==================== 合并单元格操作函数 ====================

    /**
     * 合并指定区域中相同的行或者列（支持链式调用）
     * @param {string} direction - 合并方式 -r按行 -c按列 -rm先行后列 -cm 列按上下级关系 默认为r
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a1:j1");
     * rng.z合并相同单元格("c"); // 标题按列合并
     * const rng2 = new clsRngUtils("a3:b16");
     * rng2.z合并相同单元格("rm"); // 按行合并，并且考虑前面一列的关系
     */
    z合并相同单元格(direction = "r") {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        if (direction === "-c" || direction === "c") {
            // 按列合并
            for (let col = 0; col < arr[0].length; col++) {
                let startRow = 0;
                for (let row = 1; row < arr.length; row++) {
                    if (arr[row][col] !== arr[startRow][col]) {
                        if (row - startRow > 1) {
                            range.Worksheet.Range(
                                range.Worksheet.Cells(range.Row + startRow, range.Column + col),
                                range.Worksheet.Cells(range.Row + row - 1, range.Column + col)
                            ).Merge();
                        }
                        startRow = row;
                    }
                }
                // 合并最后一组
                if (arr.length - startRow > 1) {
                    range.Worksheet.Range(
                        range.Worksheet.Cells(range.Row + startRow, range.Column + col),
                        range.Worksheet.Cells(range.Row + arr.length - 1, range.Column + col)
                    ).Merge();
                }
            }
        } else if (direction === "-rm" || direction === "rm") {
            // 按行合并，考虑前面一列的关系
            for (let row = 0; row < arr.length; row++) {
                let startCol = 0;
                for (let col = 1; col < arr[row].length; col++) {
                    if (arr[row][col] !== arr[row][startCol]) {
                        if (col - startCol > 1) {
                            range.Worksheet.Range(
                                range.Worksheet.Cells(range.Row + row, range.Column + startCol),
                                range.Worksheet.Cells(range.Row + row, range.Column + col - 1)
                            ).Merge();
                        }
                        startCol = col;
                    }
                }
                // 合并最后一组
                if (arr[row].length - startCol > 1) {
                    range.Worksheet.Range(
                        range.Worksheet.Cells(range.Row + row, range.Column + startCol),
                        range.Worksheet.Cells(range.Row + row, range.Column + arr[row].length - 1)
                    ).Merge();
                }
            }
        } else if (direction === "-cm" || direction === "cm") {
            // 列按上下级关系
            for (let col = 0; col < arr[0].length; col++) {
                let startRow = 0;
                for (let row = 1; row < arr.length; row++) {
                    if (col > 0 && arr[row][col - 1] !== arr[startRow][col - 1]) {
                        // 前一列不同，必须断开
                        if (row - startRow > 1) {
                            range.Worksheet.Range(
                                range.Worksheet.Cells(range.Row + startRow, range.Column + col),
                                range.Worksheet.Cells(range.Row + row - 1, range.Column + col)
                            ).Merge();
                        }
                        startRow = row;
                    } else if (arr[row][col] !== arr[startRow][col]) {
                        if (row - startRow > 1) {
                            range.Worksheet.Range(
                                range.Worksheet.Cells(range.Row + startRow, range.Column + col),
                                range.Worksheet.Cells(range.Row + row - 1, range.Column + col)
                            ).Merge();
                        }
                        startRow = row;
                    }
                }
                // 合并最后一组
                if (arr.length - startRow > 1) {
                    range.Worksheet.Range(
                        range.Worksheet.Cells(range.Row + startRow, range.Column + col),
                        range.Worksheet.Cells(range.Row + arr.length - 1, range.Column + col)
                    ).Merge();
                }
            }
        } else {
            // 默认按行合并
            for (let row = 0; row < arr.length; row++) {
                let startCol = 0;
                for (let col = 1; col < arr[row].length; col++) {
                    if (arr[row][col] !== arr[row][startCol]) {
                        if (col - startCol > 1) {
                            range.Worksheet.Range(
                                range.Worksheet.Cells(range.Row + row, range.Column + startCol),
                                range.Worksheet.Cells(range.Row + row, range.Column + col - 1)
                            ).Merge();
                        }
                        startCol = col;
                    }
                }
                // 合并最后一组
                if (arr[row].length - startCol > 1) {
                    range.Worksheet.Range(
                        range.Worksheet.Cells(range.Row + row, range.Column + startCol),
                        range.Worksheet.Cells(range.Row + row, range.Column + arr[row].length - 1)
                    ).Merge();
                }
            }
        }
        return this;
    }

    /**
     * 合并指定区域中相同的行或者列（英文别名，支持链式调用）
     * @param {string} direction - 合并方式
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    mergeCells(direction = "r") {
        return this.z合并相同单元格(direction);
    }

    /**
     * 取消合并指定单元格区域中每行并按要求填充（支持链式调用）
     * @param {boolean} fillAll - true: 所有行填充 false: 仅首行填充
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a1:j1");
     * rng.z取消合并填充单元格();
     * const rng2 = new clsRngUtils("a3:b16");
     * rng2.z取消合并填充单元格();
     */
    z取消合并填充单元格(fillAll = true) {
        const range = this._toRange(this.range);

        // 遍历所有合并单元格
        for (let i = 1; i <= range.Count; i++) {
            const cell = range.Cells(i);
            if (cell.MergeCells && cell.MergeArea.Cells.Count > 1) {
                const mergeArea = cell.MergeArea;
                const value = cell.Value2;

                // 取消合并
                mergeArea.UnMerge();

                // 填充值
                if (fillAll) {
                    mergeArea.Value2 = value;
                } else {
                    // 仅首行填充
                    const firstRow = mergeArea.Rows(1);
                    firstRow.Value2 = value;
                }
            }
        }
        return this;
    }

    /**
     * 取消合并指定单元格区域中每行并按要求填充（英文别名，支持链式调用）
     * @param {boolean} fillAll - true: 所有行填充 false: 仅首行填充
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    unMergeCells(fillAll = true) {
        return this.z取消合并填充单元格(fillAll);
    }

    // ==================== 插入删除函数 ====================

    /**
     * 在指定单元格区域中按指定内容插入指定行数（支持链式调用）
     * @param {any} value - 要插入的行号数组或者字符串
     * @param {number} count - 要插入的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a12:d15");
     * rng.z插入多行('*', 2); // a12:d15每行前均插入二行，插入值都为"*"
     * const rng2 = new clsRngUtils("a12:d15");
     * rng2.z插入多行(['aa','bb','cc','dd'], 1); // 每行前插入一行
     */
    z插入多行(value, count) {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        // 从下往上插入，避免索引变化
        for (let i = arr.length - 1; i >= 0; i--) {
            const row = range.Row + i;
            const insertValue = Array.isArray(value) ? (value[i] || value[0]) : value;

            for (let j = 0; j < count; j++) {
                const insertRow = range.Worksheet.Rows(row);
                insertRow.Insert();
            }

            // 填充值
            const fillRange = range.Worksheet.Range(
                range.Worksheet.Cells(row, range.Column),
                range.Worksheet.Cells(row + count - 1, range.Column + range.Columns.Count - 1)
            );

            if (typeof insertValue === 'string') {
                fillRange.Value2 = insertValue;
            } else if (Array.isArray(insertValue)) {
                for (let k = 0; k < insertValue.length && k < count; k++) {
                    range.Worksheet.Rows(row + k).Value2 = insertValue[k];
                }
            }
        }
        return this;
    }

    /**
     * 在指定单元格区域中按指定内容插入指定行数（英文别名，支持链式调用）
     * @param {any} value - 要插入的行号数组或者字符串
     * @param {number} count - 要插入的行数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    insertRows(value, count) {
        return this.z插入多行(value, count);
    }

    /**
     * 在指定单元格区域中按指定内容插入指定列数（支持链式调用）
     * @param {any} value - 要插入的列号数组或者字符串
     * @param {number} count - 要插入的列数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a12:d14");
     * rng.z插入多列('*', 2); // 每列前均插入二列
     * const rng2 = new clsRngUtils("a12:d14");
     * rng2.z插入多列([['aa'],['bb'],['cc']], 1); // 每列前插入一列
     */
    z插入多列(value, count) {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);
        const cols = arr[0].length;

        // 从右往左插入，避免索引变化
        for (let i = cols - 1; i >= 0; i--) {
            const col = range.Column + i;
            const insertValue = Array.isArray(value) && Array.isArray(value[i]) ? value[i][0] : value;

            for (let j = 0; j < count; j++) {
                const insertCol = range.Worksheet.Columns(col);
                insertCol.Insert();
            }

            // 填充值
            const fillRange = range.Worksheet.Range(
                range.Worksheet.Cells(range.Row, col),
                range.Worksheet.Cells(range.Row + range.Rows.Count - 1, col + count - 1)
            );

            if (typeof insertValue === 'string') {
                fillRange.Value2 = insertValue;
            }
        }
        return this;
    }

    /**
     * 在指定单元格区域中按指定内容插入指定列数（英文别名，支持链式调用）
     * @param {any} value - 要插入的列号数组或者字符串
     * @param {number} count - 要插入的列数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    insertCols(value, count) {
        return this.z插入多列(value, count);
    }

    /**
     * 删除指定单元格区域中的所有空白行（支持链式调用）
     * @param {boolean} entireColumn - 默认删除整列 false的时候只作用选中区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("a11:d17");
     * rng.z删除空白行(); // 删除区域内所有的空白行
     */
    z删除空白行(entireColumn = false) {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        // 找出空白行（从下往上）
        const blankRows = [];
        for (let i = arr.length - 1; i >= 0; i--) {
            let isBlank = true;
            for (let j = 0; j < arr[i].length; j++) {
                if (arr[i][j] !== null && arr[i][j] !== undefined && arr[i][j] !== '') {
                    isBlank = false;
                    break;
                }
            }
            if (isBlank) {
                blankRows.push(range.Row + i);
            }
        }

        // 删除空白行
        for (const row of blankRows) {
            if (entireColumn) {
                range.Worksheet.Rows(row).Delete();
            } else {
                range.Worksheet.Range(
                    range.Worksheet.Cells(row, range.Column),
                    range.Worksheet.Cells(row, range.Column + range.Columns.Count - 1)
                ).Delete();
            }
        }
        return this;
    }

    /**
     * 删除指定单元格区域中的所有空白行（英文别名，支持链式调用）
     * @param {boolean} entireColumn - 默认删除整列
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    delBlankRows(entireColumn = false) {
        return this.z删除空白行(entireColumn);
    }

    /**
     * 删除指定单元格区域中的所有空白列（支持链式调用）
     * @param {boolean} entireColumn - 默认删除整列 false的时候只作用选中区域
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A11:G14");
     * rng.z删除空白列(); // 删除区域内所有的空白列
     */
    z删除空白列(entireColumn = false) {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        // 找出空白列（从右往左）
        const blankCols = [];
        for (let j = arr[0].length - 1; j >= 0; j--) {
            let isBlank = true;
            for (let i = 0; i < arr.length; i++) {
                if (arr[i][j] !== null && arr[i][j] !== undefined && arr[i][j] !== '') {
                    isBlank = false;
                    break;
                }
            }
            if (isBlank) {
                blankCols.push(range.Column + j);
            }
        }

        // 删除空白列
        for (const col of blankCols) {
            if (entireColumn) {
                range.Worksheet.Columns(col).Delete();
            } else {
                range.Worksheet.Range(
                    range.Worksheet.Cells(range.Row, col),
                    range.Worksheet.Cells(range.Row + range.Rows.Count - 1, col)
                ).Delete();
            }
        }
        return this;
    }

    /**
     * 删除指定单元格区域中的所有空白列（英文别名，支持链式调用）
     * @param {boolean} entireColumn - 默认删除整列
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    delBlankCols(entireColumn = false) {
        return this.z删除空白列(entireColumn);
    }

    // ==================== 工具函数 ====================

    /**
     * 将数字列号转换为字母表示
     * @param {number} col - 要转换的数字列号
     * @returns {string} 列号的字母表示
     * @example
     * const rng = new RngUtils();
     * console.log(rng.z列号字母互转(3)); // "C"
     */
    z列号字母互转(col) {
        let result = '';
        while (col > 0) {
            col--;
            result = String.fromCharCode(65 + (col % 26)) + result;
            col = Math.floor(col / 26);
        }
        return result;
    }

    /**
     * 将数字列号转换为字母表示（英文别名）
     * @param {number} col - 要转换的数字列号
     * @returns {string} 列号的字母表示
     */
    colToAbc(col) {
        return this.z列号字母互转(col);
    }

    /**
     * union函数的增强版：对字符串地址或者单元格数组联合成一个单元格区域（支持链式调用）
     * @param {any} rng - 单元格地址或单元格数组
     * @param {Sheet} opSheet - 工作表对象，跨表的时候可以指定表，默认为ActiveSheet
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils();
     * rng.rng(rng.z联合区域('a1,a2,B4:C10'));
     * const rng2 = new clsRngUtils();
     * rng2.rng(rng2.z联合区域([Range('A1:C1'),Range("D1:D10")]));
     */
    z联合区域(rng, opSheet = null) {
        const sheet = opSheet || Application.ActiveSheet;

        if (typeof rng === 'string') {
            // 字符串地址
            const addresses = rng.split(',');
            let unionRange = null;
            for (const addr of addresses) {
                const r = sheet.Range(addr.trim());
                if (unionRange === null) {
                    unionRange = r;
                } else {
                    unionRange = Application.Union(unionRange, r);
                }
            }
            this.range = unionRange;
        } else if (Array.isArray(rng)) {
            // 单元格数组
            let unionRange = null;
            for (const r of rng) {
                if (unionRange === null) {
                    unionRange = r;
                } else {
                    unionRange = Application.Union(unionRange, r);
                }
            }
            this.range = unionRange;
        } else {
            this.range = rng;
        }
        return this;
    }

    /**
     * union函数的增强版（英文别名，支持链式调用）
     * @param {any} rng - 单元格地址或单元格数组
     * @param {Sheet} opSheet - 工作表对象
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    unionAll(rng, opSheet = null) {
        return this.z联合区域(rng, opSheet);
    }

    /**
     * 根据指定单元格区域获取最大行数组
     * @param {Range|string} rng - 单元格区域
     * @param {number} col - 选择列作为获取最大行依据，默认为全部列中的最大行
     * @returns {Array} 结果二维数组
     * @example
     * const rng = new RngUtils();
     * var arr = rng.z最大行数组(Range("a:d")); // a到d列最大行数组
     * var arr2 = rng.z最大行数组("a:d", 1); // 以第1列的最大行
     */
    z最大行数组(rng, col = null) {
        const range = this._toRange(rng);

        if (col !== null && col > 0) {
            // 以指定列的最大行
            const colRange = range.Worksheet.Columns(range.Column + col - 1);
            const maxRow = this.z最大行(colRange);
            const endRange = range.Worksheet.Range(
                range.Worksheet.Cells(range.Row, range.Column),
                range.Worksheet.Cells(maxRow, range.Column + range.Columns.Count - 1)
            );
            return this.z安全数组(endRange);
        } else {
            // 全部列中的最大行
            const safeRng = this.z安全区域(range);
            const maxRow = safeRng.Row + safeRng.Rows.Count - 1;
            const endRange = range.Worksheet.Range(
                range.Worksheet.Cells(range.Row, range.Column),
                range.Worksheet.Cells(maxRow, range.Column + range.Columns.Count - 1)
            );
            return this.z安全数组(endRange);
        }
    }

    /**
     * 根据指定单元格区域获取最大行数组（英文别名）
     * @param {Range|string} rng - 单元格区域
     * @param {number} col - 选择列作为获取最大行依据
     * @returns {Array} 结果二维数组
     */
    maxArray(rng, col = null) {
        return this.z最大行数组(rng, col);
    }

    /**
     * 按指定条件查找单元格
     * @param {Range|string} rng - 单元格对象
     * @param {any} args - 参数数组 默认按单个值 也可以传完整的Range.Find函数对应的参数
     * @returns {Array} 结果单元格一维数组
     * @example
     * const rng = new RngUtils();
     * var rs = rng.z查找单元格(Range("a11:d17"), '北京'); // 查找为北京的结果单元格数组
     * console.log(rs.unionAll().Address()); // 使用unionAll函数组合为多个单元格
     */
    z查找单元格(rng, args) {
        const range = this._toRange(rng);
        const results = [];

        let findArgs = {};
        if (typeof args === 'string' || typeof args === 'number') {
            findArgs = { What: args };
        } else {
            findArgs = args;
        }

        let found = range.Find(findArgs.What, findArgs.After, findArgs.LookIn,
                                findArgs.LookAt, findArgs.SearchOrder,
                                findArgs.SearchDirection, findArgs.MatchCase,
                                findArgs.MatchByte, findArgs.SearchFormat);

        let firstFound = null;

        while (found) {
            if (!firstFound) {
                firstFound = found;
            }

            results.push(found);

            found = range.FindNext(found);

            if (found && found.Address === firstFound.Address) {
                break;
            }
        }

        return results;
    }

    /**
     * 按指定条件查找单元格（英文别名）
     * @param {Range|string} rng - 单元格对象
     * @param {any} args - 参数数组
     * @returns {Array} 结果单元格一维数组
     */
    findRange(rng, args) {
        return this.z查找单元格(rng, args);
    }

    /**
     * 检测指定单元格是否在指定单元格区域中
     * @param {Range|string} target - 待检测的单元格
     * @param {Range|string} checkRange - 上面待检测的单元格是否在本检测单元格区域中
     * @param {function} callback - 如果存在，要执行的操作
     * @returns {boolean} 检测结果 存在: true 不存在: false
     * @example
     * const rng = new RngUtils();
     * var rs = rng.z命中单元格('c3', 'a1:d10'); // c3是不是在a1:d10的地址范围内
     * console.log(rs); // true
     * rng.z命中单元格(Range('c3'), Range('a1:d10'), x => console.log(x.Address())); // 输出 $C$3
     */
    z命中单元格(target, checkRange, callback = null) {
        const targetCell = this._toRange(target);
        const checkRng = this._toRange(checkRange);

        const result = !Application.Intersect(targetCell, checkRng) === null;

        if (result && callback) {
            callback(targetCell);
        }

        return result;
    }

    /**
     * 检测指定单元格是否在指定单元格区域中（英文别名）
     * @param {Range|string} target - 待检测的单元格
     * @param {Range|string} checkRange - 检测单元格区域
     * @param {function} callback - 命中回调函数
     * @returns {boolean} 检测结果
     */
    hitRange(target, checkRange, callback = null) {
        return this.z命中单元格(target, checkRange, callback);
    }

    /**
     * 打开多文件时，返回本文件当前表的指定单元格
     * @param {string} address - 单元格地址，注意是字符串形式，如 "a3"
     * @returns {Range} 本文件当前表的指定单元格
     * @example
     * const rng = new RngUtils();
     * var rs = rng.z本文件单元格("a3"); // 多文件打开时，指定到本文件当前表的a3
     * console.log(rs.Parent.Name); // 输出当前表a3的工作表名称
     */
    z本文件单元格(address) {
        return Application.ActiveSheet.Range(address);
    }

    /**
     * 单元格不连续列装入数组 提高不连续列数组加载速度
     * @param {Range|string} rng - 单元格对象 也可以传地址
     * @param {string} cols - 不连续列号 "f1,f3-f5,f8"
     * @returns {Array} 二维数组
     * @example
     * const rng = new RngUtils();
     * var rs = rng.z选择不连续列数组(Cells, 'f1,f3-f5'); // 选择整表有效区域的A列和C:E列装入数组
     * logjson(rs);
     */
    z选择不连续列数组(rng, cols) {
        const range = this._toRange(rng);
        const colSpecs = cols.split(',');

        let resultArray = [];

        for (const spec of colSpecs) {
            const trimmedSpec = spec.trim();

            if (trimmedSpec.includes('-')) {
                // 处理范围，如 f3-f5
                const parts = trimmedSpec.replace('f', '').split('-');
                const startCol = parseInt(parts[0]);
                const endCol = parseInt(parts[1]);

                for (let c = startCol; c <= endCol; c++) {
                    const colArray = this.z安全数组(range.Columns(c));
                    resultArray = this._mergeArrays(resultArray, colArray);
                }
            } else {
                // 处理单列，如 f1
                const col = parseInt(trimmedSpec.replace('f', ''));
                const colArray = this.z安全数组(range.Columns(col));
                resultArray = this._mergeArrays(resultArray, colArray);
            }
        }

        return resultArray;
    }

    /**
     * 单元格不连续列装入数组（英文别名）
     * @param {Range|string} rng - 单元格对象
     * @param {string} cols - 不连续列号
     * @returns {Array} 二维数组
     */
    selectColsArray(rng, cols) {
        return this.z选择不连续列数组(rng, cols);
    }

    /**
     * 合并数组（私有辅助函数）
     * @private
     * @param {Array} arr1 - 第一个数组
     * @param {Array} arr2 - 第二个数组
     * @returns {Array} 合并后的数组
     */
    _mergeArrays(arr1, arr2) {
        if (arr1.length === 0) {
            return arr2;
        }

        const result = [];
        for (let i = 0; i < Math.max(arr1.length, arr2.length); i++) {
            const row1 = arr1[i] || [];
            const row2 = arr2[i] || [];
            result.push([...row1, ...row2]);
        }
        return result;
    }

    // ==================== 排序和筛选函数 ====================

    /**
     * 单元格多列排序函数（支持链式调用）
     * @param {string} sortParams - 排序参数，如 'f3+,f4-' 表示第3列升序，第4列降序
     * @param {number} headerRows - 表头的行数，默认为1
     * @param {string} customOrder - 自定义序列，如 "上海,北京,南京,海南,西藏"
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils("A18:D24");
     * rng.z多列排序('f3+,f4-', 1); // 第3列升序，第4列降序，1行表头
     * const rng2 = new clsRngUtils("A18:D24");
     * rng2.z多列排序('2+,4-', 1, "上海,北京,南京,海南,西藏"); // 带自定义序列
     */
    z多列排序(sortParams, headerRows = 1, customOrder = "") {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        // 解析排序参数
        const params = sortParams.split(',').map(p => p.trim().toLowerCase());
        const sortFields = [];

        for (const param of params) {
            const match = param.match(/f?(\d+)([+-])?/);
            if (match) {
                const colIndex = parseInt(match[1]) - 1; // 转换为0-based索引
                const ascending = match[2] !== '-';
                sortFields.push({ colIndex, ascending });
            }
        }

        // 自定义排序顺序
        let customOrderArray = [];
        if (customOrder) {
            customOrderArray = customOrder.split(',').map(s => s.trim());
        }

        // 分离表头和数据
        const header = arr.slice(0, headerRows);
        const data = arr.slice(headerRows);

        // 排序数据
        data.sort((a, b) => {
            for (const field of sortFields) {
                const valA = a[field.colIndex];
                const valB = b[field.colIndex];

                // 处理自定义序列
                if (customOrderArray.length > 0) {
                    // 转换为字符串进行比较，确保类型匹配
                    const strA = String(valA || '');
                    const strB = String(valB || '');
                    const indexA = customOrderArray.indexOf(strA);
                    const indexB = customOrderArray.indexOf(strB);

                    if (indexA !== -1 && indexB !== -1) {
                        if (indexA !== indexB) {
                            return field.ascending ? indexA - indexB : indexB - indexA;
                        }
                    } else if (indexA !== -1) {
                        return -1;
                    } else if (indexB !== -1) {
                        return 1;
                    }
                }

                // 常规比较
                if (valA !== valB) {
                    if (typeof valA === 'number' && typeof valB === 'number') {
                        return field.ascending ? valA - valB : valB - valA;
                    } else {
                        const cmp = String(valA).localeCompare(String(valB));
                        if (cmp !== 0) {
                            return field.ascending ? cmp : -cmp;
                        }
                    }
                }
            }
            return 0;
        });

        // 合并表头和排序后的数据
        const result = [...header, ...data];

        // 写回单元格
        range.Value2 = result;
        return this;
    }

    /**
     * 单元格多列排序函数（英文别名，支持链式调用）
     * @param {string} sortParams - 排序参数
     * @param {number} headerRows - 表头的行数
     * @param {string} customOrder - 自定义序列
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    rngSortCols(sortParams, headerRows = 1, customOrder = "") {
        return this.z多列排序(sortParams, headerRows, customOrder);
    }

    /**
     * 单元格强力筛选函数（支持链式调用）
     * @param {...any} args - 多参数(列,条件回调,列,条件回调.....)
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     * @example
     * const rng = new clsRngUtils(Range("A18:D24"));
     * rng.z强力筛选(2, x => x == '北京', 4, x => x > 500); // 第2列城市为北京，第4列金额>500
     */
    z强力筛选(...args) {
        const range = this._toRange(this.range);
        const arr = this.z安全数组(range);

        // 解析参数
        const filters = [];
        for (let i = 0; i < args.length; i += 2) {
            const col = args[i] - 1; // 转换为0-based索引
            const predicate = args[i + 1];
            filters.push({ col, predicate });
        }

        // 筛选数据
        const filtered = arr.filter(row => {
            for (const filter of filters) {
                const value = row[filter.col];
                if (!filter.predicate(value)) {
                    return false;
                }
            }
            return true;
        });

        // 清空原区域
        range.ClearContents();

        // 写入筛选后的数据
        if (filtered.length > 0) {
            const targetRange = range.Worksheet.Range(
                range.Worksheet.Cells(range.Row, range.Column),
                range.Worksheet.Cells(range.Row + filtered.length - 1, range.Column + filtered[0].length - 1)
            );
            targetRange.Value2 = filtered;
        }
        return this;
    }

    /**
     * 单元格强力筛选函数（英文别名，支持链式调用）
     * @param {...any} args - 多参数
     * @returns {clsRngUtils} 返回当前实例以支持链式调用
     */
    rngFilter(...args) {
        return this.z强力筛选(...args);
    }
}

/**
 * 创建RngUtils全局实例 - 可直接使用 RngUtils.方法名() 调用
 * @example
 * RngUtils.z多列排序("A1:J15", "3+,6+,4+", 1);
 */
const RngUtils = new clsRngUtils();

/**
 * 导出RngUtils - 支持WPS JSA环境
 */
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { RngUtils };
}
// WPS JSA环境：导出为全局变量
if (typeof window !== 'undefined' || typeof Application !== 'undefined') {
    this.RngUtils = RngUtils;
}

/**
 * Global类 - 全局函数工具类
 * 提供类型判断、转换和快捷方式函数
 *
 * @class
 * @description 郑广学JSA880快速开发框架中的全局函数工具
 * @example
 * // 使用全局函数
 * var range = $("A1");
 * var arr = asArray([1, 2, 3]);
 */
class clsGlobal {
    constructor() {
        this.MODULE_NAME = "Global";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学JSA880框架";
    }

    /**
     * 快速打开JSA880帮助
     * @param {string} fxname - 函数名，如"Array2D.pad"
     * @example
     * f1("Array2D.pad")
     */
    f1(fxname) {
        Console.log(`打开帮助: ${fxname}`);
        // 在实际WPS环境中，这里会打开对应的帮助文档
    }

    /**
     * Range和Cell的快捷方式
     * @param {any} x - 单元格地址('A1')、单个数字(1)、两个数字(1,1)或单元格对象
     * @returns {Range} 单元格对象
     * @example
     * var range = $("a5");
     * console.log(range.Address()); // $A$5
     * var range2 = $(5,1);  // 第5行, 第1列
     * console.log(range2.Address()); // $A$5
     */
    $(x) {
        if (typeof x === 'string') {
            return Range(x);
        } else if (typeof x === 'number') {
            // 这里需要处理 $(1,1) 的情况
            // 但由于JavaScript不支持函数重载，需要使用arguments
            if (arguments.length === 2) {
                return Cells(arguments[0], arguments[1]);
            }
            return Cells(x, 1);
        } else if (x && x.Address) {
            // 已经是Range对象
            return x;
        }
        throw new Error("无效的参数类型");
    }

    /**
     * WorksheetFunction对象的简写
     * @param {string} path - 函数对象的路径
     * @returns {function} worksheetfunction
     * @example
     * var rs = $fx.Sum(1,2,3);
     * console.log(rs); // 6
     */
    $fx(path) {
        return Application.WorksheetFunction[path];
    }

    /**
     * 将参数转换为数组（内部使用）
     * @param {...any} args - 要转换为数组的参数
     * @returns {Array} 转换后的数组
     * @example
     * var array = $toArray("产品1", "产品2", "产品3", "产品4");
     * logjson(array,0); // ["产品1","产品2","产品3","产品4"]
     */
    $toArray(...args) {
        return args;
    }

    /**
     * 把无类型提示的数组包装为有提示的数组
     * @param {any} o - 要转换为数组的对象
     * @returns {Array} 转换后的数组
     * @example
     * var array = asArray([1,2,3,4,5]);
     * logjson(array); // [1,2,3,4,5]
     */
    asArray(o) {
        if (Array.isArray(o)) {
            return o;
        }
        return [o];
    }

    /**
     * 将日期字符串或js日期对象转换为日期对象
     * @param {any} d - 要转换的日期字符串或者日期变量
     * @returns {Date} 转换后的js日期对象
     * @example
     * var date = asDate("2023-9-21");
     * console.log(date.format()); // 2023-09-21
     */
    asDate(d) {
        if (d instanceof Date) {
            return d;
        }
        if (typeof d === 'string' || typeof d === 'number') {
            return new Date(d);
        }
        throw new Error("无法转换为日期对象");
    }

    /**
     * 将字符串转换为数字对象
     * @param {string} s - 要转换的字符串
     * @returns {Number} 转换后的数字对象
     * @example
     * var number = asNumber(555);
     * console.log(number); // 555
     */
    asNumber(s) {
        const num = Number(s);
        if (isNaN(num)) {
            throw new Error("无法转换为数字");
        }
        return num;
    }

    /**
     * 将对象转换为单元格对象
     * @param {any} rng - 要转换的对象
     * @returns {Range} 转换后的单元格对象
     * @example
     * var range = asRange("b5");
     * console.log(range.Address()); // $B$5
     * var range2 = asRange(5,2);
     * console.log(range2.Address()); // $B$5
     */
    asRange(rng) {
        if (typeof rng === 'string') {
            return Range(rng);
        } else if (typeof rng === 'number') {
            if (arguments.length === 2) {
                return Cells(arguments[0], arguments[1]);
            }
            return Cells(rng, 1);
        } else if (rng && rng.Address) {
            return rng;
        }
        throw new Error("无法转换为Range对象");
    }

    /**
     * 将对象转换为形状对象
     * @param {any} shp - 要转换的对象
     * @returns {Shape} 转换后的形状对象
     * @example
     * var shape = asShape('矩形 2');
     * console.log(shape.Name); // Rectangle 2
     */
    asShape(shp) {
        if (typeof shp === 'string') {
            return ActiveSheet.Shapes(shp);
        } else if (shp && shp.Name) {
            return shp;
        }
        throw new Error("无法转换为Shape对象");
    }

    /**
     * 将对象转换为工作表对象
     * @param {any} sht - 要转换的对象
     * @returns {Sheet} 转换后的工作表对象
     * @example
     * var sheet = asSheet("1月");
     * console.log(sheet.Name); // 1月
     */
    asSheet(sht) {
        if (typeof sht === 'string' || typeof sht === 'number') {
            return Sheets(sht);
        } else if (sht && sht.Name) {
            return sht;
        }
        throw new Error("无法转换为Sheet对象");
    }

    /**
     * 将对象转换为字符串对象
     * @param {any} s - 要转换的对象
     * @returns {String} 转换后的字符串对象
     * @example
     * var str = asString('测试');
     * console.log(str); // 测试
     */
    asString(s) {
        return String(s);
    }

    /**
     * 将对象转换为工作簿对象
     * @param {any} wbk - 要转换的对象
     * @returns {Workbook} 转换后的工作簿对象
     * @example
     * var workbook = asWorkbook("测试排序");
     * console.log(workbook.Name); // 测试排序.xlsm
     */
    asWorkbook(wbk) {
        if (typeof wbk === 'string' || typeof wbk === 'number') {
            return Workbooks(wbk);
        } else if (wbk && wbk.Name) {
            return wbk;
        }
        throw new Error("无法转换为Workbook对象");
    }

    /**
     * 将日期转换为数字对象
     * @param {any} v - 日期字符串或js日期对象
     * @returns {Number} 转换后的日期数值
     * @example
     * var num = cdate('23-9-1');
     * console.log(num); // 45170
     */
    cdate(v) {
        if (v instanceof Date) {
            // 将Date转换为Excel日期序列号
            const excelEpoch = new Date(1900, 0, 1);
            const msPerDay = 24 * 60 * 60 * 1000;
            return Math.floor((v - excelEpoch) / msPerDay) + 2;
        }
        if (typeof v === 'string') {
            const date = new Date(v);
            return this.cdate(date);
        }
        throw new Error("无法转换为日期数值");
    }

    /**
     * 将值转换为字符串对象
     * @param {any} v - 要转换的值
     * @returns {String} 转换后的字符串对象
     * @example
     * var str = cstr(1537789);
     * console.log(str); // 1537789
     */
    cstr(v) {
        return String(v);
    }

    /**
     * 检查值是否为数组
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是数组，则返回true；否则返回false
     * @example
     * var isArr = isArray([1,2,3,4,5]);
     * console.log(isArr); // true
     */
    isArray(v) {
        return Array.isArray(v);
    }

    /**
     * 检查值是否为二维数组
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是二维数组，则返回true；否则返回false
     * @example
     * var is2DArray = isArray2D([[1],[2],[3]]);
     * console.log(is2DArray); // true
     */
    isArray2D(v) {
        if (!Array.isArray(v)) {
            return false;
        }
        // 检查是否所有元素都是数组
        return v.length > 0 && Array.isArray(v[0]);
    }

    /**
     * 检查值是否为布尔值
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是布尔值，则返回true；否则返回false
     * @example
     * var isBool = isBoolean(false);
     * console.log(isBool); // true
     */
    isBoolean(v) {
        return typeof v === 'boolean';
    }

    /**
     * 检查对象是否为集合对象
     * @param {any} obj - 要检查的对象
     * @returns {Boolean} 如果对象是可遍历的集合对象，则返回true；否则返回false
     * @example
     * var isColl = isCollection(Sheets);
     * console.log(isColl); // true
     */
    isCollection(obj) {
        // 检查是否有Count和Item属性（COM集合的特征）
        return obj && typeof obj.Count === 'number' && typeof obj.Item === 'function';
    }

    /**
     * 检查值是否为日期对象
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是日期对象，则返回true；否则返回false
     * @example
     * var isDateObject = isDate(new Date());
     * console.log(isDateObject); // true
     */
    isDate(v) {
        return v instanceof Date;
    }

    /**
     * 检查值是否为空值
     * @param {any} value - 要检查的值
     * @returns {Boolean} 如果值为空值（undefined、null、空字符串），则返回true；否则返回false
     * @example
     * console.log(isEmpty(undefined)); // true
     * console.log(isEmpty('')); // true
     * console.log(isEmpty(null)); // true
     */
    isEmpty(value) {
        return value === undefined || value === null || value === '';
    }

    /**
     * 检查值是否为数值类型
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是数值类型，则返回true；否则返回false
     * @example
     * var n = isNumberic(557);
     * console.log(n); // true
     */
    isNumberic(v) {
        return typeof v === 'number' && !isNaN(v);
    }

    /**
     * 检查值是否为单元格对象
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是单元格对象，则返回true；否则返回false
     * @example
     * var isRng = isRange(Range("A1"));
     * console.log(isRng); // true
     */
    isRange(v) {
        return v && typeof v.Address === 'function' && typeof v.Value2 !== 'undefined';
    }

    /**
     * 检查值是否为正则表达式对象
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是正则表达式对象，则返回true；否则返回false
     * @example
     * var isRx = isRegex(/\d+/g);
     * console.log(isRx); // true
     */
    isRegex(v) {
        return v instanceof RegExp;
    }

    /**
     * 检查两个值是否属于同一类别
     * @param {any} x - 要比较的第一个对象
     * @param {any} y - 要比较的第二个对象
     * @returns {Boolean} 如果两个值属于同一类别，则返回true；否则返回false
     * @example
     * var isSame = isSameClass(560, 789);
     * console.log(isSame); // true
     */
    isSameClass(x, y) {
        return Object.prototype.toString.call(x) === Object.prototype.toString.call(y);
    }

    /**
     * 检查值是否为工作表对象
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是工作表对象，则返回true；否则返回false
     * @example
     * var isSht = isSheet(Sheets(1));
     * console.log(isSht); // true
     */
    isSheet(v) {
        return v && typeof v.Name === 'string' && typeof v.Range === 'function';
    }

    /**
     * 检查值是否为字符串类型
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是字符串类型，则返回true；否则返回false
     * @example
     * var isstr = isString('产品5');
     * console.log(isstr); // true
     */
    isString(v) {
        return typeof v === 'string';
    }

    /**
     * 检查值是否为工作簿对象
     * @param {any} v - 要检查的值
     * @returns {Boolean} 如果值是工作簿对象，则返回true；否则返回false
     * @example
     * var iswbk = isWorkbook(ActiveWorkbook);
     * console.log(iswbk); // true
     */
    isWorkbook(v) {
        return v && typeof v.Name === 'string' && typeof v.Sheets === 'object';
    }

    /**
     * 输出日志信息
     * @param {...any} args - 要输出的日志信息
     * @example
     * log('测试日期',0,888,'中文');
     * // 输出:
     * // 测试日期
     * // 0
     * // 888
     * // 中文
     */
    log(...args) {
        args.forEach(arg => {
            Console.log(arg);
        });
    }

    /**
     * 输出JSON格式的日志信息
     * @param {any} x - 要输出的JSON对象
     * @param {boolean} wrapopt - 是否包装JSON对象（默认为true）
     * @example
     * logjson([[1,2],[3,4],[5,6]],0);
     * // 输出: [[1,2],[3,4],[5,6]]
     */
    logjson(x, wrapopt = true) {
        if (wrapopt) {
            // 包装JSON对象，处理日期等特殊对象
            const seen = new WeakSet();
            const replacer = (key, value) => {
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
            Console.log(JSON.stringify(x, replacer, 2));
        } else {
            Console.log(JSON.stringify(x));
        }
    }

    /**
     * 获取值的类型名称
     * @param {any} x - 要获取类型名称的值
     * @returns {String} 值的类型名称
     * @example
     * var t = typeName('产品5');
     * console.log(t); // [object String]
     */
    typeName(x) {
        return Object.prototype.toString.call(x);
    }

    /**
     * 获取数组的指定维度的上界
     * @param {Array} arr - 要获取上界的数组
     * @param {number} dimension - 要获取上界的维度（默认为1）
     * @returns {Number} 指定维度的上界
     * @example
     * var upperBound = ubound([[1,2,3],[4,5,6],[7,8,9],[10,11,12]], 2);
     * console.log(upperBound); // 2
     */
    ubound(arr, dimension = 1) {
        if (!Array.isArray(arr)) {
            throw new Error("参数必须是数组");
        }
        if (dimension === 1) {
            return arr.length - 1;
        } else if (dimension === 2 && arr.length > 0 && Array.isArray(arr[0])) {
            return arr[0].length - 1;
        }
        return -1;
    }

    /**
     * 字符串及布尔值转为数值（与VBA中的val保持一致）
     * @param {String} s - 要转换的字符串
     * @returns {Number} 转换后的数值对象
     * @example
     * var value = val('5');
     * console.log(value); // 5
     */
    val(s) {
        if (typeof s === 'boolean') {
            return s ? 1 : 0;
        }
        if (typeof s === 'number') {
            return s;
        }
        if (typeof s === 'string') {
            // 移除前导空格
            s = s.trim();
            // 提取开头的数字部分
            const match = s.match(/^[-+]?[0-9]*\.?[0-9]+/);
            if (match) {
                return parseFloat(match[0]);
            }
            return 0;
        }
        return 0;
    }

    /**
     * 使用Excel计算规则对数字进行四舍五入
     * @param {number} number - 要进行四舍五入的数字
     * @param {number} decimals - 保留的小数位数（默认为2）
     * @returns {number} 四舍五入后的结果
     * @example
     * var v = $.text(round(5.786543224,3),"0.000");
     * console.log(v); // 5.787
     */
    round(number, decimals = 2) {
        const multiplier = Math.pow(10, decimals);
        return Math.round(number * multiplier) / multiplier;
    }
}

/**
 * ChainableRange类 - 支持链式调用的Range包装器
 * @class
 * @description 包装WPS Range对象，提供链式调用支持
 * @example
 * // 链式调用示例
 * $("A1").Value("你好").Font.Color(255).Interior.Color(16777215);
 * $("A1:C10").ClearFormats().Value("数据");
 */
class clsChainableRange {
    /**
     * 构造函数
     * @param {Range} range - WPS Range对象
     */
    constructor(range) {
        this._range = range;
    }

    /**
     * 获取原始Range对象（用于某些需要原生Range的场景）
     * @returns {Range} 原始Range对象
     */
    unwrap() {
        return this._range;
    }

    /**
     * 获取值（不返回链式对象）
     * @returns {any} 单元格值
     */
    getValue() {
        return this._range.Value;
    }

    /**
     * 设置值
     * @param {any} value - 要设置的值
     * @returns {clsChainableRange} 返回链式对象
     */
    Value(value) {
        if (value !== undefined) {
            this._range.Value = value;
            return this;
        }
        return this._range.Value;
    }

    /**
     * 设置或获取值（别名）
     * @param {any} value - 要设置的值
     * @returns {clsChainableRange|any} 返回链式对象或值
     */
    val(value) {
        if (value !== undefined) {
            this._range.Value = value;
            return this;
        }
        return this._range.Value;
    }

    /**
     * 设置公式
     * @param {string} formula - 公式字符串
     * @returns {clsChainableRange} 返回链式对象
     */
    Formula(formula) {
        this._range.Formula = formula;
        return this;
    }

    /**
     * 清除内容
     * @returns {clsChainableRange} 返回链式对象
     */
    ClearContents() {
        this._range.ClearContents();
        return this;
    }

    /**
     * 清除格式
     * @returns {clsChainableRange} 返回链式对象
     */
    ClearFormats() {
        this._range.ClearFormats();
        return this;
    }

    /**
     * 清除全部
     * @returns {clsChainableRange} 返回链式对象
     */
    Clear() {
        this._range.Clear();
        return this;
    }

    /**
     * 选择区域
     * @returns {clsChainableRange} 返回链式对象
     */
    Select() {
        this._range.Select();
        return this;
    }

    /**
     * 复制
     * @param {Range} [destination] - 目标区域
     * @returns {clsChainableRange} 返回链式对象
     */
    Copy(destination) {
        if (destination) {
            this._range.Copy(destination);
        } else {
            this._range.Copy();
        }
        return this;
    }

    /**
     * 粘贴
     * @param {Range} [destination] - 目标区域
     * @returns {clsChainableRange} 返回链式对象
     */
    Paste(destination) {
        if (destination) {
            destination.PasteSpecial();
        } else {
            this._range.Parent.Paste();
        }
        return this;
    }

    /**
     * 删除
     * @param {string} [shift] - 移动方向
     * @returns {clsChainableRange} 返回链式对象
     */
    Delete(shift) {
        this._range.Delete(shift);
        return this;
    }

    /**
     * 插入
     * @param {string} [shift] - 移动方向
     * @returns {clsChainableRange} 返回链式对象
     */
    Insert(shift) {
        this._range.Insert(shift);
        return this;
    }

    /**
     * 获取地址
     * @returns {string} 单元格地址
     */
    Address() {
        return this._range.Address;
    }

    /**
     * 自动列宽
     * @returns {clsChainableRange} 返回链式对象
     */
    AutoFit() {
        this._range.Columns.AutoFit();
        return this;
    }

    /**
     * 自动行高
     * @returns {clsChainableRange} 返回链式对象
     */
    AutoFitRows() {
        this._range.Rows.AutoFit();
        return this;
    }

    /**
     * 代理其他属性访问
     */
    get Font() {
        const font = this._range.Font;
        return new Proxy(this, {
            get: (target, prop) => {
                if (prop === 'set') {
                    return (props) => {
                        for (let key in props) {
                            font[key] = props[key];
                        }
                        return target;
                    };
                }
                const value = font[prop];
                if (typeof value === 'function') {
                    return value.bind(font);
                }
                // 如果是设置值
                return function(...args) {
                    if (args.length > 0) {
                        font[prop] = args[0];
                        return target;
                    }
                    return value;
                };
            },
            set: (target, prop, value) => {
                font[prop] = value;
                return true;
            }
        });
    }

    get Interior() {
        const interior = this._range.Interior;
        return new Proxy(this, {
            get: (target, prop) => {
                if (prop === 'set') {
                    return (props) => {
                        for (let key in props) {
                            interior[key] = props[key];
                        }
                        return target;
                    };
                }
                const value = interior[prop];
                if (typeof value === 'function') {
                    return value.bind(interior);
                }
                return function(...args) {
                    if (args.length > 0) {
                        interior[prop] = args[0];
                        return target;
                    }
                    return value;
                };
            },
            set: (target, prop, value) => {
                interior[prop] = value;
                return true;
            }
        });
    }

    get Borders() {
        const borders = this._range.Borders;
        return new Proxy(this, {
            get: (target, prop) => {
                if (prop === 'set') {
                    return (props) => {
                        for (let key in props) {
                            borders[key] = props[key];
                        }
                        return target;
                    };
                }
                const value = borders[prop];
                if (typeof value === 'function') {
                    return value.bind(borders);
                }
                return function(...args) {
                    if (args.length > 0) {
                        borders[prop] = args[0];
                        return target;
                    }
                    return value;
                };
            },
            set: (target, prop, value) => {
                borders[prop] = value;
                return true;
            }
        });
    }

    get Columns() {
        return this._range.Columns;
    }

    get Rows() {
        return this._range.Rows;
    }

    get Cells() {
        return this._range.Cells;
    }

    get Count() {
        return this._range.Count;
    }

    get Column() {
        return this._range.Column;
    }

    get Row() {
        return this._range.Row;
    }

    get Worksheet() {
        return this._range.Worksheet;
    }

    /**
     * 获取偏移区域
     * @param {number} rowOffset - 行偏移
     * @param {number} columnOffset - 列偏移
     * @returns {clsChainableRange} 返回链式对象
     */
    Offset(rowOffset, columnOffset) {
        return new clsChainableRange(this._range.Offset(rowOffset, columnOffset));
    }

    /**
     * 获取调整大小后的区域
     * @param {number} rowSize - 行数
     * @param {number} columnSize - 列数
     * @returns {clsChainableRange} 返回链式对象
     */
    Resize(rowSize, columnSize) {
        return new clsChainableRange(this._range.Resize(rowSize, columnSize));
    }

    /**
     * 合并单元格
     * @returns {clsChainableRange} 返回链式对象
     */
    Merge() {
        this._range.Merge();
        return this;
    }

    /**
     * 取消合并
     * @returns {clsChainableRange} 返回链式对象
     */
    UnMerge() {
        this._range.UnMerge();
        return this;
    }

    /**
     * 设置水平对齐
     * @param {string} align - 对齐方式
     * @returns {clsChainableRange} 返回链式对象
     */
    setHorizontalAlignment(align) {
        this._range.HorizontalAlignment = align;
        return this;
    }

    /**
     * 设置垂直对齐
     * @param {string} align - 对齐方式
     * @returns {clsChainableRange} 返回链式对象
     */
    setVerticalAlignment(align) {
        this._range.VerticalAlignment = align;
        return this;
    }

    /**
     * 设置列宽
     * @param {number} width - 列宽
     * @returns {clsChainableRange} 返回链式对象
     */
    setColumnWidth(width) {
        this._range.ColumnWidth = width;
        return this;
    }

    /**
     * 设置行高
     * @param {number} height - 行高
     * @returns {clsChainableRange} 返回链式对象
     */
    setRowHeight(height) {
        this._range.RowHeight = height;
        return this;
    }

    /**
     * 设置数字格式
     * @param {string} format - 格式字符串
     * @returns {clsChainableRange} 返回链式对象
     */
    setNumberFormat(format) {
        this._range.NumberFormat = format;
        return this;
    }

    /**
     * 设置文本换行
     * @param {boolean} wrap - 是否换行
     * @returns {clsChainableRange} 返回链式对象
     */
    setWrapText(wrap) {
        this._range.WrapText = wrap;
        return this;
    }

    /**
     * 调用Range上的任意方法并返回链式对象
     */
    call(methodName, ...args) {
        if (typeof this._range[methodName] === 'function') {
            this._range[methodName](...args);
        }
        return this;
    }

    /**
     * 获取Range上的任意属性
     */
    get(propName) {
        return this._range[propName];
    }

    /**
     * 设置Range上的任意属性
     */
    set(propName, value) {
        this._range[propName] = value;
        return this;
    }
}

/**
 * 全局 $ 函数 - Range和Cell的快捷方式（支持链式调用）
 * @param {any} x - 单元格地址('A1')、单个数字(1)、两个数字(1,1)或单元格对象
 * @returns {clsChainableRange} 支持链式调用的单元格对象
 * @example
 * // 基本使用
 * var range = $("a5");
 * console.log(range.Address()); // $A$5
 * var range2 = $(5,1);  // 第5行, 第1列
 * console.log(range2.Address()); // $A$5
 *
 * // 链式调用
 * $("A1").Value("你好").Font.Color(255).Interior.Color(16777215);
 * $("A1:C10").ClearFormats().Value("数据");
 * $("D1").Value(100).setNumberFormat("0.00").setHorizontalAlignment(-4108); // xlCenter
 */
function $(x) {
    let range;
    if (typeof x === 'string') {
        range = Range(x);
    } else if (typeof x === 'number') {
        // 处理 $(1,1) 的情况
        if (arguments.length === 2) {
            range = Cells(arguments[0], arguments[1]);
        } else {
            range = Cells(x, 1);
        }
    } else if (x && x.Address) {
        // 已经是Range对象
        range = x;
    } else {
        throw new Error("无效的参数类型");
    }
    // 返回支持链式调用的包装对象
    return new clsChainableRange(range);
}

/**
 * 全局 $fx 变量 - WorksheetFunction对象的简写
 */
const $fx = Application.WorksheetFunction;

/**
 * 创建Global实例（用于内部方法调用）
 */
const _GlobalInstance = new clsGlobal();

/**
 * 导出Global类 - 支持WPS JSA环境
 */
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { Global: clsGlobal, ChainableRange: clsChainableRange, $, $fx };
}
// WPS JSA环境：导出为全局变量
if (typeof window !== 'undefined' || typeof Application !== 'undefined') {
    this.Global = clsGlobal;
    this.ChainableRange = clsChainableRange;
}

// 全局函数（独立函数，方便直接调用）
function asArray(o) { return _GlobalInstance.asArray(o); }
function asDate(d) { return _GlobalInstance.asDate(d); }
function asNumber(s) { return _GlobalInstance.asNumber(s); }
function asRange(rng) { return _GlobalInstance.asRange(rng); }
function asShape(shp) { return _GlobalInstance.asShape(shp); }
function asSheet(sht) { return _GlobalInstance.asSheet(sht); }
function asString(s) { return _GlobalInstance.asString(s); }
function asWorkbook(wbk) { return _GlobalInstance.asWorkbook(wbk); }
function cdate(v) { return _GlobalInstance.cdate(v); }
function cstr(v) { return _GlobalInstance.cstr(v); }
function isArray(v) { return _GlobalInstance.isArray(v); }
function isArray2D(v) { return _GlobalInstance.isArray2D(v); }
function isBoolean(v) { return _GlobalInstance.isBoolean(v); }
function isCollection(obj) { return _GlobalInstance.isCollection(obj); }
function isDate(v) { return _GlobalInstance.isDate(v); }
function isEmpty(value) { return _GlobalInstance.isEmpty(value); }
function isNumberic(v) { return _GlobalInstance.isNumberic(v); }
function isRange(v) { return _GlobalInstance.isRange(v); }
function isRegex(v) { return _GlobalInstance.isRegex(v); }
function isSameClass(x, y) { return _GlobalInstance.isSameClass(x, y); }
function isSheet(v) { return _GlobalInstance.isSheet(v); }
function isString(v) { return _GlobalInstance.isString(v); }
function isWorkbook(v) { return _GlobalInstance.isWorkbook(v); }
function log(...args) { return _GlobalInstance.log(...args); }
function logjson(x, wrapopt) { return _GlobalInstance.logjson(x, wrapopt); }
function typeName(x) { return _GlobalInstance.typeName(x); }
function ubound(arr, dimension) { return _GlobalInstance.ubound(arr, dimension); }
function val(s) { return _GlobalInstance.val(s); }
function round(number, decimals) { return _GlobalInstance.round(number, decimals); }

// 如果直接运行此文件，则执行示例
if (typeof require !== 'undefined' && require.main === module) {
    示例使用Array2D();
}

/**
 * JSA类 - JSA通用函数工具库
 * 增强JSA能力全局可用
 *
 * @class
 * @description 郑广学JSA880快速开发框架中的通用函数工具库
 * @example
 * // 使用JSA函数
 * var arr = JSA.z转置([[1,2,3],[4,5,6]]);
 * var today = JSA.z今天();
 */
class clsJSA {
    constructor() {
        this.MODULE_NAME = "JSA";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学JSA880框架";
    }

    /**
     * 将数组转置（行列互换）
     * @param {Array} arr - 要转置的数组
     * @returns {Array} 转置后的数组
     * @example
     * var arr = JSA.z转置([[1,2,3],[4,5,6],[7,8,9],[10,11,12],[15,18,19]]);
     * logjson(arr,0);
     * // 输出: [[1,4,7,10,15],[2,5,8,11,18],[3,6,9,12,19]]
     */
    z转置(arr) {
        if (!arr || arr.length === 0) return [];
        const rows = arr.length;
        const cols = arr[0].length;
        const result = [];
        for (let j = 0; j < cols; j++) {
            result[j] = [];
            for (let i = 0; i < rows; i++) {
                result[j][i] = arr[i][j];
            }
        }
        return result;
    }

    /**
     * 将数组转置（英文别名）
     * @param {Array} arr - 要转置的数组
     * @returns {Array} 转置后的数组
     */
    transpose(arr) {
        return this.z转置(arr);
    }

    /**
     * 将文本转换为数值
     * @param {String} text - 要转换的文本
     * @returns {Number} 转换后的数值
     * @example
     * var v = JSA.z转数值("753");
     * console.log(v); // 753
     */
    z转数值(text) {
        if (typeof text === 'number') return text;
        if (typeof text === 'string') {
            // 移除前导空格
            text = text.trim();
            // 提取开头的数字部分
            const match = text.match(/^[-+]?[0-9]*\.?[0-9]+/);
            if (match) {
                return parseFloat(match[0]);
            }
            return 0;
        }
        return 0;
    }

    /**
     * 将文本转换为数值（英文别名）
     * @param {String} text - 要转换的文本
     * @returns {Number} 转换后的数值
     */
    val(text) {
        return this.z转数值(text);
    }

    /**
     * 将数组输出到单元格区域
     * @param {Array} arr - 数组
     * @param {Range|string} rng - 单元格区域
     * @param {Boolean} clearDown - 是否清空目标区下方数据
     * @returns {Range} 返回输出的完整单元格区域
     * @example
     * var rng = JSA.z写入单元格([[1,2,3],[4,5,6]], "g1", true);
     * console.log(rng.Address()); // $G$1:$I$2
     */
    z写入单元格(arr, rng, clearDown = false) {
        // 确保rng是Range对象
        let targetRng = typeof rng === 'string' ? Range(rng) : rng;

        // 写入数组
        targetRng.Value2 = arr;

        // 如果需要清空下方数据
        if (clearDown) {
            const lastRow = targetRng.Row + arr.length - 1;
            const ws = targetRng.Worksheet;
            const usedRange = ws.UsedRange;
            if (usedRange && lastRow < usedRange.Row + usedRange.Rows.Count) {
                const clearRng = ws.Range(
                    ws.Cells(lastRow + 1, targetRng.Column),
                    ws.Cells(usedRange.Row + usedRange.Rows.Count, targetRng.Column + arr[0].length - 1)
                );
                clearRng.ClearContents();
            }
        }

        // 返回实际写入的范围
        return targetRng.Worksheet.Range(
            targetRng,
            targetRng.Worksheet.Cells(targetRng.Row + arr.length - 1, targetRng.Column + arr[0].length - 1)
        );
    }

    /**
     * 将数组输出到单元格区域（英文别名，可作为数组扩展方法）
     * @param {Array} arr - 数组
     * @param {Range|string} rng - 单元格区域
     * @param {Boolean} clearDown - 是否清空目标区下方数据
     * @returns {Range} 返回输出的完整单元格区域
     */
    toRange(arr, rng, clearDown = false) {
        return this.z写入单元格(arr, rng, clearDown);
    }

    /**
     * 获取当前日期的字符串表示
     * @returns {String} 当前日期的字符串表示
     * @example
     * Console.log(JSA.z今天()); // 2024-10-05
     */
    z今天() {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        return year + "-" + month + "-" + day;
    }

    /**
     * 获取当前日期的字符串表示（英文别名）
     * @returns {String} 当前日期的字符串表示
     */
    today() {
        return this.z今天();
    }

    /**
     * 将日期对象转换为日期数值
     * @param {Date|string} d - 要转换的日期对象
     * @returns {Number} 转换后的日期数值
     * @example
     * Console.log(JSA.z转日期数值('2024-10-5')); // 45570
     */
    z转日期数值(d) {
        let date;
        if (typeof d === 'string') {
            date = new Date(d);
        } else if (d instanceof Date) {
            date = d;
        } else {
            throw new Error("无效的日期参数");
        }
        // 将Date转换为Excel日期序列号
        const excelEpoch = new Date(1900, 0, 1);
        const msPerDay = 24 * 60 * 60 * 1000;
        return Math.floor((date - excelEpoch) / msPerDay) + 2;
    }

    /**
     * 将日期对象转换为日期数值（英文别名）
     * @param {Date|string} d - 要转换的日期对象
     * @returns {Number} 转换后的日期数值
     */
    cdate(d) {
        return this.z转日期数值(d);
    }

    /**
     * 转义正则表达式中的特殊字符
     * @param {String} str - 要转义的字符串
     * @returns {String} 转义后的字符串
     * @example
     * Console.log(JSA.z转义正则('1*(2+1.3)')); // 1\*\(2\+1\.3\)
     */
    z转义正则(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * 转义正则表达式中的特殊字符（英文别名）
     * @param {String} str - 要转义的字符串
     * @returns {String} 转义后的字符串
     */
    escapeRegExp(str) {
        return this.z转义正则(str);
    }

    /**
     * 替换字符串中的所有指定内容
     * @param {String} str - 要替换的字符串
     * @param {String} find - 要查找的内容
     * @param {String} replaceWith - 替换后的内容
     * @returns {String} 替换后的字符串
     * @example
     * Console.log(JSA.z替换("你好，世界！世界很大，", "世界", "地球"));
     * // 输出: 你好，地球！地球很大，
     */
    z替换(str, find, replaceWith) {
        return str.split(find).join(replaceWith);
    }

    /**
     * 替换字符串中的所有指定内容（英文别名）
     * @param {String} str - 要替换的字符串
     * @param {String} find - 要查找的内容
     * @param {String} replaceWith - 替换后的内容
     * @returns {String} 替换后的字符串
     */
    replace(str, find, replaceWith) {
        return this.z替换(str, find, replaceWith);
    }

    /**
     * 截取字符串的指定部分
     * @param {String} str - 要截取的字符串
     * @param {Number} start - 起始位置（从1开始）
     * @param {Number} len - 截取的长度（可选）
     * @returns {String} 截取后的字符串
     * @example
     * Console.log(JSA.z截取字符("你好，世界！", 1, 2)); // 你好
     * Console.log(JSA.z截取字符("你好，世界！", 3)); // 好，世界！
     */
    z截取字符(str, start, len) {
        // 转换为0-based索引
        const startIndex = start - 1;
        if (len === undefined) {
            return str.substring(startIndex);
        }
        return str.substring(startIndex, startIndex + len);
    }

    /**
     * 截取字符串的指定部分（英文别名）
     * @param {String} str - 要截取的字符串
     * @param {Number} start - 起始位置（从1开始）
     * @param {Number} len - 截取的长度（可选）
     * @returns {String} 截取后的字符串
     */
    mid(str, start, len) {
        return this.z截取字符(str, start, len);
    }

    /**
     * 获取字符串的左侧指定长度的部分
     * @param {String} str - 要获取部分的字符串
     * @param {Number} len - 要获取的长度
     * @returns {String} 左侧指定长度的部分字符串
     * @example
     * Console.log(JSA.z左取字符("你好，世界！", 2)); // 你好
     */
    z左取字符(str, len) {
        return str.substring(0, len);
    }

    /**
     * 获取字符串的左侧指定长度的部分（英文别名）
     * @param {String} str - 要获取部分的字符串
     * @param {Number} len - 要获取的长度
     * @returns {String} 左侧指定长度的部分字符串
     */
    left(str, len) {
        return this.z左取字符(str, len);
    }

    /**
     * 获取字符串的右侧指定长度的部分
     * @param {String} str - 要获取部分的字符串
     * @param {Number} len - 要获取的长度
     * @returns {String} 右侧指定长度的部分字符串
     * @example
     * Console.log(JSA.z右取字符("你好，世界！", 3)); // 世界！
     */
    z右取字符(str, len) {
        return str.substring(str.length - len);
    }

    /**
     * 获取字符串的右侧指定长度的部分（英文别名）
     * @param {String} str - 要获取部分的字符串
     * @param {Number} len - 要获取的长度
     * @returns {String} 右侧指定长度的部分字符串
     */
    right(str, len) {
        return this.z右取字符(str, len);
    }

    /**
     * 格式化数字为指定格式的字符串
     * @param {Number} number - 要格式化的数字
     * @param {String} fmtstr - 格式化的格式字符串
     * @returns {String} 格式化后的字符串
     * @example
     * Console.log(JSA.z格式化(5.879, "￥0.00")); // ￥5.88
     */
    z格式化(number, fmtstr) {
        // 使用WorksheetFunction.Text函数进行格式化
        try {
            return Application.WorksheetFunction.Text(number, fmtstr);
        } catch (e) {
            // 如果失败，使用简单的格式化
            return String(number);
        }
    }

    /**
     * 格式化数字为指定格式的字符串（英文别名）
     * @param {Number} number - 要格式化的数字
     * @param {String} fmtstr - 格式化的格式字符串
     * @returns {String} 格式化后的字符串
     */
    text(number, fmtstr) {
        return this.z格式化(number, fmtstr);
    }

    /**
     * 在数组中查找指定关键字并返回匹配结果的索引（从1开始）
     * @param {String} keyword - 要查找的关键字
     * @param {Array} arr - 要查找的数组
     * @returns {Number} 匹配结果的索引（从1开始），未找到返回0
     * @example
     * Console.log(JSA.z查找索引("狐狸", ["兔子","狗","猫","猎豹","狐狸","熊"])); // 5
     */
    z查找索引(keyword, arr) {
        for (let i = 0; i < arr.length; i++) {
            if (String(arr[i]) === String(keyword)) {
                return i + 1; // 返回从1开始的索引
            }
        }
        return 0;
    }

    /**
     * 在数组中查找指定关键字（英文别名）
     * @param {String} keyword - 要查找的关键字
     * @param {Array} arr - 要查找的数组
     * @returns {Number} 匹配结果的索引
     */
    match(keyword, arr) {
        return this.z查找索引(keyword, arr);
    }

    /**
     * 在数组中左侧查找指定关键字并返回结果列的值
     * @param {String} keyword - 要查找的关键字
     * @param {Array} arr - 要查找的二维数组
     * @param {Number} resultCol - 结果列的索引（从1开始）
     * @param {Number} mode - 匹配模式：0精确匹配，1模糊匹配（默认0）
     * @param {String} errorValue - 遇到错误时返回的值（默认"#err"）
     * @returns {Object} 匹配结果的值
     * @example
     * var arr = [["苹果", "香蕉", "橙子"], ["狗", "桔猫", "兔子"]];
     * Console.log(JSA.z左侧查找('猫', arr, 3, 1)); // 兔子
     */
    z左侧查找(keyword, arr, resultCol, mode, errorValue) {
        if (mode === undefined) mode = 0;
        if (errorValue === undefined) errorValue = "#err";

        if (!arr || arr.length === 0) return errorValue;

        for (let i = 0; i < arr.length; i++) {
            const row = arr[i];
            if (!row || row.length === 0) continue;

            const firstCol = String(row[0]);

            if (mode === 0) {
                // 精确匹配
                if (firstCol === String(keyword)) {
                    if (resultCol > 0 && resultCol <= row.length) {
                        return row[resultCol - 1];
                    }
                    break;
                }
            } else {
                // 模糊匹配
                if (firstCol.indexOf(String(keyword)) !== -1) {
                    if (resultCol > 0 && resultCol <= row.length) {
                        return row[resultCol - 1];
                    }
                    break;
                }
            }
        }

        return errorValue;
    }

    /**
     * 在数组中左侧查找（英文别名，类似Excel的VLOOKUP）
     * @param {String} keyword - 要查找的关键字
     * @param {Array} arr - 要查找的二维数组
     * @param {Number} resultCol - 结果列的索引
     * @param {Number} mode - 匹配模式
     * @param {String} errorValue - 错误时返回的值
     * @returns {Object} 匹配结果的值
     */
    vlookup(keyword, arr, resultCol, mode, errorValue) {
        return this.z左侧查找(keyword, arr, resultCol, mode, errorValue);
    }

    /**
     * 将值转换为文本格式
     * @param {Any} v - 要转换的值
     * @returns {String} 转换后的文本格式
     * @example
     * var s = JSA.z转文本(578);
     * Console.log(s); // 578
     * console.log(typeof(s)); // string
     */
    z转文本(v) {
        return String(v);
    }

    /**
     * 将值转换为文本格式（英文别名）
     * @param {Any} v - 要转换的值
     * @returns {String} 转换后的文本格式
     */
    cstr(v) {
        return this.z转文本(v);
    }

    /**
     * 获取一个数的整数部分
     * @param {Number} v - 要获取整数部分的数
     * @returns {Number} 返回整数部分
     * @example
     * Console.log(JSA.z取整数(5.8957)); // 5
     */
    z取整数(v) {
        return Math.floor(v);
    }

    /**
     * 获取一个数的整数部分（英文别名）
     * @param {Number} v - 要获取整数部分的数
     * @returns {Number} 返回整数部分
     */
    cint(v) {
        return this.z取整数(v);
    }

    /**
     * 获取一个数的小数部分
     * @param {Number} v - 要获取小数部分的数
     * @returns {Number} 返回小数部分
     * @example
     * Console.log(JSA.z取小数(5.8957)); // 0.8957
     */
    z取小数(v) {
        return v - Math.floor(v);
    }

    /**
     * 获取一个数的小数部分（英文别名）
     * @param {Number} v - 要获取小数部分的数
     * @returns {Number} 返回小数部分
     */
    getDecimal(v) {
        return this.z取小数(v);
    }

    /**
     * 将数组转换为Excel公式字符串
     * @param {Array} arr - 要转换的数组
     * @returns {String} 转换后的Excel公式字符串
     * @example
     * console.log(JSA.z转公式数组([[1,2,3],[4,5,6]])); // {1,2,3;4,5,6}
     */
    z转公式数组(arr) {
        if (!arr || arr.length === 0) return "{}";

        const rows = arr.map(row => {
            if (Array.isArray(row)) {
                return row.join(",");
            } else {
                return String(row);
            }
        });

        return "{" + rows.join(";") + "}";
    }

    /**
     * 将数组转换为Excel公式字符串（英文别名）
     * @param {Array} arr - 要转换的数组
     * @returns {String} 转换后的Excel公式字符串
     */
    toExcelArray(arr) {
        return this.z转公式数组(arr);
    }

    /**
     * 增强的查找函数
     * @param {String} keyword - 要查找的关键字
     * @param {Array} lookupArr - 要查找的数组或单元格区域
     * @param {Array} resultArr - 结果数组或单元格区域
     * @param {String} errorValue - 遇到错误时返回的值（默认空字符串）
     * @returns {Object} 匹配结果的值
     * @example
     * var arr = [["苹果"], ["狗"], ["汽车"]];
     * var brr = [["香蕉"], ["桔猫"], ["火车"]];
     * Console.log(JSA.z增强查找('狗', arr, brr)); // 桔猫
     */
    z增强查找(keyword, lookupArr, resultArr, errorValue) {
        if (errorValue === undefined) errorValue = "";

        if (!lookupArr || lookupArr.length === 0) return errorValue;

        for (let i = 0; i < lookupArr.length; i++) {
            const lookupVal = lookupArr[i];
            // 处理二维数组
            const val = Array.isArray(lookupVal) ? lookupVal[0] : lookupVal;

            if (String(val) === String(keyword)) {
                if (resultArr && resultArr.length > i) {
                    const resultVal = resultArr[i];
                    return Array.isArray(resultVal) ? resultVal[0] : resultVal;
                }
                break;
            }
        }

        return errorValue;
    }

    /**
     * 增强的查找函数（英文别名，类似Excel的XLOOKUP）
     * @param {String} keyword - 要查找的关键字
     * @param {Array} lookupArr - 要查找的数组
     * @param {Array} resultArr - 结果数组
     * @param {String} errorValue - 错误时返回的值
     * @returns {Object} 匹配结果的值
     */
    xlookup(keyword, lookupArr, resultArr, errorValue) {
        return this.z增强查找(keyword, lookupArr, resultArr, errorValue);
    }

    /**
     * 对一组数值进行求和
     * @param {...Number} args - 要求和的多个数值
     * @returns {Number} 返回求和结果
     * @example
     * Console.log(JSA.z求和(5, 9, 11)); // 25
     */
    z求和(...args) {
        return args.reduce((sum, val) => sum + (Number(val) || 0), 0);
    }

    /**
     * 对一组数值进行求和（英文别名）
     * @param {...Number} args - 要求和的多个数值
     * @returns {Number} 返回求和结果
     */
    sum(...args) {
        return this.z求和(...args);
    }

    /**
     * 获取一组数值中的最大值
     * @param {...Number} args - 要取最大值的多个数值
     * @returns {Number} 返回最大值
     * @example
     * Console.log(JSA.z最大值(5, 18, 9, 11)); // 18
     */
    z最大值(...args) {
        return Math.max(...args.map(val => Number(val) || 0));
    }

    /**
     * 获取一组数值中的最大值（英文别名）
     * @param {...Number} args - 要取最大值的多个数值
     * @returns {Number} 返回最大值
     */
    max(...args) {
        return this.z最大值(...args);
    }

    /**
     * 获取一组数值中的最小值
     * @param {...Number} args - 要取最小值的多个数值
     * @returns {Number} 返回最小值
     * @example
     * Console.log(JSA.z最小值(9, 11, 5, 18)); // 5
     */
    z最小值(...args) {
        return Math.min(...args.map(val => Number(val) || 0));
    }

    /**
     * 获取一组数值中的最小值（英文别名）
     * @param {...Number} args - 要取最小值的多个数值
     * @returns {Number} 返回最小值
     */
    min(...args) {
        return this.z最小值(...args);
    }

    /**
     * 计算一组数值的平均值
     * @param {...Number} args - 要取平均值的多个数值
     * @returns {Number} 返回平均值
     * @example
     * Console.log(JSA.z平均值(9, 11, 5, 18)); // 10.75
     */
    z平均值(...args) {
        if (args.length === 0) return 0;
        const sum = this.z求和(...args);
        return sum / args.length;
    }

    /**
     * 计算一组数值的平均值（英文别名）
     * @param {...Number} args - 要取平均值的多个数值
     * @returns {Number} 返回平均值
     */
    average(...args) {
        return this.z平均值(...args);
    }

    /**
     * 模糊匹配字符串与模式（支持*通配符）
     * @param {String} str - 要匹配的字符串
     * @param {String} pattern - 匹配词（支持*通配符）
     * @returns {Boolean} 返回匹配结果
     * @example
     * Console.log(JSA.z模糊匹配("我们的世界很大", "*世界*")); // true
     */
    z模糊匹配(str, pattern) {
        // 转义正则表达式特殊字符（除了*）
        let regexPattern = pattern.replace(/[.+?^${}()|[\]\\]/g, '\\$&');
        // 将*替换为.*（匹配任意字符）
        regexPattern = regexPattern.replace(/\*/g, '.*');
        // 创建正则表达式
        const regex = new RegExp('^' + regexPattern + '$');
        return regex.test(str);
    }

    /**
     * 模糊匹配字符串（英文别名）
     * @param {String} str - 要匹配的字符串
     * @param {String} pattern - 匹配词
     * @returns {Boolean} 返回匹配结果
     */
    like(str, pattern) {
        return this.z模糊匹配(str, pattern);
    }

    /**
     * 生成一个数字序列一维数组
     * @param {Number} start - 序列的起始值
     * @param {Number} end - 序列的结束值
     * @param {Number} step - 序列的步长（默认1）
     * @returns {Array} 生成的数字序列一维数组
     * @example
     * logjson(JSA.z生成数字序列(5, 10, 2)); // [5, 7, 9]
     */
    z生成数字序列(start, end, step) {
        if (step === undefined) step = 1;

        const result = [];
        for (let i = start; i <= end; i += step) {
            result.push(i);
        }
        return result;
    }

    /**
     * 生成一个数字序列一维数组（英文别名）
     * @param {Number} start - 序列的起始值
     * @param {Number} end - 序列的结束值
     * @param {Number} step - 序列的步长
     * @returns {Array} 生成的数字序列一维数组
     */
    getNumberArray(start, end, step) {
        return this.z生成数字序列(start, end, step);
    }

    /**
     * 生成索引数组（别名）
     */
    getIndexs(start, end, step) {
        return this.z生成数字序列(start, end, step);
    }

    /**
     * 将数字转换为人民币大写
     * @param {Number} n - 要转换的数字
     * @returns {String} 转换后的人民币大写
     * @example
     * console.log(JSA.z人民币大写(594.12)); // 伍佰玖拾肆元壹角贰分
     */
    z人民币大写(n) {
        const digits = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
        const units = ["", "拾", "佰", "仟"];
        const bigUnits = ["", "万", "亿"];

        if (n === 0) return "零元整";

        const num = Math.abs(n);
        const integerPart = Math.floor(num);
        const decimalPart = Math.round((num - integerPart) * 100);

        let result = this._convertIntegerPart(integerPart, digits, units, bigUnits);
        result += "元";

        if (decimalPart > 0) {
            if (decimalPart >= 10) {
                const jiao = Math.floor(decimalPart / 10);
                const fen = decimalPart % 10;
                result += digits[jiao] + "角";
                if (fen > 0) {
                    result += digits[fen] + "分";
                }
            } else {
                result += digits[decimalPart] + "分";
            }
        } else {
            result += "整";
        }

        if (n < 0) result += "（负）";

        return result;
    }

    /**
     * 转换整数部分为人民币大写
     * @private
     */
    _convertIntegerPart(num, digits, units, bigUnits) {
        if (num === 0) return "";

        let result = "";
        let bigUnitIndex = 0;

        while (num > 0) {
            let section = num % 10000;
            if (section > 0) {
                let sectionResult = this._convertSection(section, digits, units);
                result = sectionResult + bigUnits[bigUnitIndex] + result;
            }
            num = Math.floor(num / 10000);
            bigUnitIndex++;
        }

        return result;
    }

    /**
     * 转换每4位一组
     * @private
     */
    _convertSection(num, digits, units) {
        let result = "";
        let unitIndex = 0;
        let lastZero = false;

        while (num > 0) {
            const digit = num % 10;
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

    /**
     * 将数字转换为人民币大写（英文别名）
     * @param {Number} n - 要转换的数字
     * @returns {String} 转换后的人民币大写
     */
    rmbdx(n) {
        return this.z人民币大写(n);
    }

    /**
     * 生成一个指定范围内的随机整数
     * @param {Number} start - 随机整数的起始值
     * @param {Number} end - 随机整数的结束值
     * @returns {Number} 生成的随机整数
     * @example
     * logjson(JSA.z随机整数(5, 10)); // 6到10之间的随机整数
     */
    z随机整数(start, end) {
        return Math.floor(Math.random() * (end - start + 1)) + start;
    }

    /**
     * 生成随机整数（英文别名）
     * @param {Number} start - 起始值
     * @param {Number} end - 结束值
     * @returns {Number} 生成的随机整数
     */
    rndInt(start, end) {
        return this.z随机整数(start, end);
    }

    /**
     * 生成一个指定范围内的随机整数一维数组
     * @param {Number} start - 随机整数的起始值
     * @param {Number} end - 随机整数的结束值
     * @param {Number} count - 随机整数的个数
     * @returns {Array} 返回生成的随机整数一维数组
     * @example
     * logjson(JSA.z随机整数数组(5, 20, 5)); // [18, 15, 6, 7, 12] 等随机数组
     */
    z随机整数数组(start, end, count) {
        const result = [];
        for (let i = 0; i < count; i++) {
            result.push(this.z随机整数(start, end));
        }
        return result;
    }

    /**
     * 生成随机整数数组（英文别名）
     * @param {Number} start - 起始值
     * @param {Number} end - 结束值
     * @param {Number} count - 个数
     * @returns {Array} 生成的随机整数一维数组
     */
    rndIntArray(start, end, count) {
        return this.z随机整数数组(start, end, count);
    }

    /**
     * 生成一个指定范围内的随机小数
     * @param {Number} min - 随机小数的最小值
     * @param {Number} max - 随机小数的最大值
     * @param {Number} decimalPlaces - 小数的位数（默认10）
     * @returns {Number} 生成的随机小数
     * @example
     * logjson(JSA.z随机小数(5, 20, 5)); // 6.46169 等随机小数
     */
    z随机小数(min, max, decimalPlaces) {
        if (decimalPlaces === undefined) decimalPlaces = 10;

        const multiplier = Math.pow(10, decimalPlaces);
        const randomNum = Math.random() * (max - min) + min;
        return Math.round(randomNum * multiplier) / multiplier;
    }

    /**
     * 生成随机小数（英文别名）
     * @param {Number} min - 最小值
     * @param {Number} max - 最大值
     * @param {Number} decimalPlaces - 小数位数
     * @returns {Number} 生成的随机小数
     */
    rndFloat(min, max, decimalPlaces) {
        return this.z随机小数(min, max, decimalPlaces);
    }

    /**
     * 生成一个指定范围内的随机小数一维数组
     * @param {Number} start - 随机小数的最小值
     * @param {Number} end - 随机小数的最大值
     * @param {Number} count - 随机小数的个数
     * @param {Number} decimalPlaces - 小数的位数（默认10）
     * @returns {Array} 生成的随机小数一维数组
     * @example
     * logjson(JSA.z随机小数数组(5, 20, 5, 2)); // [18.67, 16.77, 14.73, 10.64, 10.7] 等随机数组
     */
    z随机小数数组(start, end, count, decimalPlaces) {
        if (decimalPlaces === undefined) decimalPlaces = 10;

        const result = [];
        for (let i = 0; i < count; i++) {
            result.push(this.z随机小数(start, end, decimalPlaces));
        }
        return result;
    }

    /**
     * 生成随机小数数组（英文别名）
     * @param {Number} start - 起始值
     * @param {Number} end - 结束值
     * @param {Number} count - 个数
     * @param {Number} decimalPlaces - 小数位数
     * @returns {Array} 生成的随机小数一维数组
     */
    rndFloatArray(start, end, count, decimalPlaces) {
        return this.z随机小数数组(start, end, count, decimalPlaces);
    }

    /**
     * 随机打乱数组的顺序
     * @param {Array} array - 要打乱顺序的数组
     * @returns {Array} 打乱顺序后的数组
     * @example
     * logjson(JSA.z随机打乱([11, 20, 12, 16, 10])); // [10, 11, 12, 16, 20] 等随机排列
     */
    z随机打乱(array) {
        const result = [...array];
        for (let i = result.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            const temp = result[i];
            result[i] = result[j];
            result[j] = temp;
        }
        return result;
    }

    /**
     * 随机打乱数组（英文别名）
     * @param {Array} array - 要打乱顺序的数组
     * @returns {Array} 打乱顺序后的数组
     */
    shuffle(array) {
        return this.z随机打乱(array);
    }

    /**
     * 生成一个指定范围内的随机乱序数字序列一维数组
     * @param {Number} start - 数字序列的起始值
     * @param {Number} end - 数字序列的结束值
     * @param {Number} step - 数字序列的步长（默认1）
     * @returns {Array} 生成的随机乱序数字序列一维数组
     * @example
     * logjson(JSA.z随机乱序数字序列(5, 15, 2)); // [9, 11, 15, 5, 7, 13] 等随机序列
     */
    z随机乱序数字序列(start, end, step) {
        if (step === undefined) step = 1;

        const sequence = this.z生成数字序列(start, end, step);
        return this.z随机打乱(sequence);
    }

    /**
     * 生成随机乱序数字序列（英文别名）
     * @param {Number} start - 起始值
     * @param {Number} end - 结束值
     * @param {Number} step - 步长
     * @returns {Array} 生成的随机乱序数字序列一维数组
     */
    shuffleNumbers(start, end, step) {
        return this.z随机乱序数字序列(start, end, step);
    }

    /**
     * 延时执行函数
     * @param {Number} ts - 延时时间（毫秒）
     * @example
     * console.timeStart();
     * JSA.z延时(5);
     * console.timelog('测试延时'); // 测试延时用时：0.005秒
     */
    z延时(ts) {
        // WPS JSA环境中使用适当的方式延时
        const start = Date.now();
        while (Date.now() - start < ts) {
            // 等待
        }
    }

    /**
     * 延时执行（英文别名）
     * @param {Number} ts - 延时时间（毫秒）
     */
    delay(ts) {
        this.z延时(ts);
    }

    /**
     * 统一路径分隔符
     * @param {String} path - 要处理的路径
     * @returns {String} 处理后的路径
     * @example
     * Console.log(JSA.z统一路径分隔符("c:/Documents/123")); // c:\Documents\123
     */
    z统一路径分隔符(path) {
        return path.replace(/\//g, '\\').replace(/\\/g, '\\');
    }

    /**
     * 统一路径分隔符（英文别名）
     * @param {String} path - 要处理的路径
     * @returns {String} 处理后的路径
     */
    normalPath(path) {
        return this.z统一路径分隔符(path);
    }

    /**
     * 返回值格式化后的字符串
     * @param {number} v - 要格式化的值
     * @param {String} fmt - 格式
     * @returns {String} 格式化后的字符串
     * @example
     * Console.log(JSA.z格式化(6.899122, "0.00")); // 6.90
     */
    z格式化(v, fmt) {
        // 简单的格式化实现
        if (fmt.indexOf(".") !== -1) {
            const decimalPlaces = fmt.split(".")[1].length;
            return (Math.round(v * Math.pow(10, decimalPlaces)) / Math.pow(10, decimalPlaces)).toFixed(decimalPlaces);
        }
        return String(Math.round(v));
    }

    /**
     * 格式化值（英文别名）
     * @param {number} v - 要格式化的值
     * @param {String} fmt - 格式
     * @returns {String} 格式化后的字符串
     */
    format(v, fmt) {
        return this.z格式化(v, fmt);
    }

    /**
     * 计算两个日期之间的间隔
     * @param {Date|string} d1 - 较小的日期
     * @param {Date|string} d2 - 较大的日期
     * @param {String} format - 格式：Y年 M月 D天 YD忽略年差 MD忽略年和月 YM忽略年仅计算月差
     * @returns {String|Number} 按格式返回间隔
     * @example
     * Console.log(JSA.z日期间隔('2023-5-3', '2024-10-7', 'YD')); // 157
     * Console.log(JSA.z日期间隔('2023-5-3', '2024-10-7')); // 1年5个月4天
     */
    z日期间隔(d1, d2, format) {
        const date1 = typeof d1 === 'string' ? new Date(d1) : d1;
        const date2 = typeof d2 === 'string' ? new Date(d2) : d2;

        if (format === 'Y') {
            // 计算年差
            return date2.getFullYear() - date1.getFullYear();
        } else if (format === 'M') {
            // 计算月差
            const years = date2.getFullYear() - date1.getFullYear();
            const months = date2.getMonth() - date1.getMonth();
            return years * 12 + months;
        } else if (format === 'D') {
            // 计算天数差
            const msPerDay = 24 * 60 * 60 * 1000;
            return Math.round((date2 - date1) / msPerDay);
        } else if (format === 'YD') {
            // 忽略年差，仅计算天数差
            const tempDate = new Date(date1);
            tempDate.setFullYear(date2.getFullYear());
            const msPerDay = 24 * 60 * 60 * 1000;
            return Math.round((date2 - tempDate) / msPerDay);
        } else if (format === 'MD') {
            // 忽略年和月份差，仅计算天数差
            const tempDate = new Date(date1);
            tempDate.setFullYear(date2.getFullYear());
            tempDate.setMonth(date2.getMonth());
            const msPerDay = 24 * 60 * 60 * 1000;
            return Math.round((date2 - tempDate) / msPerDay);
        } else if (format === 'YM') {
            // 忽略年差，仅计算月差
            const months = date2.getMonth() - date1.getMonth();
            return months < 0 ? months + 12 : months;
        } else {
            // 默认：都计算
            const years = date2.getFullYear() - date1.getFullYear();
            let months = date2.getMonth() - date1.getMonth();
            let days = date2.getDate() - date1.getDate();

            if (days < 0) {
                months--;
                const prevMonth = new Date(date2.getFullYear(), date2.getMonth(), 0);
                days += prevMonth.getDate();
            }

            if (months < 0) {
                months += 12;
            }

            let result = "";
            if (years > 0) result += years + "年";
            if (months > 0) result += months + "个月";
            if (days > 0) result += days + "天";
            return result || "0天";
        }
    }

    /**
     * 计算日期间隔（英文别名）
     * @param {Date|string} d1 - 较小的日期
     * @param {Date|string} d2 - 较大的日期
     * @param {String} format - 格式
     * @returns {String|Number} 按格式返回间隔
     */
    datedif(d1, d2, format) {
        return this.z日期间隔(d1, d2, format);
    }

    /**
     * 对字符串表达式求结果
     * @param {String} expression - 字符串表达式
     * @returns {Number} 结果
     * @example
     * console.log(JSA.z表达式求值('5*6+5')); // 35
     */
    z表达式求值(expression) {
        try {
            // 注意：在WPS JSA环境中使用eval需要谨慎
            return eval(expression);
        } catch (e) {
            throw new Error("表达式求值错误: " + e.message);
        }
    }

    /**
     * 对字符串表达式求值（英文别名）
     * @param {String} expression - 字符串表达式
     * @returns {Number} 结果
     */
    eval880(expression) {
        return this.z表达式求值(expression);
    }

    /**
     * 选择二维数组中指定的列
     * @param {Array} arr - 二维数组
     * @param {Array} colIndexes - 要选择列的索引数组（可以是列号或表头名称）
     * @param {Array} newHeaders - 要指定的表头（可选）
     * @returns {Array} 返回选择后的结果二维数组
     * @example
     * var arr = [['a','b','c'], [1,2,3], [4,5,6]];
     * logjson(JSA.z选择列(arr, ['c','b','a'])); // 按表头选择并重新排序
     * logjson(JSA.z选择列(arr, ['x','z'], ['x','y','z'])); // 使用自定义表头
     */
    z选择列(arr, colIndexes, newHeaders) {
        if (!arr || arr.length === 0) return [];

        // 确定列索引
        let indexes = [];

        // 检查是否按表头选择
        if (arr.length > 0 && colIndexes.length > 0 && typeof colIndexes[0] === 'string') {
            const headers = arr[0];
            const headerMap = {};
            for (let i = 0; i < headers.length; i++) {
                headerMap[String(headers[i])] = i;
            }

            if (newHeaders && newHeaders.length > 0) {
                // 使用自定义表头
                for (let j = 0; j < colIndexes.length; j++) {
                    const col = colIndexes[j];
                    if (headerMap.hasOwnProperty(col)) {
                        indexes.push(headerMap[col]);
                    }
                }
            } else {
                // 使用原表头
                for (let j = 0; j < colIndexes.length; j++) {
                    const col = colIndexes[j];
                    if (headerMap.hasOwnProperty(col)) {
                        indexes.push(headerMap[col]);
                    }
                }
            }

            // 如果是按表头选择，第一行也按新顺序排列
            const result = [];
            if (newHeaders && newHeaders.length > 0) {
                // 使用自定义表头
                result.push(newHeaders);
            } else {
                // 使用原表头但按新顺序
                const newRow = [];
                for (let j = 0; j < colIndexes.length; j++) {
                    const col = colIndexes[j];
                    const idx = headerMap[col];
                    if (idx !== undefined) {
                        newRow.push(headers[idx]);
                    } else {
                        newRow.push(col);
                    }
                }
                result.push(newRow);
            }

            // 选择数据行
            for (let i = 1; i < arr.length; i++) {
                const row = arr[i];
                const newRow = [];
                for (let k = 0; k < indexes.length; k++) {
                    newRow.push(row[indexes[k]]);
                }
                result.push(newRow);
            }

            return result;
        } else {
            // 按列号选择（0-based）
            indexes = [];
            for (let j = 0; j < colIndexes.length; j++) {
                const idx = typeof colIndexes[j] === 'number' ? colIndexes[j] : parseInt(colIndexes[j]);
                indexes.push(idx);
            }

            const result = [];
            for (let i = 0; i < arr.length; i++) {
                const row = arr[i];
                const newRow = [];
                for (let k = 0; k < indexes.length; k++) {
                    newRow.push(row[indexes[k]]);
                }
                result.push(newRow);
            }

            return result;
        }
    }

    /**
     * 选择二维数组中指定的列（英文别名）
     * @param {Array} arr - 二维数组
     * @param {Array} colIndexes - 要选择列的索引数组
     * @param {Array} newHeaders - 要指定的表头
     * @returns {Array} 返回选择后的结果二维数组
     */
    selectCols(arr, colIndexes, newHeaders) {
        return this.z选择列(arr, colIndexes, newHeaders);
    }

    /**
     * 以字符串形式创建一个函数并依次传递参数
     * @param {String} fn - 字符串形式的函数定义
     * @param {...any} args - 函数对应的参数列表
     * @returns {*} 自定义函数的计算结果
     * @example
     * // 在WPS公式中使用
     * =jsaLambda("Math.max", A1, A2, A3)
     */
    jsaLambda(fn, ...args) {
        try {
            const func = eval(fn);
            if (typeof func === 'function') {
                return func(...args);
            }
            return func;
        } catch (e) {
            throw new Error("Lambda执行错误: " + e.message);
        }
    }

    /**
     * 将总行数按规定列数和方向排版
     * @param {Number} totalRows - 总行数
     * @param {Number} cols - 规定列数
     * @param {String} direction - 方向：'r'先行后列，'c'先列后行
     * @returns {Array} 排版后的数字序列（从0开始到总行数-1）的二维数组
     * @example
     * var rs = JSA.z矩阵分布(7, 4, 'r'); // 7个数，4列，先行后列
     * logjson(rs); // [[0,1,2,3], [4,5,6]]
     */
    z矩阵分布(totalRows, cols, direction) {
        if (direction === undefined) direction = 'r';

        const result = [];
        const numbers = [];
        for (let i = 0; i < totalRows; i++) {
            numbers.push(i);
        }

        if (direction === 'r') {
            // 先行后列
            const rows = Math.ceil(totalRows / cols);
            for (let i = 0; i < rows; i++) {
                const row = [];
                for (let j = 0; j < cols; j++) {
                    const index = i * cols + j;
                    if (index < totalRows) {
                        row.push(numbers[index]);
                    }
                }
                if (row.length > 0) {
                    result.push(row);
                }
            }
        } else {
            // 先列后行
            const rows = Math.ceil(totalRows / cols);
            for (let i = 0; i < rows; i++) {
                const row = [];
                for (let j = 0; j < cols; j++) {
                    const index = j * rows + i;
                    if (index < totalRows) {
                        row.push(numbers[index]);
                    }
                }
                if (row.length > 0) {
                    result.push(row);
                }
            }
        }

        return result;
    }

    /**
     * 矩阵分布（英文别名）
     * @param {Number} totalRows - 总行数
     * @param {Number} cols - 规定列数
     * @param {String} direction - 方向
     * @returns {Array} 排版后的数字序列二维数组
     */
    getMatrix(totalRows, cols, direction) {
        return this.z矩阵分布(totalRows, cols, direction);
    }
}

/**
 * 创建JSA全局实例 - 可直接使用 JSA.方法名() 调用
 * @example
 * JSA.z转置([[1,2,3],[4,5,6]]);
 */
const JSA = new clsJSA();

/**
 * ShtUtils - 工作表函数工具库
 * 根据 https://vbayyds.com/api/jsa880/ShtUtils.html 文档编写
 * 作者: 郑广学 (EXCEL880)
 * 版本: 1.0.0
 *
 * @description 增强工作表操作，支持表名和工作表对象作为参数
 * @example
 * // 基本使用
 * var usedRange = ShtUtils.z安全已使用区域("多表");
 * console.log(usedRange.Address()); // $A$1:$L$17
 */
class clsShtUtils {
    /**
     * 构造函数 - 创建ShtUtils实例
     * @constructor
     * @param {Worksheet|string} initialSheet - 初始工作表（表名或工作表对象，可选）
     * @example
     * // 创建空实例
     * const shtUtils = new clsShtUtils();
     * // 创建带初始工作表的实例（支持链式调用）
     * const shtUtils = new clsShtUtils("Sheet1");
     */
    constructor(initialSheet = null) {
        this.MODULE_NAME = "ShtUtils";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学JSA880框架";
        this.sheet = initialSheet ? this._getSheet(initialSheet) : null;
    }

    /**
     * 获取/设置当前工作表 - 支持链式调用
     * @param {Worksheet|string} newSheet - 新的工作表（表名或工作表对象，可选）
     * @returns {Worksheet|clsShtUtils} 传入参数时返回this，否则返回当前工作表对象
     * @example
     * // 获取当前工作表
     * const currentSheet = shtUtils.sht();
     * // 设置新工作表（支持链式调用）
     * shtUtils.sht("Sheet1").z激活表();
     */
    sht(newSheet) {
        if (newSheet !== undefined) {
            this.sheet = this._getSheet(newSheet);
            return this;
        }
        return this.sheet;
    }

    /**
     * 获取当前工作表的值
     * @returns {Array} 工作表的已使用区域值
     * @example
     * // 获取当前工作表的值
     * const values = shtUtils.sht("Sheet1").val();
     */
    val() {
        if (!this.sheet) return null;
        const usedRange = this.sheet.UsedRange;
        if (!usedRange) return [];
        return usedRange.Value2;
    }

    /**
     * 获取工作表对象（辅助方法）
     * @private
     * @param {String|Worksheet} sht - 表名或工作表对象
     * @param {Sheets} shts - 表集合
     * @returns {Worksheet} 工作表对象
     */
    _getSheet(sht, shts) {
        if (typeof sht === 'string') {
            return shts ? shts(sht) : Sheets(sht);
        }
        return sht;
    }

    /**
     * 获取工作表从A1开始的可使用区域
     * @param {String|Worksheet} 工作表 - 要获取安全已使用区域的工作表（表名或工作表对象）
     * @returns {Range} 工作表从A1开始的可使用单元格区域
     * @example
     * var usedRange = ShtUtils.z安全已使用区域("多表");
     * console.log(usedRange.Address()); // $A$1:$L$17
     * // 也可以: var usedRange = ShtUtils.z安全已使用区域(Sheets("多表"));
     */
    z安全已使用区域(工作表) {
        const sheet = this._getSheet(工作表);
        let usedRange;
        try {
            usedRange = sheet.UsedRange;
        } catch (e) {
            return sheet.Range("A1");
        }

        if (!usedRange) {
            return sheet.Range("A1");
        }

        // 获取从A1开始的已使用区域
        const lastRow = usedRange.Row + usedRange.Rows.Count - 1;
        const lastCol = usedRange.Column + usedRange.Columns.Count - 1;

        return sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, lastCol));
    }

    /**
     * 获取工作表从A1开始的可使用区域（英文别名）
     * @param {String|Worksheet} 工作表 - 要获取安全已使用区域的工作表
     * @returns {Range} 工作表从A1开始的可使用单元格区域
     */
    safeUsedRange(工作表) {
        return this.z安全已使用区域(工作表);
    }

    /**
     * 检查表集合中是否包含指定表名（支持通配符）
     * @param {String} 表名 - 要检查的表名，可以用? *通配符
     * @param {Object} 表集合 - 要检查的表集合对象，默认为Sheets
     * @returns {Boolean} 如果表集合中包含指定表名则返回true，否则返回false
     * @example
     * var includesSheet = ShtUtils.z包含表名('多*');
     * console.log(includesSheet); // true
     */
    z包含表名(表名, 表集合) {
        const shts = 表集合 || Sheets;
        const pattern = this._wildcardToRegex(表名);

        for (let i = 1; i <= shts.Count; i++) {
            if (pattern.test(shts(i).Name)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 检查表集合中是否包含指定表名（英文别名）
     * @param {String} 表名 - 要检查的表名，可以用? *通配符
     * @param {Object} 表集合 - 要检查的表集合对象
     * @returns {Boolean} 如果表集合中包含指定表名则返回true
     */
    includesSht(表名, 表集合) {
        return this.z包含表名(表名, 表集合);
    }

    /**
     * 筛选表集合中包含指定表名的表（支持通配符）
     * @param {String} 表名 - 要筛选的表名，可以用? *通配符
     * @param {Sheets} 表集合 - 要筛选的表集合对象，默认为Sheets
     * @returns {Array} 包含筛选结果的表名一维数组
     * @example
     * var filteredSheets = ShtUtils.z表名筛选('多*');
     * logjson(filteredSheets, 0); // ["多表 (3)", "多表", "多表 (2)", "多条件筛选"]
     */
    z表名筛选(表名, 表集合) {
        const shts = 表集合 || Sheets;
        const pattern = this._wildcardToRegex(表名);
        const result = [];

        for (let i = 1; i <= shts.Count; i++) {
            if (pattern.test(shts(i).Name)) {
                result.push(shts(i).Name);
            }
        }
        return result;
    }

    /**
     * 筛选表集合中包含指定表名的表（英文别名）
     * @param {String} 表名 - 要筛选的表名，可以用? *通配符
     * @param {Sheets} 表集合 - 要筛选的表集合对象
     * @returns {Array} 包含筛选结果的表名一维数组
     */
    filterShts(表名, 表集合) {
        return this.z表名筛选(表名, 表集合);
    }

    /**
     * 在表集合中查找指定的表
     * @param {String} sht - 要查找的表名
     * @param {Object} shts - 要查找的表集合对象，默认为Sheets
     * @returns {Sheet} 查找到的表对象
     * @example
     * var findSht = ShtUtils.z查找表('1月');
     * console.log(findSht.Name); // 1月
     */
    z查找表(sht, shts) {
        const sheets = shts || Sheets;
        return sheets(sht);
    }

    /**
     * 在表集合中查找指定的表（英文别名）
     * @param {String} sht - 要查找的表名
     * @param {Object} shts - 要查找的表集合对象
     * @returns {Sheet} 查找到的表对象
     */
    findSht(sht, shts) {
        return this.z查找表(sht, shts);
    }

    /**
     * 判断工作表是否为空表
     * @param {String|Worksheet} 工作表 - 要判断的工作表（表名或工作表对象）
     * @returns {Boolean} 如果工作表为空表则返回true，否则返回false
     * @example
     * var isEmpty = ShtUtils.z判断空表('Sheet1');
     * console.log(isEmpty); // true
     */
    z判断空表(工作表) {
        const sheet = this._getSheet(工作表);
        try {
            const usedRange = sheet.UsedRange;
            if (!usedRange) return true;

            // 检查是否有任何单元格有值
            for (let row = 1; row <= usedRange.Rows.Count; row++) {
                for (let col = 1; col <= usedRange.Columns.Count; col++) {
                    const cellValue = usedRange.Cells(row, col).Value;
                    if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                        return false;
                    }
                }
            }
            return true;
        } catch (e) {
            return true;
        }
    }

    /**
     * 判断工作表是否为空表（英文别名）
     * @param {String|Worksheet} 工作表 - 要判断的工作表
     * @returns {Boolean} 如果工作表为空表则返回true
     */
    isEmptySht(工作表) {
        return this.z判断空表(工作表);
    }

    /**
     * 删除指定的工作表 - 支持链式调用
     * @param {String|Worksheet} 工作表 - 要删除的工作表（表名或工作表对象，可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     * @example
     * // 传统调用方式
     * ShtUtils.z删除表('Sheet1');
     * // 链式调用方式
     * shtUtils.sht('Sheet1').z删除表();
     */
    z删除表(工作表) {
        const sheet = 工作表 ? this._getSheet(工作表) : this.sheet;
        if (!sheet) return this;
        Application.DisplayAlerts = false;
        try {
            sheet.Delete();
        } catch (e) {
            console.error("删除工作表失败:", e);
        } finally {
            Application.DisplayAlerts = true;
        }
        this.sheet = null;
        return this;
    }

    /**
     * 删除指定的工作表（英文别名）- 支持链式调用
     * @param {String|Worksheet} sht - 要删除的工作表（可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     */
    deleteSht(sht) {
        return this.z删除表(sht);
    }

    /**
     * 按照代码名称查找表集合中的表
     * @param {String} 表名 - 要查找的代码名称
     * @param {Object} 表集合 - 要查找的表集合对象，默认为Sheets
     * @returns {Worksheet} 查找到的表对象
     * @example
     * // 查看工作表的代码名称: console.log(Sheets("1月").CodeName);
     * var sheetRange = ShtUtils.z按代码名称("Sheet34");
     * console.log(sheetRange.Name); // 1月
     */
    z按代码名称(表名, 表集合) {
        const shts = 表集合 || Sheets;
        for (let i = 1; i <= shts.Count; i++) {
            const sheet = shts(i);
            if (sheet.CodeName === 表名) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * 按照代码名称查找表集合中的表（英文别名）
     * @param {String} 表名 - 要查找的代码名称
     * @param {Object} 表集合 - 要查找的表集合对象
     * @returns {Worksheet} 查找到的表对象
     */
    byCodeName(表名, 表集合) {
        return this.z按代码名称(表名, 表集合);
    }

    /**
     * 隐藏表集合中的表 - 支持链式调用
     * @param {Object|Array} 表集合 - 要隐藏的表集合对象或数组（可选，不传则使用this.sheet）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     * @example
     * // 传统调用方式
     * ShtUtils.z隐藏表([Sheets("1月"), Sheets("2月")]);
     * // 链式调用方式 - 隐藏当前实例的工作表
     * shtUtils.sht('Sheet1').z隐藏表();
     */
    z隐藏表(表集合) {
        if (!表集合) {
            // 不传参数则使用当前工作表
            if (this.sheet) {
                this.sheet.Visible = false;
            }
        } else if (Array.isArray(表集合)) {
            for (let i = 0; i < 表集合.length; i++) {
                表集合[i].Visible = false;
            }
        } else {
            表集合.Visible = false;
        }
        return this;
    }

    /**
     * 隐藏表集合中的表（英文别名）- 支持链式调用
     * @param {Object|Array} shts - 要隐藏的表集合对象或数组（可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     */
    hideSheets(shts) {
        return this.z隐藏表(shts);
    }

    /**
     * 显示表集合中的表 - 支持链式调用
     * @param {Array} 表集合 - 表名数组或工作表对象数组（可选，不传则使用this.sheet）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     * @example
     * // 传统调用方式
     * ShtUtils.z显示表([Sheets("1月"), Sheets("2月")]);
     * // 链式调用方式 - 显示当前实例的工作表
     * shtUtils.sht('Sheet1').z显示表();
     */
    z显示表(表集合) {
        if (!表集合) {
            // 不传参数则显示当前工作表
            if (this.sheet) {
                this.sheet.Visible = true;
            }
        } else if (Array.isArray(表集合)) {
            for (let i = 0; i < 表集合.length; i++) {
                const sheet = typeof 表集合[i] === 'string' ? Sheets(表集合[i]) : 表集合[i];
                if (sheet) sheet.Visible = true;
            }
        }
        return this;
    }

    /**
     * 显示表集合中的表（英文别名）- 支持链式调用
     * @param {Array} shts - 表名数组或工作表对象数组（可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     */
    showSheets(shts) {
        return this.z显示表(shts);
    }

    /**
     * 激活指定的工作表 - 支持链式调用
     * @param {String|Worksheet} 工作表 - 待激活的工作表（表名或工作表对象，可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     * @example
     * // 传统调用方式
     * ShtUtils.z激活表(Sheets("1月"));
     * // 链式调用方式
     * shtUtils.sht('Sheet1').z激活表();
     */
    z激活表(工作表) {
        const sheet = 工作表 ? this._getSheet(工作表) : this.sheet;
        if (!sheet) return this;
        sheet.Activate();
        this.sheet = sheet;
        return this;
    }

    /**
     * 激活指定的工作表（英文别名）- 支持链式调用
     * @param {String|Worksheet} 工作表 - 待激活的工作表（可选）
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     */
    shtActivate(工作表) {
        return this.z激活表(工作表);
    }

    /**
     * 返回指定工作表的最后一行的行号
     * @param {String|Worksheet} 工作表 - 要返回最后行号的工作表（表名或工作表对象）
     * @returns {Number} 最后一行的行号
     * @example
     * var n = ShtUtils.z最后一行(Sheets("1月"));
     * console.log(n); // 31
     */
    z最后一行(工作表) {
        const sheet = this._getSheet(工作表);
        try {
            const usedRange = sheet.UsedRange;
            if (!usedRange) return 0;

            return usedRange.Row + usedRange.Rows.Count - 1;
        } catch (e) {
            return 0;
        }
    }

    /**
     * 返回指定工作表的最后一行的行号（英文别名）
     * @param {String|Worksheet} 工作表 - 要返回最后行号的工作表
     * @returns {Number} 最后一行的行号
     */
    lastRow(工作表) {
        return this.z最后一行(工作表);
    }

    /**
     * 将工作表名字中包含的违规字符替换为下划线
     * 违规字符: 超过31个, 字符: :\ / ? * [ 或 ] 第一或最后一个字符用单引号
     * @param {String} 工作表名 - 待检测的工作表名
     * @returns {String} 正确的表名
     * @example
     * ShtUtils.z纠正表名("1[]2") // "1__2"
     */
    z纠正表名(工作表名) {
        let name = String(工作表名);

        // 替换违规字符
        name = name.replace(/[\\\/\?\*\[\]:]/g, '_');

        // 去除首尾单引号
        name = name.replace(/^'+|'+$/g, '');

        // 限制长度为31个字符
        if (name.length > 31) {
            name = name.substring(0, 31);
        }

        return name;
    }

    /**
     * 将工作表名字中包含的违规字符替换为下划线（英文别名）
     * @param {String} 工作表名 - 待检测的工作表名
     * @returns {String} 正确的表名
     */
    correctShtName(工作表名) {
        return this.z纠正表名(工作表名);
    }

    /**
     * 对指定的工作表数组进行排序 - 支持链式调用
     * @param {Object} shts - 要排序的工作表数组或Sheets集合（可选）
     * @param {Function|Array} sortFn - 用于排序的函数或自定义序列数组
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     * @example
     * // 传统调用方式
     * ShtUtils.z工作表排序(Sheets(["产品11","产品1","产品2"]), (a,b)=>val($.mid(a,3))-val($.mid(b,3)));
     * ShtUtils.z工作表排序(Sheets(['工程','销售','财务']), ['销售','财务','工程']); // 自定义序列
     * // 链式调用方式（需要传入工作表集合）
     * shtUtils.z工作表排序(Sheets(['工程','销售','财务']), ['销售','财务','工程']);
     */
    z工作表排序(shts, sortFn) {
        const sheets = shts || Sheets;
        // 获取工作表名称数组
        let sheetNames = [];
        if (Array.isArray(sheets)) {
            sheetNames = sheets.map(s => typeof s === 'string' ? s : s.Name);
        } else {
            // 假设是Sheets集合
            for (let i = 1; i <= sheets.Count; i++) {
                sheetNames.push(sheets(i).Name);
            }
        }

        // 排序
        if (sortFn) {
            if (Array.isArray(sortFn)) {
                // 自定义序列排序
                sheetNames.sort((a, b) => {
                    const indexA = sortFn.indexOf(a);
                    const indexB = sortFn.indexOf(b);
                    return (indexA === -1 ? 999 : indexA) - (indexB === -1 ? 999 : indexB);
                });
            } else if (typeof sortFn === 'function') {
                sheetNames.sort(sortFn);
            }
        } else {
            sheetNames.sort();
        }

        // 重新排列工作表位置
        for (let i = 0; i < sheetNames.length; i++) {
            Sheets(sheetNames[i]).Move(Sheets(i + 1));
        }
        return this;
    }

    /**
     * 对指定的工作表数组进行排序（英文别名）- 支持链式调用
     * @param {Object} shts - 要排序的工作表数组（可选）
     * @param {Function|Array} sortFn - 用于排序的函数或自定义序列
     * @returns {clsShtUtils} 返回当前实例以支持链式调用
     */
    sheetsSort(shts, sortFn) {
        return this.z工作表排序(shts, sortFn);
    }

    /**
     * 将通配符转换为正则表达式（辅助方法）
     * @private
     * @param {String} wildcard - 包含通配符的字符串
     * @returns {RegExp} 正则表达式对象
     */
    _wildcardToRegex(wildcard) {
        const pattern = wildcard.replace(/[.+^${}()|[\]\\]/g, '\\$&') // 转义正则字符
            .replace(/\*/g, '.*') // * 匹配任意字符
            .replace(/\?/g, '.'); // ? 匹配单个字符
        return new RegExp('^' + pattern + '$', 'i');
    }
}

/**
 * 创建ShtUtils全局实例 - 可直接使用 ShtUtils.方法名() 调用
 * @example
 * ShtUtils.z安全已使用区域("多表");
 */
const ShtUtils = new clsShtUtils();

/**
 * DateUtils - 日期常用函数库
 * 根据 https://vbayyds.com/api/jsa880/DateUtils.html 文档编写
 * 作者: 郑广学 (EXCEL880)
 * 版本: 1.0.0
 *
 * @description Excel日期时间与js日期时间的互转、日期的加减操作、日期格式化等
 * @example
 * // 基本使用
 * var jsdate = new Date('2023-9-10 20:05');
 * var excelDate = DateUtils.z转表格日期(jsdate);
 * console.log(excelDate); // 45179.836805555598
 */
class clsDateUtils {
    /**
     * 构造函数 - 创建DateUtils实例
     * @constructor
     * @param {Date|number|string} initialDate - 初始日期（可选）
     */
    constructor(initialDate = null) {
        this.MODULE_NAME = "DateUtils";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学JSA880框架";
        // 如果传入initialDate，存储为Date对象
        this.date = initialDate ? new Date(initialDate) : new Date();
    }

    /**
     * 获取/设置当前日期 - 支持链式调用
     * @param {Date|number|string} newDate - 新的日期（可选）
     * @returns {Date|clsDateUtils} 传入参数时返回this，否则返回当前Date对象
     * @example
     * // 获取当前日期
     * var currentDate = DateUtils.dt();
     * // 设置新日期（支持链式调用）
     * DateUtils.dt(new Date('2023-9-10')).z加天(5).val();
     */
    dt(newDate) {
        if (newDate !== undefined) {
            this.date = new Date(newDate);
            return this;
        }
        return this.date;
    }

    /**
     * 获取当前日期的值
     * @returns {Date} 当前Date对象
     * @example
     * var dateValue = DateUtils.dt(new Date('2023-9-10')).z加天(5).val();
     */
    val() {
        return this.date;
    }

    /**
     * Excel日期基准（1900年1月1日）
     * @private
     */
    get EXCEL_BASE_DATE() {
        return new Date(1900, 0, 1);
    }

    /**
     * 一天的毫秒数
     * @private
     */
    get DAY_IN_MS() {
        return 24 * 60 * 60 * 1000;
    }

    /**
     * 将JavaScript日期对象转换为Excel表格日期格式
     * @param {Date} jsdate - 要转换的JavaScript日期对象
     * @returns {Number} 转换后的Excel表格日期格式
     * @example
     * var jsdate = new Date('2023-9-10 20:05');
     * var excelDate = DateUtils.z转表格日期(jsdate);
     * console.log(excelDate); // 45179.836805555598
     */
    z转表格日期(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }

        const excelBase = this.EXCEL_BASE_DATE.getTime();
        const dateMs = jsdate.getTime();
        const daysDiff = (dateMs - excelBase) / this.DAY_IN_MS;

        // Excel有1900年2月29日的bug，需要加1天
        return daysDiff + 2;
    }

    /**
     * 将JavaScript日期对象转换为Excel日期格式（英文别名）
     * @param {Date} jsdate - 要转换的JavaScript日期对象
     * @returns {Number} 转换后的Excel日期格式
     */
    toExcelDate(jsdate) {
        return this.z转表格日期(jsdate);
    }

    /**
     * 格式化JavaScript日期对象为指定格式的字符串
     * @param {Date} jsdate - 要格式化的JavaScript日期对象
     * @param {String} fmt - 指定的日期格式字符串 y-年 M-月,d-日 时-H 分-m 秒-s 毫秒-SSS 星期-aaa
     * @returns {String} 格式化后的日期字符串
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var formattedDate = DateUtils.z日期格式化(jsdate, 'yyyy-MM-dd HH:mm:ss');
     * console.log(formattedDate); // 2023-09-10 20:05:11
     */
    z日期格式化(jsdate, fmt) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }

        const weekDays = ['日', '一', '二', '三', '四', '五', '六'];

        const result = fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, (match, year, month, day, hour, minute, second, millisecond, week) => {
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

        return result;
    }

    /**
     * 格式化JavaScript日期对象为指定格式的字符串（英文别名）
     * @param {Date} jsdate - 要格式化的JavaScript日期对象
     * @param {String} fmt - 指定的日期格式字符串
     * @returns {String} 格式化后的日期字符串
     */
    format(jsdate, fmt) {
        return this.z日期格式化(jsdate, fmt);
    }

    /**
     * 在当前日期上添加指定天数 - 支持链式调用
     * @param {Number} days - 要添加的天数
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z加天(5).val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd HH:mm:ss')); // 2023-09-15 20:05:11
     */
    z加天(days) {
        const result = new Date(this.date);
        result.setDate(result.getDate() + days);
        this.date = result;
        return this;
    }

    /**
     * 在当前日期上添加指定天数（英文别名）- 支持链式调用
     * @param {Number} days - 要添加的天数
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    addDays(days) {
        return this.z加天(days);
    }

    /**
     * 在当前日期上添加指定时间 - 支持链式调用
     * @param {String} t - 要添加的时间 按字符串格式 '1:10:15' (时:分:秒)
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z加时间("1:00:00").val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd HH:mm:ss')); // 2023-09-10 21:05:11
     */
    z加时间(t) {
        const parts = t.split(':');
        const hours = parseInt(parts[0]) || 0;
        const minutes = parseInt(parts[1]) || 0;
        const seconds = parseInt(parts[2]) || 0;

        const result = new Date(this.date);
        result.setHours(result.getHours() + hours);
        result.setMinutes(result.getMinutes() + minutes);
        result.setSeconds(result.getSeconds() + seconds);

        this.date = result;
        return this;
    }

    /**
     * 在当前日期上添加指定时间（英文别名）- 支持链式调用
     * @param {String} t - 要添加的时间 按字符串格式 '1:10:15'
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    addTimes(t) {
        return this.z加时间(t);
    }

    /**
     * 在当前日期上添加指定月份 - 支持链式调用
     * @param {Number} m - 要添加的月份
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z加月(4).val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd HH:mm:ss')); // 2024-01-10 20:05:11
     */
    z加月(m) {
        const result = new Date(this.date);
        result.setMonth(result.getMonth() + m);
        this.date = result;
        return this;
    }

    /**
     * 在当前日期上添加指定月份（英文别名）- 支持链式调用
     * @param {Number} m - 要添加的月份
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    addMonths(m) {
        return this.z加月(m);
    }

    /**
     * 在当前日期上添加指定年份 - 支持链式调用
     * @param {Number} y - 要添加的年份
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z加年(2).val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd HH:mm:ss')); // 2025-09-10 20:05:11
     */
    z加年(y) {
        const result = new Date(this.date);
        result.setFullYear(result.getFullYear() + y);
        this.date = result;
        return this;
    }

    /**
     * 在当前日期上添加指定年份（英文别名）- 支持链式调用
     * @param {Number} y - 要添加的年份
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    addYears(y) {
        return this.z加年(y);
    }

    /**
     * 获取指定日期的年份
     * @param {Date} jsdate - 要获取年份的JavaScript日期对象
     * @returns {Number} 指定日期的年份
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var year = DateUtils.z年(jsdate);
     * console.log(year); // 2023
     */
    z年(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        return jsdate.getFullYear();
    }

    /**
     * 获取指定日期的年份（英文别名）
     * @param {Date} jsdate - 要获取年份的JavaScript日期对象
     * @returns {Number} 指定日期的年份
     */
    year(jsdate) {
        return this.z年(jsdate);
    }

    /**
     * 获取指定日期的月份
     * @param {Date} jsdate - 要获取月份的JavaScript日期对象
     * @returns {Number} 指定日期的月份
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var month = DateUtils.z月(jsdate);
     * console.log(month); // 9
     */
    z月(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        return jsdate.getMonth() + 1;
    }

    /**
     * 获取指定日期的月份（英文别名）
     * @param {Date} jsdate - 要获取月份的JavaScript日期对象
     * @returns {Number} 指定日期的月份
     */
    month(jsdate) {
        return this.z月(jsdate);
    }

    /**
     * 获取指定日期的日
     * @param {Date} jsdate - 要获取日的JavaScript日期对象
     * @returns {Number} 指定日期的日
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var day = DateUtils.z日(jsdate);
     * console.log(day); // 10
     */
    z日(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        return jsdate.getDate();
    }

    /**
     * 获取指定日期的日（英文别名）
     * @param {Date} jsdate - 要获取日的JavaScript日期对象
     * @returns {Number} 指定日期的日
     */
    day(jsdate) {
        return this.z日(jsdate);
    }

    /**
     * 获取指定日期的星期几（以数字表示）
     * @param {Date} jsdate - 要获取星期的JavaScript日期对象
     * @returns {Number} 星期几（1表示星期一，7表示星期日）
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11'); // 2023年9月10日是星期日
     * var weekday = DateUtils.z星期(jsdate);
     * console.log(weekday); // 7
     */
    z星期(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        const day = jsdate.getDay();
        return day === 0 ? 7 : day; // 0(星期日)转为7
    }

    /**
     * 获取指定日期的星期几（以数字表示，英文别名）
     * @param {Date} jsdate - 要获取星期的JavaScript日期对象
     * @returns {Number} 星期几（1表示星期一，7表示星期日）
     */
    weekday(jsdate) {
        return this.z星期(jsdate);
    }

    /**
     * 获取指定日期的星期几（以中文字符串表示）
     * @param {Date} jsdate - 要获取星期的JavaScript日期对象
     * @returns {String} 星期几的中文字符串（"一"、"二"..."日"）
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11'); // 2023年9月10日是星期日
     * var weekdayCn = DateUtils.z星期中文(jsdate);
     * console.log(weekdayCn); // 日
     */
    z星期中文(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        const weekDays = ['日', '一', '二', '三', '四', '五', '六'];
        return weekDays[jsdate.getDay()];
    }

    /**
     * 获取指定日期的星期几（以中文字符串表示，英文别名）
     * @param {Date} jsdate - 要获取星期的JavaScript日期对象
     * @returns {String} 星期几的中文字符串
     */
    weekdayCn(jsdate) {
        return this.z星期中文(jsdate);
    }

    /**
     * 获取指定日期的季度
     * @param {Date} jsdate - 要获取季度的JavaScript日期对象
     * @returns {Number} 季度（1表示第一季度，2表示第二季度，3表示第三季度，4表示第四季度）
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var quarter = DateUtils.z季度(jsdate);
     * console.log(quarter); // 3
     */
    z季度(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        return Math.floor((jsdate.getMonth() + 2) / 3);
    }

    /**
     * 获取指定日期的季度（英文别名）
     * @param {Date} jsdate - 要获取季度的JavaScript日期对象
     * @returns {Number} 季度
     */
    quarter(jsdate) {
        return this.z季度(jsdate);
    }

    /**
     * 获取指定日期所在月份的天数
     * @param {Date} jsdate - 要获取天数的JavaScript日期对象
     * @returns {Number} 指定日期所在月份的天数
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var days = DateUtils.z当月天数(jsdate);
     * console.log(days); // 30
     */
    z当月天数(jsdate) {
        if (!(jsdate instanceof Date)) {
            jsdate = new Date(jsdate);
        }
        const year = jsdate.getFullYear();
        const month = jsdate.getMonth() + 1;
        return new Date(year, month, 0).getDate();
    }

    /**
     * 获取指定日期所在月份的天数（英文别名）
     * @param {Date} jsdate - 要获取天数的JavaScript日期对象
     * @returns {Number} 指定日期所在月份的天数
     */
    daysOfMonth(jsdate) {
        return this.z当月天数(jsdate);
    }

    /**
     * 将当前日期设置为所在月份的第一天 - 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z月初().val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd')); // 2023-09-01
     */
    z月初() {
        this.date = new Date(this.date.getFullYear(), this.date.getMonth(), 1);
        return this;
    }

    /**
     * 将当前日期设置为所在月份的第一天（英文别名）- 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    firstDayOfMonth() {
        return this.z月初();
    }

    /**
     * 将当前日期设置为所在月份的最后一天 - 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z月底().val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd')); // 2023-09-30
     */
    z月底() {
        this.date = new Date(this.date.getFullYear(), this.date.getMonth() + 1, 0);
        return this;
    }

    /**
     * 将当前日期设置为所在月份的最后一天（英文别名）- 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    endOfMonth() {
        return this.z月底();
    }

    /**
     * 将JavaScript日期对象转换为VBA日期数值
     * @param {Date} jsdate - 要转换的JavaScript日期对象
     * @returns {Number} 转换后的VBA日期数值
     * @example
     * var jsdate = new Date('2023-9-10 20:05:11');
     * var vbaDate = DateUtils.z转VBA日期数值(jsdate);
     * console.log(vbaDate); // 45179.836932870399
     */
    z转VBA日期数值(jsdate) {
        return this.z转表格日期(jsdate);
    }

    /**
     * 将JavaScript日期对象转换为VBA日期数值（英文别名）
     * @param {Date} jsdate - 要转换的JavaScript日期对象
     * @returns {Number} 转换后的VBA日期数值
     */
    cdate(jsdate) {
        return this.z转VBA日期数值(jsdate);
    }

    /**
     * 将Excel日期数值转换为JavaScript日期对象
     * @param {Number} xlsdate - 要转换的Excel日期数值
     * @returns {Date} 转换后的JavaScript日期对象
     * @example
     * var xlsdate = 45179.836932870399;
     * var jsDate = DateUtils.z表格日期转JS(xlsdate);
     * console.log(DateUtils.format(jsDate, 'yyyy-MM-dd HH:mm:ss')); // 2023-09-10 20:05:11
     */
    z表格日期转JS(xlsdate) {
        const excelBase = this.EXCEL_BASE_DATE.getTime();
        // 减去Excel的1900年2月29日bug的额外天数（2）
        const daysDiff = xlsdate - 2;
        const dateMs = excelBase + daysDiff * this.DAY_IN_MS;
        return new Date(dateMs);
    }

    /**
     * 将Excel日期数值转换为JavaScript日期对象（英文别名）
     * @param {Number} xlsdate - 要转换的Excel日期数值
     * @returns {Date} 转换后的JavaScript日期对象
     */
    fromExcelDate(xlsdate) {
        return this.z表格日期转JS(xlsdate);
    }

    /**
     * 获取当前日期的字符串表示（格式为YYYY-MM-DD）
     * @returns {String} 当前日期的字符串表示
     * @example
     * var today = DateUtils.z今天日期();
     * console.log(today); // 2024-09-28
     */
    z今天日期() {
        const now = new Date();
        return this.format(now, 'yyyy-MM-dd');
    }

    /**
     * 获取当前日期的字符串表示（英文别名）
     * @returns {String} 当前日期的字符串表示
     */
    today() {
        return this.z今天日期();
    }

    /**
     * 获取当前日期和时间的字符串表示（格式为YYYY-MM-DD HH:mm:ss）
     * @returns {String} 当前日期和时间的字符串表示
     * @example
     * var datetime = DateUtils.z日期时间();
     * console.log(datetime); // 2024-09-28 20:48:13
     */
    z日期时间() {
        const now = new Date();
        return this.format(now, 'yyyy-MM-dd HH:mm:ss');
    }

    /**
     * 获取当前日期和时间的字符串表示（英文别名）
     * @returns {String} 当前日期和时间的字符串表示
     */
    now() {
        return this.z日期时间();
    }

    /**
     * 获取当前时间的字符串表示（格式为HH:mm:ss）
     * @returns {String} 当前时间的字符串表示
     * @example
     * var time = DateUtils.z时间();
     * console.log(time); // 20:49:53
     */
    z时间() {
        const now = new Date();
        return this.format(now, 'HH:mm:ss');
    }

    /**
     * 获取当前时间的字符串表示（英文别名）
     * @returns {String} 当前时间的字符串表示
     */
    time() {
        return this.z时间();
    }

    /**
     * 计算两个日期之间的间隔
     * @param {Date} dmin - 较小的日期
     * @param {Date} dmax - 较大的日期
     * @returns {String} 两个日期之间的间隔字符串表示（格式: x年x个月x天）
     * @example
     * var dmin = new Date('2023-9-10 20:05:11');
     * var dmax = new Date('2023-10-6 20:05:11');
     * var interval = DateUtils.z日期间隔(dmin, dmax);
     * console.log(interval); // 0年0个月26天
     */
    z日期间隔(dmin, dmax) {
        if (!(dmin instanceof Date)) dmin = new Date(dmin);
        if (!(dmax instanceof Date)) dmax = new Date(dmax);

        // 确保dmin较小，dmax较大
        if (dmin > dmax) {
            [dmin, dmax] = [dmax, dmin];
        }

        let years = dmax.getFullYear() - dmin.getFullYear();
        let months = dmax.getMonth() - dmin.getMonth();
        let days = dmax.getDate() - dmin.getDate();

        if (days < 0) {
            months--;
            // 获取上个月的天数
            const prevMonth = new Date(dmax.getFullYear(), dmax.getMonth(), 0);
            days += prevMonth.getDate();
        }

        if (months < 0) {
            years--;
            months += 12;
        }

        return years + '年' + months + '个月' + days + '天';
    }

    /**
     * 计算两个日期之间的间隔（英文别名）
     * @param {Date} dmin - 较小的日期
     * @param {Date} dmax - 较大的日期
     * @returns {String} 两个日期之间的间隔字符串表示
     */
    datedif(dmin, dmax) {
        return this.z日期间隔(dmin, dmax);
    }

    /**
     * 将当前日期设置为只保留日期部分，时间部分置为0 - 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z只留日期().val();
     * console.log(DateUtils.format(result, 'yyyy-MM-dd')); // 2023-09-10
     */
    z只留日期() {
        this.date = new Date(this.date.getFullYear(), this.date.getMonth(), this.date.getDate());
        return this;
    }

    /**
     * 将当前日期设置为只保留日期部分（英文别名）- 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    justDate() {
        return this.z只留日期();
    }

    /**
     * 将当前日期设置为只保留时间部分，日期部分置为0 - 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     * @example
     * var result = DateUtils.dt(new Date('2023-9-10 20:05:11')).z只留时间().val();
     * console.log(result.getHours()); // 20
     */
    z只留时间() {
        const hours = this.date.getHours();
        const minutes = this.date.getMinutes();
        const seconds = this.date.getSeconds();
        const milliseconds = this.date.getMilliseconds();

        // 创建一个新的日期对象，只保留时间部分
        this.date = new Date(0, 0, 1, hours, minutes, seconds, milliseconds);
        return this;
    }

    /**
     * 将当前日期设置为只保留时间部分（英文别名）- 支持链式调用
     * @returns {clsDateUtils} 返回当前实例以支持链式调用
     */
    justTime() {
        return this.z只留时间();
    }
}

/**
 * IO - IO文件系统函数库
 * 根据 https://vbayyds.com/api/jsa880/IO.html 文档编写
 * 作者: 郑广学 (EXCEL880)
 * 版本: 1.0.0
 *
 * @description 实现多层文件遍历、文件夹遍历、封装常见文件操作函数
 * @example
 * // 基本使用
 * var isFile = IO.z是否文件('/path/to/file.txt');
 * console.log(isFile); // true
 *
 * var files = IO.z遍历文件('/path/to/folder', true, true, false);
 * console.log(files); // ['/path/to/folder/file1.txt', ...]
 */
class clsIO {
    /**
     * 构造函数 - 创建IO实例
     * @constructor
     */
    constructor() {
        this.MODULE_NAME = "IO";
        this.VERSION = "1.0.0";
        this.AUTHOR = "郑广学 (EXCEL880)";

        // 文件系统对象 (WPS JSA环境)
        this._fso = null;
    }

    /**
     * 获取文件系统对象（辅助方法）
     * @private
     * @returns {Object} 文件系统对象
     */
    _getFSO() {
        if (!this._fso) {
            // WPS JSA环境使用ActiveXObject创建FileSystemObject
            this._fso = new ActiveXObject("Scripting.FileSystemObject");
        }
        return this._fso;
    }

    /**
     * 判断给定路径是否为文件
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为文件
     * @example
     * var isFile = IO.z是否文件('/path/to/file.txt');
     * console.log(isFile); // true
     */
    z是否文件(path) {
        try {
            const fso = this._getFSO();
            if (fso.FileExists(path)) {
                return true;
            }
            // 如果路径存在且不是文件夹，则是文件
            if (fso.FolderExists(path)) {
                return false;
            }
            return false;
        } catch (e) {
            console.error("判断文件失败:", e);
            return false;
        }
    }

    /**
     * 判断给定路径是否为文件（英文别名）
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为文件
     */
    IsFile(path) {
        return this.z是否文件(path);
    }

    /**
     * 判断给定路径是否为文件夹
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为文件夹
     * @example
     * var isDirectory = IO.z是否文件夹('/path/to/folder');
     * console.log(isDirectory); // true
     */
    z是否文件夹(path) {
        try {
            const fso = this._getFSO();
            return fso.FolderExists(path);
        } catch (e) {
            console.error("判断文件夹失败:", e);
            return false;
        }
    }

    /**
     * 判断给定路径是否为文件夹（英文别名）
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为文件夹
     */
    IsDirectory(path) {
        return this.z是否文件夹(path);
    }

    /**
     * 判断给定路径是否为隐藏文件或文件夹
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为隐藏文件或文件夹
     * @example
     * var isHidden = IO.z是否隐藏('/path/to/file.txt');
     * console.log(isHidden); // false
     */
    z是否隐藏(path) {
        try {
            const fso = this._getFSO();
            if (fso.FileExists(path)) {
                const file = fso.GetFile(path);
                return file.Attributes && (file.Attributes & 2) === 2; // 2 = Hidden attribute
            }
            if (fso.FolderExists(path)) {
                const folder = fso.GetFolder(path);
                return folder.Attributes && (folder.Attributes & 2) === 2;
            }
            return false;
        } catch (e) {
            console.error("判断隐藏属性失败:", e);
            return false;
        }
    }

    /**
     * 判断给定路径是否为隐藏文件或文件夹（英文别名）
     * @param {String} path - 要判断的路径
     * @returns {Boolean} 给定路径是否为隐藏文件或文件夹
     */
    IsHidden(path) {
        return this.z是否隐藏(path);
    }

    /**
     * 获取给定路径的文件名（包含后缀）
     * @param {String} path - 要获取文件名的路径
     * @returns {String} 文件名（包含后缀）
     * @example
     * var fileName = IO.z文件名('/path/to/file.txt');
     * console.log(fileName); // 'file.txt'
     */
    z文件名(path) {
        if (!path) return '';
        const parts = path.replace(/\\/g, '/').split('/');
        return parts[parts.length - 1] || '';
    }

    /**
     * 获取给定路径的文件名（包含后缀，英文别名）
     * @param {String} path - 要获取文件名的路径
     * @returns {String} 文件名（包含后缀）
     */
    getFileName(path) {
        return this.z文件名(path);
    }

    /**
     * 获取给定路径的纯文件名（不包含后缀）
     * @param {String} path - 要获取纯文件名的路径
     * @returns {String} 纯文件名（不包含后缀）
     * @example
     * var fileName = IO.z纯文件名('/path/to/file.txt');
     * console.log(fileName); // 'file'
     */
    z纯文件名(path) {
        const fileName = this.z文件名(path);
        const lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex > 0) {
            return fileName.substring(0, lastDotIndex);
        }
        return fileName;
    }

    /**
     * 获取给定路径的纯文件名（不包含后缀，英文别名）
     * @param {String} path - 要获取纯文件名的路径
     * @returns {String} 纯文件名（不包含后缀）
     */
    getFileNameNoType(path) {
        return this.z纯文件名(path);
    }

    /**
     * 获取给定路径的文件后缀
     * @param {String} path - 要获取文件后缀的路径
     * @returns {String} 文件后缀
     * @example
     * var fileType = IO.z文件后缀('/path/to/file.txt');
     * console.log(fileType); // 'txt'
     */
    z文件后缀(path) {
        const fileName = this.z文件名(path);
        const lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex > 0 && lastDotIndex < fileName.length - 1) {
            return fileName.substring(lastDotIndex + 1);
        }
        return '';
    }

    /**
     * 获取给定路径的文件后缀（英文别名）
     * @param {String} path - 要获取文件后缀的路径
     * @returns {String} 文件后缀
     */
    getFileType(path) {
        return this.z文件后缀(path);
    }

    /**
     * 遍历给定文件夹路径下的文件
     * @param {String} 文件夹路径 - 要遍历的文件夹路径
     * @param {Boolean} 遍历子文件 - 是否遍历子文件夹中的文件
     * @param {Boolean} 不包含隐藏 - 是否包含隐藏文件和文件夹
     * @param {Boolean} 结果包含文件夹 - 遍历结果是否包含文件夹
     * @returns {Array} 包含文件路径的数组
     * @example
     * var files = IO.z遍历文件('/path/to/folder', true, true, false);
     * console.log(files); // ['/path/to/folder/file1.txt', ...]
     */
    z遍历文件(文件夹路径, 遍历子文件, 不包含隐藏, 结果包含文件夹) {
        遍历子文件 = 遍历子文件 !== undefined ? 遍历子文件 : false;
        不包含隐藏 = 不包含隐藏 !== undefined ? 不包含隐藏 : true;
        结果包含文件夹 = 结果包含文件夹 !== undefined ? 结果包含文件夹 : false;

        const result = [];
        this._traverseFiles(文件夹路径, 遍历子文件, 不包含隐藏, 结果包含文件夹, result);
        return result;
    }

    /**
     * 遍历文件辅助方法（递归）
     * @private
     */
    _traverseFiles(folderPath, recursive, excludeHidden, includeFolders, result) {
        try {
            const fso = this._getFSO();
            if (!fso.FolderExists(folderPath)) {
                return;
            }

            const folder = fso.GetFolder(folderPath);

            // 遍历文件
            const files = new Enumerator(folder.Files);
            while (!files.atEnd()) {
                const file = files.item();
                const filePath = file.Path;

                // 检查是否排除隐藏文件
                if (excludeHidden && this.z是否隐藏(filePath)) {
                    files.moveNext();
                    continue;
                }

                result.push(filePath);
                files.moveNext();
            }

            // 遍历子文件夹
            if (recursive) {
                const subFolders = new Enumerator(folder.SubFolders);
                while (!subFolders.atEnd()) {
                    const subFolder = subFolders.item();
                    const subFolderPath = subFolder.Path;

                    // 检查是否排除隐藏文件夹
                    if (excludeHidden && this.z是否隐藏(subFolderPath)) {
                        subFolders.moveNext();
                        continue;
                    }

                    // 如果需要包含文件夹
                    if (includeFolders) {
                        result.push(subFolderPath);
                    }

                    // 递归遍历子文件夹
                    this._traverseFiles(subFolderPath, recursive, excludeHidden, includeFolders, result);
                    subFolders.moveNext();
                }
            }
        } catch (e) {
            console.error("遍历文件失败:", e);
        }
    }

    /**
     * 遍历给定文件夹路径下的文件（英文别名）
     * @param {String} 文件夹路径 - 要获取文件的文件夹路径
     * @param {Boolean} 遍历子文件 - 是否遍历子文件夹中的文件
     * @param {Boolean} 不包含隐藏 - 是否包含隐藏文件和文件夹
     * @param {Boolean} 结果包含文件夹 - 获取结果是否包含文件夹
     * @returns {Array} 包含文件路径的数组
     */
    getFiles(文件夹路径, 遍历子文件, 不包含隐藏, 结果包含文件夹) {
        return this.z遍历文件(文件夹路径, 遍历子文件, 不包含隐藏, 结果包含文件夹);
    }

    /**
     * 遍历给定文件夹路径下的文件夹
     * @param {String} 文件夹路径 - 要遍历的文件夹路径
     * @param {Boolean} 遍历子文件 - 是否遍历子文件夹中的文件夹
     * @param {Boolean} 不包含隐藏 - 是否包含隐藏文件夹
     * @returns {Array} 包含文件夹路径的数组
     * @example
     * var directories = IO.z遍历文件夹('/path/to/folder', true, true);
     * console.log(directories); // ['/path/to/folder/subfolder1', ...]
     */
    z遍历文件夹(文件夹路径, 遍历子文件, 不包含隐藏) {
        遍历子文件 = 遍历子文件 !== undefined ? 遍历子文件 : false;
        不包含隐藏 = 不包含隐藏 !== undefined ? 不包含隐藏 : true;

        const result = [];
        this._traverseFolders(文件夹路径, 遍历子文件, 不包含隐藏, result);
        return result;
    }

    /**
     * 遍历文件夹辅助方法（递归）
     * @private
     */
    _traverseFolders(folderPath, recursive, excludeHidden, result) {
        try {
            const fso = this._getFSO();
            if (!fso.FolderExists(folderPath)) {
                return;
            }

            const folder = fso.GetFolder(folderPath);

            // 遍历子文件夹
            const subFolders = new Enumerator(folder.SubFolders);
            while (!subFolders.atEnd()) {
                const subFolder = subFolders.item();
                const subFolderPath = subFolder.Path;

                // 检查是否排除隐藏文件夹
                if (excludeHidden && this.z是否隐藏(subFolderPath)) {
                    subFolders.moveNext();
                    continue;
                }

                result.push(subFolderPath);

                // 递归遍历子文件夹
                if (recursive) {
                    this._traverseFolders(subFolderPath, recursive, excludeHidden, result);
                }
                subFolders.moveNext();
            }
        } catch (e) {
            console.error("遍历文件夹失败:", e);
        }
    }

    /**
     * 遍历给定文件夹路径下的文件夹（英文别名）
     * @param {String} 文件夹路径 - 要获取文件夹的文件夹路径
     * @param {Boolean} 遍历子文件 - 是否遍历子文件夹中的文件夹
     * @param {Boolean} 不包含隐藏 - 是否包含隐藏文件夹
     * @returns {Array} 包含文件夹路径的数组
     */
    getDirectorys(文件夹路径, 遍历子文件, 不包含隐藏) {
        return this.z遍历文件夹(文件夹路径, 遍历子文件, 不包含隐藏);
    }

    /**
     * 获取给定路径的上级文件夹路径
     * @param {String} path - 要获取上级文件夹路径的路径
     * @param {Number} 返回级数 - 返回上级文件夹的级数
     * @returns {String} 上级文件夹路径
     * @example
     * var parentFolder = IO.z上级文件夹('/path/to/file.txt');
     * console.log(parentFolder); // '/path/to'
     */
    z上级文件夹(path, 返回级数) {
        返回级数 = 返回级数 !== undefined ? 返回级数 : 1;
        let result = path;

        for (let i = 0; i < 返回级数; i++) {
            // 统一斜杠
            result = result.replace(/\\/g, '/');
            // 去除末尾斜杠
            result = result.replace(/\/+$/, '');
            // 获取上级目录
            const lastSlashIndex = result.lastIndexOf('/');
            if (lastSlashIndex > 0) {
                result = result.substring(0, lastSlashIndex);
            } else {
                break;
            }
        }
        return result;
    }

    /**
     * 获取给定路径的上级文件夹路径（英文别名）
     * @param {String} 路径 - 要获取上级文件夹路径的路径
     * @param {Number} 返回级数 - 返回上级文件夹的级数
     * @returns {String} 上级文件夹路径
     */
    lastDirectoty(路径, 返回级数) {
        return this.z上级文件夹(路径, 返回级数);
    }

    /**
     * 复制源路径下的文件到目标路径
     * @param {String} 源路径 - 要复制的文件路径
     * @param {String} targetpath - 目标文件路径
     * @returns {String} 复制后的文件路径
     * @example
     * var copiedFile = IO.z复制文件('/path/to/source.txt', '/path/to/target.txt');
     * console.log(copiedFile); // '/path/to/target.txt'
     */
    z复制文件(源路径, targetpath) {
        try {
            const fso = this._getFSO();
            // 确保目标文件夹存在
            const targetFolder = this.z上级文件夹(targetpath);
            if (!fso.FolderExists(targetFolder)) {
                fso.CreateFolder(targetFolder);
            }
            // 复制文件
            fso.CopyFile(源路径, targetpath, true); // true表示覆盖
            return targetpath;
        } catch (e) {
            console.error("复制文件失败:", e);
            return null;
        }
    }

    /**
     * 复制源路径下的文件到目标路径（英文别名）
     * @param {String} 源路径 - 要复制的文件路径
     * @param {String} targetpath - 目标文件路径
     * @returns {Boolean} 复制是否成功
     */
    copyFile(源路径, targetpath) {
        const result = this.z复制文件(源路径, targetpath);
        return result !== null;
    }

    /**
     * 移动源路径下的文件到目标路径
     * @param {String} 源路径 - 要移动的文件路径
     * @param {String} targetpath - 目标文件路径
     * @example
     * IO.z移动文件('/path/to/source.txt', '/path/to/target.txt');
     */
    z移动文件(源路径, targetpath) {
        try {
            const fso = this._getFSO();
            // 确保目标文件夹存在
            const targetFolder = this.z上级文件夹(targetpath);
            if (!fso.FolderExists(targetFolder)) {
                fso.CreateFolder(targetFolder);
            }
            // 移动文件
            fso.MoveFile(源路径, targetpath);
        } catch (e) {
            console.error("移动文件失败:", e);
        }
    }

    /**
     * 移动源路径下的文件到目标路径（英文别名）
     * @param {String} 源路径 - 要移动的文件路径
     * @param {String} targetpath - 目标文件路径
     * @returns {Boolean} 移动是否成功
     */
    moveFile(源路径, targetpath) {
        try {
            this.z移动文件(源路径, targetpath);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 重命名源路径下的文件或文件夹
     * @param {String} 源路径 - 要重命名的文件或文件夹路径
     * @param {String} 新名称 - 新的名称
     * @example
     * IO.z重命名('/path/to/file.txt', 'newname.txt');
     */
    z重命名(源路径, 新名称) {
        try {
            const fso = this._getFSO();
            // 构建新路径
            const parentPath = this.z上级文件夹(源路径);
            const newPath = parentPath + '/' + 新名称;

            if (fso.FileExists(源路径)) {
                fso.MoveFile(源路径, newPath);
            } else if (fso.FolderExists(源路径)) {
                fso.MoveFolder(源路径, newPath);
            }
        } catch (e) {
            console.error("重命名失败:", e);
        }
    }

    /**
     * 重命名源路径下的文件或文件夹（英文别名）
     * @param {String} 源路径 - 要重命名的文件或文件夹路径
     * @param {String} 新名称 - 新的名称
     * @returns {Boolean} 重命名是否成功
     */
    rename(源路径, 新名称) {
        try {
            this.z重命名(源路径, 新名称);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 将内容写入指定路径的文本文件中
     * @param {String} 路径 - 要写入的文本文件路径
     * @param {String} 内容 - 要写入的内容
     * @example
     * IO.z写入文本文件('/path/to/file.txt', 'Hello, World!');
     */
    z写入文本文件(路径, 内容) {
        try {
            const fso = this._getFSO();
            // 确保文件夹存在
            const folderPath = this.z上级文件夹(路径);
            if (!fso.FolderExists(folderPath)) {
                fso.CreateFolder(folderPath);
            }

            // 使用AdoDb.Stream写入UTF-8文件
            const stream = new ActiveXObject("ADODB.Stream");
            stream.Type = 2; // adTypeText (文本)
            stream.Charset = "utf-8";
            stream.Open();
            stream.WriteText(内容);
            stream.SaveToFile(路径, 2); // 2 = adSaveCreateOverWrite
            stream.Close();
        } catch (e) {
            console.error("写入文件失败:", e);
        }
    }

    /**
     * 将内容写入指定路径的文本文件中（英文别名）
     * @param {String} path - 要写入的文本文件路径
     * @param {String} content - 要写入的内容
     * @returns {Boolean} 写入是否成功
     */
    writefile(path, content) {
        try {
            this.z写入文本文件(path, content);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 读取指定路径的文件内容
     * @param {String} 路径 - 要读取的文件路径
     * @returns {String} 文件内容
     * @example
     * var fileContent = IO.z读取文件('/path/to/file.txt');
     * console.log(fileContent); // 'Hello, World!'
     */
    z读取文件(路径) {
        try {
            const fso = this._getFSO();
            if (!fso.FileExists(路径)) {
                console.error("文件不存在:", 路径);
                return "";
            }

            // 使用AdoDb.Stream读取UTF-8文件
            const stream = new ActiveXObject("ADODB.Stream");
            stream.Type = 2; // adTypeText (文本)
            stream.Charset = "utf-8";
            stream.Open();
            stream.LoadFromFile(路径);
            const content = stream.ReadText(-1); // -1 = adReadAll
            stream.Close();
            return content;
        } catch (e) {
            console.error("读取文件失败:", e);
            return "";
        }
    }

    /**
     * 读取指定路径的文件内容（英文别名）
     * @param {String} path - 要读取的文件路径
     * @returns {String} 文件内容
     */
    readfile(path) {
        return this.z读取文件(path);
    }

    /**
     * 创建文件夹
     * @param {String} path - 要创建的文件夹路径
     * @returns {String} 创建的文件夹路径
     * @example
     * var folderPath = IO.z创建文件夹('/path/to/folder');
     * console.log(folderPath); // '/path/to/folder'
     */
    z创建文件夹(path) {
        try {
            const fso = this._getFSO();
            // 如果文件夹不存在，则创建（包括父文件夹）
            if (!fso.FolderExists(path)) {
                fso.CreateFolder(path);
            }
            return path;
        } catch (e) {
            console.error("创建文件夹失败:", e);
            // 尝试递归创建
            this._createFolderRecursive(path);
            return path;
        }
    }

    /**
     * 递归创建文件夹
     * @private
     */
    _createFolderRecursive(path) {
        try {
            const fso = this._getFSO();
            if (fso.FolderExists(path)) {
                return;
            }
            // 获取父文件夹
            const parentPath = this.z上级文件夹(path);
            if (parentPath && parentPath !== path) {
                this._createFolderRecursive(parentPath);
            }
            if (!fso.FolderExists(path)) {
                fso.CreateFolder(path);
            }
        } catch (e) {
            // 忽略已存在的错误
        }
    }

    /**
     * 创建文件夹（英文别名）
     * @param {String} path - 要创建的文件夹路径
     * @returns {String} 创建的文件夹路径
     */
    MkDir2(path) {
        return this.z创建文件夹(path);
    }

    /**
     * 删除指定路径的文件
     * @param {String} path - 要删除的文件路径
     * @example
     * IO.z删除文件('/path/to/file.txt');
     */
    z删除文件(path) {
        try {
            const fso = this._getFSO();
            if (fso.FileExists(path)) {
                fso.DeleteFile(path, true); // true表示强制删除（只读文件）
            }
        } catch (e) {
            console.error("删除文件失败:", e);
        }
    }

    /**
     * 删除指定路径的文件（英文别名）
     * @param {String} path - 要删除的文件路径
     * @returns {String} 被删除的文件路径
     */
    Delete(path) {
        try {
            this.z删除文件(path);
            return path;
        } catch (e) {
            return null;
        }
    }

    /**
     * 删除指定路径的目录
     * @param {String} path - 要删除的目录路径
     * @example
     * IO.z删除目录('/path/to/folder');
     */
    z删除目录(path) {
        try {
            const fso = this._getFSO();
            if (fso.FolderExists(path)) {
                fso.DeleteFolder(path, true); // true表示强制删除（只读文件）
            }
        } catch (e) {
            console.error("删除目录失败:", e);
        }
    }

    /**
     * 递归删除指定路径的目录及其内容（英文别名）
     * @param {String} path - 要删除的目录路径
     * @returns {Boolean} 删除是否成功
     */
    deleteTree(path) {
        try {
            this.z删除目录(path);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 处理路径中的斜杠，将反斜杠转换为正斜杠
     * @param {String} path - 要处理的路径
     * @returns {String} 处理后的路径
     * @example
     * var processedPath = IO.z路径斜杠处理('C:\\path\\to\\file.txt');
     * console.log(processedPath); // 'C:/path/to/file.txt'
     */
    z路径斜杠处理(path) {
        if (!path) return '';
        return path.replace(/\\/g, '/');
    }

    /**
     * 处理路径中的斜杠（英文别名）
     * @param {String} path - 要处理的路径
     * @returns {String} 处理后的路径字符串
     */
    correctPath(path) {
        return this.z路径斜杠处理(path);
    }

    /**
     * 显示文件选择对话框
     * @param {String} path - 显示文件对话框的初始路径
     * @returns {String} 选择的文件路径
     * @example
     * var filePath = IO.showFileDialog('/path/to/folder');
     */
    showFileDialog(path) {
        try {
            var dialog = new ActiveXObject("UserAccounts.CommonDialog");
            dialog.InitialDir = path || "";
            dialog.Filter = "All Files|*.*";
            if (dialog.ShowOpen()) {
                return dialog.FileName;
            }
            return "";
        } catch (e) {
            console.error("显示文件对话框失败:", e);
            // 备用方案: 使用Application对象
            try {
                return Application.GetOpenFileName();
            } catch (e2) {
                console.error("备用文件对话框也失败:", e2);
                return "";
            }
        }
    }

    /**
     * 显示文件夹选择对话框
     * @param {String} path - 显示文件夹对话框的初始路径
     * @returns {String} 选择的文件夹路径
     * @example
     * var folderPath = IO.showFolderDialog('/path/to/folder');
     */
    showFolderDialog(path) {
        try {
            // 使用Shell.Application对象
            var shell = new ActiveXObject("Shell.Application");
            var folder = shell.BrowseForFolder(0, "请选择文件夹", 0, path);
            if (folder) {
                return folder.Self.Path;
            }
            return "";
        } catch (e) {
            console.error("显示文件夹对话框失败:", e);
            return "";
        }
    }
}

/**
 * 创建IO全局实例 - 可直接使用 IO.方法名() 调用
 * @example
 * IO.z是否文件('/path/to/file.txt');
 */
const IO = new clsIO();

/**
 * 创建DateUtils全局实例 - 可直接使用 DateUtils.方法名() 调用
 * @example
 * DateUtils.z转表格日期(new Date());
 */
const DateUtils = new clsDateUtils();

/**
 * 导出类 - 支持WPS JSA环境
 */
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { JSA, ShtUtils, Array2D, DateUtils, IO };
}
// WPS JSA环境：导出为全局变量
if (typeof window !== 'undefined' || typeof Application !== 'undefined') {
    this.JSA = JSA;
    this.ShtUtils = ShtUtils;
    this.Array2D = Array2D;
    this.DateUtils = DateUtils;
    this.IO = IO;
}
