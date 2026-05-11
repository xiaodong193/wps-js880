/**
 * ============== 共享常量模块（V3 重构版）==============
 * 作者：徐晓冬
 * 描述：定义所有模块共享的常量和工具函数
 * 版本：V3
 * 最后修改：20260501 全面重构
 *
 * V3 变更：
 * - [重构] 统一颜色管理：新增 COLORS 色板对象，所有颜色集中定义
 * - [修复] COLOR_LIGHT_BLUE / COLOR_LIGHT_GREEN / COLOR_LIGHT_YELLOW 颜色值与注释不符
 * - [重构] 应用格式() 改用查找表替代 switch-case
 * - [修复] 设置字体样式() falsy 值检查（size=0 被跳过的问题）
 * - [修复] createDataHandler() endCol 计算逻辑和硬编码 A:M 范围
 * - [优化] createDataHandler() 改为脏标记模式，避免每次 setCell 都写回整张表
 * - [清理] 移除未使用的 logjson()、arrDataFromRngExtended() 死代码
 * - [优化] XL 对象移除冗余兼容性别名（保留 xlDown 因银承模块使用）
 * - [优化] 统一错误日志级别：验证用 warn，运行时用 error
 * - [优化] 使用 Object.freeze() 冻结枚举常量对象
 *
 * 核心功能：
 * - 统一管理所有常量定义
 * - 提供 Excel/WPS 内置常量映射
 * - 定义格式、字体、颜色等样式常量（COLORS 统一色板）
 * - 提供版本信息管理
 * - 提供样式工具函数
 * ====================================================
 */

// ========== 模块名称 ==========
const MODULE_NAME = "mShared_constants";

// ========== 版本信息 ==========
const SYSTEM_VERSION = {
    MAJOR: 3,
    MINOR: 2026,
    PATCH: 5,
    BUILD: 1,  // 2026-05-01 V3 重构
    DATE: "20260501",
    toString() {
        return `${this.MAJOR}.${this.MINOR}.${this.PATCH}.${this.BUILD}`;
    }
};

const VERSION = SYSTEM_VERSION.toString();
const VERSION_DATE = SYSTEM_VERSION.DATE;

console.log(`[mShared_constants] 系统版本: ${VERSION}`);

// ========== Excel/WPS 内置常量 ==========
const XL = Object.freeze({
    // 对齐方式
    HCenter: -4108,        // 水平居中
    VCenter: -4108,        // 垂直居中
    Left: -4131,           // 左对齐
    Right: -4152,          // 右对齐
    Top: -4160,            // 顶端对齐
    Bottom: -4107,         // 底端对齐
    General: -4143,        // 常规

    // 边框样式
    Continuous: 1,         // 连续线
    Dash: -4115,           // 虚线
    Dot: -4118,            // 点线
    Double: -4119,         // 双线
    Thin: 2,               // 细线
    Medium: -4138,         // 中等线
    Thick: 4,              // 粗线
    None: -4142,           // 无边框

    // 方向
    Down: -4121,           // 向下
    Up: -4162,             // 向上
    ToLeft: -4159,         // 向左
    ToRight: -4161,        // 向右
    xlDown: -4121,         // 向下（兼容 m银行承兑汇票模块.js 使用）

    // 线条粗细
    Hairline: 1            // 极细线
});

// ========== RGB 颜色转换函数 ==========
/**
 * RGB - RGB 颜色转换函数
 *
 * 作用：将 RGB 值转换为 WPS 可用的颜色值
 * 注意：必须在 COLORS 色板之前定义
 *
 * @param {number} r - 红色分量（0-255）
 * @param {number} g - 绿色分量（0-255）
 * @param {number} b - 蓝色分量（0-255）
 * @returns {number} WPS 颜色值（r + g * 256 + b * 65536）
 */
function RGB(r, g, b) {
    if (typeof r !== 'number' || typeof g !== 'number' || typeof b !== 'number') {
        console.warn(`[mShared_constants] RGB 值必须为数字: r=${r}, g=${g}, b=${b}`);
        return 0;
    }
    r = Math.max(0, Math.min(255, Math.floor(r)));
    g = Math.max(0, Math.min(255, Math.floor(g)));
    b = Math.max(0, Math.min(255, Math.floor(b)));
    return r + g * 256 + b * 65536;
}

// ========== 统一色板 COLORS ==========
/**
 * COLORS - 统一颜色管理色板
 *
 * 作用：全项目唯一的颜色定义来源（Single Source of Truth）
 * 使用：rng.Interior.Color = COLORS.HEADER_BLUE;
 *
 * 所有模块（mParameterManager、mStyleManager 等）应引用此色板，
 * 不再各自独立计算颜色值。
 *
 * 注意：依赖上方定义的 RGB() 函数
 */
const COLORS = Object.freeze({
    // ── 基础色 ──
    WHITE:          RGB(255, 255, 255),
    BLACK:          RGB(0,   0,   0),
    RED:            RGB(255, 0,   0),
    GREEN:          RGB(0,   255, 0),
    BLUE:           RGB(0,   0,   255),
    YELLOW:         RGB(255, 255, 0),
    CYAN:           RGB(0,   255, 255),
    MAGENTA:        RGB(255, 0,   255),

    // ── 灰度 ──
    GRAY:           RGB(128, 128, 128),
    LIGHT_GRAY:     RGB(192, 192, 192),
    DARK_GRAY:      RGB(64,  64,  64),

    // ── 语义色（业务用途） ──
    HEADER_BLUE:    RGB(0,   174, 240),    // 表头背景色（天蓝色）
    LIGHT_GREEN:    RGB(144, 238, 144),    // 正向指标背景色
    LIGHT_RED:      RGB(255, 204, 204),    // 负向/警示背景色
    LIGHT_YELLOW:   RGB(255, 255, 224),    // 租前期行高亮色
    PARAM_AREA_BG:  RGB(173, 216, 230),    // 参数区标题背景色
});

// ========== 向后兼容颜色别名 ==========
// 保留旧 COLOR_* 变量名，指向 COLORS 统一色板
// 其他模块中已有的 COLOR_WHITE 等引用无需修改
const COLOR_WHITE = COLORS.WHITE;
const COLOR_BLACK = COLORS.BLACK;
const COLOR_RED = COLORS.RED;
const COLOR_GREEN = COLORS.GREEN;
const COLOR_BLUE = COLORS.BLUE;
const COLOR_YELLOW = COLORS.YELLOW;
const COLOR_CYAN = COLORS.CYAN;
const COLOR_MAGENTA = COLORS.MAGENTA;
const COLOR_GRAY = COLORS.GRAY;
const COLOR_LIGHT_GRAY = COLORS.LIGHT_GRAY;
const COLOR_DARK_GRAY = COLORS.DARK_GRAY;
const COLOR_LIGHT_BLUE = COLORS.HEADER_BLUE;      // 语义映射：LIGHT_BLUE → 表头蓝
const COLOR_LIGHT_GREEN = COLORS.LIGHT_GREEN;
const COLOR_LIGHT_RED = COLORS.LIGHT_RED;
const COLOR_LIGHT_YELLOW = COLORS.LIGHT_YELLOW;

// ========== 格式常量 ==========
const FORMAT_STANDARD = "#,##0.00";
const FORMAT_INTEGER = "0";
const FORMAT_DATE = "yyyy-mm-dd";
const FORMAT_PERCENTAGE = "0.00%";
const FORMAT_TEXT = "@";
const FORMAT_CURRENCY = "¥#,##0.00";
const FORMAT_ACCOUNTING = "_¥* #,##0.00_ ;_¥* -#,##0.00_ ;_¥* \"-\"??_ ;_ @_ ";

/**
 * FORMAT_MAP - 格式类型查找表
 *
 * 作用：替代 switch-case，新增格式只需在此添加一行
 */
const FORMAT_MAP = Object.freeze({
    Standard: FORMAT_STANDARD,
    Integer: FORMAT_INTEGER,
    Date: FORMAT_DATE,
    Text: FORMAT_TEXT,
    Percentage: FORMAT_PERCENTAGE,
    Currency: FORMAT_CURRENCY,
    Accounting: FORMAT_ACCOUNTING,
});

// ========== 字体常量 ==========
const FONT_CHINESE = "微软雅黑";
const FONT_ENGLISH = "Arial";
const FONT_DEFAULT = "宋体";
const FONT_SIZE_TITLE = 16;
const FONT_SIZE_HEADER = 12;
const FONT_SIZE_NORMAL = 10;
const FONT_SIZE_SMALL = 9;

// ========== 表格常量 ==========
const TABLE = {
    MIN_ROW_HEIGHT: 15,
    MIN_COLUMN_WIDTH: 20,
    DEFAULT_ROW_HEIGHT: 18,
    DEFAULT_COLUMN_WIDTH: 10,
    HEADER_BACKGROUND_COLOR: COLORS.HEADER_BLUE,
    HEADER_FONT_COLOR: COLORS.BLACK,
    HEADER_FONT_BOLD: true,
    HEADER_BORDER_COLOR: COLORS.GRAY
};

// ========== 列索引常量 ==========
const COL = Object.freeze({
    PERIOD: 0,              // 期次
    DATE: 1,                // 支付日
    RENT: 2,                // 租金
    PRINCIPAL: 3,           // 本金
    INTEREST: 4,            // 利息
    CUMULATIVE_PRINCIPAL: 5, // 累积偿还本金额
    PRINCIPAL_BALANCE: 6,   // 租金本金余额
    REMAINING_BALANCE: 7,   // 剩余租金余额
    PAID_RENT: 8,           // 已还租金
    MONTH_INTERVAL: 9,      // 支付日/月间隔
    CUSTOM_INTERVAL: 10,    // 支付日/月间隔-自定义
    PRINCIPAL_RATIO: 11,    // 本金比例
    RATE_PER_PERIOD: 12     // 每期适用利率
});

const CF_COL = Object.freeze({
    PERIOD: 0,              // 期次
    DATE: 1,                // 日期
    NET_CASHFLOW_1: 2,      // 净现金流1
    REMARK_1: 3,            // 净现金流1-备注
    NET_CASHFLOW_2: 4,      // 净现金流2
    REMARK_2: 5,            // 净现金流2-备注
    WIRE_TRANSFER: 6,       // 电汇放款
    RENT_PAYMENT: 7,        // 租金偿付
    DEPOSIT: 8,             // 保证金
    NOMINAL_PRICE: 9,       // 名义货价
    BROKER_FEE: 10          // 经纪人费用
});

const FORMULA_ROW = Object.freeze({
    HEADER: 0,   // 表头行
    FIRST: 1,    // 首期
    MIDDLE: 2,   // 中间期
    LAST: 3,     // 末期
    EXTRA: 4     // 预留
});

// ========== 错误级别 ==========
const ERROR_LEVELS = Object.freeze({
    INFO: "INFO",
    WARNING: "WARNING",
    ERROR: "ERROR",
    CRITICAL: "CRITICAL"
});

// ========== 样式工具函数 ==========

/**
 * 应用格式 - 应用数字格式到单元格范围
 *
 * @param {Range} rng - 单元格范围对象
 * @param {string} formatType - 格式类型（见 FORMAT_MAP）
 * @returns {boolean} 是否应用成功
 */
function 应用格式(rng, formatType) {
    try {
        if (!rng) {
            console.warn(`[mShared_constants] 应用格式: 单元格范围为空`);
            return false;
        }

        const format = FORMAT_MAP[formatType];
        if (!format) {
            console.warn(`[mShared_constants] 未知的格式类型: ${formatType}`);
            return false;
        }

        rng.NumberFormat = format;
        return true;
    } catch (error) {
        console.error(`[mShared_constants] 应用格式失败: ${error.message}`);
        return false;
    }
}

/**
 * 设置表格样式 - 设置表格基本样式（边框 + 居中 + 自动换行）
 *
 * @param {Range} rng - 单元格范围对象
 * @returns {boolean} 是否设置成功
 */
function 设置表格样式(rng) {
    try {
        if (!rng) {
            console.warn(`[mShared_constants] 设置表格样式: 单元格范围为空`);
            return false;
        }

        rng.Borders.LineStyle = XL.Continuous;
        rng.Borders.Weight = XL.Thin;
        rng.Borders.Color = COLORS.BLACK;
        rng.HorizontalAlignment = XL.HCenter;
        rng.VerticalAlignment = XL.VCenter;
        rng.WrapText = true;

        return true;
    } catch (error) {
        console.error(`[mShared_constants] 设置表格样式失败: ${error.message}`);
        return false;
    }
}

/**
 * 设置字体样式 - 设置字体样式
 *
 * @param {Range} rng - 单元格范围对象
 * @param {Object} options - 样式选项
 * @param {string} [options.name] - 字体名称
 * @param {number} [options.size] - 字体大小
 * @param {boolean} [options.bold] - 是否加粗
 * @param {boolean} [options.italic] - 是否斜体
 * @param {number} [options.color] - 字体颜色
 * @returns {boolean} 是否设置成功
 */
function 设置字体样式(rng, options = {}) {
    try {
        if (!rng) {
            console.warn(`[mShared_constants] 设置字体样式: 单元格范围为空`);
            return false;
        }

        if (options.name !== undefined) rng.Font.Name = options.name;
        if (options.size !== undefined) rng.Font.Size = options.size;
        if (options.bold !== undefined) rng.Font.Bold = options.bold;
        if (options.italic !== undefined) rng.Font.Italic = options.italic;
        if (options.color !== undefined) rng.Font.Color = options.color;

        return true;
    } catch (error) {
        console.error(`[mShared_constants] 设置字体样式失败: ${error.message}`);
        return false;
    }
}

/**
 * 设置背景颜色 - 设置单元格背景颜色
 *
 * @param {Range} rng - 单元格范围对象
 * @param {number} color - 颜色值（使用 COLORS.* 或 RGB()）
 * @returns {boolean} 是否设置成功
 */
function 设置背景颜色(rng, color) {
    try {
        if (!rng) {
            console.warn(`[mShared_constants] 设置背景颜色: 单元格范围为空`);
            return false;
        }
        rng.Interior.Color = color;
        return true;
    } catch (error) {
        console.error(`[mShared_constants] 设置背景颜色失败: ${error.message}`);
        return false;
    }
}

/**
 * 设置边框 - 设置单元格边框
 *
 * @param {Range} rng - 单元格范围对象
 * @param {Object} [options] - 边框选项
 * @param {number} [options.lineStyle=XL.Continuous] - 线条样式
 * @param {number} [options.weight=XL.Thin] - 线条粗细
 * @param {number} [options.color=COLORS.BLACK] - 边框颜色
 * @returns {boolean} 是否设置成功
 */
function 设置边框(rng, options = {}) {
    try {
        if (!rng) {
            console.warn(`[mShared_constants] 设置边框: 单元格范围为空`);
            return false;
        }

        rng.Borders.LineStyle = options.lineStyle !== undefined ? options.lineStyle : XL.Continuous;
        rng.Borders.Weight = options.weight !== undefined ? options.weight : XL.Thin;
        rng.Borders.Color = options.color !== undefined ? options.color : COLORS.BLACK;

        return true;
    } catch (error) {
        console.error(`[mShared_constants] 设置边框失败: ${error.message}`);
        return false;
    }
}

/**
 * 列号转字母 - 将列号（1-based）转为 Excel 列字母
 *
 * @param {number} colNum - 列号（1=A, 2=B, ..., 27=AA）
 * @returns {string} 列字母
 *
 * 示例：
 * colToLetter(1)   → "A"
 * colToLetter(26)  → "Z"
 * colToLetter(27)  → "AA"
 */
function colToLetter(colNum) {
    var letter = "";
    var col = colNum;
    while (col > 0) {
        col--;
        letter = String.fromCharCode(65 + (col % 26)) + letter;
        col = Math.floor(col / 26);
    }
    return letter;
}

/**
 * arrDataFromRngExtended - 增强型数组读取函数
 *
 * 作用：从 Excel 范围读取数据并返回带操作方法的数据处理器
 * 设计：脏标记模式，修改操作只更新内存，需显式调用 syncToSheet() 写回
 *
 * @param {Worksheet} sheet - 工作表对象
 * @param {number} startRow - 起始行号
 * @param {Array} arrHeaders - 表头数组
 * @returns {Object} 数据处理器对象
 */
function arrDataFromRngExtended(sheet, startRow, arrHeaders) {
    const colCount = arrHeaders.length;
    const DEFAULT_RECOVERY_ROWS = 5;

    const createEmptyArray = () => {
        const arr = [];
        for (var i = 0; i < DEFAULT_RECOVERY_ROWS; i++) {
            arr[i] = new Array(colCount).fill("");
        }
        return arr;
    };

    try {
        var usedRange;
        try {
            usedRange = sheet.UsedRange;
        } catch (e) {
            console.warn(`[mShared_constants] 无法获取已使用范围，使用空数组`);
            return createDataHandler(sheet, startRow, createEmptyArray(), colCount);
        }

        // 根据实际列数动态计算范围（不再硬编码 A:M）
        const endRow = startRow + usedRange.Rows.Count - 1;
        const endColLetter = colToLetter(colCount);
        const rng = sheet.Range(`A${startRow}:${endColLetter}${endRow}`);

        var arr = rng.Value2;

        // 标准化为二维数组
        if (!Array.isArray(arr)) arr = [[arr]];
        if (!Array.isArray(arr[0])) arr = [arr];

        return createDataHandler(sheet, startRow, arr, colCount);
    } catch (error) {
        console.error(`[mShared_constants] 增强型数组读取失败: ${error.message}`);
        return createDataHandler(sheet, startRow, createEmptyArray(), colCount);
    }
}

/**
 * createDataHandler - 创建数据处理器（内部函数）
 *
 * 设计：脏标记模式
 * - get 操作：直接读取内存数组
 * - set 操作：只更新内存 + 标记脏，不立即写回
 * - syncToSheet()：一次性写回所有脏数据
 *
 * @param {Worksheet} sheet - 工作表对象
 * @param {number} startRow - 起始行号
 * @param {Array} arr - 数据数组
 * @param {number} colCount - 列数
 * @returns {Object} 数据处理器对象
 */
function createDataHandler(sheet, startRow, arr, colCount) {
    var dirty = false;
    const endRow = startRow + arr.length - 1;
    const endColLetter = colToLetter(colCount);

    const getRange = () => sheet.Range(`A${startRow}:${endColLetter}${endRow}`);

    return {
        data: arr,

        /** 读取列（索引从1开始） */
        getColumn(colIndex) {
            return arr.map(function(row) { return row; }[colIndex - 1]);
        },

        /** 修改列（索引从1开始），标记脏 */
        setColumn(colIndex, newData) {
            for (var i = 0; i < arr.length; i++) {
                arr[i][colIndex - 1] = newData[i];
            }
            dirty = true;
        },

        /** 读取行（索引从1开始） */
        getRow(rowIndex) {
            return arr[rowIndex - 1];
        },

        /** 修改行（索引从1开始），标记脏 */
        setRow(rowIndex, newData) {
            arr[rowIndex - 1] = newData;
            dirty = true;
        },

        /** 读取单元格（索引均从1开始） */
        getCell(rowIndex, colIndex) {
            return arr[rowIndex - 1][colIndex - 1];
        },

        /** 修改单元格（索引均从1开始），标记脏 */
        setCell(rowIndex, colIndex, value) {
            arr[rowIndex - 1][colIndex - 1] = value;
            dirty = true;
        },

        /** 获取维度信息 */
        getDimensions() {
            return { rows: arr.length, cols: arr[0] ? arr[0].length : 0 };
        },

        /** 是否有未写回的修改 */
        isDirty() { return dirty; },

        /** 同步到工作表（只在脏时写入） */
        syncToSheet() {
            if (!dirty) return;
            getRange().Value2 = arr;
            dirty = false;
        },

        /** 强制同步（无论是否脏） */
        forceSync() {
            getRange().Value2 = arr;
            dirty = false;
        }
    };
}

// ========== 导出说明 ==========
// WPS JSA 不支持 ES6 的 export 语法
// 所有常量和函数都在全局作用域中，可以直接使用

// ========== 快捷函数 ==========

/**
 * $ - Range快捷访问函数
 * 用法: $("A1") 等同于 Application.ActiveSheet.Range("A1")
 *       $("Sheet1!A1:B10") 等同于 Worksheets("Sheet1").Range("A1:B10")
 * @param {string} addr - 单元格地址，支持 "A1" 或 "Sheet1!A1:B10" 格式
 * @returns {Range} WPS Range对象
 */
function $(addr) {
    if (addr.indexOf('!') > 0) {
        var parts = addr.split('!');
        return Application.Worksheets(parts[0]).Range(parts[1]);
    }
    return Application.ActiveSheet.Range(addr);
}

// ========== 安全备份机制 ==========

/**
 * backupSheetData - 数据修改前自动备份到隐藏工作表
 *
 * 在清除/修改关键数据前调用，创建隐藏备份工作表
 *
 * @param {string} sheetName - 要备份的工作表名称
 * @param {string} rangeAddr - 备份范围，如 "A5:M50"
 * @returns {string} 备份工作表名称，失败返回空字符串
 */
function backupSheetData(sheetName, rangeAddr) {
    try {
        var ws = Application.Worksheets(sheetName);
        var sourceRng = ws.Range(rangeAddr);

        // 生成备份工作表名称
        var timestamp = new Date();
        var tsStr = timestamp.getFullYear().toString() +
            (timestamp.getMonth() + 1).toString().padStart(2, '0') +
            timestamp.getDate().toString().padStart(2, '0') + '_' +
            timestamp.getHours().toString().padStart(2, '0') +
            timestamp.getMinutes().toString().padStart(2, '0') +
            timestamp.getSeconds().toString().padStart(2, '0');
        var backupName = '_bak_' + sheetName + '_' + tsStr;

        // 检查是否已存在同名备份，若存在则先删除
        var prevAlerts = Application.DisplayAlerts;
        Application.DisplayAlerts = false;
        try {
            var existing = Application.Worksheets(backupName);
            existing.Visible = true;  // 先取消隐藏
            existing.Delete();
        } catch (e) {
            // 不存在同名备份，忽略
        }
        Application.DisplayAlerts = prevAlerts;

        // 创建新工作表
        Application.DisplayAlerts = false;
        var backupWs = Application.Worksheets.Add();
        backupWs.Name = backupName;
        Application.DisplayAlerts = prevAlerts;

        // 复制数据
        sourceRng.Copy(backupWs.Range("A1"));

        // 隐藏备份工作表
        backupWs.Visible = 0;  // xlSheetVeryHidden = 0（更隐蔽）

        console.log('[备份] 数据已备份到: ' + backupName + ' (范围: ' + rangeAddr + ')');
        return backupName;
    } catch (error) {
        console.log('[备份] 备份失败: ' + error.message);
        return '';
    }
}

/**
 * $ - Range 快捷函数
 *
 * 提供更简洁的 Range 引用方式，默认使用 ActiveSheet
 *
 * @param {string} addr - 单元格地址，如 "A1:B10" 或 "C5"
 * @param {Object} [sheet] - 可选的工作表对象，默认 ActiveSheet
 * @returns {Range}
 *
 * @example
 * $("A1:B10").Value2 = data;
 * $("C5", mySheet).Interior.Color = COLOR_YELLOW;
 */
function $(addr, sheet) {
    var ws = sheet || Application.ActiveSheet;
    return ws.Range(addr);
}

/**
 * safeFromExcelDate - 安全地将 Excel OA 数值日期转为 JS Date 对象
 *
 * 替代分散在 m直租.js、m银行承兑汇票模块.js 中的 7 处重复转换守卫代码。
 * 自动检测输入类型：Excel 数值 → DateUtils 转换；已是 Date/字符串 → 原样返回。
 *
 * @param {*} rawDate - Excel OA 数值、JS Date、字符串或 undefined
 * @returns {Date|*} 转换后的 JS Date 对象，或原值（如无法识别）
 *
 * @example
 * var d = safeFromExcelDate(Range("A1").Value2);
 * var y = d.getFullYear();  // 安全调用 Date 方法
 */
function safeFromExcelDate(rawDate) {
    if (typeof rawDate !== 'number') return rawDate;
    if (typeof DateUtils !== 'undefined' && typeof DateUtils.fromExcelDate === 'function') {
        return DateUtils.fromExcelDate(rawDate);
    }
    // 兜底：OA 日期基准 1899-12-30
    var excelBase = new Date(1899, 11, 30).getTime();
    return new Date(excelBase + rawDate * 86400000);
}

/**
 * createFormulaTemplate - 创建公式模板二维数组（预分配 rows×cols 空数组）
 *
 * 替代分散在 m直租.js、m调息.js、mFormulaGenerator.js 中的手动 for-loop 初始化。
 *
 * @param {number} rows - 行数（通常 5: header/first/middle/last/total）
 * @param {number} cols - 列数（通常 13 或 14）
 * @returns {Array<Array>} 预分配的二维数组
 *
 * @example
 * var arr = createFormulaTemplate(5, 13);  // 替代 for(var i=0;i<5;i++) arr[i]=new Array(13)
 */
function createFormulaTemplate(rows, cols) {
    var arr = [];
    for (var i = 0; i < rows; i++) arr[i] = new Array(cols);
    return arr;
}

/**
 * 提示消息 - 通过 InputBox 展示消息（内容可选中复制）
 *
 * WPS MsgBox 不支持文字复制，InputBox 的默认值区域可以选中并 Ctrl+C 复制全文。
 * 用于替代 MsgBox/alert 展示需要用户复制的调试信息、错误详情等。
 *
 * @param {string} msg - 提示内容
 * @param {string} title - 对话框标题（可选，默认"系统提示"）
 *
 * @example
 * 提示消息("当前测算总期数: 36\n年利率: 4.5%", "参数确认");
 */
function 提示消息(msg, title) {
    title = title || '系统提示';
    InputBox(msg, title, '请复制上方内容...');
}

console.log(`[mShared_constants] 模块加载完成 - 版本 ${VERSION}`);