/**
 * ============== 统一错误处理模块 ==============
 * 作者：徐晓冬
 * 版本：V2.20260130
 * 描述：提供统一的错误处理、日志记录和用户通知机制
 * 
 * 核心改进：
 * - 统一的错误分类和错误码
 * - 分级日志记录（DEBUG/INFO/WARN/ERROR/FATAL）
 * - 可配置的错误处理策略
 * - 错误上报和统计
 * ====================================================
 */

// ============== 错误配置 ==============
const ERROR_CONFIG = {
    // 日志级别
    logLevel: {
        DEBUG: 0,
        INFO: 1,
        WARN: 2,
        ERROR: 3,
        FATAL: 4
    },
    
    // 当前日志级别（生产环境建议设为 WARN）
    currentLogLevel: 1, // INFO
    
    // 是否显示弹窗
    showAlert: true,
    
    // 是否在控制台输出
    logToConsole: true,
    
    // 错误发生时是否继续执行
    continueOnError: false,
    
    // 最大日志条目数
    maxLogEntries: 1000,
    
    // 模块名称
    moduleName: "ErrorHandler"
};

// ============== 错误码定义 ==============
const ERROR_CODES = {
    // 系统级错误 (1000-1999)
    SYSTEM_ERROR: 1000,
    INITIALIZATION_ERROR: 1001,
    CONFIGURATION_ERROR: 1002,
    MODULE_NOT_FOUND: 1003,
    
    // 参数错误 (2000-2999)
    INVALID_PARAMETER: 2000,
    MISSING_REQUIRED_PARAM: 2001,
    PARAM_TYPE_ERROR: 2002,
    PARAM_OUT_OF_RANGE: 2003,
    PARAM_CONVERSION_ERROR: 2004,
    
    // 数据错误 (3000-3999)
    DATA_NOT_FOUND: 3000,
    DATA_FORMAT_ERROR: 3001,
    DATA_VALIDATION_ERROR: 3002,
    DATA_INCONSISTENCY: 3003,
    
    // 计算错误 (4000-4999)
    CALCULATION_ERROR: 4000,
    DIVISION_BY_ZERO: 4001,
    INVALID_FORMULA: 4002,
    CONVERGENCE_ERROR: 4003,
    
    // 文件/IO错误 (5000-5999)
    FILE_NOT_FOUND: 5000,
    FILE_READ_ERROR: 5001,
    FILE_WRITE_ERROR: 5002,
    WORKSHEET_ERROR: 5003,
    
    // 网络错误 (6000-6999)
    NETWORK_ERROR: 6000,
    TIMEOUT_ERROR: 6001,
    API_ERROR: 6002,
    
    // 业务逻辑错误 (7000-7999)
    BUSINESS_RULE_VIOLATION: 7000,
    INVALID_REPAYMENT_METHOD: 7001,
    INVALID_DATE_RANGE: 7002,
    INSUFFICIENT_FUNDS: 7003
};

// ============== 错误信息映射 ==============
const ERROR_MESSAGES = {
    [ERROR_CODES.SYSTEM_ERROR]: "系统错误",
    [ERROR_CODES.INITIALIZATION_ERROR]: "初始化失败",
    [ERROR_CODES.CONFIGURATION_ERROR]: "配置错误",
    [ERROR_CODES.MODULE_NOT_FOUND]: "模块未找到",
    [ERROR_CODES.INVALID_PARAMETER]: "参数无效",
    [ERROR_CODES.MISSING_REQUIRED_PARAM]: "缺少必需参数",
    [ERROR_CODES.PARAM_TYPE_ERROR]: "参数类型错误",
    [ERROR_CODES.PARAM_OUT_OF_RANGE]: "参数超出范围",
    [ERROR_CODES.PARAM_CONVERSION_ERROR]: "参数转换失败",
    [ERROR_CODES.DATA_NOT_FOUND]: "数据未找到",
    [ERROR_CODES.DATA_FORMAT_ERROR]: "数据格式错误",
    [ERROR_CODES.DATA_VALIDATION_ERROR]: "数据验证失败",
    [ERROR_CODES.DATA_INCONSISTENCY]: "数据不一致",
    [ERROR_CODES.CALCULATION_ERROR]: "计算错误",
    [ERROR_CODES.DIVISION_BY_ZERO]: "除零错误",
    [ERROR_CODES.INVALID_FORMULA]: "无效公式",
    [ERROR_CODES.CONVERGENCE_ERROR]: "计算未收敛",
    [ERROR_CODES.FILE_NOT_FOUND]: "文件未找到",
    [ERROR_CODES.FILE_READ_ERROR]: "文件读取失败",
    [ERROR_CODES.FILE_WRITE_ERROR]: "文件写入失败",
    [ERROR_CODES.WORKSHEET_ERROR]: "工作表错误",
    [ERROR_CODES.NETWORK_ERROR]: "网络错误",
    [ERROR_CODES.TIMEOUT_ERROR]: "请求超时",
    [ERROR_CODES.API_ERROR]: "API调用失败",
    [ERROR_CODES.BUSINESS_RULE_VIOLATION]: "违反业务规则",
    [ERROR_CODES.INVALID_REPAYMENT_METHOD]: "无效的还款方式",
    [ERROR_CODES.INVALID_DATE_RANGE]: "无效的日期范围",
    [ERROR_CODES.INSUFFICIENT_FUNDS]: "资金不足"
};

// ============== 统一错误处理类 ==============
class clsErrorHandler {
    constructor(config = {}) {
        this.MODULE_NAME = "clsErrorHandler";
        this.config = { ...ERROR_CONFIG, ...config };
        
        // 日志存储
        this.logEntries = [];
        
        // 错误统计
        this.errorStats = {
            totalErrors: 0,
            errorCountsByCode: {},
            errorCountsByModule: {}
        };
        
        console.log(`[${this.MODULE_NAME}] 错误处理器初始化完成`);
    }
    
    /**
     * 记录日志
     * @param {number} level - 日志级别
     * @param {string} message - 日志消息
     * @param {Object} context - 上下文信息
     */
    log(level, message, context = {}) {
        // 检查日志级别
        if (level < this.config.currentLogLevel) {
            return;
        }
        
        const entry = {
            timestamp: new Date(),
            level: Object.keys(this.config.logLevel).find(function(k) { return this; }.config.logLevel[k] === level),
            message: message,
            context: {
                module: context.module || "Unknown",
                function: context.function || "Unknown",
                ...context
            }
        };
        
        // 存储日志
        this.logEntries.push(entry);
        
        // 限制日志数量
        if (this.logEntries.length > this.config.maxLogEntries) {
            this.logEntries.shift();
        }
        
        // 控制台输出
        if (this.config.logToConsole) {
            const levelName = entry.level;
            const prefix = `[${entry.timestamp.toLocaleTimeString()}] [${levelName}] [${entry.context.module}]`;
            
            switch (level) {
                case ERROR_CONFIG.logLevel.DEBUG:
                    console.log(`${prefix} ${message}`);
                    break;
                case ERROR_CONFIG.logLevel.INFO:
                    console.info(`${prefix} ${message}`);
                    break;
                case ERROR_CONFIG.logLevel.WARN:
                    console.warn(`${prefix} ${message}`);
                    break;
                case ERROR_CONFIG.logLevel.ERROR:
                case ERROR_CONFIG.logLevel.FATAL:
                    console.error(`${prefix} ${message}`);
                    break;
            }
        }
    }
    
    /**
     * 调试日志
     */
    debug(message, context) {
        this.log(ERROR_CONFIG.logLevel.DEBUG, message, context);
    }
    
    /**
     * 信息日志
     */
    info(message, context) {
        this.log(ERROR_CONFIG.logLevel.INFO, message, context);
    }
    
    /**
     * 警告日志
     */
    warn(message, context) {
        this.log(ERROR_CONFIG.logLevel.WARN, message, context);
    }
    
    /**
     * 错误日志
     */
    error(message, context) {
        this.log(ERROR_CONFIG.logLevel.ERROR, message, context);
    }
    
    /**
     * 致命错误日志
     */
    fatal(message, context) {
        this.log(ERROR_CONFIG.logLevel.FATAL, message, context);
    }
    
    /**
     * 处理错误
     * @param {Error} error - 错误对象
     * @param {number} errorCode - 错误码
     * @param {Object} context - 上下文信息
     * @returns {Object} 错误处理结果
     */
    handleError(error, errorCode = ERROR_CODES.SYSTEM_ERROR, context = {}) {
        const moduleName = context.module || "Unknown";
        const functionName = context.function || "Unknown";
        
        // 更新统计
        this.errorStats.totalErrors++;
        this.errorStats.errorCountsByCode[errorCode] = (this.errorStats.errorCountsByCode[errorCode] || 0) + 1;
        const moduleKey = `${moduleName}.${functionName}`;
        this.errorStats.errorCountsByModule[moduleKey] = (this.errorStats.errorCountsByModule[moduleKey] || 0) + 1;
        
        // 构建错误信息
        const errorMessage = ERROR_MESSAGES[errorCode] || "未知错误";
        const fullMessage = `[${moduleName}.${functionName}] ${errorMessage}: ${error.message}`;
        
        // 记录错误
        this.log(ERROR_CONFIG.logLevel.ERROR, fullMessage, {
            ...context,
            errorCode: errorCode,
            stack: error.stack
        });
        
        // 用户通知
        // 注意：errorCode 是业务错误码（1000-7000），不应用于与日志级别比较
        // 根据错误码范围判断严重程度：4000以上为计算/系统级错误，需要弹窗
        const isSevereError = errorCode >= ERROR_CODES.CALCULATION_ERROR || 
                              errorCode === ERROR_CODES.SYSTEM_ERROR ||
                              errorCode === ERROR_CODES.INITIALIZATION_ERROR;
        if (this.config.showAlert && isSevereError) {
            this.notifyUser(errorCode, errorMessage, error.message);
        }
        
        return {
            success: false,
            errorCode: errorCode,
            errorMessage: errorMessage,
            detailMessage: error.message,
            shouldContinue: this.config.continueOnError
        };
    }
    
    /**
     * 验证参数
     * @param {*} value - 参数值
     * @param {string} name - 参数名
     * @param {string} type - 期望类型
     * @param {boolean} required - 是否必需
     * @param {*} defaultValue - 默认值
     * @returns {Object} 验证结果
     */
    validateParam(value, name, type = null, required = true, defaultValue = null) {
        // 检查必需参数
        if (required && (value === null || value === undefined)) {
            return {
                valid: false,
                errorCode: ERROR_CODES.MISSING_REQUIRED_PARAM,
                message: `缺少必需参数：${name}`
            };
        }
        
        // 使用默认值
        if (!required && (value === null || value === undefined)) {
            return {
                valid: true,
                value: defaultValue
            };
        }
        
        // 类型检查
        if (type) {
            const actualType = typeof value;
            if (actualType !== type) {
                return {
                    valid: false,
                    errorCode: ERROR_CODES.PARAM_TYPE_ERROR,
                    message: `参数 ${name} 类型错误，期望 ${type}，实际 ${actualType}`
                };
            }
        }
        
        return {
            valid: true,
            value: value
        };
    }
    
    /**
     * 验证数值范围
     */
    validateRange(value, name, min = null, max = null) {
        if (min !== null && value < min) {
            return {
                valid: false,
                errorCode: ERROR_CODES.PARAM_OUT_OF_RANGE,
                message: `参数 ${name} 不能小于 ${min}`
            };
        }
        
        if (max !== null && value > max) {
            return {
                valid: false,
                errorCode: ERROR_CODES.PARAM_OUT_OF_RANGE,
                message: `参数 ${name} 不能大于 ${max}`
            };
        }
        
        return { valid: true, value: value };
    }
    
    /**
     * 用户通知
     */
    notifyUser(errorCode, errorMessage, detailMessage) {
        try {
            if (typeof alert !== 'undefined') {
                console.error(`[${errorCode}] ${errorMessage}\n详细信息：${detailMessage}`);
            }
        } catch (e) {
            console.error("显示错误弹窗失败:", e);
        }
    }
    
    /**
     * 获取日志
     */
    getLogs(level = null, module = null, startTime = null, endTime = null) {
        var logs = this.logEntries;
        
        if (level !== null) {
            var levelName = Object.keys(this.config.logLevel).find(function(k) { return this.config.logLevel[k] === level; });
            logs = logs.filter(function(entry) { return entry.level === levelName; });
        }
        
        if (module !== null) {
            logs = logs.filter(function(entry) { return entry.context.module === module; });
        }
        
        if (startTime !== null) {
            logs = logs.filter(function(entry) { return entry.timestamp >= startTime; });
        }
        
        if (endTime !== null) {
            logs = logs.filter(function(entry) { return entry.timestamp <= endTime; });
        }
        
        return logs;
    }
    
    /**
     * 获取错误统计
     */
    getErrorStats() {
        return {
            ...this.errorStats,
            topErrors: Object.entries(this.errorStats.errorCountsByCode)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10),
            topModules: Object.entries(this.errorStats.errorCountsByModule)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 10)
        };
    }
    
    /**
     * 清空日志
     */
    clearLogs() {
        this.logEntries = [];
        console.log(`[${this.MODULE_NAME}] 日志已清空`);
    }
    
    /**
     * 重置统计
     */
    resetStats() {
        this.errorStats = {
            totalErrors: 0,
            errorCountsByCode: {},
            errorCountsByModule: {}
        };
        console.log(`[${this.MODULE_NAME}] 统计已重置`);
    }
}

// ============== 全局错误处理器实例 ==============
// 创建默认实例
const g_errorHandler = new clsErrorHandler();

// ============== 便捷函数 ==============

/**
 * 快速记录错误
 */
function logError(message, module = "Unknown", func = "Unknown") {
    g_errorHandler.error(message, { module, function: func });
}

/**
 * 快速处理异常
 */
function handleException(error, errorCode, module = "Unknown", func = "Unknown") {
    return g_errorHandler.handleError(error, errorCode, { module, function: func });
}

/**
 * 包装函数：自动捕获异常
 */
function tryCatch(wrapper, errorCode = ERROR_CODES.SYSTEM_ERROR, context = {}) {
    return function(...args) {
        try {
            return wrapper.apply(this, args);
        } catch (error) {
            return g_errorHandler.handleError(error, errorCode, context);
        }
    };
}

console.log("[mErrorHandler.js] 统一错误处理模块加载完成");
