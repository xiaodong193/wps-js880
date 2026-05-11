/**
 * =======================================================================
 * clsErrorHandler - 错误处理管理器
 * =======================================================================
 * 
 * 版本: 1.0.0
 * 日期: 2026-02-07
 * 作者: JSA880框架团队
 * 
 * 功能说明:
 *   提供统一的错误处理机制，包括：
 *   - 安全执行函数（捕获异常）
 *   - 参数验证
 *   - 错误日志记录
 *   - 友好的错误提示
 * 
 * 使用示例:
 *   var result = clsErrorHandler.SafeExecute(
 *       function() { return riskyOperation(); },
 *       'riskyOperation',
 *       defaultValue
 *   );
 * 
 * =======================================================================
 */

/**
 * 错误处理管理器
 * @type {Object}
 */
var clsErrorHandler = (function() {
    'use strict';
    
    // 私有变量
    var m_errorCount = 0;
    var m_maxErrors = 100;
    var m_errorLog = [];
    
    /**
     * 格式化错误消息
     * @private
     * @param {String} context - 错误上下文
     * @param {Error} error - 错误对象
     * @returns {String} 格式化的错误消息
     */
    function formatErrorMessage(context, error) {
        var message = '[ERROR] ' + context;
        
        if (error) {
            if (error.message) {
                message += ': ' + error.message;
            }
            if (error.stack) {
                message += '\nStack: ' + error.stack;
            }
        }
        
        return message;
    }
    
    /**
     * 记录错误
     * @private
     * @param {String} context - 错误上下文
     * @param {Error} error - 错误对象
     */
    function logError(context, error) {
        var message = formatErrorMessage(context, error);
        
        // 输出到控制台
        Console.log(message);
        
        // 记录到日志
        m_errorLog.push({
            timestamp: new Date().toISOString(),
            context: context,
            message: error ? error.message : 'Unknown error',
            stack: error ? error.stack : null
        });
        
        m_errorCount++;
        
        // 限制日志大小
        if (m_errorLog.length > m_maxErrors) {
            m_errorLog.shift();
        }
    }
    
    // 公共接口
    return {
        /**
         * 安全执行函数
         * @param {Function} fn - 要执行的函数
         * @param {String} context - 错误上下文（用于日志）
         * @param {*} defaultValue - 发生错误时返回的默认值
         * @returns {*} 函数执行结果或默认值
         */
        SafeExecute: function(fn, context, defaultValue) {
            try {
                if (typeof fn !== 'function') {
                    throw new Error('第一个参数必须是函数');
                }
                
                return fn();
            } catch (error) {
                logError(context, error);
                return defaultValue;
            }
        },
        
        /**
         * 验证参数
         * @param {*} value - 要验证的值
         * @param {String} name - 参数名
         * @param {String} expectedType - 期望类型（'array', 'object', 'string', 'number' 等）
         * @param {Boolean} isRequired - 是否必需（默认 true）
         * @returns {Boolean} 是否有效
         */
        ValidateParam: function(value, name, expectedType, isRequired) {
            isRequired = (isRequired !== undefined) ? isRequired : true;
            
            // 检查必需参数
            if (isRequired && (value === null || value === undefined)) {
                Console.log('[ERROR] 参数 ' + name + ' 是必需的，但收到 ' + value);
                return false;
            }
            
            // 如果非必需且为空，则通过
            if (!isRequired && (value === null || value === undefined)) {
                return true;
            }
            
            // 检查类型
            var actualType = typeof value;
            
            if (expectedType === 'array') {
                if (!Array.isArray(value)) {
                    Console.log('[ERROR] 参数 ' + name + ' 必须是数组，但收到 ' + actualType);
                    return false;
                }
            } else if (expectedType === 'object') {
                if (actualType !== 'object' || Array.isArray(value)) {
                    Console.log('[ERROR] 参数 ' + name + ' 必须是对象，但收到 ' + actualType);
                    return false;
                }
            } else if (actualType !== expectedType) {
                Console.log('[ERROR] 参数 ' + name + ' 必须是 ' + expectedType + '，但收到 ' + actualType);
                return false;
            }
            
            return true;
        },
        
        /**
         * 验证数据数组
         * @param {*} arr - 数据数组
         * @returns {Boolean} 是否有效
         */
        ValidateDataArray: function(arr) {
            if (!arr) {
                Console.log('[ERROR] 数据不能为空');
                return false;
            }
            
            if (!Array.isArray(arr)) {
                Console.log('[ERROR] 数据必须是数组');
                return false;
            }
            
            if (arr.length === 0) {
                Console.log('[ERROR] 数组不能为空');
                return false;
            }
            
            return true;
        },
        
        /**
         * 验证字段配置
         * @param {*} fields - 字段配置
         * @returns {Boolean} 是否有效
         */
        ValidateFieldConfig: function(fields) {
            if (!fields) {
                Console.log('[ERROR] 字段配置不能为空');
                return false;
            }
            
            if (typeof fields !== 'string' && !Array.isArray(fields)) {
                Console.log('[ERROR] 字段配置必须是字符串或数组');
                return false;
            }
            
            return true;
        },
        
        /**
         * 抛出错误
         * @param {String} message - 错误消息
         */
        ThrowError: function(message) {
            throw new Error(message);
        },
        
        /**
         * 获取错误日志
         * @returns {Array} 错误日志数组
         */
        GetErrorLog: function() {
            return m_errorLog.slice();  // 返回副本
        },
        
        /**
         * 清空错误日志
         */
        ClearErrorLog: function() {
            m_errorLog = [];
            m_errorCount = 0;
        },
        
        /**
         * 获取错误计数
         * @returns {Number} 错误总数
         */
        GetErrorCount: function() {
            return m_errorCount;
        }
    };
})();

// 导出为全局对象（如果支持）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = clsErrorHandler;
}
