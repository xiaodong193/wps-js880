/**
 * =======================================================================
 * clsLogger - 日志管理器
 * =======================================================================
 * 
 * 版本: 1.0.0
 * 日期: 2026-02-07
 * 作者: JSA880框架团队
 * 
 * 功能说明:
 *   提供统一的日志管理功能，包括：
 *   - 多级别日志（DEBUG, INFO, WARN, ERROR）
 *   - 日志开关控制
 *   - 日志缓冲
 *   - 日志导出
 * 
 * 使用示例:
 *   Logger.SetEnabled(true);
 *   Logger.SetLogLevel('DEBUG');
 *   Logger.Debug('调试信息');
 *   Logger.Info('普通信息');
 *   Logger.Warn('警告信息');
 *   Logger.Error('错误信息');
 * 
 * =======================================================================
 */

/**
 * 日志管理器
 * @type {Object}
 */
var clsLogger = (function() {
    'use strict';
    
    // 私有变量
    var m_isEnabled = false;  // 默认关闭
    var m_logLevel = 'INFO';  // DEBUG, INFO, WARN, ERROR
    var m_logBuffer = [];
    var m_maxBufferSize = 1000;
    
    // 日志级别定义
    var LOG_LEVELS = {
        'DEBUG': 0,
        'INFO': 1,
        'WARN': 2,
        'ERROR': 3
    };
    
    /**
     * 获取当前日志级别数值
     * @private
     * @returns {Number} 日志级别数值
     */
    function getCurrentLogLevel() {
        return LOG_LEVELS[m_logLevel] || LOG_LEVELS.INFO;
    }
    
    /**
     * 格式化时间戳
     * @private
     * @returns {String} 格式化的时间戳
     */
    function formatTimestamp() {
        var now = new Date();
        var hours = String(now.getHours()).padStart(2, '0');
        var minutes = String(now.getMinutes()).padStart(2, '0');
        var seconds = String(now.getSeconds()).padStart(2, '0');
        var milliseconds = String(now.getMilliseconds()).padStart(3, '0');
        return hours + ':' + minutes + ':' + seconds + '.' + milliseconds;
    }
    
    /**
     * 输出日志
     * @private
     * @param {String} level - 日志级别
     * @param {String} message - 日志消息
     */
    function log(level, message) {
        var currentLevel = getCurrentLogLevel();
        var msgLevel = LOG_LEVELS[level];
        
        // 检查日志级别
        if (msgLevel < currentLevel) {
            return;  // 级别不够，不输出
        }
        
        // 检查开关
        if (!m_isEnabled && level !== 'ERROR') {
            return;  // 关闭状态下只输出错误
        }
        
        // 格式化日志
        var timestamp = formatTimestamp();
        var formattedMessage = '[' + timestamp + '] [' + level + '] ' + message;
        
        // 输出到控制台
        Console.log(formattedMessage);
        
        // 添加到缓冲区
        m_logBuffer.push({
            timestamp: timestamp,
            level: level,
            message: message
        });
        
        // 限制缓冲区大小
        if (m_logBuffer.length > m_maxBufferSize) {
            m_logBuffer.shift();
        }
    }
    
    // 公共接口
    return {
        /**
         * 设置日志开关
         * @param {Boolean} enabled - 是否启用日志
         */
        SetEnabled: function(enabled) {
            m_isEnabled = enabled;
        },
        
        /**
         * 获取日志开关状态
         * @returns {Boolean} 是否启用日志
         */
        IsEnabled: function() {
            return m_isEnabled;
        },
        
        /**
         * 设置日志级别
         * @param {String} level - 日志级别（'DEBUG', 'INFO', 'WARN', 'ERROR'）
         */
        SetLogLevel: function(level) {
            if (LOG_LEVELS[level] !== undefined) {
                m_logLevel = level;
            } else {
                Console.log('[WARN] 无效的日志级别: ' + level);
            }
        },
        
        /**
         * 获取日志级别
         * @returns {String} 当前日志级别
         */
        GetLogLevel: function() {
            return m_logLevel;
        },
        
        /**
         * 输出调试信息
         * @param {String} message - 日志消息
         */
        Debug: function(message) {
            log('DEBUG', message);
        },
        
        /**
         * 输出普通信息
         * @param {String} message - 日志消息
         */
        Info: function(message) {
            log('INFO', message);
        },
        
        /**
         * 输出警告信息
         * @param {String} message - 日志消息
         */
        Warn: function(message) {
            log('WARN', message);
        },
        
        /**
         * 输出错误信息
         * @param {String} message - 日志消息
         */
        Error: function(message) {
            log('ERROR', message);
        },
        
        /**
         * 获取日志缓冲区
         * @returns {Array} 日志缓冲区副本
         */
        GetLogBuffer: function() {
            return m_logBuffer.slice();
        },
        
        /**
         * 清空日志缓冲区
         */
        ClearBuffer: function() {
            m_logBuffer = [];
        },
        
        /**
         * 导出日志为字符串
         * @returns {String} 日志字符串
         */
        ExportToString: function() {
            var lines = [];
            for (var i = 0; i < m_logBuffer.length; i++) {
                var entry = m_logBuffer[i];
                lines.push('[' + entry.timestamp + '] [' + entry.level + '] ' + entry.message);
            }
            return lines.join('\n');
        },
        
        /**
         * 导出日志到数组（用于写入工作表）
         * @returns {Array} 二维数组 [timestamp, level, message]
         */
        ExportToArray: function() {
            var result = [];
            for (var i = 0; i < m_logBuffer.length; i++) {
                var entry = m_logBuffer[i];
                result.push([entry.timestamp, entry.level, entry.message]);
            }
            return result;
        }
    };
})();

// 全局快捷方式
var Logger = clsLogger;

// 导出为全局对象（如果支持）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = clsLogger;
}
