/**
 * =======================================================================
 * clsParameterValidator - 参数验证器
 * =======================================================================
 * 
 * 版本: 1.0.0
 * 日期: 2026-02-07
 * 作者: JSA880框架团队
 * 
 * 功能说明:
 *   提供统一的参数验证功能，包括：
 *   - 数据数组验证
 *   - 字段配置验证
 *   - 选项对象验证
 *   - 自定义验证规则
 * 
 * 使用示例:
 *   if (!clsParameterValidator.ValidatePivotParameters(data, rowFields, colFields, dataFields)) {
 *       return [];
 *   }
 * 
 * =======================================================================
 */

/**
 * 参数验证器
 * @type {Object}
 */
var clsParameterValidator = {
    /**
     * 验证透视表参数
     * @param {Array} arr - 数据数组
     * @param {Array|String} rowFields - 行字段配置
     * @param {Array|String} colFields - 列字段配置
     * @param {Array|String} dataFields - 数据字段配置
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidatePivotParameters: function(arr, rowFields, colFields, dataFields) {
        'use strict';
        
        // 验证数据数组
        var dataValidation = this.ValidateDataArray(arr);
        if (!dataValidation.isValid) {
            return dataValidation;
        }
        
        // 验证字段配置（至少需要行字段或列字段之一）
        if (!rowFields && !colFields) {
            return {
                isValid: false,
                message: '至少需要指定行字段或列字段'
            };
        }
        
        // 验证数据字段
        if (!dataFields) {
            return {
                isValid: false,
                message: '必须指定数据字段'
            };
        }
        
        // 验证行字段（如果提供）
        if (rowFields) {
            var rowValidation = this.ValidateFieldConfig(rowFields, 'rowFields');
            if (!rowValidation.isValid) {
                return rowValidation;
            }
        }
        
        // 验证列字段（如果提供）
        if (colFields) {
            var colValidation = this.ValidateFieldConfig(colFields, 'colFields');
            if (!colValidation.isValid) {
                return colValidation;
            }
        }
        
        // 验证数据字段
        var dataValidation2 = this.ValidateFieldConfig(dataFields, 'dataFields');
        if (!dataValidation2.isValid) {
            return dataValidation2;
        }
        
        return {
            isValid: true,
            message: '参数验证通过'
        };
    },
    
    /**
     * 验证数据数组
     * @param {*} arr - 数据数组
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateDataArray: function(arr) {
        'use strict';
        
        if (!arr) {
            return {
                isValid: false,
                message: '数据不能为空'
            };
        }
        
        if (!Array.isArray(arr)) {
            return {
                isValid: false,
                message: '数据必须是数组类型'
            };
        }
        
        if (arr.length === 0) {
            return {
                isValid: false,
                message: '数组不能为空'
            };
        }
        
        return {
            isValid: true,
            message: '数据数组验证通过'
        };
    },
    
    /**
     * 验证字段配置
     * @param {*} fields - 字段配置
     * @param {String} fieldName - 字段名称（用于错误消息）
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateFieldConfig: function(fields, fieldName) {
        'use strict';
        
        fieldName = fieldName || 'fields';
        
        if (!fields) {
            return {
                isValid: false,
                message: fieldName + ' 不能为空'
            };
        }
        
        if (typeof fields !== 'string' && !Array.isArray(fields)) {
            return {
                isValid: false,
                message: fieldName + ' 必须是字符串或数组'
            };
        }
        
        return {
            isValid: true,
            message: fieldName + ' 验证通过'
        };
    },
    
    /**
     * 验证选项对象
     * @param {*} options - 选项对象
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateOptions: function(options) {
        'use strict';
        
        if (options && typeof options !== 'object') {
            return {
                isValid: false,
                message: '选项必须是对象'
            };
        }
        
        return {
            isValid: true,
            message: '选项验证通过'
        };
    },
    
    /**
     * 验证数字参数
     * @param {*} value - 要验证的值
     * @param {String} name - 参数名称
     * @param {Number} minValue - 最小值（可选）
     * @param {Number} maxValue - 最大值（可选）
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateNumber: function(value, name, minValue, maxValue) {
        'use strict';
        
        if (value === null || value === undefined) {
            return {
                isValid: false,
                message: name + ' 不能为空'
            };
        }
        
        if (typeof value !== 'number') {
            return {
                isValid: false,
                message: name + ' 必须是数字'
            };
        }
        
        if (minValue !== undefined && value < minValue) {
            return {
                isValid: false,
                message: name + ' 不能小于 ' + minValue
            };
        }
        
        if (maxValue !== undefined && value > maxValue) {
            return {
                isValid: false,
                message: name + ' 不能大于 ' + maxValue
            };
        }
        
        return {
            isValid: true,
            message: name + ' 验证通过'
        };
    },
    
    /**
     * 验证布尔参数
     * @param {*} value - 要验证的值
     * @param {String} name - 参数名称
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateBoolean: function(value, name) {
        'use strict';
        
        if (value !== null && value !== undefined && typeof value !== 'boolean') {
            return {
                isValid: false,
                message: name + ' 必须是布尔值'
            };
        }
        
        return {
            isValid: true,
            message: name + ' 验证通过'
        };
    },
    
    /**
     * 验证字符串参数
     * @param {*} value - 要验证的值
     * @param {String} name - 参数名称
     * @param {Boolean} isRequired - 是否必需
     * @returns {Object} { isValid: boolean, message: string }
     */
    ValidateString: function(value, name, isRequired) {
        'use strict';
        
        isRequired = (isRequired !== undefined) ? isRequired : true;
        
        if (isRequired && (value === null || value === undefined)) {
            return {
                isValid: false,
                message: name + ' 不能为空'
            };
        }
        
        if (value !== null && value !== undefined && typeof value !== 'string') {
            return {
                isValid: false,
                message: name + ' 必须是字符串'
            };
        }
        
        return {
            isValid: true,
            message: name + ' 验证通过'
        };
    }
};

// 导出为全局对象（如果支持）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = clsParameterValidator;
}
