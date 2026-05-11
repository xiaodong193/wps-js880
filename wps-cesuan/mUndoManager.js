/**
 * ============== 统一撤销管理器模块 ==============
 * 作者：徐晓冬
 * 版本：V2.20260130
 * 描述：为整个租金测算系统提供统一的撤销/重做功能
 * 
 * 核心特性：
 * - 命令模式（Command Pattern）封装操作
 * - 支持任意操作的撤销/重做
 * - 多层级历史记录管理
 * - 与现有代码无缝集成
 * - 支持操作分组（事务）
 * - 支持操作描述和元数据
 * ====================================================
 */

// ============== 撤销管理器配置常量 ==============
const UNDO_CONFIG = {
    // 默认最大历史记录数
    DEFAULT_MAX_HISTORY: 20,
    // 最大允许的历史记录数
    ABSOLUTE_MAX_HISTORY: 100,
    // 默认操作分组名称
    DEFAULT_GROUP_NAME: "default",
    // 操作类型枚举
    OPERATION_TYPES: {
        DATA_MODIFY: "data_modify",         // 数据修改
        ROW_INSERT: "row_insert",           // 插入行
        ROW_DELETE: "row_delete",           // 删除行
        PARAM_CHANGE: "param_change",       // 参数变更
        CONFIG_CHANGE: "config_change",     // 配置变更
        TABLE_GENERATE: "table_generate",   // 生成表格
        RATE_ADJUST: "rate_adjust",         // 利率调整
        STYLE_APPLY: "style_apply",         // 样式应用
        BATCH_OPERATION: "batch_operation"  // 批量操作
    }
};

// ============== 操作命令基类 ==============
class clsCommand {
    /**
     * 构造函数
     * @param {string} type - 操作类型
     * @param {string} description - 操作描述
     * @param {Object} metadata - 操作元数据
     */
    constructor(type, description, metadata) {
        this.id = this.generateId();
        this.type = type;
        this.description = description;
        this.metadata = metadata || {};
        this.timestamp = new Date();
        this.executed = false;
    }
    
    /**
     * 生成唯一ID
     * @returns {string} 唯一ID
     */
    generateId() {
        return "cmd_" + Date.now() + "_" + Math.floor(Math.random() * 10000);
    }
    
    /**
     * 执行操作（子类必须实现）
     * @returns {boolean} 是否成功
     */
    execute() {
        throw new Error("execute() 方法必须在子类中实现");
    }
    
    /**
     * 撤销操作（子类必须实现）
     * @returns {boolean} 是否成功
     */
    undo() {
        throw new Error("undo() 方法必须在子类中实现");
    }
    
    /**
     * 重做操作（默认调用execute，子类可覆盖）
     * @returns {boolean} 是否成功
     */
    redo() {
        return this.execute();
    }
    
    /**
     * 获取操作信息
     * @returns {Object} 操作信息
     */
    getInfo() {
        return {
            id: this.id,
            type: this.type,
            description: this.description,
            timestamp: this.timestamp,
            executed: this.executed,
            metadata: this.metadata
        };
    }
}

// ============== 数据修改命令类 ==============
class clsDataModifyCommand extends clsCommand {
    /**
     * 构造函数
     * @param {Object} target - 操作目标对象
     * @param {string} property - 属性名
     * @param {*} oldValue - 旧值
     * @param {*} newValue - 新值
     * @param {string} description - 操作描述
     */
    constructor(target, property, oldValue, newValue, description) {
        super(UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY, description);
        this.target = target;
        this.property = property;
        this.oldValue = oldValue;
        this.newValue = newValue;
    }
    
    /**
     * 执行操作
     * @returns {boolean} 是否成功
     */
    execute() {
        try {
            this.target[this.property] = this.newValue;
            this.executed = true;
            return true;
        } catch (error) {
            console.error(`[clsDataModifyCommand] 执行失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * 撤销操作
     * @returns {boolean} 是否成功
     */
    undo() {
        try {
            this.target[this.property] = this.oldValue;
            return true;
        } catch (error) {
            console.error(`[clsDataModifyCommand] 撤销失败: ${error.message}`);
            return false;
        }
    }
}

// ============== 工作表数据修改命令类 ==============
class clsWorksheetModifyCommand extends clsCommand {
    /**
     * 构造函数
     * @param {Object} worksheet - 工作表对象
     * @param {Object} range - 数据范围
     * @param {*} oldData - 旧数据
     * @param {*} newData - 新数据
     * @param {string} description - 操作描述
     */
    constructor(worksheet, range, oldData, newData, description) {
        super(UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY, description);
        this.worksheet = worksheet;
        this.range = range;
        this.oldData = oldData;
        this.newData = newData;
    }
    
    /**
     * 执行操作
     * @returns {boolean} 是否成功
     */
    execute() {
        try {
            const targetRange = this.worksheet.Range(this.range);
            targetRange.Value2 = this.newData;
            this.executed = true;
            return true;
        } catch (error) {
            console.error(`[clsWorksheetModifyCommand] 执行失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * 撤销操作
     * @returns {boolean} 是否成功
     */
    undo() {
        try {
            const targetRange = this.worksheet.Range(this.range);
            targetRange.Value2 = this.oldData;
            return true;
        } catch (error) {
            console.error(`[clsWorksheetModifyCommand] 撤销失败: ${error.message}`);
            return false;
        }
    }
}

// ============== 批量操作命令类（宏命令） ==============
class clsBatchCommand extends clsCommand {
    /**
     * 构造函数
     * @param {string} description - 操作描述
     * @param {Array} commands - 子命令数组
     */
    constructor(description, commands) {
        super(UNDO_CONFIG.OPERATION_TYPES.BATCH_OPERATION, description);
        this.commands = commands || [];
        this.successCount = 0;
    }
    
    /**
     * 添加子命令
     * @param {clsCommand} command - 子命令
     */
    addCommand(command) {
        this.commands.push(command);
    }
    
    /**
     * 执行所有子命令
     * @returns {boolean} 是否全部成功
     */
    execute() {
        this.successCount = 0;
        
        for (var i = 0; i < this.commands.length; i++) {
            const cmd = this.commands[i];
            if (cmd.execute()) {
                this.successCount++;
            } else {
                // 执行失败，回滚已执行的命令
                this.rollback(i - 1);
                return false;
            }
        }
        
        this.executed = true;
        return this.successCount === this.commands.length;
    }
    
    /**
     * 撤销所有子命令（逆序）
     * @returns {boolean} 是否全部成功
     */
    undo() {
        var success = true;
        
        // 逆序撤销
        for (var i = this.commands.length - 1; i >= 0; i--) {
            const cmd = this.commands[i];
            if (!cmd.undo()) {
                success = false;
            }
        }
        
        return success;
    }
    
    /**
     * 回滚已执行的命令
     * @param {number} lastExecutedIndex - 最后成功执行的命令索引
     */
    rollback(lastExecutedIndex) {
        for (var i = lastExecutedIndex; i >= 0; i--) {
            this.commands[i].undo();
        }
    }
}

// ============== 撤销管理器类 ==============
class clsUndoManager {
    /**
     * 构造函数
     * @param {Object} options - 配置选项
     */
    constructor(options) {
        this.MODULE_NAME = "clsUndoManager";
        this.VERSION = "2.20260130";
        
        // 配置选项
        this.m_maxHistorySize = (options && options.maxHistory) || UNDO_CONFIG.DEFAULT_MAX_HISTORY;
        this.m_enableGrouping = (options && options.enableGrouping) || false;
        this.m_currentGroup = null;
        
        // 历史记录栈
        this.m_undoStack = [];
        this.m_redoStack = [];
        
        // 监听器
        this.m_listeners = {
            onUndo: [],
            onRedo: [],
            onChange: []
        };
        
        console.log(`[${this.MODULE_NAME}] 撤销管理器初始化完成 - 版本 ${this.VERSION}`);
    }
    
    /**
     * 执行命令（自动记录历史）
     * @param {clsCommand} command - 命令对象
     * @returns {boolean} 是否成功
     */
    execute(command) {
        try {
            console.log(`[${this.MODULE_NAME}] 执行操作: ${command.description}`);
            
            // 执行命令
            const result = command.execute();
            
            if (result) {
                // 清空重做栈（新操作后不能重做）
                this.m_redoStack = [];
                
                // 添加到撤销栈
                if (this.m_currentGroup && this.m_enableGrouping) {
                    // 添加到当前分组
                    this.m_currentGroup.addCommand(command);
                } else {
                    this.m_undoStack.push(command);
                    this.trimHistory();
                }
                
                // 触发变更事件
                this.notifyChange("execute", command);
                
                console.log(`[${this.MODULE_NAME}] 操作执行成功`);
            } else {
                console.error(`[${this.MODULE_NAME}] 操作执行失败`);
            }
            
            return result;
        } catch (error) {
            console.error(`[${this.MODULE_NAME}] 执行操作失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * 撤销上一步操作
     * @returns {Object} 撤销结果
     */
    undo() {
        try {
            if (this.m_undoStack.length === 0) {
                console.log(`[${this.MODULE_NAME}] 没有可撤销的操作`);
                return { success: false, message: "没有可撤销的操作" };
            }
            
            console.log(`[${this.MODULE_NAME}] ========== 开始撤销操作 ==========`);
            
            // 获取最后一个命令
            const command = this.m_undoStack.pop();
            
            // 执行撤销
            const result = command.undo();
            
            if (result) {
                // 添加到重做栈
                this.m_redoStack.push(command);
                
                // 触发撤销事件
                this.notifyChange("undo", command);
                this.triggerListener("onUndo", command);
                
                console.log(`[${this.MODULE_NAME}] 撤销成功: ${command.description}`);
                console.log(`[${this.MODULE_NAME}] ========== 撤销操作完成 ==========`);
                
                return {
                    success: true,
                    message: `已撤销: ${command.description}`,
                    command: command.getInfo()
                };
            } else {
                // 撤销失败，恢复原状
                this.m_undoStack.push(command);
                console.error(`[${this.MODULE_NAME}] 撤销失败`);
                return { success: false, message: "撤销失败" };
            }
        } catch (error) {
            console.error(`[${this.MODULE_NAME}] 撤销操作异常: ${error.message}`);
            return { success: false, message: `撤销失败: ${error.message}` };
        }
    }
    
    /**
     * 重做上一步撤销的操作
     * @returns {Object} 重做结果
     */
    redo() {
        try {
            if (this.m_redoStack.length === 0) {
                console.log(`[${this.MODULE_NAME}] 没有可重做的操作`);
                return { success: false, message: "没有可重做的操作" };
            }
            
            console.log(`[${this.MODULE_NAME}] ========== 开始重做操作 ==========`);
            
            // 获取最后一个重做命令
            const command = this.m_redoStack.pop();
            
            // 执行重做
            const result = command.redo();
            
            if (result) {
                // 添加回复原栈
                this.m_undoStack.push(command);
                this.trimHistory();
                
                // 触发重做事件
                this.notifyChange("redo", command);
                this.triggerListener("onRedo", command);
                
                console.log(`[${this.MODULE_NAME}] 重做成功: ${command.description}`);
                console.log(`[${this.MODULE_NAME}] ========== 重做操作完成 ==========`);
                
                return {
                    success: true,
                    message: `已重做: ${command.description}`,
                    command: command.getInfo()
                };
            } else {
                // 重做失败，恢复原状
                this.m_redoStack.push(command);
                console.error(`[${this.MODULE_NAME}] 重做失败`);
                return { success: false, message: "重做失败" };
            }
        } catch (error) {
            console.error(`[${this.MODULE_NAME}] 重做操作异常: ${error.message}`);
            return { success: false, message: `重做失败: ${error.message}` };
        }
    }
    
    /**
     * 开始操作分组（事务）
     * @param {string} description - 分组描述
     */
    beginGroup(description) {
        if (!this.m_enableGrouping) {
            console.warn(`[${this.MODULE_NAME}] 分组功能未启用`);
            return;
        }
        
        if (this.m_currentGroup) {
            console.warn(`[${this.MODULE_NAME}] 已有进行中的分组，自动结束并创建新分组`);
            this.endGroup();
        }
        
        this.m_currentGroup = new clsBatchCommand(description);
        console.log(`[${this.MODULE_NAME}] 开始操作分组: ${description}`);
    }
    
    /**
     * 结束操作分组（事务）
     */
    endGroup() {
        if (!this.m_enableGrouping) {
            return;
        }
        
        if (!this.m_currentGroup) {
            console.warn(`[${this.MODULE_NAME}] 没有进行中的分组`);
            return;
        }
        
        // 如果有子命令，添加到撤销栈
        if (this.m_currentGroup.commands.length > 0) {
            this.m_undoStack.push(this.m_currentGroup);
            this.trimHistory();
            this.m_redoStack = [];
            
            console.log(`[${this.MODULE_NAME}] 结束操作分组: ${this.m_currentGroup.description}，包含 ${this.m_currentGroup.commands.length} 个子操作`);
        } else {
            console.log(`[${this.MODULE_NAME}] 操作分组为空，已丢弃`);
        }
        
        this.m_currentGroup = null;
    }
    
    /**
     * 取消操作分组（不保存）
     */
    cancelGroup() {
        if (!this.m_currentGroup) {
            return;
        }
        
        console.log(`[${this.MODULE_NAME}] 取消操作分组: ${this.m_currentGroup.description}`);
        this.m_currentGroup = null;
    }
    
    /**
     * 检查是否可以撤销
     * @returns {boolean} 是否可以撤销
     */
    canUndo() {
        return this.m_undoStack.length > 0;
    }
    
    /**
     * 检查是否可以重做
     * @returns {boolean} 是否可以重做
     */
    canRedo() {
        return this.m_redoStack.length > 0;
    }
    
    /**
     * 获取撤销历史
     * @returns {Array} 历史记录数组
     */
    getUndoHistory() {
        return this.m_undoStack.map(function(cmd, index) {
            return {
                index: index,
                info: cmd.getInfo()
            };
        });
    }
    
    /**
     * 获取重做历史
     * @returns {Array} 重做历史数组
     */
    getRedoHistory() {
        return this.m_redoStack.map(function(cmd, index) {
            return {
                index: index,
                info: cmd.getInfo()
            };
        });
    }
    
    /**
     * 获取下一个可撤销操作的信息
     * @returns {Object|null} 操作信息
     */
    getNextUndoInfo() {
        if (this.m_undoStack.length === 0) {
            return null;
        }
        return this.m_undoStack[this.m_undoStack.length - 1].getInfo();
    }
    
    /**
     * 获取下一个可重做操作的信息
     * @returns {Object|null} 操作信息
     */
    getNextRedoInfo() {
        if (this.m_redoStack.length === 0) {
            return null;
        }
        return this.m_redoStack[this.m_redoStack.length - 1].getInfo();
    }
    
    /**
     * 清空所有历史记录
     */
    clear() {
        this.m_undoStack = [];
        this.m_redoStack = [];
        this.m_currentGroup = null;
        console.log(`[${this.MODULE_NAME}] 已清空所有历史记录`);
    }
    
    /**
     * 设置最大历史记录数
     * @param {number} size - 历史记录数
     */
    setMaxHistorySize(size) {
        if (typeof size === "number" && size > 0 && size <= UNDO_CONFIG.ABSOLUTE_MAX_HISTORY) {
            this.m_maxHistorySize = size;
            this.trimHistory();
            console.log(`[${this.MODULE_NAME}] 最大历史记录数已设置为: ${size}`);
        } else {
            console.warn(`[${this.MODULE_NAME}] 无效的历史记录数: ${size}，应在 1-${UNDO_CONFIG.ABSOLUTE_MAX_HISTORY} 之间`);
        }
    }
    
    /**
     * 启用/禁用分组功能
     * @param {boolean} enabled - 是否启用
     */
    setGroupingEnabled(enabled) {
        this.m_enableGrouping = enabled;
        console.log(`[${this.MODULE_NAME}] 分组功能已${enabled ? "启用" : "禁用"}`);
    }
    
    /**
     * 修剪历史记录（保持不超过最大值）
     * @private
     */
    trimHistory() {
        while (this.m_undoStack.length > this.m_maxHistorySize) {
            this.m_undoStack.shift();
        }
    }
    
    /**
     * 添加事件监听器
     * @param {string} event - 事件名称（onUndo, onRedo, onChange）
     * @param {Function} callback - 回调函数
     */
    addListener(event, callback) {
        if (this.m_listeners[event]) {
            this.m_listeners[event].push(callback);
        }
    }
    
    /**
     * 移除事件监听器
     * @param {string} event - 事件名称
     * @param {Function} callback - 回调函数
     */
    removeListener(event, callback) {
        if (this.m_listeners[event]) {
            const index = this.m_listeners[event].indexOf(callback);
            if (index > -1) {
                this.m_listeners[event].splice(index, 1);
            }
        }
    }
    
    /**
     * 触发监听器
     * @private
     * @param {string} event - 事件名称
     * @param {*} data - 事件数据
     */
    triggerListener(event, data) {
        if (this.m_listeners[event]) {
            const moduleName = this.MODULE_NAME;
            this.m_listeners[event].forEach(function(callback) {
                try {
                    callback(data);
                } catch (error) {
                    console.error(`[${moduleName}] 监听器执行失败: ${error.message}`);
                }
            });
        }
    }
    
    /**
     * 通知状态变更
     * @private
     * @param {string} action - 操作类型
     * @param {clsCommand} command - 命令对象
     */
    notifyChange(action, command) {
        this.triggerListener("onChange", {
            action: action,
            command: command.getInfo(),
            canUndo: this.canUndo(),
            canRedo: this.canRedo(),
            undoCount: this.m_undoStack.length,
            redoCount: this.m_redoStack.length
        });
    }
    
    /**
     * 获取状态信息
     * @returns {Object} 状态信息
     */
    getStatus() {
        return {
            canUndo: this.canUndo(),
            canRedo: this.canRedo(),
            undoCount: this.m_undoStack.length,
            redoCount: this.m_redoStack.length,
            maxHistory: this.m_maxHistorySize,
            groupingEnabled: this.m_enableGrouping,
            inGroup: this.m_currentGroup !== null
        };
    }
}

// ============== 通用可撤销命令类 ==============
/**
 * clsUndoableCommand - 通用可撤销命令类
 * 
 * 作用：将任意的执行/撤销函数对封装为标准命令对象
 * 设计：适配器模式，将函数适配为 clsCommand 接口
 */
class clsUndoableCommand extends clsCommand {
    /**
     * 构造函数
     * @param {Object} context - 执行上下文（this指向）
     * @param {string} type - 操作类型
     * @param {string} description - 操作描述
     * @param {Function} executeFn - 执行函数
     * @param {Function} undoFn - 撤销函数
     * @param {Object} metadata - 元数据
     */
    constructor(context, type, description, executeFn, undoFn, metadata) {
        super(type, description, metadata);
        this._context = context;
        this._executeFn = executeFn;
        this._undoFn = undoFn;
    }
    
    /**
     * 执行操作
     * @returns {boolean} 是否成功
     */
    execute() {
        try {
            const result = this._executeFn.call(this._context);
            this.executed = true;
            return result;
        } catch (error) {
            console.error(`[clsUndoableCommand] 执行失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * 撤销操作
     * @returns {boolean} 是否成功
     */
    undo() {
        try {
            return this._undoFn.call(this._context);
        } catch (error) {
            console.error(`[clsUndoableCommand] 撤销失败: ${error.message}`);
            return false;
        }
    }
}

// ============== 全局撤销管理器实例 ==============
const g_undoManager = new clsUndoManager();

// ============== 便捷函数 ==============

/**
 * 获取全局撤销管理器实例
 * @returns {clsUndoManager} 撤销管理器实例
 */
function getUndoManager() {
    return g_undoManager;
}

/**
 * 创建新的撤销管理器实例
 * @param {Object} options - 配置选项
 * @returns {clsUndoManager} 新的撤销管理器实例
 */
function createUndoManager(options) {
    return new clsUndoManager(options);
}

/**
 * 撤销上一步操作（使用全局管理器）
 * @returns {Object} 撤销结果
 */
function globalUndo() {
    return g_undoManager.undo();
}

/**
 * 重做上一步撤销的操作（使用全局管理器）
 * @returns {Object} 重做结果
 */
function globalRedo() {
    return g_undoManager.redo();
}

/**
 * 检查是否可以撤销（使用全局管理器）
 * @returns {boolean} 是否可以撤销
 */
function canGlobalUndo() {
    return g_undoManager.canUndo();
}

/**
 * 检查是否可以重做（使用全局管理器）
 * @returns {boolean} 是否可以重做
 */
function canGlobalRedo() {
    return g_undoManager.canRedo();
}

/**
 * 显示撤销/重做状态
 */
function showUndoStatus() {
    const status = g_undoManager.getStatus();
    console.log("========== 撤销管理器状态 ==========");
    console.log(`可撤销: ${status.canUndo ? "是" : "否"} (${status.undoCount} 条记录)`);
    console.log(`可重做: ${status.canRedo ? "是" : "否"} (${status.redoCount} 条记录)`);
    console.log(`最大历史: ${status.maxHistory}`);
    console.log(`分组功能: ${status.groupingEnabled ? "启用" : "禁用"}`);
    console.log(`正在进行分组: ${status.inGroup ? "是" : "否"}`);
    console.log("====================================");
}

// ============== 集成示例：可撤销混入类 ==============

/**
 * 可撤销操作的基础混入类
 * 可以被任何类继承或混入以支持撤销功能
 */
class clsUndoableMixin {
    initUndoSupport(undoManager) {
        this.m_undoManager = undoManager || g_undoManager;
        this.m_isUndoEnabled = true;
        console.log(`[${this.MODULE_NAME || "Unknown"}] 撤销功能已初始化`);
    }
    executeUndoable(type, description, doFn, undoFn, metadata) {
        if (!this.m_isUndoEnabled || !this.m_undoManager) return doFn();
        const command = new clsUndoableCommand(this, type, description, doFn, undoFn, metadata);
        return this.m_undoManager.execute(command);
    }
    setPropertyWithUndo(property, newValue, description) {
        const oldValue = this[property];
        if (oldValue === newValue) return true;
        const self = this;
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY,
            description || `修改属性 ${property}`,
            function() { self[property] = newValue; return true; },
            function() { self[property] = oldValue; return true; },
            { property, oldValue, newValue }
        );
    }
    undo() { return this.m_undoManager ? this.m_undoManager.undo() : { success: false, message: "撤销管理器未初始化" }; }
    redo() { return this.m_undoManager ? this.m_undoManager.redo() : { success: false, message: "撤销管理器未初始化" }; }
    canUndo() { return this.m_undoManager ? this.m_undoManager.canUndo() : false; }
    canRedo() { return this.m_undoManager ? this.m_undoManager.canRedo() : false; }
}

/**
 * 为 clsRentalCalculation 添加撤销支持的示例类
 */
class clsRentalCalculationWithUndo extends clsRentalCalculation {
    constructor(parameterManager, undoManager) {
        super(parameterManager);
        this.initUndoSupport(undoManager);
        console.log(`[${this.MODULE_NAME}] 撤销功能已集成`);
    }
    initUndoSupport(undoManager) {
        this.m_undoManager = undoManager || g_undoManager;
        this.m_isUndoEnabled = true;
    }
    setParameterWithUndo(paramName, newValue) {
        const self = this;
        const oldValue = this.p ? this.p.val(paramName) : null;
        return this.executeUndoable(
            UNDO_CONFIG.OPERATION_TYPES.PARAM_CHANGE,
            `修改参数 ${paramName}: ${oldValue} -> ${newValue}`,
            function() { if (self.p) { self.p.SetParameterValue(paramName, newValue); return true; } return false; },
            function() { if (self.p && oldValue !== null) { self.p.SetParameterValue(paramName, oldValue); return true; } return false; },
            { paramName, oldValue, newValue }
        );
    }
    createWorksheetCommand(description, doFn, undoFn) {
        return new clsUndoableCommand(this, UNDO_CONFIG.OPERATION_TYPES.DATA_MODIFY, description, doFn, undoFn, { target: "worksheet" });
    }
    executeUndoable(type, description, doFn, undoFn, metadata) {
        if (!this.m_isUndoEnabled || !this.m_undoManager) return doFn();
        return this.m_undoManager.execute(new clsUndoableCommand(this, type, description, doFn, undoFn, metadata));
    }
    beginTransaction(description) { if (this.m_undoManager) { this.m_undoManager.setGroupingEnabled(true); this.m_undoManager.beginGroup(description); } }
    commitTransaction() { if (this.m_undoManager) { this.m_undoManager.endGroup(); this.m_undoManager.setGroupingEnabled(false); } }
    rollbackTransaction() { if (this.m_undoManager) { this.m_undoManager.cancelGroup(); this.m_undoManager.setGroupingEnabled(false); } }
}

/**
 * 为 clsInterestRateAdjustment 添加撤销支持的示例类
 */
class clsInterestRateAdjustmentWithUndo extends clsInterestRateAdjustment {
    constructor(parameterManager, undoManager) {
        super(parameterManager);
        this.m_undoManager = undoManager || g_undoManager;
        console.log(`[${this.MODULE_NAME}] 统一撤销功能已集成`);
    }
    addAdjustmentWithUndoManager(period, newRate) {
        const self = this;
        const existingIndex = this.m_adjustmentPeriods.findIndex(item => item.period === period);
        const oldRate = existingIndex !== -1 ? this.m_adjustmentPeriods[existingIndex].newRate : null;
        const command = new clsUndoableCommand(
            this,
            UNDO_CONFIG.OPERATION_TYPES.RATE_ADJUST,
            `${oldRate !== null ? "修改" : "添加"}调息节点：第${period}期利率${(newRate * 100).toFixed(2)}%`,
            function() { return self.addAdjustmentPeriod(period, newRate); },
            function() {
                if (oldRate !== null) self.m_adjustmentPeriods[existingIndex].newRate = oldRate;
                else { const idx = self.m_adjustmentPeriods.findIndex(item => item.period === period); if (idx !== -1) self.m_adjustmentPeriods.splice(idx, 1); }
                return true;
            },
            { period, oldRate, newRate }
        );
        return this.m_undoManager.execute(command);
    }
    generateTableWithUndoManager(adjustmentArray, sourceSheetName) {
        const self = this;
        var backupData = null, worksheet = null;
        if (this.p && this.p.m_worksheet) {
            worksheet = this.p.m_worksheet;
            const startRow = this.p.RentTableStartRow, totalPeriods = this.p.TotalPeriodsCellValue;
            backupData = worksheet.Range(`A${startRow}:M${startRow + totalPeriods - 1}`).Value2;
        }
        const command = new clsUndoableCommand(
            this,
            UNDO_CONFIG.OPERATION_TYPES.TABLE_GENERATE,
            `生成调息表：${sourceSheetName}`,
            function() { return self.generateAdjustmentTable(adjustmentArray, sourceSheetName); },
            function() {
                if (worksheet && backupData) {
                    const startRow = self.p.RentTableStartRow, totalPeriods = self.p.TotalPeriodsCellValue;
                    worksheet.Range(`A${startRow}:M${startRow + totalPeriods - 1}`).Value2 = backupData;
                    worksheet.Calculate();
                }
                return true;
            },
            { adjustmentArray, sourceSheetName }
        );
        return this.m_undoManager.execute(command);
    }
}

// ============== 演示函数 ==============

/**
 * 全局撤销管理器演示
 */
function demoGlobalUndoManager() {
    console.log("========== 全局撤销管理器演示 ==========");
    const manager = getUndoManager();
    manager.execute(new clsDataModifyCommand({ value: 0 }, "value", 0, 100, "设置值为100"));
    manager.execute(new clsDataModifyCommand({ value: 100 }, "value", 100, 200, "设置值为200"));
    showUndoStatus();
    manager.undo(); showUndoStatus();
    manager.redo(); showUndoStatus();
    manager.clear();
    console.log("历史记录已清空");
    showUndoStatus();
    console.log("========== 演示结束 ==========");
}

/**
 * 批量操作（事务）演示
 */
function demoTransaction() {
    console.log("========== 事务演示 ==========");
    const manager = createUndoManager({ enableGrouping: true });
    const target = { value1: 0, value2: 0, value3: 0 };
    manager.beginGroup("批量修改三个值");
    manager.execute(new clsDataModifyCommand(target, "value1", 0, 10, "设置value1=10"));
    manager.execute(new clsDataModifyCommand(target, "value2", 0, 20, "设置value2=20"));
    manager.execute(new clsDataModifyCommand(target, "value3", 0, 30, "设置value3=30"));
    manager.endGroup();
    console.log("事务提交后的值:", target);
    console.log("撤销栈长度:", manager.m_undoStack.length);
    manager.undo();
    console.log("撤销事务后的值:", target);
    console.log("========== 演示结束 ==========");
}

/**
 * 工作表操作撤销演示
 */
function demoWorksheetUndo() {
    console.log("========== 工作表撤销演示 ==========");
    try {
        const ws = Application.ActiveSheet;
        const range = "A1:B2";
        const oldData = ws.Range(range).Value2;
        const newData = [["新值1", "新值2"], ["新值3", "新值4"]];
        const command = new clsWorksheetModifyCommand(ws, range, oldData, newData, "修改A1:B2区域数据");
        const manager = getUndoManager();
        manager.execute(command);
        console.log("数据已修改");
        if (manager.canUndo()) console.log("可以撤销修改");
    } catch (error) {
        console.log("工作表演示需要在WPS环境中运行:", error.message);
    }
    console.log("========== 演示结束 ==========");
}

console.log("[mUndoManager.js] 统一撤销管理器模块加载完成");
console.log("集成示例已合并: clsUndoableMixin / clsRentalCalculationWithUndo / clsInterestRateAdjustmentWithUndo");
console.log("演示函数: demoGlobalUndoManager() / demoTransaction() / demoWorksheetUndo()");
