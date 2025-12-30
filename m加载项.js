Attribute Module_Name = "m加载项"
/**
 * 租金测算系统加载项
 * 负责加载所有必要的模块和初始化系统
 */

// 加载共享常量
console.log("正在加载共享常量...");
// 确保共享常量已定义
if (typeof MODULE_NAME === 'undefined') {
    const MODULE_NAME = "共享常量";
    const AUTHOR = "徐晓冬";
    const VERSION = "4.2025.11.16";
    
    // 错误代码常量
    const ERR_INITIALIZE = 1000;
    const ERR_VALIDATION = 2000;
    const ERR_FORMULA = 3000;
    const ERR_RANGE = 4000;
    
    // 格式常量
    const FORMAT_STANDARD = "#,##0.00";
    const FORMAT_INTEGER = "0";
    const FORMAT_DATE = "yyyy-mm-dd";
    const FORMAT_PERCENTAGE = "0.00%";
    
    // 对齐常量
    const XL = {
        HCenter: -4108,
        VCenter: -4108,
        Left: -4131,
        Right: -4152,
        Top: -4160,
        Bottom: -4107
    };
    
    // 字体常量
    const FONT_DEFAULT = "黑体";
    const FONT_ENGLISH = "Arial";
    const FONT_CHINESE = "微软雅黑";
    const FONT_SIZE_TITLE = 14;
    const FONT_SIZE_HEADER = 12;
    const FONT_SIZE_NORMAL = 10;
    const FONT_SIZE_LARGE = 26;
    
    console.log("共享常量已定义");
}

// 加载参数管理器
console.log("正在加载参数管理器...");
if (typeof ParameterManager === 'undefined') {
    console.log("错误：ParameterManager未定义，请检查mParameterManager.js是否正确加载");
} else {
    console.log("参数管理器加载成功");
}

// 加载租金计算模块
console.log("正在加载租金计算模块...");
if (typeof RentalCalculation === 'undefined') {
    console.log("错误：RentalCalculation未定义，请检查mRentalCalculation.js是否正确加载");
} else {
    console.log("租金计算模块加载成功");
}

// 加载现金流量生成器
console.log("正在加载现金流量生成器...");
if (typeof CashFlowGenerator === 'undefined') {
    console.log("错误：CashFlowGenerator未定义，请检查CashFlowGenerator.js是否正确加载");
} else {
    console.log("现金流量生成器加载成功");
}

// 加载银行承兑汇票模块
console.log("正在加载银行承兑汇票模块...");
if (typeof BankAcceptanceModule === 'undefined') {
    console.log("错误：BankAcceptanceModule未定义，请检查mBankAcceptance.js是否正确加载");
} else {
    console.log("银行承兑汇票模块加载成功");
}

/**
 * 系统初始化函数
 */
function 系统初始化() {
    try {
        console.log("=== 租金测算系统 ===");
        console.log("版本：" + VERSION);
        console.log("作者：" + AUTHOR);
        console.log("系统初始化完成");
        return true;
    } catch (error) {
        console.log("系统初始化失败：" + error.message);
        return false;
    }
}

/**
 * 测试所有模块加载
 */
function 测试模块加载() {
    try {
        console.log("=== 模块加载测试 ===");
        
        // 测试参数管理器
        if (typeof ParameterManager !== 'undefined') {
            console.log("✅ 参数管理器：已加载");
            const p = new ParameterManager();
            console.log("✅ 参数管理器实例创建：成功");
        } else {
            console.log("❌ 参数管理器：未加载");
        }
        
        // 测试租金计算模块
        if (typeof RentalCalculation !== 'undefined') {
            console.log("✅ 租金计算模块：已加载");
        } else {
            console.log("❌ 租金计算模块：未加载");
        }
        
        // 测试现金流量生成器
        if (typeof CashFlowGenerator !== 'undefined') {
            console.log("✅ 现金流量生成器：已加载");
        } else {
            console.log("❌ 现金流量生成器：未加载");
        }
        
        // 测试银行承兑汇票模块
        if (typeof BankAcceptanceModule !== 'undefined') {
            console.log("✅ 银行承兑汇票模块：已加载");
        } else {
            console.log("❌ 银行承兑汇票模块：未加载");
        }
        
        console.log("模块加载测试完成");
        return true;
        
    } catch (error) {
        console.log("模块加载测试失败：" + error.message);
        return false;
    }
}

// 立即执行系统初始化
系统初始化();

// 可选：执行模块加载测试
// 测试模块加载();
