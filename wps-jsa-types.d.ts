/**
 * WPS JSA 类型定义文件 (.d.ts)
 *
 * 使用方法：
 * 1. 将此文件放在你的项目根目录
 * 2. 在 JS 文件顶部添加：/// <reference path="wps-jsa-types.d.ts" />
 * 3. 在 VS Code 中编写代码时将获得完整的类型提示
 *
 * 注意：这是一个声明文件，不会被编译到最终的代码中
 * 只是用于开发时的类型检查和智能提示
 */

// ============================================
// WPS Application 对象
// ============================================

declare const Application: WPSApplication;

interface WPSApplication {
    /** Excel 应用程序名称 */
    Name: string;
    /** Excel 版本 */
    Version: string;
    /** 当前活动工作簿 */
    ActiveWorkbook: WPSWorkbook;
    /** 当前活动工作表 */
    ActiveSheet: WPSSheet;
    /** 当前窗口 */
    ActiveWindow: WPSWindow;
    /** 所有工作簿集合 */
    Workbooks: WPSWorkbooks;
    /** 所有工作表集合 */
    Sheets: WPSSheets;
    /** 计算模式 */
    Calculation: XlCalculation;
    /** 显示警告 */
    DisplayAlerts: boolean;
    /** 屏幕更新 */
    ScreenUpdating: boolean;

    /** 显示消息框 */
    MessageBox(prompt: string, buttons?: any, title?: string): number;
    /** 退出 */
    Quit(): void;
    /** 运行 */
    Run(macro: string): any;
}

// ============================================
// Workbook 对象
// ============================================

interface WPSWorkbook {
    /** 工作簿名称 */
    Name: string;
    /** 工作簿路径 */
    Path: string;
    /** 是否已保存 */
    Saved: boolean;
    /** 所有工作表 */
    Worksheets: WPSSheets;
    /** 当前活动工作表 */
    ActiveSheet: WPSSheet;

    /** 保存工作簿 */
    Save(): void;
    /** 另存为 */
    SaveAs(filename: string): void;
    /** 关闭工作簿 */
    Close(saveChanges?: any): void;
}

// ============================================
// Worksheet 对象
// ============================================

interface WPSSheet {
    /** 工作表名称 */
    Name: string;
    /** 工作表索引 */
    Index: number;
    /** 是否可见 */
    Visible: any;
    /** 单元格区域 */
    Cells: WPSRange;
    /** 已使用区域 */
    UsedRange: WPSRange;
    /** 列 */
    Columns: WPSRange;
    /** 行 */
    Rows: WPSRange;

    /** 激活工作表 */
    Activate(): void;
    /** 删除工作表 */
    Delete(): void;
    /** 复制工作表 */
    Copy(before?: any, after?: any): void;
    /** 移动工作表 */
    Move(before?: any, after?: any): void;
    /** 获取区域 */
    Range(arg1: any, arg2?: any): WPSRange;
}

// ============================================
// Range 对象
// ============================================

interface WPSRange {
    /** 区域地址 */
    Address: string;
    /** 单元格值 */
    Value: any;
    /** 单元格值（推荐使用） */
    Value2: any;
    /** 公式 */
    Formula: string;
    /** 公式数组 */
    FormulaArray: any;
    /** 文本 */
    Text: string;
    /** 列索引 */
    Column: number;
    /** 行索引 */
    Row: number;
    /** 列宽 */
    ColumnWidth: number;
    /** 行高 */
    RowHeight: number;
    /** 水平对齐 */
    HorizontalAlignment: any;
    /** 垂直对齐 */
    VerticalAlignment: any;
    /** 字体 */
    Font: WPSFont;
    /** 内部样式 */
    Interior: WPSInterior;
    /** 边框 */
    Borders: WPSBorders;
    /** 合并单元格 */
    MergeCells: boolean;
    /** 区域左上角列 */
    Left: number;
    /** 区域顶部行 */
    Top: number;
    /** 区域宽度 */
    Width: number;
    /** 区域高度 */
    Height: number;
    /** 区域行数 */
    Rows: WPSRange;
    /** 区域列数 */
    Columns: WPSRange;

    /** 激活区域 */
    Activate(): void;
    /** 选择区域 */
    Select(): void;
    /** 清除内容 */
    ClearContents(): void;
    /** 清除全部 */
    Clear(): void;
    /** 复制 */
    Copy(destination?: any): void;
    /** 粘贴 */
    PasteSpecial(paste?: any): void;
    /** 自动调整列宽 */
    AutoFit(): void;
    /** 合并单元格 */
    Merge(across?: boolean): void;
    /** 取消合并 */
    UnMerge(): void;
    /** 获取指定偏移的区域 */
    Offset(rowOffset: number, columnOffset: number): WPSRange;
    /** 获取调整大小后的区域 */
    Resize(rowSize: number, columnSize: number): WPSRange;
}

// ============================================
// Font 对象
// ============================================

interface WPSFont {
    /** 字体名称 */
    Name: string;
    /** 字体大小 */
    Size: number;
    /** 是否粗体 */
    Bold: boolean;
    /** 是否斜体 */
    Italic: boolean;
    /** 下划线 */
    Underline: any;
    /** 颜色（索引） */
    ColorIndex: number;
    /** 颜色（RGB） */
    Color: number;
}

// ============================================
// Interior 对象
// ============================================

interface WPSInterior {
    /** 背景色（索引） */
    ColorIndex: number;
    /** 背景色（RGB） */
    Color: number;
    /** 填充样式 */
    Pattern: any;
}

// ============================================
// Borders 对象
// ============================================

interface WPSBorders {
    /** 左边框 */
    LeftBorder(el: any): WPSBorder;
    /** 右边框 */
    RightBorder(el: any): WPSBorder;
    /** 上边框 */
    TopBorder(el: any): WPSBorder;
    /** 下边框 */
    BottomBorder(el: any): WPSBorder;
}

interface WPSBorder {
    /** 线条样式 */
    LineStyle: any;
    /** 线条粗细 */
    Weight: any;
    /** 颜色 */
    Color: number;
    /** 颜色索引 */
    ColorIndex: number;
}

// ============================================
// Window 对象
// ============================================

interface WPSWindow {
    /** 窗口标题 */
    Caption: string;
    /** 窗口高度 */
    Height: number;
    /** 窗口宽度 */
    Width: number;
    /** 显示滚动条 */
    DisplayScrollBars: boolean;
    /** 显示网格线 */
    DisplayGridlines: boolean;
    /** 显示行号列标 */
    DisplayHeadings: boolean;
}

// ============================================
// Collections
// ============================================

interface WPSWorkbooks {
    /** 工作簿数量 */
    Count: number;
    /** 获取工作簿 */
    Item(index: any): WPSWorkbook;
    /** 添加工作簿 */
    Add(template?: any): WPSWorkbook;
    /** 打开工作簿 */
    Open(filename: string): WPSWorkbook;
    /** 关闭所有 */
    Close(): void;
}

interface WPSSheets {
    /** 工作表数量 */
    Count: number;
    /** 获取工作表 */
    Item(index: any): WPSSheet;
    /** 添加工作表 */
    Add(before?: any, after?: any, count?: any): WPSSheet;
}

// ============================================
// 枚举
// ============================================

declare enum XlCalculation {
    xlCalculationAutomatic = -4105,
    xlCalculationManual = -4135,
    xlCalculationSemiautomatic = 2
}

declare enum XlHAlign {
    xlHAlignCenter = -4108,
    xlHAlignLeft = -4131,
    xlHAlignRight = -4152,
    xlHAlignGeneral = 1
}

declare enum XlVAlign {
    xlVAlignCenter = -4108,
    xlVAlignTop = -4160,
    xlVAlignBottom = -4107
}

declare enum XlBordersIndex {
    xlEdgeLeft = 7,
    xlEdgeTop = 8,
    xlEdgeBottom = 9,
    xlEdgeRight = 10,
    xlInsideHorizontal = 12,
    xlInsideVertical = 11
}

declare enum XlLineStyle {
    xlContinuous = 1,
    xlDash = -4115,
    xlDot = -4118,
    xlDouble = -4119,
    xlNone = -4142
}

// ============================================
// 全局函数
// ============================================

/** 输出到控制台 */
declare function Console_log(message: any): void;

/** 输入框 */
declare function InputBox(prompt: string, title?: string, default?: string): string;

/** 等待 */
declare function Application_Wait(time: number): void;

// ============================================
// Console 对象
// ============================================

declare const Console: {
    log(message: any): void;
    error(message: any): void;
    warn(message: any): void;
    info(message: any): void;
};

// ============================================
// Array 扩展（JSA 支持）
// ============================================

interface Array<T> {
    /** JSA 支持 flatMap */
    flatMap<U>(callback: (value: T, index: number, array: T[]) => U[]): U[];
    /** JSA 支持 flat */
    flat(depth?: number): any[];
}

// ============================================
// 导出声明
// ============================================

declare global {
    const Application: WPSApplication;
    const Console: {
        log(message: any): void;
        error(message: any): void;
        warn(message: any): void;
        info(message: any): void;
    };
}

export {};
