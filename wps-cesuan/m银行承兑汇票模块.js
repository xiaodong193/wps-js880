/**
 * ============== 银行承兑汇票模块（重构版）==============
 * 作者：徐晓冬
 * 更新日期：2026-05-01
 * 描述：三层分离架构 — 流程控制 / 业务计算 / 表格写入
 *
 * 设计理念：
 * - 二维数组为核心：所有数据操作通过二维数组中转
 * - 公式优先：表格中体现公式，用户可看到计算逻辑
 * - 三层分离：流程层→业务层→表格层，职责清晰
 * - 复用已有模块：mShared_constants（样式/格式/常量）
 *
 * 依赖：
 * - mShared_constants.js（设置表格样式, 应用格式, 设置字体样式 等）
 * - mErrorHandler.js（g_errorHandler）
 * - JSA880.js（Array2D, $.maxArray — 可选，有则用）
 * ====================================================
 */

// ============== 配置对象 ==============
const BILL_CONFIG = {
    // ── 放款信息区域 ──
    fkHeaderRow: 10,
    fkDataStartRow: 11,
    fkDataEndCol: "J",

    // ── 综合利率区域 ──
    rateStartRow: 21,

    // ── 放款数据列索引（A~J → 0~9）──
    col: {
        date: 0,          // A: 放款日期
        amount: 1,        // B: 放款金额
        months: 2,        // C: 对付月数
        type: 3,          // D: 放款类型（银行承兑汇票/电汇）
        feeRate: 4,       // E: 手续费率
        margin: 5,        // F: 保证金金额
        interest: 6,      // G: 银承利率
        reserved: 7,      // H: 备用
        custMargin: 8,    // I: 客户保证金
        brokerFee: 9,     // J: 经纪人费用
    },

    // ── 现金流表列数 ──
    totalCols: 15,  // A~O

    // ── 现金流表头 ──
    headers: [
        "期次", "日期", "净现金流1（1-9）", "净现金流1-备注",
        "净现金流2（1-8）", "净现金流2-备注", "（1）电汇放款",
        "（2)银行承兑汇票放款-保证金", "（3）银行承兑汇票放款-尾款",
        "（4）银行承兑汇票-手续费", "（5）银行承兑汇票-利息收入",
        "（6）我司收取客户保证金/退回保证金", "（7）租金",
        "（8）名义货价", "（9）经纪人费用支付"
    ],
};


// ============== 银行承兑汇票类 ==============
class cls银行承兑汇票 {

    constructor(parameterManager) {
        this.MODULE_NAME = "银行承兑汇票模块";

        // 依赖注入参数管理器
        if (parameterManager && typeof parameterManager === 'object') {
            this.p = parameterManager;
        } else {
            if (typeof clsParameterManager === 'undefined') {
                throw new Error("[cls银行承兑汇票] clsParameterManager 未定义");
            }
            this.p = new clsParameterManager();
            this.p.Initialize("银行承兑汇票");
        }

        // 工作表引用
        this.ws = this.p.m_billSheet;
        this.wsSource = this.p.m_sourceSheet;

        // 关键行号（银承表自身的布局）
        this.RowStart = this.p.RowStart;                   // 标题行（默认26）
        this.RentTableStartRow = this.p.RentTableStartRow; // 数据起始行（默认28）

        // ★ 从源表"1租金测算表V1"读取参数，不从银承表读
        // 银承表的 B10 是表头文本"放款金额-元"，不是期数
        // FIX(BUG 1): 强制 Number() 转换，防止 Excel 文本格式导致字符串拼接
        this.totalPeriods = Number(this.wsSource.Range("B10").Value2) || 0;
        this.paymentsPerYear = Number(this.wsSource.Range("B8").Value2) || 2;

        if (this.totalPeriods <= 0) {
            throw new Error(`总期数无效: ${this.totalPeriods}（请检查源表 B10 单元格）`);
        }

        // 现金流表在源表的起始行
        this.CashFlowTablerowStart = this.RentTableStartRow + this.totalPeriods + 6;

        console.log(`[${this.MODULE_NAME}] 实例创建 - 总期数:${this.totalPeriods} (从源表读取)`);
    }


    // ╔══════════════════════════════════════════╗
    // ║  第一层：流程控制（5步编排）              ║
    // ╚══════════════════════════════════════════╝

    /**
     * 生成银承现金流量表（主入口）
     * 流程：校验 → 读数据 → 写表头 → 写数据 → 写备注 → 写利率
     */
    生成银承现金流量表() {
        try {
            // 校验
            if (!this._validate()) return false;

            this._beginUndo("生成银承现金流量表");

            // 步1：读放款数据 → 二维数组
            const fkArr = this._readFangKuan();

            // 步2：写表头
            this._writeTitle();

            // 步3：构建并写入现金流数据（二维数组 → toRange）
            const totalRows = this._writeCashFlowData(fkArr);

            // 步4：生成备注（读回数据 → 拼备注 → 写入）
            this._writeRemarks(fkArr, totalRows);

            // 步5：综合利率公式
            this._writeRates(fkArr, totalRows);

            // 步6：修复汇总区域 D1:D5 的 #REF! 引用
            this._fixSummaryFormulas();

            this._endUndo();
            console.log(`[${this.MODULE_NAME}] 银承现金流量表生成完成`);
            return true;
        } catch (error) {
            this._handleError(error, '生成银承现金流量表');
            return false;
        }
    }

    /** 清除银承表中生成的数据 */
    清除银行承兑汇票表中数据() {
        try {
            const R = BILL_CONFIG.rateStartRow;
            this.ws.Range(`A${this.RentTableStartRow}:O${this.CashFlowTablerowStart + this.totalPeriods + 101}`).Clear();
            this.ws.Range(`B${R}:B${R + 1}`).ClearContents();
            this.ws.Range(`D${R}:D${R + 2}`).ClearContents();
            this.ws.Range(`F${R}:F${R + 2}`).ClearContents();
            this.ws.Range(`H${R}:H${R + 1}`).ClearContents();
            return true;
        } catch (error) {
            this._handleError(error, '清除银行承兑汇票表中数据');
            return false;
        }
    }


    // ╔══════════════════════════════════════════╗
    // ║  第二层：业务计算（纯数组，不碰 Range）  ║
    // ╚══════════════════════════════════════════╝

    /**
     * 计算放款占用行数
     * 银承占2行（保证金+尾款），电汇占1行
     */
    _calcFangKuanRows(fkArr) {
        if (!fkArr) return 0;
        const col = BILL_CONFIG.col;
        var count = 0;
        for (const row of fkArr) {
            count += (row[col.type] === "银行承兑汇票") ? 2 : 1;
        }
        return count;
    }

    /**
     * 构建银承放款行1（放款日：保证金、手续费、客户保证金）
     *
     * @param {Object} d - 源数据行
     * @param {number} idx - 票据索引
     * @param {Object} col - 列映射配置
     * @returns {{values: Object, formulas: Object}} 行数据 { values: {colIndex: val}, formulas: {colIndex: formulaStr} }
     */
    _buildBillRow1(d, idx, col) {
        var fkDate = safeFromExcelDate(d[col.date]);
        return {
            values: {
                1: fkDate,                                 // B: 放款日期
                7: -1 * (d[col.margin] || 0),             // H: 保证金（负值=流出）
                9: -1 * ((d[col.feeRate] || 0) * (d[col.amount] || 0)),  // J: 手续费
                11: d[col.custMargin] || 0,               // L: 客户保证金
            },
            formulas: {
                0: `="0-${idx}"`,                         // A: 期次标签
                2: `=SUM(H{r}:O{r})`,                     // C: 净现金流1 = SUM(H:O)
                4: `=SUM(H{r}:O{r})`,                     // E: 净现金流2（暂同上，后续排除列9）
            }
        };
    }

    /**
     * 构建银承放款行2（对付日：尾款、利息收入）
     *
     * @param {Object} d - 源数据行
     * @param {number} idx - 票据索引
     * @param {number} row1Num - 放款行1的Excel行号（用于EDATE公式）
     * @param {Object} col - 列映射配置
     * @returns {{values: Object, formulas: Object}} 行数据
     */
    _buildBillRow2(d, idx, row1Num, col) {
        return {
            values: {
                8: -1 * ((d[col.amount] || 0) - (d[col.margin] || 0)),  // I: 尾款
                10: (d[col.interest] || 0) * (d[col.amount] || 0),      // K: 利息收入
            },
            formulas: {
                0: `="0-${idx}"`,                         // A: 期次标签
                1: `=EDATE(B${row1Num},${d[col.months] || 0})`, // B: 对付日期
                2: `=SUM(H{r}:O{r})`,                     // C: 净现金流1
                4: `=SUM(H{r}:O{r})`,                     // E: 净现金流2
            }
        };
    }

    /**
     * 构建电汇放款行
     *
     * @param {Object} d - 源数据行
     * @param {number} idx - 票据索引
     * @param {Object} col - 列映射配置
     * @returns {{values: Object, formulas: Object}} 行数据
     */
    _buildTransferRow(d, idx, col) {
        var tfDate = safeFromExcelDate(d[col.date]);
        return {
            values: {
                1: tfDate,                                 // B: 放款日期
                11: d[col.custMargin] || 0,               // L: 客户保证金
                14: d[col.brokerFee] || 0,                // O: 经纪人费用
            },
            formulas: {
                0: `="0-${idx}"`,                         // A: 期次标签
                2: `=SUM(H{r}:O{r})`,                     // C: 净现金流1
                4: `=SUM(H{r}:O{r})`,                     // E: 净现金流2
                6: `=-${d[col.amount] || 0}`,             // G: 电汇放款（负值公式）
            }
        };
    }

    /**
     * 构建还款行
     * @param {number} t - 期次（1~totalPeriods）
     * @param {Object} srcData - 从租金源表读取的数据 { date, rent, nominalPrice, brokerFee }
     */
    _buildRepayRow(t, srcData) {
        return {
            values: {
                0: t,                                     // A: 期次
                1: srcData.date,                           // B: 日期
                11: srcData.rent || 0,                     // L: 租金
                12: srcData.nominalPrice || 0,             // M: 名义货价
                13: srcData.brokerFee || 0,                // N: 经纪人费用（还款相关）
            },
            formulas: {
                2: `=SUM(H{r}:O{r})`,                     // C: 净现金流1
                4: `=SUM(H{r}:O{r})`,                     // E: 净现金流2
            }
        };
    }

    /**
     * 构建备注文本（配置驱动）
     */
    _buildRemark(rowValues) {
        // 列索引 → 备注文本映射
        const remarkMap = [
            [6,  "电汇放款"],
            [7,  "银行承兑汇票放款-保证金"],
            [8,  "银行承兑汇票放款-尾款"],
            [9,  "银行承兑汇票-手续费"],
            [10, "银行承兑汇票-利息收入"],
            // 列11：客户保证金（特殊处理）
            // 列12：租金（特殊处理）
            [13, "名义货价"],
            [14, "经纪人费用"],
        ];

        const parts = [];

        for (const [colIdx, text] of remarkMap) {
            const val = rowValues[colIdx];
            if (val !== undefined && val !== null && val !== 0 && val !== "") {
                parts.push(text);
            }
        }

        // 客户保证金（列11）：正=收取，负=退回
        const custVal = rowValues[11];
        if (custVal > 0) parts.push("我司收取客户保证金");
        else if (custVal < 0) parts.push("我司退回保证金");

        // 租金（列12）
        const rentVal = rowValues[12];
        if (rentVal !== undefined && rentVal !== 0 && rentVal !== "") {
            const period = rowValues[0];
            parts.push(`第${period}期租金`);
        }

        return parts.join("/");
    }


    // ╔══════════════════════════════════════════╗
    // ║  第三层：表格 I/O（只做读写）            ║
    // ╚══════════════════════════════════════════╝

    /** 校验前置条件 */
    _validate() {
        if (!this.ws) { console.error("找不到银承工作表。"); return false; }
        if (!this.wsSource) { console.error("找不到\"1租金测算表V1\"工作表，请先执行租金测算。"); return false; }
        if (!this.totalPeriods || this.totalPeriods <= 0) { console.error(`总期数无效(${this.totalPeriods})。`); return false; }
        return true;
    }

    /**
     * 读取放款数据 → 二维数组
     * 使用 $.maxArray（JSA880）或原生 Range.Value2
     */
    _readFangKuan() {
        try {
            const startRow = BILL_CONFIG.fkDataStartRow;
            const startCell = this.ws.Range("A" + startRow);

            // 检查是否有数据
            if (startCell.Value2 == null || startCell.Value2 === "") {
                console.log(`[${this.MODULE_NAME}] 没有放款数据`);
                return null;
            }

            // 找到数据结束行
            // FIX(BUG 3): 限制最大行数，防止 End(xlDown) 跳到最后一行(1048576)导致 OOM
            var endRow = startRow;
            try {
                endRow = startCell.End(XL.xlDown).Row;
                const MAX_DATA_ROWS = 500;
                if (endRow - BILL_CONFIG.fkHeaderRow > MAX_DATA_ROWS) {
                    endRow = BILL_CONFIG.fkHeaderRow + MAX_DATA_ROWS;
                }
            } catch (e) { /* 单行 */ }

            const n = endRow - BILL_CONFIG.fkHeaderRow;
            if (n <= 0) return null;

            // 批量读取为二维数组
            const rng = this.ws.Range(`A${startRow}:${BILL_CONFIG.fkDataEndCol}${BILL_CONFIG.fkHeaderRow + n}`);
            const arr = rng.Value2;

            // 标准化为二维数组
            if (!Array.isArray(arr)) return [[arr]];
            if (!Array.isArray(arr[0])) return [arr];

            console.log(`[${this.MODULE_NAME}] 放款数据已读取，${arr.length} 行`);
            return arr;
        } catch (error) {
            this._handleError(error, '_readFangKuan');
            return null;
        }
    }

    /**
     * 写标题 + 表头
     * 复用：设置背景颜色(), 设置字体样式(), 设置表格样式()
     */
    _writeTitle() {
        const C = BILL_CONFIG;
        const lastColLetter = String.fromCharCode(64 + C.totalCols);

        // 总标题
        const titleRange = this.ws.Range("A" + this.RowStart);
        titleRange.Value2 = "现金流及综合利率测算";
        设置背景颜色(titleRange, COLOR_WHITE);
        设置字体样式(titleRange, { name: FONT_DEFAULT, size: FONT_SIZE_TITLE, color: COLOR_BLACK });

        // 表头行（二维数组批量写入）
        const hr = this.ws.Range(`A${this.RowStart + 1}:${lastColLetter}${this.RowStart + 1}`);
        hr.Value2 = [C.headers];
        设置背景颜色(hr, COLOR_LIGHT_BLUE);
        设置字体样式(hr, { name: FONT_DEFAULT, size: FONT_SIZE_HEADER, color: COLOR_BLACK });
        hr.HorizontalAlignment = XL.HCenter;
        hr.VerticalAlignment = XL.VCenter;
        hr.WrapText = true;
    }

    /**
     * 构建并写入现金流数据
     * 核心方法：二维数组构建 → 批量写入 → 公式列单独写入
     * @returns {number} 总行数（放款行 + 还款行）
     */
    _writeCashFlowData(fkArr) {
        const col = BILL_CONFIG.col;
        const fkRows = this._calcFangKuanRows(fkArr);
        const totalRows = fkRows + this.totalPeriods;
        const startRow = this.RentTableStartRow;

        // ── 构建完整二维数组（totalRows × 15列）──
        // FIX(BUG 7): 使用 null 替代 undefined，避免 WPS JSA 对 undefined 处理不一致
        const table = [];
        for (var i = 0; i < totalRows; i++) {
            table.push(new Array(BILL_CONFIG.totalCols).fill(null));
        }

        // 收集公式（行号 → {colIndex: formulaStr}）
        const formulaMap = {};
        var k2 = 0;  // 当前写入行偏移

        // ── 放款阶段 ──
        if (fkArr) {
            for (var i = 0; i < fkArr.length; i++) {
                const d = fkArr[i];
                const isBill = d[col.type] === "银行承兑汇票";

                if (isBill) {
                    // 银承占2行
                    const row1Idx = k2;
                    const row2Idx = k2 + 1;
                    const row1Num = startRow + k2;

                    // 行1：放款日
                    const r1 = this._buildBillRow1(d, i, col);
                    this._applyRowData(table[row1Idx], r1.values);
                    formulaMap[row1Idx] = this._resolveFormulas(r1.formulas, startRow + row1Idx);

                    // 行2：对付日
                    const r2 = this._buildBillRow2(d, i, row1Num, col);
                    this._applyRowData(table[row2Idx], r2.values);
                    formulaMap[row2Idx] = this._resolveFormulas(r2.formulas, startRow + row2Idx);

                    k2 += 2;
                } else {
                    // 电汇占1行
                    const rowIdx = k2;
                    const r = this._buildTransferRow(d, i, col);
                    this._applyRowData(table[rowIdx], r.values);
                    formulaMap[rowIdx] = this._resolveFormulas(r.formulas, startRow + rowIdx);
                    k2 += 1;
                }
            }
        }
        console.log(`[${this.MODULE_NAME}] 放款占用行数: ${k2}`);

        // ── 还款阶段 ──
        for (var t = 1; t <= this.totalPeriods; t++) {
            const rowIdx = k2;
            const srcBase = this.CashFlowTablerowStart + t;

            // 从源表读取还款数据
            var repayDate = safeFromExcelDate(this.wsSource.Range("B" + (this.RentTableStartRow - 1 + t)).Value2);
            const srcData = {
                date: repayDate,
                rent: this.wsSource.Range("I" + srcBase).Value2 || 0,
                nominalPrice: this.wsSource.Range("C" + srcBase).Value2 || 0,
                brokerFee: this.wsSource.Range("K" + srcBase).Value2 || 0,
            };

            const r = this._buildRepayRow(t, srcData);
            this._applyRowData(table[rowIdx], r.values);
            formulaMap[rowIdx] = this._resolveFormulas(r.formulas, startRow + rowIdx);
            k2++;
        }

        // ── 批量写入二维数组 ──
        const endRow = startRow + totalRows - 1;
        const lastColLetter = String.fromCharCode(64 + BILL_CONFIG.totalCols);
        const dataRange = this.ws.Range(`A${startRow}:${lastColLetter}${endRow}`);
        dataRange.Value2 = table;

        // ── 写入公式列（Value2 无法写公式）──
        for (const [rowIdx, formulas] of Object.entries(formulaMap)) {
            const rowNum = startRow + parseInt(rowIdx);
            for (const [colIdx, formula] of Object.entries(formulas)) {
                const cell = this.ws.Cells(rowNum, parseInt(colIdx) + 1);  // colIdx从0开始，Cells从1开始
                if (formula.includes("R1C1")) {
                    cell.FormulaR1C1 = formula;
                } else {
                    cell.Formula = formula;
                }
            }
        }

        // ── 应用数字格式（复用 应用格式()）──
        const rangeG_O = this.ws.Range(`G${startRow}:O${endRow}`);
        const rangeC = this.ws.Range(`C${startRow}:C${endRow}`);
        const rangeE = this.ws.Range(`E${startRow}:E${endRow}`);
        const rangeB = this.ws.Range(`B${startRow}:B${endRow}`);
        应用格式(rangeG_O, "Standard");
        应用格式(rangeC, "Standard");
        应用格式(rangeE, "Standard");
        应用格式(rangeB, "Date");

        // ── 应用表格样式（复用 设置表格样式()）──
        设置表格样式(dataRange);

        return totalRows;
    }

    /**
     * 生成并写入备注
     * 读回刚写入的数据 → 构建备注文本 → 批量写入D列和F列
     */
    _writeRemarks(fkArr, totalRows) {
        try {
            const startRow = this.RentTableStartRow;
            const endRow = startRow + totalRows - 1;

            // FIX(BUG 4): 强制计算，确保公式结果已就绪
            try { Application.Calculate(); } catch (e) { /* WPS 环境 */ }

            // 读回数据（含公式计算结果）
            const lastColLetter = String.fromCharCode(64 + BILL_CONFIG.totalCols);
            const rawData = this.ws.Range(`A${startRow}:${lastColLetter}${endRow}`).Value2;

            if (!rawData || !Array.isArray(rawData)) return;

            // 构建备注数组
            const remark1Arr = [];
            const remark2Arr = [];

            for (var i = 0; i < rawData.length; i++) {
                const row = Array.isArray(rawData[i]) ? rawData[i] : [rawData[i]];
                const remark = this._buildRemark(row);
                remark1Arr.push([remark]);
                remark2Arr.push([remark]);
            }

            // 批量写入D列（净现金流1备注）和F列（净现金流2备注）
            this.ws.Range(`D${startRow}`).Resize(totalRows, 1).Value2 = remark1Arr;
            this.ws.Range(`F${startRow}`).Resize(totalRows, 1).Value2 = remark2Arr;

            // 对齐样式
            const rangeD = this.ws.Range(`D${startRow}:D${endRow}`);
            const rangeF = this.ws.Range(`F${startRow}:F${endRow}`);
            rangeD.HorizontalAlignment = XL.HCenter;
            rangeF.HorizontalAlignment = XL.HCenter;

            console.log(`[${this.MODULE_NAME}] 备注生成完成`);
        } catch (error) {
            this._handleError(error, '_writeRemarks');
        }
    }

    /**
     * 写综合利率公式（IRR / XIRR / 对比）
     * 对齐 mCashFlowGenerator.综合利率一览()
     */
    _writeRates(fkArr, totalRows) {
        try {
            const R = BILL_CONFIG.rateStartRow;
            const r = this.RentTableStartRow;
            const lastDataRow = r + totalRows - 1;

            // ── XIRR ──
            this.ws.Range(`D${R}`).Formula = `=XIRR(C${r}:C${lastDataRow},B${r}:B${lastDataRow})`;
            this.ws.Range(`D${R + 1}`).Formula = `=XIRR(E${r}:E${lastDataRow},B${r}:B${lastDataRow})`;
            this.ws.Range(`D${R + 2}`).FormulaR1C1 = "=R[-2]C-R[-1]C";

            // ── IRR（使用源表单元格引用而非硬编码数值，确保用户修改 B8 后公式自动更新）
            // FIX(BUG 5): 使用单元格引用代替硬编码 paymentsPerYear 数值
            const srcName = this.wsSource.Name;
            this.ws.Range(`B${R}`).Formula = `=IRR(C${r}:C${lastDataRow})*${srcName}!B8`;
            this.ws.Range(`B${R + 1}`).Formula = `=IRR(E${r}:E${lastDataRow})*${srcName}!B8`;

            // ── 对比区（与租金测算表对比）
            // FIX(BUG 6): 使用变量工作表名代替硬编码字符串，避免工作表重命名后公式 #REF!
            this.ws.Range(`F${R}`).Formula = `='${srcName}'!D16`;
            this.ws.Range(`F${R + 1}`).Formula = `='${srcName}'!D17`;
            this.ws.Range(`F${R + 2}`).Formula = `=F${R}-F${R + 1}`;

            // ── 差额 ──
            this.ws.Range(`H${R}`).Formula = `=D${R}-F${R}`;
            this.ws.Range(`H${R + 1}`).Formula = `=D${R + 1}-F${R + 1}`;

            // ── 百分比格式（复用 应用格式()）──
            应用格式(this.ws.Range(`B${R}:B${R + 1}`), "Percentage");
            应用格式(this.ws.Range(`D${R}:D${R + 2}`), "Percentage");
            应用格式(this.ws.Range(`F${R}:F${R + 2}`), "Percentage");
            应用格式(this.ws.Range(`H${R}:H${R + 1}`), "Percentage");

            // ── 标签 ──
            this.ws.Range(`A${R}`).Value2 = "IRR内含报酬率";
            this.ws.Range(`A${R + 1}`).Value2 = "(1)企业看IRR";
            this.ws.Range(`C${R}`).Value2 = "XIRR净内含报酬率";
            this.ws.Range(`C${R + 1}`).Value2 = "(1)企业看XIRR";
            this.ws.Range(`C${R + 2}`).Value2 = "(2)经纪人费用影响";
            this.ws.Range(`E${R}`).Value2 = "租金测算表对比";
            this.ws.Range(`E${R + 1}`).Value2 = "租金测算表XIRR";
            this.ws.Range(`E${R + 2}`).Value2 = "差异";
            this.ws.Range(`G${R}`).Value2 = "差额（银承-租金）";
            this.ws.Range(`G${R + 1}`).Value2 = "差额（银承-租金）";

            设置字体样式(this.ws.Range(`A${R}:H${R + 2}`), { name: FONT_DEFAULT, size: FONT_SIZE_NORMAL });

            console.log(`[${this.MODULE_NAME}] 综合利率公式写入完成`);
        } catch (error) {
            this._handleError(error, '_writeRates');
        }
    }


    // ╔══════════════════════════════════════════╗
    // ║  辅助方法                               ║
    // ╚══════════════════════════════════════════╝

    /**
     * 修复汇总区域 D1:D5 的 #REF! 引用
     * 
     * 问题：模板中 D1=`#REF!*银行承兑汇票!B1`、D2=`#REF!*银行承兑汇票!B2`
     * 原因：跨表引用的工作表名可能丢失或 #REF! 替代了原表名
     * 修复：重写为正确的引用公式
     * 
     * 汇总区域布局（银行承兑汇票表 Row 1~5）：
     * Row 1: A1="电汇比例" B1=0.3  C1="电汇金额"  D1=B1（引用当前表 B1）
     * Row 2: A2="电汇金额" B2=1-B1 C2="银承比例"  D2=B2*D1（=银承金额）
     * Row 3: A3="银承金额" B3=0.2  C3="银承保证金比例" D3=D2*B3（=银承保证金金额）
     * Row 4: A4="银承保证金金额" B4=1-B3 C4="银承尾款比例" D4=B4*D2（=银承尾款）
     * Row 5: A5="银承尾款" B5=0.02 C5="保证存款利率" D5=B5*D3（=保证金存款利息）
     */
    _fixSummaryFormulas() {
        try {
            // 修复 D1:D5 的公式 — 正确的引用应该指向当前表自身的 B 列
            // D1（电汇金额）= 从投放情况中 SUMIF 实际电汇金额
            // D2（银承金额）= 从投放情况中 SUMIF 实际银承金额
            this.ws.Range("D1").Formula = '=SUMIF(D11:D100,"电汇",B11:B100)';
            this.ws.Range("D2").Formula = '=SUMIF(D11:D100,"银行承兑汇票",B11:B100)';
            this.ws.Range("D3").Formula = "=D2*B3";
            this.ws.Range("D4").Formula = "=B4*D2";
            this.ws.Range("D5").Formula = "=B5*D3";

            console.log(`[${this.MODULE_NAME}] D1:D5 #REF! 引用已修复`);
        } catch (error) {
            this._handleError(error, '_fixSummaryFormulas');
        }
    }

    /**
     * 将 values 对象应用到二维数组行
     * @param {Array} tableRow - 目标二维数组行
     * @param {Object} values - {colIndex: value} 映射
     */
    _applyRowData(tableRow, values) {
        for (const [colIdx, val] of Object.entries(values)) {
            tableRow[parseInt(colIdx)] = val;
        }
    }

    /**
     * 将公式模板中的 {r} 替换为实际行号
     * @param {Object} formulas - {colIndex: formulaStr} 模板
     * @param {number} rowNum - 实际Excel行号
     * @returns {Object} 替换后的公式映射
     */
    _resolveFormulas(formulas, rowNum) {
        const resolved = {};
        for (const [colIdx, formula] of Object.entries(formulas)) {
            resolved[colIdx] = formula.replace(/\{r\}/g, String(rowNum));
        }
        return resolved;
    }

    /** 撤销管理 */
    _beginUndo(action) {
        if (typeof g_undoManager !== 'undefined') {
            try { g_undoManager.beginAction(action); } catch (e) { console.warn(`[银承] undo开始失败: ${e.message}`); }
        }
    }
    _endUndo() {
        if (typeof g_undoManager !== 'undefined') {
            try { g_undoManager.endAction(); } catch (e) { console.warn(`[银承] undo结束失败: ${e.message}`); }
        }
    }

    /** 统一错误处理 */
    _handleError(error, source) {
        const msg = `[${this.MODULE_NAME}] ${source} 失败：${error.message}`;
        if (typeof g_errorHandler !== 'undefined') {
            g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: source });
        } else {
            console.error(msg);
        }
    }

    /** 银承分类汇总 */
    // FIX(BUG 8): 从 12 列(A-L)扩展到 15 列(A-O)，原来 M/N/O 三列数据丢失
    银承分类汇总() {
        try {
            const startRow = this.RentTableStartRow;
            var maxRow;
            try { maxRow = this.ws.Range("A" + (startRow - 1)).CurrentRegion.Rows.Count; } catch (e) { maxRow = 20; }
            const endRow = startRow + maxRow - 1;
            const TOTAL_COLS = 15;  // A~O
            const arr = this.ws.Range(`A${startRow}:O${endRow}`).Value2;  // FIX: A-L → A-O
            if (!arr || arr.length === 0) return false;

            const dd = {}, hrr = [];
            var k = 0;
            for (const row of arr) {
                const key = row[1];
                if (dd[key] !== undefined) {
                    const ri = dd[key];
                    hrr[ri][0] = (hrr[ri][0] || "") + "\\" + row[0];
                    for (var c = 2; c <= TOTAL_COLS - 1; c++) hrr[ri][c] = (hrr[ri][c] || 0) + row[c];
                } else {
                    dd[key] = k;
                    hrr[k] = [...row.slice(0, TOTAL_COLS)];
                    k++;
                }
            }

            const headerRow = startRow - 1;
            const lrr = this.ws.Range(`A${headerRow}:O${headerRow}`).Value2;  // FIX: A-L → A-O
            var summarySheet;
            try { summarySheet = Application.Sheets("银承分类汇总"); } catch (e) {
                try { summarySheet = Application.Sheets.Add(); summarySheet.Name = "银承分类汇总"; } catch (e2) { summarySheet = this.ws; }
            }
            if (summarySheet) {
                summarySheet.Range("A1").Resize(1, TOTAL_COLS).Value2 = lrr;
                if (k > 0) summarySheet.Range("A2").Resize(k, TOTAL_COLS).Value2 = hrr.slice(0, k);
            }
            console.log(`[${this.MODULE_NAME}] 银承分类汇总完成，汇总行数: ${k}`);
            return true;
        } catch (error) {
            this._handleError(error, '银承分类汇总');
            return false;
        }
    }

    /** 初始化（兼容接口） */
    Initialize() {
        console.log(`[${this.MODULE_NAME}] 初始化完成`);
        return true;
    }

    // ─────────────── 兼容旧接口 ───────────────
    WriteHeaders2() { return this._writeTitle(); }
    创建银承现金流量表表头() { return this._writeTitle(); }
    银承放款现金流() { return this.生成银承现金流量表(); }
    银承综合利率一览() { /* 由主流程统一调用 */ return true; }
    银承现金流1备注update() { /* 由主流程统一调用 */ return true; }
    银承现金流2备注update() { /* 由主流程统一调用 */ return true; }
    SaveFangKuanInfoToArr() { return this._readFangKuan(); }
    nFangKuan() { var result = this._readFangKuan(); return (result && result.length) || 0; }
    arrRngCashFlow() { return null; }
    流量表数据生成() {
        const fkArr = this._readFangKuan();
        const fkRows = this._calcFangKuanRows(fkArr);
        return { totalRows: fkRows + this.totalPeriods, fangKuanRows: fkRows, repaymentRows: this.totalPeriods };
    }
}


// ============== 全局快捷函数 ==============

/** 生成银承 - 一键生成银行承兑汇票现金流量表 */
function 生成银承() {
    try {
        console.log("========== 开始生成银行承兑汇票现金流量表 ==========");
        const result = getRentSystem().生成银承现金流量表();
        if (result) {
            console.log("银行承兑汇票现金流量表生成成功！");
        } else {
            console.error("生成失败，请查看控制台日志。");
        }
        return result;
    } catch (error) {
        console.log(`生成银承 失败：${error.message}`);
        console.error(`生成失败：${error.message}`);
        return false;
    }
}

/** 测试银承 - 自动填入测试数据并生成 */
function 测试银承() {
    try {
        console.log("========== [测试] 开始银行承兑汇票模块测试 ==========");

        if (typeof cls银行承兑汇票 === 'undefined') throw new Error("cls银行承兑汇票 未定义");
        if (typeof clsParameterManager === 'undefined') throw new Error("clsParameterManager 未定义");

        // 获取或创建银承工作表
        var billSheet;
        try {
            billSheet = Application.Worksheets("银行承兑汇票");
            billSheet.Activate();
        } catch (e) {
            billSheet = Application.Worksheets.Add();
            billSheet.Name = "银行承兑汇票";
        }

        // 行11: 银行承兑汇票放款 5000万
        billSheet.Range("A11:J11").Value2 = [[
            new Date(2026, 0, 15), 50000000, 6, "银行承兑汇票",
            0.0005, 15000000, 0.03, 0, 0, 0
        ]];
        // 行12: 电汇放款 3000万
        billSheet.Range("A12:J12").Value2 = [[
            new Date(2026, 0, 15), 30000000, 0, "电汇", 0, 0, 0, 0, 50000
        ]];
        console.log("[测试银承] 测试放款数据已填入 A11:J12");

        // 确保租金测算数据存在
        try {
            const rentSheet = Application.Worksheets("1租金测算表V1");
            if (!rentSheet.Range("B4").Value2) {
                rentSheet.Range("B4").Value2 = 80000000;
                rentSheet.Range("B5").Value2 = 0.035;
                rentSheet.Range("B8").Value2 = 1;
                rentSheet.Range("B10").Value2 = 10;
                rentSheet.Range("D10").Value2 = 6;
                if (typeof 计算main === 'function') 计算main();
            }
        } catch (e) {
            console.log(`[测试银承] 警告：无法准备租金数据：${e.message}`);
        }

        const result = getRentSystem().生成银承现金流量表();
        console.log(result ? "测试成功！请查看「银行承兑汇票」工作表。" : "测试失败，请查看控制台日志。");
        return result;
    } catch (error) {
        console.log(`测试银承 失败：${error.message}\n${error.stack}`);
        console.error(`测试失败：${error.message}`);
        return false;
    }
}

/** 清除银承 - 清除银行承兑汇票表中的生成数据 */
function 清除银承() {
    try {
        const billSheet = Application.Worksheets("银行承兑汇票");
        if (!billSheet) throw new Error("找不到工作表「银行承兑汇票」");
        billSheet.Activate();
        const module = new cls银行承兑汇票();
        const result = module.清除银行承兑汇票表中数据();
        if (result) console.log("银行承兑汇票数据已清除！");
        else console.error("清除失败，请查看控制台日志。");
        return result;
    } catch (error) {
        console.error(`清除银承 失败：${error.message}`);
        return false;
    }
}
