/**
 * ============== 租金测算系统 - 调息功能扩展类 ==============
 * 作者：AI Assistant
 * 描述：继承自 RentalCalculation，增加利率调整功能
 * ====================================================
 */

class cls调息 extends RentalCalculation {
    constructor() {
        super(); // 调用父类构造函数
        this.MODULE_NAME = "cls调息";
        console.log(`[${this.MODULE_NAME}] 调息功能类实例创建`);
        
        // 新增属性：调息相关
        this.adjustmentPeriods = []; // 存储调息期次数组 [{period: 5, newRate: 0.05}, ...]
        this.p=p;
        this.adjustmentConfig = {
            isEnabled: false,           // 是否启用调息
            adjustmentType: "固定调整",  // 调整类型："固定调整"、"浮动调整"、"自定义"
            adjustmentBasis: "基准利率", // 调整依据："基准利率"、"LPR"、"固定值"
            adjustmentValue: 0,        // 调整幅度（正数为上浮，负数为下浮）
            repaymentMethod: "等额本金法",  // 还款方式，包括"等额本金法"、"等额租金法"、"本金比例法"等
            periodChgStart:3 ,// 调整起始期次,

            
        };
    }

    /**
     * 初始化调息配置
     * @param {Object} config - 调息配置对象
     * @returns {Boolean} 是否成功
     */
    InitializeAdjustment(config = {}) {
        try {
            console.log(`[${this.MODULE_NAME}] 开始初始化调息配置...`);
            
            // 合并配置
            this.adjustmentConfig = {
                ...this.adjustmentConfig,
                ...config
            };
            
            console.log(`[${this.MODULE_NAME}] 调息配置：`, this.adjustmentConfig);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 调息配置初始化失败：${error.message}`);
            return false;
        }
    }

    /**
     * 添加调息节点
     * @param {Number} period - 调息期次（从第几期开始调整）
     * @param {Number} newRate - 新利率（年化利率，例如 0.05 表示 5%）
     * @returns {Boolean} 是否成功
     */
    AddAdjustmentPeriod(period, newRate) {
        try {
            // 验证参数
            if (!Number.isInteger(period) || period < 1 || period > this.p.TotalPeriodsCellValue) {
                throw new Error(`期次参数错误：${period}，必须在 1-${this.p.TotalPeriodsCellValue} 之间`);
            }
            
            if (typeof newRate !== 'number' || newRate < 0 || newRate > 1) {
                throw new Error(`利率参数错误：${newRate}，必须在 0-1 之间`);
            }
            
            // 检查是否已存在该期次
            const existingIndex = this.adjustmentPeriods.findIndex(item => item.period === period);
            
            if (existingIndex !== -1) {
                // 更新已有记录
                this.adjustmentPeriods[existingIndex].newRate = newRate;
                console.log(`[${this.MODULE_NAME}] 更新第 ${period} 期调息利率为 ${(newRate * 100).toFixed(2)}%`);
            } else {
                // 添加新记录
                this.adjustmentPeriods.push({ period, newRate });
                console.log(`[${this.MODULE_NAME}] 添加第 ${period} 期调息节点，利率 ${(newRate * 100).toFixed(2)}%`);
            }
            
            // 按期次排序
            this.adjustmentPeriods.sort((a, b) => a.period - b.period);
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 添加调息节点失败：${error.message}`);
            return false;
        }
    }

    /**
     * 批量添加调息节点
     * @param {Array} adjustments - 调息数组 [{period: 5, newRate: 0.05}, ...]
     * @returns {Boolean} 是否成功
     */
    BatchAddAdjustments(adjustments) {
        try {
            if (!Array.isArray(adjustments)) {
                throw new Error("参数必须是数组");
            }
            
            for (const item of adjustments) {
                this.AddAdjustmentPeriod(item.period, item.newRate);
            }
            
            console.log(`[${this.MODULE_NAME}] 批量添加 ${adjustments.length} 个调息节点完成`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 批量添加调息节点失败：${error.message}`);
            return false;
        }
    }

    /**
     * 获取指定期次的适用利率
     * @param {Number} period - 期次
     * @returns {Number} 该期适用的利率
     */
    GetApplicableRate(period) {
        try {
            // 如果未启用调息，返回原始利率
            if (!this.adjustmentConfig.isEnabled || this.adjustmentPeriods.length === 0) {
                return this.p.InterestRateCellValue;
            }
            
            // 查找该期次之前最近的调息节点
            let applicableRate = this.p.InterestRateCellValue;
            
            for (const adjustment of this.adjustmentPeriods) {
                if (adjustment.period <= period) {
                    applicableRate = adjustment.newRate;
                } else {
                    break; // 已排序，后续不需要再检查
                }
            }
            
            return applicableRate;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 获取适用利率失败：${error.message}`);
            return this.p.InterestRateCellValue; // 返回默认利率
        }
    }

    /**
     * 清除所有调息节点
     */
    ClearAdjustments() {
        this.adjustmentPeriods = [];
        console.log(`[${this.MODULE_NAME}] 已清除所有调息节点`);
    }

    /**
     * 重写：等额本金法按天计息（带调息功能）
     * @returns {Array} 公式数组
     */
    调息等额本金法按天计息arr() {
        try {
            if (this.p === null) {
                throw new Error("ParameterManager未初始化");
            }

            // 如果未启用调息，调用父类方法
            if (!this.adjustmentConfig.isEnabled) {
                return super.等额本金法按天计息arr();
            }

            // 创建存储单元格公式的二维数组
            const arr = [];
            for (let i = 0; i < 5; i++) {
                arr[i] = new Array(13);
            }

            // 第1期每列公式（使用第13列存储每期适用利率）
            arr[1][1] = "1";
            arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`;
            arr[1][3] = "=ROUND(RC[1]+RC[2],2)"; // 租金
            arr[1][4] = `=ROUND(${this.p.PrincipalCellR1C1}/${this.p.TotalPeriodsCellR1C1},2)`; // 本金
            arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*RC[8]/360*(RC[-3]-${this.p.LeaseStartDateCellR1C1}),2)`; // 利息（引用M列利率）
            arr[1][6] = "=RC[-2]";
            arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`;
            arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`;
            arr[1][9] = "=RC[-6]";
            arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`;
            arr[1][13] = this.GetApplicableRate(1); // 第1期适用利率（静态值）

            // 第2期-倒数第2期
            arr[2][1] = "=R[-1]C+1";
            arr[2][2] = "=EDATE(R[-1]C, " + this.p.PaymentIntervalCellValue + ")";
            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*RC[8]/360*(RC[-3]-R[-1]C[-3]),2)`; // 利息（引用M列利率）
            arr[2][6] = "=RC[-2] + R[-1]C";
            arr[2][7] = arr[1][7];
            arr[2][8] = arr[1][8];
            arr[2][9] = "=RC[-6] + R[-1]C";
            arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
            arr[2][13] = "=R[-1]C"; // 默认继承上期利率，后续根据调息节点修改

            // 倒数第1期
            arr[3][1] = arr[2][1];
            arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")";
            arr[3][3] = arr[1][3];
            arr[3][4] = `=${this.p.PrincipalCellR1C1} - SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`;
            arr[3][5] = arr[2][5];
            arr[3][6] = arr[2][6];
            arr[3][7] = arr[1][7];
            arr[3][8] = arr[1][8];
            arr[3][9] = "=RC[-6] + R[-1]C";
            arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
            arr[3][13] = "=R[-1]C";

            console.log(`[${this.MODULE_NAME}] 带调息功能的等额本金法公式数组生成成功`);
            return arr;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 带调息功能的等额本金法公式数组生成失败：${error.message}`);
            return null;
        }
    }

    /**
     * 重写：本金比例法按天计息（带调息功能）
     * @returns {Array} 公式数组
     */
    本金比例法按天计息arr() {
        try {
            if (this.p === null) {
                throw new Error("ParameterManager未初始化");
            }

            // 如果未启用调息，调用父类方法
            if (!this.adjustmentConfig.isEnabled) {
                return super.本金比例法按天计息arr();
            }

            // 创建存储单元格公式的二维数组
            const arr = [];
            for (let i = 0; i < 5; i++) {
                arr[i] = new Array(13);
            }

            // 第1期每列公式（使用第13列存储每期适用利率）
            arr[1][1] = "1";
            arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`;
            arr[1][3] = "=ROUND(RC[1]+RC[2],2)";
            arr[1][4] = `=ROUND(${this.p.PrincipalCellR1C1}*RC[8]/100,2)`;
            arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*RC[8]/360*(RC2-${this.p.LeaseStartDateCellR1C1}),2)`; // 利息（引用M列利率）
            arr[1][6] = "=RC[-2]";
            arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`;
            arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`;
            arr[1][9] = "=RC[-6]";
            arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`;
            arr[1][12] = `=round(100/${this.p.TotalPeriodsCellR1C1},2)`;
            arr[1][13] = this.GetApplicableRate(1); // 第1期适用利率

            // 第2期-倒数第2期
            arr[2][1] = "=R[-1]C+1";
            arr[2][2] = "=EDATE(R[-1]C, " + this.p.PaymentIntervalCellValue + ")";
            arr[2][3] = arr[1][3];
            arr[2][4] = arr[1][4];
            arr[2][5] = `=ROUND(R[-1]C[2]*RC[8]/360*(RC[-3]-R[-1]C[-3]),2)`; // 利息（引用M列利率）
            arr[2][6] = "=RC[-2] + R[-1]C";
            arr[2][7] = arr[1][7];
            arr[2][8] = arr[1][8];
            arr[2][9] = "=RC[-6] + R[-1]C";
            arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
            arr[2][12] = arr[1][12];
            arr[2][13] = "=R[-1]C"; // 默认继承上期利率

            // 倒数第1期
            arr[3][1] = arr[2][1];
            arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")";
            arr[3][3] = arr[1][3];
            arr[3][4] = arr[1][4];
            arr[3][5] = arr[2][5];
            arr[3][6] = arr[2][6];
            arr[3][7] = arr[1][7];
            arr[3][8] = arr[1][8];
            arr[3][9] = "=RC[-6] + R[-1]C";
            arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
            arr[3][12] = `=100-SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`;
            arr[3][13] = "=R[-1]C";

            console.log(`[${this.MODULE_NAME}] 带调息功能的本金比例法公式数组生成成功`);
            return arr;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 带调息功能的本金比例法公式数组生成失败：${error.message}`);
            return null;
        }
    }

    /**
     * 重写：生成数据区域（增加调息列）
     * @returns {Boolean}
     */
    createDataRange() {
        try {
            // 首先调用父类方法生成基础表格
            const result = super.createDataRange();
            if (!result) {
                throw new Error("父类表格生成失败");
            }

            // 如果启用了调息功能，创建第13列（每期适用利率列）
            if (this.adjustmentConfig.isEnabled) {
                this.创建租金测算表表头(13, 13); // 添加"每期适用利率"表头
                
                // 写入利率数据到M列
                const rateRange = this.p.m_worksheet.Range(
                    `M${this.p.RentTableStartRow}:M${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
                );
                
                // 生成利率数组
                const rateData = [];
                for (let period = 1; period <= this.p.TotalPeriodsCellValue; period++) {
                    rateData.push([this.GetApplicableRate(period)]);
                }
                
                rateRange.Value2 = rateData;
                rateRange.NumberFormat = "0.00%"; // 设置为百分比格式
                this.添加框线(rateRange);
                this.设置背景颜色(rateRange, this.p.m_COLOR_YELLOW);
                
                // 更新合计行
                this.租金测算表合计行(13, 13);
                
                console.log(`[${this.MODULE_NAME}] 调息列（M列）已创建`);
                
                // 标黄调息区域
                this.标黄调息区域();
            }

            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 创建数据区域失败：${error.message}`);
            return false;
        }
    }

    /**
     * 标黄调息区域（期次序号、本金、利息）
     * 从首个调息期次开始的所有期次都要标黄
     */
    标黄调息区域() {
        try {
            if (!this.adjustmentConfig.isEnabled || this.adjustmentPeriods.length === 0) {
                console.log(`[${this.MODULE_NAME}] 未启用调息或无调息节点，跳过标黄`);
                return;
            }
            
            // 获取第一个调息节点
            const firstAdjustment = this.adjustmentPeriods[0];
            const firstAdjustmentPeriod = firstAdjustment.period;
            
            console.log(`[${this.MODULE_NAME}] 开始标黄调息区域，起始期次：${firstAdjustmentPeriod}`);
            
            // 计算行范围
            const startRow = this.p.RentTableStartRow + firstAdjustmentPeriod - 1;
            const endRow = this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1;
            
            // 标黄第1列（期次序号）
            const rangeCol1 = this.p.m_worksheet.Range(`A${startRow}:A${endRow}`);
            this.设置背景颜色(rangeCol1, this.p.m_COLOR_YELLOW);
            
            // 标黄第4列（本金）
            const rangeCol4 = this.p.m_worksheet.Range(`D${startRow}:D${endRow}`);
            this.设置背景颜色(rangeCol4, this.p.m_COLOR_YELLOW);
            
            // 标黄第5列（利息）
            const rangeCol5 = this.p.m_worksheet.Range(`E${startRow}:E${endRow}`);
            this.设置背景颜色(rangeCol5, this.p.m_COLOR_YELLOW);
            
            console.log(`[${this.MODULE_NAME}] 调息区域标黄完成（行${startRow}-${endRow}，列A、D、E）`);
            
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 标黄调息区域失败：${error.message}`);
        }
    }

    /**
     * 从工作表读取调息配置（例如从特定单元格区域读取）
     * @param {String} configRangeA1 - 配置区域地址，例如 "P2:Q10"
     * @returns {Boolean}
     */
    LoadAdjustmentsFromSheet(configRangeA1) {
        try {
            const range = this.p.m_worksheet.Range(configRangeA1);
            const data = range.Value2;
            
            if (!Array.isArray(data)) {
                throw new Error("读取数据格式错误");
            }
            
            this.ClearAdjustments();
            
            // 假设第一列是期次，第二列是新利率
            for (let i = 0; i < data.length; i++) {
                const period = data[i][0];
                const newRate = data[i][1];
                
                if (period && newRate) {
                    this.AddAdjustmentPeriod(period, newRate);
                }
            }
            
            console.log(`[${this.MODULE_NAME}] 从工作表加载 ${this.adjustmentPeriods.length} 个调息节点`);
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 从工作表加载调息配置失败：${error.message}`);
            return false;
        }
    }

    /**
     * 获取新的工作表名称
     * 自动生成以"调息表V"开头的工作表名称
     * @returns {String} 新的工作表名称
     */
    获取新工作表名() {
        try {
            const prefix = "调息表V";
            let maxNumber = 0;
            
            // 遍历所有工作表，查找已存在的调息表
            for (let i = 1; i <= Application.Worksheets.Count; i++) {
                const sheetName = Application.Worksheets(i).Name;
                
                // 检查是否以"调息表V"开头
                if (sheetName.startsWith(prefix)) {
                    const numberStr = sheetName.substring(prefix.length);
                    const number = parseInt(numberStr, 10);
                    
                    if (!isNaN(number) && number > maxNumber) {
                        maxNumber = number;
                    }
                }
            }
            
            // 生成新的工作表名称
            const newSheetName = prefix + (maxNumber + 1);
            console.log(`[${this.MODULE_NAME}] 获取新工作表名称：${newSheetName}`);
            
            return newSheetName;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 获取新工作表名称失败：${error.message}`);
            return "调息表V1"; // 返回默认名称
        }
    }

    /**
     * 复制工作表到新工作表
     * @param {String} 源工作表名 - 源工作表名称
     * @returns {Object} 新工作表对象
     */
    复制工作表(源工作表名) {
        try {
            console.log(`[${this.MODULE_NAME}] 开始复制工作表：${源工作表名}`);
            
            // 获取源工作表
            const sourceSheet = Application.Worksheets(源工作表名);
            
            // 获取新工作表名称
            const newSheetName = this.获取新工作表名();
            
            // 复制工作表（在源工作表之后复制）
            sourceSheet.Copy(null, sourceSheet);
            
            // 获取刚复制的工作表（最后一个工作表）
            const newSheet = Application.Worksheets(Application.Worksheets.Count);
            
            // 重命名为新名称
            newSheet.Name = newSheetName;
            
            console.log(`[${this.MODULE_NAME}] 工作表复制成功，新表名：${newSheetName}`);
            
            return newSheet;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 复制工作表失败：${error.message}`);
            return null;
        }
    }

    /**
     * 生成调息表
     * @param {Array} 调息数组 - 调息节点数组 [{period: 3, newRate: 0.025}, ...]
     * @param {String} 源工作表名 - 源工作表名称，默认为"1租金测算表V1"
     * @returns {Boolean} 是否成功
     */
    生成调息表(调息数组, 源工作表名 = "1租金测算表V1") {
        try {
            console.log(`[${this.MODULE_NAME}] 开始生成调息表，源工作表：${源工作表名}`);
            
            // 1. 复制源工作表到新工作表
            const newSheet = this.复制工作表(源工作表名);
            if (!newSheet) {
                throw new Error("复制工作表失败");
            }
            
            const newSheetName = newSheet.Name;
            console.log(`[${this.MODULE_NAME}] 新工作表：${newSheetName}`);
            
            // 2. 初始化到新工作表
            this.Initialize(newSheetName);
            
            // 3. 启用调息功能
            this.InitializeAdjustment({
                isEnabled: true,
                adjustmentType: "固定调整",
                adjustmentBasis: "基准利率"
            });
            
            // 4. 清除原有数据
            this.清除原有表中数据();
            
            // 5. 批量添加调息节点
            if (Array.isArray(调息数组) && 调息数组.length > 0) {
                this.BatchAddAdjustments(调息数组);
                console.log(`[${this.MODULE_NAME}] 已添加 ${调息数组.length} 个调息节点`);
            }
            
            // 6. 创建表头（包括第13列）
            this.创建租金测算表表头(1, 13);
            
            // 7. 生成数据区域（带调息功能）
            this.createDataRange();
            
            // 8. 输出调息情况
            this.调息情况();
            
            console.log(`[${this.MODULE_NAME}] 调息表生成完成：${newSheetName}`);
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 生成调息表失败：${error.message}`);
            return false;
        }
    }
    调息情况(){
        console.log("调息情况：", this.adjustmentPeriods);
        switch(this.adjustmentConfig.adjustmentType){
            case "固定调整":
                console.log("固定调整方式");
                break;
            case "浮动调整":
                console.log("浮动调整方式");
                break;
            case "自定义":
                console.log("自定义调整方式");
                break;
            default:
                console.log("未知调整方式");
        }

    }
    
    /**
     * 将数组数据写入Excel工作表
     * @param {Array} arrData - 二维数组数据
     * @param {Number} startRow - 开始行号
     * @param {Number} startCol - 开始列号
     * @returns {Boolean} 是否写入成功
     */
    writeArrDataToSheet(arrData, startRow, startCol) {
        try {
            // 参数验证
            if (!Array.isArray(arrData)) {
                throw new Error("数据必须是数组格式");
            }
            
            if (typeof startRow !== 'number' || startRow < 1) {
                throw new Error("起始行号必须是正整数");
            }
            
            if (typeof startCol !== 'number' || startCol < 1) {
                throw new Error("起始列号必须是正整数");
            }
            
            // 检查参数管理器是否已初始化
            if (!this.p || !this.p.m_worksheet) {
                throw new Error("参数管理器未初始化或工作表不可用");
            }
            
            // 获取工作表对象
            const worksheet = this.p.m_worksheet;
            
            // 验证数组维度
            const rowCount = arrData.length;
            if (rowCount === 0) {
                console.log("警告：数组为空，无需写入");
                return true;
            }
            
            const colCount = Array.isArray(arrData[0]) ? arrData[0].length : 1;
            
            // 确定目标范围
            const startCell = worksheet.Cells(startRow, startCol);
            const endRow = startRow + rowCount - 1;
            const endCol = startCol + colCount - 1;
            const targetRange = worksheet.Range(startCell, worksheet.Cells(endRow, endCol));
            
            // 写入数据
            targetRange.Value2 = arrData;
            
            console.log(`[${this.MODULE_NAME}] 成功将 ${rowCount}×${colCount} 数组数据写入工作表，起始位置: R${startRow}C${startCol}`);
            return true;
            
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 写入数组数据到工作表失败：${error.message}`);
            return false;
        }
    }
}


// 示例1：创建带调息功能的租金测算表
function 生成带调息的租金测算表() {
    // 创建实例
    const calc = new cls调息();
    
    // 初始化
    calc.Initialize();
    
    // 启用调息功能
    calc.InitializeAdjustment({
        isEnabled: true,
        adjustmentType: "固定调整",
        adjustmentBasis: "基准利率"
    });
    
    // 添加调息节点
    // 第5期开始利率调整为 4.5%
    calc.AddAdjustmentPeriod(5, 0.045);
    
    // 第10期开始利率调整为 5.2%
    calc.AddAdjustmentPeriod(10, 0.052);
    
    // 生成表格
    calc.清除原有表中数据();
    calc.创建租金测算表表头();
    calc.createDataRange();
    
    console.log("带调息功能的租金测算表生成完成");
}

// 示例2：批量添加调息节点
function 批量设置调息() {
    const calc = new cls调息();
    calc.Initialize();
    
    calc.InitializeAdjustment({ isEnabled: true });
    
    // 批量添加
    calc.BatchAddAdjustments([
        { period: 3, newRate: 0.048 },
        { period: 6, newRate: 0.051 },
        { period: 9, newRate: 0.055 }
    ]);
    
    calc.createDataRange();
}

// 示例3：从工作表读取调息配置
function 从表格加载调息配置() {
    const calc = new cls调息();
    calc.Initialize();
    
    calc.InitializeAdjustment({ isEnabled: true });
    
    // 假设 P2:Q10 区域存储了调息配置
    // P列：期次，Q列：新利率
    calc.LoadAdjustmentsFromSheet("P2:Q10");
    
    calc.createDataRange();
}
function 等额租金法调息(){
	r = new cls调息();
	r.Initialize();
	r.清除原有表中数据();
	r.创建租金测算表表头(1, 10);
	r.createDataRange();
	let arr2D = r.写入每期利率();
	r.每期适用利率();

	let arrFormula= r.等额租金法arr();//存储的公式二维数组
	let arrData = r.arrToArrData(arrFormula);//扩展数据区域的二维数组
	//logjson(arrData)
	//租金部分调整，第3列。
	//
	//
	//改变第2列和第3列的公式
    let periodChgStart = r.adjustmentConfig.periodChgStart;
    arr2D=r.调整期利率(arr2D, 0.025, periodChgStart);
	let str1 = `=ROUND(-PMT(RC[10]/${r.p.PaymentsPerYearCellR1C1},${r.p.TotalPeriodsCellR1C1}-${periodChgStart},R${r.p.RentTableStartRow+periodChgStart-1}C7,0),2)`; // 租金
	let str2 = `=ROUND(-PPMT(RC[9]/${r.p.PaymentsPerYearCellR1C1},RC[-3]-${periodChgStart},${r.p.TotalPeriodsCellR1C1}-${periodChgStart},R${r.p.RentTableStartRow+periodChgStart-1}C7,0),2)`; // 本金;
	let str3 =`0.025`
	//r.arrDataRewriteCol(arrData, 3, str2);
	
	let col3 = arrData.map((row)=>row[2]);
    let updatedCol3 = col3.map((item, index) => {
        if (index + 1 > periodChgStart) {
            // 调整后期次
            return str1;
        };
        return item
     });
// 将修改后的列数据写回arrData
	updatedCol3.forEach((value, index) => {
	    arrData[index][2] = value; // 更新每行的第3列（索引为2）
	});
	
	
	
	let col4 = arrData.map((row)=>row[3]);

    let updatedCol4 = col4.map((item, index) => {
        if (index + 1 > periodChgStart && index+1<r.p.TotalPeriodsCellValue) {
            // 调整后期次
            return str2;
        };
        return item
     });
     // 将修改后的列数据写回arrData
	updatedCol4.forEach((value, index) => {
	    arrData[index][3] = value; // 更新每行的第3列（索引为2）
	});
     
    let col13 = arrData.map((row)=>row[12]);
    let updatedCol13 = col13.map((item, index) => {
        if (index + 1 <= periodChgStart) {
            // 调整后期次
            return r.p.InterestRateCellValue;
        }else{
        	return str3;
        };
        
     });
// 将修改后的列数据写回arrData
	updatedCol13.forEach((value, index) => {
	    arrData[index][12] = value; // 更新每行的第3列（索引为2）
	});

	let sheet = r.p.m_sourceSheet;
	let RentTableStartRow = r.p.RentTableStartRow;
	let arrHeaders = r.arrHeaders;
    //r.writeArrDataToSheet(arrData, sheet, RentTableStartRow, arrHeaders);
    r.创建租金测算表表头(13, 13);
	r.writeArrDataToSheet(arrData, RentTableStartRow, 1);

}

/**
 * 测试 writeArrDataToSheet 方法
 */
function 测试writeArrDataToSheet() {
    try {
        console.log("开始测试 writeArrDataToSheet 方法...");
        
        // 创建实例
        const calc = new cls调息();
        calc.Initialize();
        
        // 准备测试数据
        const testData = [
            ["姓名", "年龄", "城市"],
            ["张三", 25, "北京"],
            ["李四", 30, "上海"],
            ["王五", 28, "广州"]
        ];
        
        // 测试写入数据
        const result = calc.writeArrDataToSheet(testData, 1, 1);
        
        if (result) {
            console.log("✅ writeArrDataToSheet 测试成功！数据已写入工作表");
        } else {
            console.log("❌ writeArrDataToSheet 测试失败！");
        }
        
        return result;
    } catch (error) {
        console.log(`❌ writeArrDataToSheet 测试异常：${error.message}`);
        return false;
    }
}
