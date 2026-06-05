/**
 * ============== 租前期租金测算类 ==============
 * 作者：徐晓冬
 * 描述：继承自clsRentalCalculation，专门处理带租前期的租金测算
 * ====================================================
 */

/**
 * 租前期租金测算类
 * @description 继承clsRentalCalculation类，扩展支持租前期功能
 *              租前期：从放款日到起租日之间的期间
 *              租前息：租前期产生的利息，在第1期租金支付日一并支付
 */
class clsPreLease extends clsRentalCalculation {
	/**
	 * 构造函数
	 * @param {Object} parameterManager - 参数管理器实例（可选）
	 */
	constructor(parameterManager) {
		super(parameterManager); // 传递参数管理器给父类
		this.MODULE_NAME = "clsPreLease";
		this.preLeaseRowCount = 0; // 租前期行数
		this.preLeaseRemainderInterest = 0; // 租前息（零头部分）
		console.log(`[${this.MODULE_NAME}] 类实例创建（继承自clsRentalCalculation）`);
	}

	// ============== 租前期辅助方法 ==============

	/**
	 * 计算租前期利息支付行（生成公式）
	 * @description 生成租前期利息支付公式数组，按天计息
	 * @returns {Array} 租前期支付数据数组，包含公式和值
	 */
	计算租前期支付数据() {
		try {
			// 获取参数
			const loanDate = this.p.LeaseStartDateCellValue; // 放款日
			const firstPaymentDate = this.p.FirstPaymentDateCellValue; // 起租日（第1期支付日）
			const preLeaseInterval = this.p.PreLeaseIntervalCellValue; // 租前期间隔（月）
			const preLeaseMonths = this.p.PreLeaseMonthsCellValue; // 租前期月数
			const principal = this.p.principalCellValue; // 租赁成本
			const annualRate = this.p.InterestRateCellValue; // 年利率

			console.log(`[${this.MODULE_NAME}] 开始计算租前期支付...`);
			console.log(`  放款日: ${loanDate}`);
			console.log(`  起租日: ${firstPaymentDate}`);
			console.log(`  租前期月数: ${preLeaseMonths}`);
			console.log(`  租前期间隔: ${preLeaseInterval}个月`);

			// 计算完整支付次数和零头月数
			const fullPayments = Math.floor(preLeaseMonths / preLeaseInterval);
			const remainderMonths = preLeaseMonths % preLeaseInterval;

			console.log(`  完整支付次数: ${fullPayments}`);
			console.log(`  零头月数: ${remainderMonths}`);

			const preLeaseRows = [];

			// 获取单元格地址
			const principalCell = this.p.PrincipalCellR1C1; // 本金单元格（R4C2）
			const rateCell = this.p.InterestRateCellR1C1; // 利率单元格
			const loanDateCell = this.p.LeaseStartDateCellR1C1; // 放款日单元格
			const preLeaseIntervalCell = this.p.PreLeaseIntervalCellR1C1; // 租前期间隔单元格

			// 生成租前期完整支付行
			for (var i = 1; i <= fullPayments; i++) {
				const row = new Array(13);
				row[0] = `租前${i}`; // 期次（静态值）

				// 支付日：公式 =EDATE(放款日, i * 租前期间隔) 或 =EDATE(上一行支付日, 租前期间隔)
				if (i === 1) {
					row[1] = `=EDATE(${loanDateCell}, ${preLeaseIntervalCell})`; // 第一期支付日
				} else {
					row[1] = `=EDATE(R[-1]C, ${preLeaseIntervalCell})`; // 后续支付日
				}

				// 租金：引用利息列 =RC[2]
				row[2] = "=RC[2]";

				// 本金：固定为0
				row[3] = 0;

				// 利息：公式 =ROUND(本金*M列利率/360*DAYS(支付日,上一支付日或放款日),2)
				if (i === 1) {
					// 第一期：从放款日开始计算
					row[4] = `=ROUND(${principalCell}*RC[8]/360*(RC[-3]-${loanDateCell}),2)`;
				} else {
					// 后续期：从上一支付日开始计算
					row[4] = `=ROUND(${principalCell}*RC[8]/360*(RC[-3]-R[-1]C[-3]),2)`;
				}

				// 累积偿还本金额：固定为0
				row[5] = 0;

				// 租金本金余额：引用本金单元格
				row[6] = `=${principalCell}`;

				// 剩余租金余额：空值
				row[7] = "";

				// 已还租金：引用租金列 =RC[-6]
				row[8] = "=RC[-6]";

				// 支付日/月间隔：=DATEDIF(上一支付日或放款日, 当前支付日, "M")
				if (i === 1) {
					row[9] = `=DATEDIF(${loanDateCell}, RC[-8], "M")`;
				} else {
					row[9] = '=DATEDIF(R[-1]C[-8], RC[-8], "M")';
				}

				// 自定义间隔：空值
				row[10] = "";

				// 本金比例：空值
				row[11] = "";

				// 每期适用利率：引用年利率单元格
				row[12] = `=${rateCell}`;

				preLeaseRows.push(row);
			}

			// 计算零头利息（将并入第1期租金）
			var remainderInterest = 0;
			if (remainderMonths > 0) {
				// 零头利息计算（用于备注）
				// 从最后完整支付日到起租日的天数
				const lastPaymentDate = this.addMonths(loanDate, fullPayments * preLeaseInterval);
				const days = this.daysBetween(lastPaymentDate, firstPaymentDate);
				remainderInterest = Math.round(principal * annualRate / 360 * days * 100) / 100;
				console.log(`  租前息（零头）: ${remainderInterest}元 (${days}天)`);
			} else {
				console.log(`  租前息（零头）: 0元（租前期已付清）`);
			}

			// 存储零头利息供后续使用
			this.preLeaseRemainderInterest = remainderInterest;
			this.preLeaseRowCount = preLeaseRows.length;

			console.log(`[${this.MODULE_NAME}] 租前期支付计算完成，生成${preLeaseRows.length}行数据`);
			return preLeaseRows;

		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租前期支付计算失败：${error.message}`);
			return [];
		}
	}

	/**
	 * 日期加月数
	 * @param {Date|string} date - 起始日期
	 * @param {number} months - 要增加的月数
	 * @returns {string} 格式化后的日期字符串 (YYYY-MM-DD)
	 */
	addMonths(date, months) {
		date = safeFromExcelDate(date);  // Excel OA数值 → JS Date
		const d = new Date(date);
		d.setMonth(d.getMonth() + months);
		const year = d.getFullYear();
		const month = String(d.getMonth() + 1).padStart(2, '0');
		const day = String(d.getDate()).padStart(2, '0');
		return `${year}-${month}-${day}`;
	}

	/**
	 * 计算两个日期之间的天数差
	 * @param {Date|string} startDate - 开始日期
	 * @param {Date|string} endDate - 结束日期
	 * @returns {number} 天数差
	 */
	daysBetween(startDate, endDate) {
		try {
			// 修复：若参数为Excel数值，先转JS Date
			startDate = safeFromExcelDate(startDate);
			endDate = safeFromExcelDate(endDate);
			const start = new Date(startDate);
			const end = new Date(endDate);

			// 验证日期有效性
			if (isNaN(start.getTime()) || isNaN(end.getTime())) {
				console.log(`[${this.MODULE_NAME}] 无效的日期: startDate=${startDate}, endDate=${endDate}`);
				return 0;
			}

			// 重置为当天的0点，避免时间部分影响计算
			start.setHours(0, 0, 0, 0);
			end.setHours(0, 0, 0, 0);

			// 计算差值（可能为负数）
			const diffTime = end - start;
			const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));

			// 返回绝对值
			return Math.abs(diffDays);
		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 日期计算失败: ${error.message}`);
			return 0;
		}
	}

	/**
	 * 写入租前期支付数据到工作表（支持公式）
	 * @param {Array} preLeaseData - 租前期支付数据数组，包含公式和值
	 * @returns {boolean} 是否成功
	 */
	写入租前期数据(preLeaseData) {
		try {
			if (!preLeaseData || preLeaseData.length === 0) {
				console.log(`[${this.MODULE_NAME}] 无租前期数据需要写入`);
				return true;
			}

			const ws = this.p.m_worksheet;
			const startRow = this.p.RentTableStartRow;

			// 逐行写入数据
			for (var i = 0; i < preLeaseData.length; i++) {
				const rowData = preLeaseData[i];
				const currentRow = startRow + i;

				// 逐列处理数据
				for (var col = 0; col < rowData.length; col++) {
					const cellValue = rowData[col];
					if (cellValue === "" || cellValue === null || cellValue === undefined) {
						continue; // 跳过空值
					}

					// 计算Excel列字母 (0=A, 1=B, ..., 12=M)
					const colLetter = String.fromCharCode(65 + col); // A=65
					const cellAddress = `${colLetter}${currentRow}`;
					const cell = ws.Range(cellAddress);

					// 判断是否为公式
					if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
						// 写入公式（使用R1C1格式）
						cell.FormulaR1C1 = cellValue;
					} else {
						// 写入值
						cell.Value2 = cellValue;
					}
				}
			}

			// 获取数据区域范围用于格式设置
			const dataRange = ws.Range(
				`${this.p.m_COL_PERIOD}${startRow}:${this.p.m_COL_PRINCIPAL_RATIO}${startRow + preLeaseData.length - 1}`
			);

			// 应用格式
			应用格式(dataRange.Columns(1), "Text"); // 期次列为文本
			应用格式(dataRange.Columns(2), "Date"); // 日期列
			for (var i = 3; i <= 5; i++) {
				应用格式(dataRange.Columns(i), "Standard"); // 金额列
			}
			应用格式(dataRange.Columns(13), "Percentage"); // 利率列

			// 添加边框
			this.添加框线(dataRange);

			// 设置租前期行的背景色
			const preLeaseRange = ws.Range(
				`A${startRow}:M${startRow + preLeaseData.length - 1}`
			);
			preLeaseRange.Interior.Color = this.p.m_COLOR_LIGHT_YELLOW || RGB(255, 255, 153);

			console.log(`[${this.MODULE_NAME}] 租前期数据写入完成，共${preLeaseData.length}行`);
			return true;

		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租前期数据写入失败：${error.message}`);
			return false;
		}
	}

	/**
	 * 处理第1期租金的租前息
	 * @description 将租前息（零头）添加到第1期租金中
	 */
	处理第1期租前息() {
		try {
			const ws = this.p.m_worksheet;
			const startRow = this.p.RentTableStartRow + this.preLeaseRowCount;
			const rentCol = this.p.m_COL_RENT; // C列
			const principalCol = this.p.m_COL_PRINCIPAL; // D列
			const interestCol = this.p.m_COL_INTEREST; // E列
			const remarksCol = "N"; // 备注列

			// 第1期租金单元格
			const firstPeriodRentCell = ws.Range(`${rentCol}${startRow}`);
			const firstPeriodPrincipalCell = ws.Range(`${principalCol}${startRow}`);
			const firstPeriodInterestCell = ws.Range(`${interestCol}${startRow}`);

			// 当前第1期租金 = 本金 + 利息（利息为0，因为起租日就是支付日）
			// 新的第1期租金 = 本金 + 租前息
			const principal = firstPeriodPrincipalCell.Value2;
			const preLeaseInterest = this.preLeaseRemainderInterest;

			// 修改第1期租金
			firstPeriodRentCell.Formula = `=${principalCol}${startRow}+${preLeaseInterest}`;
			firstPeriodInterestCell.Value2 = 0;

			// 添加备注
			const remarksCell = ws.Range(`${remarksCol}${startRow}`);
			const currentRemarks = remarksCell.Value2 || "";
			remarksCell.Value2 = `含租前息${preLeaseInterest}元`;

			console.log(`[${this.MODULE_NAME}] 第1期租前息处理完成，租前息=${preLeaseInterest}元`);
			return true;

		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 第1期租前息处理失败：${error.message}`);
			return false;
		}
	}

	/**
	 * 应用租前期行的格式
	 * @description 为租前期行设置背景色和格式
	 */
	应用租前期格式() {
		try {
			if (this.preLeaseRowCount === 0) {
				return true;
			}

			const ws = this.p.m_worksheet;
			const startRow = this.p.RentTableStartRow;
			const endRow = startRow + this.preLeaseRowCount - 1;

			// 租前期行区域
			const preLeaseRange = ws.Range(`A${startRow}:M${endRow}`);

			// 应用格式
			this.应用格式(preLeaseRange.Columns(1), "Text"); // 期次列为文本
			this.应用格式(preLeaseRange.Columns(2), "Date"); // 日期列
			for (var i = 3; i <= 5; i++) {
				this.应用格式(preLeaseRange.Columns(i), "Standard"); // 金额列
			}
			this.应用格式(preLeaseRange.Columns(13), "Percentage"); // 利率列

			// 添加边框
			this.添加框线(preLeaseRange);

			// 设置租前期行的背景色
			preLeaseRange.Interior.Color = this.p.m_COLOR_LIGHT_YELLOW || RGB(255, 255, 153);

			console.log(`[${this.MODULE_NAME}] 租前期格式应用完成：行${startRow}-${endRow}`);
			return true;

		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租前期格式应用失败：${error.message}`);
			return false;
		}
	}

	/**
	 * 生成带租前期的租金测算表（在二维数组阶段合并租前期和正常期数据）
	 * @returns {boolean} 是否成功
	 */
	生成租前期租金测算表() {
		try {
			console.log(`[${this.MODULE_NAME}] 开始生成租前期租金测算表...`);

			// 1. 创建表头
			this.创建租金测算表表头();

			// 2. 生成完整数据（租前期 + 正常期）并一次性写入
			const result = this.createDataRange();

			// 3. 处理第1期租金的租前息（如果有零头）
			if (this.preLeaseRemainderInterest > 0) {
				this.处理第1期租前息();
			}

			// 4. 应用格式到租前期行
			if (this.preLeaseRowCount > 0) {
				this.应用租前期格式();
			}

			console.log(`[${this.MODULE_NAME}] 租前期租金测算表生成完成`);
			return true;

		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租前期租金测算表生成失败：${error.message}`);
			return false;
		}
	}

	/**
	 * RGB颜色转换（已弃用，请使用 mShared_constants.js 中的全局 RGB() 函数）
	 * @deprecated 保留此方法仅为向后兼容，新代码请直接调用 RGB(r, g, b)
	 */
	RGB(r, g, b) {
		return RGB(r, g, b);
	}

	// ============== 重写的数据处理方法（支持租前期模式） ==============

	/**
	 * 重写arrToArrData方法（在二维数组阶段合并租前期和正常期数据）
	 * @param {Array} arrFormula - 正常期的公式数组
	 * @returns {Array} 合并后的完整数据数组（租前期 + 正常期）
	 */
	arrToArrData(arrFormula) {
		try {
			// 1. 生成租前期数据（如果存在）
			const preLeaseData = this.计算租前期支付数据();
			this.preLeaseRowCount = preLeaseData.length;

			// 2. 生成正常期数据
			const totalPeriods = this.p.TotalPeriodsCellValue;
			var normalPeriodData = [];
			for (var i = 0; i < totalPeriods; i++) {
				normalPeriodData[i] = new Array(arrFormula[1].length);
			}
			const maxRow = arrFormula.length - 1;
			const maxCol = arrFormula[1].length - 1;

			for (var row = 0; row < totalPeriods; row++) {
				var rowIndex = null;
				for (var col = 1; col <= maxCol && col < arrFormula[1].length; col++) {
					if (row === 0) {
						rowIndex = 1; // 首行公式
						normalPeriodData[row][col - 1] = arrFormula[rowIndex][col];
					} else if (row >= 1 && row < totalPeriods - 1) {
						rowIndex = 2; // 中间行公式
						normalPeriodData[row][col - 1] = arrFormula[rowIndex][col];
					} else if (row === totalPeriods - 1) {
						rowIndex = 3; // 最后一行公式
						normalPeriodData[row][col - 1] = arrFormula[rowIndex][col];
					}
				}
			}

			// 3. 合并租前期数据和正常期数据
			const combinedData = [];
			// 先添加租前期数据
			for (var i = 0; i < preLeaseData.length; i++) {
				combinedData.push(preLeaseData[i]);
			}
			// 再添加正常期数据
			for (var i = 0; i < normalPeriodData.length; i++) {
				combinedData.push(normalPeriodData[i]);
			}

			console.log(`[${this.MODULE_NAME}] 数据数组合并完成：租前期${preLeaseData.length}行 + 正常期${normalPeriodData.length}行 = ${combinedData.length}行`);
			return combinedData;
		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租金测算表数据数组转换失败：${error.message}`);
			return null;
		}
	}

	/**
	 * 重写arrDataToDataRange方法（处理合并后的租前期+正常期数据）
	 * @param {Array} arrData - 合并后的完整数据数组（租前期 + 正常期）
	 * @returns {Object} 写入的数据范围
	 */
	arrDataToDataRange(arrData) {
		try {
			if (arrData === null) {
				throw new Error("数据数组为空");
			}

			// 数据从表头行开始写入（包含租前期和正常期）
			const startRow = this.p.RentTableStartRow;
			const totalRows = arrData.length;

			// 数据区域范围：从表头行开始写入所有行
			const rngData = this.p.m_worksheet.Range(
				`${this.p.m_COL_PERIOD}${startRow}:${this.p.m_COL_PRINCIPAL_RATIO}${startRow + totalRows - 1}`
			);

			// 逐行写入数据（支持公式和值混合）
			for (var i = 0; i < arrData.length; i++) {
				const rowData = arrData[i];
				const currentRow = startRow + i;

				// 逐列处理数据
				for (var col = 0; col < rowData.length; col++) {
					const cellValue = rowData[col];
					if (cellValue === "" || cellValue === null || cellValue === undefined) {
						continue; // 跳过空值
					}

					// 计算Excel列字母 (0=A, 1=B, ..., 12=M)
					const colLetter = String.fromCharCode(65 + col); // A=65
					const cellAddress = `${colLetter}${currentRow}`;
					const cell = this.p.m_worksheet.Range(cellAddress);

					// 判断是否为公式
					if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
						// 写入公式（使用R1C1格式）
						cell.FormulaR1C1 = cellValue;
					} else {
						// 写入值
						cell.Value2 = cellValue;
					}
				}
			}

			console.log(`[${this.MODULE_NAME}] 数据数组写入表格完成：${totalRows}行数据写入到行${startRow}-${startRow + totalRows - 1}`);
			return rngData;
		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租金测算表数据数组写入表格失败：${error.message}`);
			return null;
		}
	}

	/**
	 * 重写createDataRange方法（支持租前期模式）
	 * 修改合计行的生成位置
	 */
	createDataRange() {
		try {
			var rng = null;
			var arrFormula = [];
			var arrData = [];

			// 检查是否已初始化
			if (this.p === null || !this.p.IsInitialized) {
				this.Initialize();
			}

			// 验证wsTarget是否已正确初始化
			if (this.WsTarget === null) {
				throw new Error("目标工作表未初始化");
			}

			// 合计行的位置 = 表头行 + 租前期行数 + 正常期次数
			const totalRowPosition = this.p.RentTableStartRow + this.preLeaseRowCount + this.p.TotalPeriodsCellValue;

			// 使用ParameterManager获取还款方式
			const repaymentMethod = this.WsTarget.Range(this.p.RepaymentMethodCellA1).Value2;
			const r1 = this.p.m_worksheet.Range(`${this.p.m_COL_PRINCIPAL_RATIO}${totalRowPosition}`);

			// 设置公式生成器参数化模式（直租模块使用M列每期适用利率 + 租前期模式）
			this.formulaGenerator.setRateReferenceMode('column');
			this.formulaGenerator.setPreLeaseConfig({
				enabled: this.preLeaseRowCount > 0,
				preLeaseRowCount: this.preLeaseRowCount,
				firstPaymentDateRef: this.p.FirstPaymentDateCellR1C1,
				preLeaseMonthsRef: this.p.PreLeaseMonthsCellR1C1
			});

			switch (repaymentMethod) {
				case "等额本息（后付）":
					arrFormula = this.formulaGenerator.generateEqualPaymentFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2]);
					this.添加框线(rng.Columns("A:J"));
					break;
				case "等额本金（按天计息）":
					arrFormula = this.formulaGenerator.generateEqualPrincipalDailyInterestFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2, 4], [12]);
					this.添加框线(rng.Columns("A:J"));
					break;
				case "等额本金（按期计息）":
					arrFormula = this.formulaGenerator.generateEqualPrincipalPeriodicInterestFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2, 4], [12]);
					this.添加框线(rng.Columns("A:J"));
					break;
				case "本金比例（按期计息）":
					arrFormula = this.formulaGenerator.generatePrincipalRatioPeriodicInterestFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2]);
					this.创建租金测算表表头(12, 12);
					this.添加框线(rng.Columns("A:J"));
					this.添加框线(rng.Columns("L:L"));
					this.设置背景颜色(rng.Columns("L:L"), this.p.m_COLOR_YELLOW);
					this.租金测算表合计行(12, 12);
					this.设置背景颜色(r1, this.p.m_COLOR_WHITE);
					break;
				case "本金比例（按天计息）":
					arrFormula = this.formulaGenerator.generatePrincipalRatioDailyInterestFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2]);
					this.创建租金测算表表头(12, 12);
					this.添加框线(rng.Columns("A:J"));
					this.添加框线(rng.Columns("L:L"));
					this.设置背景颜色(rng.Columns("L:L"), this.p.m_COLOR_YELLOW);
					this.租金测算表合计行(12, 12);
					this.设置背景颜色(r1, this.p.m_COLOR_WHITE);
					break;
				case "等额本息（先付）":
					arrFormula = this.formulaGenerator.generateEqualPaymentAdvanceFormulas();
					arrData = this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1, 2]);
					this.添加框线(rng.Columns("A:J"));
					break;
				default:
					throw new Error(`不支持的还款方式：${repaymentMethod}`);
			}

			// 应用格式和生成合计行
			this.租金测算表应用格式(rng);
			this.租金测算表合计行(1, 10);

			console.log(`[${this.MODULE_NAME}] 租金测算表生成完成`);
			return true;
		} catch (error) {
			// P1-8: 集成统一错误处理器
			if (typeof g_errorHandler !== 'undefined') {
				g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '租金测算表生成' });
			} else {
				console.error(`[${this.MODULE_NAME}] 租金测算表生成失败：${error.message}`);
			}
			return false;
		}
	}



// ============== 公式生成委托 clsFormulaGenerator（已去重 V2.1） ==============
// 等额租金法arr / 等额本息先付arr / 等额本金法按天计息arr / 本金比例法按天计息arr / 本金比例法按期计息arr
// 已委托为 this.formulaGenerator.generate*() + setPreLeaseConfig() + setRateReferenceMode("column")
// 见 createDataRange() switch-case

/**
 * 重写租金测算表合计行方法（支持租前期模式）
 * 在租前期模式下，合计行位置需要加上租前期行数
 */
	租金测算表合计行(startCol = 1, lastCol = 12) {
		try {
			// 租前期模式下：合计行位置 = 表头行 + 租前期行数 + 正常期次数
			const lastRow = this.p.RentTableStartRow + this.preLeaseRowCount + this.p.TotalPeriodsCellValue;

			// 固定该行覆盖A~L，便于使用 Columns(i) 与原有列位一致（1~12）
			const range = this.p.m_worksheet.Range(`A${lastRow}:L${lastRow}`);

			// R1C1 形式的区间求和：从"距本行 TotalPeriodsCellValue 行之上"到"上一行"
			const sumFormula = `=SUM(R[${-this.p.TotalPeriodsCellValue}]C:R[-1]C)`;

			// 按列处理规则：
			// 1: "合计"
			// 2: "-"
			// 3~5: 求和
			// 6~9: "-"
			// 10~12: 求和
			for (var i = 1; i <= 12; i++) {
				// 仅在指定列区间内处理
				if (i < startCol || i > lastCol) continue;

				if (i === 1) {
					range.Columns(1).Value2 = "合计";
					this.添加框线(range.Columns(1));
				} else if (i === 2) {
					range.Columns(2).Value2 = "-";
					this.添加框线(range.Columns(2));
				} else if ((i >= 3 && i <= 5) || (i >= 10 && i <= 12)) {
					range.Columns(i).FormulaR1C1 = sumFormula;
					this.添加框线(range.Columns(i));
				} else if (i >= 6 && i <= 9) {
					range.Columns(i).Value2 = "-";
					this.添加框线(range.Columns(i));
				}
			}

			return true;
		} catch (error) {
			// P1-8: 集成统一错误处理器
			if (typeof g_errorHandler !== 'undefined') {
				g_errorHandler.handleError(error, ERROR_CODES.CALCULATION_ERROR, { module: this.MODULE_NAME, function: '租金测算表合计行' });
			} else {
				console.error(`[${this.MODULE_NAME}] 租金测算表合计行生成失败：${error.message}`);
			}
			return false;
		}
	}
}

// ============== 快捷调用函数 ==============

/**
 * 生成带租前期的租金测算表
 */
function 生成租前期租金测算表() {
	try {
		console.log("=== 开始生成租前期租金测算表 ===");
		const r = new clsPreLease();
		r.Initialize();
		r.生成租前期租金测算表();
		return true;
	} catch (error) {
		console.log(`租前期租金测算表生成：${error.message}`);
		return false;
	}
}
