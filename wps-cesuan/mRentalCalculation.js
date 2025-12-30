/**
 * ============== 租金测算系统主文件 ==============
 * 作者：徐晓冬
 * 描述：整合所有租金测算功能的主模块
 * ====================================================
 */
 const p = new ParameterManager();
 p.Initialize();
 
 class RentalCalculation{
	constructor(){
		this.MODULE_NAME = "RentalCalculation";	
		console.log("[" + this.MODULE_NAME + "] 类实例创建");
		this.WsTarget = null;
		this.m_staticValueConversion = false;
		this.arrHeaders = ["期次", "支付日", "租金", "本金", "利息", "累积偿还本金额",
	                        "租金本金余额", "剩余租金余额", "已还租金", "支付日/月间隔",
	                        "支付日/月间隔-自定义", "本金比例", "每期适用利率"];// 一维表头数组
		this.lenghHeader = this.arrHeaders.length;// 表头长度
	}
	Initialize(sheetName = "1租金测算表V1") {
    	try {
	        // 设置参数管理器

	        this.p = p
	        this.p.Initialize(sheetName)     
	        this.WsTarget = this.p.m_worksheet;
	       
	        
	        // 输出初始化结果
	        console.log(`[${this.MODULE_NAME}] 初始化完成 - 总期数:${this.p.TotalPeriodsCellValue}, 租赁成本:${this.p.principalCellValue}`);
	        console.log("------------------------");
	        return true;
	        
    } catch (error) {
	        const errMsg = `初始化失败：${error.message}`;
	        console.log("------------------------");
	        console.log(`[${this.MODULE_NAME}] ${errMsg}`);
	        alert(`初始化失败：${errMsg}`);
	        return false;
    }
	}
	创建租金测算表表头(startCol = 1, lastCol = 11) {
	    try {
	        
	        // 设置主标题
	        const titleCell = this.p.m_worksheet.Cells(this.p.RowStart, 1)
	        titleCell.Value2 = "租金测算表";
	        titleCell.Interior.Color = this.p.m_COLOR_WHITE;
	        titleCell.Font.Name = FONT_CHINESE;
	        titleCell.Font.Size = FONT_SIZE_TITLE;
	        titleCell.Font.Color = this.p.m_COLOR_BLACK;
	        titleCell.HorizontalAlignment = XL.HCenter;
	        
	        // 定义表头数组
	        const headers = this.arrHeaders
	        
	        // 动态计算列范围
	        //const lastCol = headers.length;
	        
			const headerRange = this.p.m_worksheet.Range(
			  this.p.m_worksheet.Cells(this.p.RowStart + 1, startCol),
			  this.p.m_worksheet.Cells(this.p.RowStart + 1, lastCol)
			);
			for (let j = 1; j <= (lastCol - startCol + 1); j++) {
			  headerRange.Cells(1, j).Value2 = headers[(startCol - 1) + (j - 1)];
			  const colRange = headerRange.Columns(j);
			  colRange.Interior.Color = this.p.m_COLOR_BLUE;
			  colRange.Font.Name = FONT_DEFAULT;
			  colRange.Font.Size = FONT_SIZE_HEADER;
			  colRange.Font.Color = this.p.m_COLOR_BLACK;
			  colRange.HorizontalAlignment = XL.HCenter;
			  colRange.VerticalAlignment = XL.VCenter;
			  colRange.WrapText = true;
			  colRange.Borders.LineStyle = xlContinuous;
			  colRange.Borders.Weight = xlThin;
			  colRange.Borders.Color = this.p.m_COLOR_GRAY;
			}
			

	        
	        console.log(`[${this.MODULE_NAME}] 表头写入成功`);
	        return true;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 租金测算表表头创建失败：${error.message}`);
	        return false;
	    }
	}
	等额租金法arr() {
	    try {	        
	        // 创建存储单元格公式的二维数组，R1C1样式
	        const arr = [];
	        for (let i = 0; i < 5; i++) {
	            arr[i] = new Array(13);
	        }
	        
	        // 等额租金法涉及的Fx公式
	        // 第1期每列（字段）公式
	        arr[1][1] = "1"; // 期次
	        arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`; // EDATE公式 生成第一期支付日期
	        arr[1][3] = `=ROUND(-PMT(${this.p.InterestRateCellR1C1}/R11C2,${this.p.TotalPeriodsCellR1C1},${this.p.PrincipalCellR1C1},0),2)`; // 租金
	        arr[1][4] = `=ROUND(-PPMT(${this.p.InterestRateCellR1C1}/R11C2,RC[-3],${this.p.TotalPeriodsCellR1C1},${this.p.PrincipalCellR1C1},0),2)`; // 本金
	        arr[1][5] = "=RC[-2]-RC[-1]"; // 利息
	        arr[1][6] = "=RC[-2]"; // 累积偿还本金额
	        arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`; // 租金本金余额
	        arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`; // 剩余租金余额
	        arr[1][9] = "=RC[-6]"; // 已还租金
	        arr[1][10] = '=DATEDIF(R13C2,RC[-8], "M")'; // 偿还日期月间隔
	        
	        // 第2期-倒数第2期每列（字段）公式
	        arr[2][1] = "=R[-1]C+1"; // 期次递增
	        arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellValue + ")"; // 支付日递增
	        arr[2][3] = arr[1][3]; // 租金公式相同
	        arr[2][4] = arr[1][4]; // 本金公式相同
	        arr[2][5] = arr[1][5]; // 利息公式相同
	        arr[2][6] = "=RC[-2] + R[-1]C"; // 累积偿还本金额累加
	        arr[2][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[2][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[2][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        
	        // 倒数第1期每列（字段）公式
	        arr[3][1] = arr[2][1];
	        arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")"; // 支付日
	        arr[3][3] = arr[1][3]
	        arr[3][4] = `=${this.p.PrincipalCellR1C1} - SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`; // 最后一期本金
	        arr[3][5] = arr[1][5];
	        arr[3][6] = arr[2][6];
	        arr[3][7] = arr[1][7];
	        arr[3][8] = arr[1][8];
	        arr[3][9] = "=RC[-6] + R[-1]C";
	        arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        
	        console.log(`[${this.MODULE_NAME}] 等额租金法公式数组生成成功`);
	        return arr;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 等额租金法公式数组生成失败：${error.message}`);
	        return null;
	    }
	}
	等额本息先付arr() {
	    try {
	        // 使用已初始化的全局变量
	        if (this.p === null) {
	            throw new Error("ParameterManager未初始化");
	        }
	        
	        // 创建存储单元格公式的二维数组，R1C1样式
	        const arr = [];
	        for (let i = 0; i < 5; i++) {
	            arr[i] = new Array(13);
	        }
	        
	        // 等额本息先付涉及公式
	        // 倒数第1期每列（字段）公式
	        arr[1][1] = "1"; // 期次
	        arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`; // EDATE公式 生成第一期支付日期
	        arr[1][3] = `=ROUND(-PPMT(${this.p.InterestRateCellR1C1}/${this.p.PaymentsPerYearCellR1C1},RC[-2],${this.p.TotalPeriodsCellR1C1},${this.p.PrincipalCellR1C1}+RC[2],,1),2)`;
	        arr[1][4] = "=RC[-1]-RC[1]"; // 本金
	        arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*${this.p.InterestRateCellR1C1}*(RC[-3] - ${this.p.LeaseStartDateCellR1C1})/360,2)`; // 利息（按天计息）
	        arr[1][6] = "=RC[-2]"; // 累积偿还本金额
	        arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`; // 租金本金余额
	        arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`; // 剩余租金余额
	        arr[1][9] = "=RC[-6]"; // 已还租金
	        arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`; // 偿还日期月间隔
	        
	        // 第2期-倒数第2期每列（字段）公式
	        arr[2][1] = "=R[-1]C+1"; // 期次递增
	        arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellValue + ")"; // 支付日递增
	        arr[2][3] = `=R${this.p.RentTableStartRow}C`; // 租金（引用第一期）
	        arr[2][4] = arr[1][4]; // 本金公式相同
	        arr[2][5] = `=R[-1]C[2]*${this.p.InterestRateCellR1C1}/${this.p.PaymentsPerYearCellR1C1}`; // 利息（按期计息）
	        arr[2][6] = "=RC[-2] + R[-1]C"; // 累积偿还本金额累加
	        arr[2][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[2][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[2][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        
	        // 倒数第1期每列（字段）公式
	        arr[3][1] = arr[2][1]; // 期次公式相同
	        arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")"; // 支付日
	        arr[3][3] = arr[2][3]; // 租金公式相同
	        arr[3][4] = `=${this.p.PrincipalCellR1C1} - SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`; // 最后一期本金调整
	        arr[3][5] = arr[2][5]; // 利息公式相同
	        arr[3][6] = arr[2][6]; // 累积偿还本金额公式相同
	        arr[3][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[3][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[3][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        
	        console.log(`[${this.MODULE_NAME}] 等额本息先付公式数组生成成功`);
	        return arr;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 等额本息先付公式数组生成失败：${error.message}`);
	        return null;
	    }
	}
	等额本金法按天计息arr() {
	    try {
	        // 使用已初始化的全局变量
	        if (this.p === null) {
	            throw new Error("ParameterManager未初始化");
	        }
	        
	        // 创建存储单元格公式的二维数组，R1C1样式
	        const arr = [];
	        for (let i = 0; i < 5; i++) {
	            arr[i] = new Array(13);
	        }
	        
	        // 等额本金法的R1C1公式
	        arr[1][1] = "1"; // 第1列生成期次序号
	        arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`; // EDATE公式 生成第一期支付日期
	        arr[1][3] = "=ROUND(RC[1]+RC[2],2)"; // 租金RC= RC[1]=D列（本金），RC[2]=E列（利息）
	        arr[1][4] = `=ROUND(${this.p.PrincipalCellR1C1}/${this.p.TotalPeriodsCellR1C1},2)`; // 当期本金=总成本/总期数
	        arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*${this.p.InterestRateCellR1C1}/360*(RC[-3]-${this.p.LeaseStartDateCellR1C1}),2)`; // 利息
	        arr[1][6] = "=RC[-2]"; // 累积偿还本金额,还完第1期后的累积偿还本金额
	        arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`; // 租金本金余额，还完第1期后的租金本金余额
	        arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`; // 剩余租金余额
	        arr[1][9] = "=RC[-6]"; // 已还租金总额，等于第1期偿还的租金
	        arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`; // 得出偿还间隔/月
	        
	        // 第2期-倒数第2期每列（字段）公式
	        arr[2][1] = "=R[-1]C+1";
	        arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellValue + ")"; // 支付日递增
	        arr[2][3] = arr[1][3];
	        arr[2][4] = arr[1][4];
	        arr[2][5] = `=ROUND(R[-1]C[2]*${this.p.InterestRateCellR1C1}/360*(RC[-3]-R[-1]C[-3]),2)`; // 利息
	        arr[2][6] = "=RC[-2] + R[-1]C"; // 累积偿还本金额，本期偿还本金+上期累计数
	        arr[2][7] = arr[1][7];
	        arr[2][8] = arr[1][8];
	        arr[2][9] = "=RC[-6] + R[-1]C";
	        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
	        
	        // 倒数第1期每列（字段）公式
	        arr[3][1] = arr[2][1];
	        arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")"; // 支付日
	        arr[3][3] = arr[1][3];
	        arr[3][4] = `=${this.p.PrincipalCellR1C1} - SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`; // 最后一期本金调整
	        arr[3][5] = arr[2][5];
	        arr[3][6] = arr[2][6];
	        arr[3][7] = arr[1][7];
	        arr[3][8] = arr[1][8];
	        arr[3][9] = "=RC[-6] + R[-1]C";
	        arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")';
	        
	        return arr;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 等额本金法按天计息公式数组生成失败：${error.message}`);
	        return null;
	    }
	}
	本金比例法按天计息arr() {
	    try {
	        // 使用已初始化的全局变量
	        if (this.p === null) {
	            throw new Error("ParameterManager未初始化");
	        }
	        
	        // 创建存储单元格公式的二维数组，R1C1样式
	        const arr = [];
	        for (let i = 0; i < 5; i++) {
	            arr[i] = new Array(13);
	        }
	        
	        // 本金比例法按天计息的R1C1公式
	        // 第1期每列（字段）公式
	        arr[1][1] = "1"; // 第1列生成期次序号
	        arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`; // EDATE公式 生成第一期支付日期
	        arr[1][3] = "=ROUND(RC[1]+RC[2],2)"; // 租金 = 本金 + 利息
	        arr[1][4] = `=ROUND(${this.p.PrincipalCellR1C1}*RC[8]/100,2)`; // 当期本金=总成本*本金比例/100
	        arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*${this.p.InterestRateCellR1C1}/360*(RC2-${this.p.LeaseStartDateCellR1C1}),2)`; // 利息（按天计息）
	        arr[1][6] = "=RC[-2]"; // 累积偿还本金额
	        arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`; // 租金本金余额
	        arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`; // 剩余租金余额
	        arr[1][9] = "=RC[-6]"; // 已还租金总额
	        arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`; // 偿还日期月间隔
	        arr[1][12] = `=round(100/${this.p.TotalPeriodsCellR1C1},2)`; // 本金比例
	        
	        // 第2期-倒数第2期每列（字段）公式
	        arr[2][1] = "=R[-1]C+1"; // 期次递增
	        arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellValue + ")"; // 支付日递增
	        arr[2][3] = arr[1][3]; // 租金公式相同
	        arr[2][4] = arr[1][4]; // 本金公式相同
	        arr[2][5] = `=ROUND(R[-1]C[2]*${this.p.InterestRateCellR1C1}/360*(RC[-3]-R[-1]C[-3]),2)`; // 利息（按天计息）
	        arr[2][6] = "=RC[-2] + R[-1]C"; // 累积偿还本金额
	        arr[2][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[2][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[2][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        arr[2][12] = arr[1][12]; // 本金比例公式相同
	        
	        // 倒数第1期每列（字段）公式
	        arr[3][1] = arr[2][1];
	        arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")"; // 支付日
	        arr[3][3] = arr[1][3];
	        arr[3][4] = arr[1][4];
	        arr[3][5] = arr[2][5];
	        arr[3][6] = arr[2][6];
	        arr[3][7] = arr[1][7];
	        arr[3][8] = arr[1][8];
	        arr[3][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[3][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        arr[3][12] = `=100-SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`; // 最后一期本金比例调整
	        
	        console.log(`[${this.MODULE_NAME}] 本金比例法按天计息公式数组生成成功`);
	        return arr;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 本金比例法按天计息公式数组生成失败：${error.message}`);
	        return null;
	    }
	}
	本金比例法按期计息arr() {
	    try {
	        // 使用已初始化的全局变量
	        if (this.p === null) {
	            throw new Error("ParameterManager未初始化");
	        }
	        
	        // 创建存储单元格公式的二维数组，R1C1样式
	        const arr = [];
	        for (let i = 0; i < 5; i++) {
	            arr[i] = new Array(13);
	        }
	        
	        // 本金比例法按期计息的R1C1公式
	        // 第1期每列（字段）公式
	        arr[1][1] = "1"; // 第1列生成期次序号
	        //arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellR1C1})`; // EDATE公式 生成第一期支付日期
	        arr[1][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1}, ${this.p.PaymentIntervalCellValue})`; // EDATE公式 生成第一期支付日期
	        arr[1][3] = "=ROUND(RC[1]+RC[2],2)"; // 租金
	        arr[1][4] = `=ROUND(${this.p.PrincipalCellR1C1}*RC[8]/100,2)`; // 当期本金=总成本*本金比例/100
	        arr[1][5] = `=ROUND(${this.p.PrincipalCellR1C1}*${this.p.InterestRateCellR1C1}/${this.p.PaymentsPerYearCellR1C1},2)`; // 利息（按期计息）
	        arr[1][6] = "=RC[-2]"; // 累积偿还本金额
	        arr[1][7] = `=${this.p.PrincipalCellR1C1} - RC[-1]`; // 租金本金余额
	        arr[1][8] = `=SUM(R${this.p.RentTableStartRow}C3:R${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}C3) - RC[1]`; // 剩余租金余额
	        arr[1][9] = "=RC[-6]"; // 已还租金总额
	        arr[1][10] = `=DATEDIF(${this.p.LeaseStartDateCellR1C1},RC[-8], "M")`; // 得出偿还间隔/月
	        arr[1][12] = `=round(100/${this.p.TotalPeriodsCellR1C1},2)`; // 本金比例
	        
	        // 第2期-倒数第2期每列（字段）公式
	        arr[2][1] = "=R[-1]C+1"; // 期次递增
	        //arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellR1C1 + ")"; // 支付日递增
	        arr[2][2] = "=EDATE(R[-1]C, " +this.p.PaymentIntervalCellValue + ")"; // 支付日递增
	        arr[2][3] = arr[1][3]; // 租金公式相同
	        arr[2][4] = arr[1][4]; // 本金公式相同
	        arr[2][5] = `=ROUND(R[-1]C[2]*${this.p.InterestRateCellR1C1}/${this.p.PaymentsPerYearCellR1C1},2)`; // 利息（按期计息）
	        arr[2][6] = "=RC[-2] + R[-1]C"; // 累积偿还本金额
	        arr[2][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[2][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[2][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[2][10] = '=DATEDIF(R[-1]C[-8],RC[-8], "M")'; // 月间隔计算
	        arr[2][12] = arr[1][12]; // 本金比例公式相同
	        
	        // 倒数第1期每列（字段）公式
	        arr[3][1] = arr[2][1]; // 期次公式相同
	        //arr[3][2] = `=EDATE(${this.p.LeaseStartDateCellR1C1},${this.p.PaymentIntervalCellR1C1}*${this.p.TotalPeriodsCellR1C1})`; // 支付日
	       	arr[3][2] = "=EDATE(" + this.p.LeaseStartDateCellR1C1 + "," + this.p.PaymentIntervalCellValue + "*" + this.p.TotalPeriodsCellValue + ")"; // 支付日
	        arr[3][3] = arr[1][3]; // 租金公式相同
	        arr[3][4] = arr[1][4]; // 本金公式相同
	        arr[3][5] = arr[2][5]; // 利息（按期计息）
	        arr[3][6] = arr[2][6]; // 累积偿还本金额公式相同
	        arr[3][7] = arr[1][7]; // 租金本金余额公式相同
	        arr[3][8] = arr[1][8]; // 剩余租金余额公式相同
	        arr[3][9] = "=RC[-6] + R[-1]C"; // 已还租金累加
	        arr[3][10] = arr[2][10]; // 月间隔计算
	        arr[3][12] = `=100-SUM(R[${-this.p.TotalPeriodsCellValue + 1}]C:R[-1]C)`; // 最后一期本金比例调整
	        
	        console.log(`[${this.MODULE_NAME}] 本金比例法按期计息公式数组生成成功`);
	        return arr;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 本金比例法按期计息公式数组生成失败：${error.message}`);
	        return null;
	    }
	}
// 合计行：支持起始列与末列控制
	租金测算表合计行(startCol = 1, lastCol = 12) {
		  try {
		    const lastRow = this.p.RentTableStartRow + this.p.TotalPeriodsCellValue;
		
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
			    for (let i = 1; i <= 12; i++) {
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
			      } else {
			        // 兜底：未覆盖到的列，用"-"
			        range.Columns(i).Value2 = "-";
			        this.添加框线(range.Columns(i));
			      }
			    }
		
		    // 样式
		    range.Font.Bold = true;
		    range.NumberFormat = "#,##0.00";
            range.HorizontalAlignment = xlCenter;
            range.VerticalAlignment = xlCenter;
		    //this.添加框线(range);
		
		
		    return true;
		  } catch (error) {
		    console.log(`租金测算表合计行生成失败：${error.message}`);
		    alert(`租金测算表合计行生成失败：${error.message}`);
		    return false;
		  }
		}
	arrToArrData(arrFormula) {

		//arrFormula：存储公式的二维数组
		//arrData:表格中数据区域的二维数组
		//首先将arrFormula数组中的公式扩展arrData，然后再将arrData写入表格
		//

	    try {
			let arrData = [];// 创建存储单元格数据的二维数组
			// 初始化arrData数组，大小为总期数行 x 公式列数
			for (let i=0; i<this.p.TotalPeriodsCellValue; i++){
				arrData[i] = new Array(arrFormula[1].length);
			}
			const maxRow = arrFormula.length - 1; // 获取arrFormula行数
	        const maxCol = arrFormula[1].length - 1; // 获取arrFormula列数
	        
	        // 利用循环结构在单元格写入公式
	        // 修复：row从0开始，代表第1期（数组索引0）
	        // 修复：col从1开始，因为arrFormula数组从索引1开始填充公式（对应Excel列号）
	        // 关键修复：arrData是JavaScript数组（从0开始），arrFormula的col是Excel列号（从1开始）
	        // 所以需要将arrFormula[rowIndex][col]赋值给arrData[row][col-1]
	        for (let row = 0; row < this.p.TotalPeriodsCellValue; row++) {
	            let rowIndex = null;
	            // - `col <= maxCol`：确保列索引不超过预设的最大列数`maxCol`
	            // `col < arrFormula[1].length`：确保列索引不超过`arrFormula`数组第二行的实际长度
				for (let col = 1; col <= maxCol && col < arrFormula[1].length; col++) {
					// 修复：row从0开始，所以第0期（row=0）使用第1期公式
					if (row === 0) {
						rowIndex = 1; // 首行公式
						arrData[row][col-1] = arrFormula[rowIndex][col];//将公式赋值给数据区域数组
					} else if (row >= 1 && row < this.p.TotalPeriodsCellValue - 1) {
						rowIndex = 2; // 中间行公式
						arrData[row][col-1] = arrFormula[rowIndex][col];//将公式赋值给数据区域数组
					} else if (row === this.p.TotalPeriodsCellValue - 1) {
						rowIndex = 3; // 最后一行公式
						arrData[row][col-1] = arrFormula[rowIndex][col];//将公式赋值给数据区域数组
					}
				}
	        }
			//console.log("arrData:"+ JSON.stringify(arrData, null, 2));
			return arrData;
	        }catch (error) {
	        console.log(`[${this.MODULE_NAME}] 租金测算表数据数组转换失败：${error.message}`);
	        return null;
	    }
	}
	arrDataToDataRange(arrData){
		//将arrData数组写入表格
		//arData:表格中数据区域的二维数组
		//首先将arrData写入表格
		//然后返回数据区域Range对象
	    try {
			if (arrData === null) {
				throw new Error("数据数组为空");
			}
			//数据区域范围
			const rngData = this.p.m_worksheet.Range(
				`${this.p.m_COL_PERIOD}${this.p.RentTableStartRow}:${this.p.m_COL_PRINCIPAL_RATIO}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
			);
			
			// 修复：检查arrData的第一行第一列是否为空
			// 如果为空，说明arrData数组从索引1开始填充，需要调整
			if (arrData.length > 0 && arrData[0].length > 1 && arrData[0][0] === undefined) {
				// 创建一个新的数组，从索引1开始复制数据
				const adjustedArrData = [];
				for (let i = 0; i < arrData.length; i++) {
					adjustedArrData[i] = [];
					// 从索引1开始复制，跳过索引0的undefined
					for (let j = 1; j < arrData[i].length; j++) {
						adjustedArrData[i].push(arrData[i][j]);
					}
				}
				rngData.Value2 = adjustedArrData;
			} else {
				rngData.Value2 = arrData;
			}
			
			return rngData;
		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租金测算表数据数组写入表格失败：${error.message}`);
	        return null;
		}
	}
	arrDataRewriteCol(arrData, colIndex, newValues){
		//arrData:表格中数据区域的二维数组
		//colIndex:要修改的列索引，从0开始
		//newValues:新的列值,是个字符串
		//将arrData数组中指定列colIndex的每一行值修改为newValues、然后返回修改后的arrData数组
		try {
			if (arrData === null) {
				throw new Error("数据数组为空");
			}
			for (let row = 0; row < this.p.TotalPeriodsCellValue; row++) {
				arrData[row][colIndex] = newValues;
			}
			return arrData;
		} catch (error) {
			console.log(`[${this.MODULE_NAME}] 租金测算表数据数组指定列修改失败：${error.message}`);
	        return null;
		}
	}
	createDataRange(){
		//创建数据区域Range对象
		//返回True
	    try {
	        
	        var rng = null;
	        let arrFormula = [];
	        let arrData = [];
	        // 检查是否已初始化
	        if (this.p === null || !this.p.IsInitialized) {
	            this.Initialize();
	        }
	        
	        // 验证wsTarget是否已正确初始化
	        if (this.WsTarget === null) {
	            throw new Error("目标工作表未初始化");
	        }
	        var rng = null;
	        // 使用ParameterManager获取还款方式
	        const repaymentMethod = this.WsTarget.Range(this.p.RepaymentMethodCellA1).Value2;
	        const r1 = this.p.m_worksheet.Range(`${this.p.m_COL_PRINCIPAL_RATIO}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue}`
            	);	        
	        switch (repaymentMethod) {
	            case "等额本息（后付）":
					arrFormula= this.等额租金法arr();
					arrData= this.arrToArrData(arrFormula);
					rng= this.arrDataToDataRange(arrData);
                    this.列转化成数值以及清除(rng, [1,2]);
					this.添加框线(rng.Columns("A:J"));
	                break;
	            case "等额本金（按天计息）":
					arrFormula= this.本金比例法按天计息arr();
					arrData= this.arrToArrData(arrFormula);
					rng= this.arrDataToDataRange(arrData);
					this.列转化成数值以及清除(rng, [1,2,4], [12]);
                    this.添加框线(rng.Columns("A:J"));
	                break;
	            case "等额本金（按期计息）":
					arrFormula= this.本金比例法按期计息arr();
					arrData= this.arrToArrData(arrFormula);
					rng= this.arrDataToDataRange(arrData);
                    this.列转化成数值以及清除(rng, [1,2,4], [12]);
					this.添加框线(rng.Columns("A:J"));
	                break;
	            case "本金比例（按期计息）":
					arrFormula= this.本金比例法按期计息arr();
					arrData= this.arrToArrData(arrFormula);
					rng= this.arrDataToDataRange(arrData);
                    this.列转化成数值以及清除(rng, [1,2]);
	                this.创建租金测算表表头(12, 12)
                    this.添加框线(rng.Columns("A:J"));
                    this.添加框线(rng.Columns("L:L"));
                    this.设置背景颜色(rng.Columns("L:L"), this.p.m_COLOR_YELLOW);
	                this.租金测算表合计行(12, 12);
            		this.设置背景颜色(r1, this.p.m_COLOR_WHITE)
	                break;
	            case "本金比例（按天计息）":
					arrFormula= this.本金比例法按天计息arr();
					arrData= this.arrToArrData(arrFormula);
					rng = this.arrDataToDataRange(arrData);
                    this.列转化成数值以及清除(rng, [1,2]);
	                this.创建租金测算表表头(12, 12);
                    this.添加框线(rng.Columns("A:J"));
                    this.添加框线(rng.Columns("L:L"));
                    this.设置背景颜色(rng.Columns("L:L"), this.p.m_COLOR_YELLOW);
	                this.租金测算表合计行(12, 12)
	                this.设置背景颜色(r1, this.p.m_COLOR_WHITE)
	                break;
	            case "等额本息（先付）":
					arrFormula= this.等额本息先付arr();
					arrData= this.arrToArrData(arrFormula);
					rng= this.arrDataToDataRange(arrData);
                    this.列转化成数值以及清除(rng, [1,2]);
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
	        console.log(`[${this.MODULE_NAME}] 租金测算表生成失败：${error.message}`);
	        alert(`生成租金测算表时发生错误：${error.message}`);
	        return false;
	    }
	}
    列转化成数值以及清除(rng, convertColumns = [], clearColumns = []) {
        try {
            // 转换指定列为数值
            convertColumns.forEach(colIndex => {
                if (colIndex > 0) { // 确保列索引有效
                    const colRange = rng.Columns.Item(colIndex);
                    colRange.Value2 = colRange.Value2;
                }
            });
            
            // 清除指定列的内容
            clearColumns.forEach(colIndex => {
                if (colIndex > 0) { // 确保列索引有效
                    const colRange = rng.Columns.Item(colIndex);
                    colRange.ClearContents();
                }
            });
            
            return true;
        } catch (error) {
            console.log(`处理列数据失败：${error.message}`);
            return false;
        }
    }
	租金测算表应用格式(rng) {
	    try {
	    	
	        if (rng === null) {
	            throw new Error("目标范围为空");
	        }
	        
	        // 应用各列格式
	        let result = 应用格式(rng.Columns(1), "Integer"); // 期次列
	        if (!result) return;
	        
	        result = 应用格式(rng.Columns(2), "Date"); // 日期列
	        if (!result) return;
	        
	        // 第3-9列应用货币格式
	        for (let i = 3; i <= 9; i++) {
	            result = 应用格式(rng.Columns(i), "Standard");
	            if (!result) return;
	        }
	        
	        // 设置边框和字体
	        设置表格样式(rng);
	        
	        return true;
	    } catch (error) {
	        console.log(`租金测算表应用格式失败：${error.message}`);
	        alert(`租金测算表应用格式失败：${error.message}`);
	        return false;
	    }
	}
	清除原有表中数据() {
	    try {
	    	const WsTarget = this.p.m_worksheet;
	        if (WsTarget === null) {
	            throw new Error("目标工作表未初始化");
	        }
	        
	        // 清除参数区域
	        WsTarget.Range("B16:B17").ClearContents();
	        WsTarget.Range("D16:D18").ClearContents();
	        
	        // 清除表格数据区域
	        const clearRange = WsTarget.Range(
	            `A${this.p.RowStart}:L${this.p.CashFlowTablerowStart + this.p.TotalPeriodsCellValue + 100}`
	        );
	        clearRange.Clear();
	        
	        console.log(`[${this.MODULE_NAME}] 原有数据清除完成`);
	        return true;
	    } catch (error) {
	        console.log(`[${this.MODULE_NAME}] 数据清除失败：${error.message}`);
	        return false;
    }
	}
	添加框线(rng) {
	    try {

	        rng.Borders.LineStyle = xlContinuous;
	        rng.Borders.Color = this.p.m_COLOR_BLACK;
	        rng.Borders.Weight = xlThin;
	        rng.Borders.TintAndShade = 0;
	        return true;
	    } catch (error) {
	        console.log(`添加框线失败：${error.message}`);
	        return false;
	    }
	}
    设置背景颜色(rng, color){
	    rng.Interior.Color = color; 
    }
    自定义月间隔(targetRange, startDateCellA1){
        try{
            console.log("[" + this.MODULE_NAME + "] 开始批量生成日期公式，目标范围：" + targetRange.Address);
            console.log("[" + this.MODULE_NAME + "] 起租日单元格：" + startDateCellA1);
            let i = 0;
            let cell = null;
            let formula = "";
            let rowCount = targetRange.Rows.Count;
            // 遍历目标范围内的每个单元格
            for (let row = 1; row <= rowCount; row++) {
                cell = targetRange.Cells(row, 1);
                i = i + 1;
                if (i === 1) {
                    // 第一期：使用起租日计算，K列对应同一行的间隔值
                    formula = `=EDATE(${startDateCellA1}, K${cell.Row})`;
                    console.log("[" + this.MODULE_NAME + "] 第" + i + "期公式：" + formula);
                } else {
                    // 其他期次：使用上期支付日计算，K列对应同一行的间隔值
                    formula = "=EDATE(B" + (cell.Row - 1) + ",K" + cell.Row + ")";
                    console.log("[" + this.MODULE_NAME + "] 第" + i + "期公式：" + formula);
                }
                // 设置单元格公式
                cell.Formula = formula;
                // 设置单元格格式为日期格式
                cell.NumberFormat = "yyyy-mm-dd";
                // 添加调试信息
                console.log("[" + this.MODULE_NAME + "] 第" + i + "期单元格：" + cell.Address + "，公式设置完成:" + formula);
            }
        } catch (error) {
            console.log(`自定义月间隔失败：${error.message}`);
            return false;
        }

    }
    生成月间隔(){
        try {
            // 月间隔自定义表头生成
            this.创建租金测算表表头(11, 11);
            // 复制当前月间隔列
            const sourceRange = this.p.m_worksheet.Range(
                `${this.p.m_COL_MONTH_INTERVAL}${this.p.RentTableStartRow}:${this.p.m_COL_MONTH_INTERVAL}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
            );
            const tarRange = this.p.m_worksheet.Range(
                `${this.p.m_COL_CUSTOM_INTERVAL}${this.p.RentTableStartRow}:${this.p.m_COL_CUSTOM_INTERVAL}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
            );
            const targetRange = this.p.m_worksheet.Range(
                `${this.p.m_COL_DATE}${this.p.RentTableStartRow}:${this.p.m_COL_DATE}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
            );
            tarRange.Value2 = sourceRange.Value2;
            this.自定义月间隔(targetRange, this.p.LeaseStartDateCellA1);
            this.添加框线(tarRange);
            this.设置背景颜色(tarRange, this.p.m_COLOR_YELLOW);
            this.租金测算表合计行(11, 11);
            const r1 = this.p.m_worksheet.Range(`${this.p.m_COL_CUSTOM_INTERVAL}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue}`
            	);
            this.添加框线(r1);
            

            return true;
        } catch (error) {
            console.log(`生成月间隔失败：${error.message}`);
            return false;
        }
    }
    写入每期利率(){
    	let rowStart = this.p.RentTableStartRow;
    	let totalPeriod = this.p.TotalPeriodsCellValue;
    	let ratePerPeriod = this.p.InterestRateCellValue;
    	let arr2D = new Array(totalPeriod);
    	//生成一个二维数组，写入对应的M列
    	//第一步，生成数组
	    for(let i = 0; i < totalPeriod; i++){
	        arr2D[i] = [ratePerPeriod]; // 每行只有一个值，对应M列
	    }
	    return arr2D;
    	
    }
	调整期利率(arr2D, rate, periodn){
		//调整期利率
		//rate:新的利率
		//periodn:要调整的期次，从1开始
		try {
			// 如果arr2D未提供，先生成默认利率数组
			if (!arr2D) {
				arr2D = this.写入每期利率();
			}
			
			// 从指定的期次开始调整利率
			// 注意：periodn是从1开始的，但数组索引是从0开始，所以需要减1
			for(let i = periodn - 1; i < this.p.TotalPeriodsCellValue; i++){
				arr2D[i][0] = rate; // 每行只有一个值，对应M列
			}
			
			// 将调整后的数组写回工作表
			const rngData = this.p.m_worksheet.Range(
				`${this.p.m_COL_RatePerPeriod}${this.p.RentTableStartRow}:${this.p.m_COL_RatePerPeriod}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
			);
			rngData.Value2 = arr2D;
			
			return arr2D;
		} catch (error) {
			console.log(`调整期利率失败：${error.message}`);
	        return false;
		}
	}
	每期适用利率(){
		//生成每期适用的利率
		//数据区域范围
		const rngData = this.p.m_worksheet.Range(
			`${this.p.m_COL_RatePerPeriod}${this.p.RentTableStartRow}:${this.p.m_COL_RatePerPeriod}${this.p.RentTableStartRow + this.p.TotalPeriodsCellValue - 1}`
		);
		let arrData = this.写入每期利率();
		rngData.Value2 = arrData;
        应用格式(rngData, "Percentage");
        return true;
	}	
		
}
