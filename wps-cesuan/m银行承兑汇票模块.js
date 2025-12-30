Attribute Module_Name = "m银行承兑汇票模块"
/**
 * ============== 银行承兑汇票模块 ==============
 * 作者：徐晓冬
 * 更新日期：2025-11-21
 * 描述：本模块是银行承兑汇票生成模块，支持WPS
 * ====================================================
 */


class cls银行承兑汇票 {
    constructor() {
        // 修正：接收参数而不是直接使用全局变量p
        if (p && typeof p === 'object') {
            this.p = p;
        } else {
            // 如果没有传入参数，尝试使用全局变量p
            if (typeof p !== 'undefined') {
                this.p = p;
            } else {
                // 如果仍然没有p，设置默认值
                p = new ParameterManager();
                this.p = p;
                p.Initialize("银行承兑汇票")
            }
        }
         // 设置工作表
        this.ws = this.p.m_billSheet;
        this.wsourse = this.p.m_sourceSheet;

        this.MODULE_NAME = "银行承兑汇票模块";
        // 安全地获取属性值
        this.RowStart = this.p.RowStart;
        this.RentTableStartRow = this.p.RentTableStartRow;
        this.SheetNamerow = this.p.SheetNamerow;
        this.arrHeaders = ["期次", "日期", "净现金流1（1-9）", "净现金流1-备注",
                        "净现金流2（1-8）", "净现金流2-备注", "（1）电汇放款",
                        "（2)银行承兑汇票放款-保证金", "（3）银行承兑汇票放款-尾款",
                        "（4）银行承兑汇票-手续费", "（5）银行承兑汇票-利息收入",
                        "（6）我司收取客户保证金/退回保证金", "（7）租金",
                        "（8）名义货价", "（9）经纪人费用支付"];// 银行承兑汇票现金流量表的一维表头数组
        this.lenghHeader = this.arrHeaders.length;// 表头长度


        //this.ws = this.p.m_worksheet;
            
        // 租金测算表数据区第1行Row值
        this.RentTableStartRow = this.p.RentTableStartRow;
        this.totalPeriods = this.p.TotalPeriodsCellValue;
        this.paymentInterval = this.p.PaymentIntervalCellValue;
        this.Principal = this.p.PrincipalCellR1C1;
        this.CashFlowTablerowStart = this.p.CashFlowTablerowStart; // 银行承兑汇票Sheet的现金流量表数据区第1行Row值
        console.log(`[${this.MODULE_NAME}] 类实例创建`);
    }
    
    /**
     * 初始化模块参数
     */
    Initialize() {
        try {
            console.log(`[${this.MODULE_NAME}] 开始初始化...`);
            
            console.log(`[${this.MODULE_NAME}] 初始化完成 - 总期数:${this.totalPeriods}`);
            return true;
            
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 初始化失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 生成银行承兑汇票现金流量表
     */
    生成银承现金流量表() {
        try {
            this.创建银承现金流量表表头();
            this.银承放款现金流();
            this.银承现金流1备注update();
            this.银承现金流2备注update();
            this.银承综合利率一览();
            
            console.log(`[${this.MODULE_NAME}] 银行承兑汇票现金流量表生成完成`);
            return true;
            
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 生成银行承兑汇票现金流量表失败：${error.message}`);
            return false;
        }
    } 
    /**
     * 写入表头
     */
    WriteHeaders2(sheetName) {
        try {            
            console.log(`[${this.MODULE_NAME}] 表头长度: ${this.lenghHeader}`);
                      
            // 总标题行
            this.p.m_billSheet.Range("A" + this.RowStart).Value2 = "现金流及综合利率测算";
            const titleRange =  this.p.m_billSheet.Range("A" + (this.RowStart));
            titleRange.Interior.Color = RGB(255, 255, 255); // 设置单元格背景为蓝色
            titleRange.Font.Name = FONT_DEFAULT;            // 设置字体为黑体
            titleRange.Font.Size = FONT_SIZE_TITLE;                // 设置字号为14号
            titleRange.Font.Color = RGB(0, 0, 0);     // 设置字体颜色为黑色
            
            // 将数组写入行
            const headerRange =  this.p.m_billSheet.Range("A" + (this.RowStart + 1) + ":O" + (this.RowStart + 1));
            headerRange.Value2 = this.arrHeaders;
            
            const formatRange =  this.p.m_billSheet.Range("A" + (this.RowStart + 1) + ":O" + (this.RowStart + 1));
            formatRange.Interior.Color = RGB(0, 174, 240); // 设置单元格背景为蓝色
            formatRange.Font.Name = FONT_DEFAULT;            // 设置字体为黑体
            formatRange.Font.Size = FONT_SIZE_HEADER;                // 设置字号为12号
            formatRange.Font.Color = RGB(0, 0, 0);     // 设置字体颜色为黑色
            formatRange.HorizontalAlignment = XL.HCenter; // 设置水平居中对齐
            formatRange.VerticalAlignment = XL.VCenter;  // 设置垂直居中对齐
            formatRange.WrapText = true;               // 设置自动换行
            
            console.log(`[${this.MODULE_NAME}] 表头写入成功`);
            return true;
            
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] WriteHeaders2函数执行失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 创建银行承兑汇票现金流量表表头
     */
    创建银承现金流量表表头() {
        try {
            let result = false;
            // 调用函数
            result = this.WriteHeaders2(this.p.m_billSheetName);
            return result;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 创建银承现金流量表表头失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 计算放款次数
     */
    nFangKuan() {
        try {
            let firstNonEmptyRow = 0;
            
            // 从A11开始向下查找第一个非空单元格
            const startRange = this.p.m_billSheet.Range("A11");
            const endRange = startRange.End(xlDown);
            
            // 检查是否到达了工作表末尾（可能是空的工作表）
            firstNonEmptyRow = endRange.Row;
            
            // 如果firstNonEmptyRow小于等于11，说明可能是空的或者是第一行就有数据
            if (firstNonEmptyRow <= 11) {
                // 检查A11是否真的有数据
                if (startRange.Value2 === null || startRange.Value2 === "" || startRange.Value2 === undefined) {
                    console.log(`[${this.MODULE_NAME}] Sheet10中没有找到放款数据`);
                    return 0;
                }
            }
            
            const fangKuanNum = firstNonEmptyRow - this.p.m_billSheet.Range("A10").Row; // 计算得出放款次数
            console.log("放款次数: " + fangKuanNum);
            return fangKuanNum;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] nFangKuan函数执行失败：${error.message}`);
            return 0;
        }
    }
    
    /**
     * 保存放款信息到数组
     */
    SaveFangKuanInfoToArr() {
        try {
            //let arr = [];
            let nFangKuan = 0; // 放款次数
            let firstNonEmptyRow = 0;
            
            
            // 从A11开始向下查找第一个非空单元格
            const startRange = this.p.m_billSheet.Range("A11")
            let endRange = startRange;
            if (startRange.Value2 != null && startRange.Value2 !== ""){
            	endRange = startRange.End(xlDown);
            }
            
            
            firstNonEmptyRow = endRange.Row;
            console.log("firstNonEmptyRow: " + firstNonEmptyRow);
            
            // 检查是否有数据
            if (firstNonEmptyRow <= 11) {
                // 检查A11是否真的有数据
                if (this.ws.Range("A11").Value2 === null || 
                    this.ws.Range("A11").Value2 === "" || 
                    this.ws.Range("A11").Value2 === undefined) {
                    console.log(`[${this.MODULE_NAME}] Sheet10中没有找到放款数据`);
                    return null;
                }
                nFangKuan = 1; // 至少有一行数据
            } else {
                nFangKuan = firstNonEmptyRow - 10; // A10是起始行，所以减去10
            }
            
            console.log("nFangKuan: " + nFangKuan);
            
            // 安全地计算范围
            if (nFangKuan <= 0) {
                console.log(`[${this.MODULE_NAME}] 没有有效的放款数据`);
                return null;
            }
            
            const endRow = 10 + nFangKuan;
            console.log("读取范围: A11:J" + endRow);
            let rng = this.p.m_billSheet.Range("A11:J" + endRow)
            // 存储放款信息到数组arr
            let arr = rng.Value2;
            // 2）格式化成字符串更好看
			console.log(JSON.stringify(arr, null, 2));
			arr.forEach((row, i) => {
			  row.forEach((cell, j) => {
			    console.log(`row ${i+1}, col ${j+1}:`, cell);
			  });
			});
			
            // 将放款期次信息填写到现金流量表中
            return arr;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] SaveFangKuanInfoToArr函数执行失败：${error.message}`);
            console.log(`错误堆栈: ${error.stack}`);
            return null;
        }
    }
    //流量表数据生成-新
    //
    //
    流量表数据生成(){
    	null;
    }
    
    /**
     * 银承放款现金流
     */
    银承放款现金流() {
        try {
            let arr = null;
            let n = 0; // 放款次数
            let k2 = 0; // 计数用，用于统计放款占用行数
            let sstr = "0-"; // 字符串
            
            let k1 = 0;
            arr = this.SaveFangKuanInfoToArr();
            if (!arr) {
                console.log(`[${this.MODULE_NAME}] 放款信息数组为空，跳过放款现金流生成`);
                return true;
            }
            n = this.nFangKuan();
            
            // 构建放款的现金流量表
            for (k1 = 0; k1 < n; k1++) {
                if (this.ws.Range("D" + (11 + k1)).Value2 === "银行承兑汇票") {
                    // 银行承兑汇票放款
                    this.ws.Range("A" + (this.RentTableStartRow + k2)).Formula = "=concat(\"" + sstr + "\",\"" + k1 + "\")"; // 生成期次
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).Value2 = arr[k1][0]; // 放款日期
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).NumberFormat = "yyyy-mm-dd ;@";
                    this.ws.Range("C" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[4]:RC[12])"; // 合计
                    this.ws.Range("D" + (this.RentTableStartRow + k2)).Value2 = ""; // 空着
                    this.ws.Range("E" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[2]:RC[9])"; // 合计
                    this.ws.Range("G" + (this.RentTableStartRow + k2)).Value2 = ""; // 电汇放款，此处空着
                    this.ws.Range("H" + (this.RentTableStartRow + k2)).Value2 = -1 * (arr[k1][5] || 0); // （2)银行承兑汇票放款-保证金
                    this.ws.Range("I" + (this.RentTableStartRow + k2)).Value2 = ""; // （3）银行承兑汇票放款-尾款
                    this.ws.Range("J" + (this.RentTableStartRow + k2)).Value2 = -1 * ((arr[k1][7] || 0) * (arr[k1][1] || 0)); // （4）银行承兑汇票-手续费
                    this.ws.Range("K" + (this.RentTableStartRow + k2)).Value2 = ""; // （5）银行承兑汇票-利息收入
                    this.ws.Range("L" + (this.RentTableStartRow + k2)).Value2 = arr[k1][8] || 0; // （6）期初我司收取客户保证金/退回保证金
                    this.ws.Range("M" + (this.RentTableStartRow + k2)).Value2 = ""; // （7）租金
                    this.ws.Range("O" + (this.RentTableStartRow + k2)).Value2 = ""; // （8）经纪人费用支付
                    k2 = k2 + 1;
                    
                    // 银行承兑汇票尾款
                    this.ws.Range("A" + (this.RentTableStartRow + k2)).Formula = "=concat(\"" + sstr + "\",\"" + k1 + "\")"; // 生成期次
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).Formula = `=EDATE(${arr[k1][0]}, ${arr[k1][4] || 0})`; // 放款日期+银承对付日期
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).NumberFormat = "yyyy-mm-dd ;@";
                    this.ws.Range("C" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[4]:RC[12])"; // 合计
                    this.ws.Range("D" + (this.RentTableStartRow + k2)).Value2 = ""; // 合计备注空着
                    this.ws.Range("E" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[2]:RC[9])"; // 合计
                    this.ws.Range("G" + (this.RentTableStartRow + k2)).Value2 = ""; // 电汇放款，此处空着
                    this.ws.Range("H" + (this.RentTableStartRow + k2)).Value2 = ""; // （2)银行承兑汇票放款-保证金
                    this.ws.Range("I" + (this.RentTableStartRow + k2)).Value2 = -1 * ((arr[k1][1] || 0) - (arr[k1][5] || 0)); // （3）银行承兑汇票放款-尾款
                    this.ws.Range("J" + (this.RentTableStartRow + k2)).Value2 = ""; // （4）银行承兑汇票-手续费,放款时收取，此处为空
                    this.ws.Range("K" + (this.RentTableStartRow + k2)).Value2 = (arr[k1][6] || 0) * (arr[k1][1] || 0); // （5）银行承兑汇票-利息收入
                    this.ws.Range("L" + (this.RentTableStartRow + k2)).Value2 = ""; // （6）期初我司收取客户保证金/退回保证金,此处为空
                    this.ws.Range("M" + (this.RentTableStartRow + k2)).Value2 = ""; // （7）租金，此处为空
                    this.ws.Range("O" + (this.RentTableStartRow + k2)).Value2 = ""; // （8）经纪人费用支付,此处为空
                    k2 = k2 + 1; // 银行承兑汇票占用2行
                } else {
                    this.ws.Range("D" + (11 + k1)).Value2 = "电汇";
                    // 电汇放款
                    this.ws.Range("A" + (this.RentTableStartRow + k2)).Formula = "=concat(\"" + sstr + "\",\"" + k1 + "\")"; // 生成期次
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).Value2 = arr[k1][0]; // 放款日期
                    this.ws.Range("B" + (this.RentTableStartRow + k2)).NumberFormat = "yyyy-mm-dd ;@";
                    this.ws.Range("C" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[4]:RC[12])"; // 合计
                    this.ws.Range("D" + (this.RentTableStartRow + k2)).Value2 = ""; // 空着
                    this.ws.Range("E" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[2]:RC[9])"; // 合计
                    this.ws.Range("G" + (this.RentTableStartRow + k2)).Formula = "= -1 * " + (arr[k1][1] || 0);  // 电汇放款
                    this.ws.Range("H" + (this.RentTableStartRow + k2)).Value2 = ""; // （2)银行承兑汇票放款-保证金,此处不涉及
                    this.ws.Range("I" + (this.RentTableStartRow + k2)).Value2 = ""; // （3）银行承兑汇票放款-尾款，此处不涉及
                    this.ws.Range("J" + (this.RentTableStartRow + k2)).Value2 = ""; // （4）银行承兑汇票-手续费,此处不涉及
                    this.ws.Range("K" + (this.RentTableStartRow + k2)).Value2 = ""; // （5）银行承兑汇票-利息收入，此处不涉及
                    this.ws.Range("L" + (this.RentTableStartRow + k2)).Value2 = arr[k1][8] || 0; // （6）期初我司收取客户保证金/退回保证金
                    this.ws.Range("M" + (this.RentTableStartRow + k2)).Value2 = ""; // （7）租金，此处不涉及
                    this.ws.Range("O" + (this.RentTableStartRow + k2)).Value2 = arr[k1][9] || 0; // （8）经纪人费用支付
                    k2 = k2 + 1; // 电汇占用1行
                }
            }
            

            console.log("已占用行数K2:" + k2);
            k1 = 0; // k1计数器归零
            console.log("TotalPeriods本次总期数为：" + this.totalPeriods);
            
            for (k1 = 1; k1 <= this.totalPeriods; k1++) {
                // 还款阶段
                this.ws.Range("A" + (this.RentTableStartRow + k2)).Value2 = k1; // 生成期次
                this.ws.Range("B" + (this.RentTableStartRow + k2)).Value2 = this.p.m_sourceSheet.Range("B" + (this.RentTableStartRow - 1 + k1)).Value2; // 日期
                this.ws.Range("B" + (this.RentTableStartRow + k2)).NumberFormat = "yyyy-mm-dd ;@";
                this.ws.Range("C" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[4]:RC[12])"; // 合计
                this.ws.Range("D" + (this.RentTableStartRow + k2)).Value2 = ""; // 空着
                this.ws.Range("E" + (this.RentTableStartRow + k2)).FormulaR1C1 = "=sum(RC[2]:RC[9])"; // 合计
                this.ws.Range("G" + (this.RentTableStartRow + k2)).Value2 = ""; // 电汇放款，此处空着
                this.ws.Range("H" + (this.RentTableStartRow + k2)).Value2 = ""; // （2)银行承兑汇票放款-保证金,此处不涉及
                this.ws.Range("I" + (this.RentTableStartRow + k2)).Value2 = ""; // （3）银行承兑汇票放款-尾款，此处不涉及
                this.ws.Range("J" + (this.RentTableStartRow + k2)).Value2 = ""; // （4）银行承兑汇票-手续费,此处不涉及
                this.ws.Range("K" + (this.RentTableStartRow + k2)).Value2 = ""; // （5）银行承兑汇票-利息收入，此处不涉及
                this.ws.Range("L" + (this.RentTableStartRow + k2)).Value2 = this.p.m_sourceSheet.Range("I" + (this.CashFlowTablerowStart + k1)).Value2 || 0; // （6）期初我司收取客户保证金/退回保证金
                this.ws.Range("M" + (this.RentTableStartRow + k2)).Value2 =  this.p.m_sourceSheet.Range("C" + (this.CashFlowTablerowStart + k1)).Value2 || 0;  // （7）租金
                this.ws.Range("N" + (this.RentTableStartRow + k2)).Value2 =  this.p.m_sourceSheet.Range("J" + (this.CashFlowTablerowStart + k1)).Value2 || 0; // 名义货价
                this.ws.Range("O" + (this.RentTableStartRow + k2)).Value2 =  this.p.m_sourceSheet.Range("K" + (this.CashFlowTablerowStart + k1)).Value2 || 0; // （8）经纪人费用支付
                k2 = k2 + 1;
                console.log(k1);
            }
            
            const lastRow = this.totalPeriods + this.RentTableStartRow - 1 + k2;
            this.ws.Range("G" + this.RentTableStartRow + ":O" + lastRow).NumberFormat = "#,##0.00"; // 设置格式
            this.ws.Range("C" + this.RentTableStartRow + ":C" + lastRow).NumberFormat = "#,##0.00"; // 设置格式
            this.ws.Range("E" + this.RentTableStartRow + ":E" + lastRow).NumberFormat = "#,##0.00"; // 设置格式
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 银承放款现金流函数执行失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 银承现金流1备注
     */
    arrRngCashFlow(){
    	//采用数组编写,生成银承现金流量表数据
        let arr = [];//存放银承现金流量表数据，设计成二维数组
        //计算生成arr的行数（row值），等于 放款形成的row + 还款形成的row，
        // 计算从银承现金流表数据第1行开始，到最后一行的row值。
        //列值（col值）为headers的长度，即15列
        let RowStart= this.RentTableStartRow;
        let rngStart= this.p.m_billSheet.Range("A"+String(RowStart));
        let rngEnd= rngStart.End(xlDown);
        let rows= rngEnd.Row - RowStart +1;//总行数,作为arr的row值，row值就是数组的第1维脚标数字
        console.log("总行数,作为arr的row值，row值就是数组的第1维脚标数字:"+ rows);
        let cols = null;
        if (this.lenghHeader == null){
        	cols = this.lenghHeader;//列值（col值）为headers的长度，即15列，作为arr的第2维脚标数字
        }else{
        	cols = this.lenghHeader;//列值（col值）为headers的长度，即15列，作为arr的第2维脚标数字
        }
        
        console.log("列值（col值）为headers的长度，即15列，作为arr的第2维脚标数字:"+ cols);
        
        let sheet = this.p.m_billSheet;
        const nameOfsheet = sheet.Name;
        //let range = sheet.getRange(RowStart, 1, rows, cols); 
        let range = sheet.Range(sheet.Cells(RowStart, 1), sheet.Cells(RowStart+rows, cols)) 
        let sheetData = range.Value2; // 获取范围内的所有值
        for (let i = 0; i < rows; i++){
            let row = [];
            for (let j = 0; j < cols;j++){
                //初始化二维数组arr[i][j]的值
                //arr[i][j] = "";//单元格值
                row.push({
                    value: sheetData[i][j], //单元格值
                    row: i, //行索引
                    col: j, //列索引
                    address: `R${i+1}C${j+1}`, //单元格地址
                    sheetAddress: `${nameOfsheet}!R${i+1}C${j+1}`, //在Sheet中的实际地址,工作表+单元格地址
                    header: this.arrHeaders[j] //对应的表头名称
                });
            }
            arr.push(row);
        }
       // console.log("arr:"+ JSON.stringify(arr, null, 2));
        saveWithFileSystem("arr银承现金流量表数据.json", JSON.stringify(arr, null, 2));
        return arr;
    }
    银承现金流1备注update(){
        let arr = this.arrRngCashFlow();
        if (!arr) {
            console.log(`[${this.MODULE_NAME}] 银承现金流量表数据为空，跳过备注更新`);
            return true;
        }
        //将arr数组中 净现金流1-备注 的数据，更新到银承现金流量表中
        for (let i = 0; i < arr.length; i++){
            let cell = arr[i][3]; //净现金流1-备注 列
            let brr = []; //存放备注信息的数组
            if (arr[i][6].value !== undefined && arr[i][6].value !== null) {
                brr.push("电汇放款:" + arr[i][6].value);
            }

            if (arr[i][7].value !== undefined && arr[i][7].value !== null) {
                brr.push("银行承兑汇票放款-保证金:" + arr[i][7].value);
            }
            if (arr[i][8].value !== undefined && arr[i][8].value !== null) {
                brr.push("银行承兑汇票放款-尾款:" + arr[i][8].value);
            }
            if (arr[i][9].value !== undefined && arr[i][9].value !== null) {
                brr.push("银行承兑汇票-手续费:" + arr[i][9].value);
            }
            if (arr[i][10].value !== undefined && arr[i][10].value !== null) {
                brr.push("银行承兑汇票-利息收入:" + arr[i][10].value);
            }
            if (arr[i][11].value !== undefined && arr[i][11].value !== null) {
                if (arr[i][11].value > 0) {
                    brr.push("我司收取客户保证金:" + arr[i][11].value);
                }
                else if (arr[i][11].value < 0) {
                    brr.push("我司退回保证金:" + arr[i][11].value);
                }
            }
            if (arr[i][12].value !== undefined && arr[i][12].value !== null) {
                brr.push("第" + arr[i][0].value + "期租金:" + arr[i][12].value);
            }
            if (arr[i][13].value !== undefined && arr[i][13].value !== null) {
                brr.push("名义货价:" + arr[i][13].value);
            }
            if (arr[i][14].value !== undefined && arr[i][14].value !== null && arr[i][14].value !== 0) {
                brr.push("经纪人费用:" + arr[i][14].value);
            }
            cell.value = brr.join("/"); //将备注信息数组连接成字符串，赋值给cell.value
        }
        //将更新后的备注信息写入到银承现金流量表中
        for (let i = 0; i < arr.length; i++){
            let cell = arr[i][3]; //净现金流1-备注 列
            let excelCell = this.ws.Range("D" + (this.RentTableStartRow + i));
            excelCell.Value2 = cell.value;
            excelCell.HorizontalAlignment = -4108; //设置水平居中对齐
        }
        return true;
    }
    银承现金流2备注update(){
        let arr = arrDataFromRng(this.p.m_billSheet, this.p.RentTableStartRow, this.arrHeaders);
        if (!arr) {
            console.log(`[${this.MODULE_NAME}] 银承现金流量表数据为空，跳过备注更新`);
            return true;
        }
        //将arr数组中 净现金流2-备注 的数据，更新到银承现金流量表中
        for (let i = 0; i < arr.length; i++){
            let cell = arr[i][5]; //净现金流2-备注 列
            let brr = []; //存放备注信息的数组
            if (arr[i][6].value !== undefined && arr[i][6].value !== null) {
                brr.push("电汇放款:" + arr[i][6].value);
            }

            if (arr[i][7].value !== undefined && arr[i][7].value !== null) {
                brr.push("银行承兑汇票放款-保证金:" + arr[i][7].value);
            }
            if (arr[i][8].value !== undefined && arr[i][8].value !== null) {
                brr.push("银行承兑汇票放款-尾款:" + arr[i][8].value);
            }
            if (arr[i][9].value !== undefined && arr[i][9].value !== null) {
                brr.push("银行承兑汇票-手续费:" + arr[i][9].value);
            }
            if (arr[i][10].value !== undefined && arr[i][10].value !== null) {
                brr.push("银行承兑汇票-利息收入:" + arr[i][10].value);
            }
            if (arr[i][11].value !== undefined && arr[i][11].value !== null) {
                if (arr[i][11].value > 0) {
                    brr.push("我司收取客户保证金:" + arr[i][11].value);
                }
                else if (arr[i][11].value < 0) {
                    brr.push("我司退回保证金:" + arr[i][11].value);
                }
            }
            if (arr[i][12].value !== undefined && arr[i][12].value !== null) {
                brr.push("第" + arr[i][0].value + "期租金:" + arr[i][12].value);
            }
            if (arr[i][13].value !== undefined && arr[i][13].value !== null) {
                brr.push("名义货价:" + arr[i][13].value);
            }
            cell.value = brr.join("/"); //将备注信息数组连接成字符串，赋值给cell.value
        }
        //将更新后的备注信息写入到银承现金流量表中
        for (let i = 0; i < arr.length; i++){
            let cell = arr[i][5]; //净现金流2-备注 列
            let excelCell = this.ws.Range("F" + (this.RentTableStartRow + i));
            excelCell.Value2 = cell.value;
            excelCell.HorizontalAlignment = -4108; //设置水平居中对齐
        }
        return true;
    }

    
    /**
     * 银承分类汇总
     */
    银承分类汇总() {
        try {
            let dd = {};
            let arr = null;
            let hrr = new Array(100);
            for (let i = 0; i < 100; i++) {
                hrr[i] = new Array(15);
            }
            let k = 0;
            let i = 0;
            let x = 0;
            let row = 0;
            let maxRow = 0;
            
            maxRow = this.ws.Range("A21").CurrentRegion.Rows.Count;
            
            arr = this.ws.Range("A22:L41").Value2;
            k = 0;
            x = 1;
            i = 1;
            
            for (x = 1; x <= arr.length; x++) {
                if (dd[arr[x-1][1]]) {
                    row = dd[arr[x-1][1]];      // 如果存在。字典中的顺序序号赋值给row
                    hrr[row-1][0] = (hrr[row-1][0] || "") + "\\" + arr[x-1][0];       // 文本格式，对应累加
                    hrr[row-1][3] = (hrr[row-1][3] || "") + "\\" + arr[x-1][3];       // 文本格式，对应累加
                    hrr[row-1][2] = (hrr[row-1][2] || 0) + arr[x-1][2];       // 对应累加
                    hrr[row-1][4] = (hrr[row-1][4] || 0) + arr[x-1][4];       // 对应累加
                    hrr[row-1][5] = (hrr[row-1][5] || 0) + arr[x-1][5];       // 对应累加
                    hrr[row-1][6] = (hrr[row-1][6] || 0) + arr[x-1][6];       // 对应累加
                    hrr[row-1][7] = (hrr[row-1][7] || 0) + arr[x-1][7];       // 对应累加
                    hrr[row-1][8] = (hrr[row-1][8] || 0) + arr[x-1][8];       // 对应累加
                    hrr[row-1][9] = (hrr[row-1][9] || 0) + arr[x-1][9];       // 对应累加
                    hrr[row-1][10] = (hrr[row-1][10] || 0) + arr[x-1][10];       // 对应累加
                    hrr[row-1][11] = (hrr[row-1][11] || 0) + arr[x-1][11];       // 对应累加
                } else {
                    // 如果日期不存在 ，k是序号，dd的键是日期，value是新序号
                    k = k + 1;
                    dd[arr[x-1][1]] = k;
                    
                    hrr[k-1][0] = arr[x-1][0];
                    hrr[k-1][1] = arr[x-1][1];
                    hrr[k-1][2] = arr[x-1][2];
                    hrr[k-1][3] = arr[x-1][3];
                    hrr[k-1][4] = arr[x-1][4];
                    hrr[k-1][5] = arr[x-1][5];
                    hrr[k-1][6] = arr[x-1][6];
                    hrr[k-1][7] = arr[x-1][7];
                    hrr[k-1][8] = arr[x-1][8];
                    hrr[k-1][9] = arr[x-1][9];
                    hrr[k-1][10] = arr[x-1][10];
                    hrr[k-1][11] = arr[x-1][11];
                }
            }
            
            let lrr = null;
            // lrr为标题表头
            lrr = this.ws.Range("A21:L21").Value2;
            Application.Sheets("Sheet14").Range("A1").Resize(1, 12).Value2 = lrr;
            
            Application.Sheets("Sheet14").Range("A2").Resize(k, 12).Value2 = hrr;      // 汇总后结果输出
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 银承分类汇总函数执行失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 清除银行承兑汇票表中数据
     */
    清除银行承兑汇票表中数据() {
        try {
            this.Initialize();
            
            // 清除原有表中数据
            this.ws.Range("A" + this.RentTableStartRow + ":O" + (this.CashFlowTablerowStart + this.totalPeriods + 1 + 100)).Clear();
            this.ws.Range("B21:B22").ClearContents();
            this.ws.Range("D21:D23").ClearContents();
            this.ws.Range("F21:F23").ClearContents();
            this.ws.Range("H21:H22").ClearContents();
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 清除银行承兑汇票表中数据函数执行失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 银承综合利率一览
     */
    银承综合利率一览() {
        try {
            let n = 0;
            let k1 = 0; // 临时变量，用于计数
            let k2 = 0; // 计数用，用于统计放款占用行数
            let arr = null;
            let brr = new Array(100);
            for (let i = 0; i < 100; i++) {
                brr[i] = new Array(1);
            }
            let row = 0;
            let formula = "";

            n = this.nFangKuan();
            k2 = 0;
            
            for (k1 = 1; k1 <= n; k1++) {
                if (this.ws.Range("D" + (10 + k1)).Value2 === "银行承兑汇票") {
                    // 银行承兑汇票放款
                    k2 = k2 + 2;
                } else {
                    this.ws.Range("D" + (10 + k1)).Value2 = "电汇";
                    k2 = k2 + 1;
                }
            }
            
            //console.log("放款期间占用行数：【" + k2 + "】");
            formula = "=XIRR(C" + this.RentTableStartRow + ":C" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + "," + 
                      "B" + this.RentTableStartRow + ":B" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + ")";
            this.ws.Range("D21").Formula = formula;
            formula = "=XIRR(E" + this.RentTableStartRow + ":E" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + "," + 
                      "B" + this.RentTableStartRow + ":B" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + ")";
            this.ws.Range("D22").Formula = formula;
            formula = "=R[-2]C-R[-1]C";
            this.ws.Range("D23").FormulaR1C1 = formula;
            formula = "=IRR(C" + this.RentTableStartRow + ":C" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + ")*" + this.p.PaymentsPerYearValue;
            this.ws.Range("B21").Formula = formula;
            formula = "=IRR(E" + this.RentTableStartRow + ":E" + (this.RentTableStartRow + this.totalPeriods + k2 - 1) + ")*" + this.p.PaymentsPerYearValue;
            this.ws.Range("B22").Formula = formula;
            
            formula = "=1租金测算表V1!D16";
            this.ws.Range("F21").Formula = formula; // Sheet1.Range("D16")
            formula = "=1租金测算表V1!D17";
            this.ws.Range("F22").Formula = formula;
            formula = "=F21- F22";
            this.ws.Range("F23").Formula = formula; // Sheet1.Range("D18")
            formula = "=D21 - F21";
            this.ws.Range("H21").Formula = formula;
            formula = "=D22 - F22";
            this.ws.Range("H22").Formula = formula;
            
            this.ws.Range("B21:B22,D21:D23").NumberFormatLocal = "0.00%";
            this.ws.Range("E21:E23").NumberFormatLocal = "0.00%";
            
            return true;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] 银承综合利率一览函数执行失败：${error.message}`);
            return false;
        }
    }
    
    /**
     * 日期加法函数
     */
    DateAdd(interval, number, date) {
        try {
            const d = new Date(date);
            switch (interval.toLowerCase()) {
                case "m":
                    d.setMonth(d.getMonth() + number);
                    break;
                case "d":
                    d.setDate(d.getDate() + number);
                    break;
                case "y":
                    d.setFullYear(d.getFullYear() + number);
                    break;
                case "h":
                    d.setHours(d.getHours() + number);
                    break;
                case "n":
                    d.setMinutes(d.getMinutes() + number);
                    break;
                case "s":
                    d.setSeconds(d.getSeconds() + number);
                    break;
            }
            return d;
        } catch (error) {
            console.log(`[${this.MODULE_NAME}] DateAdd函数执行失败：${error.message}`);
            return new Date();
        }
    }
}
