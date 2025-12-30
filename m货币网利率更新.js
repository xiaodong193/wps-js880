Attribute Module_Name = "m货币网利率更新"
/**
 * ============== 模块1：通过中国货币网更新LPR利率 ==============
 * 作者：徐晓冬
 * 版本：1.00
 * 更新日期：
 * 描述：通过中国货币网更新LPR利率，支持WPS
 * ====================================================
 */

class LPRDownloader {
    constructor() {
        this.MODULE_NAME = "LPRDownloader";
        console.log("[" + this.MODULE_NAME + "] 初始化完成");
    }

    /**
     * 下载并复制LPR数据
     * @returns {boolean} 操作是否成功
     */
    DownloadAndCopyLPRData() {
        try {
            let wb = null;
            let ws = null;
            let targetWs = null;
            let filePath = "";
            let url = "";
            let currentFolder = "";
            let startDate = "";
            let endDate = "";
            let cell = null;

            // 格式化日期
            const today = new Date();
            const oneYearAgo = new Date(today);
            oneYearAgo.setFullYear(today.getFullYear() - 1);
            
            startDate = this.FormatDate(oneYearAgo);  // 一年前的日期
            endDate = this.FormatDate(today);  // 今天的日期

            console.log("[" + this.MODULE_NAME + "] 开始下载LPR数据，日期范围: " + startDate + " 到 " + endDate);

            // 设置URL和下载路径
            url = "https://www.chinamoney.com.cn/dqs/rest/cm-u-bk-currency/LprHisExcel?lang=CN&strStartDate=" + startDate + "&strEndDate=" + endDate;
            currentFolder = ThisWorkbook.Path;
            
            if (currentFolder === "") {
                console.log("请先保存当前工作簿。");
                return false;
            }
            
            filePath = currentFolder + "\\LPRData.xlsx";

            // 方法1: 直接使用WPS的Workbooks.Open方法打开URL
            if (this.DownloadWithWPSOpen(url, filePath)) {
                // 打开下载的工作簿
                wb = Workbooks.Open(filePath);
                ws = wb.Sheets("Sheet0");
                
                // 确保目标工作表存在
                try {
                    targetWs = ThisWorkbook.Sheets("贷款基础利率");
                } catch (error) {
                    targetWs = ThisWorkbook.Sheets.Add();
                    targetWs.Name = "贷款基础利率";
                }

                // 复制数据
                ws.Range("A2:C13").Copy();
                targetWs.Range("B3").PasteSpecial(-4163); // xlPasteValues = -4163
                targetWs.Range("B3:D3").NumberFormat = "0.00";
				// 清除剪贴板
				Application.CutCopyMode = false;
                // 关闭下载的工作簿
                wb.Close(false); // SaveChanges = false

                // 将C3:D14转换为数值格式，保留两位小数
                const targetRange = targetWs.Range("C3:D14");
                for (let i = 1; i <= targetRange.Rows.Count; i++) {
                    for (let j = 1; j <= targetRange.Columns.Count; j++) {
                        cell = targetRange.Cells(i, j);
                        if (this.IsNumeric(cell.Value)) {
                            cell.NumberFormat = "0.00";
                            cell.Value = parseFloat(cell.Value);
                        }
                    }
                }

                console.log(
                    "中国货币网的LPR数据已成功下载到当前工作簿文件夹并复制到'贷款基础利率'工作表。\n" +
                    "日期范围：" + startDate + " 到 " + endDate + "\n。"
                );
                
                console.log("[" + this.MODULE_NAME + "] LPR数据下载成功");
                return true;
            } else {
                console.log("下载LPR数据失败。请检查网络连接或URL是否正确。");
                return false;
            }

        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] 发生错误: " + error.message);
            console.log("[" + this.MODULE_NAME + "] 错误详情: " + error.stack);
            console.log("发生错误: " + error.message);
            return false;
        }
    }

    /**
     * 方法1: 使用WPS的Workbooks.Open直接打开URL
     */
    DownloadWithWPSOpen(url, filePath) {
        try {
            console.log("[" + this.MODULE_NAME + "] 尝试使用WPS直接打开URL");
            
            // 直接使用WPS的Workbooks.Open方法打开URL
            const tempWorkbook = Workbooks.Open(url);
            
            if (tempWorkbook) {
                // 保存到本地文件
                tempWorkbook.SaveAs(filePath);
                tempWorkbook.Close(false); // 不保存更改
                
                console.log("[" + this.MODULE_NAME + "] WPS直接打开URL成功");
                return true;
            }
            return false;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] WPS直接打开URL失败: " + error.message);
            return false;
        }
    }

    /**
     * 方法2: 使用WPS JSA的CreateObject方法
     */
    DownloadWithCreateObject(url, filePath) {
        try {
            console.log("[" + this.MODULE_NAME + "] 尝试使用CreateObject方法");
            
            // 使用WPS JSA的CreateObject创建对象
            const httpRequest = CreateObject("MSXML2.XMLHTTP");
            if (!httpRequest) {
                console.log("[" + this.MODULE_NAME + "] CreateObject创建失败");
                return false;
            }
            
            httpRequest.open("GET", url, false);
            httpRequest.send();
            
            if (httpRequest.status === 200) {
                // 使用WPS的文件操作保存数据
                const fs = CreateObject("Scripting.FileSystemObject");
                if (fs) {
                    const stream = fs.CreateTextFile(filePath, true);
                    stream.Write(httpRequest.responseText);
                    stream.Close();
                    console.log("[" + this.MODULE_NAME + "] CreateObject下载成功");
                    return true;
                }
            }
            return false;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] CreateObject方法失败: " + error.message);
            return false;
        }
    }

    /**
     * 方法3: 使用WPS JSA的FileSystemObject
     */
    DownloadWithFileSystem(url, filePath) {
        try {
            console.log("[" + this.MODULE_NAME + "] 尝试使用FileSystemObject方法");
            
            // 使用WPS JSA的文件系统对象
            const fso = Application.FileSystemObject;
            if (fso) {
                // 这里需要WPS JSA支持的文件下载方法
                // 由于WPS JSA限制，可能需要其他方法
                console.log("[" + this.MODULE_NAME + "] FileSystemObject可用，但需要具体实现");
                return false;
            }
            return false;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] FileSystemObject方法失败: " + error.message);
            return false;
        }
    }

    /**
     * 方法4: 使用WPS JSA的Shell对象
     */
    DownloadWithShell(url, filePath) {
        try {
            console.log("[" + this.MODULE_NAME + "] 尝试使用Shell方法");
            
            // 使用WPS JSA的Shell对象执行系统命令
            const shell = CreateObject("WScript.Shell");
            if (shell) {
                // 执行curl命令下载文件
                const command = 'curl -L -o "' + filePath + '" "' + url + '"';
                const result = shell.Run(command, 0, true);
                
                if (result === 0) {
                    console.log("[" + this.MODULE_NAME + "] Shell下载成功");
                    return true;
                }
            }
            return false;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] Shell方法失败: " + error.message);
            return false;
        }
    }

    /**
     * 方法5: 使用WPS JSA的WebClient
     */
    DownloadWithWebClient(url, filePath) {
        try {
            console.log("[" + this.MODULE_NAME + "] 尝试使用WebClient方法");
            
            // 使用System.Net.WebClient
            const webClient = CreateObject("System.Net.WebClient");
            if (webClient) {
                webClient.DownloadFile(url, filePath);
                console.log("[" + this.MODULE_NAME + "] WebClient下载成功");
                return true;
            }
            return false;
        } catch (error) {
            console.log("[" + this.MODULE_NAME + "] WebClient方法失败: " + error.message);
            return false;
        }
    }

    /**
     * 尝试多种下载方法
     */
    TryMultipleDownloadMethods(url, filePath) {
        console.log("[" + this.MODULE_NAME + "] 尝试多种下载方法");
        
        // 方法1: 使用WPS直接打开URL
        if (this.DownloadWithWPSOpen(url, filePath)) {
            return true;
        }
        
        // 方法2: 使用CreateObject
        if (this.DownloadWithCreateObject(url, filePath)) {
            return true;
        }
        
        // 方法3: 使用Shell
        if (this.DownloadWithShell(url, filePath)) {
            return true;
        }
        
        // 方法4: 使用WebClient
        if (this.DownloadWithWebClient(url, filePath)) {
            return true;
        }
        
        console.log("[" + this.MODULE_NAME + "] 所有下载方法都失败了");
        return false;
    }

    /**
     * 格式化日期为yyyy-mm-dd格式
     * @param {Date} date - 日期对象
     * @returns {string} 格式化后的日期字符串
     */
    FormatDate(date) {
        const year = date.getFullYear();
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const day = date.getDate().toString().padStart(2, '0');
        return year + "-" + month + "-" + day;
    }

    /**
     * 检查是否为数值
     * @param {any} value - 要检查的值
     * @returns {boolean} 是否为数值
     */
    IsNumeric(value) {
        return !isNaN(parseFloat(value)) && isFinite(value);
    }
}

// 使用示例：
// const lprDownloader = new LPRDownloader();
// lprDownloader.DownloadAndCopyLPRData();

/**
 * 全局函数 - 用于菜单或按钮调用
 */
function 更新LPR利率() {
    try {
        const lprDownloader = new LPRDownloader();
        return lprDownloader.DownloadAndCopyLPRData();
    } catch (error) {
        console.log("下载LPR数据错误: " + error.message);
        console.log("下载LPR数据时发生错误: " + error.message);
        return false;
    }
}
