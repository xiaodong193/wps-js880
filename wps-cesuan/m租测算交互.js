/**
 * @module m租金测算交互
 * @description 租赁收益测算表交互逻辑模块
 * @version 1.1.0
 * @date 2026-05-11
 */

var M租测算 = (function() {
    var m_worksheet = null;
    var m_config = {
        参数区域: { 行: 3, 列起始: 1, 列结束: 10 },
        测算起始行: 15,
        日期列: 2,
        金额列: 3
    };

    function 获取工作表(sheetName) {
        if ("string" === typeof sheetName) {
            return Worksheets(sheetName);
        }
        return ActiveSheet;
    }

    function 验证参数(sht) {
        var params = {};
        var errors = [];

        params.租赁成本 = sht.Range("B4").Value2;
        params.票面利率 = sht.Range("B5").Value2;
        params.总期数 = sht.Range("B10").Value2;
        params.支付间隔 = sht.Range("B11").Value2;

        if (null === params.租赁成本 || params.租赁成本 <= 0) {
            errors.push("租赁成本必须大于0");
        }
        if (null === params.票面利率 || params.票面利率 <= 0) {
            errors.push("票面利率必须大于0");
        }
        if (null === params.总期数 || params.总期数 <= 0) {
            errors.push("总期数必须大于0");
        }
        if (null === params.支付间隔 || params.支付间隔 <= 0) {
            errors.push("支付间隔必须大于0");
        }

        return {
            valid: errors.length === 0,
            errors: errors,
            params: params
        };
    }

    function 计算租金(params, method) {
        var principal = params.租赁成本;
        var rate = params.票面利率;
        var periods = parseInt(params.总期数, 10);

        if ("等额本息" === method) {
            return 计算等额本息(principal, rate, periods);
        } else if ("等额本金" === method) {
            return 计算等额本金(principal, rate, periods);
        }
        return 0;
    }

    function 计算等额本息(principal, annualRate, periods) {
        var monthlyRate = annualRate / 12;
        if (0 === monthlyRate) {
            return principal / periods;
        }
        var factor = Math.pow(1 + monthlyRate, periods);
        return principal * monthlyRate * factor / (factor - 1);
    }

    function 计算等额本金(principal, annualRate, periods) {
        var monthlyRate = annualRate / 12;
        var monthlyPrincipal = principal / periods;
        var payments = [];
        var remaining = principal;

        for (var i = 0; i < periods; i++) {
            var interest = remaining * monthlyRate;
            var payment = monthlyPrincipal + interest;
            payments.push(payment);
            remaining = remaining - monthlyPrincipal;
        }

        return payments;
    }

    function 生成还款计划(sht, params, method) {
        var startRow = m_config.测算起始行;
        var periods = parseInt(params.总期数, 10);
        var calculation = 计算租金(params, method);

        if ("等额本息" === method) {
            var payment = calculation;
            for (var i = 0; i < periods; i++) {
                var row = startRow + i;
                sht.Cells(row, 1).Value2 = i + 1;
                sht.Cells(row, m_config.金额列).Value2 = payment;
            }
        } else if ("等额本金" === method) {
            var payments = calculation;
            for (var j = 0; j < payments.length; j++) {
                var r = startRow + j;
                sht.Cells(r, 1).Value2 = j + 1;
                sht.Cells(r, m_config.金额列).Value2 = payments[j];
            }
        }
    }

    function 刷新测算(sheetName, method) {
        var sht = 获取工作表(sheetName);
        var validation = 验证参数(sht);

        if (!validation.valid) {
            MsgBox("参数验证失败:\n" + validation.errors.join("\n"));
            return false;
        }

        if (undefined === method) {
            method = "等额本息";
        }

        生成还款计划(sht, validation.params, method);
        return true;
    }

    function 清除数据(sheetName) {
        var sht = 获取工作表(sheetName);
        var startRow = m_config.测算起始行;
        var endRow = startRow + 200;

        sht.Range(sht.Cells(startRow, 1), sht.Cells(endRow, 10)).ClearContents();
    }

    function 导出结果(targetSheet) {
        var sourceSht = 获取工作表(null);
        var targetSht = null;

        try {
            targetSht = Worksheets(targetSheet);
        } catch (e) {
            targetSht = Worksheets.Add(After = Worksheets(Worksheets.Count));
            targetSht.Name = targetSheet;
        }

        var arr = sourceSht.Range("A1").CurrentRegion.safeArray();
        arr.toRange(targetSht.Range("A1"), true);

        return true;
    }

    function 获取利率选择(sht) {
        return sht.Range("D5").Value2;
    }

    function 更新LPR利率(sht) {
        var baseRate = sht.Range("G10").Value2;
        var basePoints = sht.Range("I11").Value2;

        if (null !== baseRate && null !== basePoints) {
            var newRate = baseRate + basePoints / 10000;
            sht.Range("B5").Value2 = newRate;
        }
    }

    function 计算IRR(cashFlows) {
        var irr = 0;
        var step = 0.0001;
        var maxIterations = 10000;
        var tolerance = 0.00001;

        for (var i = 0; i < maxIterations; i++) {
            var npv = 0;
            for (var j = 0; j < cashFlows.length; j++) {
                npv = npv + cashFlows[j] / Math.pow(1 + irr, j);
            }

            if (Math.abs(npv) < tolerance) {
                return irr * 100;
            }

            irr = irr + step;
        }

        return irr * 100;
    }

    return {
        Initialize: function(sheetName) {
            m_worksheet = 获取工作表(sheetName);
            return true;
        },
        Validate: function() {
            return 验证参数(m_worksheet);
        },
        Refresh: function(method) {
            return 刷新测算(null, method);
        },
        Clear: function() {
            return 清除数据(null);
        },
        Export: function(targetSheet) {
            return 导出结果(targetSheet);
        },
        CalculateIRR: function(cashFlows) {
            return 计算IRR(cashFlows);
        },
        UpdateLPR: function() {
            return 更新LPR利率(m_worksheet);
        },
        Config: m_config
    };
})();

function btn刷新测算_Click() {
    Application.EnableEvents = false;
    try {
        var result = M租测算.Refresh();
        if (result) {
            MsgBox("测算刷新完成");
        }
    } catch (e) {
        MsgBox("刷新失败: " + e.message);
    } finally {
        Application.EnableEvents = true;
    }
}

function btn刷新等额本金_Click() {
    Application.EnableEvents = false;
    try {
        var result = M租测算.Refresh("等额本金");
        if (result) {
            MsgBox("等额本金测算完成");
        }
    } catch (e) {
        MsgBox("刷新失败: " + e.message);
    } finally {
        Application.EnableEvents = true;
    }
}

function btn清除数据_Click() {
    if (6 === MsgBox("确定要清除所有测算数据吗？", 4 + 32)) {
        M租测算.Clear();
    }
}

function btn导出结果_Click() {
    var targetName = InputBox("请输入导出工作表名称:", "导出结果", "测算结果_" + JSA.now.replace(/[:\-]/g, "").slice(0, 8));
    if (targetName && targetName.length > 0) {
        var result = M租测算.Export(targetName);
        if (result) {
            Worksheets(targetName).Activate();
            MsgBox("导出成功: " + targetName);
        }
    }
}

function btn计算IRR_Click() {
    Application.EnableEvents = false;
    try {
        var sht = ActiveSheet;
        var startRow = M租测算.Config.测算起始行;
        var cashFlows = [];

        for (var i = 0; i < 50; i++) {
            var cellValue = sht.Cells(startRow + i, M租测算.Config.金额列).Value2;
            if (null !== cellValue && undefined !== cellValue) {
                cashFlows.push(cellValue);
            } else {
                break;
            }
        }

        if (cashFlows.length > 0) {
            var irr = M租测算.CalculateIRR(cashFlows);
            sht.Range("K4").Value2 = irr / 100;
            MsgBox("IRR计算完成: " + irr.toFixed(2) + "%");
        } else {
            MsgBox("未找到现金流数据");
        }
    } catch (e) {
        MsgBox("IRR计算失败: " + e.message);
    } finally {
        Application.EnableEvents = true;
    }
}