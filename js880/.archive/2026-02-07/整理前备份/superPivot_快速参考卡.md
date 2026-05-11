# superPivot v3.9.0 快速参考卡

## 🚀 最简代码模板

```javascript
var result = Array2D.z超级透视(
    数据,
    ['f1', '行标题'],
    ['f2', '列标题'],
    ['sum("f3")', '值标题']
);
result.toRange("A1");
```

---

## 📋 参数速查

| 参数 | 格式 | 示例 |
|------|------|------|
| **行/列字段** | `'f1,f2'` 或 `['f1,f2', '标题1,标题2']` | `['f1,f2', '大区,省份']` |
| **数据字段** | `'count(),sum("f3")'` | `['sum("f4"),count()', '销售额,订单数']` |
| **排序** | `f1+`(升) `f1-`(降) `f1#`(原序) | `['f1+,f2-', '产品,地区']` |

---

## ⚙️ 常用配置

### 基础配置
```javascript
{
    cornerTitle: '报表标题',
    layoutMode: 'outline',      // compact | outline | tabular
    rowFieldIndent: true,       // 层级缩进
    rowFieldIndentSize: 4       // 缩进空格数
}
```

### 小计总计
```javascript
{
    rowSubtotals: { enabled: true, label: '小计' },
    colSubtotals: { enabled: true, label: '小计' },
    grandTotals: { row: true, column: true, label: '总计' }
}
```

### 百分比
```javascript
{
    displayAs: {
        mode: 'percentOfGrandTotal',  // percentOfRowTotal | percentOfColTotal
        decimals: 2
    }
}
```

---

## 🎯 常见场景代码

### 场景1：简单透视
```javascript
Array2D.z超级透视(data, 'f1', 'f2', 'sum("f3")');
```

### 场景2：多层行列
```javascript
Array2D.z超级透视(
    data,
    ['f1,f2', '年,月'],
    ['f3,f4', '大区,省份'],
    ['sum("f5")', '销售额']
);
```

### 场景3：带小计总计
```javascript
Array2D.z超级透视(
    data, ['f1', '产品'], ['f2', '年份'], ['sum("f3")', '销售额'],
    1, 1, '@^@',
    { rowSubtotals: {enabled: true}, grandTotals: {row: true, column: true} }
);
```

### 场景4：百分比显示
```javascript
Array2D.z超级透视(
    data, ['f1', '产品'], ['f2', '年份'], ['sum("f3")', '占比'],
    1, 1, '@^@',
    { displayAs: {mode: 'percentOfGrandTotal', decimals: 2} }
);
```

### 场景5：多聚合
```javascript
Array2D.z超级透视(
    data, ['f1', '产品'], ['f2', '年份'],
    ['count(),sum("f3"),average("f3")', '数量,总和,平均']
);
```

---

## 📊 返回值方法

| 方法 | 用途 |
|------|------|
| `toRange("A1")` | 写入单元格 |
| `toRange("A1", true)` | 写入并合并单元格 |
| `getMeta()` | 获取元数据 |
| `val()` / `res()` | 获取原始数组 |

---

## 🔍 调试代码

```javascript
var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);

// 查看基本信息
Console.log("总行数: " + result.length);
Console.log("表头: " + JSON.stringify(result[0]));

// 查看元数据
var meta = result.getMeta();
Console.log(JSON.stringify(meta, null, 2));
```

---

## ❗ 注意事项

1. **数据格式**：第一行必须是表头
2. **字段引用**：使用 `f1`, `f2` 表示第1列、第2列
3. **数值类型**：数据字段列必须是数值
4. **内存限制**：超大数据建议分批处理

---

## 📞 问题排查

| 问题 | 解决方案 |
|------|----------|
| 结果为空 | 检查 `headerRows` 参数是否正确 |
| 数值不对 | 确保数据列没有文本字符 |
| 排序无效 | 使用 `f1+` / `f1-` / `f1#` |
| 输出慢 | 大数据量禁用屏幕更新 |

---

**版本**: v3.9.0 | **适用**: WPS JSA
