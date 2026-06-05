# 链式调用

JSA880 框架全面支持链式调用，让代码更简洁、更易读。

## 基本概念

链式调用允许你在一个表达式中连续执行多个操作：

```javascript
// 普通写法
var filtered = arr.z筛选('f2 > 100');
var sorted = filtered.z多列排序('f1+');
var result = sorted.z结果();

// 链式写法
var result = arr.z筛选('f2 > 100')
               .z多列排序('f1+')
               .z结果();
```

## 支持链式调用的方法

### 筛选操作

```javascript
// 单条件筛选
arr.z筛选('f1 > 100')

// 多条件筛选（AND）
arr.z筛选('f1 > 100')
   .z筛选('f2 === "中国"')

// OR 条件（在同一筛选中）
arr.z筛选('f1 === "A" || f1 === "B"')
```

### 排序操作

```javascript
// 单列排序
arr.z排序('f1+')  // 升序
arr.z排序('f1-')  // 降序

// 多列排序
arr.z多列排序('f1+,f2-')

// 升序/降序排序
arr.z升序排序()
arr.z降序排序()
```

### 切片操作

```javascript
// 跳过我前5行
arr.z跳过(5)

// 只取前10行
arr.z取前N个(10)

// 跳过然后取
arr.z跳过(5).z取前N个(10)

// 间隔取数
arr.z间隔取数(2)     // 每2个取1个
arr.z间隔取数(3, 1)  // 从第2个开始，每3个取1个
```

### 映射转换

```javascript
// 转换所有值
arr.z映射('x => x.map(v => String(v))')

// 提取特定列
arr.z映射('x => x.slice(0, 3)')

// 计算新列
arr.z映射('x => x.concat([x[0] * x[1]])')
```

## 完整示例

### 示例1：数据清洗流程

```javascript
var result = rawData
    .z筛选('f1 !== ""')           // 去除空行
    .z筛选('f3 > 0')              // 只保留正值
    .z多列排序('f1+,f2-')           // 按名称升序、日期降序
    .z选择列([0, 1, 2, 5])         // 只保留需要的列
    .z结果();                       // 获取最终数组

result.toRange('A1');
```

### 示例2：分组统计算法

```javascript
var summary = salesData
    .z筛选('f5 > 1000')           // 筛选大额订单
    .z分组('f1')                    // 按产品分组
    .z映射('g => [g.key, g.sum("f3"), g.count()]')  // 计算每组
    .z排序('f2-')                  // 按销售额降序
    .z结果();

summary.toRange('A1');
```

### 示例3：排名计算

```javascript
var ranked = scoreData
    .z分组排名('f3-', 'f1')        // 按班级分组，按分数降序排名
    .z筛选('f4 <= 10')             // 取前10名
    .z多列排序('f1+,f4+')         // 按班级升序、排名升序
    .z结果();
```

### 示例4：条件统计

```javascript
var stats = departmentData
    .z筛选('f2 === "销售部"')      // 筛选销售部
    .z映射('x => ({部门: x[1], 姓名: x[0], 销售额: x[3]}))')  // 投影
    .z排序('f3-')                  // 按销售额降序
    .z结果();
```

## 实用技巧

### 1. 调试链式调用

```javascript
// 在链中插入日志
var result = arr
    .z筛选('f1 > 100')
    .z映射(x => {
        Console.log('Processing:', x);
        return x;
    })
    .z排序('f2+')
    .z结果();
```

### 2. 条件跳过

```javascript
// 跳过前面连续满足条件的行
arr.z跳过前面连续满足('f1 === 0')
    .z取前面连续满足('f1 > 0')
```

### 3. 分页处理

```javascript
// 每页10条，处理所有页
var page10 = arr.z按行数分页(10)[9];  // 取第10页

// 或遍历所有页
var pages = arr.z按行数分页(100);
pages.forEach((page, i) => {
    page.toRange(`A${i * 100 + 1}`);
});
```

### 4. 与 toRange 链式输出

```javascript
// 直接输出到单元格
arr.z筛选('f1 > 0')
   .z排序('f1+')
   .toRange('A1');

// 或使用 z写入单元格
arr.z筛选('f1 > 0')
   .z写入单元格('D5', true);
```

## 方法返回值

| 方法 | 返回类型 | 说明 |
|------|----------|------|
| `z筛选` | Array2D | 返回新的 Array2D 实例 |
| `z排序` | Array2D | 返回新的 Array2D 实例 |
| `z映射` | Array2D | 返回新的 Array2D 实例 |
| `z跳过` | Array2D | 返回新的 Array2D 实例 |
| `z取前N个` | Array2D | 返回新的 Array2D 实例 |
| `z结果` | Array | 返回原生数组 |
| `toRange` | Range | 返回 Range 对象 |

## z结果() 的重要性

`z结果()` 用于将 Array2D 实例转换回原生数组：

```javascript
// 正确 - 使用 z结果() 获取数组
var arr = data.z筛选('f1 > 0').z结果();

// 错误 - 返回的是 Array2D 对象，不是数组
var wrong = data.z筛选('f1 > 0');  // 这是 Array2D 对象
```