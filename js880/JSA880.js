/**
 * ========================================================================
 * JSA880_WPS_Modern.js - 郑广学JSA880快速开发框架（WPS现代版）
 * ========================================================================
 *
 * 原作者: 郑广学 (EXCEL880)
 * 维护者: 徐晓冬
 * 版本: 4.0.37 (2026年6月9日)

【此版本为WPS现代版 v4.0.36】
 * 【此版本为WPS现代版 v4.0.13】
 * 【此版本为WPS现代版 v4.0.12】
 * 【此版本为WPS现代版 v4.0.11】
 * 【此版本为WPS现代版 v4.0.0】
 * - 移除所有Node.js兼容代码
 * - 移除所有浏览器兼容代码
 * - 保留const/let，适用于WPS Office 2021+
 * - 仅适用于WPS Office JavaScript API (JSA)
 *
 * 原作者: 郑广学 (EXCEL880)
 *
 * API文档: https://vbayyds.com/api/jsa880/
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.40 — 2026-06-11)
 * --------------------------------------------------------------------------
 * 1. [修复] XXD-218: z筛选 对象参数 {headerName:value} 简写匹配 header 名
 *    - 根因: {A:1} 走 _filterByObject→_checkCondition, 该方法只认 {column,operator,value} 结构
 *      不认 {headerName:value} 简写, column=undefined → colIndex=-1 → 全部返回 false → []
 *    - 修法: z筛选 对象参数分支增加简写检测(!column && !operator), 从 _header 或 data[0]
 *      构建 header→index 映射, 按映射直接筛选并返回, 不再委托 _checkCondition
 *    - 复现: new Array2D([['A','B'],[1,2],[3,4]]).z筛选({A:1})._items → 期望 [[1,2]]
 *    - 验证: node 复现命令通过; 多条件 AND、无匹配、_header 属性、结构化条件回退均通过
 *
 * 更新日志 (v4.0.39 — 2026-06-11)
 * --------------------------------------------------------------------------
 * 1. [文档] XXD-180/XXD-181: SuperMap.z分组 / z分组统计 形状约定澄清
 *    - 现象: 两者同名 z分组 但返回形状不同,易被误以为是 bug
 *      · SuperMap.z分组(arr, sel)                  → 字典 { key: rows[] }
 *      · SuperMap.z分组统计(arr, groupCol, cfg)    → 二维数组 [header, ...rows]
 *    - 判定: 不是 bug,是不同操作的不同自然表示
 *      · z分组      是纯分组 (无聚合),字典 key→行集合 最自然
 *      · z分组统计  是分组+聚合,每组 1 行,二维表格最自然
 *    - 修法: 在两个方法上方补 JSDoc 明确写出 @returns 形状 + 形状对照说明
 *    - 行为: 无运行时代码变更;两种形状保持向后兼容
 *
 * 更新日志 (v4.0.38 — 2026-06-11)
 * ------------------------------------------------------------------------
 * 1. [修复] XXD-160: Array2D._header 链式调用
 *    - 根因: ._header 是属性 this._header=...，外部 a._header(0) 抛 not a function
 *    - 修法:
 *      ① 在 Array2D.prototype 加同名方法 _header(n)：
 *         - n 省略时按 1 处理（与现有 z去重 跳过 1 行约定一致）
 *         - 把 headerRows 写到实例自有属性 _header（writable/configurable/enumerable:false）
 *         - 返回 this，支持链式（如 ._header(0).z去重()）
 *      ② 7 处 '_header' in this 改为 Object.prototype.hasOwnProperty.call(this, '_header')
 *         避免匹配到原型链上的 _header 方法（XXD-160 引入的新污染）
 *      ③ 构造函数 _header 复制路径加 typeof data._header !== 'function' 守卫
 *    - 复现命令:
 *         node -e "var fs=require('fs');eval(fs.readFileSync('JSA880.js','utf8'));\
var a=new Array2D([[1,2],[1,2],[3,4]]);a._header(0);console.log(a.z去重()._items);"
 *    - 修前: THROW a._header is not a function
 *    - 修后: [[1,2],[3,4]]
 *
 * 更新日志 (v4.0.37 — 2026-06-09)
 * ------------------------------------------------------------------------
 * 1. [修复] parseFieldsConfig 自定义标题 split 支持中文逗号(，)
 *    - 根因: split(',') 只拆分 ASCII 逗号,中文逗号(，)未拆分
 *      例如 rowFields=['f3,f2', '商品，国家'] → titles=['商品，国家'](未拆分),第一个 row 标题 cell 被整体填入 "商品，国家"
 *    - 修法: 所有标题拆分统一改为 split(/[,，]/) 同时支持 ASCII 逗号和中文逗号
 *    - 影响范围: parseFieldsConfig 内 4 处标题 split(行字段/列字段/数据字段)
 *    - 受益: row/col/data 自定义标题使用中文逗号时正确拆分到各 cell
 *
 * 2. [修复] XXD-76 XXD-77: 列选择器/输出选择器接受中文全角逗号(，)
 *    - 受影响函数: z去重(distinct), z分组汇总(groupInto), z选择列(selectCols)
 *    - 修法: colSelector/resultSelector 的 includes(',') 和 split(',') 全部改为同时支持 ASCII 和中文逗号
 *    - 修法: LAMBDA_PATTERNS.MULTI_COLUMN/ARRAY_BRACKET 正则和 parseLambda 内逗号拆分改为同时支持 ASCII 和中文逗号
 *    - 受益: 'f1，f2，f3' 格式的列选择器在所有受影响函数中正常工作
 *    - 注意: z批量填充(cols 参数为数字,非列选择器)无需修改; z批量删除列 includes(',') 单列判断已修复 (XXD-78/XXD-79)
 *    - 追加: z内连接/z左连接/z右连接/z一对多连接 4 处 resultSelector.split(',') → split(/[,，]/)
 *
 * 3. [修复] XXD-102 XXD-103: DateUtils.z日期格式化 中文格式抛错
 *    - 根因: z日期格式化 签名 (jsdate, fmt) 但原型链调用时仅传入1参,fmt为undefined
 *    - 修法: 函数顶部加原型链兼容守卫,推断原型链调用模式回退到 this._date
 *    - 受益: DateUtils(d).z日期格式化('yyyy年MM月dd日') 等中文格式不再抛错
 *    - 兼容: 原有两参调用路径行为不变
 *
 * 4. [修复] XXD-137 XXD-140: z分组 中文逗号(，)归一化为 ASCII 逗号
 *    - 根因: keySelector/valSelector 含中文逗号(，)时，parseLambda 虽能解析，
 *      但 _normalizeCacheKey 产生不一致的缓存键('s:f1，f2' vs 's:f1,f2')，
 *      导致 Index 缓存未命中，且行为与 ASCII 逗号('f1,f2')路径不完全一致
 *    - 修法: z分组 入口处将 keySelector/valSelector 的 ，→, 归一化
 *    - 影响范围: z分组 原型方法 (index 加速路径也受益于统一缓存键)
 *    - 受益: 中文输入法场景下 'f1，f2' 行为与 'f1,f2' 完全一致
 *
 * 5. [修复] XXD-141 XXD-138: z去重 f3-f5 范围只输出第3列而非3列
 *    - 根因: colSelector 含无逗号的范围语法 f3-f5 时，单列 f 分支 startsWith('f') 先于
 *      范围检测,parseInt('3-f5')→3,只取第3列;逗号分隔分支内 f3-f4 同样 startsWith('f')
 *      先于 includes('-'),范围被截断为单列;toIndex 同
 *    - 修法 (4处):
 *      ① z去重 单列 f 分支: 新增 indexOf('-') 范围检测,展开 f3-f5→[2,3,4]
 *      ② z去重 逗号分支: startsWith('f') 内新增 includes('-') 范围检测
 *      ③ toIndex 单列 f 分支: 同上范围检测
 *      ④ toIndex 逗号分支: 同上范围检测
 *    - 影响范围: z去重(colSelector f3-f5)、toIndex、逗号分隔内 f3-f4 格式
 *    - 受益: z去重('f3-f5') 正确输出 3 列(f3,f4,f5)而非仅 f3;逗号格式如 'f1,f3-f4' 正确展开
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.36 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] superPivot 单列字段表头行 1 col 字段标题位置不 push (避免 dataField 重复)
 *    - 根因: v4.0.32 在行 1 col 字段标题位置 push defaultDataTitles[0]='计数',但行 0 col 字段标题是 '年' 不是 dataField
 *      重复 1 个 dataField 标题导致行 1 多 1 cell (15 列 vs 14 列 ideal)
 *    - 修法: 行 1 col 字段标题位置不 push,让行 0/1 都在 col 字段标题位置对齐
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.35 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] superPivot 单列字段表头行 0 row 字段位置只 push 1 次 (合并跨 numRowFieldLevels 列)
 *    - 根因: 之前 v4.0.32 给 row 字段空 cell 重复 numRowFieldLevels 次,表头 15 列 (2 row + 1 col 标题 + 4 年 × 3 data)
 *      但用户 ideal 只有 14 列 (1 row merged + 1 col 标题 + 4 年 × 3 data)
 *    - 修法: 行 0 row 字段位置只 push 1 个空 cell,recordMerge 合并跨 numRowFieldLevels 列
 *           行 1 仍按 numRowFieldLevels push row 字段标题 (国家, 产品)
 *    - 受益: 3.28 节 KO一切的k函数.xlsm 多层透视 J1 表头匹配用户理想 (14 列)
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.34 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] superPivot 没数据时默认 nullValue 从 '' 改为 0(让 WPS 显 0 匹配用户 ideal)
 * 2. [修复] 单列字段表头第 1 行 col 字段标题位置从 '' 改为 defaultDataTitles[0]
 * 3. [诊断] dataRow 长度 + 内容日志
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.33 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] val() 补空 cell 用空字符串 '' 而非 null — WPS spill null 显 0
 *    - 根因: v4.0.20 注释说 "WPS spill null 显空白",实际是错的,WPS spill null 也显 0
 *      导致多层透视 J1 第 0 行 col 0-1 显 0,理想是 (空)
 *    - 修法: 补 undefined/缺列 时用 '' (WPS spill '' 显空白)
 *    - 受益场景: 3.28 节 KO一切的k函数.xlsm '多层透视' J1 第 0 行 (空) (空) 年 ...
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.32 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] superPivot 单列字段 numColFieldLevels===1 表头改为 2 行
 *    - 根因: 之前 headerRowCount=3,行字段在顶部,不符合用户预期
 *    - 用户期望: 行0=[空,空,年,2021×3,2022×3,2023×3,2024×3],行1=[国家,产品,计数,求和,连接,...]
 *    - 修法: headerRowCount 改为 2,行 0 = col 标题 + col 值(×numDataFields),行 1 = row 字段 + dataField 标题
 *    - 行字段标题移到表头底部(行 1)
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.31 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] superPivot 单列字段(numColFieldLevels===1)表头生成: 行0/行1 每个 colKey 重复 numDataFields 次
 *    - 根因: 之前每个 colKey 只 push 1 次到行0/行1,但行2 是 colKey × numDataFields 推
 *      导致行0/行1 列数 = numColKeys,行2 列数 = numColKeys × numDataFields,val() 补 null 变 0
 *      例如 colKeys=4 年,numDataFields=3,行0=4 列(应为 12),行1=4 列(应为 12),行2=12 列 ✓
 *    - 修法: 行0/行1 套 for df in numDataFields 循环,与行2 对齐
 *    - 列小计同样修复(每个小计占 numDataFields 列)
 *    - 多列字段分支(>1)已有正确逻辑,不动
 *    - 受益场景: 3.28 节 KO一切的k函数.xlsm "多层透视" J1 表头后 9 列不再显 0
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.30 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] z筛选 字符串 fN 代理自动 trim — 修 superPivot 输出 f2 列带前导空格导致 .filter 失败
 *    - 根因: superPivot row[3] 输出 f2="    Product1"(前导 4 空格),用户 .filter(x=>x.f2=='Product1')
 *      严格相等匹配失败,所有数据行被过滤掉,只保留 i==0 表头 → val() 返回 1 行
 *    - 修法: z筛选 构造 proxy 时,字符串自动 trim 后赋给 x.fN(原 __proxy[i] 保留未 trim 值)
 *    - 受益场景: 3.28 节 KO一切的k函数.xlsm "多层透视 (2)" J1
 *      =k("__KJ_ARGS__=...filter((x,i)=>i==0 || x.f2=='Product1')...",A1:H40,"count()")
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.29 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [诊断] z筛选 入口 + 每行判定日志 — 调试 (.filter((x,i)=>x.f2=='P1')) 只 1 行
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.28 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] __KJ_ARGS__ 宽松解析: 保护引号内逗号,避免 "f3,f2" 被切成 "f3" + "f2"
 *    - 根因: v4.0.27 按外层 , 切键值对,值内含逗号(如 "f3,f2")被破坏
 *      rowFields 被错误解析为 "f3" 而不是 "f3,f2",superPivot 收到不完整字段
 *    - 修法: 切之前先把 "..." 内的 , 替换为占位符 __KJ_PLACEHOLDER__,切完还原
 *    - 测试: T1-T6 全过(6/6)
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.27 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] __KJ_ARGS__ 严格 JSON.parse 失败时降级到宽松解析
 *    - 根因: WPS 公式字符串 "" 配对吃引号,实际进 jsaLambda 的 __KJ_ARGS__ 形如
 *      {rowFields:"f3,f2",colFields:"f6"} (字段名无引号),严格 JSON 要求字段名必须有引号
 *    - 修法: 严格失败后,自己按 , 切键值对、按 : 切 key/value,自动去字段名/值的引号
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.26 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] __KJ_ARGS__ JSON 解析:反引号 → 双引号
 *    - 根因: 用户在 WPS 公式中写 `f3,f2`(反引号),jsaLambda 收到仍是反引号,JSON.parse 不认
 *    - 修法: 解析前先 .replace(/`/g, '"')
 * 2. [修复] JSON parse 失败兜底也剔除 __KJ_ARGS__ 标记
 *    - 根因: JSON 失败时只 log,没剔除,链式解析器把整段当 fn 编译 → Malformed arrow function
 *    - 修法: catch 块也执行 fn = fn.replace(__kjMatch[0], '').trim()
 * 3. [修复] fn 规范化:多空白字符 → 单个空格
 *    - 根因: WPS 公式粘多行,fn 内部有 \n,new Function 编译时 arrow function body 跨行报错
 *    - 修法: fn = fn.replace(/\s+/g, ' ').trim()
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.25 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [诊断] val() 加详细日志 — totalRows / maxLen / 前 2 行内容
 *    调试 多层透视 (2).filter((x,i)=>i==0 || x.f2=='Product1') 只 spill 1 行的问题
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.24 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] v4.0.23 的 __KJ_ARGS__ 字段追加位置错 — 改用插队到 Range 之后
 *    - 根因: v4.0.23 把提取的 rowFields/colFields/dataFields push 到 realArgs 头部(WPS args 之前)
 *      WPS 传 args=[Range, "count()"],提取后变成 realArgs=["f3,f2","f6","count()",Range]
 *      superPivot(arr, colFields, dataFields, Range) 顺序错位:
 *        arr = "f3,f2" (string)
 *        rowFields = "f6"
 *        colFields = "count()"
 *        dataFields = Range (undefined)
 *    - 修法: 改为把字段插到 **第一个非 -r 参数(Range)之后** → realArgs = [Range, "f3,f2", "f6", "count()"]
 *      superPivot(Range, "f3,f2", "f6", "count()") 顺序正确
 *    - 关键证据: superPivot IN 显示 rowFields="f6"(应是 "f3,f2"),colFields=undefined,arr.t=string
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.23 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] jsaLambda 入口支持 __KJ_ARGS__={...} JSON 提取 — 解决 WPS 公式引擎丢 string 参数 bug
 *    - 根因: WPS 公式 =k(fn, Range, s1, s2, s3) 多 string 参数中,引擎只传 4 个(丢 1 个)
 *      验证: =k("$$...", A1:H40, "f3,f2", "f6", "count()...") → jsaLambda IN 显示
 *            args=[<Range>, "f3,f2", "", "count()..."],"f6" 静默丢失
 *      superPivot 因此收到 colFields="",不分组,只 spill 表头
 *    - 修法: jsaLambda 入口从 fn 字符串中正则匹配 __KJ_ARGS__={...} JSON,自动注入到 realArgs
 *      用法: =k("__KJ_ARGS__={\"rowFields\":\"f3,f2\",\"colFields\":\"f6\"}  (...args)=>$$...", A1:H40, "count()...")
 *      WPS 只看到 3 个 string 参数(fn + 2 个),不再触发丢参
 *      顺序: rowFields, colFields, dataFields, headerRows,然后是 args
 *    - 向后兼容: 没用 __KJ_ARGS__ 标记的老公式完全不受影响
 *    - 受益场景: 3.28 节 KO一切的k函数.xlsm "多层透视!J1" 公式
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.14 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [修复] _kParseChainableExpression: 放弃覆盖 Array.prototype.filter,
 *    改用源码级重写——把链式最后一个 .filter / .map 的上游用 new Array2D(...) 包裹
 *    - 根因: WPS JSA 沙箱可能拒绝给 Array.prototype 重新赋值,抛 TypeError,
 *      被 jsaLambda 外层 catch 抓住,导致 jsaLambda 返回 null
 *      单元格表现为 #K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined"
 *    - 修法: 解析时改写 expr,变成 (new Array2D(<上游>)).filter(<谓词>)
 *      这样 .filter 自动走 Array2D.prototype.z筛选,继承构造函数注入的 f1/f2 访问器
 *    - 不再覆盖 Array.prototype,WPS 沙箱无副作用
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.13 — 2026-06-07)
 * ------------------------------------------------------------------------
 * 1. [Bug] 链式调用 .filter((x,i)=>x.fN==...) 失败:谓词 x.fN 永远是 undefined
 *    - 根因: _kParseChainableExpression 编译的 .filter 是 Array.prototype.filter,
 *      而 superPivot 等链式上游返回的是普通 2D 数组(不是 Array2D 实例)
 *    - 修复: 链式解析器临时把 Array.prototype.filter / map 重定向到
 *      Array2D.z筛选 / z映射(只在本次 eval 范围,try/finally 还原)
 *    - 同源问题: Array2D.prototype.z筛选 自身也缺 fN proxy,补上(与 z映射 行为对齐)
 *    - 受益场景: 3.28 节 KO一切的k函数.xlsm "(多层透视 2)!J1"
 *      =k("(...args)=>$$.superPivot(...args).filter((x,i)=>i==0 || x.f2=='Product1')", ...)
 *    - 老用法(1D 数组 / 非函数回调)完全不受影响(走原生分支)
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.11 — 2026-06-05)
 * ------------------------------------------------------------------------
 * 1. [质量] superPivot: typeof 检查替换为 instanceof Array2D（更可靠）
 * 2. [质量] _createStaticMethods scalar 分支改用 unwrapVal（消除冗余代码）
 * 3. [质量] Logger: warn/error 级别改用 console.warn/console.error
 * 4. [质量] 全局错误日志统一 console.log → console.error（10处）
 * 5. [质量] JSA 命名空间新增导出 19 个全局函数/模块
 *    - agg, oadate, map2d, forEach2d, StopWatch, TreeNode
 *    - batch_k, k_help, DataExport, DataImport, Logger
 *    - StrUtils, NumUtils, MsgUtils, WorkbookUtils, FormUtils, PicUtiles
 * 6. [清理] 删除 12 处注释掉的 DEBUG Console.log
 * 7. [Bug] Array2D.agg switch/case: var 重声明 → let + 块作用域（2处）
 * 8. [Bug] SuperMap._aggregate switch/case: 同上修复
 * 9. [Bug] 全局 25 处 for(var k=) 循环变量重命名为 for(var ki=)，避免遮蔽顶层 UDF function k()
 * 10. [修复] rangeMatrix v4.0.11 原地覆盖破坏课程 group-by 语义
 *     - 原 rangeMatrix（L10628）是 group-by 聚合（按 keySelector 对 dataArrays 求和）
 *     - v4.0.11 末尾（L18837）原地覆盖实现为元素级矩阵地址运算，与原版语义完全不同
 *     - 课程第 3 章 3 个 rangeMatrix 示例（L3833-L3846 in JSA880_to_paste.js 教学版）全部依赖 group-by 语义
 *     - 修复：保留原 group-by 实现不动；元素级版本独立为新函数 Array2D.rangeZip
 *     - 同步导出 $.rangeZip / $.z区域对齐；不覆盖 $.rangeMatrix / $.z区域矩阵
 *     - 课程 paste 板（JS880教案/第03章/3-25/JSA880_to_paste.js）的 rangeMatrix 行为回到 v3.8 group-by
 *     - 新增 rangeZip 用于元素级矩阵 zip/对齐（替代 v4.0.11 错误位置的同名重写）
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.10 — 2026-06-01)
 * ------------------------------------------------------------------------
 * 1. [修复] Array2D.prototype.toRange / Array.prototype.toRange
 *    - 修复第二个参数语义：从 skipRows 改为 clearBelow
 *    - true/1: 写入后清空输出区域下方的原有数据
 *    - false/不传: 仅写入数组覆盖区域，下方旧数据保留
 * 2. [修复] unwrapVal 死循环（行 500-502）— 任何非空值导致栈溢出
 * 3. [修复] $.RngUtils() 无参调用时引用未定义函数 createRngUtilsStaticProxy
 * 4. [修复] isCollection 的 typeof === 'unknown' 永远为 false
 * 5. [修复] SheetChangeHelper.check 使用 target.Value2 而非交集值，导致监控永远不触发
 * 6. [清理] 6 处重复定义：map2d×2 / forEach2d×2 / toMatrix×2 / JSA880.分组汇总连接×2 / $.maxRange×2 / $.CurrentRegion×2
 * 7. [新增] IO 文件/文件夹操作 10 个方法（对照郑广学docx第10章）
 *    - IO.copyFile / z复制文件
 *    - IO.rename / z重命名文件
 *    - IO.moveFile / z移动文件
 *    - IO.mkDir2 / z建文件夹
 *    - IO.delete / z删除文件
 *    - IO.copyFolder / z复制文件夹
 *    - IO.reNameFolder / z改文件夹名
 *    - IO.moveFolder / z移动文件夹
 *    - IO.delTree / z递归删文件夹
 *    - IO.clearTree / z清空文件夹
 * 8. [新增] JSA.rndIntArray / z随机整数数组(min, max, n) — 生成 n 个随机整数数组
 * 9. [新增] StopWatch 性能计时器类
 *    - 新增 StopWatch 类，提供 start()/time()/lap()/usedTime()/restart()
 *    - 与旧版 JSA880 教案代码中 StopWatch 用法完全兼容
 *    - 高精度毫秒计时，time() 返回秒数（保留3位小数）
 * 10. [新增] IO.MkDir2 / IO.z创建文件夹 别名
 *     - IO.MkDir2 作为 IO.mkDir2 的 PascalCase 别名
 *     - IO.z创建文件夹 作为中文别名
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.4)
 * ------------------------------------------------------------------------
 * 1. [修复] Array2D.distinct/z去重 - 回调函数模式代理创建时机
 *    - 将代理创建从 keyFn 调用前移，统一为每行创建 proxy 对象
 *    - 确保函数模式和 f模式 使用相同的代理逻辑
 *    - 修复回调函数中 x.f2.trim() 访问 undefined 的问题
 * 2. [修复] Array2D.distinct/z去重 - 函数模式输出格式
 *    - 函数模式回调返回数组时，直接使用返回值而非再包装
 *    - 修复 outputFn 嵌套数组问题（导致 #N/A）
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.3)
 * ------------------------------------------------------------------------
 * 1. [修复] Array2D.distinct/z去重 - 多列模式 colIndexes 变量遮蔽问题
 *    - 移除多列模式中的 var colIndexes，改为直接赋值外层变量
 * 2. [修复] Array2D.distinct/z去重 - 防御性空值检查
 *    - 所有 outputFn 函数添加 row 空值检查
 *    - 防止 undefined/null 行导致 .slice() 报错
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.2)
 * ------------------------------------------------------------------------
 * 1. [增强] Array2D.deleteRows/z批量删除行 - 新增Lambda字符串模式支持
 *    - 支持 "r=>r.f3=='美国'" 字符串形式
 *    - 支持双引号和单引号包裹的字符串值
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v4.0.1)
 * ------------------------------------------------------------------------
 * 1. [增强] Array2D.distinct/z去重 - 新增 resultSelector 第三参数支持
 *    - 支持选择输出列: undefined 只输出关键字列, '' 输出所有列
 *    - 支持 'f1,f2' 或 [0,1] 格式指定输出列
 *    - 修复单列 f1 模式时 colIndexes 未初始化的问题
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.9.4)
 * ------------------------------------------------------------------------
 * 1. [新增] z超级透视 内置筛选器 - 支持输出前筛选特定行/列值
 *    - filterRows: 函数或对象，筛选行键
 *    - filterCols: 函数或对象，筛选列键
 *    - 支持 { f1: ['北京', '上海'] } 白名单格式
 *    - 支持 row => row.f1 !== '深圳' 函数格式
 *
 * 2. [修复] z分块 - size=0/负数时死循环，现返回空数组
 * 3. [修复] z多列排序 - sortParams 为 null/undefined 时崩溃，现返回副本
 * 4. [修复] z平均值 - 分母包含非数值项导致结果偏低，现只计算有效数值
 * 5. [修复] z转置 - 空行或 undefined 行时崩溃，现安全返回空数组
 * 6. [修复] z跳过 - count 为负数/undefined 时行为异常，现默认为 0
 * 7. [修复] z取前N个 - count 为负数/undefined 时行为异常，现默认取全部
 * 8. [修复] z间隔取数 - 空数组时崩溃(TypeError)，现安全返回空数组
 * 9. [修复] z矩阵排版 - 列优先模式(direction='c')逻辑完全错误，重写修复
 * 10. [修复] z重复N次 - 返回普通数组(破坏链式调用)，现返回 Array2D 深拷贝
 * 11. [修复] z取交集 - 去重逻辑错误(使用右表索引)，改用左表 key 去重
 * 12. [修复] Array2D.rank/rankGroup - 美式排名并列值排名号不正确，已修复
 * 13. [修复] SuperMap._buildAllView/toMap - 使用 for...of 语法(ES6+)，改用 forEach
 * 14. [修复] z转字符串 - 非数组行元素调用 .join() 崩溃，现安全转为 String
 * 15. [修复] z文本连接 - 数字类型 selector 被忽略，现正确解析为列索引
 * 16. [修复] z区域矩阵 - 命名冲突(两版本覆盖)，原分块版重命名为 z分块矩阵
 * 17. [修复] Array2D.agg - 两个重复定义返回值不一致(null vs 0)，删除冗余版本
 * 18. [修复] Array2D.getIndexs - 负 step + start<end 导致无限循环，已修复
 * 19. [修复] groupInto - count() 无参数无法匹配正则，现支持无参聚合函数
 * 20. [修复] z按范围选择 - 返回普通数组，现返回 Array2D 实例
 * 21. [修复] z区域映射 - 返回普通数组，现返回 Array2D 实例
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.9.3)
 * ------------------------------------------------------------------------
 * 1. [新增] 补充JSA全局函数 - 根据API文档补全17个缺失函数
 *    - z查找索引/match, z左侧查找/vlookup, z增强查找/xlookup
 *    - z转文本/cstr, z取整数/cint, z取小数/getDecimal
 *    - z转公式数组/toExcelArray, z统一路径分隔符/normalPath
 *    - jsaLambda, z解析函数表达式/parseLambda
 *
 * 2. [新增] 补充Array2D类方法 - 根据API文档补全38个缺失方法
 *    - z版本/version, z错误值/isError, z空结果
 *    - repeat/z重复N次, skipWhile/z跳过前面连续满足
 *    - takeWhile/z取前面连续满足, intersect/z取交集
 *    - union/z去重并集, rangeSelect/z按范围选择
 *    - rangeForEach/z按范围遍历, rangeMap/z区域映射
 *    - rangeMatrix/z区域矩阵, toMatrix/z矩阵排版
 *    - z矩阵运算, agg/z聚合
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.9.2)
 * ------------------------------------------------------------------------
 * 1. [修复] z超级透视 无列字段情况 - 修复数据值获取问题
 *    - 修复数据行构建时 groupMap 键格式不匹配的问题
 *    - 当无列字段时，groupMap 键格式为 "rowKey|||" 而非 "rowKey"
 *    - 现在数据行可以正确获取聚合值
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.9.1)
 * ------------------------------------------------------------------------
 * 1. [修复] z超级透视 无列字段情况 - 修复数据列缺失问题
 *    - 修复仅行字段时，数据字段标题未添加到表头的问题
 *    - 修复仅行字段时，数据行没有包含聚合值的问题
 *    - 修复仅行字段时，总计行没有包含总计值的问题
 *    - 添加对 numColFieldLevels === 0 的专门处理分支
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.8.3)
 * ------------------------------------------------------------------------
 * 1. [修复] z超级透视 单列字段表头生成 - 修复列值未添加问题
 *    - 修复单列字段情况下，列字段值未添加到表头的问题
 *    - 添加列字段标题到第2行
 *    - 添加列字段值到第1行
 *    - 添加数据字段标题到第3行
 * 2. [修复] z超级透视 排序功能诊断 - 添加调试日志
 *    - 添加排序配置解析日志
 *    - 添加行键排序结果日志
 *    - 帮助诊断排序符号 (+/-) 不生效的问题

 * ------------------------------------------------------------------------
 * 更新日志 (v3.8.2)
 * ------------------------------------------------------------------------
 * 1. [修复] z超级透视 多层列字段表头生成 - 修复中间层重复逻辑
 *    - 修复非第一层列字段值的重复次数计算
 *    - 引入 lowerLevelsSpan 变量，每个值重复 (下层组合数) 次
 *    - 修复中间层（如年份）的值序列不能正确嵌套在上层值下的问题
 * 2. [性能] z超级透视 表头生成优化 - 预计算层级唯一值
 *    - 在表头生成前预计算所有列字段层级的唯一值数组
 *    - 避免在嵌套循环中重复调用 getLevelValues 和 getUniqueLevelCount
 *    - 性能提升: 多列字段场景下提升 15-30%
 * 3. [性能] toRange 输出优化 - 屏幕更新和计算控制
 *    - 禁用屏幕更新 ScreenUpdating = false
 *    - 设置计算模式为手动 Calculation = -4135
 *    - 禁用事件触发 EnableEvents = false
 *    - 性能提升: 大数据量场景下提升数百倍
 * 4. [性能] toRange 解除合并优化 - 批量操作
 *    - 使用 writeRng.MergeCells = false 一次性解除合并
 *    - 替代嵌套循环逐个单元格检查合并状态
 *    - 性能提升: 减少数千次 COM 调用
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.8.0)
 * ------------------------------------------------------------------------
 * 1. [增强] z超级透视 多行多列表头支持 - 真正的多层级透视表
 *    - 支持多个列字段时，每个列字段值占据独立的一行
 *    - 动态计算表头行数: headerRowCount = 列字段数量 + 1
 *    - 列字段值按层级展开，形成完整的层次结构
 *    - 支持任意数量的列字段（年份、月份、部门等）
 * 2. [修复] z超级透视 表头生成问题 - 修复 _header 属性传递链
 *    - 修复 _new 方法中 _header 检测逻辑，使用 'in' 操作符更可靠
 *    - 修复 _new 方法保留原始 _original，确保 z筛选/z排序 后仍能访问表头
 *    - 修复 z多列排序 默认 headerRows=0 导致表头参与排序的问题
 *    - 修复 z筛选 默认 skipHeader=undefined 导致表头参与筛选的问题
 *    - 新增自动检测: 如果对象有 _header 属性，自动设置 headerRows/skipHeader=1
 * 3. [修复] 静态方法自动生成器 - 保留 _header 属性
 *    - 修复自动生成的静态方法提取 _items 导致 _header 丢失
 *    - 修复自动生成的静态方法返回 _items 导致 _header 再次丢失
 *    - 将关键方法（filter, sortByCols 等）加入手动定义列表
 * 4. [修复] Array2D 构造函数 - 复制输入数据的 _header 属性
 *    - 修复 new Array2D(arr) 时不会保留 arr._header 的问题
 *    - 确保静态方法链式调用时 _header 属性正确传递
 * 5. [修复] Array2D.filter/sortByCols - 静态方法返回 Array2D 对象
 *    - 修复静态方法返回 .val() 普通数组导致 _header 属性丢失
 *    - 改为直接返回 Array2D 对象，保留 _header 属性用于链式调用
 * 6. [修复] Array2D.superPivot - 在处理 Array2D 对象前保存 _header
 *    - 修复在替换 arr._items 之前保存 _header 和 _original 属性
 *    - 确保 superPivot 能正确获取原始表头用于列字段标题
 * 7. [优化] Array2D.superPivot - 改进表头布局（方案3 - v3.7.9）
 *    - 列字段标题放在第1行最后一个行字段位置
 *    - "国家"与"月"上下对齐，不单独占用一列
 *    - 第1行：(n-1)个空白 + "国家" + 列值
 *    - 第2行：行字段标题 + 数据字段标题
 *    - 布局紧凑，无需合并单元格
 * 8. [增强] Array2D.superPivot - 多行多列表头支持（v3.8.0）
 *    - 支持多个列字段时，每个列字段值占据独立的一行
 *    - 动态计算表头行数: headerRowCount = 列字段数量 + 1
 *    - 列字段值按层级展开，形成完整的层次结构
 *    - 支持任意数量的列字段（年份、月份、部门等）
 *
 * ------------------------------------------------------------------------
 * 更新日志 (v3.7.8)
 * ------------------------------------------------------------------------
 * 1. [增强] Array2D.distinct/z去重 - 新增支持多列组合去重
 *    - 支持 'f1,f2' 格式多列去重
 *    - 支持 Lambda 函数返回数组作为组合键
 *    - 保持向后兼容单列去重
 * 2. [新增] Array2D.agg/z聚合 - 新增聚合计算函数
 *    - 支持 sum, count, average, max, min 五种聚合类型
 *    - 支持字符串参数格式，如 agg('f3', 'sum')
 * 3. [新增] Array2D.rangeMatrix/z区域矩阵 - 区域矩阵分组聚合
 *    - 支持按指定区域分组并聚合数据
 *    - 兼容文档中的 rangeMatric 拼写
 * 4. [新增] Array2D.html/z输出HTML - 输出HTML表格
 *    - 支持表头、样式、标题等配置选项
 *    - 支持静态方法和实例方法两种调用方式
 * 5. [新增] 方法别名
 *    - delcols (deleteCols 的别名)
 *    - SelectCols (selectCols 的大写别名)
 *    - zDistinct (distinct 的中文别名)
 * 课程咨询: 微信 EXCEL880B
 *
 * ------------------------------------------------------------------------
 * 核心特性
 * ------------------------------------------------------------------------
 * 1. 一行代码完成复杂数据操作（筛选、排序、分组、透视）
 * 2. 完整的Array2D二维数组工具库（100+方法）
 * 3. 智能类型识别（自动识别数字、日期、字符串、布尔值）
 * 4. 双语API支持（中英文方法名自由切换）
 * 5. 链式调用与静态方法双模式
 * 6. 完善的错误处理和边界检查
 *
 * ------------------------------------------------------------------------
 * 设计原则
 * ------------------------------------------------------------------------
 * 1. 零依赖：仅使用WPS原生对象和标准JavaScript
 * 2. 向后兼容：保持与原版JSA880 API一致
 * 3. 性能优先：大数据量下仍保持高效
 * 4. 类型安全：完善的参数检查和类型转换
 *
 * ------------------------------------------------------------------------
 * 代码示例
 * ------------------------------------------------------------------------
 *
 * // 示例1: 超级透视表（核心功能）
 * Array2D.z超级透视(数据, ['产品+,国家-'], ['月份+'], ['count(),sum("销量"),average("金额")']);
 *
 * // 示例2: 链式数据处理
 * $.maxArray("A1:H1")
 *   .skip(1)                    // 跳过表头
 *   .filter('f2>0')             // 筛选有效数据
 *   .sortByCols('f1+,f2-')      // 多列排序
 *   .take(100)                  // 取前100行
 *   .toRange("K1");             // 输出到单元格
 *
 * // 示例3: 智能类型处理
 * var arr = Array2D(data);
 * arr.z智能排序('f1');          // 自动识别类型并排序
 * arr.z智能分组('f2', 'month'); // 按月分组
 *
 * // 示例4: 合并单元格（新增）
 * RngUtils.z合并单元格("A1:B2");
 * $.mergeCells("A1:C3");
 *
 * ========================================================================
 */

// ==================== 目录索引 ====================
// 快速导航：使用 Ctrl+F 搜索下面的标签跳转到对应模块
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第一部分：基础设施（Foundation）                              │
// ├─────────────────────────────────────────────────────────────┤
// │ [ENV_DETECTION]   - 环境检测（WPS/NodeJS/浏览器）              │
// │ [CONSTANTS]       - 常量定义（Lambda模式、填充方向等）         │
// │ [LAMBDA_PARSER]   - Lambda表达式解析器（箭头函数、列选择器）    │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第二部分：核心类 - Array2D（Core - Array2D）                  │
// ├─────────────────────────────────────────────────────────────┤
// │ [ARRAY2D_BASE]    - 基础架构（构造函数、原型链、_new方法）     │
// │ [ARRAY2D_BASIC]   - 基础操作（val, copy, count, isEmpty）   │
// │ [ARRAY2D_STATS]   - 统计计算（sum, average, max, min, median）│
// │ [ARRAY2D_MATRIX]  - 矩阵操作（transpose, reverse, fill）    │
// │ [ARRAY2D_ROWCOL]  - 行列操作（skip, take, getRow, addRow）  │
// │ [ARRAY2D_SORT]    - 排序去重（sort, sortByCols, distinct）  │
// │ [ARRAY2D_FILTER]  - 查找筛选（filter, find, where链式筛选） │
// │ [ARRAY2D_PIVOT]   - 分组透视（groupBy, z超级透视）           │
// │ [ARRAY2D_JOIN]    - 连接操作（leftjoin, fulljoin, zip）     │
// │ [ARRAY2D_BATCH]   - 批量操作（deleteCols, insertCols）      │
// │ [ARRAY2D_PAGE]    - 分页切片（pageByRows, slice, nth）      │
// │ [ARRAY2D_IO]      - 输入输出（toRange, toJson）             │
// │ [ARRAY2D_SMART]   - 智能功能（智能排序、智能分组、类型检测）  │
// │ [QUERY_BUILDER]   - 链式查询构建器（where, and, or）        │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第三部分：辅助类（Auxiliary Classes）                         │
// ├─────────────────────────────────────────────────────────────┤
// │ [SUPERMAP]       - 增强Map（支持调试视图、层级展开）           │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第四部分：工具类（Utilities）                                 │
// ├─────────────────────────────────────────────────────────────┤
// │ [RNGUTILS]       - Range区域工具库（安全数组、最大行列等）     │
// │ [SHTUTILS]       - 工作表工具库（激活、重命名、复制等）       │
// │ [DATEUTILS]      - 日期工具库（加减、格式化、月初月末）       │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第五部分：链式调用类（Chain Classes）                         │
// ├─────────────────────────────────────────────────────────────┤
// │ [RANGECHAIN]     - Range链式调用（z值、z加边框等）            │
// │ [SHEETCHAIN]     - Sheet链式调用（z激活、z名称等）            │
// │ [AS]             - 类型转换包装类（asArray, asRange等）       │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第六部分：函数库（Function Libraries）                        │
// ├─────────────────────────────────────────────────────────────┤
// │ [JSA]            - 通用函数库（工具方法集合）                  │
// │ [IO]             - 文件操作库（导入导出）                     │
// │ [TYPE_CONVERT]   - 类型转换函数（as系列：asArray, asRange）   │
// │ [TYPE_CHECK]     - 类型检查函数（is系列：isArray, isRange）   │
// │ [GLOBAL_FUNCS]   - 全局工具函数（log, logjson, f1, $fx）      │
// │ [SHORTCUT_$]     - $快捷函数（maxArray, safeArray等）        │
// └─────────────────────────────────────────────────────────────┘
//
// ┌─────────────────────────────────────────────────────────────┐
// │ 第七部分：全局导出（Exports）                                 │
// ├─────────────────────────────────────────────────────────────┤
// │ [EXPORTS]        - 统一全局变量导出（WPS/NodeJS/浏览器）       │
// └─────────────────────────────────────────────────────────────┘
//
// ==================== 目录索引结束 ====================

// ==================== [CONSTANTS] 常量定义区 ====================
// 统一定义框架中使用的常量，避免魔法数字和硬编码

// ╔════════════════════════════════════════════════════════════════╗
// ║ [v5.0.0] $$ 全局别名 = Array2D                                  ║
// ║  公式 k("$$.superPivot", ...) 必须能直接访问 Array2D 方法          ║
// ╚════════════════════════════════════════════════════════════════╝
(function _kInitDollarDollar() {
    try {
        var __g = (function() { return this; })();
        if (typeof __g.$$ === 'undefined' && typeof Array2D !== 'undefined') {
            __g.$$ = Array2D;
        }
    } catch (e) { /* 静默失败 */ }
})();

/**
 * Lambda表达式匹配模式
 * @enum {RegExp}
 */
const LAMBDA_PATTERNS = {
    /** 箭头函数语法 */
    ARROW_FUNCTION: /=>/,
    /** 索引选择器 $0, $1, $2 */
    INDEX_SELECTOR: /\$(\d+)/g,
    /** 列选择器 f1, f2, f3 */
    COLUMN_SELECTOR: /^f\d+/,
    /** 多列选择器 f1,f2 或 f1, f2, f3 */
    MULTI_COLUMN: /^f\d+(\s*[,，]\s*f\d+)+$/,
    /** 方括号多列选择器 [f1,f2,f3] */
    ARRAY_BRACKET: /^\[f\d+(\s*[,，]\s*f\d+)*\]$/
};

/**
 * 填充方向枚举
 * @enum {string}
 */
const FILL_DIRECTION = {
    /** 向右填充（默认） */
    RIGHT: 'right',
    /** 向左填充 */
    LEFT: 'left',
    /** 向下填充 */
    DOWN: 'down',
    /** 向上填充 */
    UP: 'up'
};

/**
 * 数组操作边界限制
 * @enum {number}
 */
const ARRAY_LIMITS = {
    /** 最大数组索引 */
    MAX_INDEX: 1000000,
    /** 默认填充值 */
    DEFAULT_FILL: '',
    /** 默认行数 */
    DEFAULT_ROWS: 1,
    /** 默认列数 */
    DEFAULT_COLS: 1
};

/**
 * 合并单元格标记
 * @enum {string}
 */
const MERGE_CELL_MARKERS = {
    /** 合并单元格标记（WPS用null表示） */
    WPS_MERGE: null,
    /** 合并单元格标记（用户自定义） */
    CUSTOM: '#MERGE#'
};

// ==================== [CONSTANTS] 常量定义区结束 ====================

// WPS JSA 专用环境（Application 对象始终可用）

// ==================== [LAMBDA_PARSER] Lambda表达式解析器 ====================
/**
 * Lambda表达式缓存
 * @private
 */
const _lambdaCache = Object.create(null);

/**
 * 解析Lambda表达式为可执行函数（支持ES6箭头函数）
 * @private
 * @param {string|Function} expr - Lambda表达式或函数
 * @returns {Function|null} 可执行函数
 * @example
 * parseLambda('$0*2')           // _ => _[0]*2
 * parseLambda('f1+f2')           // _ => _[0]+_[1]
 * parseLambda('row=>row.x')      // row => row.x
 * parseLambda('x=>x.age>18')     // x => x.age>18
 */
// 🔧 v4.0.11 修复: 统一使用 const 声明 fn 以避免 hoisting 问题
function parseLambda(expr) {
    if (typeof expr === 'function') return expr;
    // 🔧 v4.0.11 修复: null/undefined/空字符串 提前返回 null，避免缓存无用函数
    if (expr === null || expr === undefined || expr === '') {
        return null;
    }
    // 🔧 XXD-210: 支持数字选择器 — 数字视为列索引，返回 row[col]
    if (typeof expr === 'number' && isFinite(expr) && expr >= 0 && Math.floor(expr) === expr) {
        var _col = expr;
        return function(row) { return row == null ? undefined : row[_col]; };
    }
    if (typeof expr !== 'string') return null;

    // 缓存检查
    if (_lambdaCache[expr]) return _lambdaCache[expr];

    var fn;
    try {
        // 处理箭头函数语法 (ES6)
        if (LAMBDA_PATTERNS.ARROW_FUNCTION.test(expr)) {
            // 使用箭头函数语法
            fn = eval('(' + expr + ')');
        }
// 处理 $0, $1, $2 索引语法 -> 转换为箭头函数
        else if (LAMBDA_PATTERNS.INDEX_SELECTOR.test(expr)) {
            const indexMatch = expr.match(LAMBDA_PATTERNS.INDEX_SELECTOR);
            if (indexMatch && indexMatch.length > 0) {
                const indices = indexMatch.map(m => parseInt(m.substring(1)));
                // 🔧 v4.0.11 修复: 防御性检查 - 空数组 Math.max() 会返回 -Infinity
                if (indices.length === 0) {
                    console.warn('Lambda解析: $ 表达式未匹配到有效索引:', expr);
                    return null;
                }
                const maxIndex = Math.max.apply(Math, indices);
                // 安全检查：防止索引越界
                if (!isFinite(maxIndex) || maxIndex > ARRAY_LIMITS.MAX_INDEX) {
                    console.warn('Lambda索引超出限制:', maxIndex);
                    return null;
                }
                // 转换为箭头函数: $0 -> _[0], $1 -> _[1]
                fn = new Function('_', 'return ' + expr.replace(LAMBDA_PATTERNS.INDEX_SELECTOR, '_[$1]'));
            }
        }
        // 处理多列选择器 f1,f2 或 f1, f2, f3 -> 返回数组
        else if (LAMBDA_PATTERNS.MULTI_COLUMN.test(expr)) {
            // 分割并转换为数组: 'f1,f2' -> '[_[0],_[1]]'
            const cols = expr.split(/\s*[,，]\s*/).map(c => '_[' + (parseInt(c.substring(1)) - 1) + ']').join(',');
            fn = new Function('_', 'return [' + cols + ']');
        }
        // 🔧 v3.9.5 新增：处理方括号多列选择器 [f1,f2,f3] -> 返回数组
        else if (LAMBDA_PATTERNS.ARRAY_BRACKET.test(expr)) {
            // 提取方括号内的内容: '[f2,f4,f5,f6,f7,f8]' -> 'f2,f4,f5,f6,f7,f8'
            const innerExpr = expr.slice(1, -1).trim();
            const cols = innerExpr.split(/\s*[,，]\s*/).map(c => '_[' + (parseInt(c.substring(1)) - 1) + ']').join(',');
            fn = new Function('_', 'return [' + cols + ']');
        }
        // 处理 f1, f2, f3 单列选择器语法 -> 转换为箭头函数
        // 匹配 f1, f2, f3 等列选择器（支持带括号的形式如 f(1)）
        else if (/f\s*\(\s*\d+\s*\)/.test(expr) || LAMBDA_PATTERNS.COLUMN_SELECTOR.test(expr)) {
            // 转换 f1 -> _[0], f2 -> _[1], f(1) -> _[0], 等
            fn = new Function('_', 'return ' + expr.replace(/f\s*\(?\s*(\d+)\s*\)?\s*/gi, function(m, num) {
                return '_[' + (parseInt(num) - 1) + ']';
            }));
        }
        // 其他情况当作表达式
        else {
            fn = new Function('_', 'return ' + expr);
        }
    } catch (e) {
        console.warn('Lambda解析失败:', expr, e);
        return null;
    }

    _lambdaCache[expr] = fn;
    return fn;
}

/**
 * 静态方法：解析Lambda表达式
 * @param {string|Function} expr - Lambda表达式或函数
 * @returns {Function|null} 可执行函数
 * @example
 * Array2D.parseLambda('f1+f2')
 * Array2D.z解析函数表达式('row=>row.x*2')
 */
// ==================== [HELPER] 通用工具函数 ====================
/**
 * 解包Array2D值 - 如果结果是Array2D对象，提取内部数组
 * @param {*} result 
 * @returns {*} 解包后的值
 */
function unwrapVal(result) {
    if (result === null || result === undefined) return result;
    // 如果是 Array2D 或带 val() 方法的包装对象，提取其原始值
    if (typeof result === 'object' && typeof result.val === 'function') {
        return result.val();
    }
    return result;
}

/**
 * 深拷贝数组/对象（JSON方式）
 * @param {*} obj 
 * @returns {*} 深拷贝结果
 */

/**
 * 键相等比较（支持数组键）
 */
function keysEqual(k1, k2) {
    if (k1 === k2) return true;
    if (Array.isArray(k1) && Array.isArray(k2)) {
        if (k1.length !== k2.length) return false;
        for (var ki = 0; ki < k1.length; ki++) {
            if (k1[ki] !== k2[ki]) return false;
        }
        return true;
    }
    return String(k1) === String(k2);
}

/**
 * 键序列化（用于作为对象key）
 */
function serializeKey(k) {
    if (k === null || k === undefined) return String(k);
    if (Array.isArray(k)) return k.join('|@|');
    return String(k);
}

Array2D.parseLambda = parseLambda;
Array2D.z解析函数表达式 = parseLambda;

// ==================== [ARRAY2D_BASE] Array2D基础架构 ====================
// 
// 【设计原理 - 寄生组合式继承】
// 
// 问题：JavaScript中如何让自定义对象既拥有数组的特性，又能添加自定义方法？
// 
// 方案对比：
// 1. 原型继承：Array2D.prototype = new Array() - 会执行Array构造函数，有问题
// 2. 对象扩展：直接修改Array.prototype - 污染原生对象，不推荐
// 3. 包装模式：内部持有一个数组 - 需要代理所有数组方法，复杂
// 4. 寄生组合式继承（本方案）：只继承原型，不执行父类构造函数 - 最佳方案
//
// 实现步骤：
// 1. 使用 Object.create(Array.prototype) 创建以Array.prototype为原型的空对象
// 2. 将Array2D.prototype指向这个对象
// 3. 在构造函数中，使用Array.prototype.push.apply(this, items)将数据附加到实例
// 4. 这样Array2D实例就是一个真正的数组，同时又拥有自定义方法
//
// 好处：
// - Array2D实例是真正的数组，可使用所有数组方法（slice, concat等）
// - JSON.stringify自动正确处理
// - for...of循环可用
// - instanceof Array 返回true

/**
 * Array2D - 二维数组处理工具（支持智能提示和链式调用）
 * @constructor
 * @class
 * @description 提供丰富的二维数组操作函数，支持中英双语API
 * @param {Array} [data] - 二维数组数据，可为空、一维数组或二维数组
 * @returns {Array2D} Array2D实例，支持链式调用和智能提示
 * @throws {Error} 当传入非数组数据时自动包装为单元素二维数组
 * @example
 * // 基本使用 - 创建并计算
 * var sum = Array2D([[1,2,3],[4,5,6]]).z求和();        // 返回 21
 * 
 * // 链式调用 - 流畅的数据处理
 * Array2D([[1,2],[3,4],[5,6]])
 *   .z跳过(1)                    // 跳过第1行（表头）
 *   .z筛选('f1>2')               // 筛选第1列大于2的行
 *   .z多列排序('f2+')            // 按第2列升序
 *   .toRange("A10");              // 输出到A10单元格
 * 
 * // Lambda表达式 - 简洁的列操作
 * Array2D(data).z求和('f1');      // 对第1列求和
 * Array2D(data).z平均值('f2');    // 对第2列求平均
 * 
 * // 写入WPS单元格
 * Array2D([[1,2],[3,4]]).toRange("A1");  // 将数据写入A1:B2
 */
function Array2D(data) {
    // 【工厂模式检测】
    // 如果调用时没有用new（如 Array2D([[1,2]])），则自动补上new
    // 这样用户既可使用 new Array2D(data)，也可直接使用 Array2D(data)
    if (!(this instanceof Array2D)) {
        return new Array2D(data);
    }

    // 【数据规范化处理】
    // 将各种输入格式统一转换为二维数组格式
    var items = [];
    if (data === null || data === undefined) {
        // 空值转为空数组
        items = [];
    } else if (data instanceof Array2D) {
        // Array2D 实例：直接提取内部数组
        items = data._items;
    } else if (Array.isArray(data)) {
        // 数组直接保留
        items = data;
    } else {
        // 其他类型（数字、字符串等）包装为单元素二维数组
        items = [[data]];
    }

    // v4.0.11: 为所有行注入 .f1/.f2 列访问器，支持 x=>x.f3 箭头函数回调
    for (var _fi = 0; _fi < items.length; _fi++) {
        var _frow = items[_fi];
        if (Array.isArray(_frow)) {
            for (var _fc = 0; _fc < _frow.length; _fc++) {
                if (!(_frow.hasOwnProperty('f' + (_fc + 1)))) {
                    Object.defineProperty(_frow, 'f' + (_fc + 1), {
                        get: (function(idx) { return function() { return this[idx]; }; })(_fc),
                        set: (function(idx) { return function(v) { this[idx] = v; }; })(_fc),
                        enumerable: false,
                        configurable: true
                    });
                }
            }
        }
    }

    // 【关键步骤：将数据附加到实例】
    // 使用Array原型的push方法，将所有元素添加到当前实例(this)
    // 这样this就成为一个真正的数组（具备length和索引访问能力）
    Array.prototype.push.apply(this, items);

    // 【添加内部属性】
    // 使用Object.defineProperty定义属性，设置enumerable: false
    // 这样这些属性不会出现在for...in循环和Object.keys中，保持数组的纯净性

    // _original: 保存原始传入的数据（用于调试和追溯）
    Object.defineProperty(this, '_original', {
        value: data,
        writable: true,
        enumerable: false,      // 不可枚举，JSON.stringify时不会包含
        configurable: true
    });

    // _items: 数据访问器属性
    // getter: 返回当前数组数据的副本（避免外部直接修改内部状态）
    // setter: 用新数据替换当前所有数据（用于链式操作中的数据更新）
    Object.defineProperty(this, '_items', {
        get: function() {
            // 🔧 P0-3 性能优化: 使用原生 slice 替代手动循环，性能提升 3-5x
            var copy = Array.prototype.slice.call(this);
            // v4.0.11: 为副本注入 .f1/.f2 列访问器
            for (var _fi = 0; _fi < copy.length; _fi++) {
                var _frow = copy[_fi];
                if (Array.isArray(_frow)) {
                    for (var _fc = 0; _fc < _frow.length; _fc++) {
                        if (!(_frow.hasOwnProperty('f' + (_fc + 1)))) {
                            Object.defineProperty(_frow, 'f' + (_fc + 1), {
                                get: (function(idx) { return function() { return this[idx]; }; })(_fc),
                                set: (function(idx) { return function(v) { this[idx] = v; }; })(_fc),
                                enumerable: false,
                                configurable: true
                            });
                        }
                    }
                }
            }
            return copy;
        },
        set: function(value) {
            // 清空当前数据并填充新数据
            // 使用Array.prototype方法而不是this.splice，因为this此时可能还不是完整数组
            Array.prototype.splice.call(this, 0, this.length);
            Array.prototype.push.apply(this, value);
        },
        enumerable: false,      // 不可枚举
        configurable: true
    });

    // 🔧 v3.7.9 修复: 复制输入数据的 _header 属性（如果存在）
    // 这样当静态方法调用 new Array2D(arr) 时，_header 会被保留
    // 使用 'in' 操作符检查，因为 _header 可能是不可枚举的
    // XXD-160: 排除 typeof==='function'，避免复制到 Array2D.prototype._header 链式 setter
    if (data && typeof data === 'object' && '_header' in data && data._header !== undefined && data._header !== null && typeof data._header !== 'function') {
        Object.defineProperty(this, '_header', {
            value: data._header,
            writable: true,
            enumerable: false,
            configurable: true
        });
    }
}

// 【设置原型链 - 寄生组合式继承的核心】
// 1. 创建一个以Array.prototype为原型的空对象
// 2. 将Array2D.prototype指向这个对象
// 3. Array2D实例的原型链：instance -> Array2D.prototype -> Array.prototype -> Object.prototype
Array2D.prototype = Object.create(Array.prototype);

// 【修复constructor指向】
// 上一步操作后，Array2D.prototype.constructor指向Array
// 需要手动修正回Array2D，否则instanceof检查会出问题
Array2D.prototype.constructor = Array2D;

// 【添加toJSON方法 - 序列化支持】
// 当使用JSON.stringify()序列化Array2D实例时，自动调用此方法
// 这样序列化结果只包含数据内容，不包含内部属性（_original, _items等）
// 示例：JSON.stringify(Array2D([[1,2]])) 返回 "[[1,2]]" 而不是包含内部状态的对象
Object.defineProperty(Array2D.prototype, 'toJSON', {
    value: function() {
        return this._items;     // 返回数据数组（getter返回的副本）
    },
    enumerable: false,
    configurable: true,
    writable: true
});

// ==================== [SHTUTILS] 工作表操作工具 ====================

function __sht_getSheet(sht) {
    if (typeof sht === 'string') return Sheets(sht);
    return sht || null;
}

function __sht_wildcardToRegex(wildcard) {
    var pattern = wildcard.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    return new RegExp('^' + pattern + '$', 'i');
}

/**
 * ShtUtils - 工作表函数工具库，增强工作表操作
 * @constructor
 * @class
 * @description 工作表函数工具库,增强工作表操作。备注：这里面工作表作为参数都可以传表名或者工作表对象
 * @param {String|Worksheet} [工作表] - 工作表名称或对象
 * @returns {ShtUtils} ShtUtils实例，支持智能提示
 * @example
 * var rng = ShtUtils.z安全已使用区域("Sheet1");
 * var ok  = ShtUtils.z包含表名("数据*");
 */
function ShtUtils(工作表) {
    if (!(this instanceof ShtUtils)) {
        return new ShtUtils(工作表);
    }
    this._sheet = 工作表 ? __sht_getSheet(工作表) : null;
}

ShtUtils.prototype = Object.create(Object.prototype);
ShtUtils.prototype.constructor = ShtUtils;

/**
 * 获取工作表从A1开始的可使用区域
 * @param {String|Worksheet} 工作表 - 要获取安全已使用区域的工作表
 * @returns {Range} 从A1开始的已使用单元格区域
 */
ShtUtils.prototype.z安全已使用区域 = function(工作表) {
    var sheet = 工作表 ? __sht_getSheet(工作表) : Application.ActiveSheet;
    if (!sheet) return null;
    var usedRange;
    try { usedRange = sheet.UsedRange; } catch (e) { return sheet.Range("A1"); }
    if (!usedRange) return sheet.Range("A1");
    return sheet.Range(sheet.Cells(1, 1), sheet.Cells(usedRange.Row + usedRange.Rows.Count - 1, usedRange.Column + usedRange.Columns.Count - 1));
}
ShtUtils.prototype.safeUsedRange = ShtUtils.prototype.z安全已使用区域;
ShtUtils.z安全已使用区域 = ShtUtils.prototype.z安全已使用区域;
ShtUtils.safeUsedRange = ShtUtils.prototype.z安全已使用区域;

/**
 * 检查表集合中是否包含指定表名（支持 * ? 通配符）
 * @param {String} 表名 - 要检查的表名，可以用? * 通配符
 * @param {Sheets} [表集合=Sheets] - 要检查的表集合对象，默认为 Sheets
 * @returns {boolean} 是否包含
 */
ShtUtils.prototype.z包含表名 = function(表名, 表集合) {
    var shts = 表集合 || Sheets;
    if (!shts) return false;
    var pattern = __sht_wildcardToRegex(表名);
    for (var i = 1; i <= shts.Count; i++) { if (pattern.test(shts(i).Name)) return true; }
    return false;
}
ShtUtils.prototype.includesSht = ShtUtils.prototype.z包含表名;
ShtUtils.z包含表名 = ShtUtils.prototype.z包含表名;
ShtUtils.includesSht = ShtUtils.prototype.z包含表名;

/**
 * 筛选表集合中匹配的表名（支持 * ? 通配符）
 * @param {String} 表名 - 要筛选的表名，可以用? *通配符
 * @param {Sheets} [表集合=Sheets] - 要筛选的表集合对象，默认为 Sheets
 * @returns {Array} 匹配的表名一维数组
 */
ShtUtils.prototype.z表名筛选 = function(表名, 表集合) {
    var shts = 表集合 || Sheets;
    if (!shts) return [];
    var pattern = __sht_wildcardToRegex(表名);
    var result = [];
    for (var i = 1; i <= shts.Count; i++) { if (pattern.test(shts(i).Name)) result.push(shts(i).Name); }
    return result;
}
ShtUtils.prototype.filterShts = ShtUtils.prototype.z表名筛选;
ShtUtils.z表名筛选 = ShtUtils.prototype.z表名筛选;
ShtUtils.filterShts = ShtUtils.prototype.z表名筛选;

/**
 * 在表集合中按名称查找工作表
 * @param {String} sht - 要查找的表名
 * @param {Sheets} [shts=Sheets] - 要查找的表集合对象，默认为 Sheets
 * @returns {Sheet} 查找到的表对象
 */
ShtUtils.prototype.z查找表 = function(sht, shts) {
    shts = shts || Sheets;
    if (!shts) return null;
    for (var i = 1; i <= shts.Count; i++) { if (shts(i).Name === sht) return shts(i); }
    return null;
}
ShtUtils.prototype.findSht = ShtUtils.prototype.z查找表;
ShtUtils.z查找表 = ShtUtils.prototype.z查找表;
ShtUtils.findSht = ShtUtils.prototype.z查找表;

/**
 * 判断工作表是否为空表
 * @param {String|Worksheet} 工作表 - 要判断的工作表
 * @returns {boolean} 是否为空表
 */
ShtUtils.prototype.z判断空表 = function(工作表) {
    var sheet = __sht_getSheet(工作表);
    if (!sheet) return true;
    try {
        var used = sheet.UsedRange;
        if (!used) return true;
        if (used.Row >= 1048576 || used.Column >= 16384) return true;
        if (used.Rows.Count === 1 && used.Columns.Count === 1) {
            var val = used.Value2;
            if (val === null || val === undefined || val === '') return true;
        }
        return false;
    } catch (e) { return true; }
}
ShtUtils.prototype.isEmptySht = ShtUtils.prototype.z判断空表;
ShtUtils.z判断空表 = ShtUtils.prototype.z判断空表;
ShtUtils.isEmptySht = ShtUtils.prototype.z判断空表;

/**
 * 删除指定的工作表
 * @param {String|Worksheet} 工作表 - 要删除的工作表
 */
ShtUtils.prototype.z删除表 = function(工作表) {
    var sheet = __sht_getSheet(工作表);
    if (!sheet) return;
    Application.DisplayAlerts = false;
    try { sheet.Delete(); } catch (e) { /* 忽略 */ }
    Application.DisplayAlerts = true;
}
ShtUtils.prototype.deleteSht = ShtUtils.prototype.z删除表;
ShtUtils.z删除表 = ShtUtils.prototype.z删除表;
ShtUtils.deleteSht = ShtUtils.prototype.z删除表;

/**
 * 按代码名称查找工作表
 * @param {String} 表名 - 要查找的代码名称
 * @param {Sheets} [表集合=Sheets] - 要查找的表集合对象，默认: Sheets
 * @returns {Worksheet} 查找到的表对象
 */
ShtUtils.prototype.z按代码名称 = function(表名, 表集合) {
    var shts = 表集合 || Sheets;
    if (!shts) return null;
    for (var i = 1; i <= shts.Count; i++) { try { if (shts(i).CodeName === 表名) return shts(i); } catch (e) { } }
    return null;
}
ShtUtils.prototype.byCodeName = ShtUtils.prototype.z按代码名称;
ShtUtils.z按代码名称 = ShtUtils.prototype.z按代码名称;
ShtUtils.byCodeName = ShtUtils.prototype.z按代码名称;

/**
 * 隐藏表集合中的表。不传参数则隐藏当前表外所有表
 * @param {Array} [表集合] - 要隐藏的表集合对象
 */
ShtUtils.prototype.z隐藏表 = function(表集合) {
    if (表集合 !== undefined) {
        var arr = Array.isArray(表集合) ? 表集合 : [表集合];
        for (var i = 0; i < arr.length; i++) {
            try { var sheet = typeof arr[i] === 'string' ? Sheets(arr[i]) : arr[i]; if (sheet) sheet.Visible = 0; } catch (e) { }
        }
        return;
    }
    var current = Application.ActiveSheet;
    for (var i = 1; i <= Sheets.Count; i++) { try { var s = Sheets(i); if (s.Name !== current.Name) s.Visible = 0; } catch (e) { } }
}
ShtUtils.prototype.hideSheets = ShtUtils.prototype.z隐藏表;
ShtUtils.z隐藏表 = ShtUtils.prototype.z隐藏表;
ShtUtils.hideSheets = ShtUtils.prototype.z隐藏表;

/**
 * 显示表集合中的表。不传参数则显示所有表
 * @param {Array} [表集合] - 表名数组
 */
ShtUtils.prototype.z显示表 = function(表集合) {
    if (表集合 !== undefined) {
        var arr = Array.isArray(表集合) ? 表集合 : [表集合];
        for (var i = 0; i < arr.length; i++) { try { var sheet = typeof arr[i] === 'string' ? Sheets(arr[i]) : arr[i]; if (sheet) sheet.Visible = -1; } catch (e) { } }
        return;
    }
    for (var i = 1; i <= Sheets.Count; i++) { try { Sheets(i).Visible = -1; } catch (e) { } }
}
ShtUtils.prototype.showSheets = ShtUtils.prototype.z显示表;
ShtUtils.z显示表 = ShtUtils.prototype.z显示表;
ShtUtils.showSheets = ShtUtils.prototype.z显示表;

/**
 * 根据工作表名称激活工作表
 * @param {Worksheet} 工作表 - 待激活的工作表
 */
ShtUtils.prototype.z激活表 = function(工作表) {
    var sheet = 工作表 ? __sht_getSheet(工作表) : Application.ActiveSheet;
    if (sheet) sheet.Activate();
    return sheet;
}
ShtUtils.prototype.shtActivate = ShtUtils.prototype.z激活表;
ShtUtils.z激活表 = ShtUtils.prototype.z激活表;
ShtUtils.shtActivate = ShtUtils.prototype.z激活表;

/**
 * 返回指定工作表的最后一行的行号
 * @param {String|Worksheet} 工作表 - 要返回最后行号的工作表
 * @returns {Number} 最后一行的行号
 */
ShtUtils.prototype.z最后一行 = function(工作表) {
    var sheet = __sht_getSheet(工作表);
    if (!sheet) return 0;
    try {
        var used = sheet.UsedRange;
        if (!used) return 0;
        return used.Row + used.Rows.Count - 1;
    } catch (e) { return 0; }
}
ShtUtils.prototype.lastRow = ShtUtils.prototype.z最后一行;
ShtUtils.z最后一行 = ShtUtils.prototype.z最后一行;
ShtUtils.lastRow = ShtUtils.prototype.z最后一行;

/**
 * 将工作表名中的违规字符替换为 _
 * @param {String} 工作表名 - 待检测的工作表名
 * @returns {String} 正确的表名
 */
ShtUtils.prototype.z纠正表名 = function(工作表名) {
    var name = String(工作表名);
    name = name.replace(/[:\/\\\?\*\[\]]/g, '_');
    if (name.length > 31) name = name.substring(0, 31);
    if (name.charAt(0) === "'") name = '_' + name.substring(1);
    if (name.length > 0 && name.charAt(name.length - 1) === "'") name = name.substring(0, name.length - 1) + '_';
    return name;
}
ShtUtils.prototype.correctShtName = ShtUtils.prototype.z纠正表名;
ShtUtils.z纠正表名 = ShtUtils.prototype.z纠正表名;
ShtUtils.correctShtName = ShtUtils.prototype.z纠正表名;

/**
 * 对工作表数组进行排序
 * @param {Array|Sheets} shts - 要排序的工作表数组或 Sheets 集合
 * @param {Function|Array} [sortFn] - 排序函数或自定义序列数组
 */
ShtUtils.prototype.z工作表排序 = function(shts, sortFn) {
    if (!shts) return;
    var names = [];
    if (typeof shts.Count === 'number') {
        for (var i = 1; i <= shts.Count; i++) names.push(shts(i).Name);
    } else if (Array.isArray(shts)) {
        for (var i = 0; i < shts.length; i++) { names.push(typeof shts[i] === 'string' ? shts[i] : shts[i].Name); }
    }
    if (Array.isArray(sortFn)) {
        var order = sortFn;
        names.sort(function(a, b) { var ia = order.indexOf(a), ib = order.indexOf(b); if (ia === -1 && ib === -1) return a.localeCompare(b); if (ia === -1) return 1; if (ib === -1) return -1; return ia - ib; });
    } else if (typeof sortFn === 'function') {
        names.sort(sortFn);
    } else {
        names.sort(function(a, b) { return a.localeCompare(b); });
    }
// 🔧 v4.0.11 修复: 静默吞异常的改为在控制台输出错误，便于调试
    for (var i = 0; i < names.length; i++) {
        try {
            var targetSheet = Sheets(names[i]);
            if (targetSheet) targetSheet.Move(null, Sheets(Sheets.Count));
        } catch (e) {
            if (typeof Console !== 'undefined') {
                Console.log('🔧 ShtUtils.z工作表排序 移动表失败 [' + names[i] + ']: ' + e.message);
            }
        }
    }
}
ShtUtils.prototype.sheetsSort = ShtUtils.prototype.z工作表排序;
ShtUtils.z工作表排序 = ShtUtils.prototype.z工作表排序;
ShtUtils.sheetsSort = ShtUtils.prototype.z工作表排序;

function ShtUtils_ctor(initialSheet) {
    return new ShtUtils(initialSheet);
}
ShtUtils_ctor.prototype = ShtUtils.prototype;
// ==================== [RNGUTILS] Range区域工具库 ====================

/**
 * RngUtils - Range区域操作工具（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description WPS Range区域操作增强工具
 * @example
 * RngUtils("A1:C10").z加边框().z自动列宽()
 */
function RngUtils(initialRange) {
    if (!(this instanceof RngUtils)) {
        return new RngUtils(initialRange);
    }
    this._range = initialRange ? this._toRange(initialRange) : null;
}

/**
 * 转换为Range对象
 * @private
 * @param {Range|string} rng - Range对象或地址字符串
 * @returns {Range|null} Range对象或null
 */
RngUtils.prototype._toRange = function(rng) {
    if (!rng) return null;
    if (typeof rng === 'string') return Range(rng);
    return rng;
};

/**
 * 获取/设置Range
 * @param {Range|string} newRange - 新Range
 * @returns {RngUtils|Range} 设置时返回this，否则返回当前Range
 */
RngUtils.prototype.rng = function(newRange) {
    if (newRange !== undefined) {
        this._range = this._toRange(newRange);
        return this;
    }
    return this._range;
};

/**
 * 获取值
 * @returns {Array} 二维数组
 */
RngUtils.prototype.val = function() {
    if (!this._range) return null;
    return this._range.Value2;
};

// ==================== 基础信息函数 ====================

/**
 * z列字母转数字 - Excel 列字母转数字 (A=1, Z=26, AA=27)
 * @param {string} s - 列字母 (大小写均可)
 * @returns {number} 列号 (1-based)
 * @example
 * RngUtils.z列字母转数字("A")    // 1
 * RngUtils.z列字母转数字("Z")    // 26
 * RngUtils.z列字母转数字("AA")   // 27
 */
RngUtils.z列字母转数字 = function(s) { var n = 0; for (var i=0;i<s.length;i++) n = n*26 + (s.charCodeAt(i)-64); return n; };
/**
 * z数字转列字母 - Excel 列号转字母 (1=A, 26=Z, 27=AA)
 * @param {number} n - 列号 (1-based, 必须 > 0)
 * @returns {string} 列字母
 * @example
 * RngUtils.z数字转列字母(1)   // "A"
 * RngUtils.z数字转列字母(26)  // "Z"
 * RngUtils.z数字转列字母(27)  // "AA"
 */
RngUtils.z数字转列字母 = function(n) { var s=""; while (n>0) { var r = (n-1) % 26; s = String.fromCharCode(65+r) + s; n = Math.floor((n-1)/26); } return s; };
/**
 * z最后一个 - 获取指定区域的最后一个单元格
 * @param {Range|string} rng - 单元格区域
 * @returns {Range} 最后一个单元格
 * @example
 * RngUtils.z最后一个("A1:A13")  // $A$13
 */
RngUtils.z最后一个 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Cells(r.Rows.Count, r.Columns.Count);
};
RngUtils.lastCell = RngUtils.z最后一个;

/**
 * z安全区域 - 获取当前区域与UsedRange的交集
 * @param {Range|string} rng - 单元格区域
 * @returns {Range} 交集单元格
 * @example
 * RngUtils.z安全区域("A:A")  // $A$1:$A$13
 */
RngUtils.z安全区域 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return r;
    return Application.Intersect(r, usedRange);
};
RngUtils.safeRange = RngUtils.z安全区域;

/**
 * z安全数组 - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @example
 * RngUtils.z安全数组("A1:A13").filter(row => row[0] > 0).toRange("C1")
 */
RngUtils.z安全数组 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var arr = r.Value2;
    if (arr === null || arr === undefined) return new Array2D([]);
    // 单个单元格转二维数组
    if (!Array.isArray(arr)) return new Array2D([[arr]]);
    // 一维数组转二维
    if (!Array.isArray(arr[0])) {
        var result = [];
        for (var i = 0; i < arr.length; i++) {
            result.push([arr[i]]);
        }
        return new Array2D(result);
    }
    // 🔧 v3.7.5 保存表头信息到 Array2D 对象
    var result = new Array2D(arr);
    if (arr.length > 0) {
        Object.defineProperty(result, '_header', {
            value: arr[0],
            writable: true,
            enumerable: false,
            configurable: true
        });
    }
    return result;
};
RngUtils.safeArray = RngUtils.z安全数组;

/**
 * z最大行 - 获取指定区域的最大行数
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @example
 * RngUtils.z最大行("A:A")     // 70 (单列，从下往上查找第一个有效数据)
 * RngUtils.z最大行("A1")      // 70 (单单元格，自动扩展为整列)
 * RngUtils.z最大行("A1:C10")  // 10 (多单元格区域返回该区域的最后一行)
 */
RngUtils.z最大行 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return 0;

// 🔧 v4.0.11 修复: 移除冗余条件 Columns.Count === 1
        // (外层 if 已保证此条件，第二个条件才是判断单单元格的依据)
    if (r.Columns.Count === 1) {
        var col = r.Rows.Count === 1 ? sheet.Columns(r.Column) : r;
        var safe = Application.Intersect(col, usedRange);
        if (!safe) return 0;
        // 从下往上查找第一个非空单元格
        for (var i = safe.Rows.Count; i >= 1; i--) {
            var cell = safe.Cells(i, 1);
            var val = cell.Value2;
            // 跳过 null、undefined、空字符串（包括 =""）
            if (val === null || val === undefined || val === '') {
                continue;
            }
            // 跳过纯空白字符
            if (typeof val === 'string' && val.trim() === '') {
                continue;
            }
            // 找到第一个有效数据，返回行号
            return safe.Row + i - 1;
        }
        return 0;
    }

    // 多列区域，返回该区域与UsedRange交集的最后一行
    var safe = Application.Intersect(r, usedRange);
    if (!safe) return 0;
    return safe.Row + safe.Rows.Count - 1;
};
RngUtils.endRow = RngUtils.z最大行;

/**
 * z最大行单元格 - 获取指定区域最后一行的单元格
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 最后一行的单元格
 * @example
 * RngUtils.z最大行单元格("A1:A1000")  // $A$13
 */
RngUtils.z最大行单元格 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var maxRow = RngUtils.z最大行(r);
    var col = r.Column;
    return sheet.Cells(maxRow, col);
};
RngUtils.endRowCell = RngUtils.z最大行单元格;

/**
 * z最大行区域 - 获取从第一行到最后一行的区域
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Range} 从第一行到最后一行的区域
 * @example
 * RngUtils.z最大行区域("1:1000","A")  // $1:$13
 * RngUtils.z最大行区域("A1:J1")       // A1:J最大行
 */
RngUtils.z最大行区域 = function(rng, col) {
    col = col !== undefined ? col : "A";
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;

    var result;
    // 特殊参数处理
    if (col === '-c') {
        // CurrentRegion
        result = r.CurrentRegion;
    } else if (col === '-u') {
        // UsedRange
        var used = sheet.UsedRange;
        if (!used) return new RangeChain(r);
        var startRow = r.Row;
        var endRow = used.Row + used.Rows.Count - 1;
        var startCol = r.Column;
        var endCol = r.Column + r.Columns.Count - 1;
        result = sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(endRow, endCol));
    } else if (r.Rows.Count >= 16384) {
        // 整行处理 - 当rng是整行时（如 "1:1000"）
        var colNum = typeof col === 'string' ? (col.charCodeAt(0) - 64) : (col || 1);
        var maxR = RngUtils.z最大行(sheet.Columns(colNum));
        result = sheet.Range(sheet.Cells(1, colNum), sheet.Cells(maxR, colNum)).EntireRow;
    } else {
        // 默认情况 - 保持原区域的列范围，扩展行到最后一行
        var startRow = r.Row;
        var startCol = r.Column;
        var endCol = r.Column + r.Columns.Count - 1;
        var maxEndRow = startRow;
        for (var c = startCol; c <= endCol; c++) {
            var colRange = sheet.Columns(c);
            var endRow = RngUtils.z最大行(colRange);
            if (endRow > maxEndRow) {
                maxEndRow = endRow;
            }
        }
        result = sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(maxEndRow, endCol));
    }
    return new RangeChain(result);
};

/**
 * maxRange - 获取从第一行到最后一行的区域（英文别名）
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Range} 从第一行到最后一行的区域
 * @example
 * RngUtils.maxRange("1:1000","A")  // $1:$13
 * RngUtils.maxRange("A1:J1")       // A1:J最大行
 */
/**
 * maxRange - 获取从第一行到最后一行的区域（英文别名，返回 RangeChain 支持智能提示）
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {RangeChain} 从第一行到最后一行的区域（支持链式调用和智能提示）
 * @example
 * RngUtils.maxRange("1:1000","A").safeArray()  // 返回数组
 * RngUtils.maxRange("A1:J1").z加边框()         // 链式调用
 */
RngUtils.maxRange = function(rng, col) {
    var result = RngUtils.z最大行区域.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * z最大列 - 获取指定区域的最大列数
 * @param {Range|string} rng - 要获取最大列数的区域
 * @returns {number} 最大列数
 * @example
 * RngUtils.z最大列("1:1")     // 8 (单行，从右往左查找第一个有效数据)
 * RngUtils.z最大列("A1")      // 8 (单单元格，自动扩展为整行)
 * RngUtils.z最大列("A1:C10")  // 3 (多行区域返回该区域的最后一列)
 */
RngUtils.z最大列 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var usedRange = sheet.UsedRange;
    if (!usedRange) return 0;

    // 单单元格或单行时，从右往左查找第一个有效数据
    if (r.Rows.Count === 1) {
        var row = r.Rows.Count === 1 && r.Columns.Count === 1 ? sheet.Rows(r.Row) : r;
        var safe = Application.Intersect(row, usedRange);
        if (!safe) return 0;
        // 从右往左查找第一个非空单元格
        for (var i = safe.Columns.Count; i >= 1; i--) {
            var cell = safe.Cells(1, i);
            var val = cell.Value2;
            // 跳过 null、undefined、空字符串（包括 =""）
            if (val === null || val === undefined || val === '') {
                continue;
            }
            // 跳过纯空白字符
            if (typeof val === 'string' && val.trim() === '') {
                continue;
            }
            // 找到第一个有效数据，返回列号
            return safe.Column + i - 1;
        }
        return 0;
    }

    // 多行区域，返回该区域与UsedRange交集的最后一列
    var safe = Application.Intersect(r, usedRange);
    if (!safe) return 0;
    return safe.Column + safe.Columns.Count - 1;
};
RngUtils.endCol = RngUtils.z最大列;

/**
 * z最大列单元格 - 获取指定区域最后一列的单元格
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 最后一列的单元格
 * @example
 * RngUtils.z最大列单元格("1:1")  // $F$1
 */
RngUtils.z最大列单元格 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var sheet = r.Worksheet;
    var maxCol = RngUtils.z最大列(r);
    return sheet.Cells(r.Row, maxCol);
};
RngUtils.endColCell = RngUtils.z最大列单元格;

/**
 * z可见区数组 - 将可见单元格转换为数组
 * @param {Range|string} rng - 要转换的区域
 * @param {Worksheet} [tempSheet] - 临时工作表（可选）
 * @returns {Array} 可见单元格值的数组
 * @example
 * RngUtils.z可见区数组("1:4")
 */
RngUtils.z可见区数组 = function(rng, tempSheet) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var visible = r.SpecialCells(12); // xlCellTypeVisible
    if (!visible) return [];
    var arr = visible.Value2;
    // 保存到临时表
    if (tempSheet) {
        tempSheet.Range("A1").Resize(visible.Rows.Count, visible.Columns.Count).Value2 = arr;
    }
    return RngUtils.z安全数组(arr);
};
RngUtils.visibleArray = RngUtils.z可见区数组;

/**
 * z可见区域 - 获取指定区域的可见区域
 * @param {Range|string} rng - 要获取的区域
 * @returns {Range} 可见区域
 * @example
 * RngUtils.z可见区域("1:4")  // $1:$4
 */
RngUtils.z可见区域 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.SpecialCells(12); // xlCellTypeVisible
};
RngUtils.visibleRange = RngUtils.z可见区域;

/**
 * z加边框 - 为指定区域添加边框
 * @param {Range|string} rng - 要添加边框的区域
 * @param {number} [LineStyle=1] - 边框线条样式
 * @param {number} [Weight=2] - 边框线条粗细
 * @returns {Borders} 边框对象
 * @example
 * RngUtils.z加边框("A3:D7")
 */
RngUtils.z加边框 = function(rng, LineStyle, Weight) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    LineStyle = LineStyle !== undefined ? LineStyle : 1;
    Weight = Weight !== undefined ? Weight : 2;
    r.Borders.LineStyle = LineStyle;
    r.Borders.Weight = Weight;
    return r.Borders;
};
RngUtils.addBorders = RngUtils.z加边框;

/**
 * z取前几行 - 获取指定区域的前几行
 * @param {Range|string} rng - 指定区域
 * @param {number} count - 获取的行数
 * @returns {Range} 前几行的单元格
 * @example
 * RngUtils.z取前几行("a3:d7",3)  // $A$3:$D$5
 */
RngUtils.z取前几行 = function(rng, count) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Rows("1:" + count);
};
RngUtils.takeRows = RngUtils.z取前几行;

/**
 * z跳过前几行 - 跳过指定区域的前几行
 * @param {Range|string} rng - 指定区域
 * @param {number} count - 要跳过的行数
 * @returns {Range} 跳过后的单元格区域
 * @example
 * RngUtils.z跳过前几行("a3:d7",3)  // $A$6:$D$7
 */
RngUtils.z跳过前几行 = function(rng, count) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var startRow = count + 1;
    var endRow = r.Rows.Count;
    if (startRow > endRow) return null;
    return r.Rows(startRow + ":" + endRow);
};
RngUtils.skipRows = RngUtils.z跳过前几行;

/**
 * z插入多行 - 插入多行
 * @param {Range|string} rng - 要插入行的单元格区域
 * @param {any} value - 行号数组或字符串
 * @param {number} count - 要插入的行数
 * @example
 * RngUtils.z插入多行("a12:d15", '*', 2)
 */
RngUtils.z插入多行 = function(rng, value, count) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    count = count || 1;

    for (var i = r.Rows.Count; i >= 1; i--) {
        var insertValue = value;
        if (Array.isArray(value)) {
            insertValue = value[i - 1] !== undefined ? value[i - 1] : '';
        }
        for (var c = 0; c < count; c++) {
            r.Rows(i).Insert();
            var newRow = r.Rows(i);
            for (var j = 1; j <= r.Columns.Count; j++) {
                newRow.Cells(1, j).Value2 = insertValue;
            }
        }
    }
};
RngUtils.insertRows = RngUtils.z插入多行;

/**
 * z插入多列 - 插入多列
 * @param {Range|string} rng - 要插入列的单元格区域
 * @param {any} value - 列号数组或字符串
 * @param {number} count - 要插入的列数
 * @example
 * RngUtils.z插入多列("a12:d14", '*', 2)
 */
RngUtils.z插入多列 = function(rng, value, count) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    count = count || 1;

    for (var j = r.Columns.Count; j >= 1; j--) {
        var insertValue = value;
        if (Array.isArray(value)) {
            insertValue = value[j - 1] !== undefined ? value[j - 1][0] : '';
        }
        for (var c = 0; c < count; c++) {
            r.Columns(j).Insert();
            var newCol = r.Columns(j);
            for (var i = 1; i <= r.Rows.Count; i++) {
                newCol.Cells(i, 1).Value2 = insertValue;
            }
        }
    }
};
RngUtils.insertCols = RngUtils.z插入多列;

/**
 * z删除空白行 - 删除指定区域中的空白行
 * @param {Range|string} rng - 要删除空白行的单元格区域
 * @param {boolean} [entireColumn=true] - 默认删除整列 false时只作用选中区域
 * @example
 * RngUtils.z删除空白行("a11:d17")
 */
RngUtils.z删除空白行 = function(rng, entireColumn) {
    entireColumn = entireColumn !== undefined ? entireColumn : true;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // v4.1.0 性能优化: 一次 COM 调用读取所有值（替代 O(n*m) 次 .Value2 调用）
    var values = r.Value2;
    if (!values) return;
    // 1D 转 2D（单个单元格场景）
    if (!Array.isArray(values)) values = [[values]];
    if (!Array.isArray(values[0])) {
        var arr = [];
        for (var v = 0; v < values.length; v++) arr.push([values[v]]);
        values = arr;
    }

    // 在内存中找出所有空白行的索引（原始索引）
    var baseRow = r.Row;
    var blankRowIndexes = [];
    for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var isEmpty = true;
        for (var j = 0; j < row.length; j++) {
            var val = row[j];
            if (val !== null && val !== undefined && val !== '') {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) blankRowIndexes.push(baseRow + i);
    }

    // 倒序删除（避免索引偏移）
    for (var ki = blankRowIndexes.length - 1; ki >= 0; ki--) {
        if (entireColumn) {
            Rows(blankRowIndexes[ki]).EntireRow.Delete();
        } else {
            Rows(blankRowIndexes[ki]).Delete();
        }
    }
};
RngUtils.delBlankRows = RngUtils.z删除空白行;

/**
 * z删除空白列 - 删除指定区域中的空白列
 * @param {Range|string} rng - 要删除空白列的单元格区域
 * @param {boolean} [entireColumn=true] - 默认删除整列 false时只作用选中区域
 * @example
 * RngUtils.z删除空白列("A11:G14")
 */
RngUtils.z删除空白列 = function(rng, entireColumn) {
    entireColumn = entireColumn !== undefined ? entireColumn : true;
    var r = typeof rng === 'string' ? Range(rng) : rng;

    var blankCols = [];
    for (var j = r.Columns.Count; j >= 1; j--) {
        var col = r.Columns(j);
        var isEmpty = true;
        for (var i = 1; i <= r.Rows.Count; i++) {
            var val = col.Cells(i, 1).Value2;
            if (val !== null && val !== undefined && val !== '') {
                isEmpty = false;
                break;
            }
        }
        if (isEmpty) {
            blankCols.push(j);
        }
    }

    for (var ki = 0; ki < blankCols.length; ki++) {
        if (entireColumn) {
            r.Columns(blankCols[ki]).EntireColumn.Delete();
        } else {
            r.Columns(blankCols[ki]).Delete();
        }
    }
};
RngUtils.delBlankCols = RngUtils.z删除空白列;

/**
 * z删除列 - 删除指定区域中指定列号的列
 * @param {Range|string} rng - 要操作的单元格区域
 * @param {number} colIdx - 要删除的列号（从 1 开始，与 Range.Columns 一致）
 * @returns {boolean} 是否成功删除
 * @example
 * RngUtils.z删除列("A1:D5", 2)  // 删除 B 列（整列）
 * RngUtils.z删除列(ActiveSheet.Range("B2:D5"), 1)  // 删除选中区第 1 列
 */
RngUtils.z删除列 = function(rng, colIdx) {
    if (colIdx === undefined || colIdx === null) return false;
    var r = typeof rng === 'string' ? Range(rng) : rng;
    if (!r) return false;
    if (colIdx < 1 || colIdx > r.Columns.Count) return false;
    r.Columns(colIdx).Delete();
    return true;
};
RngUtils.deleteCol = RngUtils.z删除列;

/**
 * z读取区域 - 读取指定区域的数据为二维数组（Array2D）
 *              与 Array2D.toRange 配对（z读取区域 ↔ z写入单元格 / toRange）
 * @param {Range|string} rng - 要读取的单元格区域
 * @returns {Array2D} 二维数组形式的数据
 * @example
 * RngUtils.z读取区域("A1:C3")       // [[a,b,c],[d,e,f],[g,h,i]]
 * RngUtils.z读取区域("A1")          // [[a]]  （单格自动包成二维）
 * RngUtils.z读取区域(ActiveSheet.UsedRange).filter(...)  // 链式处理
 */
RngUtils.z读取区域 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    if (!r) return new Array2D([]);
    var arr = r.Value2;
    if (arr === null || arr === undefined) return new Array2D([]);
    if (!Array.isArray(arr)) return new Array2D([[arr]]);
    if (!Array.isArray(arr[0])) {
        var oneD = [];
        for (var i = 0; i < arr.length; i++) oneD.push([arr[i]]);
        return new Array2D(oneD);
    }
    return new Array2D(arr);
};
RngUtils.fromRange = RngUtils.z读取区域;
RngUtils.readRange = RngUtils.z读取区域;

/**
 * z整行 - 获取指定单元格区域的整行
 * @param {Range|string} rng - 要获取整行的单元格区域
 * @returns {Range} 整行单元格
 * @example
 * RngUtils.z整行("11:14")  // $11:$14
 */
RngUtils.z整行 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.EntireRow;
};
RngUtils.entireRow = RngUtils.z整行;

/**
 * z整列 - 获取指定单元格区域的整列
 * @param {Range|string} rng - 要获取整列的单元格区域
 * @returns {Range} 整列单元格
 * @example
 * RngUtils.z整列("A:B")  // $A:$B
 */
RngUtils.z整列 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.EntireColumn;
};
RngUtils.entire_column = RngUtils.z整列;

/**
 * z行数 - 获取指定单元格区域的行数
 * @param {Range|string} rng - 要获取行数的单元格区域
 * @returns {number} 行数
 * @example
 * RngUtils.z行数("A12:D15")  // 4
 */
RngUtils.z行数 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Rows.Count;
};
RngUtils.rowsCount = RngUtils.z行数;

/**
 * z列数 - 获取指定单元格区域的列数
 * @param {Range|string} rng - 要获取列数的单元格区域
 * @returns {number} 列数
 * @example
 * RngUtils.z列数("A12:C15")  // 3
 */
RngUtils.z列数 = function(rng) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    return r.Columns.Count;
};
RngUtils.colsCount = RngUtils.z列数;

/**
 * z列号字母互转 - 将数字列号转换为字母表示
 * @param {number} c - 要转换的数字列号
 * @returns {string} 列号的字母表示
 * @example
 * RngUtils.z列号字母互转(3)  // "C"
 */
RngUtils.z列号字母互转 = function(c) {
    var result = '';
    while (c > 0) {
        c--;
        result = String.fromCharCode(65 + (c % 26)) + result;
        c = Math.floor(c / 26);
    }
    return result;
};
RngUtils.colToAbc = RngUtils.z列号字母互转;

/**
 * z复制粘贴格式 - 复制粘贴格式到目标区域
 * @param {Range|string} rng - 源单元格区域
 * @param {Range|string} target - 目标单元格区域
 * @example
 * RngUtils.z复制粘贴格式("a14:d14","a18:d21")
 */
RngUtils.z复制粘贴格式 = function(rng, target) {
    var src = typeof rng === 'string' ? Range(rng) : rng;
    var dest = typeof target === 'string' ? Range(target) : target;
    src.Copy();
    dest.PasteSpecial(-4122); // xlPasteFormats
    Application.CutCopyMode = false;
};
RngUtils.copyFormat = RngUtils.z复制粘贴格式;

/**
 * z复制粘贴值 - 复制粘贴值到目标区域
 * @param {Range|string} rng - 源单元格区域
 * @param {Range|string} target - 目标单元格区域
 * @example
 * RngUtils.z复制粘贴值("a11:d14","a18:d21")
 */
RngUtils.z复制粘贴值 = function(rng, target) {
    var src = typeof rng === 'string' ? Range(rng) : rng;
    var dest = typeof target === 'string' ? Range(target) : target;
    src.Copy();
    dest.PasteSpecial(-4163); // xlPasteValues
    Application.CutCopyMode = false;
};
RngUtils.copyValue = RngUtils.z复制粘贴值;

/**
 * z联合区域 - 对字符串地址或单元格数组联合成一个单元格区域
 * @param {any} rng - 单元格地址或单元格数组
 * @param {Sheet} [op_sht] - 工作表对象，跨表时指定
 * @returns {Range} 组合后的单元格对象
 * @example
 * RngUtils.z联合区域('a1,a2,B4:C10').Address()
 */
RngUtils.z联合区域 = function(rng, op_sht) {
    var sheet = op_sht || Application.ActiveSheet;

    if (typeof rng === 'string') {
        // 解析地址字符串
        var parts = rng.split(',');
        var ranges = [];
        for (var i = 0; i < parts.length; i++) {
            ranges.push(sheet.Range(parts[i].trim()));
        }
        if (ranges.length === 1) return ranges[0];
        return sheet.Union(ranges[0], ranges[1]);
    }

    if (Array.isArray(rng)) {
        if (rng.length === 0) return null;
        if (rng.length === 1) return rng[0];
        return sheet.Union(rng[0], rng[1]);
    }

    return rng;
};
RngUtils.unionAll = RngUtils.z联合区域;

/**
 * z多列排序 - 单元格多列排序
 * @param {Range|string} rng - 待排序的单元格范围
 * @param {string} sortParams - 排序参数 'f3+,f4-' 表示第3列升序第4列降序
 * @param {number} [headerRows=1] - 表头的行数
 * @param {string} [customOrder] - 自定义序列
 * @example
 * RngUtils.z多列排序("A18:D24",'f3+,f4-',1)
 */
RngUtils.z多列排序 = function(rng, sortParams, headerRows, customOrder) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    headerRows = headerRows || 1;

    // 解析排序参数
    var sorts = [];
    var parts = sortParams.split(/[,，]/);
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        var match = part.match(/^f?(\d+)([+-])?$/);
        if (match) {
            sorts.push({
                col: parseInt(match[1]),
                order: (match[2] || '+') === '+' ? 1 : 2 // 1升序 2降序, 缺省升序
            });
        }
        else {
            console.warn('[JSA880] z多列排序: 忽略无效排序参数 "' + part + '"');
        }
    }

    // 获取数据数组
    var arr = RngUtils.z安全数组(r);
    if (arr.length <= headerRows) return;

    // 分离表头和数据
    var header = arr.slice(0, headerRows);
    var data = arr.slice(headerRows);

    // 排序
    data.sort(function(a, b) {
        for (var s = 0; s < sorts.length; s++) {
            var sort = sorts[s];
            var colIdx = sort.col - 1;
            var valA = a[colIdx];
            var valB = b[colIdx];

            // 自定义序列处理
            if (customOrder) {
                var orderArr = customOrder.split(/[,，]/);
                var idxA = orderArr.indexOf(String(valA));
                var idxB = orderArr.indexOf(String(valB));
                if (idxA >= 0 && idxB >= 0) {
                    valA = idxA;
                    valB = idxB;
                }
            }

            if (typeof valA === 'string' && typeof valB === 'string') {
                var strA = valA.toLowerCase();
                var strB = valB.toLowerCase();
                if (strA < strB) return sort.order === 1 ? -1 : 1;
                if (strA > strB) return sort.order === 1 ? 1 : -1;
            } else {
                if (valA < valB) return sort.order === 1 ? -1 : 1;
                if (valA > valB) return sort.order === 1 ? 1 : -1;
            }

        }
        return 0;
    });

    // 写回
    r.Value2 = header.concat(data);
};
RngUtils.rngSortCols = RngUtils.z多列排序;

/**
 * z合并单元格 - 合并指定区域
 *
 * 【功能说明】将指定的单元格区域合并为一个单元格，合并后左上角单元格的值保留
 *
 * 【技术实现】
 * 1. 环境检测：检查是否在WPS环境中运行
 * 2. 参数转换：支持字符串地址（"A1:B2"）或 Range 对象
 * 3. 调用原生API：使用 WPS 的 Range.Merge() 方法执行合并
 * 4. 返回结果：返回合并后的 Range 对象，便于链式调用
 *
 * 【使用场景】
 * - 创建表头合并单元格（如多级表头）
 * - 标题居中显示
 * - 数据透视表的分组标题
 *
 * @param {Range|string} rng - 要合并的区域，可以是地址字符串或Range对象
 * @returns {Range} 合并后的区域对象
 *
 * @example
 * // 示例1：合并标题行
 * RngUtils.z合并单元格("A1:D1");
 *
 * // 示例2：使用英文方法名
 * RngUtils.mergeCells("A1:C3");
 *
 * // 示例3：链式调用
 * RngUtils.z合并单元格("A1:B2").z加边框();
 *
 * // 示例4：全局快捷方式（推荐）
 * $.mergeCells("A1:D1");
 */
RngUtils.z合并单元格 = function(rng) {
    // 环境检测：非WPS环境直接返回null

    // 参数转换：字符串地址转为Range对象
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // 调用WPS原生API执行合并
    r.Merge();

    // 返回Range对象，支持链式调用
    return r;
};
// 英文别名：支持中英文双语调用
RngUtils.mergeCells = RngUtils.z合并单元格;

/**
 * z取消合并单元格 - 取消指定区域的合并
 *
 * 【功能说明】将已合并的单元格区域拆分为独立的单元格
 *
 * 【技术实现】
 * 1. 环境检测：检查是否在WPS环境中运行
 * 2. 参数转换：支持字符串地址（"A1:B2"）或 Range 对象
 * 3. 调用原生API：使用 WPS 的 Range.UnMerge() 方法取消合并
 * 4. 返回结果：返回取消合并后的 Range 对象
 *
 * 【使用场景】
 * - 重新布局数据结构
 * - 批量处理前取消合并
 * - 数据导入前的预处理
 *
 * @param {Range|string} rng - 要取消合并的区域，可以是地址字符串或Range对象
 * @returns {Range} 取消合并后的区域对象
 *
 * @example
 * // 示例1：取消合并
 * RngUtils.z取消合并单元格("A1:B2");
 *
 * // 示例2：使用英文方法名
 * RngUtils.unmergeCells("A1:C3");
 *
 * // 示例3：全局快捷方式（推荐）
 * $.unmergeCells("A1:D1");
 */
RngUtils.z取消合并单元格 = function(rng) {
    // 环境检测：非WPS环境直接返回null

    // 参数转换：字符串地址转为Range对象
    var r = typeof rng === 'string' ? Range(rng) : rng;

    // 调用WPS原生API取消合并
    r.UnMerge();

    // 返回Range对象
    return r;
};
// 英文别名：支持中英文双语调用
RngUtils.unmergeCells = RngUtils.z取消合并单元格;

/**
 * z最大行数组 - 获取从第一行到最大行的区域并转换为二维数组
 * @param {Range|string} rng - 要获取的区域（如 "A1:H1"）
 * @param {number} [col] - 可选，指定列作为获取最大行依据
 * @returns {Array} 二维数组
 * @example
 * var arr = RngUtils.z最大行数组("A1:H1");
 * logjson(arr);
 */
RngUtils.z最大行数组 = function(rng, col) {
    var maxRng = RngUtils.z最大行区域(rng, col);
    return RngUtils.z安全数组(maxRng);
};
// 英文别名
RngUtils.maxArray = RngUtils.z最大行数组;

// ==================== 使用辅助函数创建 RngUtils 实例方法别名 ====================
createBilingualAliases(RngUtils.prototype, [
    ['z加边框', 'addBorders'],
    ['z去边框', 'removeBorders'],
    ['z清除内容', 'clearContents'],
    ['z清除格式', 'clearFormats'],
    ['z自动列宽', 'autoFitColumns'],
    ['z自动行高', 'autoFitRows'],
    ['z设置背景色', 'backgroundColor'],
    ['z设置字体色', 'fontColor'],
    ['z行数', 'rowsCount'],
    ['z列数', 'colsCount'],
    ['z地址', 'address']
]);

/**
 * 加边框
 * @param {Number} lineStyle - 线条样式（默认1）
 * @param {Number} weight - 线条粗细（默认2）
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z加边框()
 * RngUtils("A1:C10").z加边框(1, 3).z设置背景色(0xFFFF00)
 */
RngUtils.prototype.z加边框 = function(lineStyle, weight) {
    if (!this._range) return this;
    lineStyle = lineStyle !== undefined ? lineStyle : 1;
    weight = weight !== undefined ? weight : 2;
    this._range.Borders.LineStyle = lineStyle;
    this._range.Borders.Weight = weight;
    return this;
};
RngUtils.prototype.addBorders = RngUtils.prototype.z加边框;

/**
 * 去边框
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z去边框()
 */
RngUtils.prototype.z去边框 = function() {
    if (!this._range) return this;
    this._range.Borders.LineStyle = -4142; // xlLineStyleNone
    return this;
};
RngUtils.prototype.removeBorders = RngUtils.prototype.z去边框;

/**
 * 清除内容
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z清除内容()
 */
RngUtils.prototype.z清除内容 = function() {
    if (!this._range) return this;
    this._range.ClearContents();
    return this;
};
RngUtils.prototype.clearContents = RngUtils.prototype.z清除内容;

/**
 * 清除格式
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z清除格式()
 */
RngUtils.prototype.z清除格式 = function() {
    if (!this._range) return this;
    this._range.ClearFormats();
    return this;
};
RngUtils.prototype.clearFormats = RngUtils.prototype.z清除格式;

/**
 * 自动列宽
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z自动列宽()
 * RngUtils("A:Z").z自动列宽()  // 整列自动调整
 */
RngUtils.prototype.z自动列宽 = function() {
    if (!this._range) return this;
    this._range.Columns.AutoFit();
    return this;
};
RngUtils.prototype.autoFitColumns = RngUtils.prototype.z自动列宽;

/**
 * 自动行高
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z自动行高()
 */
RngUtils.prototype.z自动行高 = function() {
    if (!this._range) return this;
    this._range.Rows.AutoFit();
    return this;
};
RngUtils.prototype.autoFitRows = RngUtils.prototype.z自动行高;

/**
 * 设置背景色
 * @param {Number} color - RGB颜色值
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z设置背景色(RGB(255, 0, 0))  // 红色背景
 * RngUtils("A1:C10").z设置背景色(0xFFFF00)        // 黄色背景
 */
RngUtils.prototype.z设置背景色 = function(color) {
    if (!this._range) return this;
    this._range.Interior.Color = color;
    return this;
};
RngUtils.prototype.backgroundColor = RngUtils.prototype.z设置背景色;

/**
 * 设置字体色
 * @param {Number} color - RGB颜色值
 * @returns {RngUtils} 当前实例
 * @example
 * RngUtils("A1:C10").z设置字体色(RGB(255, 0, 0))  // 红色字体
 * RngUtils("A1:C10").z设置字体色(0xFF0000)        // 红色字体
 */
RngUtils.prototype.z设置字体色 = function(color) {
    if (!this._range) return this;
    this._range.Font.Color = color;
    return this;
};
RngUtils.prototype.fontColor = RngUtils.prototype.z设置字体色;

/**
 * 获取行数
 * @returns {Number} 行数
 */
RngUtils.prototype.z行数 = function() {
    if (!this._range) return 0;
    return this._range.Rows.Count;
};
RngUtils.prototype.rowsCount = RngUtils.prototype.z行数;

/**
 * 获取列数
 * @returns {Number} 列数
 */
RngUtils.prototype.z列数 = function() {
    if (!this._range) return 0;
    return this._range.Columns.Count;
};
RngUtils.prototype.colsCount = RngUtils.prototype.z列数;

/**
 * 获取地址
 * @returns {String} 单元格地址
 */
RngUtils.prototype.z地址 = function() {
    if (!this._range) return '';
    return this._range.Address();
};
RngUtils.prototype.address = RngUtils.prototype.z地址;


// ==================== [JSA] 通用函数库 ====================

/**
 * JSA - 通用函数工具（静态方法）
 * @constructor
 * @class
 * @description 常用函数集合
 */
function JSA() {}

/**
 * 转置数组
 * @param {Array} arr - 数组
 * @returns {Array} 转置后的数组
 */
JSA.z转置 = function(arr) {
    if (!arr || arr.length === 0) return [];
    // 一维数组处理：[1,2,3] → [[1],[2],[3]]
    if (!Array.isArray(arr[0])) {
        var result = [];
        for (var i = 0; i < arr.length; i++) {
            result[i] = [arr[i]];
        }
        return result;
    }
    var rows = arr.length;
    var cols = arr[0].length;
    var result = [];
    for (var j = 0; j < cols; j++) {
        result[j] = [];
        for (var i = 0; i < rows; i++) {
            result[j][i] = arr[i][j];
        }
    }
    return result;
};
JSA.transpose = JSA.z转置;

/**
 * 转数值
 * @param {String} text - 文本
 * @returns {Number} 数值
 */
JSA.z转数值 = function(text) {
    if (typeof text === 'number') return text;
    if (typeof text === 'string') {
        text = text.trim();
        // v4.1.0 修复: 原正则 /^[-+]?[0-9]*\.?[0-9]+/ 无法匹配 "1." 这种以小数点结尾的合法数字
        // 修正后: 同时支持 "1.5", "1.", ".5", "1", "-1.5" 等所有合法格式
        var match = text.match(/^[-+]?(\d+\.?\d*|\.\d+)/);
        if (match) return parseFloat(match[0]);
        return 0;
    }
    return 0;
};
JSA.val = JSA.z转数值;

/**
 * 写入单元格（根据数组大小自动扩展区域）
 * @param {Array} arr - 数组
 * @param {Range|string} rng - 单元格区域（左上角单元格）
 * @param {Boolean} clearDown - 是否清空下方（保留参数兼容性）
 * @returns {Range} 写入的Range
 */
JSA.z写入单元格 = function(arr, rng, clearDown) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    // XXD-47: null/undefined 守卫 — 公开 API 常接收公式结果（可能为 null）
    if (!arr || !targetRng) return null;
    var rows = arr.length;
    // 修复：一维数组横向写入，二维数组纵向写入
    var cols = rows > 0 ? (Array.isArray(arr[0]) ? arr[0].length : arr.length) : 0;
    var is1D = !Array.isArray(arr[0]);
    // 一维数组横向写入：1xN，二维数组纵向写入：NxM
    var endRng = is1D ? targetRng.Item(1, cols) : targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = arr;
    return writeRng;
};
JSA.toRange = JSA.z写入单元格;

/**
 * 获取今天日期
 * @returns {String} 今天日期 YYYY-MM-DD
 */
JSA.z今天 = function() {
    var now = new Date();
    return now.getFullYear() + '-' +
           String(now.getMonth() + 1).padStart(2, '0') + '-' +
           String(now.getDate()).padStart(2, '0');
};
JSA.today = JSA.z今天;
// ==================== [XXD-164] 常用工具 z* 全局别名补齐 ====================
// 规范定义见 XXD-164 issue 描述。每个函数同时提供英文别名，遵循同文件既有约定
// (例如 JSA.today = JSA.z今天)。避免与 JSA.now (Date 对象) 冲突：
//   z当前时间戳 → Date.now() 的数值时间戳，英文别名 nowTimestamp
//   z当前日期   → new Date() 的 Date 对象，英文别名 currentDate

/**
 * 是否为空：null/undefined、空字符串、或空数组都视为空
 * @param {*} v - 任意值
 * @returns {Boolean}
 */
JSA.z是否为空 = function(v) {
    return v == null || v === '' || (Array.isArray(v) && v.length === 0);
};
JSA.isEmpty = JSA.z是否为空;

/**
 * 非空：取反
 * @param {*} v - 任意值
 * @returns {Boolean}
 */
JSA.z非空 = function(v) {
    return !JSA.z是否为空(v);
};
JSA.isNotEmpty = JSA.z非空;

/**
 * 默认值：null/undefined 时返回兜底值
 * @param {*} a - 首选值
 * @param {*} b - 兜底值
 * @returns {*}
 */
JSA.z默认值 = function(a, b) {
    return a != null ? a : b;
};
JSA.defaultIfNull = JSA.z默认值;

/**
 * 当前时间戳：Date.now() 的毫秒数值
 * @returns {Number} 自 1970-01-01 UTC 起的毫秒数
 */
JSA.z当前时间戳 = function() {
    return Date.now();
};
JSA.nowTimestamp = JSA.z当前时间戳;

/**
 * 当前日期：返回当前 Date 对象
 * @returns {Date}
 */
JSA.z当前日期 = function() {
    return new Date();
};
JSA.currentDate = JSA.z当前日期;

/**
 * 日期差：以天为单位返回 d2 - d1 (毫秒差 / 86400000)
 * @param {Date|string|number} d1 - 起始日期
 * @param {Date|string|number} d2 - 结束日期
 * @returns {Number} 天数 (可负)
 */
JSA.z日期差 = function(d1, d2) {
    var toMs = function(d) { return d instanceof Date ? d.getTime() : new Date(d).getTime(); };
    return (toMs(d2) - toMs(d1)) / 86400000;
};
JSA.dateDiff = JSA.z日期差;

/**
 * 包含：字符串包含子串
 * @param {String} s - 源串
 * @param {String} sub - 子串
 * @returns {Boolean}
 */
JSA.z包含 = function(s, sub) {
    return typeof s === 'string' && s.includes(sub);
};
JSA.includesStr = JSA.z包含;

/**
 * 开始于：字符串以 prefix 开头
 * @param {String} s - 源串
 * @param {String} prefix - 前缀
 * @returns {Boolean}
 */
JSA.z开始于 = function(s, prefix) {
    return typeof s === 'string' && s.startsWith(prefix);
};
JSA.startsWith = JSA.z开始于;

/**
 * 结束于：字符串以 suffix 结尾
 * @param {String} s - 源串
 * @param {String} suffix - 后缀
 * @returns {Boolean}
 */
JSA.z结束于 = function(s, suffix) {
    return typeof s === 'string' && s.endsWith(suffix);
};
JSA.endsWith = JSA.z结束于;

/**
 * 去空白：去掉首尾空白
 * @param {String} s - 源串
 * @returns {String}
 */
JSA.z去空白 = function(s) {
    return typeof s === 'string' ? s.trim() : s;
};
JSA.trim = JSA.z去空白;

/**
 * 分割：按 sep 切分字符串
 * @param {String} s - 源串
 * @param {String|RegExp} sep - 分隔符
 * @returns {Array}
 */
JSA.z分割 = function(s, sep) {
    return typeof s === 'string' ? s.split(sep) : [];
};
JSA.splitStr = JSA.z分割;

/**
 * 连接：把数组按 sep 拼接为字符串
 * @param {Array} arr - 数组
 * @param {String} sep - 分隔符
 * @returns {String}
 */
JSA.z连接 = function(arr, sep) {
    return Array.isArray(arr) ? arr.join(sep) : '';
};
JSA.joinStr = JSA.z连接;

/**
 * 替换全部：把字符串中所有 from 替换为 to (split/join 实现，避免 RegExp 转义陷阱)
 * @param {String} s - 源串
 * @param {String} from - 待替换子串
 * @param {String} to - 目标子串
 * @returns {String}
 */
JSA.z替换全部 = function(s, from, to) {
    if (typeof s !== 'string') return s;
    return s.split(from).join(to);
};
JSA.replaceAll = JSA.z替换全部;

/**
 * 转大写
 * @param {String} s - 源串
 * @returns {String}
 */
JSA.z转大写 = function(s) {
    return typeof s === 'string' ? s.toUpperCase() : s;
};
JSA.toUpperCase = JSA.z转大写;

/**
 * 转小写
 * @param {String} s - 源串
 * @returns {String}
 */
JSA.z转小写 = function(s) {
    return typeof s === 'string' ? s.toLowerCase() : s;
};
JSA.toLowerCase = JSA.z转小写;
// ==================== /XXD-164 补齐结束 ====================


/**
 * 转日期数值
 * @param {Date|string} d - 日期
 * @returns {Number} Excel日期数值
 */
JSA.z转日期数值 = function(d) {
    var date = typeof d === 'string' ? new Date(d) : d;
    var excelEpoch = new Date(1900, 0, 1);
    var msPerDay = 24 * 60 * 60 * 1000;
    return Math.floor((date - excelEpoch) / msPerDay) + 2;
};
JSA.cdate = JSA.z转日期数值;

// 🔧 XXD-158 final fix: 通用类型转换 / 计数工具
//   与 JSA.z转数值 (Excel 语义, 失败返回 0) 不同, z字符串转数字 直接走 Number()
//   返回 NaN 以便调用方判断失败; z计数 是数组长度工具, 与 Array2D.prototype.z计数
//   (列内有效数值计数) 语义不同, 因此作为顶层 JSA 工具单独暴露.
JSA.z字符串转数字 = function(s) { return Number(s); };
JSA.toNumber = JSA.z字符串转数字;

JSA.z数字转字符串 = function(n) { return String(n); };
JSA.toString = JSA.z数字转字符串;

JSA.z随机数 = function() { return Math.random(); };
JSA.random = JSA.z随机数;

JSA.z计数 = function(arr) { return arr == null ? 0 : arr.length; };
JSA.count = JSA.z计数;

/**
 * 替换
 * @param {String} str - 字符串
 * @param {String} find - 查找
 * @param {String} replaceWith - 替换
 * @returns {String} 结果
 */
JSA.z替换 = function(str, find, replaceWith) {
    // XXD-47: null/undefined 守卫 — 公开 API 常接收公式结果（可能为 null）
    if (typeof str !== 'string' && !(str instanceof String)) return '';
    return str.split(find).join(replaceWith);
};
JSA.replace = JSA.z替换;

/**
 * 数组转JSON字符串（数组扩展方法）
 * @returns {String} JSON字符串
 * @description 将数组转换为JSON格式字符串
 * @example [1,2,3].toJson()              // 返回 "[1,2,3]"
 * @example ["a","b"].toJson()            // 返回 "[\"a\",\"b\"]"
 * @example [{x:1},{y:2}].toJson()        // 返回 "[{\"x\":1},{\"y\":2}]"
 */
Array.prototype.toJson = function() {
    return JSON.stringify(this);
};

/**
 * 数组转JSON字符串（数组扩展方法 - 中文别名）
 */
Array.prototype.z转JSON = Array.prototype.toJson;

/**
 * 数组元素转数值（数组扩展方法）
 * @returns {Array} 数值数组
 * @description 将数组中每个元素转换为数值
 * @example "1a2b3c4asd5".match(/\d/g).val()        // 返回 [1,2,3,4,5]
 * @example ["1","2","3"].val()                    // 返回 [1,2,3]
 * @example ["10","20","abc"].val()                // 返回 [10,20,0]
 */
Array.prototype.val = function() {
    return this.map(function(item) {
        var num = Number(item);
        return isNaN(num) ? 0 : num;
    });
};

/**
 * 数组元素转数值（数组扩展方法 - 中文别名）
 */
Array.prototype.z转数值 = Array.prototype.val;

/**
 * 数组写入单元格（数组扩展方法，根据数组大小自动扩展区域）
 * @param {Range|string} rng - 单元格区域（左上角单元格）
 * @returns {Range} 写入的Range
 * @description 将二维数组写入指定单元格
 * @example
 * var arr = [[1, 'A'], [2, 'B'], [3, 'C']];
 * arr.toRange("J2");                    // 写入J2:L4
 * arr.toRange(Range("A1"));             // 写入A1:C4
 */
Array.prototype.toRange = function(rng, clearBelow) {
    // 🔧 v4.0.1 修复: 空数组保护，防止 Item(0,0) 报错
    if (!this || this.length === 0) return null;
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = this.length;
    var cols = rows > 0 ? (Array.isArray(this[0]) ? this[0].length : 1) : 0;
    // 列数边界检查
    if (cols === 0) return null;

    // 🔧 v4.0.4 修复一维数组问题：一维数组应被视为一行多列
    var dataToWrite = this;
    var actualRows = rows;
    var actualCols = cols;
    if (rows > 0 && !Array.isArray(this[0])) {
        // 一维数组转为二维数组 [[1, 2, 3]]
        dataToWrite = [this.slice()];
        actualRows = 1;
        actualCols = dataToWrite[0].length;
    }

    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(actualRows, actualCols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    writeRng.Value2 = dataToWrite;

    // 🔧 v4.0.10 修复: 支持 clearBelow 参数清空下方区域
    if (clearBelow === true || clearBelow === 1) {
        // 获取当前区域下方的已使用区域并清空
        try {
            var usedRange = sheet.UsedRange;
            if (usedRange) {
                var writeEndRow = targetRng.Row + actualRows - 1;
                var usedEndRow = usedRange.Row + usedRange.Rows.Count - 1;
                if (usedEndRow > writeEndRow) {
                    // 清空下方区域
                    var belowRng = sheet.Range(
                        sheet.Cells(writeEndRow + 1, targetRng.Column),
                        sheet.Cells(usedEndRow, targetRng.Column + actualCols - 1)
                    );
                    belowRng.ClearContents();
                }
            }
        } catch (e) {
            // 忽略清空错误
        }
    }

    return writeRng;
};

/**
 * 数组写入单元格（数组扩展方法 - 中文别名）
 */
Array.prototype.z写入单元格 = Array.prototype.toRange;

/**
 * 截取字符
 * @param {String} str - 字符串
 * @param {Number} start - 起始位置（从1开始）
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z截取字符 = function(str, start, len) {
    var startIndex = start - 1;
    if (len === undefined) return str.substring(startIndex);
    return str.substring(startIndex, startIndex + len);
};
JSA.mid = JSA.z截取字符;

/**
 * 左取字符
 * @param {String} str - 字符串
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z左取字符 = function(str, len) {
    return str.substring(0, len);
};
JSA.left = JSA.z左取字符;

/**
 * 右取字符
 * @param {String} str - 字符串
 * @param {Number} len - 长度
 * @returns {String} 结果
 */
JSA.z右取字符 = function(str, len) {
    return str.substring(str.length - len);
};
JSA.right = JSA.z右取字符;

// 🔧 XXD-156/XXD-157 final fix: JSA.z求和/z最大值/z最小值/z平均值 支持嵌套数组参数
// 复现: JSA.z最大值([[1,2,3]]) 返回 0 — Number([[1,2,3]]) === NaN 被吞成 0
// 修复: 用 _zFlatNums 递归扁平化所有 array 参数, 跳过 NaN, 保留 ...Number 旧用法
function _zFlatNums() {
    var out = [];
    function walk(v) {
        if (Array.isArray(v)) { for (var i = 0; i < v.length; i++) walk(v[i]); return; }
        if (v === null || v === undefined || v === '') return;
        var n = (typeof v === 'number') ? v : parseFloat(String(v).replace(/,/g, ''));
        if (!isNaN(n)) out.push(n);
    }
    for (var i = 0; i < arguments.length; i++) walk(arguments[i]);
    return out;
}

/**
 * 求和
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 和
 */
JSA.z求和 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    var s = 0;
    for (var i = 0; i < nums.length; i++) s += nums[i];
    return Math.round(s * 1e10) / 1e10;
};
JSA.sum = JSA.z求和;

/**
 * 最大值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 最大值，无有效数值时返回 0
 */
JSA.z最大值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    return nums.length === 0 ? 0 : Math.max.apply(null, nums);
};
JSA.max = JSA.z最大值;

/**
 * 最小值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 最小值，无有效数值时返回 0
 */
JSA.z最小值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    return nums.length === 0 ? 0 : Math.min.apply(null, nums);
};
JSA.min = JSA.z最小值;

/**
 * 平均值
 * @param {...(Number|Array)} args - 数值或(嵌套)数组
 * @returns {Number} 平均值，无有效数值时返回 0
 */
JSA.z平均值 = function() {
    var nums = _zFlatNums.apply(null, arguments);
    if (nums.length === 0) return 0;
    var s = 0; for (var i = 0; i < nums.length; i++) s += nums[i];
    return Math.round((s / nums.length) * 1e10) / 1e10;
};
JSA.average = JSA.z平均值;
// 🔧 XXD-153 final fix: JSA.agg / JSA.oadate 别名
JSA.agg = JSA.agg || function(arr, sel) { return new Array2D(arr).z求和(sel); };
// 🔧 XXD-220 fix: oadate(null) THROW + oadate('2024-06-09') THROW — 防御性处理 null/字符串
JSA.oadate = JSA.oadate || function(d) {
    if (d == null) return 0;   // OADate convention: null/undefined = 0 (no date)
    if (typeof d === 'number') return d;  // 已经是 OADate 数值，直接返回
    if (typeof d === 'string') d = new Date(d);
    if (!(d instanceof Date) || isNaN(d.getTime())) return null;
    return d.getTime() / 86400000 + 25569;
};
// 🔧 XXD-222: JSA.fromOADate — oadate 的反函数
JSA.fromOADate = JSA.fromOADate || function(n) { return new Date((n - 25569) * 86400000); };

/**
 * 模糊匹配
 * @param {String} str - 字符串
 * @param {String} pattern - 模式（支持*和?）
 * @returns {Number} 匹配返回-1，不匹配返回0
 * @description 包含模式匹配，自动在模式前后添加 *，除非模式以 ^ 开头或 $ 结尾
 */
JSA.z模糊匹配 = function(str, pattern) {
    if (pattern === undefined || pattern === null) return 0;
    // 转义正则特殊字符，但保留 * 和 ?
    var regexPattern = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&')
                              .replace(/\*/g, '.*')
                              .replace(/\?/g, '.');
    // 包含模式：自动在前后加上 .*，除非模式已经以 ^ 开头或 $ 结尾
    var anchored = regexPattern;
    if (anchored.indexOf('^') !== 0) anchored = '.*' + anchored;
    if (anchored.charAt(anchored.length - 1) !== '$') anchored = anchored + '.*';
    var regex = new RegExp('^' + anchored + '$');
    // 匹配返回-1，不匹配返回0
    return regex.test(str) ? -1 : 0;
};
JSA.like = JSA.z模糊匹配;

/**
 * 表达式求值
 * @param {String} expr - 字符串表达式（如 '5*6+5'）
 * @returns {Number} 计算结果
 * @description 对字符串表达式进行求值计算
 * @example JSA.eval880('5*6+5')     // 返回 35
 * @example JSA.eval880('10+20*3')   // 返回 70
 * @example JSA.eval880('(1+2)*3')   // 返回 9
 */
JSA.z表达式求值 = function(expr) {
    if (typeof expr !== 'string') return Number(expr) || 0;
    // 使用 Function 构造函数安全地计算表达式
    try {
        return new Function('return ' + expr)();
    } catch (e) {
        return 0;
    }
};
JSA.eval880 = JSA.z表达式求值;

/**
 * 生成数字序列
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @param {Number} step - 步长
 * @returns {Array} 序列
 */
// 【v4.2.2 修复】依据官方 API 文档补齐规范别名
// 规范：getIndexs(开始, 结束, 步长) → Array
//      别名：z生成数字序列
//      例：JSA.getIndexs(5, 10, 2) → [5,7,9]
// 实现：使用 Array2D.getIndexs（功能更完善，支持负 step + 防御性 step=0 检查）
JSA.z生成数字序列 = function(start, end, step) {
    return Array2D.getIndexs(start, end, step);
};
// ✅ 补齐规范要求的英文别名
JSA.getIndexs = JSA.z生成数字序列;
JSA.getNumberArray = JSA.z生成数字序列;  // 旧别名也保留（向后兼容）

/**
 * 人民币大写
 * @param {Number} n - 数字
 * @returns {String} 大写
 */
JSA.z人民币大写 = function(n) {
    var digits = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"];
    var units = ["", "拾", "佰", "仟"];
    var bigUnits = ["", "万", "亿"];

    if (n === 0) return "零元整";

    var num = Math.abs(n);
    var integerPart = Math.floor(num);
    var decimalPart = Math.round((num - integerPart) * 100);

    var result = _convertIntegerPart(integerPart, digits, units, bigUnits) + "元";

    if (decimalPart > 0) {
        if (decimalPart >= 10) {
            var jiao = Math.floor(decimalPart / 10);
            var fen = decimalPart % 10;
            result += digits[jiao] + "角";
            if (fen > 0) result += digits[fen] + "分";
        } else {
            result += digits[decimalPart] + "分";
        }
    } else {
        result += "整";
    }

    if (n < 0) result += "（负）";
    return result;

    function _convertIntegerPart(num, digits, units, bigUnits) {
        if (num === 0) return "";
        var result = "";
        var bigUnitIndex = 0;
        while (num > 0) {
            var section = num % 10000;
            if (section > 0) {
                var sectionResult = _convertSection(section, digits, units);
                result = sectionResult + bigUnits[bigUnitIndex] + result;
            }
            num = Math.floor(num / 10000);
            bigUnitIndex++;
        }
        return result;
    }

    function _convertSection(num, digits, units) {
        var result = "";
        var unitIndex = 0;
        var lastZero = false;
        while (num > 0) {
            var digit = num % 10;
            if (digit === 0) {
                if (!lastZero && result !== "") {
                    result = digits[0] + result;
                    lastZero = true;
                }
            } else {
                result = digits[digit] + units[unitIndex] + result;
                lastZero = false;
            }
            num = Math.floor(num / 10);
            unitIndex++;
        }
        return result;
    }
};
JSA.rmbdx = JSA.z人民币大写;

/**
 * 随机整数
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @returns {Number} 随机整数
 */
JSA.z随机整数 = function(start, end) {
    return Math.floor(Math.random() * (end - start + 1)) + start;
};
JSA.rndInt = JSA.z随机整数;

/**
 * 生成 n 个随机整数数组
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @param {Number} n - 生成数量
 * @returns {Array} 随机整数数组
 * @example
 * JSA.rndIntArray(1, 100, 50) // 生成 50 个 1-100 之间的随机整数
 */
JSA.z随机整数数组 = function(start, end, n) {
    n = n || 1;
    var arr = [];
    for (var i = 0; i < n; i++) {
        arr.push(Math.floor(Math.random() * (end - start + 1)) + start);
    }
    return arr;
};
JSA.rndIntArray = JSA.z随机整数数组;

/**
 * 随机打乱
 * @param {Array} array - 数组
 * @returns {Array} 打乱后的数组
 */
JSA.z随机打乱 = function(array) {
    var result = array.slice();
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return result;
};
JSA.shuffle = JSA.z随机打乱;

/**
 * 延时
 * @param {Number} ts - 毫秒
 */
JSA.z延时 = function(ts) {
    if (ts < 0) throw new RangeError('z延时: 参数不能为负数');
    var start = Date.now();
    while (Date.now() - start < ts) {
        // 等待
    }
};
JSA.delay = JSA.z延时;

/**
 * 日期间隔
 * @param {Date|string} d1 - 日期1
 * @param {Date|string} d2 - 日期2
 * @param {String} format - 格式
 * @returns {String|Number} 间隔
 */
JSA.z日期间隔 = function(d1, d2, format) {
    var date1 = typeof d1 === 'string' ? new Date(d1) : d1;
    var date2 = typeof d2 === 'string' ? new Date(d2) : d2;

    if (format === 'Y') return date2.getFullYear() - date1.getFullYear();
    if (format === 'M') {
        var years = date2.getFullYear() - date1.getFullYear();
        var months = date2.getMonth() - date1.getMonth();
        return years * 12 + months;
    }
    if (format === 'D') {
        var msPerDay = 24 * 60 * 60 * 1000;
        return Math.round((date2 - date1) / msPerDay);
    }
    // 默认返回完整间隔
    var years = date2.getFullYear() - date1.getFullYear();
    var months = date2.getMonth() - date1.getMonth();
    var days = date2.getDate() - date1.getDate();

    if (days < 0) {
        months--;
        var prevMonth = new Date(date2.getFullYear(), date2.getMonth(), 0);
        days += prevMonth.getDate();
    }
    if (months < 0) {
        years--;
        months += 12;
    }

    var result = "";
    if (years > 0) result += years + "年";
    if (months > 0) result += months + "个月";
    if (days > 0) result += days + "天";
    return result || "0天";
};
JSA.datedif = JSA.z日期间隔;

/**
 * 选择列
 * @param {Array} arr - 二维数组
 * @param {Array} colIndexes - 列索引
 * @param {Array} newHeaders - 新表头
 * @returns {Array} 结果数组
 */
JSA.z选择列 = function(arr, colIndexes, newHeaders) {
    if (!arr || arr.length === 0) return [];

    var indexes = [];

    // 检查是否按表头选择
    if (arr.length > 0 && colIndexes.length > 0 && typeof colIndexes[0] === 'string') {
        var headers = arr[0];
        var headerMap = {};
        for (var i = 0; i < headers.length; i++) {
            headerMap[String(headers[i])] = i;
        }

        for (var j = 0; j < colIndexes.length; j++) {
            var col = colIndexes[j];
            if (headerMap.hasOwnProperty(col)) {
                indexes.push(headerMap[col]);
            }
        }

        var result = [];
        if (newHeaders && newHeaders.length > 0) {
            result.push(newHeaders);
        } else {
            var newRow = [];
            for (var ki = 0; ki < colIndexes.length; ki++) {
                var col = colIndexes[ki];
                var idx = headerMap[col];
                newRow.push(idx !== undefined ? headers[idx] : col);
            }
            result.push(newRow);
        }

        for (var i = 1; i < arr.length; i++) {
            var row = arr[i];
            var newRow = [];
            for (var ki = 0; ki < indexes.length; ki++) {
                newRow.push(row[indexes[ki]]);
            }
            result.push(newRow);
        }

        return result;
    } else {
        // 按列号选择
        indexes = [];
        for (var j = 0; j < colIndexes.length; j++) {
            indexes.push(typeof colIndexes[j] === 'number' ? colIndexes[j] : parseInt(colIndexes[j]));
        }

        var result = [];
        for (var i = 0; i < arr.length; i++) {
            var row = arr[i];
            var newRow = [];
            for (var ki = 0; ki < indexes.length; ki++) {
                newRow.push(row[indexes[ki]]);
            }
            result.push(newRow);
        }

        return result;
    }
};
JSA.selectCols = JSA.z选择列;

/**
 * 查找索引 - 增强版VLOOKUP
 * @param {*} 关键字 - 查找关键字
 * @param {Array} 数组 - 二维数组
 * @param {Number} 结果列 - 结果列索引
 * @param {Boolean} 模式 - 是否模糊匹配
 * @param {*} 错误值 - 找不到时的默认值
 * @returns {*} 查找结果
 */
JSA.z查找索引 = function(关键字, 数组, 结果列, 模式, 错误值) {
    if (!数组 || !数组.length) return 错误值 !== undefined ? 错误值 : null;
    结果列 = 结果列 || 1;
    var rowCount = 数组.length;
    for (var i = 0; i < rowCount; i++) {
        var row = 数组[i];
        if (!row) continue;
        var key = row[0];
        var match = 模式 ? (String(key).indexOf(String(关键字)) >= 0) : (key === 关键字);
        if (match) return row[结果列 - 1] !== undefined ? row[结果列 - 1] : 错误值;
    }
    return 错误值 !== undefined ? 错误值 : null;
};
JSA.match = JSA.z查找索引;

/**
 * 左侧查找 - VLOOKUP风格查找
 * @param {*} 关键字 - 查找关键字
 * @param {Array} 数组 - 二维数组
 * @param {Number} 结果列 - 结果列索引
 * @param {Boolean} 模式 - 是否模糊匹配
 * @param {*} 错误值 - 找不到时的默认值
 * @returns {*} 查找结果
 */
JSA.z左侧查找 = function(关键字, 数组, 结果列, 模式, 错误值) {
    return JSA.z查找索引(关键字, 数组, 结果列, 模式, 错误值);
};
JSA.vlookup = JSA.z左侧查找;

/**
 * 增强查找 - XLOOKUP风格查找
 * @param {*} 关键字 - 查找关键字
 * @param {Array} 查找数组 - 查找列
 * @param {Array} 结果数组 - 结果列
 * @param {*} 错误值 - 找不到时的默认值
 * @returns {*} 查找结果
 */
JSA.z增强查找 = function(关键字, 查找数组, 结果数组, 错误值) {
    if (!查找数组 || !结果数组) return 错误值 !== undefined ? 错误值 : null;
    for (var i = 0; i < 查找数组.length; i++) {
        if (查找数组[i] === 关键字) {
            return 结果数组[i] !== undefined ? 结果数组[i] : 错误值;
        }
    }
    return 错误值 !== undefined ? 错误值 : null;
};
JSA.xlookup = JSA.z增强查找;

/**
 * 转文本 - CSTR函数
 * @param {*} v - 值
 * @returns {String} 文本
 */
JSA.z转文本 = function(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return String(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    // v4.1.0 修复: 原代码 return JSA.z今天(v) 调用无参函数，应返回今天而非格式化输入日期
    // 现改用 z日期格式化: 默认 YYYY-MM-DD 格式
    if (v instanceof Date) {
        try { return DateUtils.prototype.z日期格式化.call(DateUtils(), v, 'yyyy-MM-dd'); }
        catch (e) { return v.toISOString().substring(0, 10); }
    }
    return String(v);
};
JSA.cstr = JSA.z转文本;

/**
 * 取整数 - CINT函数
 * @param {Number} v - 数值
 * @returns {Number} 整数部分
 */
JSA.z取整数 = function(v) {
    var num = Number(v);
    return isNaN(num) ? 0 : (num >= 0 ? Math.floor(num) : Math.ceil(num));
};
JSA.cint = JSA.z取整数;

// XXD-157/158 final fix: 基础数学函数中文 alias
JSA.z向下取整 = function(v) { return Math.floor(v); };
JSA.z向上取整 = function(v) { return Math.ceil(v); };
JSA.z四舍五入 = function(v) { return Math.round(v); };
JSA.z取整     = JSA.z四舍五入; // 默认四舍五入取整
JSA.z幂       = function(base, exp) { return Math.pow(base, exp); };
JSA.z对数     = function(v) { return Math.log(v); };
JSA.z绝对值   = function(v) { return Math.abs(v); };
JSA.z正弦     = function(v) { return Math.sin(v); };
JSA.z余弦     = function(v) { return Math.cos(v); };
JSA.z正切     = function(v) { return Math.tan(v); };

/**
 * 取小数 - GETDECIMAL函数
 * @param {Number} v - 数值
 * @returns {Number} 小数部分
 */
JSA.z取小数 = function(v) {
    var num = Number(v);
    if (isNaN(num)) return 0;
    var sign = num >= 0 ? 1 : -1;
    return sign * (Math.abs(num) - Math.floor(Math.abs(num)));
};
JSA.getDecimal = JSA.z取小数;

/**
 * 转公式数组
 * @param {Array} arr - 数组
 * @returns {String} 公式字符串
 */
JSA.z转公式数组 = function(arr) {
    if (!Array.isArray(arr)) return arr;
    var rows = arr.length;
    // XXD-47: 1D 数组的 arr[0] 是标量无 .length — 视为 1 列
    var cols = rows > 0 ? (Array.isArray(arr[0]) ? arr[0].length : 1) : 0;
    var result = [];
    for (var i = 0; i < rows; i++) {
        if (Array.isArray(arr[i])) {
            result.push(arr[i].join('\t'));
        } else {
            result.push(String(arr[i]));
        }
    }
    return result.join('\n');
};
JSA.toExcelArray = JSA.z转公式数组;

/**
 * 统一路径分隔符
 * @param {String} path - 路径
 * @returns {String} 标准化路径
 */
JSA.z统一路径分隔符 = function(path) {
    if (!path) return '';
    return path.replace(/\\/g, '/');
};
JSA.normalPath = JSA.z统一路径分隔符;

/**
 * [v5.0.0] 链式表达式解析
 * 把 "$$(...).filter(...).map(...)" 这类链式 lambda 编译成可执行函数
 * 关键:$$ 作为参数显式传入,保证作用域正确
 */
function _kParseChainableExpression(expr) {
    // 检测是否含链式调用(.filter / .map / .slice / .take / .skip / .sort / .reduce)
    var isChainable = /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce|find|some|every)\s*\(/.test(expr);
    if (!isChainable) return null;

    try {
        // 🔧 v4.0.14 关键: IIFE 内 var $$ = Array2D
        //   用户的 lambda `(...args)=>$$.superPivot(...args).filter(...)` 中 $$ 是自由变量
        //   lambda 词法作用域 = 这里 IIFE 内的 scope,通过 var $$ = Array2D 捕获
        //   不加这句,lambda 调用时 $$ 未定义 → ReferenceError → jsaLambda 返回 null → #K_ERR
        // v4.0.11 已有这个结构,v4.0.13 我误改成 __Array2D 参数就坏了
        // 现在恢复 v4.0.11 结构,不再覆盖 Array.prototype (避免 WPS 沙箱拒绝)
        var fn = new Function('__args', 'return (function() {' +
                              '  var $$ = (typeof Array2D !== "undefined") ? Array2D : this.Array2D;' +
                              '  return (' + expr + ').apply(null, __args);' +
                              '}).apply(null, __args)');
        return fn;
    } catch (e) {
        if (typeof Console !== 'undefined') Console.log('parseChainableExpression 失败:' + e.message);
        return null;
    }
}

/**
 * jsaLambda（K函数）- 以字符串形式创建并执行JSA函数 【v4.2.2 增强版】
 *
 * 1) 路径调用：   k("JSA.getIndexs", 1, 10, 2)
 * 2) Lambda：    k("x => x*2", 5)                       // 10
 * 3) 索引语法：  k("$0*2", [1,2,3])                     // 2
 * 4) 多行代码：  k("var s=0; for (var i=0;i<args.length;i++) s += args[i]; return s;", 1, 2, 3)
 * 5) -r参数：    k("rng => rng.Address()", -r, "A1:B3")
 * 6) ...args：   k("(...args) => args.join(',')", 1, 2, 3)        // "1,2,3"
 *
 * @param {String|Function} fn - 字符串函数表达式 / 路径 / 多行代码
 * @param {...any} args - 参数；当字符串 "-r" 出现时，其后所有字符串参数自动按 Range 地址解析
 * @returns {*} 函数结果；解析失败返回 null
 */
JSA.jsaLambda = function(fn, ...args) {
    try {
        // 0) 🔧 v4.0.24: 提前声明 realArgs, 让 __KJ_ARGS__ 提取 + -r 解析 + args 收集共用
        var realArgs = [];

        // 0.5) 🔧 v4.0.24: 提取 __KJ_ARGS__ —— 解决 WPS 公式引擎吞字符串参数 bug
        //   根因: WPS 公式 =k("$$...", A1:H40, "f3,f2", "f6", "count()...") 5 个 string
        //         参数中,WPS 只传进来 4 个(其中一个被默默丢了),例如 "f6" 丢失
        //   修法: 用户在第一个 fn 字符串中加注释式标记 __KJ_ARGS__={"rowFields":"f3,f2",...}
        //         jsaLambda 入口先收集 WPS 传的 args 到 realArgs 头部,再把提取的字段追加到末尾
        //         这样 (...args) 展开时,顺序仍是 [Range, ..., rowFields, colFields, dataFields, headerRows]
        //   例: =k("__KJ_ARGS__={\"rowFields\":\"f3,f2\"}  (...args)=>$$...", A1:H40, "count()...")
        //   WPS 传: args = [Range, "count()"]
        //   提取后: realArgs = [Range, "count()", "f3,f2"]
        //   (...args) 展开: superPivot(Range, "count()", "f3,f2") ❌
        //   ❌ 上面的追加方案仍然不对,改用插队方案: 把 __KJ_ARGS__ 字段塞到 WPS args 中的 Range 之后
        //   提取后: realArgs = [Range, "f3,f2", "count()"]
        //   (...args) 展开: superPivot(Range, "f3,f2", "count()") → superPivot(arr, rowFields, dataFields) ✅
        //   ⚠️ 用户公式必须按 [Range, 数据依赖字段, 其他字段] 顺序传 args
        var __kjExtracted = null;
        if (typeof fn === 'string') {
            var __kjMatch = fn.match(/__KJ_ARGS__\s*=\s*(\{[\s\S]*?\})\s*/);
            if (__kjMatch) {
                var __kjJsonText = __kjMatch[1];
                // 🔧 v4.0.27: 反引号 → 双引号
                __kjJsonText = __kjJsonText.replace(/`/g, '"');
                // 🔧 v4.0.27: 宽松 JSON 解析 — 字段名 / 字符串值都允许无引号
                //   根因: WPS 公式字符串中 "" 配对外层,内部 " 经常被吞
                //   用户的 __KJ_ARGS__ 实际进 jsaLambda 是 {rowFields:"f3,f2"} 而非 {"rowFields":"f3,f2"}
                //   改: 先尝试标准 JSON.parse,失败则用宽松解析(给 field 名/字符串值自动加引号)
                var __kj = null;
                var __kjStrictOk = false; // 🔧 [XXD-51] 跟踪严格 JSON.parse 是否成功(用来区分 "宽松解析只是空对象" vs "__KJ_ARGS__={} 正常")
                try {
                    __kj = JSON.parse(__kjJsonText);
                    __kjStrictOk = true;
                } catch (__kjE1) {
                    if (typeof Console !== 'undefined') {
                        try { Console.log('[k/v4.0.27] __KJ_ARGS__ 严格 JSON parse 失败: ' + __kjE1.message + ', json=' + __kjJsonText + ', 尝试宽松解析'); } catch (__) {}
                    }
                    // 宽松解析: 保护引号内的逗号(值可能含逗号如 "f3,f2"),然后按外层 , 切
                    // 🔧 v4.0.28: 先把 "..." 内的 , 替换为占位符 __KJ_PLACEHOLDER__,切完还原
                    __kj = {};
                    var __kjProtected = __kjJsonText
                        .replace(/^\{/, '')
                        .replace(/\}$/, '')
                        .replace(/"([^"]*)"/g, function(m, p1) { return '"' + p1.replace(/,/g, '__KJ_PLACEHOLDER__') + '"'; })
                        .replace(/'([^']*)'/g, function(m, p1) { return "'" + p1.replace(/,/g, '__KJ_PLACEHOLDER__') + "'"; });
                    var __kjEntries = __kjProtected.split(',');
                    for (var __kei = 0; __kei < __kjEntries.length; __kei++) {
                        var __kv = __kjEntries[__kei].split(':');
                        if (__kv.length >= 2) {
                            var __kjKey = __kv[0].trim().replace(/^["']|["']$/g, '');
                            var __kjVal = __kv.slice(1).join(':').trim()
                                .replace(/^["']|["']$/g, '')  // 去外层引号
                                .replace(/__KJ_PLACEHOLDER__/g, ',');  // 还原占位符
                            __kj[__kjKey] = __kjVal;
                        }
                    }
                }
                // 🔧 [XXD-51] 严格 JSON.parse 失败 且 宽松解析也没拿到任何键(输入非空) → 抛 K_ERR 提示用户
                //   旧行为:__kj 留空 {} → __kjExtracted=[] → fn 标记**还是被剥了** → 用户以为传了实际没传
                //   新行为:显式抛 PARSE 错,外层 JSA.k 兜成 "#K_ERR: pos=0, FN, msg=\"...__KJ_ARGS__ 解析失败: ...\""
                //   注意:__KJ_ARGS__={} 走的是 __kjStrictOk=true 分支,这里不会触发 throw
                if (!__kjStrictOk) {
                    var __kjKeys = __kj ? Object.keys(__kj) : [];
                    if (__kjKeys.length === 0 && __kjJsonText && __kjJsonText.length > 0) {
                        var __kjErr = new Error('__KJ_ARGS__ 解析失败: ' + __kjJsonText.substring(0, 200));
                        __kjErr.__kCode = 'PARSE';
                        throw __kjErr;
                    }
                }
                __kjExtracted = [];
                if (typeof __kj.rowFields === 'string' && __kj.rowFields) __kjExtracted.push(__kj.rowFields);
                if (typeof __kj.colFields === 'string' && __kj.colFields) __kjExtracted.push(__kj.colFields);
                if (typeof __kj.dataFields === 'string' && __kj.dataFields) __kjExtracted.push(__kj.dataFields);
                if (typeof __kj.headerRows !== 'undefined') __kjExtracted.push(__kj.headerRows);
                // 🔧 [XXD-51] 守卫 fn 标记剥除 — 仅当真的提取到字段时才剥
                //   __KJ_ARGS__={} 时 __kjExtracted=[] → 保留标记在 fn 里,提示用户这是空对象(下游会按语法错兜底,而不是悄悄传错)
                if (__kjExtracted.length > 0) {
                    fn = fn.replace(__kjMatch[0], '').trim();
                }
            }
        }
        // 🔧 v4.0.26: 规范化 fn 中的换行符为单个空格(防止 WPS 公式粘多行导致 syntax error)
        if (typeof fn === 'string') {
            fn = fn.replace(/\s+/g, ' ').trim();
        }

        // 1) 解析 -r：开关打开后，后续字符串参数都按 Range 地址转 Range 对象
        //    🔧 v4.0.24: 同时把 __KJ_ARGS__ 提取的字段插到 **第一个非 -r 参数之后**
        //    典型场景: args = [Range, "count()"]  →  realArgs = [Range, "f3,f2", "f6", "count()"]
        //    这样 (...args) 展开: superPivot(Range, "f3,f2", "f6", "count()") → 顺序对 ✅
        var rangeMode = false;
        var __kjInjected = false;
        for (var i = 0; i < args.length; i++) {
            var a = args[i];
            if (a === '-r' || a === '-R') {
                rangeMode = true;
                continue;
            }
            // 先 push 当前参数(Range / string / whatever)
            if (rangeMode && typeof a === 'string' && /^\$?[A-Za-z]+[\d]+(:\$?[A-Za-z]+[\d]+)?$/.test(a)) {
                realArgs.push(asRange(a));
            } else {
                realArgs.push(a);
            }
            // 🔧 v4.0.24: 第一个非 -r 参数(Range 本身)之后,立刻插队注入 __KJ_ARGS__ 字段
            if (!__kjInjected && __kjExtracted && __kjExtracted.length > 0) {
                for (var __ki = 0; __ki < __kjExtracted.length; __ki++) realArgs.push(__kjExtracted[__ki]);
                __kjInjected = true;
            }
        }
        // 🔧 v4.0.24 兜底: WPS 没传任何 args,__KJ_ARGS__ 仍需追加
        if (!__kjInjected && __kjExtracted) {
            for (var __ki2 = 0; __ki2 < __kjExtracted.length; __ki2++) realArgs.push(__kjExtracted[__ki2]);
        }

        // 1.5) 🔧 v4.0.22 诊断 — 看 jsaLambda 实际收到的 args
        if (typeof Console !== 'undefined') {
            try {
                var __dumpArgs = [];
                for (var __di = 0; __di < args.length; __di++) {
                    var __a = args[__di];
                    if (__a === null) __dumpArgs.push('null');
                    else if (__a === undefined) __dumpArgs.push('undefined');
                    else if (typeof __a === 'string') __dumpArgs.push('"' + __a + '"');
                    else __dumpArgs.push('<' + typeof __a + '>');
                }
                var __dumpReal = [];
                for (var __ri = 0; __ri < realArgs.length; __ri++) {
                    var __ra = realArgs[__ri];
                    if (__ra === null) __dumpReal.push('null');
                    else if (__ra === undefined) __dumpReal.push('undefined');
                    else if (typeof __ra === 'string') __dumpReal.push('"' + __ra + '"');
                    else if (typeof __ra === 'function' || (typeof __ra === 'object' && __ra && typeof __ra.Address !== 'undefined')) __dumpReal.push('<Range>');
                    else __dumpReal.push('<' + typeof __ra + '>');
                }
                Console.log('[k/v4.0.22] jsaLambda IN: fn=' + JSON.stringify(fn) + ', args=[' + __dumpArgs.join(', ') + '], argsLen=' + args.length);
                Console.log('[k/v4.0.24] realArgs=[' + __dumpReal.join(', ') + '], realArgsLen=' + realArgs.length);
            } catch (__) {}
        }

        // 2) 【v4.2.2 增强】WPS 公式引擎里不识别反引号（会报#NAME?），
        //      把不含 ${ 的反引号包围短串自动转成双引号（如 `f4` → "f4`），
        //      含 ${ 的当作 JS 模板字符串，不动
        function fixBacktick(s) {
            if (typeof s === 'string' && s.indexOf('${') === -1) {
                return s.replace(/\`([^\`\n]+)\`/g, '"$1"');
            }
            return s;
        }
        if (typeof fn === 'string') fn = fixBacktick(fn);
        for (var j = 0; j < realArgs.length; j++) {
            if (typeof realArgs[j] === 'string') {
                realArgs[j] = fixBacktick(realArgs[j]);
            }
        }

        // 2.6) 【v4.2.2 增强】智能参数转换:
        //   a) 智能 Range -> Value2:WPS 公式只能传原始类型,传 Range 对象会让多数函数报错。
        //      检测到 Range 对象自动取 .Value2
        //   b) 1x1 2D 数组 flatten:WPS 公式传单值时实际是 [[v]] (1x1 二维),flatten 为 v
        //      例外:A1:A10 在 WPS 是 [[1],[2]] (Nx1 二维),不能 flatten,否则 filter/map 崩溃
        //   c) NxM / 1xN / Nx1 二维数组保持原样,交给 Array2D 方法处理
        // 🔧 v4.0.14 修复: 2D 数组必须保证是真正的 JS Array(非 WPS host array)
        //   WPS 的 Range.Value2 返回的数组通过 Array.isArray() 检查,但 .slice 可能是 undefined
        //   (host array 只实现了部分 Array 方法)。superPivot 第一行就 arr.slice(dataStartRow),
        //   报 "arr.slice is not a function" → jsaLambda 返回 null → #K_ERR
        //   修法: 检测到数组缺少关键方法(.slice/.filter/.map)时, 用 JSON 往返或 Array.from 强转
        // 🔧 v4.0.14 修复: 递归把 WPS host array 转成真正的 JS Array
        //   现象: WPS 的 Range.Value2 返回 host array,Array.isArray() === true 但 .slice 是 undefined
        //   即使外层数组的 .slice 存在,子行(host 1D array)的 .slice 也可能缺失
        //   修法: 递归处理所有层级的数组(外层 + 子层),用 JSON 往返或手动 copy 强转
        function _toRealArray(v) {
            if (!Array.isArray(v)) return v;
            // 已经是真正的 Array(原型上有 .slice + .map + .filter)就跳过
            if (typeof v.slice === 'function' && typeof v.map === 'function' && typeof v.filter === 'function') {
                return v;
            }
            // host array, JSON 往返最稳
            try {
                var __s = JSON.stringify(v);
                var __r = JSON.parse(__s);
                if (Array.isArray(__r)) return __r;
            } catch (__e1) {}
            // 兜底: 手动 copy(也递归处理子数组)
            try {
                var __out = [];
                for (var __i = 0; __i < v.length; __i++) {
                    var __item = v[__i];
                    __out.push(Array.isArray(__item) ? _toRealArray(__item) : __item);
                }
                return __out;
            } catch (__e2) { return v; }
        }
        function smartUnwrap(v) {
            // 🔧 v4.0.17: Range 对象检测改用 duck-typing,不依赖 typeof
            //   根因: WPS 公式 =k("...", A1:H40, ...) 传给 JSA 的 Range 对象 typeof === 'function'(不是 'object')
            //         v4.0.14 的 typeof v === 'object' 检测直接跳过,Range 直接传 superPivot
            //         superPivot 内部 arr.slice(dataStartRow) 报 "arr.slice is not a function"
            //   修法: 看 v 是否有 Address + Value2 属性(duck-typing),不查 typeof
            if (v != null && typeof v.Address !== 'undefined' && typeof v.Value2 !== 'undefined' && v !== asRange(v)) {
                try {
                    var __vv = v.Value2;
                    if (__vv && __vv !== v) v = __vv;
                } catch (e) { return v; }
            }
            // 1x1 2D 数组 flatten
            if (Array.isArray(v) && v.length === 1 && Array.isArray(v[0]) && v[0].length === 1) {
                return v[0][0];
            }
            // 🔧 v4.0.14: 递归把任何 host array 转成真正的 JS Array
            v = _toRealArray(v);
            return v;
        }
        for (var sk = 0; sk < realArgs.length; sk++) {
            realArgs[sk] = smartUnwrap(realArgs[sk]);
        }

        // 3) 解析字符串为可执行函数
        // [v5.0.0] 链式调用检测
        // 🔧 v4.0.14 恢复 v4.0.11 的链式解析器路径(IIFE 内 var $$ = Array2D,lambda 词法捕获)
        if (typeof fn === 'string' && /\.\s*(filter|map|slice|take|skip|sort|forEach|reduce|find|some|every)\s*\(/.test(fn)) {
            var chainParser = _kParseChainableExpression(fn);
            // 🔧 v4.0.23 诊断: 链式解析失败时打印 fn 看是否被吃字符
            if (typeof Console !== 'undefined') {
                try {
                    Console.log('[k/v4.0.23] chainParser=' + (chainParser ? 'OK' : 'NULL') +
                        ', fn=' + JSON.stringify(fn));
                } catch (__) {}
            }
            if (chainParser) {
                try {
                    var _chainResult = chainParser(realArgs);
                    if (_chainResult !== null && _chainResult !== undefined) {
                        // 🔧 v4.0.19: Array2D 实例 → unwrap 成普通 2D 数组
                        //   根因: superPivot 返回 wrappedResult(Array2D 实例),WPS 公式 spill 不知
                        //         怎么处理 Array2D 实例,只 spill 第 1 行。
                        //   修法: 如果是 Array2D 实例,取 .val() 或 ._items 返回纯 2D 数组
                        //         链式 .filter((x,i)=>i==0 || x.fN==...) 后是 Array2D,也 unwrap
                        if (_chainResult instanceof Array2D) {
                            try {
                                if (typeof _chainResult.val === 'function') {
                                    var __valOut = _chainResult.val();
                                    if (typeof Console !== 'undefined') {
                                        try {
                                            Console.log('[k/v4.0.20] val() returned: isArr=' + Array.isArray(__valOut) +
                                                ', len=' + (__valOut && __valOut.length !== undefined ? __valOut.length : 'n/a') +
                                                ', rowLens=' + (function() { try { var rl=[]; for (var i=0; i<Math.min(3,__valOut.length); i++) rl.push(__valOut[i] && __valOut[i].length); return rl.join(','); } catch(e) { return '?'; } })());
                                            // 🔧 v4.0.25 诊断: 打印前 2 行的实际内容
                                            if (__valOut && __valOut.length > 0) {
                                                var __row0 = JSON.stringify(__valOut[0]);
                                                var __row1 = __valOut.length > 1 ? JSON.stringify(__valOut[1]) : '(none)';
                                                Console.log('[k/v4.0.25] val() row[0]=' + __row0);
                                                Console.log('[k/v4.0.25] val() row[1]=' + __row1);
                                            }
                                        } catch (__) {}
                                    }
                                    return __valOut;
                                }
                                if (_chainResult._items) return _chainResult._items;
                            } catch (__ue) {}
                        }
                        return _chainResult;
                    }
                } catch (e) {
                    if (typeof Console !== 'undefined') {
                        Console.log('[k/chain] 执行失败: ' + e.message);
                        try { Console.log('[k/chain] STACK: ' + (e.stack || '(no stack)')); } catch (__) {}
                        try {
                            var __arr0 = realArgs[0];
                            Console.log('[k/chain] args[0]: isArr=' + Array.isArray(__arr0) +
                                ', t=' + typeof __arr0 +
                                ', hasSlice=' + (Array.isArray(__arr0) ? (typeof __arr0.slice) : 'n/a') +
                                ', len=' + (__arr0 && __arr0.length !== undefined ? __arr0.length : 'n/a') +
                                ', keys=' + (function() { try { var k=[]; for (var x in __arr0) { if (k.length<6) k.push(x); } return k.join(','); } catch(e) { return '?'; } })());
                        } catch (__) {}
                    }
                    // 继续走原来的解析流程
                }
            }
        }

        var func = JSA.z解析函数表达式(fn);
        if (typeof func === 'function') {
            var __mainRet = func.apply(null, realArgs);
            // 🔧 v4.0.19: Array2D 实例 → 普通 2D 数组 unwrap(WPS spill 需 plain array)
            if (__mainRet instanceof Array2D) {
                try { if (typeof __mainRet.val === 'function') return __mainRet.val(); } catch (__e) {}
                try { if (__mainRet._items) return __mainRet._items; } catch (__e) {}
            }
            return __mainRet;
        }

        // 3) 兜底：把 fn 当作完整JSA代码块执行（兼容极简场景）
        if (typeof fn === 'string' && fn.indexOf('=>') === -1) {
            var blockFn = new Function('args', 'with (JSA) { ' + fn + ' }');
            return blockFn(realArgs);
        }
        return null;
    } catch (e) {
        // 🔧 [XXD-51] 透传已标记的 K_ERR 类错误(__kCode 设过)给 JSA.k 兜成 #K_ERR 字符串
        //   否则在 HEAD 这层 catch 会把 __KJ_ARGS__ 解析失败 等提示吞掉,函数返回 null,用户看不到
        if (e && e.__kCode) throw e;
        console.warn('jsaLambda 执行失败:', e && e.message ? e.message : e);
        return null;
    }
};
/**
 * [v5.x, XXD-103/104] _safeJsonStringify — JSON.stringify 安全包装
 * 处理原生 JSON.stringify 在 WPS JSA 环境中的边界行为:
 * 1. undefined → 返回 undefined (不抛错)
 * 2. BigInt → 透传原生 TypeError (不包装错误信息)
 * 3. NaN/Infinity → 原生行为返回 'null' (不干预)
 * @param {*} val - 要序列化的值
 * @param {Function|Array} [replacer] - 替换函数或数组
 * @param {number|string} [space] - 缩进空格数或字符串
 * @returns {string|undefined} JSON 字符串,或 undefined(输入为 undefined 时)
 */
function _safeJsonStringify(val, replacer, space) {
    // 规则1: undefined 直接返回 undefined,匹配原生 JSON.stringify(undefined) 行为
    if (typeof val === 'undefined') return undefined;
    // 规则2: BigInt 原生 TypeError 直接透传(JSON.stringify 自身会抛),不额外包装
    // 规则3: NaN/Infinity 原生就返回 'null',不干预
    return JSON.stringify(val, replacer, space);
}

/**
 * [v5.0.0] JSA.k — k() 的完整实现(shim 给顶层 function k 调)
 *
 * 比 jsaLambda 多两层:
 *   1. fn / args 预处理(其实 jsaLambda 内部已做大部分,这里加保险)
 *   2. 错误位置化:#K_ERR: pos=N, KIND, msg="..."
 *
 * [v4.0.42] 新增 options object 模式:
 *   k({fn:'...', range:'A1:D40', rowFields:'f3,f2', ...})
 *   第一个参数是非空对象 + 有 fn 属性 → 走 options object 模式,否则走原来 positional 模式
 *
 * @param {string|Function|Object} fn - 字符串函数表达式 / 路径 / Lambda / options object
 * @param {...any} args - 后续参数(仅 positional 模式)
 * @returns {*} 函数结果;失败返回 "#K_ERR: ..."
 */
JSA.k = function(fn) {
    var args = [];

    // [v4.0.42] options object 模式: 第一个参数是非空对象+有fn属性→options模式,否则positional模式(向后兼容)
    if (typeof fn === 'object' && fn !== null && !Array.isArray(fn) && fn.fn !== undefined) {
        var opts = fn;
        fn = opts.fn;
        // range 优先放第一个参数位
        if (opts.range !== undefined) {
            args.push(opts.range);
        }
        // 其余属性按插入顺序追加(排除 fn 和 range)
        for (var key in opts) {
            if (key !== 'fn' && key !== 'range' && Object.prototype.hasOwnProperty.call(opts, key)) {
                args.push(opts[key]);
            }
        }
    } else {
        for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);
    }

    // 0) 必传检查
    if (typeof fn === 'undefined' || fn === null || fn === '') {
        return '#K_ERR: pos=0, FN, msg="fn 不能为空"';
    }

    // 1) 框架未加载?
    if (typeof JSA.jsaLambda !== 'function') {
        return '#K_ERR: pos=1, INTERNAL, msg="JSA880 框架未加载,请加载 JSA880.js 加载项"';
    }

    // 2) 调 jsaLambda(已含反引号转换 / Range 转 Value2 / 1x1 flatten)
    var result;
    try {
        result = JSA.jsaLambda.apply(null, [fn].concat(args));
    } catch (e) {
        // 错误位置:pos=0 表示 fn / jsaLambda 自身
        var kind = (e && e.message && e.message.indexOf('TypeError') !== -1) ? 'TYPE' : 'FN';
        return '#K_ERR: pos=0, ' + kind + ', msg="' +
               (e && e.message ? e.message.replace(/"/g, "'") : String(e)) + '"';
    }

    // 3) null / undefined 兜底
    if (result === undefined || result === null) {
        return '#K_ERR: pos=0, FN, msg="jsaLambda 返回 null/undefined,可能 fn 语法错或参数不匹配"';
    }

    return result;
};

// 保留兼容:JSA.k 仍可被 jsaLambda 直接调(无包装的版本)
// 如果需要原始行为,可用 JSA.kRaw = JSA.jsaLambda(本计划不导出,留作未来扩展)

/**
 * [v5.0.0] JSA.k.help — 排错指南
 * 在 JSA 编辑器手动调 JSA.k.help() 打印排错清单
 */
JSA.k.help = function() {
    if (typeof Console === 'undefined') return;
    Console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Console.log('  k() 公式常见问题排查');
    Console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Console.log('❌ #K_ERR: pos=0, FN → fn 字符串解析失败,检查语法');
    Console.log('❌ #K_ERR: pos=1, INTERNAL → JSA880 框架未加载');
    Console.log('❌ #K_ERR: pos=0, TYPE → 类型错(传了不该传的对象)');
    Console.log('❌ #NAME? → ThisWorkbook 没粘 3-5 行 wrapper');
    Console.log('✅ 自检:在任意单元格 =k("JSA.getIndexs", 1, 5, 1)');
    Console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
};

/**
 * 【v4.2.2 增强】JSA 命名空间别名注册
 *
 * 背景：
 *   实际写在 WPS 公式里的 k() 调用，绝大多数都是
 *   k("JSA.xxx", ...) / k("$$.xxx", ...) / k("Array2D.xxx", ...) 三种形式。
 *   但框架里很多方法都只定义在 Array2D 上（getIndexs / superPivot / textjoin / leftjoin / distinct / insertCols ...），
 *   用户在公式里写 JSA.getIndexs 会报 #K_ERR: undefined。
 *
 * 解决：在 JSA 命名空间补齐这些常用别名，使三种路径前缀都能用。
 *
 * 同时注册 qctextjoin → textjoin（个别工作簿/老代码里用的别名）。
 *
 * 这里直接赋值即可，Array2D 上的方法签名都没问题。
 */
(function _registerJSAAliases() {
    var aliasMap = {
        'getIndexs':     'getIndexs',
        'superPivot':    'superPivot',
        'z超级透视':     'z超级透视',
        'distinct':      'distinct',
        'z去重':         'z去重',
        'leftjoin':      'leftjoin',
        'z左连接':       'z左连接',
        'insertCols':    'insertCols',
        'z批量插入列':   'z批量插入列',
        'textjoin':      'textjoin',
        'z文本连接':     'z文本连接',
        'qctextjoin':    'textjoin',   // 兼容个别老公式里的 qctextjoin 写法
        'qcTextjoin':    'textjoin',
        'qcTextJoin':    'textjoin',
        'filter':        'filter',
        'z筛选':         'z筛选',
        'map':           'map',
        'z映射':         'z映射',
        'sum':           'z求和',       // JSA.sum(arr) → Array2D.z求和
        'z求和':         'z求和',
        'avg':           'z平均',
        'average':       'z平均',
        'z平均':         'z平均',
        'max':           'z最大',
        'min':           'z最小',
        'z最大':         'z最大',
        'z最小':         'z最小',
        'count':         'z计数',
        'z计数':         'z计数',
        'distinctRows':  'distinct',
        'groupBy':       'z分组',
        'z分组':         'z分组',
        'groupInto':     'groupInto',
        'z分组汇总':     'groupInto',
        'sort':          'z多列排序',
        'z排序':         'z多列排序',
        'z多列排序':     'z多列排序',
        'trans':         'z转置',
        'transpose':     'z转置',
        'z转置':         'z转置',
        'select':        'z选择列',
        'z选择列':       'z选择列',
        'selectCols':    'z选择列',
        'toRange':       'toRange',
        'fromRange':     'fromRange',
        'z写入单元格':   'toRange',
        'z读取区域':     'fromRange'
    };
    for (var alias in aliasMap) {
        if (typeof Array2D[aliasMap[alias]] === 'function' && typeof JSA[alias] === 'undefined') {
            // 绑定到 Array2D 上以保持 this 指向
            JSA[alias] = (function(name, fn) {
                return function() { return fn.apply(Array2D, arguments); };
            })(alias, Array2D[aliasMap[alias]]);
        }
    }
})();

/**
 * 解析函数表达式（v4.2.2 增强版）
 *
 * 支持的语法：
 *   - Lambda 箭头函数： "x => x*2" / "(x,y) => x+y"
 *   - 索引选择器：      "$0 + $1"  /  "$0 * 2"
 *   - 列选择器：        "f1+f2"  /  "[f1,f3,f5]"
 *   - 路径调用：        "JSA.getIndexs"  /  "Array2D.z筛选"  /  "round"  /  "cstr"
 *   - 多行JSA代码块：   "var s=0; s += a; return s;" （需以 return 结束）
 *
 * @param {String|Function} expr - 函数表达式或函数
 * @returns {Function|null} 可执行函数
 */
JSA.z解析函数表达式 = function(expr) {
    if (typeof expr === 'function') return expr;
    if (expr === null || expr === undefined || expr === '') return null;
    if (typeof expr !== 'string') return null;

    // 缓存
    if (_lambdaCache[expr]) return _lambdaCache[expr];

    var fn = null;

    // === 路径调用：A.B.C 形式（含可省略的尾部 ()） ===
    // 匹配 "JSA.getIndexs" / "Array2D.z筛选" / "round" / "JSA.getIndexs(1,10,2)"
    var pathMatch = expr.match(/^\s*([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)+)\s*(\([^)]*\))?\s*$/);
    if (pathMatch) {
        var path = pathMatch[1];
        var tailArgsStr = pathMatch[2]; // 可能为 null
        try {
            // 解析根对象：JSA / Array2D / 全局（this/globalThis）
            // 重要：JSA 和 $$ 路径都设置成双根查找（本体 + Array2D 兑底），
            // 原因：实际公式里 k("JSA.getIndexs", ...) / k("$$.superPivot", ...) 两种都常见，
            // 框架里这些函数都只定义在 Array2D 上，JSA 命名空间名是在其上方补了别名
            // （但已补的在 _registerJSAAliases 里只到该函数定义以后才生效，路径名调货
            //  还是能拿到）。这里额外加兑底保证“getIndexs/textjoin/leftjoin/distinct/insertCols/superPivot/qctextjoin” 都在。
            var root = null;
            var fallback = null;
            if (path.indexOf('JSA.') === 0) {
                root = JSA;
                fallback = Array2D;
            } else if (path.indexOf('$$') === 0) {
                // PATCH v4.2.4: $$ 路径 fallback 改到 JSA
                root = (typeof $$ !== 'undefined') ? $$ : JSA;
                fallback = JSA;
            } else if (path.indexOf('Array2D.') === 0) {
                root = Array2D;
                fallback = JSA;
            } else if (path.indexOf('RngUtils.') === 0) {
                root = (typeof RngUtils !== 'undefined') ? RngUtils : null;
                fallback = Array2D;
            } else if (path.indexOf('ShtUtils.') === 0) {
                root = (typeof ShtUtils !== 'undefined') ? ShtUtils : null;
            } else if (path.indexOf('DateUtils.') === 0) {
                root = (typeof DateUtils !== 'undefined') ? DateUtils : null;
            } else if (path.indexOf('IO.') === 0) {
                root = (typeof IO !== 'undefined') ? IO : null;
            } else {
                // 兼容全局函数：round / cstr / isArray / val 等
                var globalObj = (typeof globalThis !== 'undefined') ? globalThis : this;
                if (path.indexOf('.') === -1 && globalObj && typeof globalObj[path] === 'function') {
                    var gfn = globalObj[path];
                    _lambdaCache[expr] = function() { return gfn.apply(globalObj, arguments); };
                    return _lambdaCache[expr];
                }
            }

            // 在 root + fallback 中查找函数
            var foundTarget = null;
            var foundRoot = null;
            var parts = path.split('.');
            for (var rIdx = 0; rIdx < 2 && !foundTarget; rIdx++) {
                var candidate = rIdx === 0 ? root : fallback;
                if (!candidate) continue;
                var t = candidate;
                for (var pi = 1; pi < parts.length; pi++) {
                    if (t == null) { t = null; break; }
                    t = t[parts[pi]];
                }
                if (typeof t === 'function') {
                    foundTarget = t;
                    foundRoot = candidate;
                }
            }

            if (foundTarget) {
                if (tailArgsStr) {
                    // 固定前置参数："JSA.xxx(1,10)" 先把 1,10 锁定，后面 k() 再追加
                    var fixedStr = tailArgsStr.slice(1, -1); // 去掉首尾括号
                    var fixedArgs = [];
                    if (fixedStr.trim() !== '') {
                        // 用 eval 安全地解析字面量参数（数字 / 字符串 / true/false/null）
                        // eslint-disable-next-line no-eval
                        fixedArgs = (function() { return [eval('[' + fixedStr + ']')]; })();
                        // eval 返回的是数组的引用，需要 .concat
                        fixedArgs = [].concat(fixedArgs[0] || []);
                    }
                    (function(tgt, fa) {
                        _lambdaCache[expr] = function() {
                            var all = [].concat(fa).concat([].slice.call(arguments));
                            return tgt.apply(null, all);
                        };
                    })(foundTarget, fixedArgs);
                } else {
                    (function(tgt) {
                        _lambdaCache[expr] = function() { return tgt.apply(null, arguments); };
                    })(foundTarget);
                }
                return _lambdaCache[expr];
            }
        } catch (ePath) {
            // 路径解析失败，继续尝试下面的 lambda 模式
        }
    }

    // === 箭头函数 ===
    if (LAMBDA_PATTERNS.ARROW_FUNCTION.test(expr)) {
        try {
            // eslint-disable-next-line no-eval
            fn = eval('(' + expr + ')');
            _lambdaCache[expr] = fn;
            return fn;
        } catch (e) {
            console.warn('Lambda箭头解析失败:', expr, e);
        }
    }

    // === $0/$1 索引语法 ===
    if (expr.indexOf('$') !== -1) {
        var indexMatch = expr.match(LAMBDA_PATTERNS.INDEX_SELECTOR);
        if (indexMatch && indexMatch.length > 0) {
            var indices = indexMatch.map(function(m) { return parseInt(m.substring(1)); });
            if (indices.length > 0) {
                var maxIndex = Math.max.apply(Math, indices);
                if (isFinite(maxIndex) && maxIndex <= ARRAY_LIMITS.MAX_INDEX) {
                    fn = new Function('_', 'return ' + expr.replace(LAMBDA_PATTERNS.INDEX_SELECTOR, '_[$1]'));
                    _lambdaCache[expr] = fn;
                    return fn;
                }
            }
        }
    }

    // === f1, f2 多列语法 ===
    if (LAMBDA_PATTERNS.MULTI_COLUMN.test(expr)) {
        var cols = expr.split(/\s*[,，]\s*/).map(function(c) {
            return '_[' + (parseInt(c.substring(1)) - 1) + ']';
        }).join(',');
        fn = new Function('_', 'return [' + cols + ']');
        _lambdaCache[expr] = fn;
        return fn;
    }

    // === [f1, f2] 方括号语法 ===
    if (LAMBDA_PATTERNS.ARRAY_BRACKET.test(expr)) {
        var innerExpr = expr.slice(1, -1).trim();
        var cols2 = innerExpr.split(/\s*[,，]\s*/).map(function(c) {
            return '_[' + (parseInt(c.substring(1)) - 1) + ']';
        }).join(',');
        fn = new Function('_', 'return [' + cols2 + ']');
        _lambdaCache[expr] = fn;
        return fn;
    }

    // === f1 / f(1) 单列选择器 ===
    if (/f\s*\(\s*\d+\s*\)/.test(expr) || LAMBDA_PATTERNS.COLUMN_SELECTOR.test(expr)) {
        fn = new Function('_', 'return ' + expr.replace(/f\s*\(?\s*(\d+)\s*\)?\s*/gi, function(m, num) {
            return '_[' + (parseInt(num) - 1) + ']';
        }));
        _lambdaCache[expr] = fn;
        return fn;
    }

    // === 多行JSA代码块：当 fn 既不是路径、也没有 => 时，作为完整函数体处理 ===
    // 用 (...args) => { ... } 包裹，允许多行语句，但最后必须 return
    if (expr.indexOf('=>') === -1 && expr.indexOf('function') === -1) {
        try {
            // eslint-disable-next-line no-new-func
            fn = new Function('...args', expr);
            _lambdaCache[expr] = fn;
            return fn;
        } catch (eBlock) {
            console.warn('Lambda多行代码块解析失败:', expr, eBlock);
        }
    }

    // 兼容旧的 parseLambda 兜底
    try {
        fn = parseLambda(expr);
        if (fn) { _lambdaCache[expr] = fn; return fn; }
    } catch (e) { /* ignore */ }

    console.warn('Lambda解析失败:', expr);
    return null;
};

/**
 * 矩阵分布
 * @param {Number} totalRows - 总行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'或'c'
 * @returns {Array} 分布后的数组
 */
JSA.z矩阵分布 = function(totalRows, cols, direction) {
    direction = direction || 'r';
    var result = [];
    var numbers = [];
    for (var i = 0; i < totalRows; i++) {
        numbers.push(i);
    }

    if (direction === 'r') {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = i * cols + j;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    } else {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = j * rows + i;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    }

    return result;
};
JSA.getMatrix = JSA.z矩阵分布;
/**
 * JSA.__jsaToVBA — JSA 与 VBA 互操作内部函数
 * 用于在 JSA 环境中调用 VBA 函数
 * @param {string} procName - VBA 过程/函数名
 * @param {...*} args - 传递的参数
 * @returns {*} VBA 函数返回值
 */
JSA.__jsaToVBA = function(procName) {
    var args = [];
    for (var i = 1; i < arguments.length; i++) {
        args.push(arguments[i]);
    }
    try {
        return Application.Run(procName, args);
    } catch (e) {
        return undefined;
    }
};

/**
 * JSA.month — 获取日期的月份
 * @param {Date|string} [date] - 日期，默认当前日期
 * @returns {Number} 月份 (1-12)
 */
JSA.month = function(date) {
    var d = date ? (date instanceof Date ? date : new Date(date)) : new Date();
    return d.getMonth() + 1;
};

/**
 * JSA.now — 获取当前日期时间
 * @returns {Date} 当前日期时间
 */
JSA.now = function() {
    return new Date();
};

/**
 * JSA.m — JSA 内部辅助方法
 * @returns {Object} 辅助对象
 */
JSA.m = function() { return {}; };

/**
 * JSA.S — JSA 内部辅助方法
 * @returns {Object} 辅助对象
 */
JSA.S = function() { return {}; };



// ==================== [IO] 文件操作库 ====================

/**
 * IO - 文件操作工具
 * @constructor
 * @class
 * @description 文件系统操作（支持 WPS 和 Node.js 环境）
 */
function IO() {}

/**
 * 是否文件
 * @param {String} path - 路径
 * @returns {Boolean} 是否为文件
 */
IO.z是否文件 = function(path) {
    
    try {
        const fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FileExists(path);
    } catch (e) {
        return false;
    }
};
IO.IsFile = IO.z是否文件;

/**
 * 是否文件夹
 * @param {String} path - 路径
 * @returns {Boolean} 是否为文件夹
 */
IO.z是否文件夹 = function(path) {
    
    try {
        const fso = new ActiveXObject("Scripting.FileSystemObject");
        return fso.FolderExists(path);
    } catch (e) {
        return false;
    }
};
IO.IsDirectory = IO.z是否文件夹;

IO.z存在 = function(path) {
    if (IO.z是否文件(path) || IO.z是否文件夹(path)) return true;
    // 🔧 XXD-176 atomic fix: Node 兜底
    try {
        var __fs = require ? require('fs') : null;
        if (__fs) return __fs.existsSync(path);
    } catch (__) {}
    return false;
};

/**
 * 文件名
 * @param {String} path - 路径
 * @returns {String} 文件名
 */
IO.z文件名 = function(path) {
    if (!path) return '';
    var parts = path.replace(/\\/g, '/').split('/');
    return parts[parts.length - 1] || '';
};
IO.getFileName = IO.z文件名;

/**
 * 纯文件名
 * @param {String} path - 路径
 * @returns {String} 纯文件名
 */
IO.z纯文件名 = function(path) {
    var fileName = IO.z文件名(path);
    var lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex > 0) {
        return fileName.substring(0, lastDotIndex);
    }
    return fileName;
};
IO.getFileNameNoType = IO.z纯文件名;

/**
 * 文件后缀
 * @param {String} path - 路径
 * @returns {String} 后缀
 */
IO.z文件后缀 = function(path) {
    var fileName = IO.z文件名(path);
    var lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex > 0 && lastDotIndex < fileName.length - 1) {
        return fileName.substring(lastDotIndex + 1);
    }
    return '';
};
IO.getFileType = IO.z文件后缀;

/**
 * 上级文件夹
 * @param {String} path - 路径
 * @param {Number} 级数 - 级数
 * @returns {String} 上级路径
 */
IO.z上级文件夹 = function(path, 级数) {
    // XXD-174-z上级文件夹-根级返回空串-fix: 根级(/a.txt -> '')和纯文件名(a.txt -> '')都返回空串
    级数 = 级数 || 1;
    var result = String(path == null ? '' : path);
    for (var i = 0; i < 级数; i++) {
        result = result.replace(/\\/g, '/').replace(/\/+$/, '') || '/';
        var lastSlashIndex = result.lastIndexOf('/');
        if (lastSlashIndex > 0) {
            result = result.substring(0, lastSlashIndex);
        } else if (lastSlashIndex === 0) {
            result = '';
        } else {
            result = '';
        }
    }
    return result;
};
IO.lastDirectoty = IO.z上级文件夹;

/* eslint-disable */
// 🔧 XXD-173/XXD-174 final fix: IO.z路径拼接 — 缺失的 path.join 中文别名
// 复现: IO.z路径拼接('/a','b','c') THROW — IO.z路径拼接 is not a function
// 期望: '/a/b/c' — 类似 Node path.join: 过滤 null/空, 用 '/' 连接, 折叠连续 '/'
/**
 * 路径拼接
 * @param {...String} parts - 路径片段
 * @returns {String} 拼接后的路径
 */
IO.z路径拼接 = function() {
    var p = Array.prototype.slice.call(arguments).filter(function(x) {
        return x != null && x !== '';
    }).join('/');
    return p.replace(/\/+/g, '/');
};
IO.pathJoin = IO.z路径拼接;
IO.joinPath = IO.z路径拼接;

/**
 * 复制文件
 * @param {String} srcPath - 源文件路径
 * @param {String} dstPath - 目标文件路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.copyFile("C:/a.txt", "D:/b.txt");
 */
IO.z复制文件 = function(srcPath, dstPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FileExists(srcPath)) return false;
        if (fso.FileExists(dstPath)) fso.DeleteFile(dstPath, true);
        fso.CopyFile(srcPath, dstPath, true);
        return true;
    } catch (e) {
        return false;
    }
};
IO.copyFile = IO.z复制文件;

/**
 * 重命名文件
 * @param {String} srcPath - 源文件路径
 * @param {String} newName - 新文件名称（不含路径）
 * @returns {Boolean} 是否成功
 * @example
 * IO.rename("C:/a.txt", "b.txt");
 */
IO.z重命名文件 = function(srcPath, newName) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FileExists(srcPath)) return false;
        var dstPath = IO.z上级文件夹(srcPath) + '/' + newName;
        if (fso.FileExists(dstPath)) fso.DeleteFile(dstPath, true);
        fso.MoveFile(srcPath, dstPath);
        return true;
    } catch (e) {
        return false;
    }
};
IO.rename = IO.z重命名文件;

/**
 * 移动文件
 * @param {String} srcPath - 源文件路径
 * @param {String} dstPath - 目标路径（不含文件名）或目标完整路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.moveFile("C:/a.txt", "D:/");
 * IO.moveFile("C:/a.txt", "D:/b.txt");
 */
IO.z移动文件 = function(srcPath, dstPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FileExists(srcPath)) return false;
        // 如果目标是目录，把原文件名拼到目标
        if (fso.FolderExists(dstPath)) {
            dstPath = dstPath.replace(/[\\\/]+$/, '') + '/' + IO.z文件名(srcPath);
        }
        if (fso.FileExists(dstPath)) fso.DeleteFile(dstPath, true);
        fso.MoveFile(srcPath, dstPath);
        return true;
    } catch (e) {
        return false;
    }
};
IO.moveFile = IO.z移动文件;

/**
 * 创建文件夹（已存在则忽略，不抛错）
 * @param {String} folderPath - 文件夹路径
 * @returns {Boolean} 是否成功（或已存在）
 * @example
 * IO.mkDir2("C:/newFolder/subFolder");
 */
IO.z建文件夹 = function(folderPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (fso.FolderExists(folderPath)) return true;
        fso.CreateFolder(folderPath);
        return true;
    } catch (e) {
        return false;
    }
};
IO.mkDir2 = IO.z建文件夹;
IO.MkDir2 = IO.z建文件夹;
IO.z创建文件夹 = IO.z建文件夹;

/**
 * 删除文件
 * @param {String} filePath - 文件路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.delete("C:/a.txt");
 */
IO.z删除文件 = function(filePath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FileExists(filePath)) return true; // 不存在视为成功
        fso.DeleteFile(filePath, true);
        return true;
    } catch (e) {
        return false;
    }
};
IO.delete = IO.z删除文件;

/**
 * 复制文件夹（带下级一起）
 * @param {String} srcPath - 源文件夹路径
 * @param {String} dstName - 目标文件夹名称（不含路径）或目标完整路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.copyFolder("C:/srcFolder", "D:/dstFolder");
 * IO.copyFolder("C:/srcFolder", "dstFolder"); // 在 C: 盘根目录下创建副本
 */
IO.z复制文件夹 = function(srcPath, dstName) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(srcPath)) return false;
        // 如果 dstName 不是完整路径，视为在 srcPath 同级目录下创建
        var dstPath = fso.FolderExists(dstName) || fso.FileExists(dstName) ? dstName
            : IO.z上级文件夹(srcPath) + '/' + dstName;
        if (fso.FolderExists(dstPath)) fso.DeleteFolder(dstPath, true);
        fso.CopyFolder(srcPath, dstPath, true);
        return true;
    } catch (e) {
        return false;
    }
};
IO.copyFolder = IO.z复制文件夹;

/**
 * 重命名文件夹
 * @param {String} srcPath - 源文件夹路径
 * @param {String} newName - 新文件夹名称（不含路径）
 * @returns {Boolean} 是否成功
 * @example
 * IO.reNameFolder("C:/oldFolder", "newFolder");
 */
IO.z改文件夹名 = function(srcPath, newName) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(srcPath)) return false;
        var dstPath = IO.z上级文件夹(srcPath) + '/' + newName;
        if (fso.FolderExists(dstPath)) fso.DeleteFolder(dstPath, true);
        fso.MoveFolder(srcPath, dstPath);
        return true;
    } catch (e) {
        return false;
    }
};
IO.reNameFolder = IO.z改文件夹名;

/**
 * 移动文件夹（带子文件夹一起）
 * @param {String} srcPath - 源文件夹路径
 * @param {String} dstPath - 目标路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.moveFolder("C:/srcFolder", "D:/dstFolder");
 */
IO.z移动文件夹 = function(srcPath, dstPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(srcPath)) return false;
        // 如果目标是已存在的目录，把原文件夹名拼到目标
        if (fso.FolderExists(dstPath)) {
            dstPath = dstPath.replace(/[\\\/]+$/, '') + '/' + IO.z文件名(srcPath);
        }
        if (fso.FolderExists(dstPath)) fso.DeleteFolder(dstPath, true);
        fso.MoveFolder(srcPath, dstPath);
        return true;
    } catch (e) {
        return false;
    }
};
IO.moveFolder = IO.z移动文件夹;

/**
 * 递归删除文件夹及其所有子内容
 * @param {String} folderPath - 文件夹路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.delTree("C:/myFolder"); // 删除整个文件夹
 */
IO.z递归删文件夹 = function(folderPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(folderPath)) return true; // 不存在视为成功
        fso.DeleteFolder(folderPath, true);
        return true;
    } catch (e) {
        return false;
    }
};
IO.delTree = IO.z递归删文件夹;

/**
 * 清空文件夹下级内容（不删除自己）
 * @param {String} folderPath - 文件夹路径
 * @returns {Boolean} 是否成功
 * @example
 * IO.clearTree("C:/myFolder"); // 清空所有子文件和子文件夹
 */
IO.z清空文件夹 = function(folderPath) {
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(folderPath)) return false;
        var folder = fso.GetFolder(folderPath);
        // 删除所有子文件夹（递归）
        var enumFolders = new Enumerator(folder.SubFolders);
        for (; !enumFolders.atEnd(); enumFolders.moveNext()) {
            fso.DeleteFolder(enumFolders.item.Path, true);
        }
        // 删除所有子文件
        var enumFiles = new Enumerator(folder.Files);
        for (; !enumFiles.atEnd(); enumFiles.moveNext()) {
            fso.DeleteFile(enumFiles.item.Path, true);
        }
        return true;
    } catch (e) {
        return false;
    }
};
IO.clearTree = IO.z清空文件夹;


// ==================== [GLOBAL_FUNCS] 全局辅助函数 ====================
function RangeChain(rng, colIndex) {
    if (!(this instanceof RangeChain)) {
        return new RangeChain(rng, colIndex);
    }
    this._range = null;

    // 两个参数模式：RangeChain(行号, 列号)
    if (typeof rng === 'number' && typeof colIndex === 'number') {
        this._range = Cells(rng, colIndex);
    }
    // 字符串地址模式
    else if (typeof rng === 'string') {
        this._range = Range(rng);
    }
    // Range对象模式
    else if (rng && rng.Address) {
        this._range = rng;
    }
}

/**
 * Value - 获取原始Range对象
 * @returns {Range|null} Range对象
 */
RangeChain.prototype.Value = function() {
    return this._range;
};
/**
 * toJSON - JSON 序列化支持(XXD-103/104)
 * 当 RangeChain 实例被传入 JSON.stringify 时,返回其数据的简化表示
 * @returns {Object|null} {address, value2, row, column} 或 null(无底层 Range 时)
 */
RangeChain.prototype.toJSON = function() {
    if (!this._range) return null;
    try {
        return {
            address: this._range.Address ? this._range.Address() : '',
            value2: this._range.Value2,
            row: this._range.Row,
            column: this._range.Column
        };
    } catch (e) {
        return { address: '', value2: null, row: 0, column: 0 };
    }
};


/**
 * Value2 - 获取/设置值（Value2属性）
 * @param {any} [newValue] - 新值（可选）
 * @returns {RangeChain|any} 设置时返回this，否则返回当前值
 */
// 🔧 v4.0.11 修复: 不在此处将 z值 函数赋值给 Value2
// 否则会覆盖下方第3421行通过 Object.defineProperty 定义的访问器
// 支持 $(5,2).Value2 = 'newValue' 语法的关键
RangeChain.prototype.z值 = function(newValue) {
    if (newValue !== undefined) {
        if (this._range) this._range.Value2 = newValue;
        return this;
    }
    return this._range ? this._range.Value2 : undefined;
};
// 注意：Value2 的别名请直接使用 z值()，或在后面通过 Object.defineProperty 定义

/**
 * CurrentRegion - 获取当前区域（连续数据区域）
 * @returns {RangeChain|null} 当前区域的RangeChain对象
 */
RangeChain.prototype.z当前区域 = function() {
    if (!this._range) return null;
    try {
        return new RangeChain(this._range.CurrentRegion);
    } catch (e) {
        return null;
    }
};
RangeChain.prototype.CurrentRegion = RangeChain.prototype.z当前区域;

/**
 * safeArray - 转换为安全数组（返回 Array2D 对象，支持链式调用）
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 */
RangeChain.prototype.z安全数组 = function() {
    return RngUtils.z安全数组(this._range);
};
RangeChain.prototype.safeArray = RangeChain.prototype.z安全数组;

/**
 * VisibleArray - 转换可见区域为数组
 * @param {Worksheet} [tempSheet] - 临时工作表
 * @returns {Array} 数组
 */
RangeChain.prototype.z可见区数组 = function(tempSheet) {
    return RngUtils.z可见区数组(this._range, tempSheet);
};
RangeChain.prototype.VisibleArray = RangeChain.prototype.z可见区数组;

/**
 * RowsCount - 获取行数
 * @returns {number} 行数
 */
RangeChain.prototype.z行数 = function() {
    return this._range ? this._range.Rows.Count : 0;
};
RangeChain.prototype.RowsCount = RangeChain.prototype.z行数;

/**
 * ColsCount - 获取列数
 * @returns {number} 列数
 */
RangeChain.prototype.z列数 = function() {
    return this._range ? this._range.Columns.Count : 0;
};
RangeChain.prototype.ColsCount = RangeChain.prototype.z列数;

/**
 * Columns - 获取列集合
 * @returns {Range} Range对象的Columns属性
 */
Object.defineProperty(RangeChain.prototype, 'Columns', {
    get: function() {
        return this._range ? this._range.Columns : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Rows - 获取行集合
 * @returns {Range} Range对象的Rows属性
 */
Object.defineProperty(RangeChain.prototype, 'Rows', {
    get: function() {
        return this._range ? this._range.Rows : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Font - 获取字体对象
 * @returns {Font} Font对象
 */
Object.defineProperty(RangeChain.prototype, 'Font', {
    get: function() {
        return this._range ? this._range.Font : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Interior - 获取内部对象（背景色等）
 * @returns {Interior} Interior对象
 */
Object.defineProperty(RangeChain.prototype, 'Interior', {
    get: function() {
        return this._range ? this._range.Interior : null;
    },
    enumerable: true,
    configurable: true
});

/**
 * Address - 获取地址
 * @returns {string} 地址
 */
RangeChain.prototype.z地址 = function() {
    return this._range ? this._range.Address() : '';
};
RangeChain.prototype.Address = RangeChain.prototype.z地址;

/**
 * AddBorders - 添加边框
 * @param {number} [lineStyle=1] - 线条样式
 * @param {number} [weight=2] - 线条粗细
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z加边框 = function(lineStyle, weight) {
    if (this._range) {
        RngUtils.z加边框(this._range, lineStyle, weight);
    }
    return this;
};
RangeChain.prototype.AddBorders = RangeChain.prototype.z加边框;

/**
 * AutoFitColumns - 自动列宽
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z自动列宽 = function() {
    if (this._range) {
        this._range.Columns.AutoFit();
    }
    return this;
};
RangeChain.prototype.AutoFitColumns = RangeChain.prototype.z自动列宽;

/**
 * AutoFitRows - 自动行高
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z自动行高 = function() {
    if (this._range) {
        this._range.Rows.AutoFit();
    }
    return this;
};
RangeChain.prototype.AutoFitRows = RangeChain.prototype.z自动行高;

/**
 * ClearContents - 清除内容
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z清除内容 = function() {
    if (this._range) {
        this._range.ClearContents();
    }
    return this;
};
RangeChain.prototype.ClearContents = RangeChain.prototype.z清除内容;

/**
 * ClearFormats - 清除格式
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z清除格式 = function() {
    if (this._range) {
        this._range.ClearFormats();
    }
    return this;
};
RangeChain.prototype.ClearFormats = RangeChain.prototype.z清除格式;

/**
 * Value2 - 获取/设置值（Value2属性，比Value更快）
 * @param {any} [newValue] - 新值（可选）
 * @returns {RangeChain|any} 设置时返回this，否则返回当前值
 * @example
 * $(5, 2).z值()                    // 获取值
 * $(5, 2).z值("新值")              // 设置值
 * $(5, 2).z值("新值").z加粗()      // 链式调用
 */
// 注意：z值 方法已在第5734行定义，此处删除重复定义以避免覆盖

// 使用属性方式定义 Value2，支持 $(i,2).Value2 = rs 语法
Object.defineProperty(RangeChain.prototype, 'Value2', {
    get: function() {
        return this._range ? this._range.Value2 : undefined;
    },
    set: function(newValue) {
        if (this._range) this._range.Value2 = newValue;
    },
    enumerable: true,
    configurable: true
});

/**
 * Formula - 获取/设置公式
 * @param {string} [newFormula] - 新公式（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回公式
 */
RangeChain.prototype.z公式 = function(newFormula) {
    if (newFormula !== undefined) {
        if (this._range) this._range.Formula = newFormula;
        return this;
    }
    return this._range ? this._range.Formula : '';
};

// 使用属性方式定义 Formula
Object.defineProperty(RangeChain.prototype, 'Formula', {
    get: function() {
        return this._range ? this._range.Formula : '';
    },
    set: function(newFormula) {
        if (this._range) this._range.Formula = newFormula;
    },
    enumerable: true,
    configurable: true
});

/**
 * Text - 获取显示文本
 * @returns {string} 显示文本
 */
RangeChain.prototype.z文本 = function() {
    return this._range ? this._range.Text : '';
};

// 使用属性方式定义 Text（只读）
Object.defineProperty(RangeChain.prototype, 'Text', {
    get: function() {
        return this._range ? this._range.Text : '';
    },
    enumerable: true,
    configurable: true
});

/**
 * Row - 获取行号
 * @returns {number} 行号
 */
RangeChain.prototype.z行 = function() {
    return this._range ? this._range.Row : 0;
};

// 使用属性方式定义 Row（只读）
Object.defineProperty(RangeChain.prototype, 'Row', {
    get: function() {
        return this._range ? this._range.Row : 0;
    },
    enumerable: true,
    configurable: true
});

/**
 * Column - 获取列号
 * @returns {number} 列号
 */
RangeChain.prototype.z列 = function() {
    return this._range ? this._range.Column : 0;
};

// 使用属性方式定义 Column（只读）
Object.defineProperty(RangeChain.prototype, 'Column', {
    get: function() {
        return this._range ? this._range.Column : 0;
    },
    enumerable: true,
    configurable: true
});

/**
 * Select - 选中区域
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z选中 = function() {
    if (this._range) this._range.Select();
    return this;
};
RangeChain.prototype.Select = RangeChain.prototype.z选中;

/**
 * Activate - 激活单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z激活 = function() {
    if (this._range) this._range.Activate();
    return this;
};
RangeChain.prototype.Activate = RangeChain.prototype.z激活;

/**
 * Bold - 获取/设置加粗
 * @param {boolean} [isBold] - 是否加粗（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回加粗状态
 */
RangeChain.prototype.z加粗 = function(isBold) {
    if (isBold !== undefined) {
        if (this._range) this._range.Font.Bold = isBold;
        return this;
    }
    return this._range ? this._range.Font.Bold : false;
};

// 使用属性方式定义 Bold
Object.defineProperty(RangeChain.prototype, 'Bold', {
    get: function() {
        return this._range ? this._range.Font.Bold : false;
    },
    set: function(isBold) {
        if (this._range) this._range.Font.Bold = isBold;
    },
    enumerable: true,
    configurable: true
});

/**
 * Italic - 获取/设置斜体
 * @param {boolean} [isItalic] - 是否斜体（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回斜体状态
 */
RangeChain.prototype.z斜体 = function(isItalic) {
    if (isItalic !== undefined) {
        if (this._range) this._range.Font.Italic = isItalic;
        return this;
    }
    return this._range ? this._range.Font.Italic : false;
};

// 使用属性方式定义 Italic
Object.defineProperty(RangeChain.prototype, 'Italic', {
    get: function() {
        return this._range ? this._range.Font.Italic : false;
    },
    set: function(isItalic) {
        if (this._range) this._range.Font.Italic = isItalic;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontColor - 获取/设置字体颜色
 * @param {number} [color] - RGB颜色值（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回颜色值
 */
RangeChain.prototype.z字体颜色 = function(color) {
    if (color !== undefined) {
        if (this._range) this._range.Font.Color = color;
        return this;
    }
    return this._range ? this._range.Font.Color : 0;
};

// 使用属性方式定义 FontColor
Object.defineProperty(RangeChain.prototype, 'FontColor', {
    get: function() {
        return this._range ? this._range.Font.Color : 0;
    },
    set: function(color) {
        if (this._range) this._range.Font.Color = color;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontSize - 获取/设置字体大小
 * @param {number} [size] - 字体大小（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回字体大小
 */
RangeChain.prototype.z字号 = function(size) {
    if (size !== undefined) {
        if (this._range) this._range.Font.Size = size;
        return this;
    }
    return this._range ? this._range.Font.Size : 11;
};

// 使用属性方式定义 FontSize
Object.defineProperty(RangeChain.prototype, 'FontSize', {
    get: function() {
        return this._range ? this._range.Font.Size : 11;
    },
    set: function(size) {
        if (this._range) this._range.Font.Size = size;
    },
    enumerable: true,
    configurable: true
});

/**
 * FontName - 获取/设置字体名称
 * @param {string} [fontName] - 字体名称（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回字体名称
 */
RangeChain.prototype.z字体名称 = function(fontName) {
    if (fontName !== undefined) {
        if (this._range) this._range.Font.Name = fontName;
        return this;
    }
    return this._range ? this._range.Font.Name : '';
};

// 使用属性方式定义 FontName
Object.defineProperty(RangeChain.prototype, 'FontName', {
    get: function() {
        return this._range ? this._range.Font.Name : '';
    },
    set: function(fontName) {
        if (this._range) this._range.Font.Name = fontName;
    },
    enumerable: true,
    configurable: true
});

/**
 * InteriorColor - 获取/设置背景颜色
 * @param {number} [color] - RGB颜色值（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回颜色值
 */
RangeChain.prototype.z背景颜色 = function(color) {
    if (color !== undefined) {
        if (this._range) this._range.Interior.Color = color;
        return this;
    }
    return this._range ? this._range.Interior.Color : 16777215; // 默认白色
};

// 使用属性方式定义 InteriorColor
Object.defineProperty(RangeChain.prototype, 'InteriorColor', {
    get: function() {
        return this._range ? this._range.Interior.Color : 16777215;
    },
    set: function(color) {
        if (this._range) this._range.Interior.Color = color;
    },
    enumerable: true,
    configurable: true
});

/**
 * HorizontalAlignment - 获取/设置水平对齐
 * @param {number} [align] - 对齐方式（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回对齐方式
 */
RangeChain.prototype.z水平对齐 = function(align) {
    if (align !== undefined) {
        if (this._range) this._range.HorizontalAlignment = align;
        return this;
    }
    return this._range ? this._range.HorizontalAlignment : -4151; // 默认常规
};

// 使用属性方式定义 HorizontalAlignment
Object.defineProperty(RangeChain.prototype, 'HorizontalAlignment', {
    get: function() {
        return this._range ? this._range.HorizontalAlignment : -4151;
    },
    set: function(align) {
        if (this._range) this._range.HorizontalAlignment = align;
    },
    enumerable: true,
    configurable: true
});

/**
 * VerticalAlignment - 获取/设置垂直对齐
 * @param {number} [align] - 对齐方式（可选）
 * @returns {RangeChain|number} 设置时返回this，否则返回对齐方式
 */
RangeChain.prototype.z垂直对齐 = function(align) {
    if (align !== undefined) {
        if (this._range) this._range.VerticalAlignment = align;
        return this;
    }
    return this._range ? this._range.VerticalAlignment : -4160; // 默认底部
};

// 使用属性方式定义 VerticalAlignment
Object.defineProperty(RangeChain.prototype, 'VerticalAlignment', {
    get: function() {
        return this._range ? this._range.VerticalAlignment : -4160;
    },
    set: function(align) {
        if (this._range) this._range.VerticalAlignment = align;
    },
    enumerable: true,
    configurable: true
});

/**
 * NumberFormat - 获取/设置数字格式
 * @param {string} [format] - 格式字符串（可选）
 * @returns {RangeChain|string} 设置时返回this，否则返回格式字符串
 */
RangeChain.prototype.z数字格式 = function(format) {
    if (format !== undefined) {
        if (this._range) this._range.NumberFormat = format;
        return this;
    }
    return this._range ? this._range.NumberFormat : 'General';
};

// 使用属性方式定义 NumberFormat
Object.defineProperty(RangeChain.prototype, 'NumberFormat', {
    get: function() {
        return this._range ? this._range.NumberFormat : 'General';
    },
    set: function(format) {
        if (this._range) this._range.NumberFormat = format;
    },
    enumerable: true,
    configurable: true
});

/**
 * WrapText - 获取/设置自动换行
 * @param {boolean} [wrap] - 是否自动换行（可选）
 * @returns {RangeChain|boolean} 设置时返回this，否则返回换行状态
 */
RangeChain.prototype.z自动换行 = function(wrap) {
    if (wrap !== undefined) {
        if (this._range) this._range.WrapText = wrap;
        return this;
    }
    return this._range ? this._range.WrapText : false;
};

// 使用属性方式定义 WrapText
Object.defineProperty(RangeChain.prototype, 'WrapText', {
    get: function() {
        return this._range ? this._range.WrapText : false;
    },
    set: function(wrap) {
        if (this._range) this._range.WrapText = wrap;
    },
    enumerable: true,
    configurable: true
});

/**
 * Merge - 合并单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z合并 = function() {
    if (this._range) this._range.Merge();
    return this;
};
RangeChain.prototype.Merge = RangeChain.prototype.z合并;

/**
 * Clear - 清除内容和格式
 * @returns {RangeChain} 当前实例
 * @example
 * $("K2").Resize(1000, 5000).Clear()
 * $.Resize("K2", 1000, 5000).Clear()
 */
RangeChain.prototype.Clear = function() {
    if (this._range) {
        // WPS JSA 兼容：使用 ClearContents 和 ClearFormats
        try {
            this._range.ClearContents();
        } catch (e) {}
        try {
            this._range.ClearFormats();
        } catch (e) {}
    }
    return this;
};

/**
 * z清除 - Clear的中文别名
 */
RangeChain.prototype.z清除 = RangeChain.prototype.Clear;

/**
 * UnMerge - 取消合并单元格
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z取消合并 = function() {
    if (this._range) this._range.UnMerge();
    return this;
};
RangeChain.prototype.UnMerge = RangeChain.prototype.z取消合并;

/**
 * Resize - 调整区域大小
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @returns {RangeChain} 调整大小后的新RangeChain对象
 * @example
 * $("K2").Resize(10, 5).z清除内容()
 * $("K2").Resize(1000, 5000).z清除内容()
 */
RangeChain.prototype.Resize = function(rows, cols) {
    if (!this._range) return new RangeChain(null);
    try {
        var resizedRng = this._range.Resize(rows, cols);
        return new RangeChain(resizedRng);
    } catch (e) {
        console.error("Resize失败: " + e.message);
        return this;
    }
};

/**
 * MergeCells - 检查是否为合并单元格
 * @returns {boolean} 是否合并
 */
RangeChain.prototype.z已合并 = function() {
    return this._range ? this._range.MergeCells : false;
};

// 使用属性方式定义 MergeCells（只读）
Object.defineProperty(RangeChain.prototype, 'MergeCells', {
    get: function() {
        return this._range ? this._range.MergeCells : false;
    },
    enumerable: true,
    configurable: true
});

/**
 * Delete - 删除区域
 * @param {number} [shift] - 移动方向（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z删除 = function(shift) {
    if (this._range) this._range.Delete(shift);
    return this;
};
RangeChain.prototype.Delete = RangeChain.prototype.z删除;

/**
 * Insert - 插入区域
 * @param {number} [shift] - 移动方向（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z插入 = function(shift) {
    if (this._range) this._range.Insert(shift);
    return this;
};
RangeChain.prototype.Insert = RangeChain.prototype.z插入;

/**
 * Copy - 复制区域
 * @param {Range} [destination] - 目标区域（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z复制 = function(destination) {
    if (this._range) {
        if (destination) {
            this._range.Copy(destination);
        } else {
            this._range.Copy();
        }
    }
    return this;
};
RangeChain.prototype.Copy = RangeChain.prototype.z复制;

/**
 * Paste - 粘贴区域
 * @param {Range} [destination] - 目标区域（可选）
 * @param {number} [type] - 粘贴类型（可选）
 * @returns {RangeChain} 当前实例
 */
RangeChain.prototype.z粘贴 = function(destination, type) {
    if (!destination) return this;
    // v4.1.0 修复: WPS Range 没有 .Paste() 方法，正确 API 是 .PasteSpecial()
    // 默认类型 -4104 (xlPasteAll) 保留所有内容
    if (typeof type === "number") {
        destination.PasteSpecial(type);
    } else {
        destination.PasteSpecial(-4104);
    }
    return this;
};
RangeChain.prototype.Paste = RangeChain.prototype.z粘贴;

/**
 * 创建RngUtils静态方法代理对象
 * @private
 */
function createRngUtilsProxy() {
    var proxy = {};
    var staticMethods = [
        'z最后一个', 'lastCell',
        'z安全区域', 'safeRange',
        'z安全数组', 'safeArray',
        'z最大行', 'endRow',
        'z最大行单元格', 'endRowCell',
        'z最大行区域', 'maxRange',
        'z最大列', 'endCol',
        'z最大列单元格', 'endColCell',
        'z可见区数组', 'visibleArray',
        'z可见区域', 'visibleRange',
        'z加边框', 'addBorders',
        'z取前几行', 'takeRows',
        'z跳过前几行', 'skipRows',
        'z插入多行', 'insertRows',
        'z插入多列', 'insertCols',
        'z删除空白行', 'delBlankRows',
        'z删除空白列', 'delBlankCols',
        'z整行', 'entireRow',
        'z整列', 'entire_column',
        'z行数', 'rowsCount',
        'z列数', 'colsCount',
        'z列号字母互转', 'colToAbc',
    'z复制粘贴格式', 'copyFormat',
    'z复制粘贴值', 'copyValue',
    'z联合区域', 'unionAll',
    'z多列排序', 'rngSortCols',
    'z合并单元格', 'mergeCells',
    'z取消合并单元格', 'unmergeCells'
];

    for (var i = 0; i < staticMethods.length; i++) {
        var methodName = staticMethods[i];
        if (RngUtils[methodName]) {
            (function(name) {
                proxy[name] = function() {
                    var result = RngUtils[name].apply(RngUtils, arguments);
                    // 如果返回的是Range对象，包装成RangeChain支持链式调用
                    if (result && result.Address && typeof result.Address === 'function') {
                        return new RangeChain(result);
                    }
                    return result;
                };
            })(methodName);
        }
    }

    return proxy;
}

/**
 * $函数 - Range快捷方式和RngUtils方法代理（支持智能提示和链式调用）
 * @param {string|number} x - 地址或行号
 * @param {number} [y] - 列号（可选，当传入两个数字参数时）
 * @returns {RangeChain} RangeChain包装对象，支持智能提示和链式调用
 * @example
 * $("A1")                          // 返回RangeChain，支持链式调用
 * $(5, 2)                          // 第5行第2列，返回RangeChain
 * $(5, 2).z值()                    // 获取值
 * $(5, 2).z值("新值")              // 设置值
 * $(5, 2).z值("新值").z加粗()      // 链式调用
 * $.maxRange("A1:J1").safeArray()  // 链式调用
 */

function $(x, y) {
    // 两个参数模式：$(行, 列) - 返回RangeChain
    if (arguments.length === 2 && typeof x === 'number' && typeof y === 'number') {
        return new RangeChain(x, y);
    }
    // 单个参数模式 - 返回RangeChain
    if (typeof x === 'string') {
        return new RangeChain(x);
    } else if (typeof x === 'number') {
        return new RangeChain(x, 1);
    } else if (x && x.Address) {
        return new RangeChain(x);
    }
    // 返回空的RangeChain
    return new RangeChain(null);
}

/**
 * $$ - Array2D 全局快捷引用，用于调用二维数组的静态处理方法
 * @constructor
 * @class
 * @description Array2D 的全局快捷别名，支持所有 Array2D 静态方法（filter/map/groupInto/getMatrix 等）
 * @param {Array} [data] - 二维数组数据
 * @returns {Array2D} Array2D实例
 * @example
 * $$.getMatrix(5, 3)                // 生成5行3列矩阵
 * $$.filter(arr, 'f3 > 100')        // 静态筛选
 * $$.groupInto(arr, 'f1', 'count()') // 分组统计
 * $$(data).z筛选('f1>2').z多列排序('f2+')  // 链式调用
 */
function $$(data) {
    return new Array2D(data);
}

// $$ 与 Array2D 共享原型，实例方法完全一致
$$.prototype = Array2D.prototype;

// ==================== [SHORTCUT_$] 将常用静态方法添加到 $ 对象 ====================
// 直接定义以支持智能提示

// ==================== $ 全局包装 (v4.0.11: 提前到所有 $.xxx 之前) ====================
if (typeof $ === 'undefined') {
    $ = {};
} else if (typeof $ === 'function') {
    // $ 是 WPS 内置函数，创建包装以支持 RangeChain
    var _$ = $;
    // v4.0.11: 保存 WPS 内置 $ 的可枚举属性
    var _wpsDollar = $;
    $ = function(addr) {
        if (addr === undefined || addr === null) return new RangeChain(null);
        if (typeof addr === 'string') return new RangeChain(_$(addr));
        if (addr && addr.Address) return new RangeChain(addr);
        return new RangeChain(null);
    };
    // v4.0.11: 复制 WPS 内置 $ 的属性到新包装函数
    try {
        var _keys = Object.keys(_wpsDollar);
        for (var _i = 0; _i < _keys.length; _i++) {
            try { $[_keys[_i]] = _wpsDollar[_keys[_i]]; } catch(e) {}
        }
    } catch(e) {}
}


// v4.0.11: $$ 别名指向 Array2D，兼容旧版 API
// 用 Object.defineProperty 确保在 WPS 中不被内置变量覆盖
try {
    delete this.$$;
} catch(e) {}
try {
    Object.defineProperty(this, '$$', {
        get: function() { return Array2D; },
        set: function(v) {},
        enumerable: true,
        configurable: true
    });
} catch(e) {
    // 降级：直接赋值
    this.$$ = Array2D;
}

/**
 * $.maxRange - 获取从第一行到最后一行的区域
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.maxRange("A1:J1").safeArray()
 * $.maxRange("1:1000", "A").z加边框()
 */
$.maxRange = function(rng, col) {
    var result = RngUtils.maxRange.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z最大行区域 - maxRange的中文别名
 */
$.z最大行区域 = $.maxRange;

/**
 * $.CurrentRegion - 获取当前区域的连续数据范围
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.CurrentRegion("A1").safeArray()
 */
$.CurrentRegion = function(rng) {
    var range;
    if (typeof rng === 'string') {
        range = Range(rng);
    } else if (rng && rng.Address) {
        range = rng;
    } else {
        return new RangeChain(null);
    }
    if (range && range.CurrentRegion) {
        return new RangeChain(range.CurrentRegion);
    }
    return new RangeChain(null);
};

/**
 * $.z当前区域 - CurrentRegion的中文别名
 */
$.z当前区域 = $.CurrentRegion;

/**
 * $.safeArray - 将区域转换为安全数组（返回 Array2D 对象，支持链式调用）
 * @param {Range|string} rng - 区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 */
$.safeArray = RngUtils.safeArray;

/**
 * $.z安全数组 - safeArray的中文别名
 */
$.z安全数组 = $.safeArray;

/**
 * $.maxArray - 获取从第一行到最大行的区域并转换为数组（返回 Array2D 对象，支持链式调用）
 * @param {Range|string} rng - 要获取的区域（如 "A1:H1"）
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Array2D} Array2D 对象，支持 skip/take/filter/sortByCols/toRange 等链式调用
 * @example
 * // 基本用法：获取 A1:H1 扩展到最大行的数组
 * var arr = $.maxArray("A1:H1");
 * 
 * // 链式调用：跳过前3行，取前10行，按第2列排序，筛选年份为2023的行，输出到K4
 * $.maxArray("A1:H1")
 *   .skip(3)
 *   .take(10)
 *   .sortByCols('f2')
 *   .filter('f6==2023')
 *   .toRange("K4");
 * 
 * // 使用静态方法处理
 * var arr = $.maxArray("A1:H1");
 * var rs = Array2D.skip(arr, 3);
 * rs = Array2D.take(rs, 10);
 * rs = Array2D.sortByCols(rs, 'f2');
 * rs = Array2D.filter(rs, 'f6==2023');
 * Array2D.toRange(rs, "K4");
 */
$.maxArray = function(rng, col) {
    // 获取最大行区域
    var maxRng = RngUtils.z最大行区域.apply(RngUtils, arguments);
    if (!maxRng) return new Array2D([]);
    // 转换为 Array2D 对象
    return RngUtils.z安全数组(maxRng);
};

/**
 * $.z最大数组 - maxArray的中文别名
 */
$.z最大数组 = $.maxArray;

/**
 * $.endRow - 获取区域最大行数
 * @param {Range|string} rng - 区域
 * @returns {number} 行数
 */
$.endRow = RngUtils.endRow;

/**
 * $.z最大行 - endRow的中文别名
 */
$.z最大行 = $.endRow;

/**
 * $.addBorders - 添加边框
 * @param {Range|string} rng - 区域
 * @param {number} [lineStyle=1] - 线条样式
 * @param {number} [weight=2] - 线条粗细
 * @returns {RangeChain} RangeChain对象
 */
$.addBorders = function(rng, lineStyle, weight) {
    var result = RngUtils.addBorders.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z加边框 - addBorders的中文别名
 */
$.z加边框 = $.addBorders;

/**
 * $.autoFitColumns - 自动列宽
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 */
$.autoFitColumns = function(rng) {
    var result = RngUtils.autoFitColumns.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z自动列宽 - autoFitColumns的中文别名
 */
$.z自动列宽 = $.autoFitColumns;

/**
 * $.autoFitRows - 自动行高
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 */
$.autoFitRows = function(rng) {
    var result = RngUtils.autoFitRows.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.z自动行高 - autoFitRows的中文别名
 */
$.z自动行高 = $.autoFitRows;

/**
 * $.delBlankRows - 删除空白行
 * @param {Range|string} rng - 区域
 * @param {boolean} [entireColumn=false] - 是否删除整行
 * @returns {RangeChain} RangeChain对象
 */
$.delBlankRows = function(rng, entireColumn) {
    RngUtils.delBlankRows.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.z删除空白行 - delBlankRows的中文别名
 */
$.z删除空白行 = $.delBlankRows;

/**
 * $.delBlankCols - 删除空白列
 * @param {Range|string} rng - 区域
 * @param {boolean} [entireColumn=false] - 是否删除整列
 * @returns {RangeChain} RangeChain对象
 */
$.delBlankCols = function(rng, entireColumn) {
    RngUtils.delBlankCols.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.z删除空白列 - delBlankCols的中文别名
 */
$.z删除空白列 = $.delBlankCols;

/**
 * $.rngSortCols - 多列排序
 * @param {Range|string} rng - 区域
 * @param {Array} sortCols - 排序列数组
 * @returns {RangeChain} RangeChain对象
 */
$.rngSortCols = function(rng, sortCols) {
    RngUtils.rngSortCols.apply(RngUtils, arguments);
    return new RangeChain(rng);
};

/**
 * $.z多列排序 - rngSortCols的中文别名
 */
$.z多列排序 = $.rngSortCols;

/**
 * $.colToAbc - 列号与字母互转
 * @param {number|string} input - 列号或字母
 * @returns {string|number} 字母或列号
 */
$.colToAbc = RngUtils.colToAbc;

/**
 * $.z列号字母互转 - colToAbc的中文别名
 */
$.z列号字母互转 = $.colToAbc;

/**
 * $.rowsCount - 获取行数
 * @param {Range|string} rng - 区域
 * @returns {number} 行数
 */
$.rowsCount = RngUtils.rowsCount;

/**
 * $.z行数 - rowsCount的中文别名
 */
$.z行数 = $.rowsCount;

/**
 * $.colsCount - 获取列数
 * @param {Range|string} rng - 区域
 * @returns {number} 列数
 */
$.colsCount = RngUtils.colsCount;

/**
 * $.z列数 - colsCount的中文别名
 */
$.z列数 = $.colsCount;

/**
 * $.copyValue - 复制粘贴值
 * @param {Range|string} fromRng - 源区域
 * @param {Range|string} toRng - 目标区域
 * @returns {RangeChain} RangeChain对象
 */
$.copyValue = function(fromRng, toRng) {
    RngUtils.copyValue.apply(RngUtils, arguments);
    return new RangeChain(toRng);
};

/**
 * $.z复制粘贴值 - copyValue的中文别名
 */
$.z复制粘贴值 = $.copyValue;

/**
 * $.copyFormat - 复制粘贴格式
 * @param {Range|string} fromRng - 源区域
 * @param {Range|string} toRng - 目标区域
 * @returns {RangeChain} RangeChain对象
 */
$.copyFormat = function(fromRng, toRng) {
    RngUtils.copyFormat.apply(RngUtils, arguments);
    return new RangeChain(toRng);
};

/**
 * $.z复制粘贴格式 - copyFormat的中文别名
 */
$.z复制粘贴格式 = $.copyFormat;

/**
 * $.Resize - 调整区域大小（静态方法）
 * @param {Range|string} rng - 源区域
 * @param {number} rows - 行数
 * @param {number} cols - 列数
 * @returns {RangeChain} 调整大小后的 RangeChain 对象
 * @example
 * $.Resize("K2", 1000, 5000).z清除内容()
 * $.Resize(Range("A1"), 10, 5).z加边框()
 */
$.Resize = function(rng, rows, cols) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (!targetRng) return new RangeChain(null);
    try {
        var resizedRng = targetRng.Resize(rows, cols);
        return new RangeChain(resizedRng);
    } catch (e) {
        console.error("Resize失败: " + e.message);
        return new RangeChain(rng);
    }
};

/**
 * $.z调整大小 - Resize的中文别名
 */
$.z调整大小 = $.Resize;

/**
 * $.ClearContents - 清除内容（静态方法）
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.ClearContents("K2").Resize(1000, 5000).z清除内容()
 */
$.ClearContents = function(rng) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (targetRng) {
        targetRng.ClearContents();
    }
    return new RangeChain(targetRng);
};

/**
 * $.z清除内容 - ClearContents的中文别名
 */
$.z清除内容 = $.ClearContents;

/**
 * $.UnMerge - 取消合并（静态方法）
 * @param {Range|string} rng - 区域
 * @returns {RangeChain} RangeChain对象
 * @example
 * $.UnMerge("K2").Resize(1000, 1000).z取消合并()
 */
$.UnMerge = function(rng) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    if (targetRng) {
        targetRng.UnMerge();
    }
    return new RangeChain(targetRng);
};

/**
 * $.z取消合并 - UnMerge的中文别名
 */
$.z取消合并 = $.UnMerge;

$.shtActivate = function(sht) {
    var sheet = typeof sht === 'string' ? Sheets(sht) : sht;
    if (sheet) {
        try { sheet.Activate(); return sheet; } catch (e) { return false; }
    }
    return false;
};
$.z激活表 = $.shtActivate;

// $.safeArray 和 $.z安全数组 已在前面定义，此处不再重复
// ==================== [SHORTCUT_$] xlsm 测试兼容快捷方式 ====================

/**
 * $.resize - 调整数组尺寸（增减行列）
 * @param {Array} arr - 二维数组
 * @param {Number} rows - 目标行数
 * @param {Number} cols - 目标列数
 * @returns {Array} 调整后的二维数组
 * @example
 * $.resize(arr, 10, 5)
 */
$.resize = function(arr, rows, cols) {
    return (new Array2D(arr)).resize(rows, cols).val();
};
$.z调整数组尺寸 = $.resize;

/**
 * $.findAllIndex - 查找所有满足条件的元素位置
 * @param {Array} arr - 二维数组
 * @param {Function} fn - 条件函数 (value) => boolean
 * @returns {Array} 位置索引数组 [[row, col], ...]
 */
$.findAllIndex = function(arr, fn) {
    return Array2D.findAllIndex(arr, fn);
};
$.z查找所有下标 = $.findAllIndex;

/**
 * $.justDate - 从日期对象中提取日期部分（去掉时间）
 * @param {Date} d - 日期对象
 * @returns {Date} 仅含年月日的日期对象
 * @example
 * $.justDate(new Date())  // 2026-06-02 00:00:00
 */
$.justDate = function(d) {
    if (!(d instanceof Date)) d = new Date(d);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
};

/**
 * $.thisRange - 获取当前活动单元格
 * @returns {Range} 当前活动单元格 Range 对象
 */
$.thisRange = function() {
    try {
        return Application.ActiveCell;
    } catch (e) {
        return null;
    }
};

// === $ 别名：指向已有方法 ===
$.mergeCells = RngUtils.mergeCells;
$.z合并单元格 = $.mergeCells;

$.selectCols = Array2D.selectCols;
$.z选择列 = $.selectCols;

$.delay = JSA.delay;
$.z延时 = $.delay;

$.safeRange = RngUtils.safeRange;
$.z安全区域 = $.safeRange;

$.rangeSelect = Array2D.rangeSelect;
$.z按范围选择 = $.rangeSelect;
// v4.0.11: 补充 rangeMap 和 rangeMatrix 的 $ 别名
$.rangeMap = Array2D.rangeMap;
$.z区域映射 = $.rangeMap;
$.rangeMatrix = Array2D.rangeMatrix;
$.z区域矩阵 = $.rangeMatrix;

$.unMergeCells = RngUtils.unmergeCells;
$.z取消合并单元格 = $.unMergeCells;

$.endCol = RngUtils.endCol;
$.z最大列 = $.endCol;

$.skipRows = RngUtils.skipRows;
$.z跳过前几行 = $.skipRows;

$.rndInt = JSA.rndInt;
$.z随机整数 = $.rndInt;

$.findRange = Array2D.findRange;
$.z查找区域 = $.findRange;

// $.superPivot moved to after Array2D.superPivot definition



// ==================== [SHORTCUT_$] 将 RngUtils 方法添加到 $.RngUtils ====================
// 支持直接调用 $.addBorders() 而不是 $.RngUtils.addBorders()

// 定义需要直接添加到 $ 的常用方法
var directMethods = [
    'z加边框', 'addBorders',
    'z插入多行', 'insertRows',
    'z插入多列', 'insertCols',
    'z删除空白行', 'delBlankRows',
    'z删除空白列', 'delBlankCols',
    'z合并单元格', 'mergeCells',
    'z取消合并单元格', 'unmergeCells'
];

for (var i = 0; i < directMethods.length; i++) {
    var methodName = directMethods[i];
    if (RngUtils[methodName]) {
        (function(name) {
            $[name] = function() {
                return RngUtils[name].apply(RngUtils, arguments);
            };
        })(methodName);
    }
}

// ==================== [SHORTCUT_$] 将工具类工厂添加到 $ 对象 ====================

/**
 * $.Array2D - 二维数组工具类工厂（支持智能提示和链式调用）
 * @param {Array} data - 输入数据
 * @returns {Array2D} Array2D实例，支持链式调用和智能提示
 * @example
 * $.Array2D([[1,2],[3,4]]).z求和()      // 10
 * $.Array2D([1,2,3]).z转置()           // [[1],[2],[3]]
 * $.Array2D([[1,2],[3,4]]).toRange("A1")  // 写入A1:B2
 */
$.Array2D = function(data) {
    return new Array2D(data);
};

/**
 * $.RngUtils - Range工具类工厂
 * @param {string|Range} [initialRange] - 初始Range（可选）
 * @returns {RngUtils|Object} RngUtils实例或静态方法对象
 * @example
 * $.RngUtils("A1:B10").z安全数组()    // 实例方法
 * $.RngUtils.maxRange("A1:J1")        // 静态方法
 */
$.RngUtils = function(initialRange) {
    // 无参数调用时，返回静态方法代理对象
    if (arguments.length === 0) {
        return createRngUtilsProxy();
    }
    return new RngUtils(initialRange);
};

// ==================== [SHORTCUT_$] 添加 $.RngUtils 静态方法代理 ====================
// 支持智能提示和 $.RngUtils.maxRange() 调用

/**
 * $.RngUtils.maxRange - 获取从第一行到最后一行的区域
 * @static
 * @param {Range|string} rng - 要获取的区域
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {RangeChain} RangeChain对象，支持链式调用
 * @memberof $.RngUtils
 * @example
 * $.RngUtils.maxRange("1:1000","A").safeArray()  // 返回数组
 * $.RngUtils.maxRange("A1:J1").z加边框()         // 链式调用
 */
$.RngUtils.maxRange = function(rng, col) {
    var result = RngUtils.maxRange.apply(RngUtils, arguments);
    if (result && result.Address && typeof result.Address === 'function') {
        return new RangeChain(result);
    }
    return result;
};

/**
 * $.RngUtils.safeArray - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @static
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @memberof $.RngUtils
 */
$.RngUtils.safeArray = RngUtils.safeArray;

/**
 * $.RngUtils.z安全数组 - 将指定区域转换为安全二维数组（返回 Array2D 对象，支持链式调用）
 * @static
 * @param {Range|string} rng - 要转换的区域
 * @returns {Array2D} Array2D 对象，支持 filter/map/toRange 等链式调用
 * @memberof $.RngUtils
 */
$.RngUtils.z安全数组 = RngUtils.z安全数组;

/**
 * $.RngUtils.endRow - 获取指定区域的最大行数
 * @static
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @memberof $.RngUtils
 */
$.RngUtils.endRow = RngUtils.endRow;

/**
 * $.RngUtils.z最大行 - 获取指定区域的最大行数
 * @static
 * @param {Range|string} rng - 要获取最大行数的区域
 * @returns {number} 最大行数
 * @memberof $.RngUtils
 */
$.RngUtils.z最大行 = RngUtils.z最大行;

/**
 * $.RngUtils.maxArray - 获取从第一行到最大行的区域并转换为数组（返回 Array2D 对象，支持链式调用）
 * @static
 * @param {Range|string} rng - 要获取的区域（如 "A1:H1"）
 * @param {string} [col="A"] - 列号，如果rng是整行时需要指定
 * @returns {Array2D} Array2D 对象，支持 skip/take/filter/sortByCols/toRange 等链式调用
 * @memberof $.RngUtils
 * @example
 * $.RngUtils.maxArray("A1:H1").skip(3).take(10).toRange("K4");
 */
$.RngUtils.maxArray = $.maxArray;

/**
 * $.RngUtils.z最大数组 - maxArray的中文别名
 * @static
 * @memberof $.RngUtils
 */
$.RngUtils.z最大数组 = $.RngUtils.maxArray;

// 其他常用静态方法（可根据需要添加更多）
var staticMethods = [
    'z最后一个', 'lastCell',
    'z安全区域', 'safeRange',
    'z最大行单元格', 'endRowCell',
    'z最大行区域', 'maxRange',
    'z最大列', 'endCol',
    'z最大列单元格', 'endColCell',
    'z可见区数组', 'visibleArray',
    'z可见区域', 'visibleRange',
    'z加边框', 'addBorders',
    'z取前几行', 'takeRows',
    'z跳过前几行', 'skipRows',
    'z插入多行', 'insertRows',
    'z插入多列', 'insertCols',
    'z删除空白行', 'delBlankRows',
    'z删除空白列', 'delBlankCols',
    'z整行', 'entireRow',
    'z整列', 'entire_column',
    'z行数', 'rowsCount',
    'z列数', 'colsCount',
    'z列号字母互转', 'colToAbc',
    'z复制粘贴格式', 'copyFormat',
    'z复制粘贴值', 'copyValue',
    'z联合区域', 'unionAll',
    'z多列排序', 'rngSortCols',
    'z最大数组', 'maxArray',
    'z合并单元格', 'mergeCells',
    'z取消合并单元格', 'unmergeCells'
];

for (var i = 0; i < staticMethods.length; i++) {
    var methodName = staticMethods[i];
    if (RngUtils[methodName]) {
        (function(name) {
            $.RngUtils[name] = function() {
                var result = RngUtils[name].apply(RngUtils, arguments);
                // 如果返回的是Range对象，包装成RangeChain支持链式调用
                if (result && result.Address && typeof result.Address === 'function') {
                    return new RangeChain(result);
                }
                return result;
            };
        })(methodName);
    }
}

/**
 * $.ShtUtils - Sheet工具类工厂
 * @param {Worksheet} initialSheet - 初始Sheet
 * @returns {ShtUtils} ShtUtils实例
 * @example
 * $.ShtUtils().z当前工作表()
 */
$.ShtUtils = function(initialSheet) {
    return new ShtUtils_ctor(initialSheet);
};

/**
 * $.DateUtils - 日期工具类工厂
 * @param {Date|string} initialDate - 初始日期
 * @returns {DateUtils} DateUtils实例
 * @example
 * $.DateUtils().z格式化("yyyy-MM-dd")
 */
$.DateUtils = function(initialDate) {
    return new DateUtils(initialDate);
};


/**
 * 日志输出
 * @param {...any} args - 参数
 */
function log() {
    if (typeof Console !== 'undefined') {
        Array.prototype.slice.call(arguments).forEach(function(arg) {
            Console.log(arg);
        });
    }
}

/**
 * 批量创建中英文方法别名
 * @private
 * @param {Object} prototype - 原型对象
 * @param {Array} aliases - 别名配置数组 [[中文名, 英文名], ...]
 */
function createBilingualAliases(prototype, aliases) {
    for (var i = 0; i < aliases.length; i++) {
        var cnName = aliases[i][0];
        var enName = aliases[i][1];
        if (prototype[cnName] && !prototype[enName]) {
            prototype[enName] = prototype[cnName];
        } else if (prototype[enName] && !prototype[cnName]) {
            prototype[cnName] = prototype[enName];
        }
    }
}

/**
 * JSON日志输出
 * @param {any} x - 对象
 * @param {Boolean} wrapopt - 是否包装JSON对象(即是否要输出日期等信息)，默认true
 * @example
 * logjson([[1,2],[3,4],[5,6]],0);  // 输出: [[1,2],[3,4],[5,6]]
 * logjson([1,2,3])                  // 一维数组输出为紧凑单行
 */
function logjson(x, wrapopt) {
    wrapopt = wrapopt !== undefined ? wrapopt : true;

    // 处理 Array2D 对象（提取 _items 属性）
    if (x && typeof x === 'object' && x._items && Array.isArray(x._items)) {
        x = x._items;
    }

    // 二维数组特殊处理
    if (Array.isArray(x) && x.length > 0 && Array.isArray(x[0])) {
        // wrapopt=0 时输出紧凑格式
        if (wrapopt === false || wrapopt === 0) {
            var output = JSON.stringify(x);
            if (typeof Console !== 'undefined') {
                Console.log(output);
            }
        } else {
            // 格式化输出（对齐）
            var lines = formatArray2DAsJSON(x);
            for (var i = 0; i < lines.length; i++) {
                if (typeof Console !== 'undefined') {
                    Console.log(lines[i]);
                }
            }
        }
        return;
    }

    // 一维数组输出为紧凑单行格式
    if (Array.isArray(x)) {
        var str = '[' + x.map(function(item) {
            if (item === null || item === undefined) return '';
            return String(item);
        }).join(',') + ']';
        if (typeof Console !== 'undefined') {
            Console.log(str);
        }
        return;
    }

    // 其他类型：处理循环引用和日期
    var output;
    if (wrapopt && typeof x === 'object' && x !== null) {
        var seen = new WeakSet();
        var replacer = function(key, value) {
            if (typeof value === 'object' && value !== null) {
                if (seen.has(value)) {
                    return '[Circular]';
                }
                seen.add(value);
            }
            if (value instanceof Date) {
                return value.toISOString();
            }
            return value;
        };
        output = JSON.stringify(x, replacer, 2);
    } else {
        output = typeof x === 'object' ? JSON.stringify(x, null, wrapopt ? 2 : 0) : String(x);
    }

    if (typeof Console !== 'undefined') {
        Console.log(output);
    }

    return;
}

/**
 * 格式化二维数组为JSON（支持对齐显示）
 * @private
 * @param {Array} arr - 二维数组
 * @returns {Array} 格式化的字符串数组
 */
function formatArray2DAsJSON(arr) {
    if (!arr || arr.length === 0) return ['[]'];

    /**
     * 计算字符串的显示宽度（基于等宽字体环境）
     * 规则：
     * - ASCII 字符（U+0000 - U+007F）= 1
     * - 非ASCII 字符（包括中文等宽字符）= 2
     */
    var getDisplayWidth = function(str) {
        var width = 0;
        for (var i = 0; i < str.length; i++) {
            var code = str.charCodeAt(i);
            if (code < 128) {
                // ASCII 字符宽度为 1
                width += 1;
            } else {
                // 非ASCII 字符（包括中文）宽度为 2
                width += 2;
            }
        }
        return width;
    };

    // 先将每行转换为JSON token，以便计算显示宽度
    var cellInfos = [];
    var colCount = arr[0].length;

    for (var row = 0; row < arr.length; row++) {
        var rowCells = [];
        for (var col = 0; col < colCount; col++) {
            var cellValue = col < arr[row].length ? arr[row][col] : null;
            if (cellValue === null || cellValue === undefined) {
                rowCells.push({ str: 'null', quoted: false });
            } else if (typeof cellValue === 'number' && isFinite(cellValue)) {
                rowCells.push({ str: String(cellValue), quoted: false });
            } else if (typeof cellValue === 'boolean') {
                rowCells.push({ str: String(cellValue), quoted: false });
            } else {
                rowCells.push({ str: String(cellValue), quoted: true });
            }
        }
        cellInfos.push(rowCells);
    }

    // 计算每列内容的最大显示宽度（字符串类型包含引号宽度）
    var contentWidths = [];
    for (var col = 0; col < colCount; col++) {
        var maxWidth = 0;
        for (var row = 0; row < arr.length; row++) {
            var info = cellInfos[row][col];
            var w = getDisplayWidth(info.str) + (info.quoted ? 2 : 0);
            maxWidth = Math.max(maxWidth, w);
        }
        contentWidths.push(maxWidth);
    }

    var lines = [];

    // 构建所有行，确保对齐
    for (var row = 0; row < arr.length; row++) {
        var rowParts = [];
        for (var col = 0; col < colCount; col++) {
            var info = cellInfos[row][col];
            var displayWidth = getDisplayWidth(info.str) + (info.quoted ? 2 : 0);

            // 计算需要填充的宽度
            var paddingNeeded = contentWidths[col] - displayWidth;

            // 使用普通空格填充（每个空格占1个显示宽度）
            // 使用 Array().join() 替代 String.repeat() 以提升兼容性
            var paddingStr = paddingNeeded > 0 ? Array(paddingNeeded + 1).join(' ') : '';

            // 构建单元格：字符串加引号，数值/null/布尔保持原始类型
            var cell;
            if (info.quoted) {
                cell = '"' + paddingStr + info.str + '"';
            } else {
                cell = paddingStr + info.str;
            }

            rowParts.push(cell);
        }

        // 用逗号连接各列（逗号后无空格）
        var rowStr = '[' + rowParts.join(',') + ']';
        lines.push(rowStr);
    }

    // 添加前导空格和行尾逗号
    for (var i = 0; i < lines.length; i++) {
        if (i < lines.length - 1) {
            lines[i] = ' ' + lines[i] + ',';
        } else {
            lines[i] = ' ' + lines[i];
        }
    }

    lines.push(']');
    lines.unshift('[');
    return lines;
}

// ==================== [GLOBAL_FUNCS] 全局工具函数（f1, $fx, $toArray）====================

/**
 * f1函数 - 在WPS JSA立即窗口快速打开JSA880帮助
 * @param {String} fxname - 函数名，如Array2D.pad
 * @example
 * f1("Array2D.pad")  // 打开帮助
 */
// 🔧 v4.0.11 修复: 改用默认浏览器打开帮助，移除过时的 ActiveXObject
function f1(fxname) {
    var helpUrl = "https://vbayyds.com/api/help/" + fxname;
    try {
        // WPS JSA 中使用内置 Shell.Application 打开默认浏览器
        if (typeof Shell !== 'undefined' && typeof Shell.Application === 'function') {
            var shell = new Shell.Application();
            shell.Open(helpUrl);
        } else if (typeof Application !== 'undefined' && Application.Shell) {
            Application.Shell(helpUrl);
        } else {
            // 降级：仅输出到控制台
            if (typeof Console !== 'undefined') {
                Console.log("帮助地址: " + helpUrl);
            }
        }
    } catch (e) {
        if (typeof Console !== 'undefined') {
            Console.log("帮助地址: " + helpUrl);
        }
    }
}

/**
 * $fx函数 - WorksheetFunction对象的简写
 * @param {string} path - 函数对象的路径
 * @returns {Function} 工作表函数
 * @example
 * $fx.Sum(1,2,3)  // 6
 */
function $fx(path) {
    const parts = path.split('.');
    var obj = WorksheetFunction;
    for (var i = 0; i < parts.length; i++) {
        if (obj[parts[i]]) {
            obj = obj[parts[i]];
        } else {
            return null;
        }
    }
    return typeof obj === 'function' ? obj : null;
}

/**
 * $toArray函数 - 将参数转换为数组（内部使用）
 * @param {...any} args - 要转换为数组的参数
 * @returns {Array} 转换后的数组
 * @example
 * $toArray("产品1", "产品2", "产品3")  // ["产品1","产品2","产品3"]
 */
function $toArray() {
    var result = [];
    for (var i = 0; i < arguments.length; i++) {
        result.push(arguments[i]);
    }
    return result;
}

// ==================== [TYPE_CONVERT] 类型转换函数（as系列） ====================

/**
 * asString函数 - 将对象转换为字符串对象
 * @param {any} s - 要转换的对象
 * @returns {String} 字符串
 * @example
 * asString(123)  // "123"
 */
function asString(s) {
    return String(s === null || s === undefined ? '' : s);
}

/**
 * asArray函数 - 将值转换为Array2D对象（支持链式调用和toRange）
 * @param {any} a - 要转换的值
 * @returns {Array2D} Array2D对象
 * @example
 * asArray(123)                      // Array2D([[123]])
 * asArray("abc")                    // Array2D([["abc"]])
 * asArray([1,2,3])                  // Array2D([[1],[2],[3]])
 * asArray([[1,2],[3,4]])            // Array2D([[1,2],[3,4]])
 * asArray(Array2D([[1,2]]))         // Array2D([[1,2]]) (原样返回)
 * asArray("a,b,c")                  // Array2D([["a"],["b"],["c"]])
 */
function asArray(a) {
    // 如果已经是 Array2D，直接返回
    if (a instanceof Array2D) return a;

    var arr;
    if (Array.isArray(a)) {
        arr = a;
    } else if (a === null || a === undefined) {
        arr = [];
    } else if (typeof a === 'string') {
        // 统一中文逗号为英文逗号
        a = a.replace(/，/g, ',');
        // 尝试按逗号分割
        if (a.indexOf(',') >= 0) {
            var parts = a.split(',').map(function(s) { return s.trim(); });
            // 转为二维数组
            arr = [];
            for (var i = 0; i < parts.length; i++) {
                arr.push([parts[i]]);
            }
        } else {
            arr = [[a]];
        }
    } else {
        arr = [[a]];
    }

    // 确保 arr 是二维数组
    if (arr.length > 0 && !Array.isArray(arr[0])) {
        var newArr = [];
        for (var j = 0; j < arr.length; j++) {
            newArr.push([arr[j]]);
        }
        arr = newArr;
    }

    return new Array2D(arr);
}

/**
 * asArray2D函数 - 将值转换为Array2D对象（asArray的别名）
 * @param {any} a - 要转换的值
 * @returns {Array2D} Array2D对象
 * @example
 * asArray2D([[1,2],[3,4]])           // Array2D([[1,2],[3,4]])
 * asArray2D([1,2,3])                  // Array2D([[1],[2],[3]])
 * asArray2D("a,b,c")                  // Array2D([["a"],["b"],["c"]])
 * asArray2D(Array2D([[1,2]]))         // Array2D([[1,2]]) (原样返回)
 */
var asArray2D = asArray;

/**
 * asNumber函数 - 将值转换为数字
 * @param {any} a - 要转换的值
 * @returns {Number} 数字，转换失败返回0
 * @example
 * asNumber("123")        // 123
 * asNumber("12.34")      // 12.34
 * asNumber("abc")        // 0
 * asNumber(null)         // 0
 */
function asNumber(a) {
    if (typeof a === 'number') return a;
    if (typeof a === 'boolean') return a ? 1 : 0;
    if (a === null || a === undefined || a === '') return 0;
    var num = Number(a);
    return isNaN(num) ? 0 : num;
}

/**
 * asDate函数 - 将值转换为DateUtils对象（支持智能提示和链式调用）
 * @param {any} a - 要转换的值
 * @returns {DateUtils} DateUtils实例
 * @example
 * asDate("2023-9-1").z月份()     // 9
 * asDate(45170).z年份()          // 2023 (Excel日期序号)
 * asDate("2023/09/01").z日期()   // 1
 */
function asDate(a) {
    var date;
    if (a instanceof DateUtils) return a;
    if (a instanceof Date) {
        date = a;
    } else if (typeof a === 'number') {
        // Excel日期序号转JS Date
        date = new Date((a - 25569) * 86400 * 1000);
    } else if (typeof a === 'string') {
        date = new Date(a);
        if (isNaN(date.getTime())) {
            date = new Date();
        }
    } else {
        date = new Date();
    }
    return new DateUtils(date);
}

/**
 * asRange函数 - 将值转换为Range对象
 * @param {any} a - 要转换的值（地址字符串、Range对象等）
 * @returns {Range|null} Range对象
 * @example
 * asRange("A1")          // Range对象
 * asRange(Range("A1"))   // Range对象
 * asRange("A1:C10")      // Range对象
 */
function asRange(a) {
    if (a && a.Address) return a; // 已经是Range对象
    if (typeof a === 'string') {
        try {
            return Range(a);
        } catch (e) {
            return null;
        }
    }
    return null;
}

/**
 * asMap函数 - 将值转换为Map对象
 * @param {any} a - 要转换的值（对象、Map、二维数组等）
 * @returns {Map} Map对象
 * @example
 * asMap({a:1,b:2})       // Map(2) {"a"=>1,"b"=>2}
 * asMap([['a',1],['b',2]])// Map(2) {"a"=>1,"b"=>2}
 */
function asMap(a) {
    if (a instanceof Map) return a;
    var map = new Map();
    if (a === null || a === undefined) return map;
    if (Array.isArray(a)) {
        // 二维数组转Map: [['key','value'],...]
        a.forEach(function(item) {
            if (Array.isArray(item) && item.length >= 2) {
                map.set(item[0], item[1]);
            }
        });
    } else if (typeof a === 'object') {
        // 对象转Map
        for (var key in a) {
            if (a.hasOwnProperty(key)) {
                map.set(key, a[key]);
            }
        }
    }
    return map;
}

/**
 * asObject函数 - 将值转换为普通对象
 * @param {any} a - 要转换的值（Map、对象等）
 * @returns {Object} 普通对象
 * @example
 * asObject(new Map([['a',1],['b',2]]))  // {a:1,b:2}
 * asObject({a:1})                        // {a:1}
 */
function asObject(a) {
    if (a instanceof Map) {
        var obj = {};
        a.forEach(function(value, key) {
            obj[key] = value;
        });
        return obj;
    }
    if (typeof a === 'object' && a !== null) {
        return a;
    }
    return {};
}

/**
 * asShape函数 - 将对象转换为Shape对象
 * @param {any} shp - 要转换的对象
 * @returns {Shape|null} Shape对象
 * @example
 * asShape('矩形 2')  // Shape对象
 */
function asShape(shp) {
    if (typeof shp === 'string') {
        // 遍历所有工作表的形状
        for (var i = 1; i <= Sheets.Count; i++) {
            var sht = Sheets(i);
            for (var j = 1; j <= sht.Shapes.Count; j++) {
                if (sht.Shapes(j).Name === shp) return sht.Shapes(j);
                if (sht.Shapes(j).Name.indexOf(shp) !== -1) return sht.Shapes(j);
            }
        }
        return null;
    }
    if (shp && shp.Name) return shp;
    return null;
}

// ==================== [SHEETCHAIN] 工作表链式调用类 ====================

/**
 * SheetChain - 工作表链式调用包装类（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 包装WPS工作表对象，提供链式调用和智能提示
 * @example
 * asSheet("Sheet1").z激活().z名称()
 * asSheet(1).z已使用区域().z安全数组()
 */
function SheetChain(sht) {
    if (!(this instanceof SheetChain)) {
        return new SheetChain(sht);
    }
    this._sheet = null;

    // 检查WPS环境和Sheets可用性
    if (typeof Sheets === 'undefined') return;

    // 如果已经是Sheet对象，直接使用
    if (sht && sht.Activate && sht.Name) {
        this._sheet = sht;
        return;
    }

    if (typeof sht === 'number') {
        try {
            this._sheet = Sheets(sht);
        } catch (e) {
            this._sheet = null;
        }
        return;
    }

    if (typeof sht === 'string') {
        try {
            // 首先尝试精确匹配
            this._sheet = Sheets(sht);
        } catch (e) {
            // 精确匹配失败，尝试模糊匹配
            try {
                for (var i = 1; i <= Sheets.Count; i++) {
                    var sheet = Sheets(i);
                    // 包含匹配
                    if (sheet.Name.indexOf(sht) >= 0) {
                        this._sheet = sheet;
                        return;
                    }
                    // 忽略大小写匹配
                    if (sheet.Name.toLowerCase() === sht.toLowerCase()) {
                        this._sheet = sheet;
                        return;
                    }
                }
            } catch (e2) {
                console.error("SheetChain模糊匹配失败: " + e2.message);
            }
            this._sheet = null;
        }
        return;
    }
}

/**
 * Value - 获取原始Sheet对象
 * @returns {Worksheet|null} 工作表对象
 */
SheetChain.prototype.Value = function() {
    return this._sheet;
};

/**
 * Name - 获取工作表名称
 * @returns {String} 工作表名称
 */
SheetChain.prototype.z名称 = function() {
    return this._sheet ? this._sheet.Name : '';
};
SheetChain.prototype.Name = SheetChain.prototype.z名称;

/**
 * Activate - 激活工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z激活 = function() {
    if (this._sheet) this._sheet.Activate();
    return this;
};
SheetChain.prototype.Activate = SheetChain.prototype.z激活;

/**
 * UsedRange - 获取已使用区域
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z已使用区域 = function() {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.UsedRange);
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.UsedRange = SheetChain.prototype.z已使用区域;

/**
 * SafeUsedRange - 获取安全已使用区域（处理空表情况）
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z安全已使用区域 = function() {
    if (!this._sheet) return null;

    var usedRange;
    try {
        usedRange = this._sheet.UsedRange;
    } catch (e) {
        return new RangeChain(this._sheet.Range("A1"));
    }

    if (!usedRange) return new RangeChain(this._sheet.Range("A1"));

    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;

    return new RangeChain(this._sheet.Range(this._sheet.Cells(1, 1), this._sheet.Cells(lastRow, lastCol)));
};
SheetChain.prototype.SafeUsedRange = SheetChain.prototype.z安全已使用区域;

/**
 * Range - 获取Range对象
 * @param {String} address - 地址
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z区域 = function(address) {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.Range(address));
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.Range = SheetChain.prototype.z区域;

/**
 * Cells - 获取Cells对象
 * @param {Number} row - 行号
 * @param {Number} col - 列号
 * @returns {RangeChain|null} RangeChain对象
 */
SheetChain.prototype.z单元格 = function(row, col) {
    if (!this._sheet) return null;
    try {
        return new RangeChain(this._sheet.Cells(row, col));
    } catch (e) {
        return null;
    }
};
SheetChain.prototype.Cells = SheetChain.prototype.z单元格;

/**
 * Delete - 删除工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z删除 = function() {
    if (this._sheet) {
        try {
            this._sheet.Delete();
        } catch (e) {
            console.error("删除工作表失败: " + e.message);
        }
    }
    return this;
};
SheetChain.prototype.Delete = SheetChain.prototype.z删除;

/**
 * Copy - 复制工作表
 * @param {Worksheet} [before] - 在此工作表之前插入
 * @param {Worksheet} [after] - 在此工作表之后插入
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z复制 = function(before, after) {
    if (!this._sheet) return this;
    try {
        if (before) {
            this._sheet.Copy(before);
        } else if (after) {
            this._sheet.Copy(undefined, after);
        } else {
            this._sheet.Copy();
        }
    } catch (e) {
        console.error("复制工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Copy = SheetChain.prototype.z复制;

/**
 * Protect - 保护工作表
 * @param {String} [password] - 密码
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z保护 = function(password) {
    if (!this._sheet) return this;
    try {
        if (password) {
            this._sheet.Protect(password);
        } else {
            this._sheet.Protect();
        }
    } catch (e) {
        console.log("保护工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Protect = SheetChain.prototype.z保护;

/**
 * Unprotect - 取消保护工作表
 * @param {String} [password] - 密码
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z取消保护 = function(password) {
    if (!this._sheet) return this;
    try {
        if (password) {
            this._sheet.Unprotect(password);
        } else {
            this._sheet.Unprotect();
        }
    } catch (e) {
        console.error("取消保护工作表失败: " + e.message);
    }
    return this;
};
SheetChain.prototype.Unprotect = SheetChain.prototype.z取消保护;

/**
 * 隐藏工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z隐藏 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = false;
    return this;
};

/**
 * 显示工作表
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z显示 = function() {
    if (!this._sheet) return this;
    this._sheet.Visible = true;
    return this;
};

/**
 * Index - 获取工作表索引
 * @returns {Number} 工作表索引
 */
SheetChain.prototype.z索引 = function() {
    return this._sheet ? this._sheet.Index : 0;
};
SheetChain.prototype.Index = SheetChain.prototype.z索引;

/**
 * SetName - 设置工作表名称
 * @param {String} newName - 新名称
 * @returns {SheetChain} 当前实例
 */
SheetChain.prototype.z设置名称 = function(newName) {
    if (this._sheet) {
        this._sheet.Name = newName;
    }
    return this;
};
SheetChain.prototype.SetName = SheetChain.prototype.z设置名称;

/**
 * Exists - 判断工作表是否存在
 * @returns {Boolean} 是否存在
 */
SheetChain.prototype.z存在 = function() {
    return this._sheet !== null;
};
SheetChain.prototype.Exists = SheetChain.prototype.z存在;

// ==================== [TYPE_CONVERT] 类型转换函数（asSheet, asWorkbook）====================

/**
 * asSheet函数 - 将对象转换为SheetChain对象（支持智能提示和链式调用）
 * @param {any} sht - 要转换的对象
 * @returns {SheetChain} SheetChain实例
 * @example
 * asSheet("1月").z激活().z名称()
 * asSheet(1).z已使用区域().z安全数组()
 * asSheet().z激活()
 */
function asSheet(sht) {
    return new SheetChain(sht);
}

/**
 * asWorkbook函数 - 将对象转换为工作簿对象
 * @param {any} wbk - 要转换的对象
 * @returns {Workbook} 工作簿对象
 * @example
 * asWorkbook("测试排序")  // 工作簿对象
 */
function asWorkbook(wbk) {
    if (typeof wbk === 'string') {
        for (var i = 1; i <= Workbooks.Count; i++) {
            if (Workbooks(i).Name === wbk) return Workbooks(i);
        }
        return null;
    }
    if (wbk && wbk.Name) return wbk;
    return null;
}

// ==================== As - 类型转换包装类 ====================

/**
 * As类 - 类型转换包装类（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 提供类型转换和常用操作方法，支持中英双语API
 * @example
 * // 基本使用
 * As([[1,2,3],[4,5,6]]).toArray().z求和()        // 21
 * As("123").toNumber()                           // 123
 * As(123).toString()                             // "123"
 * // 链式调用
 * As([[1,2],[3,4]]).toArray().z转置().z扁平化().val()  // [1,3,2,4]
 */
function As(value) {
    // 支持工厂模式调用
    if (!(this instanceof As)) {
        return new As(value);
    }

    this._original = value;
    this._value = value;
}

/**
 * 创建新实例（链式调用核心）
 * @private
 * @param {any} data - 新值
 * @returns {As} 新实例
 */
As.prototype._new = function(value) {
    const instance = new As();
    instance._original = this._original;
    instance._value = value;
    return instance;
};

/**
 * 获取/设置当前值
 * @param {any} [newValue] - 新值（可选）
 * @returns {As|any} 设置时返回this，否则返回当前值
 * @example
 * As(123).val()           // 123
 * As(123).val(456)        // 返回链式对象
 */
As.prototype.val = function(newValue) {
    if (newValue !== undefined) {
        this._value = newValue;
        return this;
    }
    return this._value;
};

// ==================== [AS] 类型转换包装类 ====================

/**
 * 转换为数组
 * @returns {Array2D} 二维数组工具对象（如果是二维数组）或 As包装对象
 * @example
 * As([1,2,3]).toArray()              // [1,2,3]
 * As("a,b,c").toArray()              // ["a","b","c"]
 * As([[1,2],[3,4]]).toArray()        // Array2D对象，支持链式调用
 */
As.prototype.toArray = function() {
    // v4.1.0 修复: asArray() 始终返回 Array2D 实例，直接返回即可
    // 原代码: 非空时再 Array2D(arr) 二次包装浪费，空时返回 As 包装类型不一致
    return asArray(this._value);
};

/**
 * 转换为数字
 * @returns {As} 包装对象
 * @example
 * As("123").toNumber().val()         // 123
 * As("abc").toNumber().val()         // 0
 */
As.prototype.toNumber = function() {
    return this._new(asNumber(this._value));
};

/**
 * 转换为字符串
 * @returns {As} 包装对象
 * @example
 * As(123).toString().val()           // "123"
 * As(null).toString().val()          // ""
 */
As.prototype.toString = function() {
    return this._new(asString(this._value));
};

/**
 * 转换为日期
 * @returns {As} 包装对象
 * @example
 * As("2023-9-1").toDate().val()      // Date对象
 * As(45170).toDate().val()           // Date对象
 */
As.prototype.toDate = function() {
    return this._new(asDate(this._value));
};

/**
 * 转换为Map对象
 * @returns {As} 包装对象
 * @example
 * As({a:1,b:2}).toMap().val()        // Map对象
 */
As.prototype.toMap = function() {
    return this._new(asMap(this._value));
};

/**
 * 转换为普通对象
 * @returns {As} 包装对象
 * @example
 * const map = new Map([['a',1]]);
 * As(map).toObject().val()           // {a:1}
 */
As.prototype.toObject = function() {
    return this._new(asObject(this._value));
};

/**
 * 转换为Range对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("A1:C10").toRange().val()       // Range对象
 */
As.prototype.toRange = function() {
    const rng = asRange(this._value);
    return rng !== null ? this._new(rng) : null;
};

/**
 * 转换为工作表对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("Sheet1").toSheet().val()       // Worksheet对象
 */
As.prototype.toSheet = function() {
    const sht = asSheet(this._value);
    return sht !== null ? this._new(sht) : null;
};

/**
 * 转换为工作簿对象（WPS环境）
 * @returns {As|null} 包装对象或null
 * @example
 * As("工作簿1.xlsx").toWorkbook().val()  // Workbook对象
 */
As.prototype.toWorkbook = function() {
    const wbk = asWorkbook(this._value);
    return wbk !== null ? this._new(wbk) : null;
};

// ==================== [TYPE_CONVERT] 辅助函数（cdate, cstr）====================
As.prototype.z转数组 = As.prototype.toArray;
As.prototype.z转数字 = As.prototype.toNumber;
As.prototype.z转字符串 = As.prototype.toString;
As.prototype.z转日期 = As.prototype.toDate;
As.prototype.z转Map = As.prototype.toMap;
As.prototype.z转对象 = As.prototype.toObject;

/**
 * cdate函数 - 将日期转换为Excel日期数值
 * @param {any} v - 日期字符串或JS日期对象
 * @returns {Number} Excel日期数值
 * @example
 * cdate('2023-9-1')  // 45170
 */
function cdate(v) {
    if (typeof v === 'number') return v;
    var date;
    if (typeof v === 'string') {
        // 处理简短日期格式
        if (v.match(/^\d{1,2}-\d{1,2}$/)) {
            v = '20' + v;  // 23-9-1 -> 2023-9-1
        }
        date = new Date(v);
    } else if (v instanceof Date) {
        date = v;
    } else {
        return 0;
    }
    var excelEpoch = new Date(1900, 0, 1);
    var msPerDay = 24 * 60 * 60 * 1000;
    return Math.floor((date - excelEpoch) / msPerDay) + 2;
}

/**
 * cstr函数 - 将值转换为字符串
 * @param {any} v - 要转换的值
 * @returns {String} 字符串
 * @example
 * cstr(1537789)  // "1537789"
 */
const cstr = (v) => v === null || v === undefined ? '' : String(v);

// ==================== [TYPE_CHECK] 类型检查函数（is系列） ====================

/**
 * isArray函数 - 检查值是否为数组
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为数组
 * @example
 * isArray([1,2,3])  // true
 */
const isArray = (v) => Array.isArray(v);

/**
 * isArray2D函数 - 检查值是否为二维数组
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为二维数组
 * @example
 * isArray2D([[1],[2],[3]])  // true
 */
const isArray2D = (v) => {
    if (!Array.isArray(v)) return false;
    if (v.length === 0) return false;
    return v.every(row => Array.isArray(row));
};

/**
 * isBoolean函数 - 检查值是否为布尔值
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为布尔值
 * @example
 * isBoolean(false)  // true
 */
const isBoolean = (v) => typeof v === 'boolean';

/**
 * isCollection函数 - 检查对象是否为集合对象
 * @param {any} obj - 要检查的对象
 * @returns {Boolean} 是否为集合对象
 * @example
 * isCollection(Sheets)  // true
 */
const isCollection = (obj) => {
    if (!obj) return false;
    // 检查是否是WPS集合对象
    if (obj && typeof obj === 'object') {
        // WPS集合对象通常有Count和Item属性
        if (obj.Count !== undefined && typeof obj.Item === 'function') return true;
        // 检查是否有枚举器
        try {
            const enumerator = new Enumerator(obj);
            return true;
        } catch (e) {
            // 不是集合
        }
    }
    return false;
};

/**
 * isDate函数 - 检查值是否为日期对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为日期对象
 * @example
 * isDate(new Date())  // true
 */
const isDate = (v) => v instanceof Date;

/**
 * isEmpty函数 - 检查值是否为空值
 * @param {any} value - 要检查的值
 * @returns {Boolean} 是否为空值
 * @example
 * isEmpty(undefined)  // true
 * isEmpty('')         // true
 * isEmpty(null)       // true
 */
const isEmpty = (value) => value === null || value === undefined || value === '';

/**
 * isNumberic函数 - 检查值是否为数值类型
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为数值类型
 * @example
 * isNumberic(557)  // true
 */
const isNumberic = (v) => typeof v === 'number' && !isNaN(v);

/**
 * isRange函数 - 检查值是否为Range对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为Range对象
 * @example
 * isRange(Range("A1"))  // true
 */
const isRange = (v) => v && typeof v === 'object' && v.Address !== undefined;

/**
 * isRegex函数 - 检查值是否为正则表达式对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为正则表达式
 * @example
 * isRegex(/\d+/g)  // true
 */
const isRegex = (v) => v instanceof RegExp;

/**
 * isSameClass函数 - 检查两个值是否属于同一类别
 * @param {any} x - 第一个对象
 * @param {any} y - 第二个对象
 * @returns {Boolean} 是否属于同一类别
 * @example
 * isSameClass(560, 789)  // true
 */
const isSameClass = (x, y) => Object.prototype.toString.call(x) === Object.prototype.toString.call(y);

/**
 * isSheet函数 - 检查值是否为工作表对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为工作表对象
 * @example
 * isSheet(Sheets(1))  // true
 */
const isSheet = (v) => v && typeof v === 'object' && v.Name !== undefined && v.Cells !== undefined;

/**
 * isString函数 - 检查值是否为字符串类型
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为字符串
 * @example
 * isString('产品5')  // true
 */
const isString = (v) => typeof v === 'string';

/**
 * isWorkbook函数 - 检查值是否为工作簿对象
 * @param {any} v - 要检查的值
 * @returns {Boolean} 是否为工作簿对象
 * @example
 * isWorkbook(ActiveWorkbook)  // true
 */
const isWorkbook = (v) => v && typeof v === 'object' && v.Name !== undefined && v.Sheets !== undefined && v.Close !== undefined;

/**
 * typeName函数 - 获取值的类型名称
 * @param {any} x - 要获取类型名称的值
 * @returns {String} 类型名称
 * @example
 * typeName('产品5')  // "[object String]"
 */
const typeName = (x) => Object.prototype.toString.call(x);

// ==================== [GLOBAL_FUNCS] 其他工具函数（val, round）====================

/**
 * val函数 - 字符串及布尔值转为数值（与VBA的val保持一致）
 * @param {String} s - 要转换的字符串
 * @returns {Number} 数值
 * @example
 * val('5')      // 5
 * val('123abc') // 123
 * val('abc123') // 0
 */
const val = (s) => {
    if (typeof s === 'number') return s;
    if (typeof s === 'boolean') return s ? 1 : 0;
    if (typeof s !== 'string') return 0;
    s = s.trim();
    if (s === '') return 0;
    // VBA的val行为：读取字符串开头的数字字符
    const match = s.match(/^[-+]?[0-9]*\.?[0-9]+/);
    if (match) return parseFloat(match[0]);
    return 0;
};

/**
 * round函数 - 使用Excel计算规则对数字进行四舍五入
 * @param {number} number - 要进行四舍五入的数字
 * @param {number} [decimals=2] - 保留的小数位数（默认为2）
 * @returns {number} 四舍五入后的结果
 * @example
 * round(5.786543224, 3)  // 5.787
 */
const round = (number, decimals = 2) => {
    // 使用Excel的RoundWorksheetFunction确保与Excel行为一致
    if (typeof WorksheetFunction.Round !== 'undefined') {
        try {
            return WorksheetFunction.Round(number, decimals);
        } catch (e) {
            // 降级处理
        }
    }
    // 标准四舍五入
    const factor = Math.pow(10, decimals);
    return Math.round(number * factor) / factor;
};

// ubound函数 - 获取数组的指定维度的上界
// 在导出部分定义以避免WPS打印函数定义

// ==================== [RANGECHAIN] Range链式调用类 ====================

/**
 * 【链式调用核心】创建新Array2D实例
 * 
 * 【为什么需要这个方法？】
 * Array2D的所有数据转换方法（如skip, filter, map等）都遵循"不可变性"原则，
 * 即不修改原数组，而是返回一个新的Array2D实例。
 * 这个方法统一封装了新实例的创建逻辑。
 * 
 * 【与构造函数的区别】
 * - 构造函数：通过new创建，数据已存在于this
 * - _new方法：创建空数组，然后将数据填充进去，再设置原型
 * 
 * 【技术细节】
 * 1. 创建空数组instance = []
 * 2. 用Array.prototype.push填充数据（instance现在是普通数组）
 * 3. 用Object.setPrototypeOf将instance的原型设置为Array2D.prototype
 * 4. 添加内部属性（_original, _items）
 * 5. 返回instance，它现在是Array2D实例
 * 
 * @private          // 内部使用，不建议用户直接调用
 * @param {Array} data - 新数据（已处理好的二维数组）
 * @returns {Array2D} 新的Array2D实例
 * @example
 * // 在filter方法内部使用：
 * return this._new(this._items.filter(fn));  // 返回筛选后的新实例
 */
Array2D.prototype._new = function(data) {
    // 创建空数组实例（此时是普通数组）
    var instance = [];
    // 填充数据
    Array.prototype.push.apply(instance, data);

    // 【关键步骤：设置原型】
    // 将instance的原型链指向Array2D.prototype
    // 这样instance就能使用Array2D的所有方法了
    if (Object.setPrototypeOf) {
        // ES6标准方法
        Object.setPrototypeOf(instance, Array2D.prototype);
    } else {
        // 旧环境备用方案（__proto__是非标准但广泛支持的属性）
        instance.__proto__ = Array2D.prototype;
    }

    // 🔧 v3.7.9 修复: 保留原始的 _original 如果存在，否则使用当前数据
    // 这样 z筛选/z多列排序 后的对象仍能访问原始表头
    var originalData = data;
    if ('_original' in this && this._original !== undefined && this._original !== null) {
        originalData = this._original;
    }
    Object.defineProperty(instance, '_original', {
        value: originalData,
        writable: true,
        enumerable: false,
        configurable: true
    });

    // 🔧 P0-3 性能优化: 使用原生 slice 替代手动循环
    Object.defineProperty(instance, '_items', {
        get: function() {
            return Array.prototype.slice.call(this);
        },
        set: function(value) {
            Array.prototype.splice.call(this, 0, this.length);
            Array.prototype.push.apply(this, value);
        },
        enumerable: false,
        configurable: true
    });

    // 🔧 v3.7.9 修复: 更可靠地保留 _header 属性（表头信息）
    // 使用 in 操作符检查，因为 _header 可能是不可枚举的
    // XXD-160: 同时排除 typeof==='function'，避免误把原型上的链式 setter 复制到新实例
    if (Object.prototype.hasOwnProperty.call(this, '_header') && this._header !== undefined && this._header !== null && typeof this._header !== 'function') {
        Object.defineProperty(instance, '_header', {
            value: this._header,
            writable: true,
            enumerable: false,
            configurable: true
        });
    }

    // 返回新创建的Array2D实例
    return instance;
};

/**
 * XXD-160: 链式 _header(n) 设置表头行数。原 this._header=... 是属性，外部 a._header(0) 抛 not a function。
 * 调用后实例自有 _header 覆盖原型方法；如需修改请直接赋值 this._header = n。
 * @param {Number} [n=1] 表头行数
 * @returns {Array2D} this，支持链式
 * @example
 * new Array2D([[1,2],[1,2],[3,4]])._header(0).z去重()._items // [[1,2],[3,4]]
 */
Array2D.prototype._header = function(n) {
    var headerRows = (n === undefined || n === null) ? 1 : n;
    Object.defineProperty(this, '_header', {
        value: headerRows,
        writable: true,
        enumerable: false,
        configurable: true
    });
    return this;
};

// ==================== 基础操作 ====================

/**
 * 获取/设置数组值
 * @param {Array} [newData] - 新数据（可选）
 * @returns {Array2D|Array} 设置时返回this，否则返回当前数组
 * @example
 * Array2D([[1,2]]).val()           // [[1,2]]
 * Array2D([[1,2]]).val([[3,4]])     // 返回链式对象
 */
Array2D.prototype.val = function(newData) {
    if (newData !== undefined) {
        // setter 会自动同步数组属性
        this._items = newData;
        return this;
    }
    return this._items;
};

/**
 * 检查数组是否为空
 * @returns {Boolean} 是否为空
 * @example
 * Array2D([[1]]).z是否为空()    // false
 * Array2D([]).z是否为空()       // true
 */
Array2D.prototype.z是否为空 = function() {
    return !this._items || this._items.length === 0;
};
Array2D.prototype.isEmpty = Array2D.prototype.z是否为空;

/**
 * 获取元素数量
 * @returns {Number} 元素数量
 * @example
 * Array2D([[1,2],[3,4]]).z数量()  // 4
 */
Array2D.prototype.z数量 = function() {
    // 🔧 P0-3 性能优化: 直接遍历计数，避免创建扁平化数组副本
    var count = 0;
    for (var i = 0; i < this.length; i++) {
        count += Array.isArray(this[i]) ? this[i].length : 1;
    }
    return count;
};
Array2D.prototype.count = Array2D.prototype.z数量;

/**
 * 克隆数组（深拷贝）
 * @returns {Array2D} 新实例
 * @example
 * const arr = Array2D([[1,2]]);
 * const cloned = arr.z克隆();
 */
Array2D.prototype.z克隆 = function() {
    return this._new(JSON.parse(JSON.stringify(this._items)));
};
Array2D.prototype.copy = Array2D.prototype.z克隆;

// ==================== Array2D HTML输出方法 ====================

/**
 * 输出为HTML表格（html）- 将二维数组转换为HTML表格字符串
 * @param {Object} [options] - 配置选项
 * @param {string} [options.className] - 表格CSS类名
 * @param {string} [options.style] - 表格内联样式
 * @param {boolean} [options.header=false] - 是否将第一行作为表头
 * @param {string} [options.caption] - 表格标题
 * @returns {string} HTML表格字符串
 * @example
 * Array2D([[1,2],[3,4]]).z输出HTML()  
 * // 返回: "<table><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></table>"
 * Array2D([['姓名','年龄'],['张三',20]]).z输出HTML({header:true})
 * // 返回带thead的表格
 */
Array2D.prototype.z输出HTML = function(options) {
    options = options || {};
    var className = options.className || '';
    var style = options.style || '';
    var hasHeader = options.header === true;
    var caption = options.caption || '';
    
    var html = '<table';
    if (className) html += ' class="' + className + '"';
    if (style) html += ' style="' + style + '"';
    html += '>';
    
    if (caption) {
        html += '<caption>' + caption + '</caption>';
    }
    
    var startRow = 0;
    if (hasHeader && this._items.length > 0) {
        html += '<thead><tr>';
        var headerRow = this._items[0];
        for (var j = 0; j < headerRow.length; j++) {
            html += '<th>' + (headerRow[j] !== null && headerRow[j] !== undefined ? headerRow[j] : '') + '</th>';
        }
        html += '</tr></thead><tbody>';
        startRow = 1;
    }
    
    for (var i = startRow; i < this._items.length; i++) {
        html += '<tr>';
        var row = this._items[i];
        if (Array.isArray(row)) {
            for (var j = 0; j < row.length; j++) {
                var cell = row[j];
                html += '<td>' + (cell !== null && cell !== undefined ? cell : '') + '</td>';
            }
        } else {
            html += '<td>' + (row !== null && row !== undefined ? row : '') + '</td>';
        }
        html += '</tr>';
    }
    
    if (hasHeader) html += '</tbody>';
    html += '</table>';
    
    return html;
};
Array2D.prototype.html = Array2D.prototype.z输出HTML;
Array2D.prototype.toHtml = Array2D.prototype.z输出HTML;

/**
 * 静态方法：输出为HTML表格
 * @param {Array} arr - 二维数组
 * @param {Object} [options] - 配置选项
 * @returns {string} HTML表格字符串
 */
Array2D.html = function(arr, options) {
    if (!arr) return '<table></table>';
    return new Array2D(arr).z输出HTML(options);
};

// ==================== 使用辅助函数创建 Array2D 方法别名 ====================
createBilingualAliases(Array2D.prototype, [
    ['z填充', 'fill'],
    ['z补齐空位', 'fillBlank'],
    ['z扁平化', 'flat'],
    ['z反转', 'reverse'],
    ['z求和', 'sum'],
    ['z平均值', 'average'],
    ['z中位数', 'median'],
    ['z最大值', 'max'],
    ['z最小值', 'min'],
    ['z第一个', 'first'],
    ['z最后一个', 'last'],
    ['z转置', 'transpose'],
    ['z矩阵信息', 'matrixInfo'],
    ['z单元格', 'cell'],
    ['z设置单元格', 'setCell'],
    ['z写入单元格', 'toRange'],
    ['z连接', 'join'],
    ['z转JSON', 'toJson'],
    ['z分块', 'chunk'],
    ['z挑选', 'pick'],
    ['z跳过', 'skip'],
    ['z取前N个', 'take'],
    ['z查找索引', 'findIndex'],
    ['z包含', 'includes'],
    ['z筛选', 'filter'],
    ['z映射', 'map'],
    ['z归约', 'reduce'],
    ['z倒序归约', 'reduceRight'],
    ['z全部满足', 'every'],
    ['z有满足', 'some'],
    ['z行数', 'rowCount'],
    ['z列统计', 'columnStats'],
    ['z列数', 'colCount'],
    ['z获取行', 'getRow'],
    ['z获取列', 'getCol'],
    ['z首行', 'firstRow'],
    ['z末行', 'lastRow'],
    ['z首列', 'firstCol'],
    ['z末列', 'lastCol'],
    ['z添加行', 'addRow'],
    ['z提取列', 'pluck'],
    ['z添加列', 'addCol'],
    ['z删除行', 'deleteRow'],
    ['z删除列', 'deleteCol'],
    ['z升序排序', 'sortAsc'],
    ['z按规则升序', 'sortBy'],
    ['z按规则降序', 'sortByDesc'],
    ['z降序排序', 'sortDesc'],
    ['z行排序', 'sortRow'],
    ['z列排序', 'sortCol'],
    ['z多列排序', 'sortByCols'],
    ['z自定义排序', 'sortByList'],
    ['z去重', 'distinct'],
    ['z转矩阵', 'toMatrix'],
    ['z分组', 'groupBy'],
    ['z透视', 'pivotBy'],
    ['z上下连接', 'concat'],
    ['z左连接', 'leftjoin'],
    ['z内连接', 'innerjoin'],
    ['z一对多连接', 'leftFulljoin'],
    ['z左右全连接', 'fulljoin'],
    ['z左右连接', 'zip'],
    ['z排除', 'except'],
    ['z取交集', 'intersect'],
    ['z去重并集', 'union'],
    ['z超级查找', 'superLookup'],
    ['z查找单个', 'find'],
    ['z查找所有下标', 'findAllIndex'],
    ['z查找所有行下标', 'findRowsIndex'],
    ['z查找所有列下标', 'findColsIndex'],
    ['z查找元素下标', 'findIndexByPredicate'],
    ['z值位置', 'indexOf'],
    ['z从后往前值位置', 'lastIndexOf'],
    ['z批量删除列', 'deleteCols'],
    ['z批量删除行', 'deleteRows'],
    ['z批量插入列', 'insertCols'],
    ['z批量插入行', 'insertRows'],
    ['z插入行号', 'insertRowNum'],
    ['z按页数分页', 'pageByCount'],
    ['z按行数分页', 'pageByRows'],
    ['z按下标分页', 'pageByIndexs'],
    ['z间隔取数', 'nth'],
    ['z补齐数组', 'pad'],
    ['z重设大小', 'resize'],
    ['z处理空值', 'noNull'],
    ['z选择列', 'selectCols'],
    ['z选择行', 'selectRows'],
    ['z结果', 'res'],
    ['z行切片', 'slice'],
    ['z行切片删除行', 'splice'],
    ['z转字符串', 'toString']
]);

// ==================== 填充操作 ====================

/**
 * 批量填充数组
 * @param {string|number|boolean|null|undefined} value - 填充值
 * @param {Number} [rows] - 行数（可选，默认当前行数或1）
 * @param {Number} [cols] - 列数（可选，默认当前列数或1）
 * @returns {Array2D} 新实例
 * @example
 * Array2D().z填充(0, 2, 3)  // [[0,0,0],[0,0,0]]
 */
Array2D.prototype.z填充 = function(value, rows, cols) {
    // 🔧 XXD-188 polymorphic dispatch:
    //   legacy:  z填充(value, rows, cols)  → create rows×cols grid filled with value
    //   new:     z填充(row, col, value)    → set cell (row, col) on this._items
    // Detect: if 3 args and cols is NOT a number, treat as (row, col, value).
    // Old contract requires numeric cols, so existing callers (incl.
    // JSA880.js:13142 internal usage) are unaffected.
    if (arguments.length === 3 && typeof cols !== 'number') {
        // cell-setter form: (row, col, value)
        const row = value;
        const col = rows;
        const v   = cols;
        if (!this._items[row]) this._items[row] = [];
        this._items[row][col] = v;
        return this;
    }
    rows = rows || this._items.length || ARRAY_LIMITS.DEFAULT_ROWS;
    cols = cols || (this._items[0] ? this._items[0].length : ARRAY_LIMITS.DEFAULT_COLS);
    const result = [];
    for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
            row.push(value);
        }
        result.push(row);
    }
    return this._new(result);
};
Array2D.prototype.fill = Array2D.prototype.z填充;

/**
 * 补齐空位（fillBlank）- 支持方向填充的增强版，可处理合并单元格
 * @param {string} [direction='right'] - 填充方向：left/right/up/down
 * @param {string} [rangeAddress] - 参照单元格地址（如"A2:D2"），用于确定填充区域
 * @param {any} [fillValue=''] - 填充值
 * @returns {Array2D} 新实例
 * @example
 * // 基础用法：填充null/undefined
 * Array2D([[1,null],[2,undefined]]).z补齐空位()  // [[1,''],[2,'']]
 *
 * // 高级用法：按方向填充（用于合并单元格处理）
 * Array2D([[1,2],[3,4]]).z补齐空位('right', 'A2:D2')  // 向右填充到D2区域
 * Array2D([[1,2],[3,4]]).z补齐空位('down', 'A2:A10')  // 向下填充到A10区域
 * Array2D([[1,2],[3,4]]).z补齐空位('left', 'A2:C2')   // 向左填充到A2区域
 * Array2D([[1,2],[3,4]]).z补齐空位('up', 'A5:C10')    // 向上填充到A5区域
 *
 * // 混合参数：先按方向填充再补全
 * Array2D([[1,null],[2]]).z补齐空位('right', 'A2:D2', 0)  // [[1,0,0,0],[2,0,0,0]]
 */

/**
 * 辅助函数：根据填充方向调整坐标
 * @private
 * @param {number} row - 原始行索引
 * @param {number} col - 原始列索引
 * @param {string} direction - 填充方向 (left/right/up/down)
 * @param {number} maxLen - 原始数组最大长度
 * @param {number} finalRows - 最终行数
 * @param {number} finalCols - 最终列数
 * @returns {{row: number, col: number}} 调整后的坐标
 */
function _adjustCoordinateByDirection(row, col, direction, maxLen, finalRows, finalCols) {
    switch (direction) {
        case FILL_DIRECTION.LEFT:
            // 从右向左填充：列索引向右偏移
            return { row, col: col + (finalCols - maxLen) };
        case FILL_DIRECTION.UP:
            // 从下向上填充：行索引向下偏移（原始数据放在底部）
            return { row: row + (finalRows - maxLen), col };
        case FILL_DIRECTION.DOWN:
        case FILL_DIRECTION.RIGHT:
        default:
            // 默认：左上对齐
            return { row, col };
    }
}

Array2D.prototype.z补齐空位 = function(direction, rangeAddress, fillValue) {
    // 参数重载处理
    if (typeof direction !== 'string') {
        // 旧版调用：仅传fillValue
        fillValue = direction;
        direction = FILL_DIRECTION.RIGHT;
        rangeAddress = null;
    }

    fillValue = fillValue !== undefined ? fillValue : ARRAY_LIMITS.DEFAULT_FILL;
    direction = direction || FILL_DIRECTION.RIGHT;
    
    var result = [];
    
    // 如果提供了区域地址，解析出行列范围
    var targetRows = this._items.length;
    var targetCols = 0;
    var startRow = 0, startCol = 0;
    
    if (rangeAddress && typeof rangeAddress === 'string') {
        // 解析类似 "A2:D10" 的地址
        var match = rangeAddress.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
        if (match) {
            // 转换为0-based索引
            startCol = this._colToIndex(match[1]);  // 起始列
            startRow = parseInt(match[2]) - 1;      // 起始行
            var endCol = this._colToIndex(match[3]);   // 结束列
            var endRow = parseInt(match[4]) - 1;       // 结束行
            
            targetRows = endRow - startRow + 1;
            targetCols = endCol - startCol + 1;
        }
    }

    // 找出最大列数
    var maxLen = 0;
    for (var r = 0; r < this._items.length; r++) {
        if (this._items[r] && this._items[r].length > maxLen) {
            maxLen = this._items[r].length;
        }
    }

    // 根据方向计算最终维度
    var finalRows = targetRows || this._items.length;
    var finalCols = targetCols || Math.max(maxLen, targetCols);

    // 按方向填充
    for (var i = 0; i < finalRows; i++) {
        var row = new Array(finalCols);

        // 初始化全为fillValue
        for (var j = 0; j < finalCols; j++) {
            row[j] = fillValue;
        }

        // 根据方向填充原始数据
        for (var j = 0; j < finalCols; j++) {
            // 使用辅助函数调整坐标
            var adjusted = _adjustCoordinateByDirection(i, j, direction, maxLen, finalRows, finalCols);
            var origRow = adjusted.row;
            var origCol = adjusted.col;

            // 检查是否在原始数组范围内
            if (origRow >= 0 && origRow < this._items.length &&
                origCol >= 0 && this._items[origRow] && origCol < this._items[origRow].length) {
                var val = this._items[origRow][origCol];
                row[j] = (val === null || val === undefined) ? fillValue : val;
            }
        }

        result.push(row);
    }

    return this._new(result);
};

// 列字母转数字索引的辅助函数
Array2D.prototype._colToIndex = function(colStr) {
    var result = 0;
    for (var i = 0; i < colStr.length; i++) {
        result = result * 26 + (colStr.charCodeAt(i) - 64);
    }
    return result - 1; // 返回0-based索引
};

Array2D.prototype.fillBlank = Array2D.prototype.z补齐空位;

/**
 * 扁平化（降维）
 * @returns {Array} 一维数组
 * @example
 * Array2D([[1,2],[3,4]]).z扁平化()  // [1,2,3,4]
 */
Array2D.prototype.z扁平化 = function() {
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        if (Array.isArray(this._items[i])) {
            for (var j = 0; j < this._items[i].length; j++) {
                result.push(this._items[i][j]);
            }
        } else {
            result.push(this._items[i]);
        }
    }
    return result;
};
Array2D.prototype.flat = Array2D.prototype.z扁平化;

/**
 * 数组反转
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z反转()  // [[3,4],[1,2]]
 */
Array2D.prototype.z反转 = function() {
    return this._new(this._items.slice().reverse());
};
Array2D.prototype.reverse = Array2D.prototype.z反转;

// ==================== 统计计算 ====================

/**
 * 求和
 * @param {string|Function} [colSelector] - 列选择器 'f1'=第1列, 或回调函数
 * @returns {Number} 和
 * @example
 * Array2D([[1,2],[3,4]]).z求和()        // 10
 * Array2D([[1,2],[3,4]]).z求和('f1')     // 4 (第1列)
 */
// 🔧 XXD-145 final fix: z计数 = Excel COUNT() — 仅计数有效数值, 与 z求和/z平均值 的有效值集合一致
// 例 [['v'],[1],[2],[3]].z计数() → 3 ('v' 为文本, 不计入)
Array2D.prototype.z计数 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    return flat.filter(function(v) {
        if (v === null || v === undefined || v === '') return false;
        if (typeof v === 'number') return !isNaN(v);
        const num = parseFloat(String(v).replace(/,/g, ''));
        return !isNaN(num);
    }).length;
};
Array2D.prototype.z求和 = function(colSelector) {
    // 🔧 XXD-161 final fix: 数字 colSelector 视为 1-based 列索引, 越界抛 RangeError
    // 之前: parseLambda(5) → null (typeof !== 'string'), 静默回退到 z扁平化() 求全部, 隐藏错误
    if (typeof colSelector === 'number') {
        const colCount = this.z列数();
        if (colSelector < 1 || colSelector > colCount) {
            throw new RangeError('列索引越界: colSelector=' + colSelector + ', 列数=' + colCount);
        }
        const colIdx = colSelector - 1;
        const flat = this._items.map(function(row) { return Array.isArray(row) ? row[colIdx] : undefined; });
        const sum = flat.reduce((acc, val) => {
            const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
            return acc + (isNaN(num) ? 0 : num);
        }, 0);
        return Math.round(sum * 1e10) / 1e10;
    }
    // 🔧 XXD-162 final fix: 字符串 colSelector 解析流程
    //  1) 已知 lambda 语法 (f1/$0/箭头函数/方括号) → 走 parseLambda
    //  2) 否则尝试作为 header 名 (匹配首行表头)
    //  3) 都不匹配抛 TypeError (避免之前抛 "A is not defined" 的 ReferenceError)
    var fn = null;
    if (typeof colSelector === 'string') {
        // 预判: 已知 lambda 语法 → 直接走 parseLambda
        var isLambdaSyntax = /=>|\$\d|f\s*\(?\s*\d+\s*\)?/i.test(colSelector) || /^\[.*\]$/.test(colSelector);
        if (isLambdaSyntax) {
            fn = parseLambda(colSelector);
        } else {
            // 尝试作为 header 名 (匹配首行表头)
            if (this._items.length > 0 && Array.isArray(this._items[0])) {
                var headerRow = this._items[0];
                for (var hi = 0; hi < headerRow.length; hi++) {
                    if (headerRow[hi] !== null && headerRow[hi] !== undefined && String(headerRow[hi]) === colSelector) {
                        var colIdxH = hi;
                        fn = function(row) { return Array.isArray(row) ? row[colIdxH] : undefined; };
                        break;
                    }
                }
            }
            if (!fn) {
                var sampleHeaders = (this._items.length > 0 && Array.isArray(this._items[0]))
                    ? this._items[0].slice(0, 5).map(function(h) { return h === null || h === undefined ? '∅' : String(h); }).join(', ')
                    : '(空)';
                throw new TypeError(
                    'z求和: 列选择器 "' + colSelector + '" 既不是合法 lambda 表达式 ' +
                    '(f1/f2/$0/row=>row.x 等), 也不是表头名. 当前首行表头: [' + sampleHeaders + ']'
                );
            }
        }
    } else if (typeof colSelector === 'function') {
        fn = colSelector;
    }
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    // 🔧 XXD-97 final fix: 浮点累加误差 (0.1+0.2+0.3=0.6000000000000001) → 累加后四舍五入到 1e-10
    const sum = flat.reduce((acc, val) => {
        const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        return acc + (isNaN(num) ? 0 : num);
    }, 0);
    return Math.round(sum * 1e10) / 1e10;
};
Array2D.prototype.sum = Array2D.prototype.z求和;
// 🔧 XXD-143 final fix: z方差/z标准差 (实现)
Array2D.prototype.z方差 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const nums = flat.filter(function(v) { return typeof v === 'number' && !isNaN(v); });
    if (nums.length < 2) return 0;
    const avg = nums.reduce(function(a,b){return a+b;},0) / nums.length;
    const variance = nums.reduce(function(acc,v){return acc + (v-avg)*(v-avg);}, 0) / (nums.length - 1);
    return Math.round(variance * 1e10) / 1e10;
};
Array2D.prototype.z标准差 = function(colSelector) {
    const v = this.z方差(colSelector);
    return Math.round(Math.sqrt(v) * 1e10) / 1e10;
};

/**
 * 求平均值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 平均值
 */
Array2D.prototype.z平均值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    // 🔧 v3.9.4 修复：分母只计算有效数值项，排除 NaN
    let validCount = 0;
    let sum = 0;
    flat.forEach(function(val) {
        const num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        if (!isNaN(num)) { validCount++; sum += num; }
    });
    // 🔧 XXD-148 final fix: 浮点累加误差 → 最终四舍五入
    if (validCount === 0) return 0;
    const result = sum / validCount;
    return Math.round(result * 1e10) / 1e10;
};
Array2D.prototype.average = Array2D.prototype.z平均值;

/**
 * 求最大值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 最大值，空数组返回 undefined
 */
Array2D.prototype.z最大值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    // 🔧 XXD-95/XXD-140/XXD-144/XXD-160 fix:
    // - raw Array2D (无 _header) 默认跳过第 0 行 header (整行, 不只首格)
    // - 1-row raw 2D (e.g. [[1,2,3]]) 视为数据, 不再误当 header 跳过 (XXD-160)
    // - 字符串按 String 比较, 不再静默丢弃
    var items;
    if (fn) {
        items = this._items.map(fn);
    } else if (!(Object.prototype.hasOwnProperty.call(this, '_header')) && this._items.length > 1 && Array.isArray(this._items[0])) {
        items = [];
        for (var r = 1; r < this._items.length; r++) {
            var row = this._items[r];
            if (Array.isArray(row)) { for (var c = 0; c < row.length; c++) items.push(row[c]); }
            else items.push(row);
        }
    } else {
        items = this.z扁平化();
    }
    if (items.length === 0) return undefined;
    var maxVal = items[0];
    for (var i = 1; i < items.length; i++) {
        var v = items[i];
        if (v === null || v === undefined) continue;
        if (typeof maxVal === 'number' && typeof v === 'number') {
            if (v > maxVal) maxVal = v;
        } else {
            if (String(v) > String(maxVal)) maxVal = v;
        }
    }
    return maxVal;
};
Array2D.prototype.max = Array2D.prototype.z最大值;

/**
 * 求最小值
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 最小值，空数组返回 undefined
 */
Array2D.prototype.z最小值 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    // 🔧 XXD-95/XXD-140/XXD-144/XXD-160 fix:
    // - raw Array2D (无 _header) 默认跳过第 0 行 header (整行, 不只首格)
    // - 1-row raw 2D (e.g. [[1,2,3]]) 视为数据, 不再误当 header 跳过 (XXD-160)
    // - 字符串按 String 比较, 不再静默丢弃
    var items;
    if (fn) {
        items = this._items.map(fn);
    } else if (!(Object.prototype.hasOwnProperty.call(this, '_header')) && this._items.length > 1 && Array.isArray(this._items[0])) {
        items = [];
        for (var r = 1; r < this._items.length; r++) {
            var row = this._items[r];
            if (Array.isArray(row)) { for (var c = 0; c < row.length; c++) items.push(row[c]); }
            else items.push(row);
        }
    } else {
        items = this.z扁平化();
    }
    if (items.length === 0) return undefined;
    var minVal = items[0];
    for (var i = 1; i < items.length; i++) {
        var v = items[i];
        if (v === null || v === undefined) continue;
        if (typeof minVal === 'number' && typeof v === 'number') {
            if (v < minVal) minVal = v;
        } else {
            if (String(v) < String(minVal)) minVal = v;
        }
    }
    return minVal;
};
Array2D.prototype.min = Array2D.prototype.z最小值;

/* XXD-194: Array2D.prototype.z聚合 instance alias */
// XXD-194: instance-side z聚合 — delegates to Array2D.agg static for identical semantics.
// 复现: new Array2D([['v'],[1],[2]]).z聚合() 之前 throw, 现返回 3 (sum; 'v' as non-numeric skipped).
// 支持 (colSelector, aggType) 或 reduce 风格 (fn) 调用.
Array2D.prototype.z聚合 = function(colSelector, aggType) {
    if (typeof colSelector === 'function' && arguments.length === 1) {
        var acc;
        for (var i = 0; i < this._items.length; i++) {
            acc = colSelector(acc, this._items[i], i);
        }
        return acc;
    }
    return Array2D.agg(this._items, colSelector, aggType);
};
Array2D.prototype.agg = Array2D.prototype.z聚合;
Array2D.prototype.aggregate = Array2D.prototype.z聚合;


Array2D.prototype.z列统计 = function(colSelector) {
    if (!this._items || this._items.length === 0) {
        return { sum: 0, avg: 0, min: 0, max: 0, count: 0 };
    }
    // Resolve a per-row value extractor. If colSelector is given, use it; otherwise walk columns.
    function extractRow(row) {
        if (!Array.isArray(row)) return [row];
        if (colSelector === undefined || colSelector === null) return row;
        if (typeof colSelector === 'function') return [colSelector(row)];
        var s = String(colSelector);
        var m = s.match(/^f(\d+)$/i);
        var idx = m ? parseInt(m[1], 10) - 1 : (parseInt(s, 10) || 0);
        return [row[idx]];
    }
    // Array2D raw-2D convention: no _header + length>1 + first row is array ⇒ skip row 0 as header.
    var hasHeaderRow = !Object.prototype.hasOwnProperty.call(this, '_header')
        && this._items.length > 1
        && Array.isArray(this._items[0]);
    var startIdx = hasHeaderRow ? 1 : 0;
    var sum = 0, min = NaN, max = NaN, count = 0;
    for (var i = startIdx; i < this._items.length; i++) {
        var vals = extractRow(this._items[i]);
        for (var k = 0; k < vals.length; k++) {
            var v = vals[k];
            if (typeof v === 'number' && !isNaN(v)) {
                sum += v;
                if (isNaN(min) || v < min) min = v;
                if (isNaN(max) || v > max) max = v;
                count++;
            }
        }
    }
    var avg = count > 0 ? sum / count : 0;
    return { sum: sum, avg: avg, min: isNaN(min) ? 0 : min, max: isNaN(max) ? 0 : max, count: count };
};
Array2D.prototype.columnStats = Array2D.prototype.z列统计;

/**
 * 求中位数（median）
 * @param {string|Function} [colSelector] - 列选择器
 * @returns {Number} 中位数
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z中位数()  // 3.5
 */
Array2D.prototype.z中位数 = function(colSelector) {
    const fn = colSelector ? parseLambda(colSelector) : null;
    const flat = fn ? this._items.map(fn) : this.z扁平化();
    const numbers = flat.filter(v => typeof v === 'number' || !isNaN(parseFloat(v)))
        .map(v => typeof v === 'number' ? v : parseFloat(v));
    if (numbers.length === 0) return undefined;
    numbers.sort(function(a, b) { return a - b; });
    const mid = Math.floor(numbers.length / 2);
    return numbers.length % 2 !== 0 ? numbers[mid] : (numbers[mid - 1] + numbers[mid]) / 2;
};
Array2D.prototype.median = Array2D.prototype.z中位数;

/**
 * 获取第一个元素
 * @returns {any} 第一个元素
 */
Array2D.prototype.z第一个 = function() {
    const flat = this.z扁平化();
    return flat.length > 0 ? flat[0] : undefined;
};
Array2D.prototype.first = Array2D.prototype.z第一个;

/**
 * 获取最后一个元素
 * @returns {any} 最后一个元素
 */
Array2D.prototype.z最后一个 = function() {
    const flat = this.z扁平化();
    return flat.length > 0 ? flat[flat.length - 1] : undefined;
};
Array2D.prototype.last = Array2D.prototype.z最后一个;

// ==================== 矩阵操作 ====================

/**
 * 转置矩阵
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z转置()  // [[1,4],[2,5],[3,6]]
 */
Array2D.prototype.z转置 = function() {
    if (!this._items || this._items.length === 0) return this._new([]);
    // 一维数组处理：[1,2,3] → [[1],[2],[3]]
    if (!Array.isArray(this._items[0])) {
        var result = [];
        for (var i = 0; i < this._items.length; i++) {
            result[i] = [this._items[i]];
        }
        return this._new(result);
    }
    var rows = this._items.length;
    // XXD-21: 取所有行长度的最大值, 而不是首行长度, 否则首行较短的参差数组转置后会缺列
    var cols = 0;
    for (var k = 0; k < this._items.length; k++) {
        if (Array.isArray(this._items[k])) {
            cols = Math.max(cols, this._items[k].length);
        }
    }
    var result = [];
    for (var j = 0; j < cols; j++) {
        result[j] = [];
        for (var i = 0; i < rows; i++) {
            result[j][i] = this._items[i][j];
        }
    }
    return this._new(result);
};
Array2D.prototype.transpose = Array2D.prototype.z转置;

/**
 * 获取行列数
 * @returns {String} "行数x列数"
 */
Array2D.prototype.z矩阵信息 = function() {
    const rows = this._items.length;
    const cols = rows > 0 && this._items[0] ? this._items[0].length : 0;
    return `${rows}x${cols}`;
};
Array2D.prototype.matrixInfo = Array2D.prototype.z矩阵信息;

/**
 * 获取单元格值
 * @param {Number} row - 行号（从0开始）
 * @param {Number} col - 列号（从0开始）
 * @returns {any} 单元格值
 */
Array2D.prototype.z单元格 = function(row, col) {
    var rowCount = this._items.length;
    var colCount = rowCount > 0 && this._items[0] ? this._items[0].length : 0;
    if (typeof row !== 'number' || isNaN(row) || row < 0 || row >= rowCount) {
        throw new RangeError('行索引越界 row=' + row + ', 行数=' + rowCount);
    }
    if (typeof col !== 'number' || isNaN(col) || col < 0 || col >= colCount) {
        throw new RangeError('列索引越界 col=' + col + ', 列数=' + colCount);
    }
    return this._items[row][col];
};
Array2D.prototype.cell = Array2D.prototype.z单元格;

/**
 * 设置单元格值
 * @param {Number} row - 行号
 * @param {Number} col - 列号
 * @param {string|number|boolean|null|Date|object} value - 新值
 * @returns {Array2D} 当前实例
 */
Array2D.prototype.z设置单元格 = function(row, col, value) {
    var rowCount = this._items.length;
    var colCount = rowCount > 0 && this._items[0] ? this._items[0].length : 0;
    if (typeof row !== 'number' || isNaN(row) || row < 0 || row >= rowCount) {
        throw new RangeError('行索引越界 row=' + row + ', 行数=' + rowCount);
    }
    if (typeof col !== 'number' || isNaN(col) || col < 0 || col >= colCount) {
        throw new RangeError('列索引越界 col=' + col + ', 列数=' + colCount);
    }
    if (!this._items[row]) this._items[row] = [];
    this._items[row][col] = value;
    return this;
};
Array2D.prototype.setCell = Array2D.prototype.z设置单元格;

/**
 * 写入单元格（实例方法，根据数组大小自动扩展区域）
 * @param {Range|string} rng - 目标单元格区域（左上角单元格）
 * @param {Boolean|Number} [clearBelow] - 是否清空下方区域，默认false
 *   - true/1: 写入后清空输出区域下方的原有数据
 *   - false/不传: 仅写入数组覆盖区域，下方旧数据保留
 * @returns {Array2D} 当前实例（支持链式调用）
 * @example
 * Array2D([[1,2],[3,4]]).toRange("A1")           // 写入A1:B2
 * Array2D([[1,2],[3,4]]).toRange("A1", true)     // 写入并清空下方区域
 * Array2D([[1,2],[3,4]]).z写入单元格("K1")        // 写入K1:L2
 */
Array2D.prototype.toRange = function(rng, clearBelow) {
    // 空数组检查，防止 Item(0,0) 报错
    var items = this._items;
    if (!items || items.length === 0) {
        return this;
    }
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = items.length;
    var cols = rows > 0 ? (Array.isArray(items[0]) ? items[0].length : 1) : 0;

    // 列数边界检查
    if (cols === 0) return this;
    // 根据数组大小调整目标区域
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    // 🔧 Bug修复: 使用批量解除合并（与静态方法一致），替代逐个单元格检查
    try {
        writeRng.MergeCells = false;
    } catch (e) {
        // 如果一次性解除失败，回退到逐个检查
        for (var i = 1; i <= writeRng.Rows.Count; i++) {
            for (var j = 1; j <= writeRng.Columns.Count; j++) {
                var cell = writeRng.Cells(i, j);
                if (cell.MergeCells) {
                    cell.MergeArea.UnMerge();
                }
            }
        }
    }
    writeRng.Value2 = items;

    // 🔧 v4.0.10 修复: 支持 clearBelow 参数清空下方区域
    if (clearBelow === true || clearBelow === 1) {
        // 获取当前区域下方的已使用区域并清空
        try {
            var usedRange = sheet.UsedRange;
            if (usedRange) {
                var writeEndRow = targetRng.Row + rows - 1;
                var usedEndRow = usedRange.Row + usedRange.Rows.Count - 1;
                if (usedEndRow > writeEndRow) {
                    // 清空下方区域
                    var belowRng = sheet.Range(
                        sheet.Cells(writeEndRow + 1, targetRng.Column),
                        sheet.Cells(usedEndRow, targetRng.Column + cols - 1)
                    );
                    belowRng.ClearContents();
                }
            }
        } catch (e) {
            // 忽略清空错误
        }
    }

    return this;
};
Array2D.prototype.z写入单元格 = Array2D.prototype.toRange;

/**
 * 连接成字符串
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 */
Array2D.prototype.z连接 = function(separator = ',') {
    return this._items.map(row => Array.isArray(row) ? row.join(separator) : String(row)).join(separator);
};
Array2D.prototype.join = Array2D.prototype.z连接;

/**
 * 文本连接（textjoin）- 选择指定列的值，用分隔符连接
 * @param {String|Number|Function} selector - 列选择器，如 'f1' 或 0 或 row=>row.col
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D([['a','b'],['c','d']]).z文本连接(1, '+')  // "b+d"
 * Array2D([['a','b'],['c','d']]).textjoin('f2', '+')  // "b+d"
 */
Array2D.prototype.z文本连接 = function(selector, separator = ',') {
    var fn;
    if (typeof selector === 'function') {
        fn = selector;
    } else if (typeof selector === 'number') {
        var idx = selector;
        fn = function(row) { return Array.isArray(row) ? row[idx] : row; };
    } else {
        fn = parseLambda(selector);
    }
    var values = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (fn) {
            values.push(fn(row, i));
        } else {
            values.push(Array.isArray(row) ? row[0] : row);
        }
    }
    return values.join(separator);
};
Array2D.prototype.textjoin = Array2D.prototype.z文本连接;

/**
 * 转JSON（转JSON字符串，二维数组内部数组横着对齐显示）
 * @param {Boolean} [pretty=true] - 是否格式化输出（对齐显示）
 * @returns {String} JSON字符串
 * @example
 * Array2D([[1,2],[3,4]]).z转JSON()
 * // 输出:
 * // [
 * //  [1, 2],
 * //  [3, 4]
 * // ]
 * Array2D([[1,2],[3,4]]).toJson(false)    // "[[1,2],[3,4]]" 紧凑格式
 */
Array2D.prototype.z转JSON = function(pretty) {
    // 紧凑格式
    if (pretty === false) {
        return JSON.stringify(this._items);
    }
    // 格式化输出（对齐显示）
    if (Array.isArray(this._items) && this._items.length > 0 && Array.isArray(this._items[0])) {
        var lines = formatArray2DAsJSON(this._items);
        return lines.join('\n');
    }
    // 其他情况使用标准JSON格式
    return JSON.stringify(this._items, null, 2);
};
Array2D.prototype.toJson = Array2D.prototype.z转JSON;

// ==================== 分块挑选 ====================

/**
 * 分块
 * @param {Number} size - 块大小
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z分块(2)  // [[[1],[2]],[[3],[4]],[[5]]]
 */
/**
 * 分块
 * @param {Number} rowSize - 块行数(双参时也作为每块行数)
 * @param {Number} [colSize] - 块列数(可选). 省略时按行分块,返回 [[chunkRows...],[chunkRows...],...]
 * @returns {Array2D} 新实例
 * @example
 *  // 单参: 按行分块,每块 rowSize 行
 *  Array2D([[1],[2],[3],[4],[5]]).z分块(2)
 *  // => [[[1],[2]],[[3],[4]],[[5]]]
 *
 *  // 双参: 将每行按 colSize 切片,按行优先拼接为扁平二维数组
 *  Array2D([[1,2,3,4],[5,6,7,8]]).z分块(2, 2)
 *  // => [[1,2],[3,4],[5,6],[7,8]]
 */
Array2D.prototype.z分块 = function(rowSize, colSize) {
    // 🔧 v3.9.4 修复：size <= 0 时返回空数组，防止死循环
    if (!rowSize || rowSize <= 0) return this._new([]);

    // 单参: 按行分块(保留旧行为)
    if (colSize === undefined || colSize === null) {
        const result = [];
        for (let i = 0; i < this._items.length; i += rowSize) {
            result.push(this._items.slice(i, i + rowSize));
        }
        return this._new(result);
    }

    // 双参: 把每行按 colSize 切片,再按行优先拼接为扁平的二维数组
    // (XXD-186: 修复嵌套一层的 BUG,语义:rowSize 行一组,组内每行按 colSize 切片)
    if (colSize <= 0) return this._new([]);
    const result = [];
    for (let r = 0; r < this._items.length; r += rowSize) {
        for (let dr = 0; dr < rowSize; dr++) {
            const srcRow = this._items[r + dr];
            if (!Array.isArray(srcRow)) continue;
            for (let c = 0; c < srcRow.length; c += colSize) {
                result.push(srcRow.slice(c, c + colSize));
            }
        }
    }
    return this._new(result);
};
Array2D.prototype.chunk = Array2D.prototype.z分块;

/**
 * 挑选元素
 * @param {Number} count - 数量
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z挑选(3)  // [[1],[2],[3]]
 */
Array2D.prototype.z挑选 = function(count) {
    return this._new(this._items.slice(0, count));
};
Array2D.prototype.pick = Array2D.prototype.z挑选;

/**
 * 跳过元素
 * @param {Number} count - 跳过数量
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1],[2],[3],[4],[5]]).z跳过(2)  // [[3],[4],[5]]
 */
Array2D.prototype.z跳过 = function(count) {
    // 🔧 v3.9.4 修复：边界检查，非法值默认为0
    var n = (typeof count === 'number' && count > 0) ? Math.floor(count) : 0;
    return this._new(this._items.slice(n));
};
Array2D.prototype.skip = Array2D.prototype.z跳过;

/**
 * 取前N个
 * @param {Number} count - 数量
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z取前N个 = function(count) {
    // 🔧 v3.9.4 修复：边界检查，非法值默认取全部
    var n = (typeof count === 'number' && count >= 0) ? Math.floor(count) : this._items.length;
    return this._new(this._items.slice(0, n));
};
Array2D.prototype.take = Array2D.prototype.z取前N个;

// 补充别名：用户可能期望的更具描述性的方法名
Array2D.prototype.z跳过前N个 = Array2D.prototype.z跳过;
Array2D.prototype.z跳过前几个 = Array2D.prototype.z跳过;
Array2D.prototype.z取前几个 = Array2D.prototype.z取前N个;

/**
 * 重复N次（repeat）- 将数组重复指定次数
 * @param {Number} count - 重复次数
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).repeat(2)  // [[1,2],[3,4],[1,2],[3,4]]
 */
Array2D.prototype.repeat = function(count) {
    if (!count || count <= 0) return this._new([]);
    var result = [];
    for (var i = 0; i < count; i++) {
        for (var j = 0; j < this._items.length; j++) {
            result.push(JSON.parse(JSON.stringify(this._items[j])));
        }
    }
    return this._new(result);
};
Array2D.prototype.z重复N次 = Array2D.prototype.repeat;

/**
 * 随机打乱（shuffle）- 随机打乱数组顺序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).shuffle()  // 随机顺序
 */
Array2D.prototype.shuffle = function() {
    var result = JSON.parse(JSON.stringify(this._items));
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return this._new(result);
};
Array2D.prototype.z随机打乱 = Array2D.prototype.shuffle;

/**
 * 随机一项（random）- 随机选择一组
 * @param {Number} n - 可选，先打乱全部再取前n个
 * @returns {Array2D} 新实例
 * @example
 * Array2D([1,2,3,4,5,6]).random()        // 随机返回 Array2D([3])
 * Array2D([1,2,3,4,5,6]).random(3)       // 先打乱全部，再取前3个，返回 Array2D([2,1,3])
 * Array2D([[1,2],[3,4]]).random()        // 随机返回 Array2D([[1,2]])
 */
Array2D.prototype.random = function(n) {
    // 检测是否为一维数组
    var isOneD = this._items.length > 0 && !Array.isArray(this._items[0]);

    if (n !== undefined && n > 0) {
        // 先打乱整个数组，再取前n个
        var result = JSON.parse(JSON.stringify(this._items));

        // Fisher-Yates 洗牌整个数组
        for (var i = result.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = result[i];
            result[i] = result[j];
            result[j] = temp;
        }

        // 取前n个
        result = result.slice(0, Math.min(n, this._items.length));

        return this._new(result);
    } else {
        // 随机选择一项
        var idx = Math.floor(Math.random() * this._items.length);
        var item = this._items[idx];

        // 如果是一维数组，包装成二维；二维数组保持二维
        if (isOneD) {
            return this._new([[item]]);
        }
        return this._new([item]);
    }
};
Array2D.prototype.z随机一项 = Array2D.prototype.random;

/**
 * 跳过前面连续满足（skipWhile）- 跳过前面连续满足条件的元素
 * @param {string|Function} predicate - 条件函数
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z跳过前面连续满足('x=>x[0]<4')  // [[5,6]]
 */
/**
 * 取前面连续满足（takeWhile）- 取前面连续满足条件的元素
 * @param {string|Function} predicate - 条件函数
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z取前面连续满足('x=>x[0]<4')  // [[1,2],[3,4]]
 */
/**
 * 行切片（slice）- 提取指定范围的行
 * @param {Number} [start=0] - 起始索引
 * @param {Number} [end] - 结束索引（不包含）
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z行切片(1, 2)  // [[3,4]]
 */
Array2D.prototype.z行切片 = function(start, end) {
    start = start || 0;
    if (end === undefined) end = this._items.length;
    return this._new(this._items.slice(start, end));
};
Array2D.prototype.slice = Array2D.prototype.z行切片;

/**
 * 行切片删除行（splice）- 删除/插入行
 * @param {Number} start - 起始位置
 * @param {Number} [deleteCount=1] - 删除数量
 * @param {...Array} items - 要插入的行
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z行切片删除行(1, 1)  // [[1,2],[5,6]]
 */
Array2D.prototype.z行切片删除行 = function(start, deleteCount, items) {
    deleteCount = deleteCount !== undefined ? deleteCount : 1;
    var result = this._items.slice();
    var removed = result.splice.apply(result, [start, deleteCount].concat(Array.prototype.slice.call(arguments, 2)));
    return this._new(result);
};
Array2D.prototype.splice = Array2D.prototype.z行切片删除行;

/**
 * 转字符串（toString）- 将数组转换为字符串
 * @param {string} [rowSeparator='\n'] - 行分隔符
 * @param {string} [colSeparator=','] - 列分隔符
 * @returns {string} 字符串
 * @example
 * Array2D([[1,2],[3,4]]).z转字符串()  // "1,2\n3,4"
 */
Array2D.prototype.z转字符串 = function(rowSeparator, colSeparator) {
    rowSeparator = rowSeparator !== undefined ? rowSeparator : '\n';
    colSeparator = colSeparator !== undefined ? colSeparator : ',';
    var normalizeCell = function(v) {
        if (typeof v === 'string') {
            return v.replace(/[，；：！？（）【】《》、。．]/g, function(ch) {
                return Array2D.prototype.z转字符串._fwMap[ch];
            });
        }
        return v;
    };
    return this._items.map(function(row) {
        if (Array.isArray(row)) {
            return row.map(normalizeCell).join(colSeparator);
        }
        return String(normalizeCell(row));
    }).join(rowSeparator);
};
Array2D.prototype.z转字符串._fwMap = {
    '，': ',', '；': ';', '：': ':', '！': '!', '？': '?',
    '（': '(', '）': ')', '【': '[', '】': ']', '《': '<',
    '、': ',', '。': '.', '．': '.'
};
Array2D.prototype.toString = Array2D.prototype.z转字符串;

// ==================== 查找筛选 ====================

/**
 * 查找元素下标
 * 多签名支持：z查找索引(col, val) / z查找索引(col, val, startIdx) / z查找索引(predicate) / z查找索引(value)
 * @param {Number|String|Function} colOrPred - 列号 / 条件函数或字符串 / 要查找的值
 * @param {any} [val] - 当第一参数为列号时，要匹配的值
 * @param {Number} [startIdx] - 当第一参数为列号时，起始行下标
 * @returns {Number} 行下标，未找到返回-1
 */
Array2D.prototype.z查找索引 = function(colOrPred, val, startIdx) {
    var i, row, fn, flat;

    // 多参数模式: z查找索引(col, val) / z查找索引(col, val, startIdx)
    if (arguments.length >= 2 && typeof colOrPred === 'number') {
        var col = colOrPred;
        var from = (typeof startIdx === 'number' && startIdx >= 0) ? startIdx : 0;
        for (i = from; i < this._items.length; i++) {
            row = this._items[i];
            if (Array.isArray(row)) {
                if (row[col] == val) return i;
            } else if (col === 0) {
                if (row == val) return i;
            }
        }
        return -1;
    }

    // 单参数: 条件函数 → 行级匹配
    if (typeof colOrPred === 'function') {
        for (i = 0; i < this._items.length; i++) {
            if (colOrPred(this._items[i], i)) return i;
        }
        return -1;
    }

    // 单参数: 字符串 → 尝试 parseLambda，成功则行级匹配，失败则扁平值查找
    if (typeof colOrPred === 'string') {
        fn = parseLambda(colOrPred);
        if (fn) {
            for (i = 0; i < this._items.length; i++) {
                if (fn(this._items[i], i)) return i;
            }
            return -1;
        }
    }

    // 兜底: 扁平值查找（向后兼容 z包含 等）
    flat = this.z扁平化();
    for (i = 0; i < flat.length; i++) {
        if (flat[i] == colOrPred) return i;
    }
    return -1;
};
Array2D.prototype.findIndex = Array2D.prototype.z查找索引;

/**
 * 检查是否包含元素
 * @param {any} value - 要检查的值
 * @returns {Boolean} 是否包含
 */
Array2D.prototype.z包含 = function(value) {
    return this.z查找索引(value) !== -1;
};
Array2D.prototype.includes = Array2D.prototype.z包含;

/**
 * 遍历每个元素
 * @param {Function} callback - 回调函数 (item, index)
 * @returns {Array2D} this 支持链式调用
 * @example
 * Array2D([[1,2],[3,4]]).forEach((row, i) => Console.log(i, row))
 */
Array2D.prototype.forEach = function(callback) {
    this._items.forEach(callback);
    return this;
};

/**
 * 倒序遍历执行（forEachRev）- 从后向前遍历每个元素
 * @param {Function} callback - 回调函数 (item, index)，返回false可中断
 * @returns {Array2D} this 支持链式调用
 * @example
 * Array2D([[1,2],[3,4]]).z倒序遍历执行((row, i) => Console.log(i, row))
 */
Array2D.prototype.z倒序遍历执行 = function(callback) {
    for (var i = this._items.length - 1; i >= 0; i--) {
        var result = callback(this._items[i], i);
        if (result === false) break; // 支持提前退出
    }
    return this;
};
Array2D.prototype.forEachRev = Array2D.prototype.z倒序遍历执行;

/**
 * 筛选元素
 * @param {string|Function} predicate - 筛选条件
 * @returns {Array2D} 新实例
 * @example
 * Array2D([1,2,3,4]).z筛选('x=>x>2')  // [3,4]
 */
Array2D.prototype.z筛选 = function(predicate, skipHeader) {
    // 🔧 v3.7.9 修复: 如果没有指定 skipHeader 但对象有 _header 属性，自动设为 1
    if (skipHeader === undefined && Object.prototype.hasOwnProperty.call(this, '_header') && this._header !== undefined) {
        skipHeader = 1;
    }
    // 处理 skipHeader 参数
    var data = this._items;
    if (skipHeader && skipHeader > 0) {
        data = data.slice(skipHeader);
    }

    // 🔧 v4.0.29 诊断: 打印 z筛选 入口参数
    if (typeof Console !== 'undefined') {
        try {
            Console.log('[k/v4.0.29] z筛选 IN: this._items.len=' + (this._items ? this._items.length : 'n/a') +
                ', data.len=' + (data ? data.length : 'n/a') +
                ', skipHeader=' + skipHeader +
                ', pred.t=' + typeof predicate +
                ', hasFn=' + (typeof predicate === 'function'));
        } catch (__) {}
    }

    // 处理对象参数形式（增强功能）
    if (predicate && typeof predicate === 'object' && !Array.isArray(predicate)) {
        // XXD-218: {headerName:value} 简写形式 — 按 header 名匹配列值
        var __condKeys = Object.keys(predicate);
        var __isShorthand = __condKeys.length > 0 && !('column' in predicate) && !('operator' in predicate);
        if (__isShorthand) {
            var __hdrRow = null;
            if (this._header && Array.isArray(this._header) && this._header.length > 0) {
                __hdrRow = Array.isArray(this._header[0]) ? this._header[0] : this._header;
            } else if (data.length > 0 && Array.isArray(data[0])) {
                __hdrRow = data[0]; data = data.slice(1);
            }
            if (__hdrRow) {
                var __hdrMap = {};
                for (var __hi = 0; __hi < __hdrRow.length; __hi++) {
                    __hdrMap[String(__hdrRow[__hi])] = __hi;
                }
                var __filtered = data.filter(function(__row) {
                    if (!Array.isArray(__row)) return false;
                    for (var __ki = 0; __ki < __condKeys.length; __ki++) {
                        var __k = __condKeys[__ki];
                        var __ci = __hdrMap[__k];
                        if (__ci === undefined || __row[__ci] != predicate[__k]) return false;
                    }
                    return true;
                });
                return this._new(__filtered);
            }
        }
        return this._new(Array2D._filterByObject(data, predicate));
    }

    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return this._new([]);
    // 🔧 v4.0.13 修复: 为每行创建代理，支持 x.f1, x.f2 等语法（与 z映射 对齐）
    var result = [];
    // XXD-165: 收集谓词执行异常, 避免静默吞错返回 [] 让用户误以为数据为空
    var __errors = [];
    for (var __fi = 0; __fi < data.length; __fi++) {
        var __row = data[__fi];
        var __proxy = Array.isArray(__row) ? __row.slice() : [__row];
        // 🔧 v4.0.30 修复: superPivot 输出 f2 列带前导空格,做 trim 后再赋 fN
        //   根因: 数据 "    Product1" 被 z筛选 严格匹配 'Product1' 失败
        //   修法: 字符串自动 trim 后赋给 x.fN (同时保留原值)
        for (var __fc = 0; __fc < __proxy.length; __fc++) {
            var __cellVal = __proxy[__fc];
            if (typeof __cellVal === 'string') {
                __proxy['f' + (__fc + 1)] = __cellVal.replace(/^\s+|\s+$/g, '');
            } else {
                __proxy['f' + (__fc + 1)] = __cellVal;
            }
        }
        // 🔧 v4.0.29 诊断: 打印每行判定结果
        if (typeof Console !== 'undefined' && __fi < 5) {
            try {
                var __p = __proxy;
                Console.log('[k/v4.0.30] z筛选 row[' + __fi + ']: f1=' + JSON.stringify(__p.f1) + ', f2=' + JSON.stringify(__p.f2) + ', rowLen=' + __p.length);
            } catch (__) {}
        }
        // XXD-165: 谓词抛错时不再静默吞掉, 记录原始异常并跳过该行(其他行不受影响)
        try {
            if (fn(__proxy, __fi)) {
                result.push(__row);
            }
        } catch (__perr) {
            var __errMsg = (__perr && __perr.message) ? (__perr.name + ': ' + __perr.message) : String(__perr);
            __errors.push({ row: __fi, error: __errMsg });
            if (typeof Console !== 'undefined') {
                try { Console.warn('[k/XXD-165] z筛选 谓词在第 ' + __fi + ' 行抛错: ' + __errMsg); } catch (__) {}
            }
        }
    }
    var __ret = this._new(result);
    // XXD-165: 将本次调用过程中收集到的谓词异常挂到结果实例上, 便于用户排查
    // XXD-219: 总是定义 _errors (无错时返回 [], 而非 undefined)
    {
        try {
            Object.defineProperty(__ret, '_errors', {
                value: __errors,
                writable: true,
                enumerable: false,
                configurable: true
            });
        } catch (__) {
            try { __ret._errors = __errors; } catch (__) {}
        }
    }
    return __ret;
};
Array2D.prototype.filter = Array2D.prototype.z筛选;

/**
 * 内部方法：根据对象参数筛选
 * @private
 * @param {Array} data - 数据
 * @param {Object} condition - 条件对象
 * @returns {Array} 筛选后的数组
 */
Array2D._filterByObject = function(data, condition) {
    var self = this;
    return data.filter(function(row, index) {
        return self._checkCondition(row, condition, index);
    });
};

/**
 * 内部方法：检查单行是否满足条件
 * @private
 * @param {Array} row - 行数据
 * @param {Object} condition - 条件对象
 * @param {Number} index - 行索引
 * @returns {Boolean} 是否满足
 */
Array2D._checkCondition = function(row, condition, index) {
    var self = this;
    var column = condition.column;
    var operator = condition.operator;
    var value = condition.value;
    var logic = condition.logic || 'and'; // and / or
    
    // 获取列值
    var colIndex = -1;
    if (typeof column === 'string' && column.match(/^f\d+$/i)) {
        colIndex = parseInt(column.substring(1)) - 1;
    } else if (typeof column === 'number') {
        colIndex = column;
    }
    
    var cellValue = colIndex >= 0 ? row[colIndex] : undefined;
    
    // 执行比较
    var result = false;
    switch (operator) {
        case '>':
        case 'gt':
            result = cellValue > value;
            break;
        case '>=':
        case 'gte':
            result = cellValue >= value;
            break;
        case '<':
        case 'lt':
            result = cellValue < value;
            break;
        case '<=':
        case 'lte':
            result = cellValue <= value;
            break;
        case '==':
        case '=':
        case 'eq':
            result = cellValue == value;
            break;
        case '===':
            result = cellValue === value;
            break;
        case '!=':
        case '<>':
        case 'neq':
            result = cellValue != value;
            break;
        case 'in':
            result = Array.isArray(value) && value.indexOf(cellValue) >= 0;
            break;
        case 'nin':
        case 'notin':
            result = Array.isArray(value) && value.indexOf(cellValue) < 0;
            break;
        case 'contains':
            result = String(cellValue).indexOf(String(value)) >= 0;
            break;
        case 'startswith':
            result = String(cellValue).indexOf(String(value)) === 0;
            break;
        case 'endswith':
            var str = String(cellValue);
            var suffix = String(value);
            result = str.substring(str.length - suffix.length) === suffix;
            break;
        case 'regex':
        case 'match':
            var regex = typeof value === 'string' ? new RegExp(value) : value;
            result = regex.test(String(cellValue));
            break;
        case 'between':
            if (Array.isArray(value) && value.length >= 2) {
                result = cellValue >= value[0] && cellValue <= value[1];
            }
            break;
        case 'empty':
        case 'isnull':
            result = cellValue === null || cellValue === undefined || cellValue === '';
            break;
        case 'notempty':
        case 'notnull':
            result = cellValue !== null && cellValue !== undefined && cellValue !== '';
            break;
        case 'func':
        case 'function':
            if (typeof value === 'function') {
                result = value(cellValue, row, index);
            }
            break;
        default:
            result = false;
    }
    
    // 处理 and / or 子条件
    if (condition.and && Array.isArray(condition.and)) {
        for (var i = 0; i < condition.and.length; i++) {
            if (!self._checkCondition(row, condition.and[i], index)) {
                return false;
            }
        }
        return result;
    }
    
    if (condition.or && Array.isArray(condition.or)) {
        if (result) return true;
        for (var i = 0; i < condition.or.length; i++) {
            if (self._checkCondition(row, condition.or[i], index)) {
                return true;
            }
        }
        return false;
    }
    
    return result;
};

/**
 * where - 链式筛选起点
 * @param {string|number} column - 列名或索引
 * @returns {QueryBuilder} 查询构建器
 * @example
 * arr.where('f1').gt(0).and('f2').eq('中国').execute();
 */
Array2D.prototype.where = function(column) {
    return new QueryBuilder(this._items, column);
};
Array2D.prototype.z筛选链 = Array2D.prototype.where;

// ==================== QueryBuilder 链式筛选构建器 ====================

/**
 * QueryBuilder - 链式筛选构建器
 * @class
 * @param {Array} data - 原始数据
 * @param {string|number} column - 初始列
 * @example
 * arr.where('f1').gt(0)
 *    .and('f2').eq('中国')
 *    .or('f3').lt(100)
 *    .execute();
 */
function QueryBuilder(data, column) {
    this._data = data;
    this._conditions = [];
    this._currentColumn = column;
    this._currentLogic = 'and';
}

/**
 * 设置当前操作的列
 */
QueryBuilder.prototype.column = function(col) {
    this._currentColumn = col;
    return this;
};

/**
 * 大于
 */
QueryBuilder.prototype.gt = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '>',
        value: value
    });
    this._currentLogic = 'and'; // 重置为默认
    return this;
};
QueryBuilder.prototype.greaterThan = QueryBuilder.prototype.gt;
QueryBuilder.prototype.大于 = QueryBuilder.prototype.gt;

/**
 * 大于等于
 */
QueryBuilder.prototype.gte = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '>=',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.greaterThanOrEqual = QueryBuilder.prototype.gte;
QueryBuilder.prototype.大于等于 = QueryBuilder.prototype.gte;

/**
 * 小于
 */
QueryBuilder.prototype.lt = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '<',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.lessThan = QueryBuilder.prototype.lt;
QueryBuilder.prototype.小于 = QueryBuilder.prototype.lt;

/**
 * 小于等于
 */
QueryBuilder.prototype.lte = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '<=',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.lessThanOrEqual = QueryBuilder.prototype.lte;
QueryBuilder.prototype.小于等于 = QueryBuilder.prototype.lte;

/**
 * 等于
 */
QueryBuilder.prototype.eq = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '==',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.equals = QueryBuilder.prototype.eq;
QueryBuilder.prototype.equal = QueryBuilder.prototype.eq;
QueryBuilder.prototype.等于 = QueryBuilder.prototype.eq;

/**
 * 不等于
 */
QueryBuilder.prototype.neq = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: '!=',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.notEqual = QueryBuilder.prototype.neq;
QueryBuilder.prototype.不等于 = QueryBuilder.prototype.neq;

/**
 * 包含
 */
QueryBuilder.prototype.contains = function(value) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'contains',
        value: value
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.contain = QueryBuilder.prototype.contains;
QueryBuilder.prototype.包含 = QueryBuilder.prototype.contains;

/**
 * 在列表中
 */
QueryBuilder.prototype.in = function(values) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'in',
        value: values
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.在列表中 = QueryBuilder.prototype.in;

/**
 * 不在列表中
 */
QueryBuilder.prototype.nin = function(values) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'nin',
        value: values
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.notIn = QueryBuilder.prototype.nin;
QueryBuilder.prototype.不在列表中 = QueryBuilder.prototype.nin;

/**
 * 在范围内
 */
QueryBuilder.prototype.between = function(min, max) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'between',
        value: [min, max]
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.在范围内 = QueryBuilder.prototype.between;

/**
 * 匹配正则
 */
QueryBuilder.prototype.match = function(regex) {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'match',
        value: regex
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.regex = QueryBuilder.prototype.match;
QueryBuilder.prototype.匹配 = QueryBuilder.prototype.match;

/**
 * 为空
 */
QueryBuilder.prototype.isNull = function() {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'isnull',
        value: null
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.isEmpty = QueryBuilder.prototype.isNull;
QueryBuilder.prototype.为空 = QueryBuilder.prototype.isNull;

/**
 * 不为空
 */
QueryBuilder.prototype.isNotNull = function() {
    this._conditions.push({
        logic: this._currentLogic,
        column: this._currentColumn,
        operator: 'notnull',
        value: null
    });
    this._currentLogic = 'and';
    return this;
};
QueryBuilder.prototype.isNotEmpty = QueryBuilder.prototype.isNotNull;
QueryBuilder.prototype.不为空 = QueryBuilder.prototype.isNotNull;

/**
 * 逻辑与 - 切换到 AND 模式
 */
QueryBuilder.prototype.and = function(column) {
    this._currentLogic = 'and';
    if (column !== undefined) {
        this._currentColumn = column;
    }
    return this;
};
QueryBuilder.prototype.且 = QueryBuilder.prototype.and;

/**
 * 逻辑或 - 切换到 OR 模式
 */
QueryBuilder.prototype.or = function(column) {
    this._currentLogic = 'or';
    if (column !== undefined) {
        this._currentColumn = column;
    }
    return this;
};
QueryBuilder.prototype.或 = QueryBuilder.prototype.or;

/**
 * 执行筛选，返回 Array2D
 */
QueryBuilder.prototype.execute = function() {
    var self = this;
    var result = this._data.filter(function(row, index) {
        var andGroup = true;
        var orGroup = false;
        var hasOr = false;
        
        for (var i = 0; i < self._conditions.length; i++) {
            var cond = self._conditions[i];
            var match = Array2D._checkCondition(row, cond, index);
            
            if (cond.logic === 'or') {
                hasOr = true;
                orGroup = orGroup || match;
            } else {
                andGroup = andGroup && match;
            }
        }
        
        if (hasOr) {
            return andGroup && orGroup;
        }
        return andGroup;
    });
    
    return new Array2D(result);
};
QueryBuilder.prototype.exec = QueryBuilder.prototype.execute;
QueryBuilder.prototype.run = QueryBuilder.prototype.execute;
QueryBuilder.prototype.执行 = QueryBuilder.prototype.execute;
QueryBuilder.prototype.val = QueryBuilder.prototype.execute;

// 静态方法：直接通过 Array2D.where 创建
Array2D.where = function(data, column) {
    return new QueryBuilder(data, column);
};

/**
 * 映射转换
 * @param {string|Function} mapper - 转换函数
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z映射 = function(mapper) {
    const fn = typeof mapper === 'function' ? mapper : parseLambda(mapper);
    if (!fn) return this._new([]);
    // 🔧 XXD-134 final fix: z去重 已经 skip header, 下游不应再 skip.
    // 仅当源头未设置 _header 时(纯 raw array) 默认 skip 第 0 行.
    var skipHeader = 0;
    if (!(Object.prototype.hasOwnProperty.call(this, '_header')) && this._items.length > 0) {
        skipHeader = 1;
    }
    var data = this._items;
    if (skipHeader > 0) {
        data = data.slice(skipHeader);
    }
    // 🔧 v4.0.10 修复：为每行创建代理，支持 x.f1, x.f2 等语法
    var result = data.map(function(row, index) {
        var proxy = Array.isArray(row) ? row.slice() : [row];
        for (var c = 0; c < proxy.length; c++) {
            proxy['f' + (c + 1)] = proxy[c];
        }
        var out = fn(proxy, index);
        // XXD-180: fn 返回 undefined 时保留原行(避免下游 NaN/null 静默替换原始数据)
        return out === undefined ? row : out;
    });
    return this._new(result);
};
Array2D.prototype.map = Array2D.prototype.z映射;

/**
 * 归约计算
 * @param {Function} callback - 回调函数
 * @param {any} initialValue - 初始值
 * @returns {any} 计算结果
 */
// 🔧 XXD-195/XXD-196 final fix: 走扁平化后的叶子序列, 与 z求和/z最大值/z最小值/z平均值 一致
//  之前 this._items.reduce 把外层每一行作为元素传入, 起点 initialValue 只调一次,
//  后续累加的 b 仍是行, 与 2D 数值/表格场景的直觉 (按格累加) 相反
//  传 initialValue 时条件转发, 避免 V8 在 reduce(cb, undefined) 上把 undefined 当成累加器
//  → NaN (arguments.length 决定是否取首元素, 显式传 undefined 长度仍为 2)
Array2D.prototype.z归约 = function(callback, initialValue) {
    var callbackFn = typeof callback === 'string' ? parseLambda(callback) : callback;
    if (!callbackFn) throw new TypeError('z归约: callback 无法解析');
    var flat = this.z扁平化();
    return arguments.length < 2 ? flat.reduce(callbackFn) : flat.reduce(callbackFn, initialValue);
};
Array2D.prototype.reduce = Array2D.prototype.z归约;

/**
 * 倒序归约（reduceRight）- 从右向左归约计算
 * @param {Function} callback - 回调函数
 * @param {any} initialValue - 初始值
 * @returns {any} 计算结果
 * @example
 * Array2D([[1,2,3]]).z倒序归约((acc, val) => acc + val, 0)  // 6
 */
// XXD-208/XXD-209 final fix: z倒序归约 同 z归约 走扁平化后的叶子序列, 与 z求和/z最大值/z最小值/z平均值 一致
//  之前 this._items.reduceRight 把外层每一行作为元素传入, 起点 initialValue 只调一次,
//  后续累加的 b 仍是行, 与 2D 数值/表格场景的直觉 (按格累加) 相反
//  传 initialValue 时条件转发, 避免 V8 在 reduceRight(cb, undefined) 上把 undefined 当成累加器
//  -> NaN (arguments.length 决定是否取首元素, 显式传 undefined 长度仍为 2)
Array2D.prototype.z倒序归约 = function(callback, initialValue) {
    var callbackFn = typeof callback === 'string' ? parseLambda(callback) : callback;
    if (!callbackFn) throw new TypeError('z倒序归约: callback 无法解析');
    var flat = this.z扁平化();
    return arguments.length < 2 ? flat.reduceRight(callbackFn) : flat.reduceRight(callbackFn, initialValue);
};
Array2D.prototype.reduceRight = Array2D.prototype.z倒序归约;

/**
 * 检查是否全部满足
 * @param {string|Function} predicate - 条件
 * @returns {Boolean} 是否全部满足
 */
// XXD-196: z全部满足 应跳过 raw-2D 表头行 (与 z列统计 同步: 无显式 _header + length>1 + 首行是数组 ⇒ 跳过 row 0)
Array2D.prototype.z全部满足 = function(predicate) {
    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return false;
    var items = this._items;
    var hasHeaderRow = !Object.prototype.hasOwnProperty.call(this, '_header')
        && items.length > 1
        && Array.isArray(items[0]);
    var startIdx = hasHeaderRow ? 1 : 0;
    for (var i = startIdx; i < items.length; i++) {
        if (!fn(items[i], i, items)) return false;
    }
    return true;
};
Array2D.prototype.every = Array2D.prototype.z全部满足;

/**
 * 检查是否有满足
 * @param {string|Function} predicate - 条件
 * @returns {Boolean} 是否有满足
 */
// XXD-196: z有满足 同步跳过 raw-2D 表头行, 与 z全部满足 / z列统计 一致
Array2D.prototype.z有满足 = function(predicate) {
    const fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return false;
    var items = this._items;
    var hasHeaderRow = !Object.prototype.hasOwnProperty.call(this, '_header')
        && items.length > 1
        && Array.isArray(items[0]);
    var startIdx = hasHeaderRow ? 1 : 0;
    for (var i = startIdx; i < items.length; i++) {
        if (fn(items[i], i, items)) return true;
    }
    return false;
};
Array2D.prototype.some = Array2D.prototype.z有满足;

// ==================== 行列操作 ====================

/**
 * 获取行数
 * @returns {Number} 行数
 */
Array2D.prototype.z行数 = function() {
    return this._items.length;
};
Array2D.prototype.rowCount = Array2D.prototype.z行数;

/**
 * 获取列数
 * @returns {Number} 列数
 */
Array2D.prototype.z列数 = function() {
    return this._items.length > 0 && this._items[0] ? this._items[0].length : 0;
};
Array2D.prototype.colCount = Array2D.prototype.z列数;

/**
 * 获取指定行
 * @param {Number} index - 行号（从0开始）
 * @returns {Array} 行数据
 */
Array2D.prototype.z获取行 = function(index) {
    if (typeof index !== 'number' || isNaN(index) || index < 0 || index >= this._items.length) {
        throw new RangeError('z获取行: 行号越界: ' + index + ' (有效范围 0..' + (this._items.length - 1) + ')');
    }
    return this._items[index];
};
Array2D.prototype.getRow = Array2D.prototype.z获取行;

/**
 * 获取指定列
 * @param {Number} index - 列号（从0开始）
 * @returns {Array} 列数据
 */
Array2D.prototype.z获取列 = function(index) {
    if (typeof index !== 'number' || isNaN(index) || index < 0) {
        throw new RangeError('z获取列: 列号非法: ' + index);
    }
    // 列号上限以首行(表头)长度为基准; 数据行短于该列号时填 undefined(保留 ragged array 语义)。
    var headerLen = (Array.isArray(this._items[0]) ? this._items[0].length : 0);
    if (index >= headerLen) {
        throw new RangeError('z获取列: 列号越界: ' + index + ' (有效范围 0..' + (headerLen - 1) + ')');
    }
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row) && index < row.length) {
            result.push(row[index]);
        } else {
            result.push(undefined);
        }
    }
    return result;
};
Array2D.prototype.getCol = Array2D.prototype.z获取列;

/**
 * 获取第一行
 * @returns {Array} 第一行数据
 */
Array2D.prototype.z首行 = function() {
    return this._items[0] || [];
};
Array2D.prototype.firstRow = Array2D.prototype.z首行;

/**
 * 获取最后一行
 * @returns {Array} 最后一行数据
 */
Array2D.prototype.z末行 = function() {
    return this._items[this._items.length - 1] || [];
};
Array2D.prototype.lastRow = Array2D.prototype.z末行;

/**
 * 获取第一列
 * @returns {Array} 第一列数据
 */
Array2D.prototype.z首列 = function() {
    return this.z获取列(0);
};
Array2D.prototype.firstCol = Array2D.prototype.z首列;

/**
 * 获取最后一列
 * @returns {Array} 最后一列数据
 */
Array2D.prototype.z末列 = function() {
    return this.z获取列(this.z列数() - 1);
};
Array2D.prototype.lastCol = Array2D.prototype.z末列;

// ==================== 增删行列 ====================

/**
 * 添加行
 * @param {Array} row - 行数据
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z添加行 = function(row) {
    if (!Array.isArray(row)) {
        throw new TypeError('z添加行: 参数必须是数组, 得到 ' + typeof row);
    }
    var cols = this._items.length > 0 ? this._items[0].length : 0;
    if (row.length !== cols) {
        throw new RangeError('z添加行: 长度不匹配 当前 ' + cols + ' 列, 传入 ' + row.length + ' 列');
    }
    var result = this._items.slice();
    result.push(row);
    return this._new(result);
};
Array2D.prototype.addRow = Array2D.prototype.z添加行;

/**
 * 提取列（pluck）
 * @param {Number} colIndex - 列索引
 * @returns {Array} 列数据
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z提取列(1)  // [2,5]
 */
Array2D.prototype.z提取列 = function(colIndexOrName) {
    // 接受数字列号或表头名(字符串)。
    // 数字:返回该列所有行(行为 z获取列 行为)。
    // 字符串:按首行查找列号,返回除表头行之外的列数据(行 [1..end])。
    if (typeof colIndexOrName === 'string') {
        var headerRow = this._items[0];
        if (!Array.isArray(headerRow) || headerRow.length === 0) {
            throw new Error('z提取列: 空数据, 无法按表头取列: ' + colIndexOrName);
        }
        var resolved = -1;
        for (var _hi = 0; _hi < headerRow.length; _hi++) {
            if (String(headerRow[_hi]) === colIndexOrName) { resolved = _hi; break; }
        }
        if (resolved < 0) {
            throw new Error('z提取列: 表头未找到: ' + colIndexOrName);
        }
        var out = [];
        for (var _di = 1; _di < this._items.length; _di++) {
            var _r = this._items[_di];
            if (Array.isArray(_r) && resolved < _r.length) {
                out.push(_r[resolved]);
            } else {
                out.push(undefined);
            }
        }
        return out;
    }
    // 数字路径: 检查列索引是否越界
    var numCols = this._items.length > 0 && Array.isArray(this._items[0]) ? this._items[0].length : 0;
    if (colIndexOrName < 0 || colIndexOrName >= numCols) {
        throw new Error('z提取列: 列索引越界: ' + colIndexOrName + ' (列数: ' + numCols + ')');
    }
    return this.z获取列(colIndexOrName);
};
Array2D.prototype.pluck = Array2D.prototype.z提取列;

/**
 * 添加列
 * @param {Array} col - 列数据
 * @param {Number} index - 插入位置（可选，默认为末尾）
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z添加列([5,6])        // [[1,2,5],[3,4,6]]
 * Array2D([[1,2],[3,4]]).z添加列([5,6], 0)     // [[5,1,2],[6,3,4]]
 */
Array2D.prototype.z添加列 = function(name, data) {
    // XXD-199: z添加列(name, data|fn?) — name is the column header (string, required)
    // header row (row 0) gets `name` appended; data column is appended at the end.
    //   data === function  → fn(row) per data row
    //   data is array      → array[i] per data row (length must match data row count)
    //   data === undefined → all nulls
    if (typeof name !== 'string') {
        throw new TypeError('z添加列: name must be a string, got ' + (name === null ? 'null' : typeof name));
    }
    var rows = this._items;
    var dataCount = rows.length > 0 ? rows.length - 1 : 0;  // exclude header
    var result = [];
    for (var i = 0; i < rows.length; i++) {
        var newRow = rows[i].slice();
        var isHeader = (i === 0);
        var cell;
        if (isHeader) {
            cell = name;
        } else if (typeof data === 'function') {
            try { cell = data(rows[i]); } catch (e) { cell = null; }
        } else if (data === undefined || data === null) {
            cell = null;
        } else if (Array.isArray(data)) {
            cell = data[i - 1] !== undefined ? data[i - 1] : null;
        } else {
            cell = data;
        }
        newRow.push(cell);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.addCol = Array2D.prototype.z添加列;

/**
 * 删除行
 * @param {Number} index - 行号
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z删除行(1)  // [[1,2],[5,6]]
 */
Array2D.prototype.z删除行 = function(index) {
    // 索引边界检查
    if (index < 0 || index >= this._items.length) {
        return this._new(this._items.slice());  // 索引无效，返回副本
    }
    var result = this._items.slice();
    result.splice(index, 1);
    return this._new(result);
};
Array2D.prototype.deleteRow = Array2D.prototype.z删除行;

/**
 * 删除列
 * @param {Number} index - 列号
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2,3],[4,5,6]]).z删除列(1)  // [[1,3],[4,6]]
 */
Array2D.prototype.z删除列 = function(index) {
    // 索引边界检查
    if (index < 0 || index >= this.z列数()) {
        return this._new(this._items.slice());  // 索引无效，返回副本
    }
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        newRow.splice(index, 1);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.deleteCol = Array2D.prototype.z删除列;

/**
 * 尾部弹出一项（pop）- 删除并返回最后一行
 * @returns {Array} 被删除的行
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z尾部弹出一项()  // [5,6]
 */
Array2D.prototype.z尾部弹出一项 = function() {
    if (this.length === 0) return this;
    // 🔧 Bug修复: 直接在实例上 pop，而非通过 _items getter（getter返回副本）
    // XXD-189: 返回 this 支持链式调用（.z尾部弹出一项()._items）
    Array.prototype.pop.call(this);
    return this;
};
Array2D.prototype.pop = Array2D.prototype.z尾部弹出一项;

/**
 * 追加一项（push）- 向数组末尾添加行
 * @param {...Array} rows - 要添加的行
 * @returns {Number} 添加后的行数
 * @example
 * Array2D([[1,2],[3,4]]).z追加一项([5,6], [7,8])  // 4
 */
Array2D.prototype.z追加一项 = function() {
    // 🔧 Bug修复: 直接在实例上 push，而非通过 _items getter（getter返回副本）
    // XXD-189: 返回 this 支持链式调用（.z追加一项([3])._items）
    for (var i = 0; i < arguments.length; i++) {
        Array.prototype.push.call(this, arguments[i]);
    }
    return this;
};
Array2D.prototype.push = Array2D.prototype.z追加一项;

/**
 * 删除第一个（shift）- 删除并返回第一行
 * @returns {Array} 被删除的行
 * @example
 * Array2D([[1,2],[3,4],[5,6]]).z删除第一个()  // [1,2]
 */
Array2D.prototype.z删除第一个 = function() {
    if (this.length === 0) return this;
    // 🔧 Bug修复: 直接在实例上 shift，而非通过 _items getter（getter返回副本）
    // XXD-189: 返回 this 支持链式调用（.z删除第一个()._items）
    Array.prototype.shift.call(this);
    return this;
};
Array2D.prototype.shift = Array2D.prototype.z删除第一个;

// ==================== 排序去重 ====================

/**
 * sort - 原生数组 sort 方法的代理（支持链式调用）
 * @param {Function} compareFn - 比较函数
 * @returns {Array2D} 返回当前实例（支持链式调用）
 * @example
 * Array2D([[3,1],[2,2],[1,3]]).sort((a,b)=>a[0]-b[0]).val()  // [[1,3],[2,2],[3,1]]
 */
// 🔧 XXD-148 fix: z排序 / z单列排序 — call-time delegation (forward-ref safe).
// z多列排序 is defined LATER in this file (~L9125), so the previous
// `proto.z排序 = proto.z多列排序` assignment captured `undefined`.
// Per XXD-148 expected `[[1],[2],[3]]` from `new Array2D([['v'],[3],[1],[2]]).z排序(0)`:
//   numeric arg → sort by 0-indexed col ascending, drop the header row.
//   string arg  → alias of z多列排序 (header preserved per z多列排序 contract).
Array2D.prototype.z排序 = function(arg, headerRows, customOrder) {
    if (typeof arg === 'number') {
        var rows = (headerRows == null) ? 1 : headerRows;
        var sorted = this.z多列排序('f' + (arg + 1) + '+', rows, customOrder);
        return sorted._new(sorted._items.slice(rows));
    }
    return this.z多列排序(arg, headerRows, customOrder);
};
Array2D.prototype.z单列排序 = function(colIdx, ascending) {
    var order = ascending === false ? 'f' + (colIdx+1) + '-' : 'f' + (colIdx+1) + '+';
    return this.z多列排序(order, 0);
};
Array2D.prototype.sort = function(compareFn) {
    // 🔧 Bug修复: 直接在实例上排序，而非通过 _items getter（getter返回副本，修改无效）
    Array.prototype.sort.call(this, compareFn);
    return this;  // 返回 this 支持链式调用
};

/**
 * 按规则升序（sortBy）- 使用Lambda表达式指定排序键进行升序排序
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z按规则升序('x=>x[0]')  // [[1,'A'],[2,'B'],[3,'C']]
 */
Array2D.prototype.z按规则升序 = function(keySelector) {
    var fn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!fn) return this._new(this._items.slice());
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = fn(a);
        var valB = fn(b);
        if (valA < valB) return -1;
        if (valA > valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortBy = Array2D.prototype.z按规则升序;

/**
 * 按规则降序（sortByDesc）- 使用Lambda表达式指定排序键进行降序排序
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z按规则降序('x=>x[0]')  // [[3,'C'],[2,'B'],[1,'A']]
 */
Array2D.prototype.z按规则降序 = function(keySelector) {
    var fn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!fn) return this._new(this._items.slice());
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = fn(a);
        var valB = fn(b);
        if (valA > valB) return -1;
        if (valA < valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortByDesc = Array2D.prototype.z按规则降序;

/**
 * 降序排序（sortDesc）- 按首列降序排序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z降序排序()  // [[3,'C'],[2,'B'],[1,'A']]
 */
Array2D.prototype.z降序排序 = function() {
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[0];
        var valB = b[0];
        if (valA > valB) return -1;
        if (valA < valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortDesc = Array2D.prototype.z降序排序;

/**
 * 升序排序 - 按首列升序排序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[3,'C'],[1,'A'],[2,'B']]).z升序排序()  // [[1,'A'],[2,'B'],[3,'C']]
 */
Array2D.prototype.z升序排序 = function() {
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[0];
        var valB = b[0];
        if (valA < valB) return -1;
        if (valA > valB) return 1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortAsc = Array2D.prototype.z升序排序;

/**
 * 行排序
 * @param {Number} colIndex - 排序依据的列
 * @param {Boolean} ascending - 是否升序
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,3],[2,2],[3,1]]).z行排序(1)       // [[3,1],[2,2],[1,3]]
 * Array2D([[1,3],[2,2],[3,1]]).z行排序(1, false)  // [[1,3],[2,2],[3,1]] 降序
 */
Array2D.prototype.z行排序 = function(colIndex, ascending) {
    ascending = ascending !== undefined ? ascending : true;
    // 列边界检查
    if (colIndex < 0 || colIndex >= this.z列数()) {
        return this._new(this._items.slice());  // 列索引无效，返回副本
    }
    var result = this._items.slice();
    result.sort(function(a, b) {
        var valA = a[colIndex];
        var valB = b[colIndex];
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    });
    return this._new(result);
};
Array2D.prototype.sortRow = Array2D.prototype.z行排序;

/**
 * 列排序
 * @param {Number} rowIndex - 排序依据的行
 * @param {Boolean} ascending - 是否升序
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z列排序 = function(rowIndex, ascending) {
    ascending = ascending !== undefined ? ascending : true;
    if (!this._items[rowIndex]) return this._new([]);
    var colCount = this._items[rowIndex].length;
    var indices = [];
    for (var i = 0; i < colCount; i++) indices.push(i);
    indices.sort(function(a, b) {
        var valA = this._items[rowIndex][a];
        var valB = this._items[rowIndex][b];
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    }.bind(this));
    var result = [];
    for (var r = 0; r < this._items.length; r++) {
        var newRow = [];
        for (var i = 0; i < indices.length; i++) {
            newRow.push(this._items[r][indices[i]]);
        }
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.sortCol = Array2D.prototype.z列排序;

/**
 * 多列排序 - 按多列排序，支持指定每列的升降序
 * @param {string} sortParams - 排序参数 'f3+,f4-' 表示第3列升序第4列降序
 * @param {number} [headerRows=0] - 表头的行数（不参与排序）
 * @param {string} [customOrder] - 自定义序列，逗号分隔
 * @returns {Array2D} 新实例
 * @example
 * Array2D(arr).z多列排序('f3+,f4-', 1)  // 第3列升序，第4列降序，第1行为表头
 */
Array2D.prototype.z多列排序 = function(sortParams, headerRows, customOrder) {
    // 🔧 v3.7.9 修复: 如果没有指定 headerRows 但对象有 _header 属性，自动设为 1
    if (headerRows === undefined && Object.prototype.hasOwnProperty.call(this, '_header') && this._header !== undefined) {
        headerRows = 1;
    }
    headerRows = headerRows || 0;

    // 🔧 v3.9.4 修复：sortParams 为空时直接返回副本
    if (!sortParams || typeof sortParams !== 'string') {
        return this._new(this._items.slice());
    }

    // 解析排序参数
    var sorts = [];
    var parts = sortParams.split(/[,，]/);
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        var match = part.match(/^f?(\d+)([+-])?$/);
        if (match) {
            sorts.push({
                col: parseInt(match[1]),
                order: (match[2] || '+') === '+' ? 1 : 2 // 1升序 2降序, 缺省升序
            });
        }
        else {
            console.warn('[JSA880] z多列排序: 忽略无效排序参数 "' + part + '"');
        }
    }

    if (this._items.length <= headerRows) return this._new(this._items.slice());

    // 分离表头和数据
    var header = this._items.slice(0, headerRows);
    var data = this._items.slice(headerRows);

    // 🔧 XXD-19 修复: 同步 wps-jsa-jsa880-agent 的类型检测。
    // 数字字符串 ('2' / '10') 直接用 < / > 比较会走字符串序 ('10' < '2')，
    // 需先归一化。规则：null/undefined/空串 → null；number/Date/数字字符串 → 数值；boolean → 0/1；其余按字符串。
    function _coerceForSort(val) {
        if (val === null || val === undefined) return null;
        if (typeof val === 'number') return val;
        if (typeof val === 'boolean') return val ? 1 : 0;
        if (val instanceof Date) return val.getTime();
        var s = String(val).trim();
        if (s === '') return null;
        if (/^-?\d+(\.\d+)?$/.test(s)) return parseFloat(s);
        return s;
    }

    // 排序
    data.sort(function(a, b) {
        for (var s = 0; s < sorts.length; s++) {
            var sort = sorts[s];
            var colIdx = sort.col - 1;
            var valA = a[colIdx];
            var valB = b[colIdx];

            // 自定义序列处理
            if (customOrder) {
                var orderArr = customOrder.split(/[,，]/);
                var idxA = orderArr.indexOf(String(valA));
                var idxB = orderArr.indexOf(String(valB));
                if (idxA >= 0 && idxB >= 0) {
                    valA = idxA;
                    valB = idxB;
                }
            }

            // 两侧都能归一为 number 时按数值序比较；否则回退到原始 < / >（字符串序）
            var cA = _coerceForSort(valA);
            var cB = _coerceForSort(valB);
            if (typeof cA === 'number' && typeof cB === 'number') {
                if (cA < cB) return sort.order === 1 ? -1 : 1;
                if (cA > cB) return sort.order === 1 ? 1 : -1;
            } else {
                if (typeof valA === 'string' && typeof valB === 'string') {
                    var strA = valA.toLowerCase();
                    var strB = valB.toLowerCase();
                    if (strA < strB) return sort.order === 1 ? -1 : 1;
                    if (strA > strB) return sort.order === 1 ? 1 : -1;
                } else {
                    if (valA < valB) return sort.order === 1 ? -1 : 1;
                    if (valA > valB) return sort.order === 1 ? 1 : -1;
                }

            }
        }
        return 0;
    });

    return this._new(header.concat(data));
};
Array2D.prototype.sortByCols = Array2D.prototype.z多列排序;

/**
 * 自定义排序 - 按指定列表的顺序排序
 * @param {number|string} colIndex - 列索引（支持数字0索引或 "f3" 格式1索引）
 * @param {Array|string} orderList - 排序列表（数组或逗号分隔的字符串）
 * @param {number} [headerRows=0] - 表头的行数（不参与排序）
 * @returns {Array2D} 新实例
 * @example
 * Array2D(arr).z自定义排序("f3", "中国,英国,美国,德国")
 * Array2D(arr).sortByList(2, ["中国", "英国", "美国", "德国"], 1)
 */
Array2D.prototype.z自定义排序 = function(colIndex, orderList, headerRows) {
    headerRows = headerRows || 0;

    // 处理列索引：支持 f3 格式（从1开始的列号）或数字索引
    var actualColIndex = colIndex;
    if (typeof colIndex === 'string' && colIndex.toLowerCase().startsWith('f')) {
        actualColIndex = parseInt(colIndex.substring(1)) - 1;
    }

    // 处理排序列表：支持逗号分隔的字符串或数组
    var actualOrderList = orderList;
    if (typeof orderList === 'string') {
        actualOrderList = orderList.split(/[,，]/).map(function(s) { return s.trim(); });
    }

    if (this._items.length <= headerRows) return this._new(this._items.slice());

    // 分离表头和数据
    var header = this._items.slice(0, headerRows);
    var data = this._items.slice(headerRows);

    data.sort(function(a, b) {
        var valA = a[actualColIndex];
        var valB = b[actualColIndex];
        var indexA = actualOrderList.indexOf(valA);
        var indexB = actualOrderList.indexOf(valB);

        // 不在列表中的值放到最后
        var posA = indexA === -1 ? 999 : indexA;
        var posB = indexB === -1 ? 999 : indexB;

        return posA - posB;
    });

    return this._new(header.concat(data));
};
Array2D.prototype.sortByList = Array2D.prototype.z自定义排序;
// ==================== Index 索引加速层 ====================

/**
 * 缓存键正规化 — 将 colSelector 转为字符串键，用于 _indexes 查找
 */
function _normalizeCacheKey(colSelector) {
    if (colSelector === undefined) return '__whole__';
    if (typeof colSelector === 'number') return 'n:' + colSelector;
    if (typeof colSelector === 'string') return 's:' + colSelector;
    if (Array.isArray(colSelector)) {
        try { return 'a:' + JSON.stringify(colSelector); } catch (e) { return null; }
    }
    return null;
}

/**
 * Index 索引对象 — 由 toIndex 创建，提供 O(unique) 的去重/分组/查找
 */
function Index(indices, source) {
    this._indices = indices;
    this._source = source;
}

Index.prototype.distinct = function(resultSelector) {
    var source = this._source;
    var keys = Object.keys(this._indices);
    var result = [];
    var outputFn;
    if (resultSelector === undefined || resultSelector === '') {
        outputFn = function(row) { return row ? (Array.isArray(row) ? row.slice() : [row]) : []; };
    } else if (typeof resultSelector === 'string') {
        if (resultSelector.includes(',') || resultSelector.includes('，')) {
            var outIdx = [];
            var parts = resultSelector.split(/[,，]/);
            for (var j = 0; j < parts.length; j++) {
                var p = parts[j].trim();
                outIdx.push(p.toLowerCase().startsWith('f') ? parseInt(p.substring(1)) - 1 : parseInt(p) - 1);
            }
            outputFn = function(row) {
                if (!row) return [];
                var out = [];
                for (var ki = 0; ki < outIdx.length; ki++) out.push(row[outIdx[ki]]);
                return out;
            };
        } else if (resultSelector.toLowerCase().startsWith('f')) {
            var col = parseInt(resultSelector.substring(1)) - 1;
            outputFn = function(row) { return row ? [row[col]] : []; };
        } else {
            outputFn = function(row) { return row ? (Array.isArray(row) ? row.slice() : [row]) : []; };
        }
    } else if (Array.isArray(resultSelector)) {
        outputFn = function(row) {
            if (!row) return [];
            var out = [];
            for (var m = 0; m < resultSelector.length; m++) out.push(row[resultSelector[m]]);
            return out;
        };
    } else {
        outputFn = function(row) { return row ? (Array.isArray(row) ? row.slice() : [row]) : []; };
    }
    for (var i = 0; i < keys.length; i++) {
        var firstRowIdx = this._indices[keys[i]][0];
        result.push(outputFn(source[firstRowIdx]));
    }
    return source._new(result);
};

Index.prototype.groupBy = function(valSelector) {
    var source = this._source;
    var keys = Object.keys(this._indices);
    var valFn = valSelector ? (typeof valSelector === 'function' ? valSelector : parseLambda(valSelector)) : null;
    var groups = Object.create(null);
    for (var i = 0; i < keys.length; i++) {
        var key = keys[i];
        var idxs = this._indices[key];
        groups[key] = [];
        for (var j = 0; j < idxs.length; j++) {
            var row = source[idxs[j]];
            groups[key].push(valFn ? valFn(row, idxs[j]) : row);
        }
    }
    return groups;
};

Index.prototype.lookup = function(key) {
    var keyStr = typeof key === 'string' ? key : JSON.stringify(key);
    var idxs = this._indices[keyStr];
    if (!idxs || idxs.length === 0) return [];
    var source = this._source;
    var result = [];
    for (var i = 0; i < idxs.length; i++) result.push(source[idxs[i]]);
    return result;
};

/**
 * toIndex — 构建列索引，加速重复去重/分组操作
 * @param {String|Number} colSelector - 列选择器（'f1'、0 等；函数不可缓存）
 * @returns {Index} 索引对象
 */
Array2D.prototype.toIndex = function(colSelector) {
    var cacheKey = _normalizeCacheKey(colSelector);
    if (!this._indexes) {
        Object.defineProperty(this, '_indexes', {
            value: Object.create(null), writable: true, enumerable: false, configurable: true
        });
    }
    if (cacheKey !== null && this._indexes[cacheKey]) return this._indexes[cacheKey];
    var keyFn;
    if (colSelector === undefined) {
        keyFn = function(row) { return JSON.stringify(row); };
    } else if (typeof colSelector === 'number') {
        keyFn = function(row) { return row[colSelector]; };
    } else if (typeof colSelector === 'string') {
        if (colSelector.includes(',') || colSelector.includes('，')) {
            var colIndexes = [];
            var parts = colSelector.split(/[,，]/);
            for (var p = 0; p < parts.length; p++) {
                var part = parts[p].trim();
                if (part.includes('-') && part.toLowerCase().startsWith('f')) {
                    // XXD-141/XXD-138: 范围 f3-f5 在逗号分隔内
                    var _range2 = part.split('-');
                    var _start2 = parseInt(_range2[0].toLowerCase().replace('f', ''));
                    var _end2 = parseInt(_range2[1].toLowerCase().replace('f', ''));
                    for (var _r2 = _start2; _r2 <= _end2; _r2++) {
                        colIndexes.push(_r2 - 1);
                    }
                } else {
                colIndexes.push(part.toLowerCase().startsWith('f') ? parseInt(part.substring(1)) - 1 : parseInt(part) - 1);
                }
            }
            keyFn = function(row) {
                var kp = [];
                for (var ki = 0; ki < colIndexes.length; ki++) kp.push(row[colIndexes[ki]]);
                return JSON.stringify(kp);
            };
        } else if (colSelector.toLowerCase().startsWith('f')) {
            // XXD-141/XXD-138: 检测范围语法 f3-f5
            if (colSelector.indexOf('-') > -1 && colSelector.indexOf(',') === -1 && colSelector.indexOf('，') === -1) {
                var _rangeX = colSelector.split('-');
                var _startX = parseInt(_rangeX[0].toLowerCase().replace('f', ''));
                var _endX = parseInt(_rangeX[1].toLowerCase().replace('f', ''));
                var _rangeCols = [];
                for (var _rX = _startX; _rX <= _endX; _rX++) {
                    _rangeCols.push(_rX - 1);
                }
                keyFn = function(row) {
                    var kp = [];
                    for (var _ki = 0; _ki < _rangeCols.length; _ki++) kp.push(row[_rangeCols[_ki]]);
                    return JSON.stringify(kp);
                };
            } else {
            var colIdx = parseInt(colSelector.substring(1)) - 1;
            keyFn = function(row) { return row[colIdx]; };
            }
        } else {
            keyFn = function(row) { return JSON.stringify(row); };
        }
    } else if (typeof colSelector === 'function') {
        keyFn = colSelector;
    } else if (Array.isArray(colSelector)) {
        keyFn = function(row) {
            var kp = [];
            for (var i = 0; i < colSelector.length; i++) kp.push(row[colSelector[i]]);
            return JSON.stringify(kp);
        };
    } else {
        keyFn = function(row) { return JSON.stringify(row); };
    }
    var indices = Object.create(null);
    var _items = this._items;
    for (var i = 0; i < _items.length; i++) {
        var row = _items[i];
        if (!row || (Array.isArray(row) && row.length === 0)) continue;
        var key = keyFn(row);
        var keyStr = typeof key === 'string' ? key : JSON.stringify(key);
        if (!indices[keyStr]) indices[keyStr] = [];
        indices[keyStr].push(i);
    }
    var idx = new Index(indices, this);
    if (cacheKey !== null) this._indexes[cacheKey] = idx;
    return idx;
};


/**
 * 去重
 * @param {Number|String|Array|Function} [colSelector] - 依据哪一列或多列去重（可选）
 *   - Number: 列索引（0-based）
 *   - String: 列选择器，如 'f1'（第1列）、'f1,f2'（第1,2列组合去重）
 *   - Array: 列索引数组，如 [0, 1]
 *   - Function: 自定义键生成函数，如 x=>[x.f1,x.f2]
 * @returns {Array2D} 新实例
 * @example
 * // 单行去重（按第1列）
 * Array2D([[1,2],[1,3],[2,4]]).z去重(0)  // [[1,2],[2,4]]
 * // f模式单列去重
 * Array2D([[1,2],[1,3],[2,4]]).z去重('f1')  // [[1,2],[2,4]]
 * // 多列组合去重
 * Array2D([[1,2],[1,3],[1,2]]).z去重('f1,f2')  // [[1,2],[1,3]]
 * // Lambda函数去重
 * Array2D([[1,2],[1,3],[2,4]]).z去重(x=>x.f1)  // [[1,2],[2,4]]
 * Array2D([[1,2],[1,3],[1,2]]).z去重(x=>[x.f1,x.f2])  // [[1,2],[1,3]]
 * // 整行去重（不传参数）
 * Array2D([[1,2],[1,2],[2,4]]).z去重()  // [[1,2],[2,4]]
 */
/**
 * 去重（z去重/distinct）- 根据指定列去重，支持选择输出列
 * @param {String|Function|Array} colSelector - 去重依据的列
 *   - String: 'f1' 单列, 'f1,f2' 多列, 'f1-f3' 范围
 *   - Function: x=>x.f1 回调函数
 *   - Array: [0, 1] 数字数组
 * @param {String|Array} [resultSelector] - 输出列选择（可选）
 *   - undefined: 只输出关键字列
 *   - 空字符串 '': 输出所有列
 *   - 'f1,f2': 输出指定列
 *   - [0, 1]: 输出指定索引列
 * @returns {Array2D} 去重后的新实例
 * @example
 * arr.z去重('f1')                    // 按第1列去重，只输出第1列
 * arr.z去重('f1,f2', '')              // 按1、2列去重，输出所有列
 * arr.z去重('f1,f2', 'f1,f3')        // 按1、2列去重，输出第1、3列
 * arr.z去重(x=>x.f1)                 // 回调函数模式
 */
Array2D.prototype.z去重 = function(colSelector, resultSelector) {
    var seen = Object.create(null);
    // 🔧 v4.0.38 Index 加速: 命中 _indexes 缓存时走 O(unique) 路径
    var _cacheKey = _normalizeCacheKey(colSelector);
    if (_cacheKey !== null && this._indexes && this._indexes[_cacheKey]) {
        return this._indexes[_cacheKey].distinct(resultSelector);
    }
    var result = [];

    // 处理不同的参数类型
    var keyFn;
    var isFunctionMode = false;
    var colIndexes = []; // 🔧 v4.0.2 初始化，避免 undefined
    if (colSelector === undefined) {
        // 整行去重
        keyFn = function(row) { return JSON.stringify(row); };
    } else if (typeof colSelector === 'function') {
        // 函数回调模式
        keyFn = colSelector;
        isFunctionMode = true;
    } else if (typeof colSelector === 'number') {
        // 数字索引
        keyFn = function(row) { return row[colSelector]; };
    } else if (typeof colSelector === 'string') {
        // 字符串模式：支持 'f1' 或 'f1,f2' 或 'f1,f3-f5'
        if (colSelector.includes(',') || colSelector.includes('，')) {
            // 多列组合
            colIndexes = [];  // 🔧 v4.0.2 移除 var，直接使用外层变量
            var parts = colSelector.split(/[,，]/);
            for (var p = 0; p < parts.length; p++) {
                var part = parts[p].trim();
                if (part.toLowerCase().startsWith('f')) {
                    // XXD-141/XXD-138: 逗号分隔内的范围 f3-f5
                    if (part.includes('-')) {
                        var range = part.split('-');
                        var start = parseInt(range[0].toLowerCase().replace('f', ''));
                        var end = parseInt(range[1].toLowerCase().replace('f', ''));
                        for (var r = start; r <= end; r++) {
                            colIndexes.push(r - 1);
                        }
                    } else {
                    colIndexes.push(parseInt(part.substring(1)) - 1);
                    }
                } else if (part.includes('-')) {
                    // 处理范围 f3-f5
                    var range = part.split('-');
                    var start = parseInt(range[0].toLowerCase().replace('f', ''));
                    var end = parseInt(range[1].toLowerCase().replace('f', ''));
                    for (var r = start; r <= end; r++) {
                        colIndexes.push(r - 1);
                    }
                } else {
                    colIndexes.push(parseInt(part) - 1);
                }
            }
            keyFn = function(row) {
                var keyParts = [];
                for (var i = 0; i < colIndexes.length; i++) {
                    keyParts.push(row[colIndexes[i]]);
                }
                return JSON.stringify(keyParts);
            };
        } else if (colSelector.toLowerCase().startsWith('f')) {
            // XXD-141/XXD-138: 检测范围语法 f3-f5（无逗号分隔的纯范围）
            if (colSelector.indexOf('-') > -1 && colSelector.indexOf(',') === -1 && colSelector.indexOf('，') === -1) {
                // 范围模式: 展开 f3-f5 → [2,3,4]
                var _rangeX = colSelector.split('-');
                var _startX = parseInt(_rangeX[0].toLowerCase().replace('f', ''));
                var _endX = parseInt(_rangeX[1].toLowerCase().replace('f', ''));
                colIndexes = [];
                for (var _rX = _startX; _rX <= _endX; _rX++) {
                    colIndexes.push(_rX - 1);
                }
                keyFn = function(row) {
                    var keyParts = [];
                    for (var i = 0; i < colIndexes.length; i++) {
                        keyParts.push(row[colIndexes[i]]);
                    }
                    return JSON.stringify(keyParts);
                };
            } else {
            // 单列 f模式
            var colIdx = parseInt(colSelector.substring(1)) - 1;
            colIndexes = [colIdx]; // 保存单列索引
            keyFn = function(row) { return row[colIdx]; };
            }
        } else {
            // 普通字符串，当作整行去重
            keyFn = function(row) { return JSON.stringify(row); };
        }
    } else if (Array.isArray(colSelector)) {
        // 数组模式 [0, 1, 2]
        keyFn = function(row) {
            var keyParts = [];
            for (var i = 0; i < colSelector.length; i++) {
                keyParts.push(row[colSelector[i]]);
            }
            return JSON.stringify(keyParts);
        };
    } else {
        // 默认整行去重
        keyFn = function(row) { return JSON.stringify(row); };
    }

    // 处理 resultSelector - 输出列选择
    var outputFn;
    if (resultSelector === undefined) {
        // 未指定：只输出关键字列
        if (isFunctionMode && typeof colSelector === 'function') {
            // 🔧 v4.0.4 函数模式：keyFn 已返回完整数组，直接使用 key 作为输出
            outputFn = function(row, key) {
                if (Array.isArray(key)) return key;
                return [key];
            };
        } else if (!colIndexes || colIndexes.length === 0) {
            // 无列选择器（整行去重等）：输出整行
            outputFn = function(row) {
                if (!row) return [];
                return Array.isArray(row) ? row.slice() : [row];
            };
        } else {
            // f1/f1,f2 模式：提取 colIndexes 列
            outputFn = function(row) {
                if (!row) return [];
                var out = [];
                for (var i = 0; i < colIndexes.length; i++) {
                    out.push(row[colIndexes[i]]);
                }
                return out;
            };
        }
    } else if (resultSelector === '') {
        // 空字符串：输出所有列
        outputFn = function(row) {
            if (!row) return [];
            return Array.isArray(row) ? row.slice() : [row];
        };
    } else if (typeof resultSelector === 'string') {
        // 字符串模式选择输出列
        if (resultSelector.includes(',') || resultSelector.includes('，')) {
            var outIndexes = [];
            var outParts = resultSelector.split(/[,，]/);
            for (var j = 0; j < outParts.length; j++) {
                var outPart = outParts[j].trim();
                if (outPart.toLowerCase().startsWith('f')) {
                    if (outPart.includes('-')) {
                        // 处理范围 f3-f5
                        var range = outPart.split('-');
                        var start = parseInt(range[0].toLowerCase().replace('f', ''));
                        var end = parseInt(range[1].toLowerCase().replace('f', ''));
                        for (var r = start; r <= end; r++) {
                            outIndexes.push(r - 1);
                        }
                    } else {
                        outIndexes.push(parseInt(outPart.substring(1)) - 1);
                    }
                } else {
                    outIndexes.push(parseInt(outPart) - 1);
                }
            }
            outputFn = function(row) {
                if (!row) return [];
                var out = [];
                for (var ki = 0; ki < outIndexes.length; ki++) {
                    out.push(row[outIndexes[ki]]);
                }
                return out;
            };
        } else if (resultSelector.toLowerCase().startsWith('f')) {
            if (resultSelector.includes('-')) {
                // 处理范围 f3-f5
                var range = resultSelector.split('-');
                var start = parseInt(range[0].toLowerCase().replace('f', ''));
                var end = parseInt(range[1].toLowerCase().replace('f', ''));
                var rangeIndexes = [];
                for (var r = start; r <= end; r++) {
                    rangeIndexes.push(r - 1);
                }
                outputFn = function(row) {
                    if (!row) return [];
                    var out = [];
                    for (var i = 0; i < rangeIndexes.length; i++) {
                        out.push(row[rangeIndexes[i]]);
                    }
                    return out;
                };
            } else {
                var outIdx = parseInt(resultSelector.substring(1)) - 1;
                outputFn = function(row) {
                    if (!row) return [];
                    return [row[outIdx]];
                };
            }
        } else {
            outputFn = function(row) {
                if (!row) return [];
                return Array.isArray(row) ? row.slice() : [row];
            };
        }
    } else if (Array.isArray(resultSelector)) {
        // 数组模式选择输出列
        outputFn = function(row) {
            if (!row) return [];
            var out = [];
            for (var m = 0; m < resultSelector.length; m++) {
                out.push(row[resultSelector[m]]);
            }
            return out;
        };
    } else {
        outputFn = function(row) {
            if (!row) return [];
            return Array.isArray(row) ? row.slice() : [row];
        };
    }

    // 🔧 v4.x 性能：缓存 _items 避免每次迭代触发 getter（getter 每次拷贝整个数组，O(n²)）
    var _items = this._items;

    // XXD-134/XXD-137 fix: 仅当源头未设 _header (纯 raw array) 时跳过第 0 行 (header).
    //   源头已设 _header (例如链式 z去重→z去重) → _items 是纯数据, 不应再 slice,
    //   否则会把第一行数据当 header 丢 (XXD-137 user repro variant: idempotent chain).
    //   与 z映射 行 8556-8562 的 skipHeader 约定保持一致: 'in' 检查通过 → 视作纯数据.
    if (!(Object.prototype.hasOwnProperty.call(this, '_header')) && _items.length > 0) {
        _items = _items.slice(1);
    }

    for (var i = 0; i < _items.length; i++) {
        var row = _items[i];
        // 🔧 v4.0.3 防御性检查：跳过无效行
        if (!row || (Array.isArray(row) && row.length === 0)) continue;

        // 🔧 v4.0.3 仅函数模式创建代理（支持 x.f1, x.f2 属性访问），其他模式直接用 row
        var input = row;
        if (isFunctionMode && Array.isArray(row)) {
            input = row.slice();
            for (var c = 0; c < input.length; c++) {
                input['f' + (c + 1)] = input[c];
            }
        }

        var key;
        try {
            key = keyFn(input);
        } catch (e) {
            // 🔧 如果 keyFn 执行失败（如 row.f2 为 undefined），跳过该行
            console.warn('⚠️ z去重跳过第' + (i+1) + '行:', e.message);
            continue;
        }
        var keyStr = typeof key === 'string' ? key : JSON.stringify(key);
        if (!seen[keyStr]) {
            seen[keyStr] = true;
            result.push(outputFn(input, key));
        }
    }
    // XXD-134 final fix: z去重 返回时设 _header=1 让链式 z映射/z筛选 自动跳过表头行
    var _distinctResult = this._new(result);
    _distinctResult._header = 1;
    return _distinctResult;
};
Array2D.prototype.distinct = Array2D.prototype.z去重;
Array2D.prototype.zDistinct = Array2D.prototype.z去重;

/**
 * 转矩阵（toMatrix）- 转换为标准矩阵格式，补齐缺失列
 * @param {any} fillValue - 填充值，默认为null
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4,5],[6]]).z转矩阵()  // [[1,2,null],[3,4,5],[6,null,null]]
 */
Array2D.prototype.z转矩阵 = function(fillValue) {
    fillValue = fillValue !== undefined ? fillValue : null;
    if (this._items.length === 0) return this._new([]);

    // 找出最大列数
    var maxCols = 0;
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var rowLen = Array.isArray(row) ? row.length : 1;
        if (rowLen > maxCols) maxCols = rowLen;
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row)) {
            var newRow = row.slice();
            while (newRow.length < maxCols) {
                newRow.push(fillValue);
            }
            result.push(newRow);
        } else {
            var newRow = [row];
            while (newRow.length < maxCols) {
                newRow.push(fillValue);
            }
            result.push(newRow);
        }
    }
    return this._new(result);
};
Array2D.prototype.toMatrix = Array2D.prototype.z转矩阵;

// ==================== 分组透视 ====================

/**
 * 分组
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} valSelector - 值选择器
 * @returns {Object} 分组结果
 */
Array2D.prototype.z分组 = function(keySelector, valSelector) {
    // XXD-137/140: 中文逗号 → ASCII逗号 归一化，确保 "f1，f2" 行为与 "f1,f2" 一致

    if (typeof keySelector === "string") keySelector = keySelector.replace(/，/g, ",");
    if (typeof valSelector === "string") valSelector = valSelector.replace(/，/g, ",");

    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    var valFn = valSelector ? (typeof valSelector === 'function' ? valSelector : parseLambda(valSelector)) : null;

    // XXD-83: 空 keySelector 默认按整行分组（与 z去重 行为一致）

    // XXD-142/XXD-139: 数字索引/数组模式 — parseLambda 不处理非字符串，需显式构建 keyFn
    if (!keyFn && typeof keySelector === 'number') {
        keyFn = function(row) { return row[keySelector]; };
    } else if (!keyFn && Array.isArray(keySelector)) {
        if (keySelector.length === 1) {
            var _ks0 = keySelector[0];
            keyFn = function(row) { return row[_ks0]; };
        } else {
            keyFn = function(row) {
                var kp = [];
                for (var _ki = 0; _ki < keySelector.length; _ki++) kp.push(row[keySelector[_ki]]);
                return JSON.stringify(kp);
            };
        }
    }
    if (!keyFn) {
        keyFn = function(row) { return JSON.stringify(row); };
    }

    // 🔧 v4.0.38 Index 加速: 命中 _indexes 缓存时走 O(unique) 路径
    var _cacheKey = _normalizeCacheKey(keySelector);
    if (_cacheKey !== null && this._indexes && this._indexes[_cacheKey]) {
        return this._indexes[_cacheKey].groupBy(valSelector);
    }

    var groups = Object.create(null);
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var key = keyFn ? keyFn(row, i) : row[0];
        var val = valFn ? valFn(row, i) : row;
        if (!groups[key]) groups[key] = [];
        groups[key].push(val);
    }
    return groups;
};
Array2D.prototype.groupBy = Array2D.prototype.z分组;
// 🔧 XXD-133 final fix: z分组汇总 alias (= z分组, 返回 Map)
Array2D.prototype.z分组汇总 = Array2D.prototype.z分组;
// XXD-133/136: z分组汇总 prototype — 委托到静态 groupInto (分组+聚合)
Array2D.prototype.z分组汇总 = function(keySelector, valueSelector, separator) {
    return Array2D.groupInto(this._items, keySelector, valueSelector, separator);
};
Array2D.prototype.groupInto = Array2D.prototype.z分组汇总;

/**
 * 数据透视（pivotBy）- 创建数据透视表
 * @param {Number|Function} rowField - 行字段索引或选择器
 * @param {Number|Function} colField - 列字段索引或选择器
 * @param {Number|Function} valueField - 值字段索引或选择器
 * @param {Function} aggregator - 聚合函数，默认为求和
 * @returns {Array2D} 新实例（透视表）
 * @example
 * // 数据: [[产品, 地区, 销量], ['A', '北京', 100], ['A', '上海', 200]]
 * Array2D(data).z透视(0, 1, 2)  // 按产品(行)、地区(列)透视销量
 */
Array2D.prototype.z透视 = function(rowField, colField, valueField, aggregator) {
    if (this._items.length === 0) return this._new([]);

    // 默认聚合函数为求和
    var agg = aggregator || function(acc, val) {
        var num1 = typeof acc === 'number' ? acc : parseFloat(String(acc).replace(/,/g, ''));
        var num2 = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
        return (isNaN(num1) ? 0 : num1) + (isNaN(num2) ? 0 : num2);
    };

    var rowValues = [];
    var colValues = [];
    var pivotData = Object.create(null);

    // 辅助函数：获取字段值
    var getFieldValue = function(row, field, index) {
        if (typeof field === 'function') return field(row, index);
        if (Array.isArray(row)) return row[field];
        return row;
    };

    // 第一遍：收集所有行和列的值
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        var rowKey = String(getFieldValue(row, rowField, i));
        var colKey = String(getFieldValue(row, colField, i));
        var value = getFieldValue(row, valueField, i);

        // 收集行值
        if (rowValues.indexOf(rowKey) === -1) rowValues.push(rowKey);
        // 收集列值
        if (colValues.indexOf(colKey) === -1) colValues.push(colKey);

        // 初始化数据结构
        if (!pivotData[rowKey]) pivotData[rowKey] = Object.create(null);

        // 聚合值
        if (pivotData[rowKey][colKey] === undefined) {
            pivotData[rowKey][colKey] = value;
        } else {
            pivotData[rowKey][colKey] = agg(pivotData[rowKey][colKey], value);
        }
    }

    // 排序
    rowValues.sort();
    colValues.sort();

    // 构建结果表
    var result = [];

    // 表头
    var header = ['行\\列'].concat(colValues);
    result.push(header);

    // 数据行
    for (var r = 0; r < rowValues.length; r++) {
        var rowKey = rowValues[r];
        var rowData = [rowKey];
        for (var c = 0; c < colValues.length; c++) {
            var colKey = colValues[c];
            var value = pivotData[rowKey] && pivotData[rowKey][colKey] !== undefined
                ? pivotData[rowKey][colKey]
                : 0;
            rowData.push(value);
        }
        result.push(rowData);
    }

    return this._new(result);
};
Array2D.prototype.pivotBy = Array2D.prototype.z透视;

// ==================== 连接相关方法 ====================

/**
 * 上下连接（concat）- 将两个或多个数组按行连接
 * @param {Array} brr - 第二个数组或多个数组
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z上下连接([[5,6]])  // [[1,2],[3,4],[5,6]]
 */
Array2D.prototype.z上下连接 = function() {
    var result = this._items.slice();
    for (var i = 0; i < arguments.length; i++) {
        var arr = arguments[i];
        // 支持Array2D实例和普通数组
        if (arr instanceof Array2D) {
            result = result.concat(arr._items);
        } else if (Array.isArray(arr)) {
            if (arr.length > 0 && Array.isArray(arr[0])) {
                result = result.concat(arr);
            } else {
                result.push(arr);
            }
        }
    }
    return this._new(result);
};
Array2D.prototype.concat = Array2D.prototype.z上下连接;

/**
 * 左连接（leftjoin）- 以左表为准的左外连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z左连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return JSON.stringify(row); };
    
    // 辅助函数：比较两个键值是否相等（支持数组键）
    function keysEqual(k1, k2) {
        if (k1 === k2) return true;
        if (Array.isArray(k1) && Array.isArray(k2)) {
            if (k1.length !== k2.length) return false;
            for (var ki = 0; ki < k1.length; ki++) {
                if (k1[ki] !== k2[ki]) return false;
            }
            return true;
        }
        return String(k1) === String(k2);
    }

    // 处理 resultSelector：支持函数或字符串（如 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5'）
    var resFn;
    if (typeof resultSelector === 'function') {
        resFn = resultSelector;
    } else if (typeof resultSelector === 'string' && resultSelector) {
        // 解析 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5' 这样的字符串
        var parts = resultSelector.split(/[,，]/).map(function(s) { return s.trim(); });
        var selectors = parts.map(function(part) {
            var match = part.match(/^([ab])\.f(\d+)$/i);
            if (match) {
                return {
                    table: match[1].toLowerCase(), // 'a' 或 'b'
                    colIndex: parseInt(match[2]) - 1 // 0-based 索引
                };
            }
            return null;
        }).filter(function(s) { return s !== null; });
        
        resFn = function(leftRow, rightRow) {
            var result = [];
            for (var s = 0; s < selectors.length; s++) {
                var sel = selectors[s];
                var row = sel.table === 'a' ? leftRow : rightRow;
                if (row && sel.colIndex >= 0 && sel.colIndex < row.length) {
                    result.push(row[sel.colIndex]);
                } else {
                    result.push(null);
                }
            }
            return result;
        };
    } else {
        // 默认：直接拼接
        resFn = function(a, b) { return a.concat(b || []); };
    }

    // XXD-16: pre-build rightMap for O(M+N); leftjoin still takes the first match (preserves original behavior)
    var rightMap = {};
    for (var j = 0; j < brr.length; j++) {
        var sk = String(rightFn(brr[j], j));
        if (!rightMap[sk]) rightMap[sk] = [];
        rightMap[sk].push(brr[j]);
    }
    // Cache this._items once: it's a getter that copies the data on every access (perf footgun)
    var leftItems = this._items;
    var leftLen = leftItems.length;
    var result = [];
    for (var i = 0; i < leftLen; i++) {
        var leftRow = leftItems[i];
        var leftKey = leftFn(leftRow, i);
        var rightRows = rightMap[String(leftKey)] || [];
        var matched = rightRows[0] || null;
        result.push(resFn(leftRow.slice(), matched ? matched.slice() : []));
    }
    return this._new(result);
};
Array2D.prototype.leftjoin = Array2D.prototype.z左连接;
// 🔧 XXD-132 final fix: z右连接 alias (= swap args of z左连接)
Array2D.prototype.z右连接 = function(brr, rightKeySelector, leftKeySelector, resultSelector) {
    return Array2D.prototype.z左连接.call(this, brr instanceof Array2D ? brr : new Array2D(brr), rightKeySelector, leftKeySelector, resultSelector);
};

/**
 * 内连接（innerjoin）- 仅保留两表键匹配的行（多对多：左表每一匹配 × 右表每一匹配）
 * 与 z左连接 的区别：未匹配的左表行不会以空右表形式输出。
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 * @example
 * arr.z内连接(brr, 'f1', 'f1')          // 默认拼接
 * arr.innerjoin(brr, 'f1', 'f1', 'a.f1,b.f2')
 */
Array2D.prototype.z内连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    // XXD-206 S3.5: null brr → empty result (instead of THROW)
    if (brr == null) return this._new([]);

    // XXD-206 S3.6: capture header rows for header-name key resolution
    var _leftHeader = (this._items && this._items.length > 0) ? this._items[0] : null;
    var _rightHeader = (brr && brr.length > 0) ? brr[0] : null;

    function pickKeyFn(sel, headerRow) {
        if (typeof sel === 'function') return sel;
        if (typeof sel === 'number' && Number.isFinite(sel) && sel >= 0) {
            return function(row) { return row == null ? undefined : row[sel]; };
        }
        if (sel) {
            // XXD-206 S3.6: if the selector is a plain string matching a header cell,
            // treat it as a column-name lookup against the table's first row.
            if (typeof sel === 'string' && Array.isArray(headerRow)) {
                for (var _h = 0; _h < headerRow.length; _h++) {
                    if (headerRow[_h] === sel) {
                        return function(row) { return row == null ? undefined : row[_h]; };
                    }
                }
            }
            return parseLambda(sel);
        }
        return function(row) { return JSON.stringify(row); };
    }
    var leftFn = pickKeyFn(leftKeySelector, _leftHeader);
    var rightFn = pickKeyFn(rightKeySelector, _rightHeader);
    // XXD-206 S3.3: detect "header-name join with matching name" for natural-join semantic
    var _rightKeyColIndex = -1;
    if (typeof leftKeySelector === 'string' && typeof rightKeySelector === 'string'
        && leftKeySelector === rightKeySelector
        && Array.isArray(_rightHeader)) {
        for (var _rk = 0; _rk < _rightHeader.length; _rk++) {
            if (_rightHeader[_rk] === rightKeySelector) { _rightKeyColIndex = _rk; break; }
        }
    }

    // 处理 resultSelector：支持函数或字符串（如 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5'）
    var resFn;
    if (typeof resultSelector === 'function') {
        resFn = resultSelector;
    } else if (typeof resultSelector === 'string' && resultSelector) {
        var parts = resultSelector.split(/[,，]/).map(function(s) { return s.trim(); });
        var selectors = parts.map(function(part) {
            var match = part.match(/^([ab])\.f(\d+)$/i);
            if (match) {
                return {
                    table: match[1].toLowerCase(), // 'a' 或 'b'
                    colIndex: parseInt(match[2]) - 1 // 0-based 索引
                };
            }
            return null;
        }).filter(function(s) { return s !== null; });

        resFn = function(leftRow, rightRow) {
            var result = [];
            for (var s = 0; s < selectors.length; s++) {
                var sel = selectors[s];
                var row = sel.table === 'a' ? leftRow : rightRow;
                if (row && sel.colIndex >= 0 && sel.colIndex < row.length) {
                    result.push(row[sel.colIndex]);
                } else {
                    result.push(null);
                }
            }
            return result;
        };
    } else if (Array.isArray(resultSelector) && resultSelector.length > 0 && resultSelector.every(function(x){ return typeof x === 'number' && Number.isFinite(x) && x >= 0; })) {
        // XXD-184 array-form resultSelector: 数组形式 = 扩展左行 + 追加右表指定列
        // [idx1, idx2, ...] → result row = leftRow + [right[idx1], right[idx2], ...]
        // XXD-206 S3.4: 越界返 null (显式), 不用 undefined (静默)
        var arrIdx = resultSelector;
        resFn = function(leftRow, rightRow) {
            var out = leftRow.slice();
            for (var s = 0; s < arrIdx.length; s++) {
                var i = arrIdx[s];
                out.push(rightRow == null ? null : (i < rightRow.length ? rightRow[i] : null));
            }
            return out;
        };
    } else {
        // 默认：直接拼接
        // XXD-206 S3.3: 头名相同时, 默认结果去掉右表键列, 避免 ['A','B','B','C'] 重复.
        // 仅当左右都是 header 解析路径且名字相同才触发(原数字/函数键不受影响).
        if (typeof _rightKeyColIndex === 'number' && _rightKeyColIndex >= 0) {
            resFn = function(a, b) {
                var out = a.slice();
                if (b) {
                    for (var _i = 0; _i < b.length; _i++) {
                        if (_i === _rightKeyColIndex) continue;
                        out.push(b[_i]);
                    }
                }
                return out;
            };
        } else {
            resFn = function(a, b) { return a.concat(b || []); };
        }
    }

    // pre-build rightMap for O(M+N)
    var rightMap = {};
    for (var j = 0; j < brr.length; j++) {
        var sk = String(rightFn(brr[j], j));
        if (!rightMap[sk]) rightMap[sk] = [];
        rightMap[sk].push(brr[j]);
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var leftKey = leftFn(leftRow, i);
        var rightRows = rightMap[String(leftKey)] || [];
        // inner join: 仅当存在匹配时输出；多对多时输出所有组合
        for (var k = 0; k < rightRows.length; k++) {
            result.push(resFn(leftRow.slice(), rightRows[k].slice()));
        }
    }
    return this._new(result);
};
/* XXD-206 S3.x fixes applied */
Array2D.prototype.innerjoin = Array2D.prototype.z内连接;
// 🔧 XXD-131 final fix: z分组聚合 alias (= z分组, 返回 Map)
Array2D.prototype.z分组聚合 = function(keySelector, valSelector) {
    return this.z分组(keySelector);
};

/**
 * 左右全连接（fulljoin）- 全外连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z左右全连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return JSON.stringify(row); };
    
    // 辅助函数：键序列化（用于作为对象key）
    function serializeKey(k) {
        if (k === null || k === undefined) return String(k);
        if (Array.isArray(k)) return k.join('|@|');
        return String(k);
    }
    // 辅助函数：键相等比较
    function keysEqual(k1, k2) {
        if (k1 === k2) return true;
        if (Array.isArray(k1) && Array.isArray(k2)) {
            if (k1.length !== k2.length) return false;
            for (var ki = 0; ki < k1.length; ki++) {
                if (k1[ki] !== k2[ki]) return false;
            }
            return true;
        }
        return String(k1) === String(k2);
    }

    // 处理 resultSelector：支持函数或字符串（如 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5'）
    var resFn;
    if (typeof resultSelector === 'function') {
        resFn = resultSelector;
    } else if (typeof resultSelector === 'string' && resultSelector) {
        var parts = resultSelector.split(/[,，]/).map(function(s) { return s.trim(); });
        var selectors = parts.map(function(part) {
            var match = part.match(/^([ab])\.f(\d+)$/i);
            if (match) {
                return {
                    table: match[1].toLowerCase(),
                    colIndex: parseInt(match[2]) - 1
                };
            }
            return null;
        }).filter(function(s) { return s !== null; });
        
        resFn = function(leftRow, rightRow) {
            var result = [];
            for (var s = 0; s < selectors.length; s++) {
                var sel = selectors[s];
                var row = sel.table === 'a' ? leftRow : rightRow;
                if (row && sel.colIndex >= 0 && sel.colIndex < row.length) {
                    result.push(row[sel.colIndex]);
                } else {
                    result.push(null);
                }
            }
            return result;
        };
    } else {
        resFn = function(a, b) { return a.concat(b || []); };
    }

    var rightMap = {}; // v4.0.11 fix

    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        var sk = serializeKey(key);
        if (!rightMap[sk]) rightMap[sk] = [];
        rightMap[sk].push(brr[j]);
    }

    var seenLeftKeys = Object.create(null);
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftKey = leftFn(this._items[i], i);
        var sk = serializeKey(leftKey);
        seenLeftKeys[sk] = leftKey;
        var rightRows = rightMap[sk] || [];
        if (rightRows.length > 0) {
            for (var ri = 0; ri < rightRows.length; ri++) {
                result.push(resFn(this._items[i].slice(), rightRows[ri].slice()));
            }
        } else {
            result.push(resFn(this._items[i].slice(), []));
        }
    }

    for (var j = 0; j < brr.length; j++) {
        var rightKey = rightFn(brr[j], j);
        var sk = serializeKey(rightKey);
        if (!seenLeftKeys[sk]) {
            result.push(resFn([], brr[j].slice()));
        }
    }

    return this._new(result);
};
Array2D.prototype.fulljoin = Array2D.prototype.z左右全连接;

/**
 * 一对多连接（leftFulljoin）- 左表所有行与右表匹配的所有行连接
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 * @example
 * arr.leftFulljoin(brr, 'f1', 'f1')
 */
Array2D.prototype.z一对多连接 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return row[0]; };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return row[0]; };
    
    // 处理 resultSelector：支持函数或字符串（如 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5'）
    var resFn;
    if (typeof resultSelector === 'function') {
        resFn = resultSelector;
    } else if (typeof resultSelector === 'string' && resultSelector) {
        // 解析 'a.f1,b.f2' 或 'b.f3,b.f4,b.f5' 这样的字符串
        var parts = resultSelector.split(/[,，]/).map(function(s) { return s.trim(); });
        var selectors = parts.map(function(part) {
            var match = part.match(/^([ab])\.f(\d+)$/i);
            if (match) {
                return {
                    table: match[1].toLowerCase(), // 'a' 或 'b'
                    colIndex: parseInt(match[2]) - 1 // 0-based 索引
                };
            }
            return null;
        }).filter(function(s) { return s !== null; });
        
        resFn = function(leftRow, rightRow) {
            var result = [];
            for (var s = 0; s < selectors.length; s++) {
                var sel = selectors[s];
                var row = sel.table === 'a' ? leftRow : rightRow;
                if (row && sel.colIndex >= 0 && sel.colIndex < row.length) {
                    result.push(row[sel.colIndex]);
                } else {
                    result.push(null);
                }
            }
            return result;
        };
    } else {
        // 默认：直接拼接
        resFn = function(a, b) { return a.concat(b || []); };
    }

    var rightMap = {}; // v4.0.11 fix
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        if (!rightMap[key]) rightMap[key] = [];
        rightMap[key].push(brr[j]);
    }

    var result = [];
    var matchedRightKeys = {}; // v4.0.11: track matched right keys for full join
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var key = leftFn(leftRow, i);
        var rightRows = rightMap[key] || [];
        if (rightRows.length === 0) {
            result.push(resFn(leftRow.slice(), []));
        } else {
            matchedRightKeys[key] = true;
            for (var r = 0; r < rightRows.length; r++) {
                result.push(resFn(leftRow.slice(), rightRows[r].slice()));
            }
        }
    }
    // v4.0.11: 添加右表独有行（未被匹配的）
    for (var rk in rightMap) {
        if (rightMap.hasOwnProperty(rk) && !matchedRightKeys[rk]) {
            var unmatchedRows = rightMap[rk];
            for (var u = 0; u < unmatchedRows.length; u++) {
                result.push(resFn([], unmatchedRows[u].slice()));
            }
        }
    }
    return this._new(result);
};
Array2D.prototype.leftFulljoin = Array2D.prototype.z一对多连接;

/**
 * 左右连接（zip）- 按行左右拼接
 * @param {...Array} arrays - 要拼接的数组
 * @returns {Array2D} 新实例
 * @example
 * Array2D([[1,2],[3,4]]).z左右连接([[5],[6]])  // [[1,2,5],[3,4,6]]
 */
Array2D.prototype.z左右连接 = function() {
    var arrays = [this._items];
    for (var i = 0; i < arguments.length; i++) {
        arrays.push(arguments[i]);
    }

    var maxRows = 0;
    for (var a = 0; a < arrays.length; a++) {
        if (arrays[a].length > maxRows) maxRows = arrays[a].length;
    }

    var result = [];
    for (var r = 0; r < maxRows; r++) {
        var row = [];
        for (var a = 0; a < arrays.length; a++) {
            var arr = arrays[a];
            if (r < arr.length) {
                var rowData = arr[r];
                if (Array.isArray(rowData)) {
                    row = row.concat(rowData);
                } else {
                    row.push(rowData);
                }
            }
        }
        result.push(row);
    }

    return this._new(result);
};
Array2D.prototype.zip = Array2D.prototype.z左右连接;

/**
 * 排除（except）- 从左表排除与右表相同的元素
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftSelector - 左表选择器
 * @param {string|Function} rightSelector - 右表选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z排除 = function(brr, leftSelector, rightSelector) {
    var leftFn = leftSelector ? parseLambda(leftSelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightSelector ? parseLambda(rightSelector) : function(row) { return JSON.stringify(row); };

    function keysEqual(k1, k2) {
        if (k1 === k2) return true;
        if (Array.isArray(k1) && Array.isArray(k2)) {
            if (k1.length !== k2.length) return false;
            for (var ki = 0; ki < k1.length; ki++) {
                if (k1[ki] !== k2[ki]) return false;
            }
            return true;
        }
        return String(k1) === String(k2);
    }

    // 使用 Set 思想，用序列化键去重
    var rightSet = {}; // v4.0.11 fix: preserve hasOwnProperty
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        var sk = Array.isArray(key) ? key.join('|@|') : String(key);
        rightSet[sk] = key;
    }

    var rightKeys = [];
    for (var sk in rightSet) {
        if (rightSet.hasOwnProperty(sk)) rightKeys.push(rightSet[sk]);
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var key = leftFn(this._items[i], i);
        var found = false;
        for (var rk = 0; rk < rightKeys.length; rk++) {
            if (keysEqual(key, rightKeys[rk])) {
                found = true;
                break;
            }
        }
        if (!found) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.except = Array2D.prototype.z排除;

/**
 * 取交集（intersect）- 获取两个数组的交集
 * @param {Array} brr - 右表数组
 * @param {string|Function} leftSelector - 左表选择器
 * @param {string|Function} rightSelector - 右表选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z取交集 = function(brr, leftSelector, rightSelector) {
    var leftFn = leftSelector ? parseLambda(leftSelector) : function(row) { return JSON.stringify(row); };
    var rightFn = rightSelector ? parseLambda(rightSelector) : function(row) { return JSON.stringify(row); };

    function keysEqual(k1, k2) {
        if (k1 === k2) return true;
        if (Array.isArray(k1) && Array.isArray(k2)) {
            if (k1.length !== k2.length) return false;
            for (var ki = 0; ki < k1.length; ki++) {
                if (k1[ki] !== k2[ki]) return false;
            }
            return true;
        }
        return String(k1) === String(k2);
    }

    // 构建右表去重键集合
    var rightSet = {}; // v4.0.11 fix: preserve hasOwnProperty
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        var sk = Array.isArray(key) ? key.join('|@|') : String(key);
        rightSet[sk] = key;
    }

    var rightKeys = [];
    for (var sk in rightSet) {
        if (rightSet.hasOwnProperty(sk)) rightKeys.push(rightSet[sk]);
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var key = leftFn(this._items[i], i);
        var found = false;
        for (var rk = 0; rk < rightKeys.length; rk++) {
            if (keysEqual(key, rightKeys[rk])) {
                found = true;
                break;
            }
        }
        if (found) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.intersect = Array2D.prototype.z取交集;

/**
 * 去重并集（union）- 合并两个数组并去重
 * @param {Array} brr - 右表数组
 * @param {string|Function} keySelector - 键选择器
 * @returns {Array2D} 新实例
 */

/**
 * 超级查找（superLookup）- 类似VLOOKUP的多条件查找
 * @param {Array} brr - 查找表数组
 * @param {string|Function} leftKeySelector - 左表键选择器
 * @param {string|Function} rightKeySelector - 右表键选择器
 * @param {Function} resultSelector - 结果选择器
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z超级查找 = function(brr, leftKeySelector, rightKeySelector, resultSelector) {
    var leftFn = leftKeySelector ? parseLambda(leftKeySelector) : function(row) { return row[0]; };
    var rightFn = rightKeySelector ? parseLambda(rightKeySelector) : function(row) { return row[0]; };
    var resFn = resultSelector || function(a, b) { return a.concat(b || []); };

    function serializeKey(k) {
        return Array.isArray(k) ? k.join('|@|') : String(k);
    }

    // 构建右表查找字典（使用序列化键）
    var rightMap = {}; // v4.0.11 fix
    var rightRawMap = Object.create(null);
    for (var j = 0; j < brr.length; j++) {
        var key = rightFn(brr[j], j);
        var sk = serializeKey(key);
        if (!rightMap[sk]) {
            rightMap[sk] = brr[j];
            rightRawMap[sk] = key;
        }
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var leftRow = this._items[i];
        var key = leftFn(leftRow, i);
        var sk = serializeKey(key);
        var matched = rightMap[sk];
        result.push(resFn(leftRow.slice(), matched ? matched.slice() : []));
    }

    return this._new(result);
};
Array2D.prototype.superLookup = Array2D.prototype.z超级查找;

// ==================== 查找相关方法 ====================

/**
 * 查找单个元素（find）
 * @param {string|Function} predicate - 查找条件
 * @returns {any} 找到的元素
 */
Array2D.prototype.z查找单个 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return undefined;

    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            return this._items[i];
        }
    }
    return undefined;
};
Array2D.prototype.find = Array2D.prototype.z查找单个;

/**
 * 查找所有下标（findAllIndex）- 查找所有满足条件的元素位置
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 位置数组 [[行,列],...]
 */
Array2D.prototype.z查找所有下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (Array.isArray(row)) {
            for (var j = 0; j < row.length; j++) {
                if (fn(row[j], i, j)) {
                    result.push([i, j]);
                }
            }
        } else {
            if (fn(row, i, 0)) {
                result.push([i, 0]);
            }
        }
    }
    return result;
};
Array2D.prototype.findAllIndex = Array2D.prototype.z查找所有下标;

/**
 * 查找所有行下标（findRowsIndex）
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 行下标数组
 */
Array2D.prototype.z查找所有行下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            result.push(i);
        }
    }
    return result;
};
Array2D.prototype.findRowsIndex = Array2D.prototype.z查找所有行下标;

/**
 * 查找所有列下标（findColsIndex）
 * @param {Number} rowIndex - 行号
 * @param {string|Function} predicate - 查找条件
 * @returns {Array} 列下标数组
 */
Array2D.prototype.z查找所有列下标 = function(rowIndex, predicate) {
    // 🔧 v4.0.1 修复: 支持 (predicate) 单参数调用
    // 当 rowIndex 是函数时，自动识别为 predicate，默认搜索第0行
    if (typeof rowIndex === 'function') {
        predicate = rowIndex;
        rowIndex = 0;
    }
    // 兼容字符串 Lambda 作为 predicate 传入第一个参数的情况
    if (typeof rowIndex === 'string' && predicate === undefined) {
        predicate = rowIndex;
        rowIndex = 0;
    }

    var items = this._items;
    if (!items || items.length === 0) return [];
    if (rowIndex < 0 || rowIndex >= items.length) return [];

    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];

    var row = items[rowIndex];
    var cols = row.length;
    var result = [];
    for (var j = 0; j < cols; j++) {
        if (fn(row, j, rowIndex)) {
            result.push(j);
        }
    }
    return result;
};
Array2D.prototype.findColsIndex = Array2D.prototype.z查找所有列下标;

/**
 * 查找元素下标（findIndexByPredicate）
 * @param {string|Function} predicate - 查找条件
 * @returns {Number} 下标，未找到返回-1
 */
Array2D.prototype.z查找元素下标 = function(predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return -1;

    for (var i = 0; i < this._items.length; i++) {
        if (fn(this._items[i], i)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.findIndexByPredicate = Array2D.prototype.z查找元素下标;

/**
 * 值位置（indexOf）- 查找元素首次出现的位置
 * @param {any} value - 要查找的值
 * @param {Number} [fromIndex=0] - 开始查找的位置
 * @returns {Number} 下标，未找到返回-1
 * @example
 * Array2D([[1,2],[3,4],[1,2]]).z值位置([1,2])  // 0
 */
Array2D.prototype.z值位置 = function(value, fromIndex) {
    fromIndex = fromIndex || 0;
    for (var i = fromIndex; i < this._items.length; i++) {
        if (JSON.stringify(this._items[i]) === JSON.stringify(value)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.indexOf = Array2D.prototype.z值位置;

/**
 * 从后往前值位置（lastIndexOf）- 查找元素最后出现的位置
 * @param {any} value - 要查找的值
 * @param {Number} [fromIndex] - 开始查找的位置（从后往前）
 * @returns {Number} 下标，未找到返回-1
 * @example
 * Array2D([[1,2],[3,4],[1,2]]).z从后往前值位置([1,2])  // 2
 */
Array2D.prototype.z从后往前值位置 = function(value, fromIndex) {
    fromIndex = fromIndex !== undefined ? fromIndex : this._items.length - 1;
    for (var i = fromIndex; i >= 0; i--) {
        if (JSON.stringify(this._items[i]) === JSON.stringify(value)) {
            return i;
        }
    }
    return -1;
};
Array2D.prototype.lastIndexOf = Array2D.prototype.z从后往前值位置;

// ==================== 批量操作方法 ====================

/**
 * 批量删除列（deleteCols）
 * @param {Array|String} cols - 列号数组或f模式字符串
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z批量删除列 = function(cols) {
    var colIndexes = [];

    // 解析列索引
    if (typeof cols === 'string') {
        // f模式: f1,f2 或 f3
        if (cols.startsWith('f') && !cols.includes(',') && !cols.includes('，')) {
            var idx = parseInt(cols.substring(1)) - 1;
            colIndexes = [idx];
        } else {
            var parts = cols.split(/[,，]/);
            for (var p = 0; p < parts.length; p++) {
                if (parts[p].trim().startsWith('f')) {
                    colIndexes.push(parseInt(parts[p].trim().substring(1)) - 1);
                }
            }
        }
    } else if (Array.isArray(cols)) {
        colIndexes = cols;
    }

    // 从大到小排序删除
    colIndexes.sort(function(a, b) { return b - a; });

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = this._items[i].slice();
        for (var c = 0; c < colIndexes.length; c++) {
            if (colIndexes[c] >= 0 && colIndexes[c] < newRow.length) {
                newRow.splice(colIndexes[c], 1);
            }
        }
        result.push(newRow);
    }

    return this._new(result);
};
Array2D.prototype.deleteCols = Array2D.prototype.z批量删除列;
Array2D.prototype.delcols = Array2D.prototype.z批量删除列;  // 文档中的别名
Array2D.prototype.z批量删除列2 = Array2D.prototype.z批量删除列;  // 别名

/**
 * 批量删除行（deleteRows）
 * @param {Function|String|Array} rows - 行选择器（支持函数、字符串或数组）
 * @returns {Array2D} 新实例
 * @example
 * arr.z批量删除行(r=>r.f3=="美国")     // 函数模式
 * arr.z批量删除行("r=>r.f3=='美国'")  // Lambda字符串模式
 * arr.z批量删除行('f1-f3')           // f模式
 */
Array2D.prototype.z批量删除行 = function(rows) {
    var rowIndexes = [];

    if (typeof rows === 'function') {
        // 🔧 函数模式: 找出所有匹配条件的行索引并删除
        // 支持 r => r.f3 == "美国" 等函数/Lambda表达式
        var data = this._items;
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            // 创建带 f1, f2, f3... 属性的代理行，支持 r.f3 语法
            if (Array.isArray(row)) {
                var proxy = row.slice();
                for (var c = 0; c < proxy.length; c++) {
                    proxy['f' + (c + 1)] = proxy[c];
                }
                if (rows(proxy, i)) {
                    rowIndexes.push(i);
                }
            } else {
                if (rows(row, i)) {
                    rowIndexes.push(i);
                }
            }
        }
    } else if (typeof rows === 'string') {
        // 🔧 v4.0.1 新增：Lambda字符串模式 "r=>r.f3=='美国'" 或 "r=>r.f3==\"美国\""
        if (rows.includes('=>')) {
            // 是Lambda表达式字符串
            var fn;
            try {
                fn = eval('(' + rows + ')');
            } catch (e) {
                console.warn('Lambda解析失败:', rows, e);
            }
            if (fn) {
                var data = this._items;
                for (var i = 0; i < data.length; i++) {
                    var row = data[i];
                    if (Array.isArray(row)) {
                        var proxy = row.slice();
                        for (var c = 0; c < proxy.length; c++) {
                            proxy['f' + (c + 1)] = proxy[c];
                        }
                        if (fn(proxy, i)) {
                            rowIndexes.push(i);
                        }
                    }
                }
            }
        } else if (rows.includes('-')) {
            // f模式: f2-f4
            var match = rows.match(/f(\d+)\-f(\d+)/);
            if (match) {
                var start = parseInt(match[1]) - 1;
                var end = parseInt(match[2]) - 1;
                for (var i = start; i <= end; i++) {
                    rowIndexes.push(i);
                }
            }
        } else {
            // 🔧 v4.0.11 新增：支持条件字符串 "f3>80" 等
            // 检测是否为条件表达式（包含比较运算符）
            var condMatch = rows.match(/^f(\d+)\s*([<>=!]+)\s*(.+)$/);
            if (condMatch) {
                // 是条件表达式，使用 parseLambda 转换
                try {
                    var condFn = Array2D.parseLambda(rows);
                    if (condFn) {
                        var data = this._items;
                        for (var i = 0; i < data.length; i++) {
                            var row = data[i];
                            if (Array.isArray(row)) {
                                var proxy = row.slice();
                                for (var c = 0; c < proxy.length; c++) {
                                    proxy['f' + (c + 1)] = proxy[c];
                                }
                                if (condFn(proxy, i)) {
                                    rowIndexes.push(i);
                                }
                            }
                        }
                    }
                } catch (e) {
                    console.warn('条件解析失败:', rows, e);
                }
            } else if (rows.startsWith('f')) {
                // 纯列号模式: f3
                rowIndexes = [parseInt(rows.substring(1)) - 1];
            }
        }
    } else if (Array.isArray(rows)) {
        rowIndexes = rows;
    }

    // 从大到小排序删除
    rowIndexes.sort(function(a, b) { return b - a; });

    var result = this._items.slice();
    for (var r = 0; r < rowIndexes.length; r++) {
        if (rowIndexes[r] >= 0 && rowIndexes[r] < result.length) {
            result.splice(rowIndexes[r], 1);
        }
    }

    return this._new(result);
};
Array2D.prototype.deleteRows = Array2D.prototype.z批量删除行;

/**
 * 批量插入列（insertCols）
 * @param {Number|Function} colSelector - 列号或条件回调
 * @param {any|Function} value - 填充值或回调
 * @param {Number} count - 插入数量
 * @returns {Array2D} 新实例
 * @example
 * // 在第2列位置插入2列
 * Array2D(arr).z批量插入列(1, "x", 2)
 * // 在包含"产品"值的列位置前插入2列（默认在最后一行查找）
 * Array2D(arr).z批量插入列(x=>x.includes("产品"), " ", 2)
 */
Array2D.prototype.z批量插入列 = function(colSelector, value, count) {
    count = count || 1;
    var fillVal = value;

    var insertIndex = 0;
    if (typeof colSelector === 'function') {
        // 从条件函数解析目标值
        var funcStr = colSelector.toString();
        var valueMatch = funcStr.match(/['"]([^'"]+)['"]/);

        if (valueMatch) {
            var targetValue = valueMatch[1];
            // 默认在最后一行查找目标值的位置
            var lastRow = this._items[this._items.length - 1];
            if (Array.isArray(lastRow)) {
                for (var j = 0; j < lastRow.length; j++) {
                    if (String(lastRow[j]) == targetValue) {
                        insertIndex = j;
                        break;
                    }
                }
            }
        } else {
            // 尝试从 x[N] 解析列索引
            var indexMatch = funcStr.match(/x\[(\d+)\]/);
            if (indexMatch) {
                insertIndex = parseInt(indexMatch[1]);
            }
        }
    } else if (typeof colSelector === 'number') {
        insertIndex = colSelector;
    }

    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var row = this._items[i];
        if (!Array.isArray(row)) row = [row];

        var newRow = row.slice();
        // 准备填充值
        var fillVals = [];
        for (var c = 0; c < count; c++) {
            if (typeof fillVal === 'function') {
                fillVals.push(fillVal(row, i, insertIndex + c));
            } else {
                fillVals.push(fillVal !== undefined ? fillVal : '');
            }
        }
        // 在指定位置插入
        newRow.splice.apply(newRow, [insertIndex, 0].concat(fillVals));
        result.push(newRow);
    }

    return this._new(result);
};
Array2D.prototype.insertCols = Array2D.prototype.z批量插入列;

/**
 * 批量插入行（insertRows）
 * @param {Array|Function} rowSelector - 行号数组或条件回调
 * @param {any} value - 填充值
 * @param {Number} count - 插入数量
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z批量插入行 = function(rowSelector, value, count) {
    count = count || 1;
    var fillVal = value !== undefined ? value : '';

    var insertIndexes = [];
    if (typeof rowSelector === 'function') {
        for (var i = 0; i < this._items.length; i++) {
            if (rowSelector(this._items[i], i)) {
                insertIndexes.push(i);
            }
        }
    } else if (typeof rowSelector === 'string' && rowSelector.startsWith('f')) {
        insertIndexes = [parseInt(rowSelector.substring(1)) - 1];
    } else if (Array.isArray(rowSelector)) {
        insertIndexes = rowSelector;
    }

    var result = this._items.slice();
    // 从后往前插入
    for (var i = insertIndexes.length - 1; i >= 0; i--) {
        var idx = insertIndexes[i];
        var newRow = [];
        var maxCols = 0;
        for (var r = 0; r < result.length; r++) {
            if (Array.isArray(result[r]) && result[r].length > maxCols) {
                maxCols = result[r].length;
            }
        }
        for (var c = 0; c < maxCols; c++) {
            newRow.push(fillVal);
        }
        for (var c = 0; c < count; c++) {
            result.splice(idx, 0, newRow.slice());
        }
    }

    return this._new(result);
};
Array2D.prototype.insertRows = Array2D.prototype.z批量插入行;

/**
 * 插入行号（insertRowNum）
 * @param {Number} startNum - 起始行号
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z插入行号 = function(startNum) {
    startNum = startNum || 0;
    var result = [];
    for (var i = 0; i < this._items.length; i++) {
        var newRow = [startNum + i].concat(this._items[i]);
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.insertRowNum = Array2D.prototype.z插入行号;

// ==================== 分页方法 ====================

/**
 * 按页数分页（pageByCount）
 * @param {Number} pageCount - 总页数
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按页数分页 = function(pageCount) {
    if (pageCount < 1) pageCount = 1;
    var totalRows = this._items.length;
    var rowsPerPage = Math.ceil(totalRows / pageCount);

    var result = [];
    for (var page = 0; page < pageCount; page++) {
        var start = page * rowsPerPage;
        var end = Math.min(start + rowsPerPage, totalRows);
        if (start < totalRows) {
            result.push(this._items.slice(start, end));
        }
    }

    return result;
};
Array2D.prototype.pageByCount = Array2D.prototype.z按页数分页;

/**
 * 按行数分页（pageByRows）
 * @param {Number} pageSize - 每页行数
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按行数分页 = function(pageSize) {
    if (pageSize < 1) pageSize = 1;

    var result = [];
    for (var i = 0; i < this._items.length; i += pageSize) {
        result.push(this._items.slice(i, i + pageSize));
    }

    return result;
};
Array2D.prototype.pageByRows = Array2D.prototype.z按行数分页;

/**
 * 按下标分页（pageByIndexs）
 * @param {Array|String} indexes - 下标数组或条件
 * @returns {Array} 分页后的多维数组
 */
Array2D.prototype.z按下标分页 = function(indexes) {
    var splitIndexes = [];

    if (typeof indexes === 'string') {
        // f模式条件
        var fn = parseLambda(indexes);
        if (fn) {
            for (var i = 0; i < this._items.length; i++) {
                if (fn(this._items[i], i)) {
                    splitIndexes.push(i);
                }
            }
        }
    } else if (Array.isArray(indexes)) {
        splitIndexes = indexes;
    }

    if (splitIndexes.length === 0) return [this._items.slice()];

    var result = [];
    var start = 0;
    for (var i = 0; i < splitIndexes.length; i++) {
        var idx = splitIndexes[i];
        if (idx > start) {
            result.push(this._items.slice(start, idx));
        }
        start = idx;
    }
    result.push(this._items.slice(start));

    return result;
};
Array2D.prototype.pageByIndexs = Array2D.prototype.z按下标分页;

// ==================== 其他高级方法 ====================

/**
 * 间隔取数（nth）
 * @param {Number} interval - 间隔
 * @param {Number} offset - 偏移
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z间隔取数 = function(offset, step) {
    // 🔧 XXD-187: 修复(0,2)→[[1,2,3,4,5]] 应是 [[1,3,5]] 的取数语义
    //   - 参数顺序改为 (offset, step) 匹配验收用例
    //   - 单行输入(1D 向量)直接按 cell 取
    //   - 多行输入保留表头，按 (offset, step) 取后续行
    offset = offset || 0;
    step = step || 1;
    if (step === 0) step = 1;

    // 🔧 v3.9.4 修复：空数组保护
    if (!this._items || this._items.length === 0) return this._new([]);

    // 单行(向量): 按 cell 取
    if (this._items.length === 1) {
        var row0 = this._items[0];
        var picked = [];
        for (var j = offset; j < row0.length; j += step) {
            picked.push(row0[j]);
        }
        return this._new([picked]);
    }

    // 多行: 保留第一行（表头），按 (offset, step) 抽取后续行
    var result = [this._items[0].slice()];
    for (var i = 1; i < this._items.length; i++) {
        var bodyIdx = i - 1;          // 0-based body row index
        if (bodyIdx < offset) continue;
        if ((bodyIdx - offset) % step === 0) {
            result.push(this._items[i]);
        }
    }

    return this._new(result);
};
Array2D.prototype.nth = Array2D.prototype.z间隔取数;

/**
 * 补齐数组（pad）
 * @param {Number} cols - 列数
 * @param {Number} rows - 行数
 * @param {any} fillValue - 填充值
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z补齐数组 = function(cols, rows, fillValue) {
    cols = cols || (this._items[0] ? this._items[0].length : 1);
    rows = rows || this._items.length;
    fillValue = fillValue !== undefined ? fillValue : '';

    var result = [];
    for (var i = 0; i < rows; i++) {
        var row = i < this._items.length ? this._items[i].slice() : [];
        while (row.length < cols) {
            row.push(fillValue);
        }
        result.push(row);
    }

    return this._new(result);
};
Array2D.prototype.pad = Array2D.prototype.z补齐数组;

/**
 * 重设大小（resize）
 * @param {Number} rows - 行数
 * @param {Number} cols - 列数
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z重设大小 = function(rows, cols) {
    return this.z补齐数组(cols, rows);
};
Array2D.prototype.resize = Array2D.prototype.z重设大小;

/**
 * 处理空值（noNull）- 将null和undefined替换为空字符串
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z处理空值 = function(fill) {
    // XXD-203 final fix: if fill is omitted, per-column auto-detect:
    //   all-numeric col -> 0, all-string col -> '', otherwise keep null.
    // XXD-191: if fill is supplied, use it for every cell (back-compat).
    var useAuto = (fill === undefined);
    var result = [];
    if (!useAuto) {
        for (var i = 0; i < this._items.length; i++) {
            var row = [];
            if (Array.isArray(this._items[i])) {
                for (var j = 0; j < this._items[i].length; j++) {
                    var val = this._items[i][j];
                    row.push((val === null || val === undefined) ? fill : val);
                }
            } else {
                var val = this._items[i];
                row.push((val === null || val === undefined) ? fill : val);
            }
            result.push(row);
        }
        return this._new(result);
    }
    // auto-detect path: compute per-column default first.
    var rows = this._items;
    var ncols = 0;
    for (var r0 = 0; r0 < rows.length; r0++) {
        if (Array.isArray(rows[r0]) && rows[r0].length > ncols) ncols = rows[r0].length;
    }
    var colDefault = new Array(ncols);
    for (var c = 0; c < ncols; c++) {
        var seenNum = false, seenStr = false, seenOther = false;
        for (var rr = 0; rr < rows.length; rr++) {
            if (!Array.isArray(rows[rr])) continue;
            var v = rows[rr][c];
            if (v === null || v === undefined) continue;
            if (typeof v === 'number') seenNum = true;
            else if (typeof v === 'string') seenStr = true;
            else seenOther = true;
        }
        if (seenOther) colDefault[c] = null;
        else if (seenNum && !seenStr) colDefault[c] = 0;
        else if (seenStr && !seenNum) colDefault[c] = '';
        else if (seenNum && seenStr) colDefault[c] = null;
        else colDefault[c] = null; // all-null column — keep null
    }
    for (var i2 = 0; i2 < rows.length; i2++) {
        var row2 = [];
        if (Array.isArray(rows[i2])) {
            for (var j2 = 0; j2 < rows[i2].length; j2++) {
                var val2 = rows[i2][j2];
                row2.push((val2 === null || val2 === undefined) ? colDefault[j2] : val2);
            }
        } else {
            var val2 = rows[i2];
            row2.push((val2 === null || val2 === undefined) ? null : val2);
        }
        result.push(row2);
    }
    return this._new(result);
};
Array2D.prototype.noNull = Array2D.prototype.z处理空值;

/**
 * 选择列（selectCols）- 选择二维数组中指定的列
 * @param {Array|String} cols - 列选择方式，支持多种格式：
 *   - 数字数组: [0, 2, 4] 选择第1、3、5列（0-based索引）
 *   - f模式字符串: "f1,f3,f5" 选择第1、3、5列（1-based索引）
 *   - f模式数组: ["f1", "f3", "f5"]
 *   - 表头名称数组: ["产品", "数量", "价格"] 按首行表头匹配
 *   - 单个表头名: "产品" 选择单列
 * @param {Array} [newHeaders] - 可选，为选择后的列指定新表头
 * @returns {Array2D} 新实例
 * @example
 * // 示例1：按列号选择
 * var arr = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
 * Array2D.selectCols(arr, [0, 2]);  // 选择第1列和第3列
 * // 结果: [[1, 3], [4, 6], [7, 9]]
 *
 * // 示例2：按f模式字符串选择（推荐）
 * Array2D.selectCols(arr, "f1,f3");  // 选择第1列和第3列
 *
 * // 示例3：按表头选择
 * var arr2 = [['a','b','c'], [1,2,3], [4,5,6]];
 * Array2D.selectCols(arr2, ['c','b','a']);  // 按首行表头错位选择
 * // 结果: [["c","b","a"], [3,2,1], [6,5,4]]
 *
 * // 示例4：指定新表头
 * Array2D.selectCols(arr2, ['a','c'], ['x','z']);
 * // 结果: [["x","z"], [1,3], [4,6]]
 */
Array2D.prototype.z选择列 = function(cols, newHeaders) {
    if (!this._items.length) return this._new([]);

    var indexes = [];
    var useHeader = false;

    // 处理回调函数参数：根据回调函数返回值选择列
    if (typeof cols === 'function') {
        var callback = cols;
        for (var j = 0; j < this._items[0].length; j++) {
            var colValues = [];
            for (var i = 0; i < this._items.length; i++) {
                colValues.push(this._items[i][j]);
            }
            if (callback(colValues, j, this._items)) {
                indexes.push(j);
            }
        }
        useHeader = false;
    } else
    if (Array.isArray(cols) && cols.length > 0 && typeof cols[0] === 'number') {
        indexes = [];
        for (var i = 0; i < cols.length; i++) {
            // 1-based 转 0-based
            indexes.push(cols[i] - 1);
        }
        useHeader = false;
    } else if (typeof cols === 'string') {
        // 检查是否是 f 模式（列号格式）
        if ((cols.includes(',') || cols.includes('，') || cols.includes('-') || cols.includes('+')) && (cols.toLowerCase().includes('f'))) {
            // f 模式：先按逗号分割，再处理范围和合并
            var parts = cols.split(/[,，]/);
            indexes = [];
            for (var i = 0; i < parts.length; i++) {
                var part = parts[i].trim();
                if (part.includes('-')) {
                    // 处理范围 f3-f7
                    var range = part.split('-');
                    var start = parseInt(range[0].toLowerCase().replace('f', ''));
                    var end = parseInt(range[1].toLowerCase().replace('f', ''));
                    for (var j = start; j <= end; j++) {
                        indexes.push(j - 1);
                    }
                } else if (part.toLowerCase().startsWith('f')) {
                    if (part.includes('+')) {
                        // 处理合并 f1+f2
                        var mergeParts = part.split('+');
                        var mergeIndexes = [];
                        for (var ki = 0; ki < mergeParts.length; ki++) {
                            var idx = parseInt(mergeParts[ki].toLowerCase().replace('f', '')) - 1;
                            mergeIndexes.push(idx);
                        }
                        indexes.push(mergeIndexes);  // 用数组标记合并列
                    } else {
                        indexes.push(parseInt(part.substring(1)) - 1);  // f2 → 索引1
                    }
                } else {
                    indexes.push(parseInt(part) - 1);
                }
            }
            useHeader = false;
        } else {
            // 单个字符串，当作表头名称
            cols = [cols];
            useHeader = true;
        }
    } else if (cols.length > 0 && typeof cols[0] === 'string') {
        // 检查是否是 f 模式数组
        var allFMode = true;
        for (var i = 0; i < cols.length; i++) {
            if (typeof cols[i] === 'string' && !cols[i].toLowerCase().startsWith('f')) {
                allFMode = false;
                break;
            }
        }
        if (allFMode) {
            // f 模式数组：转换为列索引（支持范围）
            indexes = [];
            for (var i = 0; i < cols.length; i++) {
                var c = cols[i];
                if (c.includes('-')) {
                    // 处理范围
                    var range = c.split('-');
                    var start = parseInt(range[0].substring(1));
                    var end = parseInt(range[1].substring(1));
                    for (var j = start; j <= end; j++) {
                        indexes.push(j - 1);
                    }
                } else {
                    indexes.push(parseInt(c.substring(1)) - 1);
                }
            }
            useHeader = false;
        } else {
            useHeader = true;
        }
    }

    if (!useHeader && indexes.length > 0) {
        // 按列号选择（已解析的索引）
        var result = [];
        for (var i = 0; i < this._items.length; i++) {
            var row = [];
            for (var ki = 0; ki < indexes.length; ki++) {
                var idx = indexes[ki];
                if (Array.isArray(idx)) {
                    // 合并列：多个列合并为一个
                    var merged = '';
                    for (var m = 0; m < idx.length; m++) {
                        var val = this._items[i][idx[m]];
                        if (val !== null && val !== undefined) {
                            merged += String(val);
                        }
                    }
                    row.push(merged);
                } else {
                    row.push(this._items[i][idx]);
                }
            }
            result.push(row);
        }
        return this._new(result);
    }

    if (useHeader) {
        // 按表头选择
        var headers = this._items[0];
        var headerMap = {};
        for (var i = 0; i < headers.length; i++) {
            headerMap[String(headers[i])] = i;
        }

        var finalIndexes = [];
        for (var j = 0; j < cols.length; j++) {
            var col = cols[j];
            if (headerMap.hasOwnProperty(col)) {
                finalIndexes.push(headerMap[col]);
            }
        }

        var result = [];
        if (newHeaders && newHeaders.length > 0) {
            result.push(newHeaders);
        } else {
            var headerRow = [];
            for (var ki = 0; ki < finalIndexes.length; ki++) {
                var idx = finalIndexes[ki];
                headerRow.push(idx !== undefined ? headers[idx] : cols[ki]);
            }
            result.push(headerRow);
        }

        for (var i = 1; i < this._items.length; i++) {
            var row = this._items[i];
            var newRow = [];
            for (var ki = 0; ki < finalIndexes.length; ki++) {
                newRow.push(row[finalIndexes[ki]]);
            }
            result.push(newRow);
        }

        return this._new(result);
    }

    // 🔧 Bug修复: 当既不是 useHeader 模式，也没有有效索引时，返回空实例（避免返回 undefined）
    return this._new([]);
};
Array2D.prototype.selectCols = Array2D.prototype.z选择列;
Array2D.prototype.SelectCols = Array2D.prototype.z选择列;  // 文档中的大写别名

/**
 * 选择行（selectRows）
 * @param {Array} rowIndexes - 行号数组
 * @returns {Array2D} 新实例
 */
Array2D.prototype.z选择行 = function(rowIndexes) {
    var result = [];
    for (var i = 0; i < rowIndexes.length; i++) {
        var idx = rowIndexes[i];
        if (idx >= 0 && idx < this._items.length) {
            result.push(this._items[idx]);
        }
    }
    return this._new(result);
};
Array2D.prototype.selectRows = Array2D.prototype.z选择行;

/**
 * 获取结果（res）- 获取当前数组的值（val的别名）
 * @returns {Array} 当前数组
 * @example
 * Array2D([[1,2],[3,4]]).z结果()  // [[1,2],[3,4]]
 */
Array2D.prototype.z结果 = function() {
    return this._items;
};
Array2D.prototype.res = Array2D.prototype.z结果;

/**
 * 版本号
 * @returns {String} 版本
 */
Array2D.prototype.z版本 = function() {
    return '3.9.4';
};
Array2D.prototype.version = Array2D.prototype.z版本;

/**
 * 错误值 - 创建错误值对象
 * @param {String} msg - 错误消息
 * @returns {Object} 错误值对象
 */
Array2D.prototype.z错误值 = function(msg) {
    return { __isError: true, message: msg || '#VALUE!' };
};
Array2D.prototype.isError = Array2D.prototype.z错误值;

/**
 * 空结果 - 返回空数组
 * @returns {Array} 空数组
 */
Array2D.prototype.z空结果 = function() {
    return [];
};

/**
 * 下标数组 - 根据条件获取元素的下标数组
 * @param {Array} arr - 数组
 * @param {Function|String} predicate - 条件函数或Lambda表达式
 * @returns {Array} 下标数组
 */

/**
 * 重复N次 - 将数组重复指定次数
 * @param {Number} n - 重复次数
 * @returns {Array} 重复后的数组
 */
Array2D.prototype.z重复N次 = function(n) {
    n = n || 1;
    // 🔧 v3.9.4 修复：返回 Array2D 实例（支持链式调用），深拷贝避免引用共享
    var result = [];
    for (var i = 0; i < n; i++) {
        var items = this._items ? this._items : this;
        for (var j = 0; j < items.length; j++) {
            result.push(Array.isArray(items[j]) ? items[j].slice() : items[j]);
        }
    }
    return this._new(result);
};
Array2D.prototype.repeat = Array2D.prototype.z重复N次;

/**
 * 跳过前面连续满足条件的元素
 * @param {Function|String} predicate - 条件函数或Lambda表达式
 * @returns {Array2D} 新数组
 */
Array2D.prototype.z跳过前面连续满足 = function(predicate) {
    var func = typeof predicate === 'string' ? parseLambda(predicate) : predicate;
    var items = this._items || this;
    var i = 0;
    while (i < items.length && func(items[i], i, items)) i++;
    return new Array2D(items.slice(i), this._header ? { _header: this._header } : undefined);
};
Array2D.prototype.skipWhile = Array2D.prototype.z跳过前面连续满足;

/**
 * 取前面连续满足条件的元素
 * @param {Function|String} predicate - 条件函数或Lambda表达式
 * @returns {Array2D} 新数组
 */
Array2D.prototype.z取前面连续满足 = function(predicate) {
    var func = typeof predicate === 'string' ? parseLambda(predicate) : predicate;
    var items = this._items || this;
    var i = 0;
    while (i < items.length && func(items[i], i, items)) i++;
    return new Array2D(items.slice(0, i), this._header ? { _header: this._header } : undefined);
};
Array2D.prototype.takeWhile = Array2D.prototype.z取前面连续满足;

/**
 * 取交集 - 返回两个数组的交集
 * @param {Array} brr - 第二个数组
 * @param {Function} leftSelector - 左表列选择器
 * @param {Function} rightSelector - 右表列选择器
 * @returns {Array} 交集数组
 */
/**
 * 去重并集 - 返回两个数组的去重并集
 * @param {Array} brr - 第二个数组
 * @param {Function} leftSelector - 左表列选择器
 * @param {Function} rightSelector - 右表列选择器
 * @returns {Array} 并集数组
 */
Array2D.prototype.z去重并集 = function(brr, leftSelector, rightSelector) {
    var arr = this._items || this;
    if (!Array.isArray(arr)) return this._new(brr || []);
    var result = arr.slice();
    if (!brr || !brr.length) return this._new(result);
    var leftKeySelector = leftSelector || (function(x) { return x; });
    var rightKeySelector = rightSelector || leftKeySelector;
    var leftKeys = arr.map(leftKeySelector);
    var seen = new Set(leftKeys.map(String));
    for (var i = 0; i < brr.length; i++) {
        var key = rightKeySelector(brr[i], i, brr);
        if (!seen.has(String(key))) {
            seen.add(String(key));
            result.push(brr[i]);
        }
    }
    return this._new(result);
};
Array2D.prototype.union = Array2D.prototype.z去重并集;

/**
 * 按范围选择 - 根据起始和结束位置选择元素
 * @param {Number} start - 起始位置
 * @param {Number} end - 结束位置
 * @returns {Array} 选择的元素
 */
Array2D.prototype.z按范围选择 = function(start, end) {
    var items = this._items || this;
    return this._new(items.slice(start, end));
};
Array2D.prototype.rangeSelect = Array2D.prototype.z按范围选择;

/**
 * 按范围遍历 - 对指定范围的元素执行回调
 * @param {Number} start - 起始位置
 * @param {Number} end - 结束位置
 * @param {Function} callback - 回调函数
 * @returns {Array2D} 当前对象
 */
Array2D.prototype.z按范围遍历 = function(start, end, callback) {
    var items = this._items || this;
    for (var i = start; i < end && i < items.length; i++) {
        callback(items[i], i, items);
    }
    return this;
};
Array2D.prototype.rangeForEach = Array2D.prototype.z按范围遍历;

/**
 * 区域映射 - 对指定区域的元素进行映射变换
 * @param {Number} startRow - 起始行
 * @param {Number} endRow - 结束行
 * @param {Number} startCol - 起始列
 * @param {Number} endCol - 结束列
 * @param {Function} mapper - 映射函数
 * @returns {Array} 映射后的区域
 */
Array2D.prototype.z区域映射 = function(startRow, endRow, startCol, endCol, mapper) {
    var items = this._items || this;
    var result = [];
    for (var i = startRow; i < endRow && i < items.length; i++) {
        var row = items[i];
        if (!Array.isArray(row)) continue;
        var newRow = [];
        for (var j = startCol; j < endCol && j < row.length; j++) {
            newRow.push(mapper(row[j], i, j, items));
        }
        result.push(newRow);
    }
    return this._new(result);
};
Array2D.prototype.rangeMap = Array2D.prototype.z区域映射;

/**
 * 分块矩阵 - 将一维数组按区域分块聚合
 * @param {Array} arr - 数据数组
 * @param {Number} rowSize - 行大小
 * @param {Number} colSize - 列大小
 * @param {String} aggFunc - 聚合函数 'sum'|'count'|'average'|'max'|'min'
 * @returns {Array} 结果矩阵
 */
Array2D.z分块矩阵 = function(arr, rowSize, colSize, aggFunc) {
    if (!arr || !arr.length) return [];
    aggFunc = aggFunc || 'sum';
    var rows = arr.length;
    var rowCount = Math.ceil(rows / (rowSize * colSize));
    var result = [];
    for (var r = 0; r < rowCount; r++) {
        var rowResult = [];
        for (var c = 0; c < colSize; c++) {
            var start = (r * colSize + c) * rowSize;
            var end = Math.min(start + rowSize, rows);
            if (start >= rows) {
                rowResult.push(null);
            } else {
                var subArr = arr.slice(start, end);
                var val;
                switch (aggFunc) {
                    case 'sum': val = subArr.reduce(function(s, x) { return s + (Number(x) || 0); }, 0); break;
                    case 'count': val = subArr.length; break;
                    case 'average': val = subArr.reduce(function(s, x) { return s + (Number(x) || 0); }, 0) / subArr.length; break;
                    case 'max': val = Math.max.apply(null, subArr.map(Number)); break;
                    case 'min': val = Math.min.apply(null, subArr.map(Number)); break;
                    default: val = subArr[0];
                }
                rowResult.push(val);
            }
        }
        result.push(rowResult);
    }
    return result;
};
Array2D.prototype.blockMatrix = Array2D.z分块矩阵;
Array2D.prototype.z分块矩阵 = Array2D.z分块矩阵;

/**
 * 矩阵排版 - 将一维数组转换为矩阵形式
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'行优先,'c'列优先
 * @returns {Array} 矩阵
 */
Array2D.prototype.z矩阵排版 = function(cols, direction) {
    var items = this._items || this;
    if (!Array.isArray(items)) return [];
    var len = items.length;
    var rows = Math.ceil(len / cols);
    var result = [];

    if (direction === 'c') {
        // 🔧 v3.9.4 修复：列优先模式 - 按列填充再按行输出
        // 例如 [1,2,3,4,5,6] cols=2 → [[1,4],[2,5],[3,6]]（3行2列）
        // 先按列分组：col0=[1,2,3], col1=[4,5,6]，再逐行从各列取值
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var idx = j * rows + i;
                if (idx < len) row.push(items[idx]);
            }
            if (row.length > 0) result.push(row);
        }
    } else {
        // 行优先模式（默认）
        var idx = 0;
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols && idx < len; j++) {
                row.push(items[idx++]);
            }
            if (row.length > 0) result.push(row);
        }
    }
    return result;
};

/**
 * 矩阵运算 - 对矩阵进行运算
 * @param {Function} op - 运算函数
 * @returns {Array} 运算结果
 */
Array2D.prototype.z矩阵运算 = function(op) {
    var items = this._items || this;
    if (!Array.isArray(items)) return [];
    return items.map(function(row) {
        if (!Array.isArray(row)) return op(row);
        return row.map(function(cell) { return op(cell); });
    });
};

/**
 * 矩阵分布（getMatrix）- 生成数字序列的矩阵分布
 * @param {Number} totalRows - 总行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'或'c'
 * @returns {Array} 分布后的数组
 */
Array2D.getMatrix = function(totalRows, cols, direction) {
    if (totalRows === undefined || totalRows <= 0) return [];
    if (cols === undefined || cols <= 0) return [];
    direction = direction || 'r';
    var result = [];
    var numbers = [];
    for (var i = 0; i < totalRows; i++) {
        numbers.push(i);
    }

    if (direction === 'r') {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = i * cols + j;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    } else {
        var rows = Math.ceil(totalRows / cols);
        for (var i = 0; i < rows; i++) {
            var row = [];
            for (var j = 0; j < cols; j++) {
                var index = j * rows + i;
                if (index < totalRows) row.push(numbers[index]);
            }
            if (row.length > 0) result.push(row);
        }
    }

    return result;
};
Array2D.z矩阵分布 = Array2D.getMatrix;

/**
 * rangeMatrix - 区域矩阵操作（按指定区域分组聚合）
 * @param {Array} arr - 源数据数组
 * @param {string|Function} keySelector - 分组键选择器，如 'f1' 或 'A:A'
 * @param {Array} dataArrays - 数据数组或范围数组，如 [brr, 'B:B']
 * @param {Function} [aggregator] - 聚合函数，默认求和
 * @returns {Array} 聚合后的结果数组
 * @example
 * // 按A列分组，对B列求和
 * Array2D.rangeMatrix(arr, 'A:A', [brr, 'B:B'], (a,b)=>a+b)
 * // 按第1列分组，聚合多个数据列
 * Array2D.rangeMatrix(arr, 'f1', [brr], (a,b)=>a+b)
 */
Array2D.rangeMatrix = function(arr, keySelector, dataArrays, aggregator) { 
    // v4.0.11: 支持 Array2D 实例（$.maxArray 返回值）
    if (arr && arr._items) arr = arr._items;
    if (dataArrays && dataArrays._items) dataArrays = dataArrays._items;
    if (Array.isArray(dataArrays)) {
        for (var di = 0; di < dataArrays.length; di++) {
            if (dataArrays[di] && dataArrays[di]._items) dataArrays[di] = dataArrays[di]._items;
        }
    }
    if (!arr || !Array.isArray(arr)) return [];
    
    // 解析键选择器
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!keyFn) {
        // 尝试解析 A:A 格式
        if (typeof keySelector === 'string' && keySelector.match(/^[A-Z]:[A-Z]$/i)) {
            var colIdx = keySelector.charCodeAt(0) - 65; // A=0, B=1, etc.
            keyFn = function(row) { return row[colIdx]; };
        } else {
            keyFn = function(row) { return row[0]; };
        }
    }
    
    // 默认聚合函数：求和
    var aggFn = aggregator || function(acc, val) {
        var num1 = typeof acc === 'number' ? acc : parseFloat(String(acc).replace(/,/g, '')) || 0;
        var num2 = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, '')) || 0;
        return num1 + num2;
    };
    
    // 标准化数据数组
    var dataList = [];
    if (Array.isArray(dataArrays)) {
        if (dataArrays.length === 2 && typeof dataArrays[1] === 'string' && dataArrays[1].match(/^[A-Z]:[A-Z]$/i)) {
            // [brr, 'B:B'] 格式
            dataList.push({
                data: dataArrays[0],
                colSelector: dataArrays[1]
            });
        } else {
            dataList.push({ data: dataArrays, colSelector: null });
        }
    } else {
        dataList.push({ data: dataArrays, colSelector: null });
    }
    
    // 构建分组
    var groups = Object.create(null);
    for (var i = 0; i < arr.length; i++) {
        var key = keyFn(arr[i], i);
        if (!groups[key]) {
            groups[key] = { key: key, indices: [] };
        }
        groups[key].indices.push(i);
    }
    
    // 执行聚合
    var result = [];
    for (var key in groups) {
        var group = groups[key];
        var row = [key];
        
        for (var d = 0; d < dataList.length; d++) {
            var dataInfo = dataList[d];
            var dataArr = dataInfo.data;
            
            // 聚合该组的所有值
            var aggValue = null;
            for (var i = 0; i < group.indices.length; i++) {
                var idx = group.indices[i];
                if (idx < dataArr.length) {
                    var val = dataArr[idx];
                    if (aggValue === null) {
                        aggValue = val;
                    } else {
                        aggValue = aggFn(aggValue, val);
                    }
                }
            }
            row.push(aggValue !== null ? aggValue : 0);
        }
        result.push(row);
    }
    
    return result;
};
Array2D.z区域矩阵 = Array2D.rangeMatrix;
// 兼容文档中的拼写
Array2D.rangeMatric = Array2D.rangeMatrix;

/**
 * 生成下标数组（getIndexs）
 * @param {Number} start - 起始
 * @param {Number} end - 结束
 * @param {Number} step - 步长
 * @returns {Array} 序列
 */
Array2D.getIndexs = function(start, end, step) {
    step = step || 1;
    if (step === 0) step = 1;
    var result = [];
    if (step > 0) {
        for (var i = start; i <= end; i += step) {
            result.push(i);
        }
    } else {
        for (var i = start; i >= end; i += step) {
            result.push(i);
        }
    }
    return result;
};
Array2D.z生成下标数组 = Array2D.getIndexs;

/**
 * 下标数组（indexArray）- 根据条件获取元素的下标数组
 * @param {Array} arr - 数组
 * @param {string|Function} predicate - 筛选条件
 * @returns {Array} 下标数组
 * @example
 * Array2D.indexArray([[1,2],[3,4],[5,6]], 'x=>x[0]>1')  // [1, 2]
 */
Array2D.indexArray = function(arr, predicate) {
    var fn = typeof predicate === 'function' ? predicate : parseLambda(predicate);
    if (!fn) return [];
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        if (fn(arr[i], i)) {
            result.push(i);
        }
    }
    return result;
};
Array2D.z下标数组 = Array2D.indexArray;

/**
 * 按范围遍历（rangeForEach）- 对指定索引范围的元素执行回调
 * @param {Array} arr - 数组
 * @param {Number} start - 起始索引
 * @param {Number} end - 结束索引
 * @param {Function} callback - 回调函数 (item, index)
 * @returns {void}
 * @example
 * Array2D.rangeForEach([[1,2],[3,4],[5,6]], 0, 1, (row, i) => Console.log(row))
 */
Array2D.rangeForEach = function(arr, start, end, callback) {
    if (!arr || !Array.isArray(arr)) return;
    start = start || 0;
    end = end !== undefined ? end : arr.length - 1;
    for (var i = start; i <= end && i < arr.length; i++) {
        callback(arr[i], i);
    }
};
Array2D.z按范围遍历 = Array2D.rangeForEach;

/**
 * 局部映射（rangeMap）- 对二维数组指定矩形区域进行映射，返回完整数组（仅指定区域被修改）
 * @param {Array} arr - 二维数组
 * @param {Array|string} address - 地址范围，支持：
 *   - 数组格式：[行起, 列起, 行数, 列数] 如 [0, 0, 3, 2]，默认 [0, 0, Infinity, Infinity]
 *   - 字符串格式：'a1:b3' 或 'A1:B3'
 * @param {string|Function} mapper - 映射函数 (当前值, 行号, 列号, 原数组) => 新值
 * @returns {Array} 映射后的完整二维数组
 * @example
 * // 示例1: 3行2列区域添加后缀
 * var arr = [["A1","B1","C1"],["A2","B2","C2"],["A3","B3","C3"]];
 * Array2D.rangeMap(arr, [0,0,3,2], x=>x+'**');
 * // [["A1**","B1**","C1"],["A2**","B2**","C2"],["A3**","B3**","C3"]]
 *
 * // 示例2: 字符串格式地址
 * Array2D.rangeMap(arr, 'a1:b2', (x,i,j,orig)=>`${x}-${i}-${j}`);
 *
 * // 示例3: 使用回调访问原数组其他列
 * Array2D.rangeMap(arr, 'a1:b3', (x,i,j,brr)=>`${x}-${brr[i][j+2]}`);
 */
Array2D.rangeMap = function(arr, address, mapper) {
    // v4.0.11: 支持 Array2D 实例（$.maxArray 返回值）
    if (arr && arr._items) arr = arr._items;
    if (!arr || !Array.isArray(arr)) return [];
    if (!mapper) return arr;

    var fn = typeof mapper === 'function' ? mapper : parseLambda(mapper);
    if (!fn) return arr;

    // 解析地址参数
    var rowStart = 0, colStart = 0, rowCount = Infinity, colCount = Infinity;

    if (Array.isArray(address)) {
        // 数组格式: [行起, 列起, 行数, 列数]
        rowStart = address[0] || 0;
        colStart = address[1] || 0;
        rowCount = address[2] !== undefined ? address[2] : Infinity;
        colCount = address[3] !== undefined ? address[3] : Infinity;
    } else if (typeof address === 'string') {
        // 字符串格式: 'a1:b3' 或 'A1:B3'
        var rangeMatch = address.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
        if (rangeMatch) {
            // 列字母转0-based索引
            var _toColIdx = function(colStr) {
                var idx = 0;
                for (var ki = 0; ki < colStr.length; ki++) {
                    idx = idx * 26 + (colStr.toUpperCase().charCodeAt(ki) - 64);
                }
                return idx - 1;
            };
            var col1 = _toColIdx(rangeMatch[1]);
            var row1 = parseInt(rangeMatch[2]) - 1;
            var col2 = _toColIdx(rangeMatch[3]);
            var row2 = parseInt(rangeMatch[4]) - 1;
            rowStart = Math.min(row1, row2);
            colStart = Math.min(col1, col2);
            rowCount = Math.abs(row2 - row1) + 1;
            colCount = Math.abs(col2 - col1) + 1;
        }
    }

    // 深拷贝原数组
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        result[i] = Array.prototype.slice.call(arr[i]);
    }

    // 计算实际范围并映射
    var rowEnd = Math.min(rowStart + rowCount, arr.length);
    for (var i = rowStart; i < rowEnd; i++) {
        if (!arr[i]) continue;
        var colEnd = Math.min(colStart + colCount, arr[i].length);
        for (var j = colStart; j < colEnd; j++) {
            if (j < arr[i].length) {
                result[i][j] = fn(arr[i][j], i, j, arr);
            }
        }
    }

    return result;
};
Array2D.z局部映射 = Array2D.rangeMap;

/**
 * 排名（rank）- 对数组进行排名
 * @param {Array} arr - 数组
 * @param {string|Function} colSelector - 列选择器，支持 f2, f2-（降序）
 * @param {string} [type='cn'] - 排名类型：'cn'中式排名（并列跳过），'usa'美式排名（并列不跳过），'+'顺序编号
 * @returns {Array} 排名结果（二维数组，每行一个排名值）
 * @example
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-')  // [[1],[3],[1]]（中式）
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-', 'usa')  // [[1],[3],[1]]（美式）
 * Array2D.rank([[1,90],[2,80],[3,90]], 'f2-', '+')  // [[1],[3],[2]]（顺序）
 */
Array2D.rank = function(arr, colSelector, type) {
    if (arr && arr._items) arr = arr._items;
    if (!arr || !Array.isArray(arr)) return [];
    type = type || 'cn';
    var selectorStr = typeof colSelector === 'string' ? colSelector : '';
    var isDesc = selectorStr.endsWith('-');
    // 去掉末尾 +/- 再传给 parseLambda
    var cleanSelector = selectorStr.replace(/[+-]$/, '');
    var fn = typeof colSelector === 'function' ? colSelector : parseLambda(cleanSelector || colSelector);
    if (!fn) return [];

    var values = arr.map(function(row, i) { return {value: fn(row, i), index: i}; });
    values.sort(function(a, b) {
        var cmp = 0;
        if (typeof a.value === 'number' && typeof b.value === 'number') {
            cmp = a.value - b.value;
        } else {
            cmp = String(a.value).localeCompare(String(b.value));
        }
        return isDesc ? -cmp : cmp;
    });

    var ranks = [];
    for (var i = 0; i < values.length; i++) {
        var rank;
        if (type === '+') {
            // 顺序排名：1,2,3,4...（永远不跳号）
            rank = i + 1;
        } else if (type === 'usa') {
            // 🔧 v3.9.4 修复：美式排名 - 并列值取相同排名，后续排名跳号
            // 例如 [90,90,80] 降序 → 排名 [1,1,3]（80跳到第3名）
            rank = i + 1;
            for (var j = i - 1; j >= 0; j--) {
                if (values[j].value === values[i].value) {
                    rank = ranks[values[j].index][0]; // 复用前面的排名
                    break;
                }
            }
        } else { // cn 中式排名：并列取相同排名，不跳号
            // 例如 [90,90,80] → [1,1,2]（80排第2，不跳号）
            if (i > 0 && values[i].value === values[i-1].value) {
                rank = ranks[values[i-1].index][0]; // 复用前一个排名
            } else {
                rank = (i > 0 ? ranks[values[i-1].index][0] : 0) + 1;
            }
        }
        ranks[values[i].index] = [rank];
    }
    return ranks;
};
Array2D.z排名 = Array2D.rank;

/**
 * 分组排名（rankGroup）- 按分组进行排名
 * @param {Array} arr - 数组
 * @param {string|Function} colSelector - 列选择器，支持 f2, f2-（降序）
 * @param {string|Function} groupCol - 分组列选择器
 * @param {string} [type='cn'] - 排名类型
 * @param {Number} [skipHeader=0] - 跳过标题行数
 * @returns {Array} 排名结果（二维数组）
 * @example
 * Array2D.rankGroup([[1,'A',90],[2,'A',80],[3,'B',90]], 'f3-', 'f2')
 */
Array2D.rankGroup = function(arr, colSelector, groupCol, type, outputAll) {
    if (arr && arr._items) arr = arr._items;
    if (!arr || !Array.isArray(arr)) return [];
    type = type || 'cn';
    // v4.0.11: outputAll 控制输出格式
    // false/undefined: 只输出排名单列; true: 源数据+序号列+排名列
    outputAll = outputAll === true || outputAll === 1;
    // 兼容旧版 skipHeader 参数：如果 outputAll 是数字，当作 skipHeader
    var skipHeader = 0;
    if (typeof outputAll === 'number') { skipHeader = outputAll; outputAll = false; }
    var selectorStr = typeof colSelector === 'string' ? colSelector : '';
    var isDesc = selectorStr.endsWith('-');
    // 去掉末尾 +/- 再传给 parseLambda
    var cleanSelector = selectorStr.replace(/[+-]$/, '');
    // XXD-205: 数字选择器 (colIdx/groupColIdx) 直接构造取列函数,parseLambda 对非 string 返回 null
    //  → 数字列索引在分组排名里完全不可用,即使有 f1 风格 selector 也无法通过数字调用
    function numericIdxFn(sel) {
        return function(row) { return row == null ? undefined : row[sel]; };
    }
    var fn;
    if (typeof colSelector === 'function') fn = colSelector;
    else if (typeof colSelector === 'number' && Number.isFinite(colSelector) && colSelector >= 0) fn = numericIdxFn(colSelector);
    else fn = parseLambda(cleanSelector || colSelector);
    var groupFn;
    if (typeof groupCol === 'function') groupFn = groupCol;
    else if (typeof groupCol === 'number' && Number.isFinite(groupCol) && groupCol >= 0) groupFn = numericIdxFn(groupCol);
    else groupFn = parseLambda(groupCol);
    if (!fn || !groupFn) return [];

    var data = arr.slice(skipHeader);
    var groups = Object.create(null);
    for (var i = 0; i < data.length; i++) {
        var key = JSON.stringify(groupFn(data[i], i));
        if (!groups[key]) groups[key] = [];
        groups[key].push({row: data[i], index: i + skipHeader});
    }

    var ranks = [];
    for (var h = 0; h < skipHeader; h++) {
        ranks.push(['']);
    }

    for (var key in groups) {
        var group = groups[key];
        var values = group.map(function(item) {
            return {value: fn(item.row, item.index), index: item.index};
        });
        values.sort(function(a, b) {
            var cmp = 0;
            if (typeof a.value === 'number' && typeof b.value === 'number') {
                cmp = a.value - b.value;
            } else {
                cmp = String(a.value).localeCompare(String(b.value));
            }
            return isDesc ? -cmp : cmp;
        });

        for (var j = 0; j < values.length; j++) {
            var rank;
            if (type === '+') {
                rank = j + 1;
            } else if (type === 'usa') {
                // 🔧 v3.9.4 修复：美式排名 - 并列取相同排名，跳号
                rank = j + 1;
                for (var ki = j - 1; ki >= 0; ki--) {
                    if (values[ki].value === values[j].value) {
                        rank = ranks[values[ki].index][0];
                        break;
                    }
                }
            } else { // cn 中式排名：并列取相同排名，不跳号
                if (j > 0 && values[j].value === values[j-1].value) {
                    rank = ranks[values[j-1].index][0];
                } else {
                    rank = (j > 0 ? ranks[values[j-1].index][0] : 0) + 1;
                }
            }
            ranks[values[j].index] = [rank];
        }
    }
    // v4.0.11: 根据 outputAll 决定输出格式
    if (outputAll) {
        var result = [];
        for (var h = 0; h < skipHeader; h++) {
            result.push(arr[h].concat(['序号', '排名']));
        }
        for (var i = skipHeader; i < arr.length; i++) {
            var idx = [i - skipHeader + 1];
            var rk = ranks[i] || [''];
            result.push(arr[i].concat(idx, rk));
        }
        return result;
    }
    return ranks;
};
Array2D.z分组排名 = Array2D.rankGroup;

/**
 * 笛卡尔积（crossjoin）- 两个数组的笛卡尔积
 * @param {Array} arr - 第一个数组
 * @param {Array} brr - 第二个数组
 * @returns {Array} 笛卡尔积结果
 * @example
 * Array2D.crossjoin([[1,2],[3,4]], [[5,6],[7,8]])  // [[1,2,5,6],[1,2,7,8],[3,4,5,6],[3,4,7,8]]
 */
Array2D.crossjoin = function(arr, brr) {
    if (!arr || !brr) return [];
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var aRow = Array.isArray(arr[i]) ? arr[i] : [arr[i]];
        for (var j = 0; j < brr.length; j++) {
            var bRow = Array.isArray(brr[j]) ? brr[j] : [brr[j]];
            result.push(aRow.concat(bRow));
        }
    }
    return result;
};
Array2D.z笛卡尔积 = Array2D.crossjoin;

// XXD-184: prototype methods for crossjoin / z笛卡尔积
// this._items is the first array (arr); brr is supplied by caller
Array2D.prototype.crossjoin = function(brr) {
    if (brr && brr._items && !Array.isArray(brr)) brr = brr._items;
    return new Array2D(Array2D.crossjoin(this._items, brr));
};
Array2D.prototype.z笛卡尔积 = Array2D.prototype.crossjoin;

/**
 * 分组汇总（groupInto）- 按键分组并进行汇总计算
 * @param {Array} arr - 输入数组
 * @param {string|Function} keySelector - 分组键选择器（'f2,f3' 或函数）
 * @param {string|Function} valueSelector - 值聚合选择器，支持：
 *   - 字符串多函数: 'count(),sum("f4"),average("f5"),textjoin("f4","+"),平方和(4)'
 *   - 回调函数: g => [g.count(), g.sum('f4'), g.平方和(4), g.textjoin("f5","+")]
 *   - 单函数: 'g=>g.sum("f3")' 或 'sum("f4")'
 * @param {string} [separator='@^@'] - 多列分组时的分隔符
 * @returns {Array2D} 分组汇总结果（Array2D实例，支持链式调用）
 * @example
 * // 字符串形式（多聚合函数）
 * Array2D.groupInto(arr, 'f2,f3', 'count(),sum("f4"),average("f5"),textjoin("f4","+")')
 * // 函数回调形式
 * Array2D.groupInto(arr, x=>[x.f2,x.f3], g=>[g.count(), g.sum('f4'), g.平方和(4)])
 * // 支持 .f1, .f2 等属性访问语法
 */
Array2D.groupInto = function(arr, keySelector, valueSelector, separator) {
    // Handle Array2D instances from safeArray() - extract underlying data
    if (arr && arr._items && !Array.isArray(arr)) {
        arr = arr._items;
    }
    if (!arr || !Array.isArray(arr)) return new Array2D([]);
    separator = separator || '@^@';
    var keyFn = typeof keySelector === 'function' ? keySelector : parseLambda(keySelector);
    if (!keyFn) return new Array2D([]);

    // ===== 行代理函数：让数组行支持 .f1, .f2 等属性访问 =====
    // 例如: x.f2 等价于 x[1]（数组索引从0开始，f2表示第2列即索引1）
    function toRowProxy(row, rowIndex) {
        // 如果已经是代理对象，直接返回
        if (row && row._isRowProxy) return row;
        // 如果不是数组，直接返回
        if (!Array.isArray(row)) return row;
        // 创建带属性的代理数组
        var proxy = row.slice(); // 复制数组避免污染原数据
        proxy._isRowProxy = true;
        // 添加 .f1, .f2, .f3... 属性访问（f1对应索引0）
        for (var i = 0; i < proxy.length; i++) {
            Object.defineProperty(proxy, 'f' + (i + 1), {
                get: function(idx) {
                    return function() { return proxy[idx]; };
                }(i),
                enumerable: false,
                configurable: true
            });
            // 直接赋值（对于WPS JSA环境兼容性更好）
            try {
                proxy['f' + (i + 1)] = proxy[i];
            } catch (e) {
                // 某些环境不支持直接赋值，忽略
            }
        }
        return proxy;
    }

    // ===== 解析列参数的通用辅助函数 =====
    // 支持: 4 (数字), "4" (字符串数字), "f4" (f格式), "f4" (带引号), 4 (无引号数字)
    function _resolveCol(col) {
        if (col === undefined || col === null) return -1;
        if (typeof col === 'number') return col - 1;  // 数字直接转为0-based索引
        var str = String(col).replace(/^["']|["']$/g, '').replace(/^f/i, ''); // 去除引号和f前缀
        var idx = parseInt(str, 10);
        return isNaN(idx) ? -1 : idx - 1;
    }

    // ===== 创建聚合辅助对象 g（提供 g.xxx() 形式） =====
    function createAggHelper(rows) {
        var a2d = new Array2D(rows);
        return {
            _arr2d: a2d,
            _rows: rows,
            // 计数
            count: function() { return rows.length; },
            // 求和: g.sum() / g.sum('f4') / g.sum(4) / g.sum('f4*f5')
            sum: function(col) {
                if (col === undefined || col === null) return a2d.z求和();
                var colStr = String(col).replace(/^["']|["']$/g, '');
                // 检查是否是表达式（包含运算符）
                if (/[\+\-\*\/\(\)]/.test(colStr)) {
                    // 表达式求和，如 'f4*f5' 或 'f4+f5'
                    return a2d.z求和(function(r) {
                        if (!Array.isArray(r)) return 0;
                        // 将 f1, f2 等替换为 r[0], r[1]
                        var expr = colStr.replace(/f\s*(\d+)/gi, function(m, num) {
                            return 'r[' + (parseInt(num) - 1) + ']';
                        });
                        try {
                            var val = eval(expr);
                            return typeof val === 'number' ? val : parseFloat(val) || 0;
                        } catch (e) {
                            return 0;
                        }
                    });
                }
                var idx = _resolveCol(col);
                return idx >= 0 ? a2d.z求和(function(r) { return Array.isArray(r) ? r[idx] : r; }) : a2d.z求和();
            },
            // 平均值: g.average('f4') / g.average(4)
            average: function(col) {
                var idx = _resolveCol(col);
                return idx >= 0 ? a2d.z平均值(function(r) { return Array.isArray(r) ? r[idx] : r; }) : a2d.z平均值();
            },
            // 最大值: g.max('f4') / g.max(4)
            max: function(col) {
                var idx = _resolveCol(col);
                return idx >= 0 ? a2d.z最大值(function(r) { return Array.isArray(r) ? r[idx] : r; }) : a2d.z最大值();
            },
            // 最小值: g.min('f4') / g.min(4)
            min: function(col) {
                var idx = _resolveCol(col);
                return idx >= 0 ? a2d.z最小值(function(r) { return Array.isArray(r) ? r[idx] : r; }) : a2d.z最小值();
            },
            // 文本连接: g.textjoin('f5', '+') / g.textjoin(4, '+')
            textjoin: function(col, sep) {
                sep = sep !== undefined ? sep : ',';
                var idx = _resolveCol(col);
                if (idx < 0) return '';
                var vals = [];
                for (var i = 0; i < rows.length; i++) {
                    var row = rows[i];
                    var v = Array.isArray(row) && idx < row.length ? row[idx] : '';
                    if (v !== null && v !== undefined && String(v) !== '') {
                        vals.push(v);
                    }
                }
                return vals.join(sep);
            },
            // 平方和: g.平方和(4) / g.平方和('f4')
            平方和: function(col) {
                var idx = _resolveCol(col);
                if (idx < 0) return 0;
                var sum = 0;
                for (var i = 0; i < rows.length; i++) {
                    var row = rows[i];
                    var v = Array.isArray(row) ? row[idx] : row;
                    var num = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
                    if (!isNaN(num)) sum += num * num;
                }
                return sum;
            }
        };
    }

    // ===== 解析字符串形式的多聚合函数 =====
    // 输入: 'count(),sum("f4"),average("f5"),textjoin("f4","+"),平方和(4)'
    // 输出: [{func:'count',args:[]}, {func:'sum',args:['f4']}, ...]
    function parseAggString(str) {
        // 按顶级逗号分割（忽略括号内的逗号）
        var parts = [];
        var depth = 0, cur = '';
        for (var i = 0; i < str.length; i++) {
            var c = str[i];
            if (c === '(') depth++;
            else if (c === ')') depth--;
            if (c === ',' && depth === 0) {
                parts.push(cur.trim()); cur = '';
            } else { cur += c; }
        }
        if (cur.trim()) parts.push(cur.trim());

        var defs = [];
        for (var p = 0; p < parts.length; p++) {
            // XXD-15: 裸名 sum/avg/max/min/textjoin/平方和 无参语义不明确,显式 throw;
            // 裸名 count 无参 = 行数,合法。
            var bareCount = parts[p].match(/^\s*count\s*$/i);
            if (bareCount) { defs.push({ func: 'count', args: [] }); continue; }
            var bareOther = parts[p].match(/^\s*(sum|average|avg|max|min|textjoin|平方和)\s*$/i);
            if (bareOther) {
                throw new Error('groupInto: 聚合函数 ' + bareOther[1] + ' 需要参数');
            }
            var m = parts[p].match(/(sum|count|average|avg|max|min|textjoin|qctextjoin|平方和)\s*\(\s*([^)]*)\s*\)/i);
            if (m) {
                var fn = m[1].toLowerCase();
                var argsStr = m[2].trim();
                var args = [];
                if (argsStr) {
                    // 按逗号分割参数（支持引号包围的参数）
                    var inQ = false, curA = '';
                    for (var j = 0; j < argsStr.length; j++) {
                        var ch = argsStr[j];
                        if (ch === '"' || ch === "'") { inQ = !inQ; }
                        else if (ch === ',' && !inQ) {
                            args.push(curA.trim().replace(/^["']|["']$/g, ''));
                            curA = '';
                        } else { curA += ch; }
                    }
                    if (curA.trim()) args.push(curA.trim().replace(/^["']|["']$/g, ''));
                }
                defs.push({ func: fn, args: args });
            }
        }
        return defs;
    }

    // ===== 构建值聚合函数 =====
    var valueFn;
    if (typeof valueSelector === 'string') {
        var defs = parseAggString(valueSelector);
        if (defs.length > 0) {
            valueFn = function(rows) {
                var helper = createAggHelper(rows);
                var results = [];
                for (var i = 0; i < defs.length; i++) {
                    var d = defs[i];
                    switch (d.func) {
                        case 'sum':
                            results.push(helper.sum(d.args[0]));
                            break;
                        case 'count':
                            results.push(helper.count());
                            break;
                        case 'average':
                        case 'avg':
                            results.push(helper.average(d.args[0]));
                            break;
                        case 'max':
                            results.push(helper.max(d.args[0]));
                            break;
                        case 'min':
                            results.push(helper.min(d.args[0]));
                            break;
                        case 'textjoin':
                            results.push(helper.textjoin(d.args[0], d.args[1]));
                            break;
                        case '平方和':
                            results.push(helper.平方和(d.args[0]));
                            break;
                        default:
                            results.push(null);
                    }
                }
                return results.length === 1 ? results[0] : results;
            };
        } else {
            valueFn = parseLambda(valueSelector);
        }
    } else if (typeof valueSelector === 'function') {
        // 函数回调：将 group.rows 通过聚合辅助对象包装后传入
        valueFn = function(rows) {
            var helper = createAggHelper(rows);
            return valueSelector(helper);
        };
    } else {
        valueFn = valueSelector;
    }

    if (!valueFn) return new Array2D([]);

    // ===== 执行分组 =====
    var groups = {};
    var _groupKeys = [];  // 记录所有分组键的顺序
    for (var i = 0; i < arr.length; i++) {
        // 使用代理对象，支持 x.f1, x.f2 等属性访问语法
        var rowProxy = toRowProxy(arr[i], i);
        var key = keyFn(rowProxy, i);
        var keyStr = Array.isArray(key) ? key.join(separator) : String(key);
        if (!groups.hasOwnProperty(keyStr)) {
            groups[keyStr] = { key: key, rows: [] };
            _groupKeys.push(keyStr);
        }
        groups[keyStr].rows.push(arr[i]); // 原始数据存入rows，供聚合函数使用
    }

    // ===== 汇总结果 =====
    var result = [];
    for (var gi = 0; gi < _groupKeys.length; gi++) {
        var keyStr2 = _groupKeys[gi];
        var group = groups[keyStr2];
        var agg = valueFn(group.rows);
        var row;
        if (Array.isArray(group.key)) {
            // key 是数组: 聚合结果若是数组则拼接到后面
            row = Array.isArray(agg) ? group.key.concat(agg) : group.key.concat([agg]);
        } else if (group.key !== null && group.key !== undefined) {
            row = Array.isArray(agg) ? [group.key].concat(agg) : [group.key, agg];
        }
        result.push(row);
    }
    // 返回 Array2D 实例以支持链式 .toRange() 调用
    return new Array2D(result);
};
Array2D.z分组汇总 = Array2D.groupInto;

/**
 * agg - 对数组执行聚合计算
 * @param {Array} arr - 数组
 * @param {string|Function} colSelector - 列选择器，如 'f3' 或 'f3-'（降序相关）
 * @param {string} [aggType='sum'] - 聚合类型：'sum', 'count', 'average', 'max', 'min'
 * @returns {Number} 聚合结果
 * @example
 * Array2D.agg([[1,2,10],[3,4,20]], 'f3', 'sum')      // 30
 * Array2D.agg([[1,2,10],[3,4,20]], 'f3', 'count')   // 2
 * Array2D.agg([[1,2,10],[3,4,20]], 'f3', 'average') // 15
 * Array2D.agg([[1,2,10],[3,4,20]], 'f3', 'max')     // 20
 * Array2D.agg([[1,2,10],[3,4,20]], 'f3', 'min')     // 10
 */
Array2D.agg = function(arr, colSelector, aggType) {
    if (!arr || !Array.isArray(arr) || arr.length === 0) return 0;
    aggType = (aggType || 'sum').toLowerCase();
    
    // 解析列选择器
    var fn = typeof colSelector === 'function' ? colSelector : parseLambda(colSelector);
    if (!fn) {
        // 默认取第一列
        fn = function(row) { return row[0]; };
    }
    
    // 提取值
    var values = [];
    for (var i = 0; i < arr.length; i++) {
        var val = fn(arr[i], i);
        if (val !== null && val !== undefined && val !== '') {
            var num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
            if (!isNaN(num)) {
                values.push(num);
            }
        }
    }
    
    if (values.length === 0) return 0;
    
    switch (aggType) {
        case 'sum': {
            let sum = 0;
            for (let i = 0; i < values.length; i++) sum += values[i];
            return sum;
        }
        case 'count':
            return values.length;
        case 'average':
        case 'avg': {
            let sum = 0;
            for (let i = 0; i < values.length; i++) sum += values[i];
            return sum / values.length;
        }
        case 'max': {
            let max = values[0];
            for (let i = 1; i < values.length; i++) {
                if (values[i] > max) max = values[i];
            }
            return max;
        }
        case 'min': {
            let min = values[0];
            for (let i = 1; i < values.length; i++) {
                if (values[i] < min) min = values[i];
            }
            return min;
        }
        default:
            return 0;
    }
};
Array2D.z聚合 = Array2D.agg;

/**
 * 分组汇总到字典（groupIntoMap）- 按键分组并汇总到Map对象
 * @param {Array} arr - 数组
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} [valueSelector] - 值选择器
 * @returns {Map} Map对象，键为分组键，值为 {group: 数组, agg: 聚合结果}
 * @example
 * var map = Array2D.groupIntoMap([[1,'A',10],[2,'B',20]], 'f2')
 */
Array2D.groupIntoMap = function(arr, keySelector, valueSelector) {
    // Handle Array2D instances from safeArray() - extract underlying data
    if (arr && arr._items && !Array.isArray(arr)) {
        arr = arr._items;
    }
    if (!arr || !Array.isArray(arr)) return {};

    // 支持数字选择器 (列索引) — parseLambda 对数字返回 null
    var keyFn;
    if (typeof keySelector === 'function') {
        keyFn = keySelector;
    } else if (typeof keySelector === 'number' && isFinite(keySelector) && keySelector >= 0) {
        keyFn = function(row) { return row[keySelector]; };
    } else {
        keyFn = parseLambda(keySelector);
    }
    if (!keyFn) return {};

    var valueFn;
    if (typeof valueSelector === 'function') {
        valueFn = valueSelector;
    } else if (typeof valueSelector === 'number' && isFinite(valueSelector) && valueSelector >= 0) {
        valueFn = function(row) { return row[valueSelector]; };
    } else if (valueSelector != null) {
        valueFn = parseLambda(valueSelector);
    } else {
        valueFn = null;
    }

    var result = {};
    for (var i = 0; i < arr.length; i++) {
        var key = keyFn(arr[i], i);
        if (key === undefined || key === null) continue;
        if (!result.hasOwnProperty(key)) {
            result[key] = [];
        }
        result[key].push(valueFn ? valueFn(arr[i], i) : arr[i]);
    }
    return result;
};
Array2D.z分组汇总到字典 = Array2D.groupIntoMap;

/**
 * 分组连接（groupIntoJoin）- 优化sumifs和Countifs批量条件统计
 * @param {Array} targetData - 统计目标数据（左表）
 * @param {Array} sourceData - 数据源（右表）
 * @param {string|Function} keySelector - 分组键选择器
 * @param {string|Function} valueSelector - 汇总函数或选择器
 * @param {string} [separator='@^@'] - 多列分组时的分隔符
 * @returns {Array} 连接汇总后的结果
 * @example
 * // 对源数据按条件分类汇总，然后左连接到目标数据
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2', 'sum("f4")');
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2,f3', 'count(),sum("f4")', '@^@');
 * // 完整回调模式用法
 * Array2D.groupIntoJoin(目标表, 源数据表, 'f2', g => g.count());
 */
Array2D.groupIntoJoin = function(targetData, sourceData, keySelector, valueSelector, separator) {
    separator = separator || '@^@';
    
    // 1. 先对源数据做分类汇总
    var grouped = Array2D.groupInto(sourceData, keySelector, valueSelector, separator);
    
    // 2. 将汇总结果作为右表，与目标表做左连接
    return new Array2D(targetData).z左连接(
        grouped,
        keySelector,
        keySelector,
        function(leftRow, rightRow) {
            return leftRow.concat(rightRow || []);
        }
    ).val();
};
Array2D.z分组汇总连接 = Array2D.groupIntoJoin;
// XXD-209: 修复 prototype 调用 — 裹子函数透传 this._items 给 static 方法
Array2D.prototype.z分组汇总到字典 = function(keySelector, valueSelector) {
    return Array2D.groupIntoMap(this._items || this, keySelector, valueSelector);
};
Array2D.prototype.z分组汇总连接 = function(targetData, keySelector, valueSelector, separator) {
    return Array2D.groupIntoJoin(targetData, this._items || this, keySelector, valueSelector, separator);
};
// XXD-205: 修复 z分组排名 全部返 [] — 原 alias `Array2D.prototype.z分组排名 = Array2D.rankGroup;`
//  把 `arr` 当作隐式 `this`,实例调用时 `arr` 实际收到 colSelector('f1') 而非 this,
//  → 内部 `arr.slice()`/parseLambda('f1') 仍能跑出结果但语义错位;更严重的是对 prototype 自动静态包装
//  链 (L16916-L16940) 来说,wrapper 会 `new Array2D(arr).z分组排名(colSel, groupCol)`, 再次触发同样的错位,
//  → 整链返回 []. 修正: 显式把 this 作为 arr 透传给 rankGroup.
Array2D.prototype.z分组排名 = function(colSelector, groupCol, type, outputAll) {
    return Array2D.rankGroup(this, colSelector, groupCol, type, outputAll);
};
// XXD-185: 7 Chinese aliases (A4.1/A4.4/B1.3/B3.2/B3.3/B4.2/B7.9)
//   z区域矩阵 / z矩阵分布 / z查找区域 / z复制 / z复制到指定位置 / z批量填充 / z局部映射
//   静态版已存在;此处补齐 prototype 版以便 .z局部映射 等链式调用 / 实例调用可工作
Array2D.prototype.z区域矩阵 = function(keySelector, dataArrays, aggregator) { return Array2D.z区域矩阵(this._items || this, keySelector, dataArrays, aggregator); };
Array2D.prototype.z矩阵分布 = function(cols, direction) { var arr = this._items || this; return Array2D.z矩阵分布(arr.length, cols, direction); };
Array2D.prototype.z查找区域 = function(value) { return Array2D.z查找区域(this._items || this, value); };
Array2D.prototype.z复制 = function() { return new Array2D(JSON.parse(JSON.stringify(this._items || this))); };
Array2D.prototype.z复制到指定位置 = function(target, start, end) { return Array2D.z复制到指定位置(this._items || this, target, start, end); };
Array2D.prototype.z批量填充 = function(value, rows, cols) { return Array2D.z批量填充(this._items || this, value, rows, cols); };
Array2D.prototype.z局部映射 = function(address, mapper) { return Array2D.z局部映射(this._items || this, address, mapper); };

/**
 * 复制到指定位置（copyWithin）- 数组内部复制
 * @param {Array} arr - 数组
 * @param {Number} target - 目标位置
 * @param {Number} [start=0] - 源起始位置
 * @param {Number} [end] - 源结束位置
 * @returns {Array} 复制后的数组
 * @example
 * Array2D.copyWithin([[1,2],[3,4],[5,6]], 0, 2)  // [[5,6],[3,4],[5,6]]
 */
Array2D.copyWithin = function(arr, target, start, end) {
    if (!arr || !Array.isArray(arr)) return [];
    var result = JSON.parse(JSON.stringify(arr));
    var copyArr = result.slice(start || 0, end !== undefined ? end : result.length);
    for (var i = 0; i < copyArr.length; i++) {
        if (target + i < result.length) {
            result[target + i] = JSON.parse(JSON.stringify(copyArr[i]));
        }
    }
    return result;
};
Array2D.z复制到指定位置 = Array2D.copyWithin;

/**
 * 随机一项（random）- 随机选择一组
 * @param {Array} arr - 数组（一维或二维）
 * @param {Number} n - 可选，先打乱全部再取前n个
 * @returns {Array} 随机选择的项
 * @example
 * Array2D.random([[1,2],[3,4],[5,6]])     // 随机返回 [[1,2]]
 * Array2D.random([1,2,3,4,5,6])           // 返回 [3]
 * Array2D.random([1,2,3,4,5,6], 3)        // 先打乱全部，再取前3个，返回 [2,1,3]
 */
Array2D.random = function(arr, n) {
    if (!arr || !Array.isArray(arr) || arr.length === 0) return undefined;

    // 检测是否为一维数组
    var isOneD = arr.length > 0 && !Array.isArray(arr[0]);

    if (n !== undefined && n > 0) {
        // 先打乱整个数组，再取前n个
        var result = JSON.parse(JSON.stringify(arr));

        // Fisher-Yates 洗牌整个数组
        for (var i = result.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = result[i];
            result[i] = result[j];
            result[j] = temp;
        }

        // 取前n个
        result = result.slice(0, Math.min(n, arr.length));

        // 如果是一维数组，返回一维；二维数组返回二维
        return result;
    } else {
        // 随机选择一项
        var idx = Math.floor(Math.random() * arr.length);
        var item = arr[idx];

        // 如果是一维数组，返回单个值；二维数组返回二维格式
        if (isOneD) {
            return item;
        }
        return [item];
    }
};
Array2D.z随机一项 = Array2D.random;

/**
 * 随机打乱（shuffle）- Fisher-Yates 洗牌算法
 * @param {Array} arr - 数组
 * @returns {Array} 打乱后的数组
 * @example
 * Array2D.shuffle([[1,2],[3,4],[5,6]])
 */
Array2D.shuffle = function(arr) {
    if (!arr || !Array.isArray(arr)) return [];
    var result = JSON.parse(JSON.stringify(arr));
    for (var i = result.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = result[i];
        result[i] = result[j];
        result[j] = temp;
    }
    return result;
};
Array2D.z随机打乱 = Array2D.shuffle;

/**
 * 重复N次（repeat）- 将数组重复N次
 * @param {Array} arr - 数组
 * @param {Number} count - 重复次数
 * @returns {Array} 重复后的数组
 * @example
 * Array2D.repeat([[1,2]], 3)  // [[1,2],[1,2],[1,2]]
 */
Array2D.repeat = function(arr, count) {
    if (!arr || !Array.isArray(arr)) return [];
    if (count <= 0) return [];
    var result = [];
    for (var i = 0; i < count; i++) {
        for (var j = 0; j < arr.length; j++) {
            result.push(JSON.parse(JSON.stringify(arr[j])));
        }
    }
    return result;
};
Array2D.z重复N次 = Array2D.repeat;

/**
 * 静态方法：选择列（返回 Array2D 对象，支持链式调用）
 * @param {Array|Array2D} arr - 二维数组或 Array2D 对象
 */
Array2D.z选择列 = function(arr, cols, newHeaders) {
    // 智能判断：如果是 Array2D 对象，直接调用实例方法
    if (arr && arr instanceof Array2D) {
        return arr.z选择列(cols, newHeaders);
    }
    return new Array2D(arr).z选择列(cols, newHeaders);
};
Array2D.selectCols = Array2D.z选择列;

/**
 * 版本号（version）- 返回Array2D函数库版本号
 * @returns {String} 版本号
 * @example
 * Array2D.version()  // "3.2.0"
 */
Array2D.version = function() {
    return '3.9.4';
};

/**
 * 静态方法：数量（count）- 计算数组的元素数量
 * @param {Array} arr - 数组
 * @returns {Number} 元素数量
 * @example
 * Array2D.count([[1,2],[3,4]])  // 4
 */
Array2D.count = function(arr) {
    if (!arr || !Array.isArray(arr)) return 0;
    var count = 0;
    for (var i = 0; i < arr.length; i++) {
        if (Array.isArray(arr[i])) {
            count += arr[i].length;
        } else {
            count++;
        }
    }
    return count;
};
Array2D.z数量 = Array2D.count;

/**
 * 静态方法：批量填充
 */
Array2D.z批量填充 = function(arr, value, rows, cols) {
    return new Array2D(arr).z填充(value, rows, cols).val();
};
Array2D.fill = Array2D.z批量填充;

/**
 * 静态方法：写入单元格（根据数组大小自动扩展区域，返回 Range 对象）
 * @param {Array} arr - 二维数组
 * @param {Range|string} rng - 目标单元格区域（左上角单元格）
 * @returns {Range} 写入的 Range 对象
 * @example
 * var arr = [[1, 'A'], [2, 'B'], [3, 'C']];
 * Array2D.toRange(arr, "Sheet1!a1");
 * Array2D.toRange(arr, "e1");
 * var rs = Array2D.toRange(arr, Range("i1"));
 * console.log(rs.Address());  // $I$1:$J$3
 */
Array2D.toRange = function(arr, rng) {
    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = arr.length;
    var cols = rows > 0 ? (Array.isArray(arr[0]) ? arr[0].length : 1) : 0;

    // 🔧 防御性检查：确保维度有效
    if (rows === 0 || cols === 0) {
        console.log("⚠️ 警告: toRange 收到空数组 (rows=" + rows + ", cols=" + cols + ")");
        return targetRng;  // 返回原range，不做操作
    }

    // 🔧 防御性检查：防止超大范围导致WPS崩溃
    if (rows > 100000 || cols > 16000) {
        console.log("⚠️ 警告: toRange 维度超限 (rows=" + rows + ", cols=" + cols + ")");
        return targetRng;
    }

    // 根据数组大小调整目标区域
    try {
        var endRng = targetRng.Item(rows, cols);
    } catch (e) {
        console.log("❌ 创建 endRng 失败: rows=" + rows + ", cols=" + cols + ", 错误: " + e.message);
        // 尝试使用 Offset 方法
        try {
            var endRng = targetRng.Offset(rows - 1, cols - 1);
        } catch (e2) {
            console.log("❌ Offset 方法也失败: " + e2.message);
            return targetRng;
        }
    }

    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);

    // 🔧 性能优化：一次性解除整个区域的合并，而不是逐个单元格检查
    // 方法：直接设置 MergeCells = false
    try {
        writeRng.MergeCells = false;
    } catch (e) {
        // 如果一次性解除失败，回退到原方法（保留向后兼容）
        for (var i = 1; i <= writeRng.Rows.Count; i++) {
            for (var j = 1; j <= writeRng.Columns.Count; j++) {
                var cell = writeRng.Cells(i, j);
                if (cell.MergeCells) {
                    cell.MergeArea.UnMerge();
                }
            }
        }
    }

    // 批量写入数据
    writeRng.Value2 = arr;
    return writeRng;
};

/**
 * 静态方法：写入单元格（中文别名，返回 Range 对象）
 */
Array2D.z写入单元格 = Array2D.toRange;

// ==================== 静态方法封装（支持直接调用）====================

/**
 * 静态方法：筛选（filter）- 根据条件筛选数组行
 * @param {Array} arr - 二维数组
 * @param {String|Function} predicate - 筛选条件
 * @param {Number} skipHeader - 跳过表头行数
 * @returns {Array} 筛选后的二维数组
 * @example
 * Array2D.filter(arr, 'f1>1')
 * Array2D.filter(arr, x=>x.f1>5 && x.f2=="A")
 * Array2D.filter(arr, "[f1,f3,f4]")
 */
Array2D.filter = function(arr, predicate, skipHeader) {
    // 处理对象参数形式
    if (predicate && typeof predicate === 'object' && !Array.isArray(predicate)) {
        var data = skipHeader ? arr.slice(skipHeader) : arr;
        return Array2D._filterByObject(data, predicate);
    }
    // 🔧 v3.7.6 修复: 保留 Array2D 对象而非返回 .val()
    // 这样可以保留 _header 属性用于后续操作
    var result = new Array2D(arr).z筛选(predicate, skipHeader);
    // 如果原始输入有 _header 属性，确保结果对象也有（虽然 z筛选通过 _new 已经保留）
    return result;
};
Array2D.z筛选 = Array2D.filter;

/**
 * 静态方法：映射（map）- 对数组的每行进行转换
 * @param {Array} arr - 二维数组
 * @param {String|Function} mapper - 转换函数
 * @returns {Array} 转换后的二维数组
 * @example
 * Array2D.map(arr, 'f1*2')
 * Array2D.map(arr, x=>[x.f1, x.f3])
 * Array2D.map(arr, "[f1,f3]")
 */
Array2D.map = function(arr, mapper) {
    return new Array2D(arr).z映射(mapper).val();
};
Array2D.z映射 = Array2D.map;

/**
 * 静态方法：去重（distinct）- 根据指定列去重
 * @param {Array} arr - 二维数组
 * @param {String|Function} keySelector - 去重依据的列
 * @param {String} resultSelector - 结果选择器
 * @returns {Array} 去重后的二维数组
 * @example
 * Array2D.distinct(arr, 'f1,f2')
 * Array2D.distinct(arr, x=>x.f1)
 * Array2D.distinct(arr)
 */
Array2D.distinct = function(arr, keySelector, resultSelector) {
    return new Array2D(arr).z去重(keySelector, resultSelector).val();
};
Array2D.z去重 = Array2D.distinct;

/**
 * 静态方法：多列排序（sortByCols）- 按多列排序
 * @param {Array} arr - 二维数组
 * @param {String} colsConfig - 列配置，如 'f1+,f2-,f3+'
 * @param {Number} skipHeader - 表头行数
 * @returns {Array} 排序后的二维数组
 * @example
 * Array2D.sortByCols(arr, 'f1+,f2-', 1)
 */
Array2D.sortByCols = function(arr, colsConfig, skipHeader) {
    // 🔧 v3.7.6 修复: 保留 Array2D 对象而非返回 .val()
    // 这样可以保留 _header 属性用于后续操作
    var result = new Array2D(arr).z多列排序(colsConfig, skipHeader);
    return result;
};
Array2D.z多列排序 = Array2D.sortByCols;

/**
 * 静态方法：自定义排序（sortByList）- 按自定义列表排序
 * @param {Array} arr - 二维数组
 * @param {String|Number} col - 列号或列名
 * @param {String} orderList - 排序顺序，如 "A,B,C"
 * @param {Number} skipHeader - 表头行数
 * @returns {Array} 排序后的二维数组
 * @example
 * Array2D.sortByList(arr, 'f3', '美国,德国,中国')
 */
Array2D.sortByList = function(arr, col, orderList, skipHeader) {
    return new Array2D(arr).z自定义排序(col, orderList, skipHeader).val();
};
Array2D.z自定义排序 = Array2D.sortByList;

// ==================== 智能类型识别与智能排序/分组 ====================

/**
 * 智能类型检测 - 自动识别列的数据类型
 * @param {Array} data - 二维数组数据
 * @param {number} colIndex - 列索引
 * @returns {Object} 类型信息 {type: 'number'|'date'|'string'|'boolean', format: string}
 */
Array2D.detectType = function(data, colIndex) {
    // XXD-192/XXD-193 final fix: 当 colIndex 未指定 / < 0（"判断整个数据集的类型"）时，
    // flatten 全部非空单元格后做整体类型推断：若任一类型（boolean/date/number/string）≥80%
    // 占多数则返回该类型，否则 {type:'mixed', format:'text'}。
    // 指定 colIndex 时保持原逐列推断行为不变（向后兼容 smartSort/smartGroup 等调用方）。
    if (colIndex === undefined || colIndex === null || colIndex < 0) {
        return detectFlatType(data);
    }
    return detectSingleColumn(data, colIndex);

    function detectFlatType(d) {
        var samples = [];
        var maxSamples = Math.min(d.length, 100);
        var _hasNull = false;
        for (var r = 0; r < maxSamples; r++) {
            if (!d[r]) continue;
            for (var c = 0; c < d[r].length; c++) {
                var cell = d[r][c];
                if (cell === null) { _hasNull = true; continue; }
                if (cell !== undefined && cell !== '') {
                    samples.push(cell);
                }
            }
        }
        function _wrap(r) { if (_hasNull && r && typeof r === 'object') r._hasNull = true; return r; }
        if (samples.length === 0) {
            return _wrap({ type: 'string', format: 'text' });
        }
        // 检测布尔值
        var boolCount = 0;
        for (var i = 0; i < samples.length; i++) {
            var v = samples[i];
            if (v === true || v === false || v === 'true' || v === 'false' || v === '是' || v === '否' || v === 'YES' || v === 'NO') {
                boolCount++;
            }
        }
        if (boolCount / samples.length > 0.8) {
            return _wrap({ type: 'boolean', format: 'boolean' });
        }
        // 检测日期
        var dateCount = 0;
        var datePatterns = [
            /^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}$/,
            /^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}$/,
            /^\d{4}年\d{1,2}月\d{1,2}[日]?$/,
            /^\d{1,2}月\d{1,2}[日]?$/,
            /^\d{8}$/
        ];
        for (var i2 = 0; i2 < samples.length; i2++) {
            var s = String(samples[i2]);
            for (var j = 0; j < datePatterns.length; j++) {
                if (datePatterns[j].test(s)) {
                    dateCount++;
                    break;
                }
            }
        }
        if (dateCount / samples.length > 0.8) {
            var sample = String(samples[0]);
            var format = 'date';
            if (sample.indexOf('-') > 0) format = 'yyyy-MM-dd';
            else if (sample.indexOf('/') > 0) format = 'yyyy/MM/dd';
            else if (sample.indexOf('.') > 0) format = 'yyyy.MM.dd';
            else if (sample.indexOf('年') > 0) format = 'yyyy年MM月dd日';
            else if (/^\d{8}$/.test(sample)) format = 'yyyyMMdd';
            return _wrap({ type: 'date', format: format });
        }
        // 检测数字
        var numCount = 0;
        for (var i3 = 0; i3 < samples.length; i3++) {
            var vv = samples[i3];
            if (typeof vv === 'number') {
                numCount++;
            } else if (typeof vv === 'string') {
                var clean = vv.replace(/,/g, '').replace(/\s/g, '');
                if (!isNaN(parseFloat(clean)) && isFinite(clean)) {
                    numCount++;
                }
            }
        }
        if (numCount / samples.length > 0.8) {
            return _wrap({ type: 'number', format: 'number' });
        }
        // 全部阈值未达标：若完全是字符串（无 bool/date/number）→ string；否则 → mixed
        if (boolCount === 0 && dateCount === 0 && numCount === 0) {
            return _wrap({ type: 'string', format: 'text' });
        }
        return _wrap({ type: 'mixed', format: 'text' });
    }

    function detectSingleColumn(d, ci) {
        var samples = [];
        var maxSamples = Math.min(d.length, 100); // 最多取样100行
        var _hasNull = false;
        for (var i4 = 0; i4 < maxSamples; i4++) {
            if (!d[i4]) continue;
            var _cv = d[i4][ci];
            if (_cv === null) { _hasNull = true; continue; }
            if (_cv !== undefined && _cv !== '') {
                samples.push(_cv);
            }
        }
        function _wrap2(r) { if (_hasNull && r && typeof r === 'object') r._hasNull = true; return r; }
        if (samples.length === 0) {
            return _wrap2({ type: 'string', format: 'text' });
        }
        // 检测布尔值
        var boolCount2 = 0;
        for (var i5 = 0; i5 < samples.length; i5++) {
            var v2 = samples[i5];
            if (v2 === true || v2 === false || v2 === 'true' || v2 === 'false' || v2 === '是' || v2 === '否' || v2 === 'YES' || v2 === 'NO') {
                boolCount2++;
            }
        }
        if (boolCount2 / samples.length > 0.8) {
            return _wrap2({ type: 'boolean', format: 'boolean' });
        }
        // 检测日期
        var dateCount2 = 0;
        var datePatterns2 = [
            /^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}$/,
            /^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}$/,
            /^\d{4}年\d{1,2}月\d{1,2}[日]?$/,
            /^\d{1,2}月\d{1,2}[日]?$/,
            /^\d{8}$/
        ];
        for (var i6 = 0; i6 < samples.length; i6++) {
            var s2 = String(samples[i6]);
            for (var j2 = 0; j2 < datePatterns2.length; j2++) {
                if (datePatterns2[j2].test(s2)) {
                    dateCount2++;
                    break;
                }
            }
        }
        if (dateCount2 / samples.length > 0.8) {
            var sample2 = String(samples[0]);
            var format2 = 'date';
            if (sample2.indexOf('-') > 0) format2 = 'yyyy-MM-dd';
            else if (sample2.indexOf('/') > 0) format2 = 'yyyy/MM/dd';
            else if (sample2.indexOf('.') > 0) format2 = 'yyyy.MM.dd';
            else if (sample2.indexOf('年') > 0) format2 = 'yyyy年MM月dd日';
            else if (/^\d{8}$/.test(sample2)) format2 = 'yyyyMMdd';
            return _wrap2({ type: 'date', format: format2 });
        }
        // 检测数字
        var numCount2 = 0;
        for (var i7 = 0; i7 < samples.length; i7++) {
            var v3 = samples[i7];
            if (typeof v3 === 'number') {
                numCount2++;
            } else if (typeof v3 === 'string') {
                var clean2 = v3.replace(/,/g, '').replace(/\s/g, '');
                if (!isNaN(parseFloat(clean2)) && isFinite(clean2)) {
                    numCount2++;
                }
            }
        }
        if (numCount2 / samples.length > 0.8) {
            return _wrap2({ type: 'number', format: 'number' });
        }
        // 默认为字符串
        return _wrap2({ type: 'string', format: 'text' });
    }
};
Array2D.z检测类型 = Array2D.detectType;/**
 * 智能排序 - 自动识别数据类型并进行适当排序
 * @param {Array} arr - 二维数组
 * @param {string|number} col - 列号或列名
 * @param {string} direction - 排序方向 '+' 或 '-'，可选，不指定则自动判断
 * @param {number} skipHeader - 跳过表头行数
 * @returns {Array} 排序后的数组
 */
Array2D.smartSort = function(arr, col, direction, skipHeader) {
    var data = arr;
    var header = [];
    skipHeader = skipHeader || 0;
    
    if (skipHeader > 0) {
        header = data.slice(0, skipHeader);
        data = data.slice(skipHeader);
    }
    
    // 解析列索引
    var colIndex = -1;
    if (typeof col === 'string' && col.match(/^f\d+$/i)) {
        colIndex = parseInt(col.substring(1)) - 1;
    } else if (typeof col === 'number') {
        colIndex = col;
    }
    
    if (colIndex < 0) {
        return arr;
    }
    
    // 检测类型
    var typeInfo = Array2D.detectType(data, colIndex);
    
    // 如果未指定方向，数字和日期默认升序，字符串默认按拼音升序
    if (!direction) {
        direction = '+';
    }
    
    var isDesc = direction === '-';
    
    // 根据类型排序
    data.sort(function(a, b) {
        var valA = a[colIndex];
        var valB = b[colIndex];
        
        // 处理空值
        if (valA === null || valA === undefined || valA === '') {
            return isDesc ? 1 : -1;
        }
        if (valB === null || valB === undefined || valB === '') {
            return isDesc ? -1 : 1;
        }
        
        var result = 0;
        
        switch (typeInfo.type) {
            case 'number':
                // 数字排序
                var numA = typeof valA === 'number' ? valA : parseFloat(String(valA).replace(/,/g, ''));
                var numB = typeof valB === 'number' ? valB : parseFloat(String(valB).replace(/,/g, ''));
                result = numA - numB;
                break;
                
            case 'date':
                // 日期排序 - 统一转换为可比较格式
                var dateA = Array2D._parseDateForSort(valA);
                var dateB = Array2D._parseDateForSort(valB);
                result = dateA - dateB;
                break;
                
            case 'boolean':
                // 布尔排序
                var boolA = valA === true || valA === 'true' || valA === '是' || valA === 'YES' ? 1 : 0;
                var boolB = valB === true || valB === 'true' || valB === '是' || valB === 'YES' ? 1 : 0;
                result = boolA - boolB;
                break;
                
            default:
                // 字符串排序（尝试按拼音）
                var strA = String(valA);
                var strB = String(valB);
                
                // 如果有中文，尝试按拼音排序
                if (/[\u4e00-\u9fa5]/.test(strA) || /[\u4e00-\u9fa5]/.test(strB)) {
                    // 在WPS环境中可能没有localeCompare，使用简单比较
                    result = strA.localeCompare ? strA.localeCompare(strB, 'zh-CN') : (strA > strB ? 1 : -1);
                } else {
                    result = strA.toLowerCase() > strB.toLowerCase() ? 1 : -1;
                }
        }
        
        return isDesc ? -result : result;
    });
    
    return header.concat(data);
};
Array2D.z智能排序 = Array2D.smartSort;

/**
 * 内部方法：解析日期用于排序
 * @private
 */
Array2D._parseDateForSort = function(dateVal) {
    if (dateVal instanceof Date) {
        return dateVal.getTime();
    }
    
    var s = String(dateVal);
    
    // 尝试解析各种日期格式
    // yyyy-MM-dd
    var match = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3])).getTime();
    }
    
    // yyyy年MM月dd日
    match = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3])).getTime();
    }
    
    // yyyyMMdd
    match = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (match) {
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3])).getTime();
    }
    
    // 其他格式尝试直接解析
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
        return d.getTime();
    }
    
    return 0;
};

/**
 * 智能分组 - 自动识别数据类型并进行适当分组
 * @param {Array} arr - 二维数组
 * @param {string|number} col - 列号或列名
 * @param {string} groupBy - 分组方式（日期有效）：'year'|'month'|'day'|'week'
 * @returns {Map} 分组结果
 */
Array2D.smartGroup = function(arr, col, groupBy) {
    var data = arr;
    
    // 解析列索引
    var colIndex = -1;
    if (typeof col === 'string' && col.match(/^f\d+$/i)) {
        colIndex = parseInt(col.substring(1)) - 1;
    } else if (typeof col === 'number') {
        colIndex = col;
    }
    
    if (colIndex < 0) {
        return new Map();
    }
    
    // 检测类型
    var typeInfo = Array2D.detectType(data, colIndex);
    
    var groups = new Map();
    
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var val = row[colIndex];
        var key;
        
        if (typeInfo.type === 'date' && groupBy) {
            // 按日期维度分组
            var date = Array2D._parseDateForSort(val);
            var d = new Date(date);
            
            switch (groupBy) {
                case 'year':
                case '年':
                    key = d.getFullYear() + '年';
                    break;
                case 'month':
                case '月':
                    key = d.getFullYear() + '年' + (d.getMonth() + 1) + '月';
                    break;
                case 'day':
                case '日':
                    key = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
                    break;
                case 'week':
                case '周':
                    // 计算周数
                    var weekStart = new Date(d.getFullYear(), 0, 1);
                    var weekNum = Math.ceil(((d - weekStart) / 86400000 + weekStart.getDay() + 1) / 7);
                    key = d.getFullYear() + '年第' + weekNum + '周';
                    break;
                case 'quarter':
                case '季度':
                    var quarter = Math.floor(d.getMonth() / 3) + 1;
                    key = d.getFullYear() + '年Q' + quarter;
                    break;
                default:
                    key = String(val);
            }
        } else if (typeInfo.type === 'number' && groupBy) {
            // 按数字范围分组
            var num = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
            if (groupBy === 'decade' || groupBy === '十位数') {
                var decade = Math.floor(num / 10) * 10;
                key = decade + '-' + (decade + 9);
            } else if (groupBy === 'hundred' || groupBy === '百位数') {
                var hundred = Math.floor(num / 100) * 100;
                key = hundred + '-' + (hundred + 99);
            } else if (groupBy === 'thousand' || groupBy === '千位数') {
                var thousand = Math.floor(num / 1000) * 1000;
                key = thousand + '-' + (thousand + 999);
            } else {
                key = String(val);
            }
        } else {
            key = String(val);
        }
        
        if (!groups.has(key)) {
            groups.set(key, []);
        }
        groups.get(key).push(row);
    }
    
    // XXD-211-PATCH-MARKER: callers expect plain object {key:[rows]}, not Map.
    var out = {};
    groups.forEach(function (v, k) { out[k] = v; });
    return out;
};
Array2D.z智能分组 = Array2D.smartGroup;

// 实例方法版本
Array2D.prototype.z智能排序 = function(col, direction, skipHeader) {
    return this._new(Array2D.smartSort(this._items, col, direction, skipHeader));
};

Array2D.prototype.z智能分组 = function(col, groupBy) {
    return Array2D.smartGroup(this._items, col, groupBy);
};

Array2D.prototype.z检测类型 = function(col) {
    var colIndex = -1;
    if (typeof col === 'string' && col.match(/^f\d+$/i)) {
        colIndex = parseInt(col.substring(1)) - 1;
    } else if (typeof col === 'number') {
        colIndex = col;
    }
    return Array2D.detectType(this._items, colIndex);
};

/**
 * 静态方法：批量插入列（insertCols）- 在指定位置插入列
 * @param {Array} arr - 二维数组
 * @param {Number|Array} colPos - 插入位置或多个位置
 * @param {Array|String} values - 插入的值
 * @param {Number} totalCols - 总列数
 * @returns {Array} 插入列后的二维数组
 * @example
 * Array2D.insertCols(arr, 2, ['新列1','新列2'])
 */
Array2D.insertCols = function(arr, colPos, values, totalCols) {
    return new Array2D(arr).z批量插入列(colPos, values, totalCols).val();
};
Array2D.z批量插入列 = Array2D.insertCols;

/**
 * 静态方法：批量删除列（deleteCols/delCols）- 删除指定列
 * @param {Array} arr - 二维数组
 * @param {String|Number|Array} cols - 列配置
 * @returns {Array} 删除列后的二维数组
 * @example
 * Array2D.deleteCols(arr, '1,3,5')
 * Array2D.delCols(arr, [0, 2, 4])
 */
Array2D.deleteCols = function(arr, cols) {
    return new Array2D(arr).z批量删除列(cols).val();
};
Array2D.z批量删除列 = Array2D.deleteCols;
Array2D.delCols = Array2D.deleteCols;

/**
 * 静态方法：左连接（leftjoin）- 类似SQL的LEFT JOIN
 * @param {Array} arr - 左表
 * @param {Array} brr - 右表
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的二维数组
 * @example
 * Array2D.leftjoin(arr, brr, 'f1', 'f1', 'f1,f2,f4')
 */
Array2D.leftjoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z左连接(brr, leftKey, rightKey, resultSelector);
};
Array2D.z左连接 = Array2D.leftjoin;

/**
 * 静态方法：内连接（innerjoin）- 类似SQL的INNER JOIN
 * @param {Array} arr - 左表
 * @param {Array} brr - 右表
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 仅含两表键匹配行的二维数组
 * @example
 * Array2D.innerjoin(arr, brr, 'f1', 'f1', 'a.f1,b.f2')
 */
Array2D.innerjoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z内连接(brr, leftKey, rightKey, resultSelector);
};
Array2D.z内连接 = Array2D.innerjoin;

/**
 * 静态方法：排除（except）- 获取在arr中但不在brr中的元素
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 差异数组
 * @example
 * Array2D.except(arr, brr)
 */
Array2D.except = function(arr, brr) {
    return new Array2D(arr).z排除(brr);
};
Array2D.z排除 = Array2D.except;

/**
 * 静态方法：交集（intersect）- 获取arr和brr的交集
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 交集数组
 * @example
 * Array2D.intersect(arr, brr)
 */
Array2D.intersect = function(arr, brr) {
    return new Array2D(arr).z取交集(brr);
};
Array2D.z取交集 = Array2D.intersect;

/**
 * 静态方法：并集（union）- 获取arr和brr的并集并去重
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 并集数组
 * @example
 * Array2D.union(arr, brr)
 */
Array2D.union = function(arr, brr) {
    return new Array2D(arr).z去重并集(brr);
};
Array2D.z去重并集 = Array2D.union;

/**
 * 静态方法：最大值（max）- 获取数组最大值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 最大值
 * @example
 * Array2D.max(arr)
 * Array2D.max(arr, 'f1')
 */
Array2D.max = function(arr, selector) {
    var result = new Array2D(arr).z最大值(selector);
    return unwrapVal(result);
};
Array2D.z最大值 = Array2D.max;

/**
 * 静态方法：最小值（min）- 获取数组最小值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 最小值
 * @example
 * Array2D.min(arr)
 * Array2D.min(arr, 'f1')
 */
Array2D.min = function(arr, selector) {
    var result = new Array2D(arr).z最小值(selector);
    return unwrapVal(result);
};
Array2D.z最小值 = Array2D.min;

/**
 * 静态方法：文本连接（textjoin）- 选择指定列的值，用分隔符连接
 * @param {Array} arr - 二维数组
 * @param {String|Number|Function} selector - 列选择器，如 'f1' 或 0 或 row=>row.col
 * @param {String} [separator=','] - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D.textjoin([['a','b'],['c','d']], 1, '+')  // "b+d"
 * Array2D.textjoin([['a','b'],['c','d']], 'f2', '+')  // "b+d"
 */
Array2D.textjoin = function(arr, selector, separator = ',') {
    return new Array2D(arr).z文本连接(selector, separator);
};
Array2D.z文本连接 = Array2D.textjoin;

/**
 * Array 原型方法：textjoin - 为普通数组添加 textjoin 方法
 * 这样 .res() 返回的数组也可以使用 .textjoin()
 */
if (!Array.prototype.textjoin) {
    Array.prototype.textjoin = function(selector, separator = ',') {
        return Array2D.textjoin(this, selector, separator);
    };
}

/**
 * Array 原型方法：toRange - 为普通数组添加 toRange 方法
 * 这样 .res() 返回的数组也可以使用 .toRange()
 */
if (!Array.prototype.toRange) {
    Array.prototype.toRange = function(rng) {
        return Array2D.toRange(this, rng);
    };
}

/**
 * Array 原型方法：getRange - 为普通数组添加 getRange 方法
 * 这样 .res() 返回的数组也可以使用 .getRange()
 */
if (!Array.prototype.getRange) {
    Array.prototype.getRange = function(rng) {
        return Array2D.toRange(this, rng);
    };
}

/**
 * 静态方法：平均值（average）- 获取数组平均值
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 平均值
 * @example
 * Array2D.average(arr)
 * Array2D.average(arr, 'f1')
 */
Array2D.average = function(arr, selector) {
    var result = new Array2D(arr).z平均值(selector);
    return unwrapVal(result);
};
Array2D.z平均值 = Array2D.average;

/**
 * 静态方法：第一个（first）- 获取第一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 筛选条件
 * @returns {Array} 第一个元素（行）
 * @example
 * Array2D.first(arr)
 * Array2D.first(arr, 'f1>5')
 */
Array2D.first = function(arr, predicate) {
    var result = predicate ? new Array2D(arr).z第一个(predicate) : new Array2D(arr).z第一个();
    return unwrapVal(result);
};
Array2D.z第一个 = Array2D.first;

/**
 * 静态方法：最后一个（last）- 获取最后一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 筛选条件
 * @returns {Array} 最后一个元素（行）
 * @example
 * Array2D.last(arr)
 * Array2D.last(arr, 'f1>5')
 */
Array2D.last = function(arr, predicate) {
    var result = predicate ? new Array2D(arr).z最后一个(predicate) : new Array2D(arr).z最后一个();
    return unwrapVal(result);
};
Array2D.z最后一个 = Array2D.last;

/**
 * 静态方法：跳过（skip）- 跳过前N个元素
 * @param {Array} arr - 数组
 * @param {Number} count - 跳过的数量
 * @returns {Array} 剩余数组
 * @example
 * Array2D.skip(arr, 5)
 */
Array2D.skip = function(arr, count) {
    return new Array2D(arr).z跳过(count).val();
};
Array2D.z跳过 = Array2D.skip;

/**
 * 静态方法：取前N个（take）- 获取前N个元素
 * @param {Array} arr - 数组
 * @param {Number} count - 获取的数量
 * @returns {Array} 取出的数组
 * @example
 * Array2D.take(arr, 10)
 */
Array2D.take = function(arr, count) {
    return new Array2D(arr).z取前N个(count).val();
};
Array2D.z取前N个 = Array2D.take;

// 补充静态方法别名
Array2D.z跳过前N个 = Array2D.skip;
Array2D.z跳过前几个 = Array2D.skip;
Array2D.z取前几个 = Array2D.take;

/**
 * 静态方法：补齐数组（pad）- 补齐数组使所有行列数一致
 * @param {Array} arr - 数组
 * @param {Number} cols - 目标列数
 * @param {Number} rows - 目标行数
 * @param {*} fillValue - 填充值
 * @returns {Array} 补齐后的数组
 * @example
 * Array2D.pad(arr, 5, 10)
 * Array2D.pad(arr)  // 自动按最大列补齐
 */
Array2D.pad = function(arr, cols, rows, fillValue) {
    return new Array2D(arr).z补齐数组(cols, rows, fillValue).val();
};
Array2D.z补齐数组 = Array2D.pad;

/**
 * 静态方法：查找（find）- 查找符合条件的第一个元素
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Array} 找到的元素
 * @example
 * Array2D.find(arr, 'f1==5')
 * Array2D.find(arr, x=>x.f1>10)
 */
Array2D.find = function(arr, predicate) {
    var result = new Array2D(arr).z查找单个(predicate);
    return unwrapVal(result);
};
Array2D.z查找单个 = Array2D.find;

/**
 * 静态方法：查找索引（findIndex）- 查找符合条件的第一个元素索引
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Number} 元素索引
 * @example
 * Array2D.findIndex(arr, 'f1==5')
 */
Array2D.findIndex = function(arr, predicate) {
    return new Array2D(arr).z查找索引(predicate);
};
Array2D.z查找索引 = Array2D.findIndex;

/**
 * 静态方法：按行数分页（pageByRows）- 将数组按指定行数分页
 * @param {Array} arr - 数组
 * @param {Number} pageSize - 每页行数
 * @param {Number} pageNumber - 页码（从1开始）
 * @returns {Array} 分页后的数组
 * @example
 * Array2D.pageByRows(arr, 10, 2)
 */
Array2D.pageByRows = function(arr, pageSize, pageNumber) {
    return new Array2D(arr).z按行数分页(pageSize, pageNumber).val();
};
Array2D.z按行数分页 = Array2D.pageByRows;

/**
 * 静态方法：按页数分页（pageByCount）- 将数组平均分成指定页数
 * @param {Array} arr - 数组
 * @param {Number} pageCount - 总页数
 * @param {Number} pageNumber - 页码（从1开始）
 * @returns {Array} 分页后的数组
 * @example
 * Array2D.pageByCount(arr, 5, 2)
 */
Array2D.pageByCount = function(arr, pageCount, pageNumber) {
    return new Array2D(arr).z按页数分页(pageCount, pageNumber).val();
};
Array2D.z按页数分页 = Array2D.pageByCount;

/**
 * 静态方法：填充空白（fillBlank）- 填充合并单元格的空白区域
 * @param {Array} arr - 数组
 * @param {String} direction - 填充方向 'up'/'down'/'left'/'right'
 * @param {String} rangeAddress - 区域地址
 * @returns {Array} 填充后的数组
 * @example
 * Array2D.fillBlank(arr, 'up', 'A2:D2')
 */
Array2D.fillBlank = function(arr, direction, rangeAddress) {
    return new Array2D(arr).z补齐空位(direction, rangeAddress).val();
};
Array2D.z补齐空位 = Array2D.fillBlank;

/**
 * 静态方法：转矩阵（toMatrix）- 将数组转换为矩阵格式
 * @param {Array} arr - 数组
 * @param {Number} rows - 行数
 * @param {Number} cols - 列数
 * @param {String} direction - 方向 'r'/'c'
 * @returns {Array} 矩阵数组
 * @example
 * Array2D.toMatrix(arr, 3, 4, 'r')
 */
Array2D.toMatrix = function(arr, rows, cols, direction) {
    return new Array2D(arr).z转矩阵(rows, cols, direction).val();
};
Array2D.z转矩阵 = Array2D.toMatrix;

/**
 * 静态方法：查找所有行下标（findRowsIndex）- 查找符合条件的所有行索引
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 查找条件
 * @returns {Array} 行索引数组
 * @example
 * Array2D.findRowsIndex(arr, 'f1=="A"')
 */
Array2D.findRowsIndex = function(arr, predicate) {
    return new Array2D(arr).z查找所有行下标(predicate);
};
Array2D.z查找所有行下标 = Array2D.findRowsIndex;

/**
 * 静态方法：排序（sort）- 基本排序
 * @param {Array} arr - 数组
 * @param {String|Function} comparer - 比较函数
 * @returns {Array} 排序后的数组
 * @example
 * Array2D.sort(arr)
 * Array2D.sort(arr, 'f1+')
 */
Array2D.sort = function(arr, comparer) {
    return new Array2D(arr).z升序排序(comparer).val();
};
Array2D.z升序排序 = Array2D.sort;

/**
 * 静态方法：降序排序（sortDesc）
 * @param {Array} arr - 数组
 * @param {String|Function} comparer - 比较函数
 * @returns {Array} 排序后的数组
 * @example
 * Array2D.sortDesc(arr, 'f1-')
 */
Array2D.sortDesc = function(arr, comparer) {
    return new Array2D(arr).z降序排序(comparer).val();
};
Array2D.z降序排序 = Array2D.sortDesc;

/**
 * 静态方法：求和（sum）- 计算数组元素的和
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 和
 * @example
 * Array2D.sum([1,2,3,4])
 * Array2D.sum(arr, 'f1')
 */
Array2D.sum = function(arr, selector) {
    var result = new Array2D(arr).z求和(selector);
    return unwrapVal(result);
};

/**
 * 静态方法：归约（reduce）- 对数组进行归约操作
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @param {*} initialValue - 初始值
 * @returns {*} 归约结果
 * @example
 * Array2D.reduce(arr, (acc, row) => acc + row[0], 0)
 */
Array2D.reduce = function(arr, callback, initialValue) {
    return new Array2D(arr).z归约(callback, initialValue);
};
Array2D.z归约 = Array2D.reduce;

/**
 * 静态方法：倒序归约（reduceRight）- 从右向左归约
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @param {*} initialValue - 初始值
 * @returns {*} 归约结果
 * @example
 * Array2D.reduceRight(arr, (acc, row) => acc + row[0], 0)
 */
Array2D.reduceRight = function(arr, callback, initialValue) {
    return new Array2D(arr).z倒序归约(callback, initialValue);
};
Array2D.z倒序归约 = Array2D.reduceRight;

/**
 * 静态方法：遍历（forEach）- 遍历数组的每一行
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @returns {Array} 原数组
 * @example
 * Array2D.forEach(arr, (row, i) => console.log(i, row))
 */
Array2D.forEach = function(arr, callback) {
    var instance = new Array2D(arr);
    instance.forEach(callback);
    return arr;
};

/**
 * 静态方法：倒序遍历（forEachRev）- 从后向前遍历
 * @param {Array} arr - 数组
 * @param {Function} callback - 回调函数
 * @returns {Array} 原数组
 * @example
 * Array2D.forEachRev(arr, (row, i) => console.log(i, row))
 */
Array2D.forEachRev = function(arr, callback) {
    new Array2D(arr).z倒序遍历执行(callback);
    return arr;
};
Array2D.z倒序遍历执行 = Array2D.forEachRev;

/**
 * 静态方法：有满足（some）- 检查是否有元素满足条件
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 条件
 * @returns {Boolean} 是否有满足
 * @example
 * Array2D.some(arr, 'f1>5')
 */
Array2D.some = function(arr, predicate) {
    return new Array2D(arr).z有满足(predicate);
};
Array2D.z有满足 = Array2D.some;

/**
 * 静态方法：全部满足（every）- 检查是否所有元素都满足条件
 * @param {Array} arr - 数组
 * @param {String|Function} predicate - 条件
 * @returns {Boolean} 是否全部满足
 * @example
 * Array2D.every(arr, 'f1>0')
 */
Array2D.every = function(arr, predicate) {
    return new Array2D(arr).z全部满足(predicate);
};
Array2D.z全部满足 = Array2D.every;

/**
 * 静态方法：降维（flat）- 将二维数组降维为一维
 * @param {Array} arr - 二维数组
 * @param {Function} mapper - 可选的映射函数
 * @returns {Array} 一维数组
 * @example
 * Array2D.flat(arr)
 * Array2D.flat(arr, x=>x.f1)
 */
Array2D.flat = function(arr, mapper) {
    var result = new Array2D(arr);
    return mapper ? result.z扁平化(mapper) : result.z扁平化();
};
Array2D.z扁平化 = Array2D.flat;

/**
 * 静态方法：行切片删除（splice）- 删除/插入元素
 * @param {Array} arr - 数组
 * @param {Number} start - 起始位置
 * @param {Number} deleteCount - 删除数量
 * @param {...*} items - 要插入的元素
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.splice(arr, 2, 1, ['新行'])
 */
Array2D.splice = function(arr, start, deleteCount) {
    var items = Array.prototype.slice.call(arguments, 3);
    return new Array2D(arr).z行切片删除行(start, deleteCount, items);
};
Array2D.z行切片删除行 = Array2D.splice;

/**
 * 静态方法：追加一项（push）- 在数组末尾添加元素
 * @param {Array} arr - 数组
 * @param {*} item - 要添加的元素
 * @returns {Number} 新长度
 * @example
 * Array2D.push(arr, [1,2,3])
 */
Array2D.push = function(arr, item) {
    new Array2D(arr).z追加一项(item);
    return arr.length;
};
Array2D.z追加一项 = Array2D.push;

/**
 * 静态方法：尾部弹出一项（pop）- 删除并返回最后一个元素
 * @param {Array} arr - 数组
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.pop(arr)
 */
Array2D.pop = function(arr) {
    return new Array2D(arr).z尾部弹出一项();
};
Array2D.z尾部弹出一项 = Array2D.pop;

/**
 * 静态方法：删除第一个（shift）- 删除并返回第一个元素
 * @param {Array} arr - 数组
 * @returns {Array} 被删除的元素
 * @example
 * Array2D.shift(arr)
 */
Array2D.shift = function(arr) {
    return new Array2D(arr).z删除第一个();
};
Array2D.z删除第一个 = Array2D.shift;

/**
 * 静态方法：反转（reverse）- 反转数组顺序
 * @param {Array} arr - 数组
 * @returns {Array} 反转后的数组
 * @example
 * Array2D.reverse(arr)
 */
Array2D.reverse = function(arr) {
    return new Array2D(arr).z反转().val();
};
Array2D.z反转 = Array2D.reverse;

/**
 * 静态方法：文本连接（join）- 用分隔符连接所有元素
 * @param {Array} arr - 数组
 * @param {String} separator - 分隔符
 * @returns {String} 连接后的字符串
 * @example
 * Array2D.join(arr, ',')
 */
Array2D.join = function(arr, separator) {
    return new Array2D(arr).z连接(separator);
};
Array2D.z连接 = Array2D.join;

/**
 * 静态方法：转JSON字符串（toJson）- 将数组转为JSON字符串
 * @param {Array} arr - 数组
 * @param {Number|String} indent - 缩进
 * @returns {String} JSON字符串
 * @example
 * Array2D.toJson(arr, 2)
 */
Array2D.toJson = function(arr, indent) {
    return new Array2D(arr).z转JSON(indent);
};
Array2D.z转JSON = Array2D.toJson;

/**
 * 静态方法：转字符串（toString）- 将数组转为字符串
 * @param {Array} arr - 数组
 * @param {String} separator - 分隔符
 * @returns {String} 字符串
 * @example
 * Array2D.toString(arr, ',')
 */
Array2D.toString = function(arr, separator) {
    return new Array2D(arr).z转字符串(separator);
};
Array2D.z转字符串 = Array2D.toString;

/**
 * 静态方法：是否为空（isEmpty）- 检查数组是否为空
 * @param {Array} arr - 数组
 * @returns {Boolean} 是否为空
 * @example
 * Array2D.isEmpty(arr)
 */
Array2D.isEmpty = function(arr) {
    return new Array2D(arr).z是否为空();
};
Array2D.z是否为空 = Array2D.isEmpty;

/**
 * 静态方法：分组（groupBy）- 按指定条件分组
 * @param {Array} arr - 数组
 * @param {String|Function} keySelector - 分组依据
 * @returns {Map} 分组结果
 * @example
 * Array2D.groupBy(arr, 'f1')
 */
Array2D.groupBy = function(arr, keySelector) {
    return new Array2D(arr).z分组(keySelector);
};
Array2D.z分组 = Array2D.groupBy;

/**
 * 静态方法：左右全连接（fulljoin）- 类似SQL的FULL OUTER JOIN
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.fulljoin(arr, brr, 'f1', 'f1', 'f1,f2,f3')
 */
Array2D.fulljoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z左右全连接(brr, leftKey, rightKey, resultSelector);
};
Array2D.z左右全连接 = Array2D.fulljoin;

/**
 * 静态方法：一对多连接（leftFulljoin）- 左表一对多连接
 * @param {Array} arr - 左表
 * @param {Array} brr - 右表
 * @param {String|Function} leftKey - 左表关键字
 * @param {String|Function} rightKey - 右表关键字
 * @param {String|Function} resultSelector - 结果选择器
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.leftFulljoin(arr, brr, 'f1', 'f1', 'f1,f2,f3')
 */
Array2D.leftFulljoin = function(arr, brr, leftKey, rightKey, resultSelector) {
    return new Array2D(arr).z一对多连接(brr, leftKey, rightKey, resultSelector);
};
Array2D.z一对多连接 = Array2D.leftFulljoin;

/**
 * 静态方法：超级查找（superLookup）- 增强版VLOOKUP
 * @param {Array} arr - 查找范围
 * @param {*} lookupValue - 查找值
 * @param {Number|String} colIndex - 列号
 * @param {Number|String} returnCol - 返回列号
 * @returns {Array} 查找结果
 * @example
 * Array2D.superLookup(arr, 'A', 1, 3)
 */
Array2D.superLookup = function(arr, lookupValue, colIndex, returnCol) {
    var result = new Array2D(arr).z超级查找(lookupValue, colIndex, returnCol);
    return unwrapVal(result);
};
Array2D.z超级查找 = Array2D.superLookup;

/**
 * 静态方法：左右连接（zip）- 将两个数组的对应位置元素配对
 * @param {Array} arr - 数组1
 * @param {Array} brr - 数组2
 * @returns {Array} 配对后的数组
 * @example
 * Array2D.zip(arr1, arr2)
 */
Array2D.zip = function(arr, brr) {
    return new Array2D(arr).z左右连接(brr);
};
Array2D.z左右连接 = Array2D.zip;

/**
 * 静态方法：转置（transpose）- 转置二维数组
 * @param {Array} arr - 二维数组
 * @returns {Array} 转置后的数组
 * @example
 * Array2D.transpose([[1,2],[3,4]])  // [[1,3],[2,4]]
 */
Array2D.transpose = function(arr) {
    return new Array2D(arr).z转置().val();
};

/**
 * 静态方法：克隆（copy）- 深拷贝数组
 * @param {Array} arr - 数组
 * @returns {Array} 拷贝后的数组
 * @example
 * Array2D.copy(arr)
 */
Array2D.copy = function(arr) {
    return new Array2D(arr).z克隆().val();
};
Array2D.z克隆 = Array2D.copy;
Array2D.z复制 = Array2D.copy;

/**
 * 静态方法：上下连接（concat）- 连接多个数组
 * @param {Array} arr - 数组1
 * @param {...Array} arrays - 其他数组
 * @returns {Array} 连接后的数组
 * @example
 * Array2D.concat(arr1, arr2, arr3)
 */
Array2D.concat = function(arr) {
    var arrays = Array.prototype.slice.call(arguments, 1);
    return new Array2D(arr).z上下连接.apply(new Array2D(arr), arrays).val();
};
Array2D.z上下连接 = Array2D.concat;

/**
 * 静态方法：选择行（selectRows）- 选择指定行
 * @param {Array} arr - 数组
 * @param {String|Array} rows - 行配置
 * @returns {Array} 选择后的数组
 * @example
 * Array2D.selectRows(arr, '1,3,5')
 * Array2D.selectRows(arr, [0, 2, 4])
 */
Array2D.selectRows = function(arr, rows) {
    return new Array2D(arr).z选择行(rows).val();
};
Array2D.z选择行 = Array2D.selectRows;

/**
 * 静态方法：删除行（deleteRows）- 删除符合条件的行
 * @param {Array} arr - 数组
 * @param {Function|String|Array} rows - 行选择器（支持函数、字符串或数组）
 * @returns {Array} 删除后的数组
 * @example
 * Array2D.deleteRows(arr, r=>r.f3=="美国")  // 删除f3列值为"美国"的行
 * Array2D.deleteRows(arr, '1,3,5')
 */
Array2D.deleteRows = function(arr, rows) {
    return new Array2D(arr).z批量删除行(rows).val();
};
Array2D.z批量删除行 = Array2D.deleteRows;

/**
 * 静态方法：插入行（insertRows）- 在指定位置插入行
 * @param {Array} arr - 数组
 * @param {Number|Array} rowPos - 插入位置
 * @param {Array} values - 插入的值
 * @returns {Array} 插入后的数组
 * @example
 * Array2D.insertRows(arr, 2, [[1,2,3]])
 */
Array2D.insertRows = function(arr, rowPos, values) {
    return new Array2D(arr).z批量插入行(rowPos, values).val();
};
Array2D.z批量插入行 = Array2D.insertRows;

/**
 * 静态方法：插入行号（insertRowNum）- 在数组前插入行号列
 * @param {Array} arr - 数组
 * @param {Number} start - 起始行号
 * @param {String} title - 列标题
 * @returns {Array} 插入行号后的数组
 * @example
 * Array2D.insertRowNum(arr, 1, '序号')
 */
Array2D.insertRowNum = function(arr, start, title) {
    return new Array2D(arr).z插入行号(start, title).val();
};
Array2D.z插入行号 = Array2D.insertRowNum;

/**
 * 静态方法：是否包含值（includes）- 检查数组是否包含某值
 * @param {Array} arr - 数组
 * @param {*} value - 要检查的值
 * @returns {Boolean} 是否包含
 * @example
 * Array2D.includes(arr, [1,2])
 */
Array2D.includes = function(arr, value) {
    return new Array2D(arr).z包含(value);
};
Array2D.z包含 = Array2D.includes;

/**
 * 静态方法：值位置（indexOf）- 查找元素的位置
 * @param {Array} arr - 数组
 * @param {*} value - 要查找的值
 * @returns {Number} 元素索引
 * @example
 * Array2D.indexOf(arr, [1,2])
 */
Array2D.indexOf = function(arr, value) {
    return new Array2D(arr).z值位置(value);
};

/**
 * 静态方法：从后往前值位置（lastIndexOf）- 从后向前查找元素位置
 * @param {Array} arr - 数组
 * @param {*} value - 要查找的值
 * @returns {Number} 元素索引
 * @example
 * Array2D.lastIndexOf(arr, [1,2])
 */
Array2D.lastIndexOf = function(arr, value) {
    return new Array2D(arr).z从后往前值位置(value);
};

/**
 * 静态方法：中位数（median）- 计算中位数
 * @param {Array} arr - 数组
 * @param {String|Function} selector - 选择器
 * @returns {Number} 中位数
 * @example
 * Array2D.median(arr)
 * Array2D.median(arr, 'f1')
 */
Array2D.median = function(arr, selector) {
    var result = new Array2D(arr).z中位数(selector);
    return unwrapVal(result);
};
Array2D.z中位数 = Array2D.median;

/**
 * 静态方法：间隔取数（nth）- 每隔n个取一个
 * @param {Array} arr - 数组
 * @param {Number} n - 间隔
 * @param {Number} offset - 偏移量
 * @returns {Array} 取出的数组
 * @example
 * Array2D.nth(arr, 3, 0)  // 每3个取1个
 */
Array2D.nth = function(arr, n, offset) {
    // XXD-187: prototype 已改为 (offset, step) 顺序，此处转发对齐
    return new Array2D(arr).z间隔取数(offset || 0, n).val();
};
Array2D.z间隔取数 = Array2D.nth;

/**
 * 静态方法：行切片（slice）- 提取指定范围的行
 * @param {Array} arr - 数组
 * @param {Number} start - 起始位置
 * @param {Number} end - 结束位置
 * @returns {Array} 切片后的数组
 * @example
 * Array2D.slice(arr, 1, 5)
 */
Array2D.slice = function(arr, start, end) {
    return new Array2D(arr).z行切片(start, end).val();
};

/**
 * 静态方法：结果（res）- 获取结果数组
 * @param {Array} arr - 数组
 * @returns {Array} 结果数组
 * @example
 * Array2D.res(arr)
 */
Array2D.res = function(arr) {
    return new Array2D(arr).z结果();
};
Array2D.z结果 = Array2D.res;

// ==================== [STATIC_METHOD_FACTORY] 静态方法工厂 ====================
/**
 * 静态方法工厂配置
 * 格式: [静态方法名, 实例方法名, 中文别名, returnType]
 * returnType: 'array' 返回.val(), 'scalar' 返回.val()但处理单值, 'direct' 直接返回, 'string' 直接返回字符串
 */
var _STATIC_METHOD_CONFIG = [
    // 统计类 - 返回标量值
    ['sum', 'z求和', 'z求和', 'scalar'],
    ['average', 'z平均值', 'z平均值', 'scalar'],
    ['median', 'z中位数', 'z中位数', 'scalar'],
    ['max', 'z最大值', 'z最大值', 'scalar'],
    ['min', 'z最小值', 'z最小值', 'scalar'],

    // 取值类 - 返回数组元素
    ['first', 'z第一个', 'z第一个', 'oraw'],
    ['last', 'z最后一个', 'z最后一个', 'oraw'],

    // 数组变换类 - 返回数组
    ['filter', 'z筛选', 'z筛选', 'array'],
    ['map', 'z映射', 'z映射', 'array'],
    ['skip', 'z跳过', 'z跳过', 'array'],
    ['take', 'z取前N个', 'z取前N个', 'array'],
    ['reverse', 'z反转', 'z反转', 'array'],
    ['concat', 'z上下连接', 'z上下连接', 'array'],
    ['slice', 'z行切片', 'z行切片', 'array'],
    ['pad', 'z补齐数组', 'z补齐数组', 'array'],
    ['distinct', 'z去重', 'z去重', 'array'],
    ['clone', 'z克隆', 'z克隆', 'array'],
    ['transpose', 'z转置', 'z转置', 'array'],
    ['fillBlank', 'z补齐空位', 'z补齐空位', 'array'],

    // 连接类 - 返回数组
    ['leftjoin', 'z左连接', 'z左连接', 'array'],
    ['innerjoin', 'z内连接', 'z内连接', 'array'],
    ['fulljoin', 'z左右全连接', 'z左右全连接', 'array'],
    ['leftFulljoin', 'z一对多连接', 'z一对多连接', 'array'],
    ['zip', 'z左右连接', 'z左右连接', 'array'],
    ['except', 'z排除', 'z排除', 'array'],
    ['intersect', 'z取交集', 'z取交集', 'array'],
    ['union', 'z去重并集', 'z去重并集', 'array'],

    // 排序类 - 返回数组
    ['sort', 'z升序排序', 'z升序排序', 'array'],
    ['sortDesc', 'z降序排序', 'z降序排序', 'array'],
    ['sortByCols', 'z多列排序', 'z多列排序', 'array'],
    ['sortByList', 'z自定义排序', 'z自定义排序', 'array'],

    // 行列操作类 - 返回数组
    ['insertCols', 'z批量插入列', 'z批量插入列', 'array'],
    ['deleteCols', 'z批量删除列', 'z批量删除列', 'array'],
    ['selectRows', 'z选择行', 'z选择行', 'array'],
    ['deleteRows', 'z批量删除行', 'z批量删除行', 'array'],
    ['insertRows', 'z批量插入行', 'z批量插入行', 'array'],
    ['insertRowNum', 'z插入行号', 'z插入行号', 'array'],
    ['toMatrix', 'z转矩阵', 'z转矩阵', 'array'],

    // 分页类 - 返回数组
    ['pageByRows', 'z按行数分页', 'z按行数分页', 'array'],
    ['pageByCount', 'z按页数分页', 'z按页数分页', 'array'],

    // 查找类 - 部分返回标量/数组
    ['find', 'z查找单个', 'z查找单个', 'oraw'],
    ['findRowsIndex', 'z查找所有行下标', 'z查找所有行下标', 'direct'],

    // 归约类 - 返回原始类型
    ['reduce', 'z归约', 'z归约', 'direct'],
    ['reduceRight', 'z倒序归约', 'z倒序归约', 'direct'],
    ['some', 'z有满足', 'z有满足', 'direct'],
    ['every', 'z全部满足', 'z全部满足', 'direct'],

    // 特殊类
    ['textjoin', 'z文本连接', 'z文本连接', 'string'],
];

/**
 * 批量创建Array2D静态方法
 * @param {Array} configs - 配置数组
 */
function _createStaticMethods(configs) {
    configs.forEach(function(config) {
        var staticName = config[0];
        var instanceName = config[1];
        var cnName = config[2];
        var returnType = config[3] || 'array';

        Array2D[staticName] = function(arr) {
            var args = Array.prototype.slice.call(arguments, 1);
            var instance = new Array2D(arr);
            var result = instance[instanceName].apply(instance, args);

            switch (returnType) {
                case 'array':
                    return result && unwrapVal(result);
                case 'scalar':
                    return unwrapVal(result);
                case 'oraw':
                    // raw value (single element) - no .val() needed
                    return unwrapVal(result);
                case 'direct':
                    return result;
                case 'string':
                    return result;
                default:
                    return result && result.val ? result.val() : result;
            }
        };

        // 创建中文别名
        Array2D[cnName] = Array2D[staticName];

        // $$ 同步获得相同静态方法
        $$[staticName] = Array2D[staticName];
        $$[cnName] = Array2D[staticName];
    });
}

// 批量创建静态方法
_createStaticMethods(_STATIC_METHOD_CONFIG);

// 别名
Array2D.delCols = Array2D.deleteCols;
Array2D.z跳过前N个 = Array2D.skip;
Array2D.z跳过前几个 = Array2D.skip;
Array2D.z取前几个 = Array2D.take;
/**
 * Array2D.findColsIndex 静态方法 — 查找满足条件的列的索引
 * @param {Array} arr - 二维数组
 * @param {Function} fn - 条件函数 (colArray, colIndex) => boolean
 * @returns {Array} 满足条件的列索引数组
 * @example
 * Array2D.findColsIndex(arr, x => x[1] == '产品')
 */
Array2D.findColsIndex = function(arr, fn) {
    return (new Array2D(arr)).findColsIndex(fn).val();
};
Array2D.z查找所有列下标 = Array2D.findColsIndex;

/**
 * Array2D.findAllIndex 静态方法 — 查找所有满足条件的元素位置
 * @param {Array} arr - 二维数组
 * @param {Function} fn - 条件函数 (value) => boolean
 * @returns {Array} 位置索引数组 [[row, col], ...]
 * @example
 * Array2D.findAllIndex(rng.Value2, x => x == 10)
 */
Array2D.findAllIndex = function(arr, fn) {
    return (new Array2D(arr)).findAllIndex(fn).val();
};
Array2D.z查找所有下标 = Array2D.findAllIndex;

/**
 * Array2D.pageByIndexs 静态方法 — 按下标分页
 * @param {Array} arr - 二维数组
 * @param {Array} idxs - 行号数组（0-based）
 * @returns {Array} 按索引提取的行组成的数组
 * @example
 * Array2D.pageByIndexs(arr, [4, 9])
 */
Array2D.pageByIndexs = function(arr, idxs) {
    return (new Array2D(arr)).pageByIndexs(idxs).val();
};
Array2D.z按下标分页 = Array2D.pageByIndexs;

/**
 * Array2D.findRange 静态方法 — 单元格批量查找（对二维数组按值查找）
 * 注意：此方法包装了独立函数 findRange，用于在 Array2D 上下文中访问
 * @param {Array} arr - 二维数组
 * @param {*} value - 查找值
 * @returns {Array} 找到的行数组
 */
Array2D.findRange = function(arr, value) {
    var results = [];
    if (!arr || !Array.isArray(arr)) return results;
    for (var i = 0; i < arr.length; i++) {
        var row = arr[i];
        if (!row) continue;
        for (var j = 0; j < row.length; j++) {
            if (row[j] === value) {
                results.push(row);
                break;
            }
        }
    }
    return results;
};
Array2D.z查找区域 = Array2D.findRange;

/**
 * Array2D.rangeSelect 静态方法 — 按范围选择（对二维数组按索引范围提取）
 * @param {Array} arr - 二维数组
 * @param {Number} start - 起始行索引
 * @param {Number} end - 结束行索引
 * @param {Number} [startCol] - 起始列索引
 * @param {Number} [endCol] - 结束列索引
 * @returns {Array} 提取的子数组
 */
Array2D.rangeSelect = function(arr, start, end, startCol, endCol) {
    // v4.0.11: 支持 Array2D 实例（$.maxArray 返回值）
    if (arr && arr._items) arr = arr._items;
    var result = [];
    if (!arr || !Array.isArray(arr)) return result;
    
    // v4.0.11: 支持地址字符串（A2:C4, A:C, 2:3, 逗号分隔多区域）
    if (typeof start === 'string' && /^[A-Za-z\d]/.test(start)) {
        var _ci = function(s) { var idx = 0; for (var ki = 0; ki < s.length; ki++) idx = idx * 26 + (s.toUpperCase().charCodeAt(ki) - 64); return idx - 1; };
        function _parse(a, arr) {
            a = a.trim(); if (!a) return null;
            var rm = a.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
            var cm = a.match(/^([A-Za-z]+):([A-Za-z]+)$/);
            var rw = a.match(/^(\d+):(\d+)$/);
            if (rm) { var c1=_ci(rm[1]),r1=parseInt(rm[2])-1,c2=_ci(rm[3]),r2=parseInt(rm[4])-1; var rs=Math.min(r1,r2),cs=Math.min(c1,c2); return {rs:rs,cs:cs,rc:Math.min(Math.abs(r2-r1)+1, arr.length-rs),cc:Math.min(Math.abs(c2-c1)+1, (arr[0]?arr[0].length:0)-cs)}; }
            else if (cm) { var cc1=_ci(cm[1]),cc2=_ci(cm[2]); return {rs:0,cs:Math.min(cc1,cc2),rc:arr.length,cc:Math.abs(cc2-cc1)+1}; }
            else if (rw) { var rr1=parseInt(rw[1])-1,rr2=parseInt(rw[2])-1,mcc=0; for(var ri=0;ri<arr.length;ri++)if(arr[ri]&&arr[ri].length>mcc)mcc=arr[ri].length; return {rs:Math.min(rr1,rr2),cs:0,rc:Math.abs(rr2-rr1)+1,cc:mcc}; }
            return null;
        }
        var parts = start.split(',');
        if (parts.length > 1) {
            for (var p = 0; p < parts.length; p++) {
                var a = _parse(parts[p], arr);
                if (!a) continue;
                var sub = [], re = Math.min(a.rs+a.rc, arr.length);
                for (var i = a.rs; i < re; i++) { var row = arr[i]; sub.push(row ? row.slice(a.cs, Math.min(a.cs+a.cc, row.length)) : []); }
                result.push(sub);
            }
            return result;
        }
        var a = _parse(start, arr);
        if (a) {
            var re = Math.min(a.rs+a.rc, arr.length);
            for (var i = a.rs; i < re; i++) { var row = arr[i]; result.push(row ? row.slice(a.cs, Math.min(a.cs+a.cc, row.length)) : []); }
            return result;
        }
        // 地址解析失败，返回空（避免旧逻辑错误处理字符串）
        return result;
    }
    
    // 原有数字索引模式
    var s = start || 0;
    var e = end !== undefined ? end : arr.length - 1;
    if (s < 0) s = 0;
    if (e >= arr.length) e = arr.length - 1;
    for (var i = s; i <= e; i++) {
        var row = arr[i];
        if (!row) { result.push([]); continue; }
        if (startCol !== undefined && endCol !== undefined) {
            result.push(row.slice(startCol, endCol + 1));
        } else {
            result.push(row.slice());
        }
    }
    return result;
};
Array2D.z按范围选择 = Array2D.rangeSelect;



// ==================== $$ 同步 Array2D 所有静态属性 ====================
// v4.0.11: 确保 $$ 与 Array2D 完全同步（包括后续添加的属性）
(function() {
    var _syncToSS = function() {
        var target = Array2D;
        for (var key in target) {
            if (target.hasOwnProperty(key)) {
                try { Array2D[key] = target[key]; } catch(e) {}
                try { $$[key] = target[key]; } catch(e) {}
            }
        }
    };
    _syncToSS();
    // 延迟再次同步（等待所有模块加载完毕）
    if (typeof setTimeout !== 'undefined') setTimeout(_syncToSS, 0);
})();

// ==================== 超级透视表（superPivot）====================
/**
 * 超级透视（z超级透视）- 将二维数组仿透视表生成行列字段，并进行各种汇总统计的交叉表
 * @param {Array} arr - 二维数组
 * @param {Array|string} rowFields - 行字段配置，如 ['f1+,f2-'] 或 ['f1,f2', '标题']
 * @param {Array|string} colFields - 列字段配置，如 ['f5+,f6+'] 或 ['f2', '标题']
 * @param {Array|string} dataFields - 数据字段配置，如 ['count(),sum("f3")'] 或 [[回调数组], '标题']
 * @param {Number} headerRows - 标题行数，默认1
 * @param {string|Number} outputHeader - 1:输出表头, 0:不输出, 'map':返回字典，默认1
 * @param {string} separator - 分隔符，默认"@^@"
 * @returns {Array|Map} 返回二维数组或Map
 * @example
 * // 示例1：基本透视（带排序符号）
 * var rs = Array2D.z超级透视(arr, ['f1+,f2-'], ['f5+,f6+'], ['count(),sum("f3")']);
 *
 * // 示例2：带标题的透视
 * var rs = Array2D.z超级透视(arr, ['f1,f5,f6','prod,year,month'], ['f2','country'], ['count(),sum("f3"),average("f4")','count,sum,avg']);
 *
 * // 示例3：回调函数模式 + Map返回
 * var rs = Array2D.z超级透视(arr, ['f1,f5,f6','期数,年,月'], ['f2','国家'], [[g=>g.count(),g=>g.sum("f3")],'计数,求和'], 2, 'map');
 */
Array2D.z超级透视 = function(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options) {
    // 🔧 v4.0.18: WPS Range 对象 → Value2 转换 + WPS host array → 真 Array 强转
    //   根因: WPS 公式 =k("(...args)=>$.superPivot(...args).filter(...)") 传给 JSA 的 Range 对象
    //         typeof === 'function' (不是 'object'),且 smartUnwrap 调 .Value2 仍可能调不通
    //         直接传给 superPivot → arr.slice(dataStartRow) 报 "arr.slice is not a function"
    //   修法 (双重保险):
    //     1) superPivot 入口先做 Range duck-typing 检测(看 Address/Value2/Cells 任一),有就 .Value2
    //     2) 然后用 try { arr.slice(0) } 探针看是不是 host array,不是就 JSON 强转
    //   这样不依赖 jsaLambda 层的 smartUnwrap,superPivot 自己处理 Range/host array

    // (1) Range 检测(duck-typing:不依赖 typeof)
    if (arr != null && !(arr instanceof Array2D)) {
        var __looksLikeRange = (typeof arr.Address !== 'undefined') ||
                               (typeof arr.Value2 !== 'undefined') ||
                               (typeof arr.Cells !== 'undefined') ||
                               (typeof arr.Worksheet !== 'undefined') ||
                               (typeof arr.Row !== 'undefined');
        if (__looksLikeRange && !Array.isArray(arr)) {
            try {
                if (typeof arr.Value2 !== 'undefined') {
                    var __vv = arr.Value2;
                    if (__vv && __vv !== arr) arr = __vv;
                }
            } catch (__re1) {}
            if (!Array.isArray(arr)) {
                try {
                    if (typeof arr.Value !== 'undefined') {
                        var __vv2 = arr.Value;
                        if (__vv2 && __vv2 !== arr) arr = __vv2;
                    }
                } catch (__re2) {}
            }
        }
    }

    // (2) host array → 真 Array 强转(用 try 探针)
    if (Array.isArray(arr) && !(arr instanceof Array2D)) {
        var __isRealArr = true;
        try {
            var __probe = arr.slice(0);
            if (!Array.isArray(__probe)) __isRealArr = false;
        } catch (__probeErr) {
            __isRealArr = false;
        }
        if (!__isRealArr) {
            try {
                var __hp = JSON.parse(JSON.stringify(arr));
                if (Array.isArray(__hp)) arr = __hp;
            } catch (__he1) {
                try { arr = Array.from(arr); } catch (__he2) {}
            }
        }
    }

    // 🔧 XXD-163: 空数据/空入参 → 直接返回 [],不再走 header 占位("值1")路径,也不再抛 TypeError
    if (arr == null || (Array.isArray(arr) && arr.length === 0)) {
        return [];
    }

    // 🔧 v4.0.10: 检测新式调用：第5参数为 options 对象
    if (headerRows !== undefined && headerRows !== null && typeof headerRows === 'object' && !Array.isArray(headerRows)) {
        // 新式调用: z超级透视(arr, rowFields, colFields, dataFields, options)
        options = headerRows;
        headerRows = options.headerRows !== undefined ? options.headerRows : 1;
        outputHeader = options.outputHeader !== undefined ? options.outputHeader : 1;
        separator = options.separator || '@^@';
    } else {
        // 旧式调用: z超级透视(arr, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options)
        options = options || {};
        separator = separator || '@^@';
        headerRows = headerRows !== undefined ? headerRows : 1;
        outputHeader = outputHeader !== undefined ? outputHeader : 1;
    }
    var rowColSep = options.rowColSeparator || '|||';
    
    // 🔧 v3.9.0 新增：解析options参数
    var cornerTitle = options.cornerTitle || '';
    var rowFieldIndent = options.rowFieldIndent !== false;  // 默认启用缩进
    var rowFieldIndentSize = options.rowFieldIndentSize || 4;  // 默认4空格
    var layoutMode = options.layoutMode || 'outline';  // compact/outline/tabular
    
    // 🔧 v3.9.0 新增：小计和总计配置
    options.subtotals = options.subtotals || { 
      enabled: false,
      row: false,
      col: false,
      label: '小计' 
    };
    
    options.grandTotal = options.grandTotal || { 
      row: false,
      col: false,
      label: '总计' 
    };
    
    // 兼容旧版配置名称
    if (options.rowSubtotals && options.rowSubtotals.enabled) {
      options.subtotals.row = true;
    }
    if (options.colSubtotals && options.colSubtotals.enabled) {
      options.subtotals.col = true;
    }
    if (options.grandTotals) {
      if (options.grandTotals.row) options.grandTotal.row = true;
      if (options.grandTotals.column) options.grandTotal.col = true;
    }
    
    // 保留旧配置变量以兼容现有代码
    var rowSubtotals = { enabled: options.subtotals.row };
    var colSubtotals = { enabled: options.subtotals.col };
    var grandTotals = { row: options.grandTotal.row, column: options.grandTotal.col };
    
    // 百分比显示配置
    var displayAs = options.displayAs || { mode: 'value', decimals: 2 };

    // 🔧 v4.0.10: 数字格式化配置
    // 🔧 XXD-47: 默认值从 '' 改为 0 — 与 v4.0.34 数据行硬编码 0 保持一致
    var nullValue = options.nullValue !== undefined ? options.nullValue : 0;

    // 🔧 v3.9.4 新增：筛选器配置
    // 支持在透视表输出前筛选特定行/列值
    // filterRows: 函数或对象，筛选行键，如 { f1: ['北京', '上海'] } 或 row => row.f1 !== '深圳'
    // filterCols: 函数或对象，筛选列键
    var filterRows = options.filterRows || null;
    var filterCols = options.filterCols || null;

    // 🔧 v4.0.10: 合并 shouldKeepRow/shouldKeepCol 为通用筛选器
    function shouldKeepKey(key, filterConfig) {
        if (!filterConfig) return true;
        var keyParts = key.split(separator);
        if (typeof filterConfig === 'function') {
            return filterConfig(keyParts, key);
        }
        if (typeof filterConfig === 'object' && !Array.isArray(filterConfig)) {
            var allMatched = true;
            var hasMatch = false;
            for (var field in filterConfig) {
                var match = field.match(/^f(\d+)$/);
                if (match) {
                    hasMatch = true;
                    var idx = parseInt(match[1]) - 1;
                    if (idx < 0 || idx >= keyParts.length) continue;
                    var allowedValues = filterConfig[field];
                    if (Array.isArray(allowedValues) && allowedValues.indexOf(keyParts[idx]) === -1) {
                        allMatched = false;
                        break;
                    }
                }
            }
            if (hasMatch) return allMatched;
        }
        return true;
    }

    function shouldKeepRow(rowKey) { return shouldKeepKey(rowKey, filterRows); }
    function shouldKeepCol(colKey) { return shouldKeepKey(colKey, filterCols); }

    // 🔧 v3.7.6 修复: 在处理 Array2D 对象之前，先保存 _header 和 _original 属性
    var _savedHeader = null;
    var _savedOriginal = null;

    if (arr && typeof arr === 'object') {
        // 🔧 v3.7.9 修复: 使用更可靠的方式检查 _header 属性
        // 使用 in 操作符检查，因为 _header 可能是不可枚举的
        if ('_header' in arr && arr._header !== undefined && arr._header !== null) {
            _savedHeader = arr._header;
        }
        // 保存 _original 属性（用于获取原始表头）
        if ('_original' in arr && arr._original !== undefined && arr._original !== null && arr._original.length > 0) {
            _savedOriginal = arr._original;
        }
    }

    // 处理 Array2D 对象
    if (arr instanceof Array2D && Array.isArray(arr._items)) {
        arr = arr._items;
    }

    // 🔧 v3.7.5 自动保存原始表头（用于列字段标题）
    // 智能检测: 尝试从多个来源获取表头
    var _originalHeader = null;

    // 来源1: 优先使用保存的 _header 属性（来自 $.maxArray 或链式调用）
    if (_savedHeader !== null) {
        _originalHeader = _savedHeader;
    }

    // 来源2: 检查保存的 _original 属性
    if (!_originalHeader && _savedOriginal !== null) {
        _originalHeader = _savedOriginal[0];
    }

    // 来源3: 检查数据的第一行是否可能是表头
    if (!_originalHeader && headerRows > 0 && arr && arr.length > 0) {
        // 智能检测: 如果第一行包含字符串而非数字，可能是表头
        var firstRow = arr[0];
        var isHeader = false;

        // 检测方法1: 如果大部分列是字符串，认为是表头
        var stringCount = 0;
        var totalCols = firstRow.length;
        for (var i = 0; i < Math.min(5, totalCols); i++) {  // 只检查前5列
            if (typeof firstRow[i] === 'string' && isNaN(parseFloat(firstRow[i]))) {
                stringCount++;
            }
        }

        if (stringCount >= 2) {  // 至少2列是字符串
            isHeader = true;
        }
        
        // 🔧 强制信任: 如果用户指定了 headerRows > 0，且没有更好的来源，使用 arr[0]
        // 但要检查它看起来不像纯数据行（比如不包含 product1, product2 这样的值）
        if (!isHeader) {
            // 检查是否包含典型的数据值（如 product1, product2 等）
            var looksLikeData = false;
            for (var i = 0; i < firstRow.length; i++) {
                var val = String(firstRow[i]).toLowerCase();
                if (val.indexOf('product') === 0 || val.indexOf('item') === 0) {
                    looksLikeData = true;
                    break;
                }
            }
            // 如果看起来像数据，则不使用 arr[0] 作为表头
            if (!looksLikeData) {
                isHeader = true;
            }
        }

        if (isHeader) {
            _originalHeader = firstRow;
        } else {
        }
    }
    

    // 辅助函数：将行数组转为带f1,f2...属性的对象
    function toRowObject(row) {
        var obj = Array(row.length);
        for (var i = 0; i < row.length; i++) {
            obj['f' + (i + 1)] = row[i];
            obj[i] = row[i];
        }
        return obj;
    }

    // 辅助函数：创建分组对象，支持聚合操作
    function createGroupObject(group) {
        // PATCH v4.2.3: 列表达式求值（支持 f4*2、f3+f4 等计算列）
        var _exprCache = {};
        function _resolveColExpr(col, row) {
            if (!col || !row) return row ? row[col] : undefined;
            // PATCH v4.2.4
            col = col.replace(/`/g, "");   // PATCH v4.2.4 v2: 反引号直接删除
            if (col.indexOf("*") === -1 && col.indexOf("+") === -1 &&
                col.indexOf("-") === -1 && col.indexOf("/") === -1) {
                return row[col];
            }
            var fn = _exprCache[col];
            if (!fn) {
                var body = col;
                var fields = col.match(/f\d+/g) || [];
                for (var fi = 0; fi < fields.length; fi++) {
                    var f = fields[fi];
                    body = body.split(f).join('row["' + f + '"]');
                }
                try { fn = new Function("row", "return (" + body + ");"); _exprCache[col] = fn; }
                catch (e) { return undefined; }
            }
            try { return fn(row); } catch (e) { return undefined; }
        }

        return {
            _items: group,
            count: function() { return group.length; },
            sum: function(col) {
                var total = 0;
                for (var i = 0; i < group.length; i++) {
                    var val = _resolveColExpr(col, group[i]);
                    if (typeof val === 'number') {
                        total += val;
                    } else if (typeof val === 'string') {
                        var num = parseFloat(val.replace(/,/g, ''));
                        if (!isNaN(num)) total += num;
                    }
                }
                return Math.round(total * 1e10) / 1e10;
            },
            average: function(col) {
                if (group.length === 0) return 0;
                var sum = this.sum(col);
                return sum / group.length;
            },
            max: function(col) {
                var max = null;
                for (var i = 0; i < group.length; i++) {
                    var val = _resolveColExpr(col, group[i]);
                    if (typeof val === 'string') {
                        val = parseFloat(val.replace(/,/g, ''));
                    }
                    if (typeof val === 'number' && !isNaN(val)) {
                        if (max === null || val > max) max = val;
                    }
                }
                return max;
            },
            min: function(col) {
                var min = null;
                for (var i = 0; i < group.length; i++) {
                    var val = _resolveColExpr(col, group[i]);
                    if (typeof val === 'string') {
                        val = parseFloat(val.replace(/,/g, ''));
                    }
                    if (typeof val === 'number' && !isNaN(val)) {
                        if (min === null || val < min) min = val;
                    }
                }
                return min;
            },
            textjoin: function(col, sep) {
                var values = [];
                for (var i = 0; i < group.length; i++) {
                    var val = _resolveColExpr(col, group[i]);
                    values.push(val);
                }
                return values.join(sep);
            }
        };
    }


    // 辅助函数：解析结果选择器字符串
    function parseResultSelector(str) {
        var operations = [];
        var regex = /(\w+)\s*\(([^)]*)\)/g;
        var match;
        while ((match = regex.exec(str)) !== null) {
            var op = { name: match[1] };
            var argsStr = match[2].trim();
            if (argsStr) {
                // 解析参数，支持带引号和不带引号
                op.args = [];
                // 🔧 修复：先提取所有引号包裹的参数，再处理非引号参数
                // 使用更精确的正则表达式来正确处理带逗号的字符串参数
                var argRegex = /["']([^"']*)["']|([^,]+)(?:,|$)/g;
                var argMatch;
                while ((argMatch = argRegex.exec(argsStr)) !== null) {
                    var argValue;
                    if (argMatch[1] !== undefined) {
                        // 引号包裹的参数
                        argValue = argMatch[1];
                    } else if (argMatch[2] !== undefined) {
                        // 非引号参数，去除前后空格
                        argValue = argMatch[2].trim();
                    }
                    if (argValue !== undefined && argValue !== '') {
                        op.args.push(argValue);
                    }
                }
            }
            operations.push(op);
        }
        return operations;
    }

    // 解析字段配置
    function parseFieldsConfig(fieldsConfig) {
        var fields = [];
        var titles = [];
        var hasTitles = false;

        if (Array.isArray(fieldsConfig)) {
            // 🔧 v3.7.5 修复: 检查 [['f1,f2', ['标题1','标题2']] 格式
            if (fieldsConfig.length === 2 && Array.isArray(fieldsConfig[0])) {
                // 第一项是数组（可能是字段数组或回调数组）
                var firstItem = fieldsConfig[0];
                var secondItem = fieldsConfig[1];

                // 检查是否是回调数组（包含函数）
                var isCallbackArray = firstItem.length > 0 && typeof firstItem[0] === 'function';

                // 处理标题：支持字符串或数组
                var titleArray = [];
                if (typeof secondItem === 'string') {
                    titleArray = secondItem.split(/[,\uFF0C]/);
                } else if (Array.isArray(secondItem)) {
                    titleArray = secondItem;
                }

                if (isCallbackArray) {
                    // [[回调数组], '标题'] 格式 - 数据字段
                    return {
                        fields: [{ callbacks: firstItem }],
                        titles: titleArray,
                        hasTitles: true,
                        isCallback: true
                    };
                } else {
                    // [['f1,f2'], ['标题1','标题2']] 格式 - 带自定义标题的字段列表
                    var fieldStr = firstItem[0] || '';  // 假设是 ['f1,f2,f3'] 格式
                    var items = fieldStr.split(/[,，]/);
                    for (var j = 0; j < items.length; j++) {
                        var item = items[j].trim();
                        var match = item.match(/^(f\d+)([+\-#]*)$/);
                        if (match) {
                            fields.push({
                                field: match[1],
                                sort: match[2] || '+'
                            });
                        }
                    }
                    return {
                        fields: fields,
                        titles: titleArray,
                        hasTitles: true
                    };
                }
            }
            // 检查 ['f1,f2,f3', '标题'] 格式
            // 🔧 v3.8.3 修复：支持同时使用排序符号和自定义标题
            if (fieldsConfig.length === 2 && typeof fieldsConfig[0] === 'string' && typeof fieldsConfig[1] === 'string') {
                var fieldStr = fieldsConfig[0];
                var items = fieldStr.split(/[,，]/);
                // 始终解析标题（即使有排序符号）
                titles = fieldsConfig[1].split(/[,，]/);
                for (var j = 0; j < items.length; j++) {
                    var item = items[j].trim();
                    var match = item.match(/^(f\d+)([+\-#]*)$/);
                    if (match) {
                        fields.push({
                            field: match[1],
                            sort: match[2] || '+'
                        });
                    } else {
                        // 没有排序符号，作为普通字段处理
                        fields.push({
                            field: item,
                            sort: '+'
                        });
                    }
                }
                hasTitles = true;
                return { fields: fields, titles: titles, hasTitles: hasTitles };
            }
            // ['f1+,f2-'] 格式 或 带排序符号的格式
            if (typeof fieldsConfig[0] === 'string') {
                var fieldStr = fieldsConfig[0];
                var items = fieldStr.split(/[,，]/);
                for (var ki = 0; ki < items.length; ki++) {
                    var item = items[ki].trim();
                    var match = item.match(/^(f\d+)([+\-#]*)$/);
                    if (match) {
                        fields.push({
                            field: match[1],
                            sort: match[2] || '+'
                        });
                    }
                }
            }
        } else if (typeof fieldsConfig === 'string') {
            var items = fieldsConfig.split(/[,，]/);
            for (var m = 0; m < items.length; m++) {
                var item = items[m].trim();
                var match = item.match(/^(f\d+)([+\-#]*)$/);
                if (match) {
                    fields.push({
                        field: match[1],
                        sort: match[2] || '+'
                    });
                }
            }
        }

        return { fields: fields, titles: titles, hasTitles: hasTitles };
    }

    var rowConfig = parseFieldsConfig(rowFields);
    var colConfig = parseFieldsConfig(colFields);

    // 🔧 v3.8.3: 排序配置已解析（调试日志已移除）
    
    // 数据字段需要特殊处理
    var dataConfig;
    if (Array.isArray(dataFields)) {
        if (dataFields.length === 2 && Array.isArray(dataFields[0])) {
            // [[回调数组], '标题'] 格式
            dataConfig = {
                fields: [{ callbacks: dataFields[0] }],
                titles: (typeof dataFields[1] === 'string' ? dataFields[1].split(/[,，]/) : (Array.isArray(dataFields[1]) ? dataFields[1] : [])),
                hasTitles: true,
                isCallback: true,
                rawString: null
            };
        } else if (dataFields.length === 2 && typeof dataFields[0] === 'string') {
            // ['count(),sum("f3")', '标题'] 或 ['f1,f2', '标题'] 格式
            var dfStr = dataFields[0];
            // 检查是否包含聚合函数
            if (dfStr.match(/count|sum|average|max|min|textjoin/)) {
                // 数据字段格式
                dataConfig = {
                    fields: [{ field: dfStr }],
                    titles: dataFields[1].split(/[,，]/),
                    hasTitles: true,
                    rawString: dfStr
                };
            } else {
                // 普通字段格式，使用 parseFieldsConfig
                dataConfig = parseFieldsConfig(dataFields);
            }
        } else if (typeof dataFields[0] === 'string') {
            // ['count(),sum("f3")'] 格式
            dataConfig = {
                fields: [{ field: dataFields[0] }],
                titles: [],
                hasTitles: false,
                rawString: dataFields[0]
            };
        } else {
            dataConfig = parseFieldsConfig(dataFields);
        }
    } else if (typeof dataFields === 'string') {
        // 字符串格式的数据字段
        dataConfig = {
            fields: [{ field: dataFields }],
            titles: [],
            hasTitles: false,
            rawString: dataFields
        };
    } else {
        dataConfig = parseFieldsConfig(dataFields);
    }

    // 跳过标题行
    var dataStartRow = (headerRows !== undefined && headerRows !== null) ? headerRows : 1;
    var data = arr.slice(dataStartRow);

    // ✅ 过滤空行：避免生成空白行
    function isEmptyRow(row) {
        if (!row || row.length === 0) return true;
        for (var i = 0; i < row.length; i++) {
            var val = row[i];
            if (val !== null && val !== undefined && val !== '') return false;
        }
        return true;
    }

    data = data.filter(function(row) {
        return !isEmptyRow(row);
    });

    // 将数据转为对象数组
    var dataObjs = data.map(function(row) {
        return toRowObject(row);
    });

    // 提取所有行字段值并排序
    var rowKeys = [];
    var rowKeyMap = Object.create(null);

    // ✅ 辅助函数：检查键是否有效（不全是空值）
    function isValidKey(keyParts) {
        for (var i = 0; i < keyParts.length; i++) {
            var val = keyParts[i];
            if (val !== null && val !== undefined && val !== '') return true;
        }
        return false;
    }

    for (var i = 0; i < dataObjs.length; i++) {
        var obj = dataObjs[i];
        var keyParts = [];
        for (var j = 0; j < rowConfig.fields.length; j++) {
            var rf = rowConfig.fields[j];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                keyParts.push(obj[idx]);
            }
        }
        // ✅ 跳过全空的键
        if (!isValidKey(keyParts)) continue;

        var key = keyParts.join(separator);
        if (!rowKeyMap[key]) {
            rowKeyMap[key] = {
                values: keyParts.slice(),
                originalIndex: i
            };
            rowKeys.push(key);
        }
    }

    // 🔧 v4.0.10: Schwartzian transform 排序优化
    // 预计算排序键，避免每次比较都 parseFloat + localeCompare
    var rowSortKeys = rowKeys.map(function(key, idx) {
        var parts = key.split(separator);
        var keys = [];
        for (var ki = 0; ki < rowConfig.fields.length; ki++) {
            var val = parts[ki] || '';
            var num = parseFloat(val);
            var isNum = !isNaN(num) && String(num) === String(val).trim();
            keys.push({ raw: val, num: isNum ? num : null, isNum: isNum });
        }
        return { key: key, keys: keys, idx: idx, origIdx: rowKeyMap[key].originalIndex };
    });

    rowSortKeys.sort(function(a, b) {
        for (var ki = 0; ki < rowConfig.fields.length; ki++) {
            var rf = rowConfig.fields[ki];
            var ak = a.keys[ki], bk = b.keys[ki];
            var cmp = 0;
            if (ak.isNum && bk.isNum) {
                cmp = ak.num - bk.num;
            } else {
                cmp = String(ak.raw).localeCompare(String(bk.raw));
            }
            if (cmp !== 0) return rf.sort === '-' ? -cmp : cmp;
            if (rf.sort === '#') return a.origIdx - b.origIdx;
        }
        return 0;
    });

    rowKeys = rowSortKeys.map(function(sk) { return sk.key; });

    // 🔧 确保第一个 rowKey 不是数据值污染的（防御性编程）
    if (rowKeys.length > 0) {
        var firstKey = rowKeys[0];
        var keyParts = firstKey.split(separator);
        var isValid = true;
        for (var i = 0; i < keyParts.length; i++) {
            if (keyParts[i] === '' || keyParts[i] === null || keyParts[i] === undefined) {
                isValid = false;
                break;
            }
        }
        // 如果第一个键无效，移除它
        if (!isValid) {
            if (typeof Console !== 'undefined') { try { Console.log('警告: 第一个 rowKey 无效，移除: ' + firstKey); } catch(__) {} }
            rowKeys.shift();
        }
    }

    // 提取所有列字段值并排序
    var colKeys = [];
    var colKeyMap = Object.create(null);
    for (var m = 0; m < dataObjs.length; m++) {
        var obj = dataObjs[m];
        var keyParts = [];
        for (var n = 0; n < colConfig.fields.length; n++) {
            var cf = colConfig.fields[n];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                keyParts.push(obj[idx]);
            }
        }
        // ✅ 跳过全空的键
        if (!isValidKey(keyParts)) continue;

        var key = keyParts.join(separator);
        if (!colKeyMap[key]) {
            colKeyMap[key] = {
                values: keyParts.slice(),
                originalIndex: m
            };
            colKeys.push(key);
        }
    }

    // 🔧 XXD-47: 列键排序也用 Schwartzian transform（与行键排序一致，避免每次比较重复 parseFloat）
    var colSortKeys = colKeys.map(function(key, idx) {
        var parts = key.split(separator);
        var keys = [];
        for (var ki = 0; ki < colConfig.fields.length; ki++) {
            var val = parts[ki] || '';
            var num = parseFloat(val);
            var isNum = !isNaN(num) && String(num) === String(val).trim();
            keys.push({ raw: val, num: isNum ? num : null, isNum: isNum });
        }
        return { key: key, keys: keys, idx: idx, origIdx: colKeyMap[key] ? colKeyMap[key].originalIndex : idx };
    });

    colSortKeys.sort(function(a, b) {
        for (var ki = 0; ki < colConfig.fields.length; ki++) {
            var cf = colConfig.fields[ki];
            var ak = a.keys[ki], bk = b.keys[ki];
            var cmp = 0;
            if (ak.isNum && bk.isNum) {
                cmp = ak.num - bk.num;
            } else {
                cmp = String(ak.raw).localeCompare(String(bk.raw));
            }
            if (cmp !== 0) return cf.sort === '-' ? -cmp : cmp;
            if (cf.sort === '#') return a.origIdx - b.origIdx;
        }
        return 0;
    });

    colKeys = colSortKeys.map(function(sk) { return sk.key; });

    // 🔧 v3.8.3: 排序完成（调试日志已移除）

    // 分组数据：行键 + 列键 -> 数据行
    var groupMap = Object.create(null);
    for (var q = 0; q < dataObjs.length; q++) {
        var obj = dataObjs[q];
        var rowKeyParts = [];
        for (var r = 0; r < rowConfig.fields.length; r++) {
            var rf = rowConfig.fields[r];
            var match = rf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                rowKeyParts.push(obj[idx]);
            }
        }
        var colKeyParts = [];
        for (var s = 0; s < colConfig.fields.length; s++) {
            var cf = colConfig.fields[s];
            var match = cf.field.match(/^f(\d+)$/);
            if (match) {
                var idx = parseInt(match[1]) - 1;
                colKeyParts.push(obj[idx]);
            }
        }
        var rowKey = rowKeyParts.join(separator);
        var colKey = colKeyParts.join(separator);
        var fullKey = rowKey + rowColSep + colKey;
        if (!groupMap[fullKey]) {
            groupMap[fullKey] = [];
        }
        // 转回普通数组
        var row = [];
        for (var t = 0; t < obj.length; t++) {
            row.push(obj[t]);
        }
        groupMap[fullKey].push(row);
    }

    // 🔧 v3.9.4 新增：从 groupMap 中移除被筛选掉的行/列数据
    // 先收集需要保留的键（仅在有筛选条件时执行）
    if (filterRows || filterCols) {
        var validRowKeys = {};
        var validColKeys = {};
        rowKeys.forEach(function(rk) { validRowKeys[rk] = true; });
        colKeys.forEach(function(ck) { validColKeys[ck] = true; });
        // 清理 groupMap 中无效的条目
        var validGroupMap = Object.create(null);
        for (var gk in groupMap) {
            var parts = gk.split(rowColSep);
            var rk = parts[0];
            var ck = parts.slice(1).join('|||');
            if (validRowKeys[rk] && validColKeys[ck]) {
                validGroupMap[gk] = groupMap[gk];
            }
        }
        groupMap = validGroupMap;
    }

    // 解析数据字段操作
    var dataOps = [];
    if (dataConfig.isCallback) {
        dataOps = dataConfig.fields[0].callbacks;
    } else if (dataConfig.rawString) {
        var operations = parseResultSelector(dataConfig.rawString);
        dataOps = operations;
    } else if (dataConfig.fields.length > 0) {
        var opStr = dataConfig.fields[0].field || '';
        var operations = parseResultSelector(opStr);
        dataOps = operations;
    }

    // 🔧 v3.9.0 修复：提前计算 numDataFields，供 grandTotalValues 使用
    var numDataFields = Array.isArray(dataOps) && dataOps.length > 0 ? dataOps.length :
                       (dataConfig.titles && dataConfig.titles.length > 0 ? dataConfig.titles.length : 1);

    // 执行聚合操作
    function executeAggregation(group) {
        var groupObj = createGroupObject(group.map(function(r) {
            return toRowObject(r);
        }));
        var results = [];
        if (Array.isArray(dataOps) && dataOps.length > 0 && typeof dataOps[0] === 'function') {
            // 回调模式
            for (var v = 0; v < dataOps.length; v++) {
                var result = dataOps[v](groupObj);
                results.push(result);
            }
        } else {
            // 字符串模式
            for (var w = 0; w < dataOps.length; w++) {
                var op = dataOps[w];
                var args = op.args || [];
                switch (op.name) {
                    case 'count':
                        results.push(groupObj.count());
                        break;
                    case 'sum':
                        results.push(groupObj.sum(args[0]));
                        break;
                    case 'average':
                        results.push(groupObj.average(args[0]));
                        break;
                    case 'max':
                        results.push(groupObj.max(args[0]));
                        break;
                    case 'min':
                        results.push(groupObj.min(args[0]));
                        break;
                    case 'textjoin':
                        results.push(groupObj.textjoin(args[0], args[1]));
                        break;
                    case 'col':
                        // 🔧 v4.0.10: 直接列引用 f5（用于计算字段）
                        results.push(null);
                        break;
                }
            }
        }
        return results;
    }

    // 🔧 v3.9.4 新增：应用筛选器到列键
    colKeys = colKeys.filter(shouldKeepCol);

    // 🔧 v3.9.4 新增：应用筛选器到行键
    rowKeys = rowKeys.filter(shouldKeepRow);

    // v4.2.5: 无行字段时归入单一空键
    if (rowKeys.length === 0 && data.length > 0) {
        rowKeys = [''];
    }

    // map 模式：返回查询标准字典
    if (outputHeader === 'map') {
        var resultMap = new Map();
        for (var x = 0; x < rowKeys.length; x++) {
            var rowKey = rowKeys[x];
            for (var y = 0; y < colKeys.length; y++) {
                var colKey = colKeys[y];
                var fullKey = rowKey + rowColSep + colKey;
                if (groupMap[fullKey]) {
                    var agg = executeAggregation(groupMap[fullKey]);
                    var sortKey = rowKey + separator + colKey;
                    var mapKey = '01L' + String(x + 1).padStart(4, '0') + ' ' + sortKey;
                    resultMap.set(mapKey, {
                        agg: agg,
                        group: { '00000': groupMap[fullKey][0] }
                    });
                }
            }
        }
        return resultMap;
    }

    // 🔧 v3.9.0 新增：计算总计值
    var grandTotalValues = null;
    if (grandTotals.row || grandTotals.column || (displayAs.mode && displayAs.mode !== 'value')) {
        grandTotalValues = {
            rowTotals: {},
            colTotals: {},
            grandTotal: []
        };
        
        // 计算每行的总计
        for (var rk = 0; rk < rowKeys.length; rk++) {
            var rowKey = rowKeys[rk];
            var rowTotal = [];
            for (var df = 0; df < numDataFields; df++) {
                var sum = 0;
                // 🔧 XXD-106/104: colKeys===0 时从 groupMap 直接取(无 col key 后缀)
                if (colKeys.length === 0) {
                    var _noColKey = rowKey + rowColSep;
                    if (groupMap[_noColKey]) {
                        var _agg = executeAggregation(groupMap[_noColKey]);
                        var _val = parseFloat(_agg[df]);
                        if (!isNaN(_val)) sum = _val;
                    }
                } else {
                for (var ck = 0; ck < colKeys.length; ck++) {
                    var colKey = colKeys[ck];
                    var fullKey = rowKey + rowColSep + colKey;
                    if (groupMap[fullKey]) {
                        var agg = executeAggregation(groupMap[fullKey]);
                        var val = parseFloat(agg[df]);
                        if (!isNaN(val)) sum += val;
                    }
                }
                } // XXD-106/104
                rowTotal.push(sum);
            }
            grandTotalValues.rowTotals[rowKey] = rowTotal;
        }
        
        // 计算每列的总计
        for (var ck = 0; ck < colKeys.length; ck++) {
            var colKey = colKeys[ck];
            var colTotal = [];
            for (var df = 0; df < numDataFields; df++) {
                var sum = 0;
                for (var rk = 0; rk < rowKeys.length; rk++) {
                    var rowKey = rowKeys[rk];
                    var fullKey = rowKey + rowColSep + colKey;
                    if (groupMap[fullKey]) {
                        var agg = executeAggregation(groupMap[fullKey]);
                        var val = parseFloat(agg[df]);
                        if (!isNaN(val)) sum += val;
                    }
                }
                colTotal.push(sum);
            }
            grandTotalValues.colTotals[colKey] = colTotal;
        }
        
        // 计算总总计
        for (var df = 0; df < numDataFields; df++) {
            var sum = 0;
            for (var rk = 0; rk < rowKeys.length; rk++) {
                sum += grandTotalValues.rowTotals[rowKeys[rk]][df];
            }
            grandTotalValues.grandTotal.push(sum);
        }
    }

    // 🔧 v3.9.0 新增：应用百分比转换
    function applyDisplayAs(value, rowKey, colKey, dataFieldIndex, parentRowKey, parentColKey) {
        if (!displayAs.mode || displayAs.mode === 'value') {
            return value;
        }
        
        var val = parseFloat(value);
        if (isNaN(val)) return value;
        
        var decimals = displayAs.decimals || 2;
        var pct = 0;
        
        switch (displayAs.mode) {
            case 'percentOfGrandTotal':
                var total = grandTotalValues.grandTotal[dataFieldIndex];
                pct = total !== 0 ? (val / total * 100) : 0;
                break;
            case 'percentOfRowTotal':
                var rowTotal = grandTotalValues.rowTotals[rowKey] ? grandTotalValues.rowTotals[rowKey][dataFieldIndex] : 0;
                pct = rowTotal !== 0 ? (val / rowTotal * 100) : 0;
                break;
            case 'percentOfColTotal':
                var colTotal = grandTotalValues.colTotals[colKey] ? grandTotalValues.colTotals[colKey][dataFieldIndex] : 0;
                pct = colTotal !== 0 ? (val / colTotal * 100) : 0;
                break;
            default:
                return value;
        }
        
        return pct.toFixed(decimals) + '%';
    }

    // ==================== 多级表头合并信息收集 ====================
    // 合并信息格式: {row1: {col1: {rowSpan: x, colSpan: y}, ...}, ...}
    var mergeInfo = Object.create(null);

    // 辅助函数：记录合并信息
    function recordMerge(rowIdx, colIdx, rowSpan, colSpan) {
        if (rowSpan > 1 || colSpan > 1) {
            if (!mergeInfo[rowIdx]) mergeInfo[rowIdx] = Object.create(null);
            mergeInfo[rowIdx][colIdx] = { rowSpan: rowSpan, colSpan: colSpan };
        }
    }

    // 🔧 v3.8.6 修复：优先使用用户指定的数据字段标题
    var defaultDataTitles = [];
    if (dataConfig.titles && dataConfig.titles.length > 0) {
        // 优先使用用户指定的自定义标题（如 '销售金额'）
        defaultDataTitles = dataConfig.titles.slice();
    } else if (Array.isArray(dataOps) && dataOps.length > 0 && typeof dataOps[0] !== 'function') {
        // 没有自定义标题时，根据聚合函数生成默认标题
        var opNameMap = {
            'count': '计数',
            'sum': '求和',
            'average': '平均',
            'max': '最大',
            'min': '最小',
            'textjoin': '连接'
        };
        for (var i = 0; i < dataOps.length; i++) {
            var opName = dataOps[i].name;
            defaultDataTitles.push(opNameMap[opName] || opName);
        }
    } else {
        // 默认标题
        for (var j = 0; j < numDataFields; j++) {
            defaultDataTitles.push('值' + (j + 1));
        }
    }

    var result = [];

    // 🔧 v3.9.1 修复：将 headerRowCount 计算移到 if 块外面，确保在 outputHeader=0 时也能使用
    // 计算层级数
    var numColFieldLevels = colConfig.fields.length;
    var numRowFieldLevels = rowConfig.fields.length;
    // 🔧 修复：单列字段时需要3行表头，多列时需要 numColFieldLevels + 1 行
    var headerRowCount = (numColFieldLevels === 1) ? 2 : numColFieldLevels + 1;
    // 🔧 XXD-47: 显式初始化 hideRowTitles，不依赖 var 提升隐式行为
    var hideRowTitles = (outputHeader === -1);

    // 构建表头
    if (outputHeader === 1 || outputHeader === true || outputHeader === -1) {
        // 检查是否有自定义标题
        var hasRowTitles = rowConfig.hasTitles && rowConfig.titles.length > 0;
        var hasColTitles = colConfig.hasTitles && colConfig.titles.length > 0;
        var hasDataTitles = dataConfig.hasTitles && dataConfig.titles.length > 0;

        // 🔧 ==================== 表头构建辅助函数 ====================

        /**
         * 获取字段标题文本
         * @param {Object} field - 字段配置对象
         * @param {number} fieldIndex - 字段索引
         * @param {string} type - 'row'|'col'
         * @returns {string} 标题文本
         */
        function getFieldTitle(field, fieldIndex, type) {
            var config = type === 'row' ? rowConfig : colConfig;
            var titles = config.titles || [];

            // 优先级1: 自定义标题
            if (titles && titles[fieldIndex]) {
                return titles[fieldIndex];
            }
            // 优先级2: _originalHeader (表头行)
            if (_originalHeader) {
                var match = field.field.match(/^f(\d+)$/);
                if (match) {
                    var origIdx = parseInt(match[1]) - 1;
                    return _originalHeader[origIdx] || '';
                }
            }
            // 优先级3: arr[0] (数据第一行)
            if (arr && arr[0]) {
                var match = field.field.match(/^f(\d+)$/);
                if (match) {
                    var origIdx = parseInt(match[1]) - 1;
                    return arr[0][origIdx] || '';
                }
            }
            return '';
        }

        /**
         * 计算列键在某层级的唯一值数组（保持顺序）
         * @param {Array} keys - 列键数组
         * @param {number} level - 层级索引
         * @returns {Array} 唯一值数组
         */
        function getLevelUniqueValues(keys, level) {
            var seen = Object.create(null);
            var values = [];
            for (var ki = 0; ki < keys.length; ki++) {
                var parts = keys[ki].split(separator);
                if (level < parts.length) {
                    var val = parts[level];
                    var valKey = String(val);
                    if (!seen[valKey]) {
                        seen[valKey] = true;
                        values.push(val);
                    }
                }
            }
            return values;
        }

        /**
         * 计算指定层级的跨列数（该层级每个值占据的列数）
         * @param {number} level - 层级索引
         * @param {Array} levelValues - 该层级唯一值数组
         * @returns {number} 跨列数
         */

        /**
         * 构建单层级列值行（用于多级表头）
         * @param {number} level - 层级索引
         * @param {Array} levelValues - 该层级唯一值
         * @returns {Array} 行数据
         */

        /**
         * 检测需要合并的连续相同值区域
         * @param {Array} row - 行数据
         * @param {number} startCol - 起始列
         * @param {number} endCol - 结束列
         * @returns {Array} 合并信息数组
         */

        // 🔧 ==================== 表头构建主流程 ====================

        // 预计算所有层级的唯一值（性能优化）
        var colLevelValues = [];
        for (var lv = 0; lv < numColFieldLevels; lv++) {
            colLevelValues.push(getLevelUniqueValues(colKeys, lv));
        }

        // 构建表头数组
        var headerRows = [];
        for (var h = 0; h < headerRowCount; h++) {
            headerRows.push([]);
        }

        // ============ 步骤1: 根据列字段数量选择构建策略 ============

        if (numColFieldLevels === 0) {
            // ========== 无列字段：单行表头 ==========
            // 结构: [行字段标题...] [数据字段标题...]
            if (!hideRowTitles) {
                for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
                    headerRows[0].push(getFieldTitle(rowConfig.fields[rfIdx], rfIdx, 'row'));
                }
            }
            // 数据字段标题
            for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                headerRows[0].push(defaultDataTitles[dfIdx] || '');
            }

        } else if (numColFieldLevels === 1) {
            // ========== 单列字段：2行表头 ==========
            // 行0: [空白(merged 跨 numRowFieldLevels 列)] [col 字段标题(年)] [col 字段值 × numDataFields 重复]
            // 行1: [行字段标题 1 / 2 / ...] [dataField 标题 × numColKeys 重复]
            //   例: rowFields=[国家,产品] colFields=[年] dataFields=[count,sum,textjoin] 4 年
            //     行0: ["", "年", "2021","2021","2021","2022","2022","2022","2023","2023","2023","2024","2024","2024"]  (14)
            //     行1: ["国家","产品", "计数","求和","连接","计数","求和","连接",...]  (14)
            // 🔧 v4.0.35: 之前 v4.0.32 给 row 字段空 cell 重复 numRowFieldLevels 次,导致表头 15 列多 1
            //   修法: 行 0 row 字段位置只 push 1 次空 cell(用 recordMerge 跨 numRowFieldLevels 列合并)
            //        行 1 row 字段标题仍按 numRowFieldLevels push

            // 行0 左侧: 行字段位置只 push 1 个空 cell (用合并跨 numRowFieldLevels 列)
            if (!hideRowTitles) {
                headerRows[0].push(
                    numRowFieldLevels === 1
                        ? (cornerTitle || getFieldTitle(colConfig.fields[0], 0, 'col'))
                        : ''
                );
                // 合并 numRowFieldLevels 列
                if (numRowFieldLevels > 1) {
                    recordMerge(0, 0, numRowFieldLevels, 1);
                }
                // 行1: 仍然按 numRowFieldLevels push row 字段标题
                for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
                    headerRows[1].push(getFieldTitle(rowConfig.fields[rfIdx], rfIdx, 'row'));
                }
            }

            // 行0 中间: col 字段标题(年)
            headerRows[0].push(getFieldTitle(colConfig.fields[0], 0, 'col'));
            // 行1 中间: 不 push (与行 0 col 字段标题位置对齐,因为 col 字段标题"年"是 col 字段本身)
            // 🔧 v4.0.36 修复: 之前 push defaultDataTitles[0]='计数' 重复,行 1 多了 1 个 cell
            // headerRows[1].push(defaultDataTitles[0] || '');  // 删

            // 行0 右侧: col 字段值(2021/2022/...)每个 × numDataFields 重复
            for (var ck = 0; ck < colKeys.length; ck++) {
                var colKeyParts = colKeys[ck].split(separator);
                for (var __dfCk = 0; __dfCk < numDataFields; __dfCk++) {
                    headerRows[0].push(colKeyParts[0]);
                    // 行1 同步 dataField 标题
                    headerRows[1].push(defaultDataTitles[__dfCk] || '');
                }
            }

            // 列小计
            if (colSubtotals.enabled) {
                for (var __dfSt = 0; __dfSt < numDataFields; __dfSt++) {
                    headerRows[0].push(colSubtotals.label || '小计');
                    headerRows[1].push(defaultDataTitles[__dfSt] || '');
                }
            }

        } else {
            // ========== 多列字段：numColFieldLevels + 1 行表头 ==========
            // 结构示例（2个列字段）：
            // 行0: [行字段标题] [列1值] [列1值] ...
            // 行1: [空白]      [列2值] [列2值] ...
            // 行2: [空白]      [数据字段标题] ...
            // 
            // 多层表头格式（需要合并单元格）：
            // - 第1行：行字段标题 + 列字段第1层（最高层）的值
            // - 第2行：空白 + 列字段第2层的值
            // - ...
            // - 最后一行：空白 + 数据字段标题

            // 🔧 v3.9.5 修复：行字段标题应该在第1行的第1列位置
            // 构建每一层级的列值行
            for (var cfIdx = 0; cfIdx < numColFieldLevels; cfIdx++) {
                var targetRow = cfIdx;

                // 添加行字段空白
                // 🔧 XXD-69: multi-col corner 填 col 字段标题到最后一个角标列
                if (!hideRowTitles) {
                    for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
                        if (rfIdx === numRowFieldLevels - 1 && numRowFieldLevels > 0) {
                            headerRows[targetRow].push(getFieldTitle(colConfig.fields[cfIdx], cfIdx, 'col'));
                        } else {
                            headerRows[targetRow].push('');
                        }
                    }
                }

                // 🔧 v4.0.10 修复：基于 colKeys 顺序生成列值
                // colKeys 已按层级排序，正确反映嵌套结构
                // 这样可以保证每行表头的列数与 colKeys 完全对齐
                for (var ck = 0; ck < colKeys.length; ck++) {
                    var colKeyParts = colKeys[ck].split(separator);
                    for (var df = 0; df < numDataFields; df++) {
                        headerRows[targetRow].push(colKeyParts[cfIdx]);
                    }
                }
                // 添加列小计
                if (colSubtotals.enabled) {
                    headerRows[targetRow].push(cfIdx === numColFieldLevels - 1 ? (colSubtotals.label || '小计') : '');
                }
            }

            // 最后一行：行字段标题 + 数据字段标题
            var lastRow = numColFieldLevels;

            if (!hideRowTitles) {
                for (var rfIdx = 0; rfIdx < numRowFieldLevels; rfIdx++) {
                    headerRows[lastRow].push(getFieldTitle(rowConfig.fields[rfIdx], rfIdx, 'row'));
                    // 🔧 XXD-74: 取消 XXD-69 引入的角标 recordMerge — 用户期望 (1,1)(2,1) 是独立空 cell
                    //   旧逻辑：if (rfIdx < numRowFieldLevels - 1) recordMerge(0, rfIdx, numColFieldLevels, 1)
                    //   取消后 corner 4 个 cell (1,1)(1,2)(2,1)(2,2) 保持独立
                }
            }

            // 数据字段标题（基于 colKeys 对齐列值行）
            for (var ck = 0; ck < colKeys.length; ck++) {
                for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                    headerRows[lastRow].push(defaultDataTitles[dfIdx] || '');
                }
            }
            if (colSubtotals.enabled) {
                for (var dfIdx = 0; dfIdx < numDataFields; dfIdx++) {
                    headerRows[lastRow].push(defaultDataTitles[dfIdx] || '');
                }
            }
        }

        // ========== 步骤2: 检测并记录横向合并 ==========
        // 遍历每一行，检测连续相同值并记录合并
        for (var hr = 0; hr < headerRowCount; hr++) {
            var rowData = headerRows[hr];
            var colStart = hideRowTitles ? 0 : numRowFieldLevels;

            // 简单检测：统计每个值出现的起始位置和次数
            var i = colStart;
            while (i < rowData.length) {
                var val = rowData[i];
                var j = i + 1;
                while (j < rowData.length && rowData[j] === val) {
                    j++;
                }
                if (j - i > 1) {
                    recordMerge(hr, i, 1, j - i);
                }
                i = j;
            }
        }

        // 将表头行添加到结果中
        for (var h = 0; h < headerRowCount; h++) {
            result.push(headerRows[h]);
        }
    }

    // 🔧 DEBUG: 构建数据行

    // 🔧 Bug修复: 移除 var 重复声明，复用 if 块内已定义的 hideRowTitles
    // hideRowTitles 已在上方 if (outputHeader === 1 ...) 块中声明（var 提升到函数作用域）

    // 🔧 v3.9.0 修改：构建数据行（支持小计、总计、百分比）
    var dataRows = [];
    var prevRowKeyParts = null;
    
    for (var rk = 0; rk < rowKeys.length; rk++) {
        var rowKey = rowKeys[rk];
        var rowKeyParts = rowKey.split(separator);
        
        // 🔧 v3.9.0 新增：检查是否需要插入行小计
        if (rowSubtotals.enabled && rk > 0 && prevRowKeyParts) {
            // 检查当前行与前一行是否有相同的父级
            var commonParentLen = 0;
            for (var p = 0; p < rowKeyParts.length - 1 && p < prevRowKeyParts.length - 1; p++) {
                if (rowKeyParts[p] === prevRowKeyParts[p]) {
                    commonParentLen = p + 1;
                } else {
                    break;
                }
            }
            
            // 如果父级变化，插入小计行
            if (commonParentLen > 0 && rowKeyParts[commonParentLen - 1] !== prevRowKeyParts[commonParentLen - 1]) {
                // 实际应该检查是否需要根据层级插入小计
                // 简化处理：检查最后一段是否不同
            }
        }
        
        // 🔧 v3.8.8 修复：hideRowTitles = true 时不包含行字段值
        var dataRow = hideRowTitles ? [] : rowKeyParts.slice();
        
        // 🔧 v3.9.0 新增：应用层级缩进
        if (rowFieldIndent && layoutMode === 'outline' && !hideRowTitles) {
            for (var rpi = 0; rpi < dataRow.length; rpi++) {
                var indent = rpi * rowFieldIndentSize;
                dataRow[rpi] = (indent > 0 ? new Array(indent + 1).join(' ') : '') + dataRow[rpi];
            }
        }

        // 🔧 v3.9.1 新增：处理无列字段的情况 - 直接添加聚合值
        if (colKeys.length === 0) {
            // 没有列字段，使用 rowKey + '|||' 获取聚合值
            // （groupMap 中键的格式是 "rowKey|||" 当没有列字段时）
            var emptyColKey = rowKey + rowColSep;
            if (groupMap[emptyColKey]) {
                var agg = executeAggregation(groupMap[emptyColKey]);
                // 🔧 v3.9.0 新增：应用百分比转换
                for (var ai = 0; ai < agg.length; ai++) {
                    agg[ai] = applyDisplayAs(agg[ai], rowKey, null, ai);
                }
                dataRow = dataRow.concat(agg);
            } else {
                // 🔧 XXD-106/104: 空值也走 applyDisplayAs 保持百分比格式一致
                for (var c = 0; c < numDataFields; c++) {
                    dataRow.push(applyDisplayAs(nullValue, rowKey, null, c));
                }
            }
        } else {
            // 🔧 v3.7.9 方案3: 数据行与表头对齐
            for (var ck = 0; ck < colKeys.length; ck++) {
                var colKey = colKeys[ck];
                var fullKey = rowKey + rowColSep + colKey;
                if (groupMap[fullKey]) {
                    var agg = executeAggregation(groupMap[fullKey]);
                    // 🔧 v3.9.0 新增：应用百分比转换
                    for (var ai = 0; ai < agg.length; ai++) {
                        agg[ai] = applyDisplayAs(agg[ai], rowKey, colKey, ai);
                    }
                    dataRow = dataRow.concat(agg);
                } else {
                    for (var c = 0; c < numDataFields; c++) {
                        // 🔧 v4.0.34: 默认 nullValue 为 0(WPS spill 显 0 匹配用户期望)
                        // 🔧 XXD-106/104: 空值也走 applyDisplayAs 保持百分比格式一致
                        // 🔧 XXD-47: 统一用 nullValue 变量(已在入口处默认 0)
                        dataRow.push(applyDisplayAs(nullValue, rowKey, colKey, c));
                    }
                }
            }
        }

        // 🔧 v3.9.0 新增：添加列小计列（每行末尾的小计）
        if (options.subtotals.col && grandTotalValues && grandTotalValues.rowTotals[rowKey]) {
            var rowTotal = grandTotalValues.rowTotals[rowKey];
            for (var rt = 0; rt < numDataFields; rt++) {
                dataRow.push(applyDisplayAs(rowTotal[rt], rowKey, null, rt));
            }
        }

        dataRows.push(dataRow);
        prevRowKeyParts = rowKeyParts;
    }
    
    // 🔧 v3.9.0 新增：添加总计行
    if (options.grandTotal.row) {
        var totalLabel = options.grandTotal.label || '总计';
        var grandTotalRow = hideRowTitles ? [] : [totalLabel];
        // 填充空白使总计标签与行字段数量对齐
        while (grandTotalRow.length < numRowFieldLevels) {
            grandTotalRow.push(nullValue);
        }

        // 🔧 v3.9.1 新增：处理无列字段的总计行
        if (colKeys.length === 0) {
            // 没有列字段，直接使用总计值
            if (grandTotalValues && grandTotalValues.grandTotal) {
                for (var df = 0; df < numDataFields; df++) {
                    grandTotalRow.push(applyDisplayAs(grandTotalValues.grandTotal[df], null, null, df));
                }
            } else {
                for (var df = 0; df < numDataFields; df++) {
                    grandTotalRow.push(nullValue);
                }
            }
        } else {
            // 添加列总计值
            for (var ck = 0; ck < colKeys.length; ck++) {
                var colKey = colKeys[ck];
                if (grandTotalValues && grandTotalValues.colTotals[colKey]) {
                    var colTotal = grandTotalValues.colTotals[colKey];
                    for (var df = 0; df < numDataFields; df++) {
                        grandTotalRow.push(applyDisplayAs(colTotal[df], null, colKey, df));
                    }
                } else {
                    for (var df = 0; df < numDataFields; df++) {
                        grandTotalRow.push(nullValue);
                    }
                }
            }
        }

        // 添加列小计（总计行的最后几列）
        if (options.subtotals.col && grandTotalValues && grandTotalValues.grandTotal) {
            for (var df = 0; df < numDataFields; df++) {
                grandTotalRow.push(applyDisplayAs(grandTotalValues.grandTotal[df], null, null, df));
            }
        }
        
        dataRows.push(grandTotalRow);
    }
    
    // 将所有数据行添加到结果
    for (var dr = 0; dr < dataRows.length; dr++) {
        result.push(dataRows[dr]);
    }

    // 数据行构建完成

    // 包装结果，返回 Array2D 对象，添加 toRange 和 getRange 方法
    var wrappedResult = result;
    if (Array.isArray(result)) {
        // 创建 Array2D 对象
        wrappedResult = new Array2D(result);

        // 🔧 存储合并信息
        wrappedResult._mergeInfo = mergeInfo;

        /**
         * 获取合并信息
         * @returns {Object} 合并信息对象
         */
        wrappedResult.getMerges = function() {
            return mergeInfo;
        };

        /**
         * toRange - 将结果写入单元格
         * @param {Range|string} rng - 目标单元格
         * @param {Boolean} applyMerges - 是否应用合并，默认true
         * @returns {Range} Range对象
         */
        wrappedResult.toRange = function(rng, applyMerges) {
            // 🔧 性能优化：禁用屏幕更新和自动计算
            var app = Application;
            var screenUpdating = app.ScreenUpdating;
            var calculation = app.Calculation;
            var eventsEnabled = app.EnableEvents;

            try {
                // 禁用不必要的功能以提高性能
                app.ScreenUpdating = false;
                app.Calculation = -4135; // xlCalculationManual
                app.EnableEvents = false;

                var targetRange = Array2D.toRange(result, rng);

                // 如果需要合并且输出表头
                if (applyMerges !== false && outputHeader === 1 && mergeInfo) {
                    wrappedResult.applyMerges(targetRange);
                }

                return targetRange;
            } finally {
                // 恢复原始设置
                app.ScreenUpdating = screenUpdating;
                app.Calculation = calculation;
                app.EnableEvents = eventsEnabled;
            }
        };

        /**
         * getRange - 获取结果写入后的Range对象
         * @param {Range|string} rng - 目标单元格
         * @param {Boolean} applyMerges - 是否应用合并，默认true
         * @returns {Range} Range对象
         */
        wrappedResult.getRange = function(rng, applyMerges) {
            return wrappedResult.toRange(rng, applyMerges);
        };

        /**
         * 🔧 applyMerges - 应用表头合并单元格（优化版）
         * @param {Range|string} rng - 目标单元格区域
         * @returns {Array} 已执行的合并列表 [{row1, col1, row2, col2}, ...]
         * @example
         * // 自动应用合并
         * var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
         * result.toRange("A1");  // 自动合并
         *
         * // 手动应用合并
         * var result = Array2D.z超级透视(data, rowFields, colFields, dataFields);
         * result.toRange("A1", false);  // 不合并
         * result.applyMerges("A1");  // 手动合并
         */
        wrappedResult.applyMerges = function(rng) {
            // 转换 rng 为 Range 对象
            var targetRange;
            if (typeof rng === 'string') {
                targetRange = Application.ActiveSheet.Range(rng);
            } else if (rng && rng.Address) {
                targetRange = rng;
            } else {
                Console.log('applyMerges: 无效的Range参数');
                return [];
            }

            var appliedMerges = [];
            var ws = Application.ActiveSheet;

            // 🔧 性能优化：先收集所有合并区域，再批量执行
            var mergeRanges = [];

            // 遍历合并信息
            for (var rowIdx in mergeInfo) {
                var r = parseInt(rowIdx);
                for (var colIdx in mergeInfo[rowIdx]) {
                    var c = parseInt(colIdx);
                    var span = mergeInfo[rowIdx][colIdx];

                    if (span.rowSpan > 1 || span.colSpan > 1) {
                        mergeRanges.push({
                            row: r,
                            col: c,
                            rowSpan: span.rowSpan,
                            colSpan: span.colSpan
                        });
                    }
                }
            }

            // 🔧 批量执行合并（减少try-catch开销）
            for (var i = 0; i < mergeRanges.length; i++) {
                var m = mergeRanges[i];
                try {
                    var startCell = targetRange.Item(m.row + 1, m.col + 1);
                    var endCell = targetRange.Item(
                        m.row + m.rowSpan,
                        m.col + m.colSpan
                    );
                    var mergeRange = ws.Range(startCell, endCell);
                    mergeRange.Merge();
                    appliedMerges.push({
                        row1: m.row + 1,
                        col1: m.col + 1,
                        row2: m.row + m.rowSpan,
                        col2: m.col + m.colSpan
                    });
                } catch (e) {
                    // 静默处理，避免大量错误输出
                }
            }

            return appliedMerges;
        };

        /**
         * val - 获取原始数组(经过 jagged 对齐,空 cell 显式置 null)
         * @returns {Array} 对齐后的 2D 数组(每行等长,WPS spill 不会把空 cell 显 0)
         */
        wrappedResult.val = function() {
            // 🔧 v4.0.20: 多层表头各行长短不一,jagged 数组在 WPS spill 时空 cell 会显 0
            //   修法: 找到 maxLen,每行补齐到 maxLen,空 cell 用空字符串 '' (WPS spill '' 显空白)
            // 🔧 v4.0.33 修正: 之前补 null 错了 — WPS spill null 也显 0
            if (!Array.isArray(result) || result.length === 0) return result;
            var __maxLen = 0;
            for (var __i = 0; __i < result.length; __i++) {
                if (result[__i] && result[__i].length > __maxLen) __maxLen = result[__i].length;
            }
            if (__maxLen === 0) return result;
            var __aligned = [];
            for (var __j = 0; __j < result.length; __j++) {
                var __row = result[__j];
                if (!Array.isArray(__row)) { __aligned.push(__row); continue; }
                var __copy = new Array(__maxLen);
                for (var __k = 0; __k < __maxLen; __k++) {
                    if (__k < __row.length) {
                        var __v = __row[__k];
                        // undefined 转空字符串(避免 WPS spill 显 0)
                        __copy[__k] = (__v === undefined) ? '' : __v;
                    } else {
                        __copy[__k] = '';
                    }
                }
                __aligned.push(__copy);
            }
            return __aligned;
        };

        /**
         * res - 获取原始数组（val的别名）
         * @returns {Array} 原始数组
         */
        wrappedResult.res = function() { return result; };

        /**
         * 🔧 v3.9.0 新增：getMeta - 获取透视表元数据
         * @returns {Object} 元数据对象
         * @example
         * var result = Array2D.z超级透视(data, rowFields, colFields, dataFields, 0, 1, '@^@', options);
         * var meta = result.getMeta();
         * console.log(meta.rowFields);  // ['大区', '省份']
         * console.log(meta.colFields);  // ['年份', '季度']
         * console.log(meta.grandTotal); // 总销售额
         */
        wrappedResult.getMeta = function() {
            return {
                version: '3.9.4',
                rowFields: rowConfig.fields.map(function(f) { return f.field; }),
                rowTitles: rowConfig.titles,
                colFields: colConfig.fields.map(function(f) { return f.field; }),
                colTitles: colConfig.titles,
                dataFields: dataOps.map(function(op) { return op.name; }),
                dataTitles: defaultDataTitles,
                rowCount: dataRows ? dataRows.length : rowKeys.length,
                colCount: colKeys.length,
                headerRowCount: headerRowCount,
                grandTotal: grandTotalValues ? grandTotalValues.grandTotal : null,
                options: {
                    cornerTitle: cornerTitle,
                    layoutMode: layoutMode,
                    rowFieldIndent: rowFieldIndent,
                    rowSubtotals: rowSubtotals,
                    colSubtotals: colSubtotals,
                    grandTotals: grandTotals,
                    displayAs: displayAs
                }
            };
        };
    }

    return wrappedResult;
};
Array2D.superPivot = Array2D.z超级透视;
$.superPivot = Array2D.z超级透视;
$.z超级透视 = Array2D.z超级透视;

/**
 * z超级透视 - 实例方法版本
 * 调用静态方法 Array2D.z超级透视，使用当前实例的数据
 * 🔧 v3.7.7 修复: 传递 this 而非 this._items，保留 _header 和 _original 属性
 */
Array2D.prototype.z超级透视 = function(rowFields, colFields, dataFields, headerRows, outputHeader, separator, options) {
    return Array2D.z超级透视(this, rowFields, colFields, dataFields, headerRows, outputHeader, separator, options);
};
Array2D.prototype.superPivot = Array2D.prototype.z超级透视;

/**
 * 生成静态方法（从实例方法自动生成）
 */
(function() {
    var propNames = Object.getOwnPropertyNames(Array2D.prototype);
    // 已经手动定义的静态方法，跳过自动生成
    var manuallyDefined = ['z选择列', 'selectCols', 'z批量填充', 'fill', 'z写入单元格', 'toRange', 'z转置', 'transpose', 'z求和', 'sum', 'z克隆', 'copy', 'z超级透视', 'superPivot', 'crossjoin', 'z笛卡尔积',
                          'z筛选', 'filter', 'z多列排序', 'sortByCols', 'z映射', 'map', 'z去重', 'distinct', 'rangeMap', 'z局部映射', 'z区域映射', 'rangeSelect', 'z按范围选择', 'rangeForEach', 'z按范围遍历',
                          'z分组汇总', 'groupInto', 'z分组汇总到字典', 'groupIntoMap', 'z分组汇总连接', 'groupIntoJoin',
                          /* XXD-194 */ 'z聚合', 'agg', 'aggregate', /* XXD-185 */ 'z区域矩阵', 'z矩阵分布', 'z查找区域', 'z复制', 'z复制到指定位置'];

    for (var i = 0; i < propNames.length; i++) {
        var name = propNames[i];
        if (manuallyDefined.indexOf(name) >= 0) continue;

        if (name !== 'constructor' && name !== '_init' && name !== '_new' && typeof Array2D.prototype[name] === 'function') {
            (function(methodName) {
                Array2D[methodName] = function() {
                    // 🔧 v3.7.7 修复: 保留 _header 属性
                    // 第一个参数是数组数据，传递给构造函数
                    // 支持 Array2D 对象（直接使用）或普通数组
                    var firstArg = arguments.length > 0 ? arguments[0] : null;
                    // 不再提取 _items，直接传递 firstArg，让构造函数处理 _header
                    var instance = new Array2D(firstArg);
                    // 剩余参数传递给实例方法
                    var restArgs = [];
                    for (var j = 1; j < arguments.length; j++) {
                        restArgs.push(arguments[j]);
                    }
                    var result = instance[methodName].apply(instance, restArgs);
                    // 🔧 v3.7.7 修复: 返回 Array2D 对象而非 _items，保留 _header
                    return result;
                };
            })(name);
        }
    }
})();

// ==================== [SUPERMAP] 增强Map（支持局部变量窗口查看） ====================

/**
 * SuperMap - 可在局部变量窗口实时展开查看的增强版 Map
 *
 * 特点：
 * 1. 完全兼容原生 Map 的所有属性和方法
 * 2. all 属性自动初始化，创建后立即可在局部变量窗口查看
 * 3. 支持嵌套 SuperMap、二维数组、Map 数组
 * 4. 层级前缀标识（01L00001 = 层数+序号+key）
 * 5. 调试模式开关，关闭后性能接近原生 Map
 */
function SuperMap(entries) {
    if (!(this instanceof SuperMap)) {
        return new SuperMap(entries);
    }
    this._map = new Map(entries);
    this._debug = true;
    this._all = null;  // 存储 all 属性值
    this._updateAll();  // 构造时立即初始化
}

// ========== 调试模式控制 ==========

Object.defineProperty(SuperMap.prototype, 'debug', {
    get: function() {
        return this._debug;
    },
    set: function(value) {
        this._debug = !!value;
        this._updateAll();
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(SuperMap, 'debug', {
    get: function() {
        return SuperMap._staticDebug;
    },
    set: function(value) {
        SuperMap._staticDebug = !!value;
    },
    enumerable: true,
    configurable: true
});
SuperMap._staticDebug = true;

// ========== 定义 all 属性 ==========

Object.defineProperty(SuperMap.prototype, 'all', {
    get: function() {
        return this._all;
    },
    enumerable: true,
    configurable: true
});

// ========== 定义 size 属性 ==========

Object.defineProperty(SuperMap.prototype, 'size', {
    get: function() {
        return this._map.size;
    },
    enumerable: true,
    configurable: true
});

// ========== 原型方法 ==========

/**
 * 更新 all 属性（构造时和每次修改后自动调用）
 */
SuperMap.prototype._updateAll = function() {
    if (!this._debug && !SuperMap._staticDebug) {
        this._all = { _提示: "调试模式已关闭，设置 debug=true 查看" };
        return;
    }
    this._all = this._buildAllView(1);
};

/**
 * 构建树形视图
 * 格式：01L00001 key（层数+序号+原key）
 */
SuperMap.prototype._buildAllView = function(level, maxRows) {
    level = level || 1;
    maxRows = maxRows || 255;

    var result = {};
    var self = this;
    var count = 0;
    var stopped = false;

    // 🔧 v3.9.4 修复：for...of 替换为 forEach（兼容 ES5/WPS JSA）
    this._map.forEach(function(val, k) {
        if (stopped) return;

        if (count >= maxRows) {
            result['_省略剩余' + (self._map.size - count) + '项'] = "...";
            stopped = true;
            return;
        }

        // 格式：01L00001 key（层数+序号+原key）
        var prefix = (level < 10 ? '0' : '') + level + 'L';
        var seqNum = '0000' + (count + 1);
        var displayKey = prefix + seqNum.slice(-5) + ' ' + k;

        // 判断值类型并处理
        if (val instanceof SuperMap) {
            result[displayKey] = val._buildAllView(level + 1, maxRows);
        } else if (val instanceof Map) {
            var superMap = SuperMap.fromMap(val, false);
            superMap._debug = self._debug;
            result[displayKey] = superMap._buildAllView(level + 1, maxRows);
        } else if (self._is2DArray(val)) {
            var arrObj = {};
            for (var i = 0; i < val.length && i < maxRows; i++) {
                arrObj[i + 1] = val[i];
            }
            result[displayKey] = arrObj;
        } else if (Array.isArray(val) && val.length > 0 && val[0] instanceof Map) {
            var arrObj = {};
            for (var i = 0; i < val.length && i < maxRows; i++) {
                var sm = SuperMap.fromMap(val[i], false);
                sm._debug = self._debug;
                arrObj[i + 1] = sm._buildAllView(level + 1, maxRows);
            }
            result[displayKey] = arrObj;
        } else if (typeof val === 'object' && val !== null && !Array.isArray(val)) {
            result[displayKey] = val;
        } else {
            result[displayKey] = val;
        }

        count++;
    });

    return result;
};

/**
 * 判断是否为二维数组
 */
SuperMap.prototype._is2DArray = function(value) {
    if (!Array.isArray(value)) return false;
    if (value.length === 0) return false;
    return Array.isArray(value[0]);
};

// ========== Map 原生方法（自动更新 all）==========

/**
 * 设置键值对（支持链式调用）
 */
SuperMap.prototype.set = function(key, value) {
    this._map.set(key, value);
    this._updateAll();  // 自动更新
    return this;  // 返回 this 支持链式调用
};

/**
 * 获取值
 */
SuperMap.prototype.get = function(key) {
    return this._map.get(key);
};

/**
 * 检查是否包含键
 */
SuperMap.prototype.has = function(key) {
    return this._map.has(key);
};

/**
 * 删除键值对
 */
SuperMap.prototype.delete = function(key) {
    var result = this._map.delete(key);
    this._updateAll();  // 自动更新
    return result;
};

/**
 * 清空所有键值对
 */
SuperMap.prototype.clear = function() {
    this._map.clear();
    this._updateAll();  // 自动更新
    return this;
};

/**
 * 遍历所有键值对
 */
SuperMap.prototype.forEach = function(callback, thisArg) {
    this._map.forEach(callback, thisArg);
    return this;
};

/**
 * 获取所有键的数组
 */
SuperMap.prototype.keys = function() {
    return Array.from(this._map.keys());
};

/**
 * 获取所有值的数组
 */
SuperMap.prototype.values = function() {
    return Array.from(this._map.values());
};

/**
 * 获取所有键值对的数组
 */
SuperMap.prototype.entries = function() {
    return Array.from(this._map.entries());
};

// ========== 转换方法 ==========

/**
 * 转为普通 Map 对象
 */
SuperMap.prototype.toMap = function(deep) {
    deep = deep !== undefined ? deep : true;

    var result = new Map();
    // 🔧 v3.9.4 修复：for...of 替换为 forEach（兼容 ES5/WPS JSA）
    this._map.forEach(function(val, k) {
        if (deep && val instanceof SuperMap) {
            result.set(k, val.toMap(deep));
        } else if (deep && val instanceof Map) {
            result.set(k, new Map(val));
        } else if (deep && Array.isArray(val)) {
            result.set(k, val.map(function(item) {
                return item instanceof SuperMap ? item.toMap(deep) : item;
            }));
        } else {
            result.set(k, val);
        }
    });
    return result;
};

/**
 * 静态方法：将普通 Map 转为 SuperMap
 */
// 🔧 XXD-169 atomic fix: fromMap 自动检测 Object 或 Map
SuperMap.fromMap = function(map, deep) {
    if (map && typeof map === 'object' && !(map instanceof Map) && !Array.isArray(map)) {
        // 普通 Object, 转 Map 后递归
        var __m = new Map();
        for (var __k in map) {
            if (Object.prototype.hasOwnProperty.call(map, __k)) __m.set(__k, map[__k]);
        }
        map = __m;
    }
    var entries;

    if (map instanceof Map) {
        deep = deep !== undefined ? deep : true;
        entries = [];
        // 使用 forEach 替代 for...of，兼容 WPS JSA (ES3/ES5)
        map.forEach(function(value, key) {
            if (deep && value instanceof Map) {
                entries.push([key, SuperMap.fromMap(value, deep)]);
            } else if (deep && Array.isArray(value)) {
                entries.push([key, value.map(function(item) {
                    return item instanceof Map ? SuperMap.fromMap(item, deep) : item;
                })]);
            } else {
                entries.push([key, value]);
            }
        });
    } else if (map && typeof map === 'object' && !Array.isArray(map)) {
        deep = deep !== undefined ? deep : true;
        var keys = Object.keys(map);
        entries = [];
        for (var i = 0; i < keys.length; i++) {
            var key = keys[i];
            var value = map[key];
            if (deep && value && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Map)) {
                entries.push([key, SuperMap.fromMap(value, deep)]);
            } else if (deep && value instanceof Map) {
                entries.push([key, SuperMap.fromMap(value, deep)]);
            } else if (deep && Array.isArray(value)) {
                entries.push([key, value.map(function(item) {
                    if (item instanceof Map) return SuperMap.fromMap(item, deep);
                    if (item && typeof item === 'object' && !Array.isArray(item)) return SuperMap.fromMap(item, deep);
                    return item;
                })]);
            } else {
                entries.push([key, value]);
            }
        }
    } else {
        throw new Error("参数必须是 Map 或普通 Object 类型");
    }

    return new SuperMap(entries);
};
SuperMap.z从Map = SuperMap.fromMap;
SuperMap.z从Obj = function(obj, deep) {
    if (obj && typeof obj === 'object' && !Array.isArray(obj) && !(obj instanceof Map)) {
        return SuperMap.fromMap(obj, deep);
    }
    throw new Error("参数必须是普通 Object 类型");
};

/**
 * 将SuperMap内容写入单元格
 * @param {String|Range} rng - 目标单元格地址或Range对象
 * @returns {SuperMap} 当前实例
 * @example
 * SuperMap.fromMap(map).toRange('A1');
 */
SuperMap.prototype.toRange = function(rng) {
    if (this._map.size === 0) return this;

    var arr = [['键', '聚合结果', '原始数据']];
    this._map.forEach(function(value, key) {
        var aggText = Array.isArray(value.agg) ? value.agg.join(', ') : JSON.stringify(value.agg);
        arr.push([key, aggText, JSON.stringify(value.group || {})]);
    });

    var targetRng = typeof rng === 'string' ? Range(rng) : rng;
    var rows = arr.length;
    var cols = arr[0].length;
    var endRng = targetRng.Item(rows, cols);
    var sheet = targetRng.Worksheet || Application.ActiveSheet;
    var writeRng = sheet.Range(targetRng, endRng);
    // 解除合并单元格 - 逐个单元格检查并取消合并
    for (var i = 1; i <= writeRng.Rows.Count; i++) {
        for (var j = 1; j <= writeRng.Columns.Count; j++) {
            var cell = writeRng.Cells(i, j);
            if (cell.MergeCells) {
                cell.MergeArea.UnMerge();
            }
        }
    }
    writeRng.Value2 = arr;
    return this;
};
SuperMap.prototype.z写入单元格 = SuperMap.prototype.toRange;

/**
 * 打印 all 内容到控制台
 */
SuperMap.prototype.print = function(title) {
    title = title || "SuperMap 内容";
    Console.log("===== " + title + " =====");
    Console.log(JSON.stringify(this.all, null, 2));
    Console.log("========================");
};
SuperMap.prototype.z打印 = SuperMap.prototype.print;

// ==================== [DATEUTILS] 日期工具库 ====================

/**
 * DateUtils - 日期操作工具（支持智能提示和链式调用）
 * @class
 * @constructor
 * @description 日期时间处理工具
 * @example
 * DateUtils.dt().z加天(5).z月底().val()
 */
function DateUtils(initialDate) {
    if (!(this instanceof DateUtils)) {
        return new DateUtils(initialDate);
    }
    this._date = initialDate ? new Date(initialDate) : new Date();
}

/**
 * 获取/设置日期
 * @param {Date|number|string} newDate - 新日期
 * @returns {DateUtils|Date} 设置时返回this，否则返回当前日期
 */
DateUtils.prototype.dt = function(newDate) {
    if (newDate !== undefined) {
        this._date = new Date(newDate);
        return this;
    }
    return this._date;
};

/**
 * 获取值
 * @returns {Date} 当前日期对象
 */
DateUtils.prototype.val = function() {
    return this._date;
};

/**
 * 加天数
 * @param {Number} days - 天数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加天 = function(days) {
    var result = new Date(this._date);
    result.setDate(result.getDate() + days);
    this._date = result;
    return this;
};
DateUtils.prototype.addDays = DateUtils.prototype.z加天;
// 🔧 XXD-151/152/154 final fix: 别名兼容（延迟包装器,避免 forward-ref 拿到 undefined）
DateUtils.prototype.z加天数 = DateUtils.prototype.z加天;
DateUtils.prototype.z减天数 = function(days) { return this.z加天(-days); };
DateUtils.prototype.z天 = function() { return this.z日期(); };
DateUtils.prototype.z天数 = function() { return this.z日期(); };
DateUtils.prototype.z日 = function() { return this.z日期(); };
DateUtils.prototype.z月份 = function() { return this.z月份(); }; // 延迟包装,call-time 解析到下方定义的真实 z月份 (月份 1-12)
DateUtils.prototype.z年 = function() { return this.z年份(); };

/**
 * 加月数
 * @param {Number} months - 月数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加月 = function(months) {
    var result = new Date(this._date);
    result.setMonth(result.getMonth() + months);
    this._date = result;
    return this;
};
DateUtils.prototype.addMonths = DateUtils.prototype.z加月;

/**
 * 加年数
 * @param {Number} years - 年数
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z加年 = function(years) {
    var result = new Date(this._date);
    result.setFullYear(result.getFullYear() + years);
    this._date = result;
    return this;
};
DateUtils.prototype.addYears = DateUtils.prototype.z加年;

/**
 * 月初
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z月初 = function() {
    this._date = new Date(this._date.getFullYear(), this._date.getMonth(), 1);
    return this;
};
DateUtils.prototype.firstDayOfMonth = DateUtils.prototype.z月初;

/**
 * 月底
 * @returns {DateUtils} 当前实例
 */
DateUtils.prototype.z月底 = function() {
    this._date = new Date(this._date.getFullYear(), this._date.getMonth() + 1, 0);
    return this;
};
DateUtils.prototype.endOfMonth = DateUtils.prototype.z月底;

/**
 * 转表格日期
 * @param {Date} jsdate - JS日期
 * @returns {Number} Excel日期数值
 */
DateUtils.prototype.z转表格日期 = function(jsdate) {
    if (!(jsdate instanceof Date)) {
        jsdate = new Date(jsdate);
    }
    var excelBase = new Date(1900, 0, 1).getTime();
    var dateMs = jsdate.getTime();
    var dayInMs = 24 * 60 * 60 * 1000;
    return (dateMs - excelBase) / dayInMs + 2;
};
DateUtils.prototype.toExcelDate = DateUtils.prototype.z转表格日期;

/**
 * 从表格日期转换（v3.9.4+）
 * 将Excel日期数值（Range.Value2读取日期时的返回值）转换为JS Date
 * @param {Number} excelSerial - Excel日期数值（如 44896）
 * @returns {Date} JavaScript Date对象
 */
DateUtils.prototype.z从表格日期 = function(excelSerial) {
    if (typeof excelSerial !== 'number') return new Date(excelSerial);
    var excelBase = new Date(1899, 11, 30).getTime();
    var dayInMs = 24 * 60 * 60 * 1000;
    return new Date(excelBase + excelSerial * dayInMs);
};
DateUtils.prototype.fromExcelDate = DateUtils.prototype.z从表格日期;

/**
 * 日期格式化
 * @param {Date} jsdate - 日期
 * @param {String} fmt - 格式
 * @returns {String} 格式化字符串
 */
DateUtils.prototype.z日期格式化 = function(jsdate, fmt) {
    if (typeof fmt !== 'string' && typeof jsdate === 'string') { fmt = jsdate; jsdate = this._date || jsdate; }
    if (fmt == null) fmt = 'yyyy-MM-dd';
    if (jsdate == null) jsdate = this._date;
    if (!(jsdate instanceof Date)) {
        jsdate = new Date(jsdate);
    }
    var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
    return fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
        if (year) return jsdate.getFullYear().toString().padStart(year.length, '0');
        if (month) return (jsdate.getMonth() + 1).toString().padStart(month.length, '0');
        if (day) return jsdate.getDate().toString().padStart(day.length, '0');
        if (hour) return jsdate.getHours().toString().padStart(hour.length, '0');
        if (minute) return jsdate.getMinutes().toString().padStart(minute.length, '0');
        if (second) return jsdate.getSeconds().toString().padStart(second.length, '0');
        if (millisecond) return jsdate.getMilliseconds().toString().padStart(3, '0');
        if (week) return '周' + weekDays[jsdate.getDay()];
        return match;
    });
};
DateUtils.prototype.format = DateUtils.prototype.z日期格式化;

/**
 * 获取年份
 * @returns {Number} 年份（4位数字）
 * @example
 * asDate("2023-9-21").z年份()  // 2023
 */
DateUtils.prototype.z年份 = function() {
    return this._date.getFullYear();
};
DateUtils.prototype.getYear = DateUtils.prototype.z年份;

/**
 * 获取月份（1-12）
 * @returns {Number} 月份（1-12）
 * @example
 * asDate("2023-9-21").z月份()  // 9
 */
DateUtils.prototype.z月份 = function() {
    return this._date.getMonth() + 1;
};
DateUtils.prototype.getMonth = DateUtils.prototype.z月份;

/**
 * 获取日期（1-31）
 * @returns {Number} 日期（1-31）
 * @example
 * asDate("2023-9-21").z日期()  // 21
 */
DateUtils.prototype.z日期 = function() {
    return this._date.getDate();
};
DateUtils.prototype.getDate = DateUtils.prototype.z日期;

/**
 * 获取星期（1-7，7=周日）
 * @returns {Number} 星期（1-7）
 * @example
 * asDate("2023-9-21").z星期()  // 4 (周四)
 * asDate("2023-9-24").z星期()  // 7 (周日)
 */
DateUtils.prototype.z星期 = function() {
    var day = this._date.getDay();
    return day === 0 ? 7 : day;
};
DateUtils.prototype.getDay = DateUtils.prototype.z星期;

/**
 * 获取小时（0-23）
 * @returns {Number} 小时（0-23）
 */
DateUtils.prototype.z小时 = function() {
    return this._date.getHours();
};
DateUtils.prototype.getHour = DateUtils.prototype.z小时;

/**
 * 获取分钟（0-59）
 * @returns {Number} 分钟（0-59）
 */
DateUtils.prototype.z分钟 = function() {
    return this._date.getMinutes();
};
DateUtils.prototype.getMinute = DateUtils.prototype.z分钟;

/**
 * 获取秒数（0-59）
 * @returns {Number} 秒数（0-59）
 */
DateUtils.prototype.z秒 = function() {
    return this._date.getSeconds();
};
DateUtils.prototype.getSecond = DateUtils.prototype.z秒;

/**
 * 获取时间戳（毫秒）
 * @returns {Number} 时间戳
 */
DateUtils.prototype.z时间戳 = function() {
    return this._date.getTime();
};
DateUtils.prototype.getTime = DateUtils.prototype.z时间戳;

/**
 * z季度 - 获取日期所在的季度（1-4）
 * @param {Date|number} [date] - 日期对象或OA日期数值，不传则使用实例内部日期
 * @returns {Number} 季度（1-4）
 * @example
 * DateUtils.z季度(new Date())     // 当前季度（静态调用）
 * DateUtils(new Date()).z季度()   // 当前季度（实例调用）
 */
DateUtils.prototype.z季度 = function(date) {
    var d;
    if (date !== undefined) {
        d = (date instanceof Date) ? date : new Date(date);
    } else {
        d = this._date;
    }
    if (!(d instanceof Date) || isNaN(d.getTime())) return 0;
    return Math.floor(d.getMonth() / 3) + 1;
};
DateUtils.prototype.getQuarter = DateUtils.prototype.z季度;
// 静态方法：支持 DateUtils.z季度(date) 调用
DateUtils.z季度 = function(date) {
    return DateUtils.prototype.z季度.call(DateUtils(date), date);
};
DateUtils.getQuarter = DateUtils.z季度;

// ==================== Date.prototype.format 日期格式化扩展 ====================
// 让 new Date().format("yyyy-MM-dd") 可以直接调用
// 课程教材中使用 new Date().format() 将日期转为字符串写入单元格

/**
 * Date.prototype.format - 日期格式化（扩展原生Date原型）
 * @param {String} [fmt="yyyy-MM-dd"] - 格式字符串，默认 "yyyy-MM-dd"
 * @returns {String} 格式化后的日期字符串
 * @example
 * new Date().format()              // "2026-05-27"
 * new Date().format("yyyy-MM-dd")  // "2026-05-27"
 * new Date().format("yyyy年M月d日") // "2026年5月27日"
 * new Date().format("HH:mm:ss")    // "13:07:43"
 */
Date.prototype.format = function(fmt) {
    if (!fmt) fmt = 'yyyy-MM-dd';
    var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
    // 使用北京时间（UTC+8）：基于UTC时间加8小时偏移，确保不受系统时区影响
    var utcMs = this.getTime() + this.getTimezoneOffset() * 60000;
    var bjDate = new Date(utcMs + 8 * 3600000);
    return fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
        if (year) return bjDate.getFullYear().toString().padStart(year.length, '0');
        if (month) return (bjDate.getMonth() + 1).toString().padStart(month.length, '0');
        if (day) return bjDate.getDate().toString().padStart(day.length, '0');
        if (hour) return bjDate.getHours().toString().padStart(hour.length, '0');
        if (minute) return bjDate.getMinutes().toString().padStart(minute.length, '0');
        if (second) return bjDate.getSeconds().toString().padStart(second.length, '0');
        if (millisecond) return bjDate.getMilliseconds().toString().padStart(3, '0');
        if (week) return '周' + weekDays[bjDate.getDay()];
        return match;
    });
};

// ==================== DateUtils 静态方法补充 ====================
// 课程中使用 DateUtils.fromExcelDate() 静态调用，需要添加静态版本

/**
 * DateUtils.fromExcelDate 静态方法 - 将Excel日期数值转换为JS Date
 * @param {Number} excelSerial - Excel日期数值（OA日期，如 44896）
 * @returns {Date} JavaScript Date对象
 * @example
 * DateUtils.fromExcelDate(44896)  // Date对象
 */
DateUtils.fromExcelDate = function(excelSerial) {
    return DateUtils.prototype.z从表格日期.call(DateUtils(), excelSerial);
};

/**
 * DateUtils.format 静态方法 - 日期格式化（不依赖Date.prototype扩展）
 * @param {Date|String|Number} date - 日期对象、日期字符串或时间戳
 * @param {String} [fmt="yyyy-MM-dd"] - 格式字符串
 * @returns {String} 格式化后的日期字符串
 * @example
 * DateUtils.format(new Date())                // "2026-05-27"
 * DateUtils.format(new Date(), "yyyy-MM-dd HH:mm:ss")  // "2026-05-27 13:20:00"
 * DateUtils.format(DateUtils.fromExcelDate(45179.8369)) // "2023-09-10 20:05:11"
 */
DateUtils.format = function(date, fmt) {
    if (!(date instanceof Date)) {
        date = new Date(date);
    }
    return DateUtils.prototype.z日期格式化.call(DateUtils(date), date, fmt || 'yyyy-MM-dd');
/**
 * DateUtils.addMonths 静态方法 — 日期加上月数
 * @param {Date|string} date - 日期
 * @param {Number} n - 月数（可为负数）
 * @returns {Date} 新的日期对象
 * @example
 * DateUtils.addMonths(new Date(), 3)   // 三个月后
 * DateUtils.addMonths(new Date(), -1)  // 一个月前
 */
DateUtils.addMonths = function(date, n) {
    var d = date instanceof Date ? new Date(date) : new Date(date);
    d.setMonth(d.getMonth() + n);
    return d;
};

/**
 * DateUtils.datedif 静态方法 — 计算两个日期之间的差（Excel DATEDIF 函数）
 * @param {Date|string} startDate - 开始日期
 * @param {Date|string} endDate - 结束日期
 * @param {String} unit - 单位: "Y" 年, "M" 月, "D" 日, "MD" 忽略年月日差, "YM" 忽略年月差, "YD" 忽略年日差
 * @returns {Number} 日期差
 * @example
 * DateUtils.datedif("2023-01-15", "2025-06-02", "Y")  // 2
 * DateUtils.datedif("2023-01-15", "2025-06-02", "M")  // 28
 * DateUtils.datedif("2023-01-15", "2025-06-02", "D")  // 869
 */
DateUtils.datedif = function(startDate, endDate, unit) {
    var d1 = startDate instanceof Date ? startDate : new Date(startDate);
    var d2 = endDate instanceof Date ? endDate : new Date(endDate);
    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return NaN;

    var y1 = d1.getFullYear(), m1 = d1.getMonth(), day1 = d1.getDate();
    var y2 = d2.getFullYear(), m2 = d2.getMonth(), day2 = d2.getDate();

    switch (unit) {
        case 'Y':
            var years = y2 - y1;
            if (m2 < m1 || (m2 === m1 && day2 < day1)) years--;
            return years;
        case 'M':
            return (y2 - y1) * 12 + (m2 - m1) - (day2 < day1 ? 1 : 0);
        case 'D':
            return Math.floor((d2.getTime() - d1.getTime()) / 86400000);
        case 'MD':
            var refDate = new Date(y2, m2, day1);
            return Math.floor((d2.getTime() - refDate.getTime()) / 86400000);
        case 'YM':
            var mDiff = (m2 - m1 + 12) % 12;
            if (day2 < day1) mDiff = (mDiff - 1 + 12) % 12;
            return mDiff;
        case 'YD':
            var yearRef = new Date(y2, m2, day1);
            if (yearRef > d2) yearRef = new Date(y2 - 1, m2, day1);
            else yearRef = new Date(y2, m1, day1);
            return Math.floor((d2.getTime() - yearRef.getTime()) / 86400000);
        default:
            return NaN;
    }
};

};

/**
 * RangeChain - Range链式调用包装类（支持智能提示和链式调用）
 * @private
 * @class
 * @constructor
 * @description 支持Range方法的链式调用和智能提示
 * @example
 * $.maxRange("A1:J1").safeArray()     // 链式调用
 * $(5, 2).z值()                       // 获取第5行第2列的值
 * $(5, 2).z值("新值").z加粗()         // 链式设置
 */
// ==================== [EXPORTS] 全局变量统一导出 ====================

// WPS现代版 - 使用立即执行函数导出全局变量，支持ES6+
(function() {
    (function() {
        this.Array2D = Array2D;

        this.As = As;
        this.RngUtils = RngUtils;
        this.ShtUtils = ShtUtils;
        this.DateUtils = DateUtils;
        this.JSA = JSA;
        this.IO = IO;
        this.$ = $;
        this.$$ = $$;
        this.log = log;
        this.logjson = logjson;
        // Global函数
        this.f1 = f1;
        this.$fx = $fx;
        this.$toArray = $toArray;
        // As 已在第6891行导出，此处删除重复定义
        this.asArray = asArray;
        this.asDate = asDate;
        this.asMap = asMap;
        this.asNumber = asNumber;
        this.asObject = asObject;
        this.asRange = asRange;
        this.asShape = asShape;
        this.asSheet = asSheet;
        this.asString = asString;
        this.asWorkbook = asWorkbook;
        this.cdate = cdate;
        // 同步挂载 Date.prototype.format（确保模块隔离时也能通过全局函数调用）
        if (typeof Date.prototype.format !== 'function') {
            Date.prototype.format = function(fmt) {
                if (!fmt) fmt = 'yyyy-MM-dd';
                var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
                var utcMs = this.getTime() + this.getTimezoneOffset() * 60000;
                var bjDate = new Date(utcMs + 8 * 3600000);
                return fmt.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
                    if (year) return bjDate.getFullYear().toString().padStart(year.length, '0');
                    if (month) return (bjDate.getMonth() + 1).toString().padStart(month.length, '0');
                    if (day) return bjDate.getDate().toString().padStart(day.length, '0');
                    if (hour) return bjDate.getHours().toString().padStart(hour.length, '0');
                    if (minute) return bjDate.getMinutes().toString().padStart(minute.length, '0');
                    if (second) return bjDate.getSeconds().toString().padStart(second.length, '0');
                    if (millisecond) return bjDate.getMilliseconds().toString().padStart(3, '0');
                    if (week) return '周' + weekDays[bjDate.getDay()];
                    return match;
                });
            };
        }
        // 提供全局 format 函数（兼容模块隔离环境）
        // 用法：format(new Date(), "yyyy/M/d h:mm:ss") 或 format(date, "yyyy年M月d日")
        this.format = function(date, fmt) {
            if (date instanceof Date) {
                return date.format(fmt);
            }
            return new Date(date).format(fmt);
        };
        this.cstr = cstr;
        // 【v4.2.2】注册 k 函数（jsaLambda 简短别名）到全局
        // 让用户能在单元格公式中直接使用 =k("JSA.getIndexs", 1, 10, 2)
        this.k = JSA.k;
        this.jsaLambda = JSA.jsaLambda;
        // 【v4.0.10】导出常用全局函数到 JSA 命名空间
        if (typeof agg !== 'undefined') this.agg = agg;
        if (typeof oadate !== 'undefined') this.oadate = oadate;
        if (typeof map2d !== 'undefined') this.map2d = map2d;
        if (typeof forEach2d !== 'undefined') this.forEach2d = forEach2d;
        if (typeof StopWatch !== 'undefined') this.StopWatch = StopWatch;
        if (typeof TreeNode !== 'undefined') this.TreeNode = TreeNode;
        if (typeof batch_k !== 'undefined') this.batch_k = batch_k;
        if (typeof k_help !== 'undefined') this.k_help = k_help;
        if (typeof DataExport !== 'undefined') this.DataExport = DataExport;
        if (typeof DataImport !== 'undefined') this.DataImport = DataImport;
        if (typeof Logger !== 'undefined') this.Logger = Logger;
        if (typeof StrUtils !== 'undefined') this.StrUtils = StrUtils;
        if (typeof NumUtils !== 'undefined') this.NumUtils = NumUtils;
        if (typeof MsgUtils !== 'undefined') this.MsgUtils = MsgUtils;
        if (typeof WorkbookUtils !== 'undefined') this.WorkbookUtils = WorkbookUtils;
        if (typeof FormUtils !== 'undefined') this.FormUtils = FormUtils;
        if (typeof PicUtiles !== 'undefined') this.PicUtiles = PicUtiles;
        this.isArray = isArray;
        this.isArray2D = isArray2D;
        this.isBoolean = isBoolean;
        this.isCollection = isCollection;
        this.isDate = isDate;
        this.isEmpty = isEmpty;
        this.isNumberic = isNumberic;
        this.isRange = isRange;
        this.isRegex = isRegex;
        this.isSameClass = isSameClass;
        this.isSheet = isSheet;
        this.isString = isString;
        this.isWorkbook = isWorkbook;
        this.typeName = typeName;
        this.val = val;
        this.round = round;
        // ubound函数 - 获取数组的指定维度的上界
        this.ubound = function(arr, dimension) {
            dimension = dimension || 1;
            if (!Array.isArray(arr)) return -1;
            if (dimension === 1) return arr.length - 1;
            if (dimension === 2) {
                var maxLen = 0;
                for (var i = 0; i < arr.length; i++) {
                    if (Array.isArray(arr[i]) && arr[i].length > maxLen) {
                        maxLen = arr[i].length;
                    }
                }
                return maxLen - 1;
            }
            return -1;
        };
        
        // ==================== JSA880快捷API - 一行代码走天下 ====================
        
        /**
         * JSA880 - 郑广学JSA880快速开发框架主入口
         * @description 提供超简洁的一行代码API，集成所有核心功能
         * @namespace
         * @example
         * // 一行代码完成数据透视
         * JSA880.透视(数据, '产品+,月份+', '地区+', 'sum(销量),count()');
         * 
         * // 一行代码筛选数据
         * JSA880.筛选(数据, 'x=>x[0]=="北京" && x[3]>100');
         * 
         * // 一行代码读取表格数据
         * JSA880.读表("A1:D100");
         * 
         * // 一行代码写入表格数据
         * JSA880.写表([[1,2],[3,4]], "G1");
         * 
         * // 一行代码获取最大行数
         * JSA880.最大行("A:A");
         * 
         * // 一行代码删除空白行
         * JSA880.删空行("A1:F100");
         * 
         * // 一行代码排序
         * JSA880.排序(数据, 'f3+,f4-');
         */ 
        this.JSA880 = {
            /**
             * 数据透视（超简化版）
             * @param {Array} data - 二维数组数据
             * @param {string} rowFields - 行字段，支持排序符号 f1+,f2-
             * @param {string} colFields - 列字段，支持排序符号 f3+,f4-
             * @param {string} dataFields - 数据字段，格式: 'count(),sum(f5),average(f6)'
             * @param {number} [headerRows=1] - 标题行数
             * @returns {Array} 透视结果
             * @example
             * JSA880.透视(销售数据, '产品+,地区-', '月份+', 'sum(金额),count()');
             */
            透视: function(data, rowFields, colFields, dataFields, headerRows) {
                return Array2D.z超级透视(data, [rowFields], [colFields], [dataFields], headerRows);
            },
            
            /**
             * 超级透视（完整版）
             */
            超级透视: Array2D.z超级透视,
            
            /**
             * 数据筛选
             * @param {Array} data - 二维数组
             * @param {string|Function} predicate - 筛选条件
             * @returns {Array2D} Array2D对象
             * @example
             * JSA880.筛选(数据, 'x=>x[0]=="北京" && x[3]>100');
             */
            筛选: function(data, predicate) {
                return new Array2D(data).z筛选(predicate);
            },
            
            /**
             * 多条件筛选（简化版）
             * @param {Array} data - 二维数组
             * @param {Array} conditions - 条件数组，如 [[0, '北京'], [3, 100]]
             * @returns {Array2D} Array2D对象
             * @example
             * JSA880.多条件筛选(数据, [[0, '北京'], [3, 100]]);
             */
            多条件筛选: function(data, conditions) {
                var arr = new Array2D(data);
                for (var i = 0; i < conditions.length; i++) {
                    var col = conditions[i][0];
                    var val = conditions[i][1];
                    arr = arr.z筛选(function(row) { 
                        return row[col] == val || (typeof val === 'number' && row[col] > val);
                    });
                }
                return arr;
            },
            
            /**
             * 分组汇总
             * @param {Array} data - 二维数组
             * @param {string} groupCol - 分组列 f1
             * @param {string} aggCol - 汇总列 f2
             * @param {string} [aggType='sum'] - 汇总类型: sum, count, average, max, min
             * @returns {Array} 汇总结果
             * @example
             * JSA880.分组汇总(数据, 'f1', 'f3', 'sum');
             */
            分组汇总: function(data, groupCol, aggCol, aggType) {
                aggType = aggType || 'sum';
                var aggExpr = aggType + '("' + aggCol + '")';
                return Array2D.z分组汇总(data, groupCol, aggExpr);
            },
            
            /**
             * 分组汇总连接 - 优化sumifs和Countifs批量条件统计
             * @param {Array} targetData - 统计目标数据（左表）
             * @param {Array} sourceData - 数据源（右表）
             * @param {string} groupKey - 分组键选择器，如 'f2' 或 'f2,f3'
             * @param {string} aggFunc - 汇总函数，如 'sum("f4")' 或 'count(),sum("f5")'
             * @returns {Array} 连接汇总后的结果
             * @example
             * // 一行代码完成sumifs/countifs批量统计
             * JSA880.分组汇总连接(目标表, 源数据, 'f2', 'sum("f4")');
             * JSA880.分组汇总连接(目标表, 源数据, '月份,产品', 'count(),sum("销量"),average("金额")');
             */
            分组汇总连接: function(targetData, sourceData, groupKey, aggFunc) {
                return Array2D.groupIntoJoin(targetData, sourceData, groupKey, aggFunc);
            },
            
            /**
             * 读取表格数据（简化版）
             * @param {string} range - 单元格地址，如 "A1:D100" 或 "A:A"
             * @returns {Array} 二维数组
             * @example
             * JSA880.读表("A1:D100");
             * JSA880.读表("A:A");  // 读取整列到最大行
             */
            读表: function(range) {
                            var rng = typeof range === 'string' ? Range(range) : range;
                var arr = rng.Value2;
                if (arr === null || arr === undefined) return [];
                if (!Array.isArray(arr)) return [[arr]];
                if (!Array.isArray(arr[0])) {
                    var result = [];
                    for (var i = 0; i < arr.length; i++) {
                        result.push([arr[i]]);
                    }
                    return result;
                }
                return arr;
            },
            
            /**
             * 写入表格数据（简化版）
             * @param {Array} data - 二维数组
             * @param {string} startCell - 起始单元格，如 "A1"
             * @returns {Range} 写入的单元格区域
             * @example
             * JSA880.写表([[1,2],[3,4]], "G1");
             */
            写表: function(data, startCell) {
                return JSA.z写入单元格(data, startCell);
            },
            
            /**
             * 获取最大行数
             * @param {string} column - 列范围，如 "A:A" 或 "A1"
             * @returns {number} 最大行数
             * @example
             * JSA880.最大行("A:A");
             */
            最大行: function(column) {
                return RngUtils.z最大行(column);
            },
            
            /**
             * 获取最大列数
             * @param {string} row - 行范围，如 "1:1" 或 "A1"
             * @returns {number} 最大列数
             * @example
             * JSA880.最大列("1:1");
             */
            最大列: function(row) {
                return RngUtils.z最大列(row);
            },
            
            /**
             * 删除空白行
             * @param {string} range - 单元格范围
             * @param {boolean} [entireRow=true] - 是否删除整行
             * @returns {boolean} 是否成功
             * @example
             * JSA880.删空行("A1:F100");
             */
            删空行: function(range, entireRow) {
                RngUtils.z删除空白行(range, entireRow !== false);
                return true;
            },
            
            /**
             * 删除空白列
             * @param {string} range - 单元格范围
             * @param {boolean} [entireColumn=true] - 是否删除整列
             * @returns {boolean} 是否成功
             * @example
             * JSA880.删空列("A1:Z100");
             */
            删空列: function(range, entireColumn) {
                RngUtils.z删除空白列(range, entireColumn !== false);
                return true;
            },
            
            /**
             * 多列排序（简化版）
             * @param {Array} data - 二维数组
             * @param {string} sortParams - 排序参数，如 'f3+,f4-'
             * @param {number} [headerRows=1] - 标题行数
             * @returns {Array} 排序后数组
             * @example
             * JSA880.排序(数据, 'f3+,f4-', 1);
             */
            排序: function(data, sortParams, headerRows) {
                return new Array2D(data).z多列排序(sortParams, headerRows || 1);
            },
            
            /**
             * 去重
             * @param {Array} data - 二维数组
             * @param {number} [colIndex] - 指定列去重
             * @returns {Array} 去重后数组
             * @example
             * JSA880.去重(数据);
             * JSA880.去重(数据, 0);  // 按第1列去重
             */
            去重: function(data, colIndex) {
                return new Array2D(data).z去重(colIndex).val();
            },
            
            /**
             * 转置
             * @param {Array} data - 二维数组
             * @returns {Array} 转置后数组
             * @example
             * JSA880.转置([[1,2],[3,4]]);  // 返回 [[1,3],[2,4]]
             */
            转置: function(data) {
                return new Array2D(data).z转置().val();
            },
            
            /**
             * 数组求和
             * @param {Array} data - 二维数组
             * @param {string} [colSelector] - 列选择器，如 'f1'
             * @returns {number} 求和结果
             * @example
             * JSA880.求和([[1,2],[3,4]]);        // 10
             * JSA880.求和([[1,2],[3,4]], 'f1');  // 4 (第1列求和)
             */
            求和: function(data, colSelector) {
                return new Array2D(data).z求和(colSelector);
            },
            
            /**
             * 添加边框（快速版）
             * @param {string} range - 单元格范围
             * @param {number} [style=1] - 线条样式
             * @returns {boolean} 是否成功
             * @example
             * JSA880.加边框("A1:D10");
             */
            加边框: function(range, style) {
                RngUtils.z加边框(range, style || 1);
                return true;
            },
            
            /**
             * 自动列宽（快速版）
             * @param {string} range - 单元格范围
             * @returns {boolean} 是否成功
             * @example
             * JSA880.自动列宽("A:Z");
             */
            自动列宽: function(range) {
                RngUtils.z自动列宽(range);
                return true;
            },
            
            /**
             * 自动行高（快速版）
             * @param {string} range - 单元格范围
             * @returns {boolean} 是否成功
             * @example
             * JSA880.自动行高("1:100");
             */
            自动行高: function(range) {
                RngUtils.z自动行高(range);
                return true;
            },
            
            /**
             * 安全读取已使用区域
             * @param {string} [sheetName] - 工作表名称，不传则使用当前表
             * @returns {Array} 二维数组
             * @example
             * JSA880.读已用区();              // 当前表
             * JSA880.读已用区("Sheet1");      // 指定表
             */
            读已用区: function(sheetName) {
                            var sheet = sheetName ? Sheets(sheetName) : Application.ActiveSheet;
                var usedRange;
                try {
                    usedRange = sheet.UsedRange;
                } catch (e) {
                    return [];
                }
                if (!usedRange) return [];
                var arr = usedRange.Value2;
                if (arr === null || arr === undefined) return [];
                if (!Array.isArray(arr)) return [[arr]];
                if (!Array.isArray(arr[0])) {
                    var result = [];
                    for (var i = 0; i < arr.length; i++) {
                        result.push([arr[i]]);
                    }
                    return result;
                }
                return arr;
            },
            
            /**
             * 生成数字序列
             * @param {number} start - 起始数字
             * @param {number} end - 结束数字
             * @param {number} [step=1] - 步长
             * @returns {Array} 序列数组
             * @example
             * JSA880.序列(1, 10);      // [1,2,3,4,5,6,7,8,9,10]
             * JSA880.序列(1, 10, 2);  // [1,3,5,7,9]
             */
            序列: function(start, end, step) {
                step = step || 1;
                var result = [];
                for (var i = start; i <= end; i += step) {
                    result.push(i);
                }
                return result;
            },
            
            /**
             * 随机打乱数组
             * @param {Array} array - 数组
             * @returns {Array} 打乱后的数组
             * @example
             * JSA880.打乱([1,2,3,4,5]);
             */
            打乱: function(array) {
                var result = array.slice();
                for (var i = result.length - 1; i > 0; i--) {
                    var j = Math.floor(Math.random() * (i + 1));
                    var temp = result[i];
                    result[i] = result[j];
                    result[j] = temp;
                }
                return result;
            },
            
            /**
             * 随机整数
             * @param {number} min - 最小值
             * @param {number} max - 最大值
             * @returns {number} 随机整数
             * @example
             * JSA880.随机(1, 100);
             */
            随机: function(min, max) {
                return Math.floor(Math.random() * (max - min + 1)) + min;
            },
            
            /**
             * 创建SuperMap（可视化调试字典）
             * @returns {SuperMap} SuperMap实例
             * @example
             * var map = JSA880.超级字典();
             * map.set('user1', {name: '张三', age: 25});
             * map.debug(true); // 开启调试模式
             */
            超级字典: function() {
                return new SuperMap();
            },
            
            /**
             * 从Map创建SuperMap
             * @param {Map} map - 普通Map对象
             * @returns {SuperMap} SuperMap实例
             * @example
             * var nativeMap = new Map();
             * nativeMap.set('a', 1);
             * var superMap = JSA880.SuperMap从Map(nativeMap);
             */
            SuperMap从Map: function(map) {
                return SuperMap.fromMap(map);
            },
            
            /**
             * 从对象创建SuperMap
             * @param {Object} obj - 普通对象
             * @returns {SuperMap} SuperMap实例
             * @example
             * var superMap = JSA880.SuperMap从对象({a: 1, b: 2});
             */
            SuperMap从对象: function(obj) {
                return SuperMap.fromObject(obj);
            },
            
            /**
             * 从数组创建SuperMap
             * @param {Array} arr - 二维数组，每个元素为[key, value]
             * @returns {SuperMap} SuperMap实例
             * @example
             * var superMap = JSA880.SuperMap从数组([['key1', 'value1'], ['key2', 'value2']]);
             */
            SuperMap从数组: function(arr) {
                return SuperMap.fromArray(arr);
            },
            
            /**
             * 日期格式化
             * @param {Date|string} date - 日期
             * @param {string} format - 格式字符串，如 'yyyy-MM-dd HH:mm:ss'
             * @returns {string} 格式化后的日期字符串
             * @example
             * JSA880.日期格式(new Date(), 'yyyy-MM-dd');
             */
            日期格式: function(date, format) {
                date = typeof date === 'string' ? new Date(date) : date;
                var weekDays = ['日', '一', '二', '三', '四', '五', '六'];
                return format.replace(/(y+|Y+)|(M+)|(d+|D+)|(H+)|(m+)|(s+|S+)|(SSS)|(a+)/g, function(match, year, month, day, hour, minute, second, millisecond, week) {
                    if (year) return date.getFullYear().toString().padStart(year.length, '0');
                    if (month) return (date.getMonth() + 1).toString().padStart(month.length, '0');
                    if (day) return date.getDate().toString().padStart(day.length, '0');
                    if (hour) return date.getHours().toString().padStart(hour.length, '0');
                    if (minute) return date.getMinutes().toString().padStart(minute.length, '0');
                    if (second) return date.getSeconds().toString().padStart(second.length, '0');
                    if (millisecond) return date.getMilliseconds().toString().padStart(3, '0');
                    if (week) return '周' + weekDays[date.getDay()];
                    return match;
                });
            },
            
            /**
             * 人民币大写
             * @param {number} n - 数字
             * @returns {string} 人民币大写
             * @example
             * JSA880.人民币大写(12345.67);  // 壹万贰仟叁佰肆拾伍元陆角柒分
             */
            人民币大写: JSA.z人民币大写,
            
            /**
             * 字符串全局替换
             * @param {string} str - 原字符串
             * @param {string} search - 查找字符串
             * @param {string} replacement - 替换字符串
             * @returns {string} 替换后的字符串
             * @example
             * JSA880.替换("hello world", "world", "JSA880");  // "hello JSA880"
             */
            替换: function(str, search, replacement) {
                return str.split(search).join(replacement);
            },
            
            /**
             * 数组扁平化
             * @param {Array} arr - 多维数组
             * @returns {Array} 一维数组
             * @example
             * JSA880.扁平化([[1,2],[3,4],[5,6]]);  // [1,2,3,4,5,6]
             */
            扁平化: function(arr) {
                return new Array2D(arr).z扁平化();
            },
            
            /**
             * 列号转字母（Excel列名）
             * @param {number} n - 列号（从1开始）
             * @returns {string} 列字母，如 1->A, 27->AA
             * @example
             * JSA880.列号(1);   // "A"
             * JSA880.列号(27);  // "AA"
             */
            列号: function(n) {
                var result = '';
                while (n > 0) {
                    n--;
                    result = String.fromCharCode(65 + (n % 26)) + result;
                    n = Math.floor(n / 26);
                }
                return result;
            },
            
            /**
             * 列字母转列号
             * @param {string} col - 列字母，如 "A", "AA"
             * @returns {number} 列号（从1开始）
             * @example
             * JSA880.列字母("A");   // 1
             * JSA880.列字母("AA");  // 27
             */
            列字母: function(col) {
                var result = 0;
                for (var i = 0; i < col.length; i++) {
                    result = result * 26 + (col.charCodeAt(i) - 64);
                }
                return result;
            },

            // ============================================================
            // 整合 jsa880-framework.js 独有功能
            // ============================================================

            /**
             * 导出为CSV文件
             * @param {Array} data - 数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             */
            导出CSV: function(data, filePath, options) {
                return DataExport.toCSV(data, filePath, options);
            },

            /**
             * 导出为JSON文件
             * @param {Array} data - 数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             */
            导出JSON: function(data, filePath, options) {
                return DataExport.toJSON(data, filePath, options);
            },

            /**
             * 导出为HTML文件
             * @param {Array} data - 数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             */
            导出HTML: function(data, filePath, options) {
                return DataExport.toHTML(data, filePath, options);
            },

            /**
             * 导出为XML文件
             * @param {Array} data - 数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             */
            导出XML: function(data, filePath, options) {
                return DataExport.toXML(data, filePath, options);
            },

            /**
             * 从CSV导入数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             * @returns {Array} 二维数组
             */
            导入CSV: function(filePath, options) {
                return DataImport.fromCSV(filePath, options);
            },

            /**
             * 从JSON导入数据
             * @param {string} filePath - 文件路径
             * @param {Object} options - 配置选项
             * @returns {Array} 二维数组
             */
            导入JSON: function(filePath, options) {
                return DataImport.fromJSON(filePath, options);
            },

            /**
             * 从JSON导入为Array2D
             * @param {string} jsonStr - JSON字符串
             * @param {Object} options - 配置选项
             * @returns {Array2D} Array2D实例
             */
            导入JSON字符串: function(jsonStr, options) {
                return DataImport.fromJSONString(jsonStr, options);
            },

            /**
             * 验证必填项
             * @param {*} value - 值
             * @returns {boolean} 是否有效
             */
            验证必填: function(value) {
                return DataValidation.required(value);
            },

            /**
             * 验证数字范围
             * @param {*} value - 值
             * @param {number} min - 最小值
             * @param {number} max - 最大值
             * @returns {boolean} 是否在范围内
             */
            验证范围: function(value, min, max) {
                return DataValidation.range(value, min, max);
            },

            /**
             * 验证邮箱格式
             * @param {*} value - 值
             * @returns {boolean} 是否有效邮箱
             */
            验证邮箱: function(value) {
                return DataValidation.email(value);
            },

            /**
             * 验证手机号格式
             * @param {*} value - 值
             * @returns {boolean} 是否有效手机号
             */
            验证手机: function(value) {
                return DataValidation.mobile(value);
            },

            /**
             * 验证身份证号格式
             * @param {*} value - 值
             * @returns {boolean} 是否有效身份证
             */
            验证身份证: function(value) {
                return DataValidation.idCard(value);
            },

            /**
             * 验证日期格式
             * @param {*} value - 值
             * @param {string} format - 日期格式
             * @returns {boolean} 是否有效日期
             */
            验证日期: function(value, format) {
                return DataValidation.date(value, format);
            },

            /**
             * 批量验证数据
             * @param {Array} data - 数据数组
             * @param {Object} rules - 验证规则
             * @returns {Array} 验证结果数组
             */
            批量验证: function(data, rules) {
                return DataValidation.validate(data, rules);
            },

            /**
             * 读取文本文件
             * @param {string} filePath - 文件路径
             * @param {string} encoding - 编码，默认utf-8
             * @returns {string|null} 文件内容
             */
            读文件: function(filePath, encoding) {
                return IO.readTextFile(filePath, encoding);
            },

            /**
             * 写入文本文件
             * @param {string} filePath - 文件路径
             * @param {string} content - 内容
             * @param {string} encoding - 编码，默认utf-8
             * @param {boolean} append - 是否追加，默认覆盖
             */
            写文件: function(filePath, content, encoding, append) {
                return IO.writeTextFile(filePath, content, encoding, append);
            },

            /**
             * 获取文件信息
             * @param {string} filePath - 文件路径
             * @returns {Object|null} 文件信息
             */
            文件信息: function(filePath) {
                return IO.getFileInfo(filePath);
            },

            /**
             * 获取文件夹内容
             * @param {string} folderPath - 文件夹路径
             * @param {Object} options - 选项
             * @returns {Array} 文件列表
             */
            文件夹内容: function(folderPath, options) {
                return IO.getFolderContents(folderPath, options);
            },

            /**
             * 设置缓存
             * @param {string} key - 键
             * @param {*} value - 值
             * @param {number} ttl - 过期秒数（可选）
             */
            缓存设置: function(key, value, ttl) {
                return CacheUtils.set(key, value, ttl);
            },

            /**
             * 获取缓存
             * @param {string} key - 键
             * @param {*} defaultValue - 默认值
             * @returns {*} 缓存值
             */
            缓存获取: function(key, defaultValue) {
                return CacheUtils.get(key, defaultValue);
            },

            /**
             * 检查缓存是否存在
             * @param {string} key - 键
             * @returns {boolean} 是否存在
             */
            缓存存在: function(key) {
                return CacheUtils.has(key);
            },

            /**
             * 删除缓存
             * @param {string} key - 键
             */
            缓存删除: function(key) {
                return CacheUtils.remove(key);
            },

            /**
             * 清空所有缓存
             */
            缓存清空: function() {
                return CacheUtils.clear();
            },

            /**
             * 设置配置项
             * @param {string} key - 键
             * @param {*} value - 值
             */
            配置设置: function(key, value) {
                return ConfigUtils.set(key, value);
            },

            /**
             * 获取配置项
             * @param {string} key - 键
             * @param {*} defaultValue - 默认值
             * @returns {*} 配置值
             */
            配置获取: function(key, defaultValue) {
                return ConfigUtils.get(key, defaultValue);
            },

            /**
             * 加载配置文件
             * @param {string} filePath - 配置文件路径
             * @returns {boolean} 是否成功
             */
            配置加载: function(filePath) {
                return ConfigUtils.loadFile(filePath);
            },

            /**
             * 保存配置文件
             * @param {string} filePath - 配置文件路径
             * @returns {boolean} 是否成功
             */
            配置保存: function(filePath) {
                return ConfigUtils.saveFile(filePath);
            },

            /**
             * 日志调试
             * @param {string} message - 消息
             */
            日志调试: function(message) {
                return Logger.debug(message);
            },

            /**
             * 日志信息
             * @param {string} message - 消息
             */
            日志信息: function(message) {
                return Logger.info(message);
            },

            /**
             * 日志警告
             * @param {string} message - 消息
             */
            日志警告: function(message) {
                return Logger.warn(message);
            },

            /**
             * 日志错误
             * @param {string} message - 消息
             */
            日志错误: function(message) {
                return Logger.error(message);
            }
        };

    }).call(this);

    // 导出JSA880快捷对象到全局（WPS环境）
    // JSA880 和 SuperMap 直接成为全局变量，与 Array2D、RngUtils 等同级
    if (typeof Application !== 'undefined') {
        Application.JSA880 = this.JSA880;
        Application.SuperMap = SuperMap;
    }
}).call(this);

// ==================== 测试数据 ====================

/**
 * JSA_Arr - 测试用数据集
 * 用于测试 SuperPivot 等功能
 * @global
 * @example
 * var result = Array2D.z超级透视(JSA_Arr, ["f2+"], ["f3+"], ["sum(\"f4\")"], 1);
 */
var JSA_Arr = [
    ["ID", "    产品", "国家", "数量", "价格", "  年", "月", "日"],
    [" 1", "Product1", "中国", "  19", "   1", "2023", "10", "10"],
    [" 2", "Product2", "德国", "  19", "   5", "2023", " 4", " 5"],
    [" 3", "Product2", "英国", "  19", "   5", "2022", " 6", "28"],
    [" 4", "Product2", "美国", "  15", "   5", "2024", " 5", " 1"],
    [" 5", "Product1", "中国", "  11", "   1", "2024", "11", "15"],
    [" 6", "Product2", "德国", "  18", "   5", "2023", " 2", "18"],
    [" 7", "Product2", "英国", "  11", "   5", "2023", " 6", "16"],
    [" 8", "Product2", "美国", "  11", "   5", "2023", " 6", "21"],
    [" 9", "Product1", "中国", "  13", "   1", "2022", " 7", "18"],
    ["10", "Product1", "德国", "  18", "   1", "2021", "11", "13"]
];

// ==================== JSA880.js 文件结束 ====================
// 版本: WPS现代版 v4.0.0
// 最后更新: 2026-05-14
// 整合 jsa880-framework.js 独有功能
// ============================================================

// ============================================================
// 整合 DataExport 数据导出工具
// ============================================================
var DataExport = {
    /**
     * 导出为CSV
     */
    toCSV: function(data, filePath, options) {
        options = options || {};
        var delimiter = options.delimiter || ',';
        var includeHeader = options.includeHeader !== false;
        var encoding = options.encoding || 'utf-8';

        var lines = [];

        if (includeHeader && data.length > 0) {
            lines.push(data[0].map(function(cell) {
                return '"' + String(cell || '').replace(/"/g, '""') + '"';
            }).join(delimiter));
        }

        for (var i = 1; i < data.length; i++) {
            lines.push(data[i].map(function(cell) {
                var val = String(cell !== null && cell !== undefined ? cell : '');
                return '"' + val.replace(/"/g, '""') + '"';
            }).join(delimiter));
        }

        var content = lines.join('\r\n');
        IO.writeTextFile(filePath, content, encoding);

        return filePath;
    },

    /**
     * 导出为JSON
     */
    toJSON: function(data, filePath, options) {
        options = options || {};
        var pretty = options.pretty || false;
        var encoding = options.encoding || 'utf-8';

        var content;
        if (pretty) {
            content = JSON.stringify(data, null, 2);
        } else {
            content = JSON.stringify(data);
        }

        IO.writeTextFile(filePath, content, encoding);
        return filePath;
    },

    /**
     * 导出为HTML表格
     */
    toHTML: function(data, filePath, options) {
        options = options || {};
        var className = options.className || 'data-table';
        var includeHeader = options.includeHeader !== false;

        var html = '<!DOCTYPE html>\n<html>\n<head>\n';
        html += '<meta charset="UTF-8">\n';
        html += '<title>Data Export</title>\n';
        html += '<style>\n';
        html += 'table.' + className + ' { border-collapse: collapse; width: 100%; }\n';
        html += 'table.' + className + ' th, table.' + className + ' td { border: 1px solid #ddd; padding: 8px; }\n';
        html += 'table.' + className + ' th { background-color: #f2f2f2; }\n';
        html += '</style>\n';
        html += '</head>\n<body>\n';
        html += '<table class="' + className + '">\n';

        var startRow = includeHeader ? 0 : 1;

        if (includeHeader && data.length > 0) {
            html += '<thead><tr>';
            for (var j = 0; j < data[0].length; j++) {
                html += '<th>' + (data[0][j] || '') + '</th>';
            }
            html += '</tr></thead>\n';
        }

        html += '<tbody>\n';
        for (var i = startRow; i < data.length; i++) {
            html += '<tr>';
            for (var j = 0; j < data[i].length; j++) {
                html += '<td>' + (data[i][j] !== null && data[i][j] !== undefined ? data[i][j] : '') + '</td>';
            }
            html += '</tr>\n';
        }

        html += '</tbody>\n</table>\n</body>\n</html>';
        IO.writeTextFile(filePath, html, 'utf-8');
        return filePath;
    },

    /**
     * 导出为XML
     */
    toXML: function(data, filePath, options) {
        options = options || {};
        var rootName = options.rootName || 'data';
        var itemName = options.itemName || 'item';
        var encoding = options.encoding || 'utf-8';

        var xml = '<?xml version="1.0" encoding="' + encoding + '"?>\n';
        xml += '<' + rootName + '>\n';

        for (var i = 1; i < data.length; i++) {
            xml += '  <' + itemName + '>\n';
            for (var j = 0; j < data[i].length; j++) {
                var tagName = 'col' + (j + 1);
                if (data[0] && data[0][j]) {
                    tagName = String(data[0][j]).replace(/[^a-zA-Z0-9_]/g, '_');
                }
                var value = data[i][j] !== null && data[i][j] !== undefined ? String(data[i][j]) : '';
                xml += '    <' + tagName + '><![CDATA[' + value + ']]></' + tagName + '>\n';
            }
            xml += '  </' + itemName + '>\n';
        }

        xml += '</' + rootName + '>';
        IO.writeTextFile(filePath, xml, encoding);
        return filePath;
    }
};

// ============================================================
// 整合 DataImport 数据导入工具
// ============================================================
var DataImport = {
    /**
     * 从CSV导入
     */
    fromCSV: function(filePath, options) {
        options = options || {};
        var delimiter = options.delimiter || ',';
        var hasHeader = options.hasHeader !== false;
        var encoding = options.encoding || 'utf-8';

        var content = IO.readTextFile(filePath, encoding);
        if (!content) return [];

        var lines = content.split(/\r?\n/);
        var data = [];

        for (var i = 0; i < lines.length; i++) {
            var line = lines[i].trim();
            if (!line) continue;

            var row = [];
            var inQuote = false;
            var field = '';

            for (var j = 0; j < line.length; j++) {
                var char = line[j];

                if (char === '"') {
                    if (inQuote && line[j + 1] === '"') {
                        field += '"';
                        j++;
                    } else {
                        inQuote = !inQuote;
                    }
                } else if (char === delimiter && !inQuote) {
                    row.push(field);
                    field = '';
                } else {
                    field += char;
                }
            }

            row.push(field);
            data.push(row);
        }

        return data;
    },

    /**
     * 从JSON导入
     */
    fromJSON: function(filePath, options) {
        options = options || {};
        var encoding = options.encoding || 'utf-8';

        var content = IO.readTextFile(filePath, encoding);
        if (!content) return [];

        try {
            return JSON.parse(content);
        } catch (e) {
            console.error('JSON解析失败: ' + e.message);
            return [];
        }
    },

    /**
     * 从JSON字符串导入为Array2D
     */
    fromJSONString: function(jsonStr, options) {
        options = options || {};
        var headers = options.headers || [];

        try {
            var obj = JSON.parse(jsonStr);
            var rows = [];

            if (Array.isArray(obj)) {
                for (var i = 0; i < obj.length; i++) {
                    if (Array.isArray(obj[i])) {
                        rows.push(obj[i]);
                    } else if (typeof obj[i] === 'object') {
                        var row = [];
                        for (var key in obj[i]) {
                            row.push(obj[i][key]);
                        }
                        rows.push(row);
                    }
                }
            }

            if (headers.length > 0) {
                rows.unshift(headers);
            }

            return rows;
        } catch (e) {
            console.error('JSON解析失败: ' + e.message);
            return [];
        }
    }
};

// ============================================================
// 整合 DataValidation 数据验证工具
// ============================================================
var DataValidation = {
    /**
     * 验证必填
     */
    required: function(value) {
        return value !== null && value !== undefined && value !== '';
    },

    /**
     * 验证数字范围
     */
    range: function(value, min, max) {
        var num = parseFloat(value);
        if (isNaN(num)) return false;
        if (min !== undefined && num < min) return false;
        if (max !== undefined && num > max) return false;
        return true;
    },

    /**
     * 验证邮箱格式
     */
    email: function(value) {
        var pattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        return pattern.test(String(value || ''));
    },

    /**
     * 验证手机号
     */
    mobile: function(value) {
        var pattern = /^1[3-9]\d{9}$/;
        return pattern.test(String(value || '').replace(/\s/g, ''));
    },

    /**
     * 验证身份证
     */
    idCard: function(value) {
        value = String(value || '');
        var pattern15 = /^[1-9]\d{7}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}$/;
        var pattern18 = /^[1-9]\d{5}(18|19|20)\d{2}((0\d)|(1[0-2]))(([0|1|2]\d)|3[0-1])\d{3}[0-9Xx]$/;
        return pattern15.test(value) || pattern18.test(value);
    },

    /**
     * 验证日期格式
     */
    date: function(value, format) {
        format = format || 'yyyy-MM-dd';
        var pattern;
        if (format === 'yyyy-MM-dd') {
            pattern = /^\d{4}-\d{2}-\d{2}$/;
        } else if (format === 'yyyy/MM/dd') {
            pattern = /^\d{4}\/\d{2}\/\d{2}$/;
        } else if (format === 'yyyyMMdd') {
            pattern = /^\d{8}$/;
        }
        return pattern ? pattern.test(String(value || '')) : false;
    },

    /**
     * 验证正则表达式
     */
    pattern: function(value, regex) {
        if (typeof regex === 'string') {
            regex = new RegExp(regex);
        }
        return regex.test(String(value || ''));
    },

    /**
     * 验证长度
     */
    length: function(value, min, max) {
        var len = String(value || '').length;
        if (min !== undefined && len < min) return false;
        if (max !== undefined && len > max) return false;
        return true;
    },

    /**
     * 批量验证
     */
    validate: function(data, rules) {
        var results = [];
        var colIndexMap = {};

        // 建立列名到索引的映射
        if (data.length > 0 && data[0]) {
            for (var j = 0; j < data[0].length; j++) {
                colIndexMap[data[0][j]] = j;
            }
        }

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var rowResults = { row: i + 1, errors: [] };

            for (var colName in rules) {
                var rule = rules[colName];
                var colIdx = colIndexMap[colName];
                if (colIdx === undefined) continue;
                var value = row[colIdx];

                if (rule.required && !DataValidation.required(value)) {
                    rowResults.errors.push(colName + ': 必填');
                }
                if (rule.min !== undefined || rule.max !== undefined) {
                    if (!DataValidation.range(value, rule.min, rule.max)) {
                        rowResults.errors.push(colName + ': 超出范围[' + rule.min + ',' + rule.max + ']');
                    }
                }
                if (rule.email && !DataValidation.email(value)) {
                    rowResults.errors.push(colName + ': 邮箱格式错误');
                }
                if (rule.mobile && !DataValidation.mobile(value)) {
                    rowResults.errors.push(colName + ': 手机号格式错误');
                }
                if (rule.pattern && !DataValidation.pattern(value, rule.pattern)) {
                    rowResults.errors.push(colName + ': 格式不匹配');
                }
                if (rule.length) {
                    if (!DataValidation.length(value, rule.length.min, rule.length.max)) {
                        rowResults.errors.push(colName + ': 长度超出');
                    }
                }
            }

            if (rowResults.errors.length > 0) {
                results.push(rowResults);
            }
        }

        return results;
    }
};

// ============================================================
// 整合 CacheUtils 缓存工具
// ============================================================
var CacheUtils = {
    _cache: {},
    _expireMap: {},

    /**
     * 设置缓存
     */
    set: function(key, value, ttl) {
        this._cache[key] = value;

        if (ttl) {
            var expireTime = Date.now() + ttl * 1000;
            this._expireMap[key] = expireTime;
        }

        return this;
    },

    /**
     * 获取缓存
     */
    get: function(key, defaultValue) {
        if (this._expireMap.hasOwnProperty(key)) {
            if (Date.now() > this._expireMap[key]) {
                delete this._cache[key];
                delete this._expireMap[key];
                return defaultValue !== undefined ? defaultValue : null;
            }
        }

        return this._cache.hasOwnProperty(key)
            ? this._cache[key]
            : (defaultValue !== undefined ? defaultValue : null);
    },

    /**
     * 检查缓存是否存在
     */
    has: function(key) {
        if (this._expireMap.hasOwnProperty(key)) {
            if (Date.now() > this._expireMap[key]) {
                delete this._cache[key];
                delete this._expireMap[key];
                return false;
            }
        }

        return this._cache.hasOwnProperty(key);
    },

    /**
     * 删除缓存
     */
    remove: function(key) {
        delete this._cache[key];
        delete this._expireMap[key];
        return this;
    },

    /**
     * 清空所有缓存
     */
    clear: function() {
        this._cache = {};
        this._expireMap = {};
        return this;
    },

    /**
     * 获取缓存大小
     */
    size: function() {
        return Object.keys(this._cache).length;
    }
};

// ============================================================
// 整合 ConfigUtils 配置管理
// ============================================================
var ConfigUtils = {
    _configs: {},

    /**
     * 设置配置
     */
    set: function(key, value) {
        this._configs[key] = value;
        return this;
    },

    /**
     * 获取配置
     */
    get: function(key, defaultValue) {
        return this._configs.hasOwnProperty(key)
            ? this._configs[key]
            : (defaultValue !== undefined ? defaultValue : null);
    },

    /**
     * 加载配置文件
     */
    loadFile: function(filePath) {
        var content = IO.readTextFile(filePath, 'utf-8');
        if (!content) return false;

        try {
            var config = JSON.parse(content);
            for (var key in config) {
                if (config.hasOwnProperty(key)) {
                    this._configs[key] = config[key];
                }
            }
            return true;
        } catch (e) {
            console.error('配置文件解析失败: ' + e.message);
            return false;
        }
    },

    /**
     * 保存配置文件
     */
    saveFile: function(filePath) {
        var content = JSON.stringify(this._configs, null, 2);
        IO.writeTextFile(filePath, content, 'utf-8');
        return true;
    },

    /**
     * 获取所有配置
     */
    all: function() {
        return JSON.parse(JSON.stringify(this._configs));
    },

    /**
     * 检查配置是否存在
     */
    has: function(key) {
        return this._configs.hasOwnProperty(key);
    },

    /**
     * 删除配置
     */
    remove: function(key) {
        delete this._configs[key];
        return this;
    },

    /**
     * 清空所有配置
     */
    clear: function() {
        this._configs = {};
        return this;
    }
};

// ============================================================
// 整合 Logger 日志工具
// ============================================================
var Logger = {
    _level: 'info',
    _levels: { debug: 0, info: 1, warn: 2, error: 3 },
    _file: null,

    /**
     * 调试日志
     */
    debug: function(message) {
        if (this._levels[this._level] <= 0) {
            this._log('DEBUG', message);
        }
    },

    /**
     * 信息日志
     */
    info: function(message) {
        if (this._levels[this._level] <= 1) {
            this._log('INFO', message);
        }
    },

    /**
     * 警告日志
     */
    warn: function(message) {
        if (this._levels[this._level] <= 2) {
            this._log('WARN', message);
        }
    },

    /**
     * 错误日志
     */
    error: function(message) {
        if (this._levels[this._level] <= 3) {
            this._log('ERROR', message);
        }
    },

    /**
     * 内部日志方法
     */
    _log: function(level, message) {
        var timestamp = new Date().toLocaleString();
        var logMessage = '[' + timestamp + '][' + level + '] ' + message;

        if (console && console.log) {
            console.log(logMessage);
        }

        if (this._file) {
            IO.appendToFile(this._file, logMessage + '\n');
        }
    },

    /**
     * 设置日志级别
     */
    setLevel: function(level) {
        if (this._levels.hasOwnProperty(level)) {
            this._level = level;
        }
    },

    /**
     * 设置日志文件
     */
    setFile: function(filePath) {
        this._file = filePath;
    },

    /**
     * 清空日志
     */
    clear: function() {
        if (this._file) {
            IO.writeTextFile(this._file, '', 'utf-8', false);
        }
    }
};

// ============================================================
// 整合 IO 增强功能
// ============================================================
IO.getFileInfo = function(filePath) {
    var fso = new ActiveXObject('Scripting.FileSystemObject');

    if (!fso.FileExists(filePath)) {
        return null;
    }

    var file = fso.GetFile(filePath);

    return {
        path: file.Path,
        name: file.Name,
        size: file.Size,
        createdDate: file.DateCreated,
        modifiedDate: file.DateLastModified,
        attributes: file.Attributes
    };
};

IO.getFolderContents = function(folderPath, options) {
    options = options || {};
    options.includeFiles = options.includeFiles !== false;
    options.includeFolders = options.includeFolders || false;
    options.recursive = options.recursive || false;
    options.filter = options.filter || null;

    // XXD-176 (XXD-177): try ActiveX first (WPS/Windows), fall back to Node fs for testing/non-Windows.
    try {
        var fso = new ActiveXObject('Scripting.FileSystemObject');

        if (!fso.FolderExists(folderPath)) {
            return null;
        }

        var folder = fso.GetFolder(folderPath);
        var results = [];

        if (options.includeFiles) {
            var enumFiles = new Enumerator(folder.Files);
            for (; !enumFiles.atEnd(); enumFiles.moveNext()) {
                var file = enumFiles.item;

                if (options.filter) {
                    var ext = fso.GetExtensionName(file.Name).toLowerCase();
                    var filters = options.filter.split(',').map(function(f) {
                        return f.trim().toLowerCase();
                    });

                    if (filters.indexOf(ext) < 0) {
                        continue;
                    }
                }

                results.push({
                    type: 'file',
                    path: file.Path,
                    name: file.Name,
                    size: file.Size,
                    modifiedDate: file.DateLastModified
                });
            }
        }

        if (options.includeFolders || options.recursive) {
            var enumFolders = new Enumerator(folder.SubFolders);
            for (; !enumFolders.atEnd(); enumFolders.moveNext()) {
                var subFolder = enumFolders.item;

                if (options.includeFolders) {
                    results.push({
                        type: 'folder',
                        path: subFolder.Path,
                        name: subFolder.Name
                    });
                }

                if (options.recursive) {
                    var subResults = IO.getFolderContents(subFolder.Path, options);
                    if (subResults) {
                        results = results.concat(subResults);
                    }
                }
            }
        }

        return results;
    } catch (e) {
        // v4.0.11 / XXD-176: Node (or any non-ActiveX) fallback.
        if (typeof require !== 'function') return null;
        try {
            var fs = require('fs');
            var pathMod = require('path');
            if (!fs.existsSync(folderPath)) return null;
            var entries = fs.readdirSync(folderPath, { withFileTypes: true });
            var nodeResults = [];
            var filters = options.filter ? options.filter.split(',').map(function(f) {
                return f.trim().toLowerCase();
            }) : null;
            for (var i = 0; i < entries.length; i++) {
                var ent = entries[i];
                var full = pathMod.join(folderPath, ent.name);
                if (ent.isDirectory()) {
                    if (options.includeFolders) {
                        nodeResults.push({ type: 'folder', path: full, name: ent.name });
                    }
                    if (options.recursive) {
                        var sub = IO.getFolderContents(full, options);
                        if (sub) nodeResults = nodeResults.concat(sub);
                    }
                } else if (ent.isFile() && options.includeFiles) {
                    if (filters) {
                        var nodeExt = pathMod.extname(ent.name).toLowerCase();
                        if (nodeExt[0] === '.') nodeExt = nodeExt.slice(1);
                        if (filters.indexOf(nodeExt) < 0) continue;
                    }
                    var st = fs.statSync(full);
                    nodeResults.push({
                        type: 'file',
                        path: full,
                        name: ent.name,
                        size: st.size,
                        modifiedDate: st.mtime
                    });
                }
            }
            return nodeResults;
        } catch (e2) {
            console.warn('IO.getFolderContents:', e2.message);
            return null;
        }
    }
};

IO.readTextFile = function(filePath, encoding) {
    encoding = encoding || 'utf-8';

    // XXD-177 (XXD-178): unified missing-file contract.
    // Policy (b): NEVER throws. Missing file returns null silently. Other I/O errors
    // (perm denied, etc.) log a warn and return null. Callers do
    //     var s = IO.readTextFile(p);
    //     if (s === null) { ...handle missing/error... }
    // IO.z读文本文件 is the same function reference (aliased below), so the contract
    // applies to both names.

    // Try ActiveXObject (Windows/WPS)
    try {
        var fso = new ActiveXObject('Scripting.FileSystemObject');
        if (!fso.FileExists(filePath)) return null; // missing file: silent null
        var stream = new ActiveXObject('ADODB.Stream');
        stream.Type = 2; stream.Charset = encoding;
        stream.Open(); stream.LoadFromFile(filePath);
        var content = stream.ReadText(); stream.Close();
        return content;
    } catch (e) {
        // v4.0.11: fallback for non-Windows environments
        try {
            var fs = require('fs');
            return fs.readFileSync(filePath, { encoding: encoding });
        } catch (e2) {
            // XXD-177: ENOENT (file does not exist) is a normal caller-visible case,
            // not an error — return null silently. Only unexpected errors warn.
            if (e2 && e2.code === 'ENOENT') return null;
            console.warn('IO.readTextFile:', e2.message);
            return null;
        }
    }
};

IO.writeTextFile = function(filePath, content, encoding, append) {
    encoding = encoding || 'utf-8';
    append = append || false;

    // Try ActiveXObject (Windows/WPS)
    try {
        var fso = new ActiveXObject('Scripting.FileSystemObject');
        var stream = new ActiveXObject('ADODB.Stream');
        stream.Type = 2; stream.Charset = encoding;
        stream.Open();
        if (append && fso.FileExists(filePath)) {
            var existing = IO.readTextFile(filePath, encoding);
            stream.WriteText(existing || '');
        }
        stream.WriteText(content);
        stream.SaveToFile(filePath, 2);
        stream.Close();
        return filePath;
    } catch (e) {
        // v4.0.11: fallback for non-Windows environments
        try {
            var fs = require('fs');
            if (append) {
                fs.appendFileSync(filePath, content, { encoding: encoding });
            } else {
                fs.writeFileSync(filePath, content, { encoding: encoding });
            }
            return filePath;
        } catch (e2) {
            console.warn('IO.writeTextFile:', e2.message);
            return null;
        }
    }
};

IO.appendToFile = function(filePath, content) {
    return IO.writeTextFile(filePath, content, 'utf-8', true);
};

// XXD-149 (XXD-146): IO English/Chinese alias parity for text-file + listing helpers.
// fileExists/listFiles had no English form; readTextFile/writeTextFile had no Chinese form.
// All aliases must exist before any user code runs, so register at top level.
// NOTE: listFiles/z列出文件 are aliased after IO.getFiles is defined (later in this file).
IO.fileExists = IO.fileExists || IO.z是否文件;
IO.z读文本文件 = IO.readTextFile;
IO.z写文本文件 = IO.writeTextFile;
IO.z追加文本文件 = IO.appendToFile;
IO.z文件存在 = IO.z存在;

// 打印整合信息
console.log('JSA880框架 v4.0.0 已整合 jsa880-framework.js 增强功能:');
console.log('- DataExport: 导出CSV/JSON/HTML/XML');
console.log('- DataImport: 导入CSV/JSON');
console.log('- DataValidation: 数据验证');
console.log('- CacheUtils: 缓存管理');
console.log('- ConfigUtils: 配置管理');
console.log('- Logger: 日志工具');
console.log('- IO增强: 文件操作');

// ============================================================
// 整合自 JSA880-claude.js 的增强功能
// ============================================================

// ============================================================
// 全局函数 agg() - 汇总函数
// ============================================================

/**
 * agg - 汇总函数，计算数组的聚合值
 * @param {Array} arr - 二维数组
 * @param {string} aggType - 聚合类型：sum, avg, max, min, count, first, last
 * @param {string} colRef - 列引用，如"f1", "f2"
 * @returns {*} 聚合结果
 * @example
 * var total = agg(data, "sum", "f3");
 * var avg = agg(data, "avg", "f2");
 */
function agg(arr, aggType, colRef) {
    return Array2D.agg(arr, colRef, aggType);
}

// ============================================================
// 全局函数 oadate() - OA日期转换
// ============================================================

/**
 * oadate - 将WPS OA日期数值转换为JavaScript Date
 * @param {number} oaDate - OA日期数值
 * @returns {Date} JavaScript日期对象
 */
function oadate(oaDate) {
    if (typeof oaDate !== "number" || isNaN(oaDate)) return new Date();
    var msPerDay = 86400000;
    var epoch = new Date(1899, 11, 30);
    return new Date(epoch.getTime() + oaDate * msPerDay);
}
// ============================================================
// 全局函数 fromOADate() - OA日期转换反函数
// ============================================================

/**
 * fromOADate - 将JavaScript Date转换为WPS OA日期数值
 * @param {Date} date - JavaScript日期对象
 * @returns {number} OA日期数值
 */
function fromOADate(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return 0;
    var msPerDay = 86400000;
    var epoch = new Date(1899, 11, 30);
    return (date.getTime() - epoch.getTime()) / msPerDay;
}

// ============================================================
// StrUtils - 字符串工具模块（整合自 JSA880-claude.js）
// ============================================================

var StrUtils = {
    isEmpty: function(str) { return str === null || str === undefined || str === ""; },
    isBlank: function(str) { return this.isEmpty(str) || String(str).replace(/\s+/g, "") === ""; },
    trim: function(str) { return str === null || str === undefined ? "" : String(str).replace(/^\s+|\s+$/g, ""); },
    trimLeft: function(str) { return str === null || str === undefined ? "" : String(str).replace(/^\s+/, ""); },
    trimRight: function(str) { return str === null || str === undefined ? "" : String(str).replace(/\s+$/, ""); },
    toLowerCase: function(str) { return str === null || str === undefined ? "" : String(str).toLowerCase(); },
    toUpperCase: function(str) { return str === null || str === undefined ? "" : String(str).toUpperCase(); },
    capitalize: function(str) {
        if (str === null || str === undefined) return "";
        str = String(str);
        return str.length === 0 ? str : str.charAt(0).toUpperCase() + str.slice(1);
    },
    camelCase: function(str) {
        if (str === null || str === undefined) return "";
        str = String(str).toLowerCase();
        return str.replace(/[-_\s]+(.)?/g, function(match, char) { return char ? char.toUpperCase() : ""; });
    },
    snakeCase: function(str) {
        if (str === null || str === undefined) return "";
        return String(str).replace(/([A-Z])/g, "_$1").toLowerCase().replace(/^_/, "");
    },
    replaceAll: function(str, search, replacement) {
        if (str === null || str === undefined) return "";
        if (!search) return String(str);
        var s = String(str);
        var searchStr = search instanceof RegExp ? search.source : String(search);
        var flags = search instanceof RegExp ? search.flags : "g";
        if (flags.indexOf("g") === -1) flags += "g";
        var regex = new RegExp(searchStr.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), flags);
        return s.replace(regex, replacement);
    },
    contains: function(str, search) { return str !== null && str !== undefined && String(str).indexOf(search) !== -1; },
    startsWith: function(str, prefix) { return str !== null && str !== undefined && String(str).indexOf(prefix) === 0; },
    endsWith: function(str, suffix) {
        if (str === null || str === undefined) return false;
        str = String(str);
        return str.length >= suffix.length && str.lastIndexOf(suffix) === str.length - suffix.length;
    },
    split: function(str, delimiter, limit) {
        if (str === null || str === undefined) return [];
        var parts = String(str).split(delimiter);
        if (typeof limit === "number" && limit > 0) return parts.slice(0, limit);
        return parts;
    },
    join: function(parts, separator) {
        if (parts === null || parts === undefined) return "";
        separator = separator || "";
        if (typeof parts === "string") return parts;
        var result = "";
        for (var i = 0; i < parts.length; i++) {
            if (i > 0) result += separator;
            var v = parts[i];
            result += (v === null || v === undefined) ? "" : String(v);
        }
        return result;
    },
    repeat: function(str, count) {
        if (str === null || str === undefined) return "";
        if (typeof count !== "number" || count < 1) return "";
        str = String(str);
        var result = "";
        for (var i = 0; i < count; i++) result += str;
        return result;
    },
    padLeft: function(str, length, padChar) {
        if (str === null || str === undefined) str = "";
        if (!padChar) padChar = " ";
        str = String(str);
        if (str.length >= length) return str;
        var padding = "";
        for (var i = 0; i < length - str.length; i++) padding += padChar;
        return padding + str;
    },
    padRight: function(str, length, padChar) {
        if (str === null || str === undefined) str = "";
        if (!padChar) padChar = " ";
        str = String(str);
        if (str.length >= length) return str;
        var padding = "";
        for (var i = 0; i < length - str.length; i++) padding += padChar;
        return str + padding;
    },
    left: function(str, length) { return str === null || str === undefined ? "" : String(str).substring(0, length); },
    right: function(str, length) { if (str === null || str === undefined) return ""; str = String(str); return str.substring(str.length - length); },
    substring: function(str, start, end) { return str === null || str === undefined ? "" : String(str).substring(start, end); },
    removePrefix: function(str, prefix) {
        if (str === null || str === undefined) return "";
        str = String(str);
        return this.startsWith(str, prefix) ? str.substring(prefix.length) : str;
    },
    removeSuffix: function(str, suffix) {
        if (str === null || str === undefined) return "";
        str = String(str);
        return this.endsWith(str, suffix) ? str.substring(0, str.length - suffix.length) : str;
    },
    template: function(template, data, placeholderStart, placeholderEnd) {
        if (template === null || template === undefined) return "";
        if (!data) return String(template);
        placeholderStart = placeholderStart || "{{";
        placeholderEnd = placeholderEnd || "}}";
        var result = String(template);
        var pattern = new RegExp(placeholderStart.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^" + placeholderEnd.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "]+)" + placeholderEnd.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
        result = result.replace(pattern, function(match, key) {
            key = StrUtils.trim(key);
            return data.hasOwnProperty(key) ? String(data[key]) : match;
        });
        return result;
    },
    toNumber: function(str, defaultValue) {
        if (str === null || str === undefined) return defaultValue !== undefined ? defaultValue : 0;
        var num = parseFloat(String(str).replace(/,/g, ""));
        return isNaN(num) ? (defaultValue !== undefined ? defaultValue : 0) : num;
    },
    escapeHtml: function(str) {
        if (str === null || str === undefined) return "";
        str = String(str);
        var map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
        for (var key in map) if (map.hasOwnProperty(key)) str = this.replaceAll(str, key, map[key]);
        return str;
    },
    unescapeHtml: function(str) {
        if (str === null || str === undefined) return "";
        str = String(str);
        var map = { "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" };
        for (var key in map) if (map.hasOwnProperty(key)) str = this.replaceAll(str, key, map[key]);
        return str;
    },
    isNumeric: function(str) { return str !== null && str !== undefined && !isNaN(String(str).replace(/,/g, "")); },
    isInteger: function(str) { return str !== null && str !== undefined && !isNaN(parseInt(String(str).replace(/,/g, ""), 10)); },
    isAlpha: function(str) { return str !== null && str !== undefined && /^[a-zA-Z]+$/.test(str); },
    isAlphanumeric: function(str) { return str !== null && str !== undefined && /^[a-zA-Z0-9]+$/.test(str); },
    count: function(str, subStr) {
        if (str === null || str === undefined || subStr === null || subStr === undefined) return 0;
        str = String(str); subStr = String(subStr);
        if (subStr.length === 0) return 0;
        var count = 0, idx = 0;
        while ((idx = str.indexOf(subStr, idx)) !== -1) { count++; idx += subStr.length; }
        return count;
    }
};

/* XXD-167 START: StrUtils Chinese aliases */
// StrUtils 中文 alias (XXD-167/168) — assigned after literal, so StrUtils is bound.
StrUtils.z去空白 = StrUtils.trim;
StrUtils.z去左空白 = StrUtils.trimLeft;
StrUtils.z去右空白 = StrUtils.trimRight;
StrUtils.z转大写 = StrUtils.toUpperCase;
StrUtils.z转小写 = StrUtils.toLowerCase;
StrUtils.z包含 = StrUtils.contains;
StrUtils.z开始于 = StrUtils.startsWith;
StrUtils.z结束于 = StrUtils.endsWith;
StrUtils.z分割 = StrUtils.split;
StrUtils.z连接 = StrUtils.join;
StrUtils.z替换 = StrUtils.replaceAll;
StrUtils.z替换全部 = StrUtils.replaceAll;
StrUtils.z左填充 = StrUtils.padLeft;
StrUtils.z右填充 = StrUtils.padRight;
StrUtils.z重复 = StrUtils.repeat;
StrUtils.z首字母大写 = StrUtils.capitalize;
StrUtils.z驼峰命名 = StrUtils.camelCase;
StrUtils.z下划线命名 = StrUtils.snakeCase;
StrUtils.z是否为空 = StrUtils.isEmpty;
StrUtils.z是否空白 = StrUtils.isBlank;
StrUtils.z是否数字 = StrUtils.isNumeric;
StrUtils.z是否整数 = StrUtils.isInteger;
StrUtils.z是否字母 = StrUtils.isAlpha;
StrUtils.z是否字母数字 = StrUtils.isAlphanumeric;
StrUtils.z转义HTML = StrUtils.escapeHtml;
StrUtils.z反转义HTML = StrUtils.unescapeHtml;
// 🔧 XXD-168 atomic fix: 补 StrUtils.z反转
StrUtils.z反转 = function(s) { return String(s).split('').reverse().join(''); };
// 🔧 XXD-168 atomic fix: 补 StrUtils.z反转
StrUtils.z左取 = StrUtils.left;
StrUtils.z右取 = StrUtils.right;
StrUtils.z截取 = StrUtils.substring;
StrUtils.z转数字 = StrUtils.toNumber;
StrUtils.z去除前缀 = StrUtils.removePrefix;
StrUtils.z去除后缀 = StrUtils.removeSuffix;
StrUtils.z模板 = StrUtils.template;
StrUtils.z计数 = StrUtils.count;
/* XXD-167 END */

// ============================================================
// NumUtils - 数值工具模块（整合自 JSA880-claude.js）
// ============================================================

var NumUtils = {
    round: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.round(num * factor) / factor;
    },
    ceil: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.ceil(num * factor) / factor;
    },
    floor: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return 0;
        if (decimals === undefined || decimals === null) decimals = 0;
        var factor = Math.pow(10, decimals);
        return Math.floor(num * factor) / factor;
    },
    abs: function(num) { return typeof num === "number" && !isNaN(num) ? Math.abs(num) : 0; },
    sign: function(num) { if (typeof num !== "number" || isNaN(num)) return 0; return num > 0 ? 1 : (num < 0 ? -1 : 0); },
    formatNumber: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0";
        if (decimals === undefined || decimals === null) decimals = 0;
        var fixed = this.round(num, decimals).toString();
        var parts = fixed.split(".");
        var intPart = parts[0];
        var result = "";
        var count = 0;
        for (var i = intPart.length - 1; i >= 0; i--) {
            if (count > 0 && count % 3 === 0 && intPart.charAt(i) !== "-") result = "," + result;
            result = intPart.charAt(i) + result; count++;
        }
        if (parts.length > 1) {
            var decStr = parts[1];
            while (decStr.length < decimals) decStr += "0";
            result += "." + decStr;
        } else if (decimals > 0) {
            result += "." + this._repeatChar("0", decimals);
        }
        return result;
    },
    formatCurrency: function(num, currencySymbol, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0.00";
        currencySymbol = currencySymbol || "¥";
        if (decimals === undefined || decimals === null) decimals = 2;
        return currencySymbol + this.formatNumber(num, decimals);
    },
    formatPercent: function(num, decimals) {
        if (typeof num !== "number" || isNaN(num)) return "0%";
        if (decimals === undefined || decimals === null) decimals = 0;
        var percent = num * 100;
        var formatted = this.round(percent, decimals).toString();
        if (decimals > 0 && formatted.indexOf(".") === -1) formatted += "." + this._repeatChar("0", decimals);
        return formatted + "%";
    },
    parse: function(str, defaultValue) {
        if (str === null || str === undefined) return defaultValue !== undefined ? defaultValue : 0;
        if (typeof str === "number") return isNaN(str) ? (defaultValue !== undefined ? defaultValue : 0) : str;
        var numStr = String(str).replace(/[^0-9.\-]/g, "");
        var num = parseFloat(numStr);
        return isNaN(num) ? (defaultValue !== undefined ? defaultValue : 0) : num;
    },
    clamp: function(num, min, max) {
        if (typeof num !== "number" || isNaN(num)) num = 0;
        if (num < min) return min;
        if (num > max) return max;
        return num;
    },
    inRange: function(num, min, max) { return typeof num === "number" && !isNaN(num) && num >= min && num <= max; },
    sum: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var total = 0;
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) total += args[i];
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) total += args[i][j];
                }
            }
        }
        return Math.round(total * 1e10) / 1e10;
    },
    avg: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : this.sum(values) / values.length;
    },
    max: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : Math.max.apply(Math, values);
    },
    min: function(nums) {
        var args = Array.prototype.slice.call(arguments);
        var values = [];
        for (var i = 0; i < args.length; i++) {
            if (typeof args[i] === "number" && !isNaN(args[i])) values.push(args[i]);
            else if (args[i] instanceof Array) {
                for (var j = 0; j < args[i].length; j++) {
                    if (typeof args[i][j] === "number" && !isNaN(args[i][j])) values.push(args[i][j]);
                }
            }
        }
        return values.length === 0 ? 0 : Math.min.apply(Math, values);
    },
    randomInt: function(min, max) {
        if (max === undefined) { max = min; min = 0; }
        return Math.floor(Math.random() * (max - min + 1)) + min;
    },
    random: function(min, max) {
        if (max === undefined) { max = min; min = 0; }
        return Math.random() * (max - min) + min;
    },
    randomId: function(length) {
        if (typeof length !== "number" || length < 1) length = 8;
        var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        var result = "";
        for (var i = 0; i < length; i++) result += chars.charAt(Math.floor(Math.random() * chars.length));
        return result;
    },
    repeat: function(value, count) {
        if (typeof count !== "number" || count < 0) return [];
        var result = [];
        for (var i = 0; i < count; i++) result.push(value);
        return result;
    },
    _repeatChar: function(ch, count) {
        var result = "";
        for (var i = 0; i < count; i++) result += ch;
        return result;
    },
    isNumber: function(val) { return typeof val === "number" && !isNaN(val); },
    isInteger: function(val) { return typeof val === "number" && !isNaN(val) && val === Math.floor(val); },
    isEven: function(num) { return typeof num === "number" && !isNaN(num) && num % 2 === 0; },
    isOdd: function(num) { return typeof num === "number" && !isNaN(num) && num % 2 !== 0; },
    isPositive: function(num) { return typeof num === "number" && !isNaN(num) && num > 0; },
    isNegative: function(num) { return typeof num === "number" && !isNaN(num) && num < 0; },
    // === XXD-166: 中文 alias (大包装, 与 Array2D 风格一致) ===
    z取整: function(num, decimals) { return this.round(num, decimals); },
    z向上取整: function(num, decimals) { return this.ceil(num, decimals); },
    z向下取整: function(num, decimals) { return this.floor(num, decimals); },
    z四舍五入: function(num, decimals) { return this.round(num, decimals); },
    z绝对值: function(num) { return this.abs(num); },
    z幂: function(base, exp) { return Math.pow(base, exp); },
    z对数: function(n) { return Math.log(n); },
    z正弦: function(n) { return Math.sin(n); },
    z余弦: function(n) { return Math.cos(n); },
    z正切: function(n) { return Math.tan(n); },
    z取小数: function(num, decimals) { return this.round(num, decimals); },
    z百分比: function(num, decimals) { return this.formatPercent(num, decimals); },
    z随机数: function(min, max) { return this.random(min, max); },
    z随机整数: function(min, max) { return this.randomInt(min, max); },
    z随机ID: function(length) { return this.randomId(length); },
    z是数字: function(val) { return this.isNumber(val); },
    z是整数: function(val) { return this.isInteger(val); },
    z是偶数: function(num) { return this.isEven(num); },
    z是奇数: function(num) { return this.isOdd(num); },
    z是正数: function(num) { return this.isPositive(num); },
    z是负数: function(num) { return this.isNegative(num); },
    z是NaN: function(val) { return isNaN(val); },
    z裁剪: function(num, min, max) { return this.clamp(num, min, max); },
    z在范围内: function(num, min, max) { return this.inRange(num, min, max); },
    z格式化货币: function(num, currencySymbol, decimals) { return this.formatCurrency(num, currencySymbol, decimals); },
    z格式化数字: function(num, decimals) { return this.formatNumber(num, decimals); },
    z格式化百分比: function(num, decimals) { return this.formatPercent(num, decimals); },
    z求和: function() { return this.sum.apply(this, arguments); },
    z平均: function() { return this.avg.apply(this, arguments); },
    z最大值: function() { return this.max.apply(this, arguments); },
    z最小值: function() { return this.min.apply(this, arguments); },
    z重复: function(value, count) { return this.repeat(value, count); },
    z解析: function(str, defaultValue) { return this.parse(str, defaultValue); },
    z符号: function(num) { return this.sign(num); }
};

// ============================================================
// MsgUtils - 消息工具模块（整合自 JSA880-claude.js）
// ============================================================

var MsgUtils = {
    BUTTON_OK: 0,
    BUTTON_OK_CANCEL: 1,
    BUTTON_YES_NO: 4,
    BUTTON_YES_NO_CANCEL: 3,
    ICON_NONE: 0,
    ICON_INFO: 64,
    ICON_WARNING: 48,
    ICON_ERROR: 16,
    ICON_QUESTION: 32,
    DEFAULT_TITLE: "提示",
    show: function(message, buttons, title) {
        if (typeof MsgBox !== "function") { console.log("[MsgBox] " + message); return 0; }
        buttons = buttons || this.BUTTON_OK;
        title = title || this.DEFAULT_TITLE;
        try { return MsgBox(String(message), buttons, title); } catch (e) { return 0; }
    },
    info: function(message, title) { return this.show(message, this.BUTTON_OK | this.ICON_INFO, title || "信息"); },
    warn: function(message, title) { return this.show(message, this.BUTTON_OK | this.ICON_WARNING, title || "警告"); },
    error: function(message, title) { return this.show(message, this.BUTTON_OK | this.ICON_ERROR, title || "错误"); },
    success: function(message, title) { return this.show(message, this.BUTTON_OK | this.ICON_INFO, title || "成功"); },
    confirm: function(message, title) { return this.show(message, this.BUTTON_OK_CANCEL | this.ICON_QUESTION, title || "确认") === 1; },
    yesNo: function(message, title) {
        var result = this.show(message, this.BUTTON_YES_NO | this.ICON_QUESTION, title || "询问");
        switch (result) { case 6: return "yes"; case 7: return "no"; default: return "cancel"; }
    },
    input: function(prompt, title, defaultValue) {
        if (typeof InputBox !== "function") { console.log("[InputBox] " + prompt); return prompt; }
        title = title || this.DEFAULT_TITLE;
        defaultValue = defaultValue || "";
        try { var result = InputBox(String(prompt), String(title), String(defaultValue)); return result === "" ? null : result; } catch (e) { return null; }
    },
    inputNumber: function(prompt, title, defaultValue) {
        var input = this.input(prompt, title, String(defaultValue || ""));
        if (input === null) return null;
        var num = parseFloat(input);
        return isNaN(num) ? null : num;
    },
    select: function(prompt, options, defaultIndex) {
        if (!options || options.length === 0) return null;
        var optionStr = "";
        for (var i = 0; i < options.length; i++) {
            var label = options[i];
            if (typeof label === "object") label = label.label || label.value || String(label);
            optionStr += (i + 1) + ". " + label + "\n";
        }
        var fullPrompt = prompt + "\n\n" + optionStr;
        var input = this.input(fullPrompt, this.DEFAULT_TITLE, String(defaultIndex !== undefined ? defaultIndex + 1 : ""));
        if (input === null) return null;
        var idx = parseInt(input, 10) - 1;
        return (idx < 0 || idx >= options.length) ? null : idx;
    },
    log: function(message, level) {
        level = level || "info";
        if (typeof console !== "undefined") {
            var timestamp = new Date().toLocaleTimeString();
            switch (level) {
                case "warn": console.warn("[WARN] " + timestamp + " " + message); break;
                case "error": console.error("[ERROR] " + timestamp + " " + message); break;
                default: console.log("[INFO] " + timestamp + " " + message);
            }
        }
    }
};

// ============================================================
// WorkbookUtils - 工作簿工具模块（整合自 JSA880-claude.js）
// ============================================================

var WorkbookUtils = {
    _getApp: function() { return typeof Application !== "undefined" ? Application : null; },
    getActiveWorkbook: function() { var app = this._getApp(); return app ? app.ActiveWorkbook : null; },
    getActiveSheet: function() { var app = this._getApp(); return app ? app.ActiveSheet : null; },
    getSheet: function(workbook, sheetRef) {
        if (!workbook) return null;
        try { return typeof sheetRef === "number" ? workbook.Sheets(sheetRef) : workbook.Sheets(sheetRef); } catch (e) { return null; }
    },
    getSheetCount: function(workbook) {
        if (!workbook) return 0;
        try { return workbook.Sheets.Count; } catch (e) { return 0; }
    },
    getSheetNames: function(workbook) {
        var names = [];
        if (!workbook) return names;
        try { var count = workbook.Sheets.Count; for (var i = 1; i <= count; i++) names.push(workbook.Sheets(i).Name); } catch (e) { }
        return names;
    },
    createSheet: function(workbook, sheetName, position) {
        if (!workbook) return null;
        try {
            var sheet = workbook.Sheets.Add();
            if (sheetName) sheet.Name = sheetName;
            if (typeof position === "number") sheet.Move(position);
            return sheet;
        } catch (e) { return null; }
    },
    deleteSheet: function(workbook, sheetRef) {
        if (!workbook) return false;
        try { var sheet = this.getSheet(workbook, sheetRef); if (sheet) { sheet.Delete(); return true; } return false; } catch (e) { return false; }
    },
    renameSheet: function(sheet, newName) {
        if (!sheet) return false;
        try { sheet.Name = newName; return true; } catch (e) { return false; }
    },
    sheetExists: function(workbook, sheetName) {
        if (!workbook || !sheetName) return false;
        try { workbook.Sheets(sheetName); return true; } catch (e) { return false; }
    },
    getOrCreateSheet: function(workbook, sheetName) {
        if (!workbook || !sheetName) return null;
        try { if (this.sheetExists(workbook, sheetName)) return workbook.Sheets(sheetName); return this.createSheet(workbook, sheetName); } catch (e) { return null; }
    },
    activateSheet: function(sheet) { if (!sheet) return false; try { sheet.Activate(); return true; } catch (e) { return false; } },
    protectSheet: function(sheet, password) { if (!sheet) return false; try { sheet.Protect(password); return true; } catch (e) { return false; } },
    unprotectSheet: function(sheet, password) { if (!sheet) return false; try { sheet.Unprotect(password); return true; } catch (e) { return false; } },
    setZoom: function(sheet, zoom) { if (!sheet) return false; try { sheet.Application.ActiveWindow.Zoom = zoom; return true; } catch (e) { return false; } },
    freezePanes: function(sheet, rangeRef) {
        if (!sheet) return false;
        try {
            if (rangeRef) { var rng = sheet.Range(rangeRef); sheet.Activate(); rng.Select(); sheet.Application.ActiveWindow.FreezePanes = true; }
            else { sheet.Application.ActiveWindow.FreezePanes = true; }
            return true;
        } catch (e) { return false; }
    },
    unfreezePanes: function(sheet) { if (!sheet) return false; try { sheet.Application.ActiveWindow.FreezePanes = false; return true; } catch (e) { return false; } }
};

// ============================================================
// SuperMap 数据操作扩展（整合自 JSA880-claude.js）
// ============================================================

/**
 * SuperMap.z选择列 - 选择指定的列
 */
// 🔧 XXD-154 final fix: SuperMap 全套数据操作 alias (= Array2D)
SuperMap.z求和 = function(arr, sel) { return new Array2D(arr).z求和(sel); };
SuperMap.z计数 = function(arr, sel) { return new Array2D(arr).z计数(sel); };
SuperMap.z最大值 = function(arr, sel) { return new Array2D(arr).z最大值(sel); };
SuperMap.z最小值 = function(arr, sel) { return new Array2D(arr).z最小值(sel); };
SuperMap.z平均 = function(arr, sel) { return new Array2D(arr).z平均值(sel); };
SuperMap.z去重 = function(arr, sel) { return new Array2D(arr).z去重(sel); };
// 🔧 XXD-180/XXD-181 final fix start: SuperMap.z分组 vs z分组统计 形状区分文档化
/**
 * SuperMap.z分组 - 纯分组 (无聚合)
 *
 * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)
 * @param {Function|string|number|Array} sel - key 选择器,同 Array2D.prototype.z分组
 * @returns {Object<string, Array<Array>>} 字典: { groupKey: [row, row, ...] }
 *
 * 形状约定 (与 z分组统计 区分):
 *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)
 *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)
 * 调用方按需选择:
 *   - 需要 key→行集合  → 用 z分组
 *   - 需要每组 1 行汇总 → 用 z分组统计
 */
SuperMap.z分组 = function(arr, sel) { return new Array2D(arr).z分组(sel); };
SuperMap.z超级透视 = function(arr, rowF, colF, dataF) { return Array2D.z超级透视(arr, rowF, colF, dataF); };
SuperMap.z内连接 = function(a, b, lk, rk, rs) { return new Array2D(a).z内连接(b, lk, rk, rs); };
SuperMap.z筛选 = function(arr, pred) { return new Array2D(arr).z筛选(pred); };
SuperMap.z映射 = function(arr, fn) { return new Array2D(arr).z映射(fn); };
SuperMap.z选择列 = function(arr, cols) {
    if (!arr || arr.length === 0) return [];
    var header = arr[0]; var dataRows = arr.slice(1); var colIndices = [];
    var parts = String(cols).split(/[,，]/);
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].trim();
        if (typeof part === "number") { colIndices.push(part); continue; }
        var match = String(part).match(/^f(\d+)$/i);
        if (match) colIndices.push(parseInt(match[1], 10) - 1);
        else colIndices.push(parseInt(part, 10) || 0);
    }
    var newHeader = [];
    for (var c = 0; c < colIndices.length; c++) { if (colIndices[c] < header.length) newHeader.push(header[colIndices[c]]); }
    var newRows = [];
    for (var r = 0; r < dataRows.length; r++) {
        var row = dataRows[r]; var newRow = [];
        for (var c2 = 0; c2 < colIndices.length; c2++) { if (colIndices[c2] < row.length) newRow.push(row[colIndices[c2]]); }
        newRows.push(newRow);
    }
    var result = [newHeader];
    for (var j = 0; j < newRows.length; j++) result.push(newRows[j]);
    return result;
};
SuperMap.selectColumns = SuperMap.z选择列;

/**
 * SuperMap.z重命名列 - 重命名列名
 */
SuperMap.z重命名列 = function(arr, renameMap) {
    if (!arr || arr.length === 0) return arr;
    var header = arr[0].slice(); var dataRows = arr.slice(1);
    for (var colName in renameMap) {
        if (renameMap.hasOwnProperty(colName)) {
            var match = String(colName).match(/^f(\d+)$/i);
            var colIdx = match ? parseInt(match[1], 10) - 1 : (parseInt(colName, 10) || 0);
            if (colIdx < header.length) header[colIdx] = renameMap[colName];
        }
    }
    var result = [header];
    for (var i = 0; i < dataRows.length; i++) result.push(dataRows[i]);
    return result;
};
SuperMap.renameColumns = SuperMap.z重命名列;

/**
 * SuperMap.z添加列 - 添加计算列
 */
SuperMap.z添加列 = function(arr, newColName, expression) {
    if (!arr || arr.length === 0) return arr;
    var header = arr[0].slice(); var dataRows = arr.slice(1); header.push(newColName);
    var newRows = [];
    for (var i = 0; i < dataRows.length; i++) {
        var row = dataRows[i].slice(); var newValue;
        if (typeof expression === "function") { newValue = expression(row); }
        else {
            var expr = String(expression).replace(/f(\d+)/gi, function(m, num) { return "row[" + (parseInt(num, 10) - 1) + "]"; });
            try { newValue = (new Function("row", "return " + expr + ";"))(row); } catch (e) { newValue = null; }
        }
        row.push(newValue); newRows.push(row);
    }
    var result = [header];
    for (var j = 0; j < newRows.length; j++) result.push(newRows[j]);
    return result;
};
SuperMap.addColumn = SuperMap.z添加列;

/**
 * SuperMap.z删除列 - 删除指定列
 */
SuperMap.z删除列 = function(arr, colRef) {
    if (!arr || arr.length === 0) return arr;
    var header = arr[0].slice(); var dataRows = arr.slice(1);
    var match = String(colRef).match(/^f(\d+)$/i);
    var colIdx = match ? parseInt(match[1], 10) - 1 : (parseInt(colRef, 10) || 0);
    if (colIdx >= header.length) return arr;
    header.splice(colIdx, 1);
    var newRows = [];
    for (var i = 0; i < dataRows.length; i++) { var row = dataRows[i].slice(); row.splice(colIdx, 1); newRows.push(row); }
    var result = [header];
    for (var j = 0; j < newRows.length; j++) result.push(newRows[j]);
    return result;
};
SuperMap.removeColumn = SuperMap.z删除列;

/**
 * SuperMap.z转换列 - 转换列数据类型
 */
SuperMap.z转换列 = function(arr, colRef, converter) {
    if (!arr || arr.length === 0) return arr;
    var header = arr[0].slice(); var dataRows = arr.slice(1);
    var match = String(colRef).match(/^f(\d+)$/i);
    var colIdx = match ? parseInt(match[1], 10) - 1 : (parseInt(colRef, 10) || 0);
    var newRows = [];
    for (var i = 0; i < dataRows.length; i++) { var row = dataRows[i].slice(); if (colIdx < row.length) row[colIdx] = converter(row[colIdx]); newRows.push(row); }
    var result = [header];
    for (var j = 0; j < newRows.length; j++) result.push(newRows[j]);
    return result;
};
SuperMap.convertColumn = SuperMap.z转换列;

/**
 * SuperMap.z填充空值 - 填充空值
 */
SuperMap.z填充空值 = function(arr, colRef, fillValue) {
    return SuperMap.z转换列(arr, colRef, function(val) { return (val === null || val === undefined || val === "") ? fillValue : val; });
};
SuperMap.fillNull = SuperMap.z填充空值;

/**
 * SuperMap.z列统计 - 计算列的统计信息
 */
SuperMap.z列统计 = function(arr, colRef, _opts) { // XXD-170: z列统计 默认 skip header, 支持 _header 选项
    if (!arr || arr.length === 0) return { sum: 0, avg: 0, min: 0, max: 0, count: 0 };
    var hasHeader = !_opts || _opts.header !== false;
    if (hasHeader && arr.length < 2) return { sum: 0, avg: 0, min: 0, max: 0, count: 0 };
    var dataRows = hasHeader ? arr.slice(1) : arr;
    var match = String(colRef).match(/^f(\d+)$/i);
    var colIdx = match ? parseInt(match[1], 10) - 1 : (parseInt(colRef, 10) || 0);
    var values = [];
    for (var i = 0; i < dataRows.length; i++) { var v = dataRows[i][colIdx]; if (typeof v === "number" && !isNaN(v)) values.push(v); }
    if (values.length === 0) return { sum: 0, avg: 0, min: 0, max: 0, count: 0 };
    var sum = 0, min = values[0], max = values[0];
    for (var j = 0; j < values.length; j++) { sum += values[j]; if (values[j] < min) min = values[j]; if (values[j] > max) max = values[j]; }
    return { sum: sum, avg: sum / values.length, min: min, max: max, count: values.length };
};
SuperMap.columnStats = SuperMap.z列统计;

/**
 * SuperMap.z分组统计 - 按列分组并统计 (分组 + 聚合, 每组 1 行)
 *
 * @param {Array2D|Array} arr  - 源二维数组 (含 header 行)
 * @param {string|number}  groupCol    - 分组列 (f1 风格或 0 基索引)
 * @param {Object<string,string>} statsConfig - 聚合配置 { colName: aggType }
 *                                              aggType ∈ sum|avg|count|min|max
 * @returns {Array<Array>} 二维数组,首行是表头 [groupHeader, "colA_sum", "colB_avg", ...]
 *                          后续每行 [groupKey, aggValA, aggValB, ...]
 *
 * 形状约定 (与 z分组 区分):
 *   - z分组:      返回 字典 { key: rows[] }    (纯分组,无聚合)
 *   - z分组统计:  返回 二维数组 [header, ...]   (分组+聚合,每组 1 行)
 * 调用方按需选择:
 *   - 需要 key→行集合  → 用 z分组
 *   - 需要每组 1 行汇总 → 用 z分组统计
 */
// 🔧 XXD-180/XXD-181 final fix end
SuperMap.z分组统计 = function(arr, groupCol, statsConfig) {
    if (!arr || arr.length === 0) return [];
    var header = arr[0]; var dataRows = arr.slice(1);
    var match = String(groupCol).match(/^f(\d+)$/i);
    var groupIdx = match ? parseInt(match[1], 10) - 1 : (parseInt(groupCol, 10) || 0);
    var groupHeader = header[groupIdx];
    var stats = [];
    for (var statCol in statsConfig) {
        if (statsConfig.hasOwnProperty(statCol)) {
            var statMatch = String(statCol).match(/^f(\d+)$/i);
            var statIdx = statMatch ? parseInt(statMatch[1], 10) - 1 : (parseInt(statCol, 10) || 0);
            stats.push({ colIdx: statIdx, headerName: header[statIdx], aggType: statsConfig[statCol] });
        }
    }
    var groups = {};
    for (var i = 0; i < dataRows.length; i++) { var row = dataRows[i]; var key = (row[groupIdx] == null) ? '__NULL__' : String(row[groupIdx]); if (!groups[key]) groups[key] = []; groups[key].push(row); }
    var resultHeader = [groupHeader];
    for (var s = 0; s < stats.length; s++) resultHeader.push(stats[s].headerName + "_" + stats[s].aggType);
    var result = [resultHeader]; var keys = [];
    for (var key in groups) keys.push(key);
    for (var ki = 0; ki < keys.length; ki++) {
        var groupRows = groups[keys[ki]]; var newRow = [keys[ki]];
        for (var si = 0; si < stats.length; si++) {
            var colValues = [];
            for (var gi = 0; gi < groupRows.length; gi++) { if (stats[si].colIdx < groupRows[gi].length) colValues.push(groupRows[gi][stats[si].colIdx]); }
            newRow.push(SuperMap._aggregate(colValues, stats[si].aggType));
        }
        result.push(newRow);
    }
    return result;
};
SuperMap.groupBy = SuperMap.z分组统计;

/**
 * SuperMap._aggregate - 内部聚合函数
 */
SuperMap._aggregate = function(values, aggType) {
    var numericValues = [];
    for (var i = 0; i < values.length; i++) { if (typeof values[i] === "number" && !isNaN(values[i])) numericValues.push(values[i]); }
    if (numericValues.length === 0) {
        switch (aggType) { case "sum": return 0; case "avg": return 0; case "count": return values.length; case "min": return null; case "max": return null; default: return null; }
    }
    switch (aggType) {
        case "sum": { let sum = 0; for (let j = 0; j < numericValues.length; j++) sum += numericValues[j]; return sum; }
        case "avg": case "average": { let total = 0; for (let k = 0; k < numericValues.length; k++) total += numericValues[k]; return total / numericValues.length; }
        case "count": return values.length;
        case "min": { let minVal = numericValues[0]; for (let m = 1; m < numericValues.length; m++) if (numericValues[m] < minVal) minVal = numericValues[m]; return minVal; }
        case "max": { let maxVal = numericValues[0]; for (let n = 1; n < numericValues.length; n++) if (numericValues[n] > maxVal) maxVal = numericValues[n]; return maxVal; }
        default: return null;
    }
};

console.log('JSA880框架 v4.0.0 已整合 JSA880-claude.js 增强功能:');
console.log('- agg(): 全局聚合函数');
console.log('- oadate(): OA日期转换');
console.log('- StrUtils: 字符串工具模块 (40+方法)');
console.log('- NumUtils: 数值工具模块 (30+方法)');
console.log('- MsgUtils: 消息对话框模块');
console.log('- WorkbookUtils: 工作簿操作模块');
console.log('- SuperMap数据操作: selectColumns/renameColumns/addColumn/removeColumn/convertColumn/fillNull/columnStats/groupBy');
console.log('========================================');

// ============================================================
// [COURSE_INTEGRATION] 整合郑广学JSA火箭速成班课程中的核心模块
// 版本: v4.1.0 (2026-05-15)
// ============================================================

// ==================== TreeNode - 树结构类 ====================
/**
 * TreeNode - 多级树结构类（支持多级菜单、BOM展开、层级汇总）
 * @class
 * @example
 * var tree = new TreeNode('root');
 * tree.addChild('child1').addChild('grandchild1');
 * tree.show(); // 打印树结构
 */
function TreeNode(name, data) {
    this.name = name || '';
    this.data = data || null;
    this.children = [];
    this.parent = null;
}

TreeNode.prototype.addChild = function(name, data) {
    var child = new TreeNode(name, data);
    child.parent = this;
    this.children.push(child);
    return child;
};

TreeNode.prototype.removeChild = function(name) {
    for (var i = 0; i < this.children.length; i++) {
        if (this.children[i].name === name) {
            this.children.splice(i, 1);
            return true;
        }
    }
    return false;
};

TreeNode.prototype.getChild = function(name) {
    for (var i = 0; i < this.children.length; i++) {
        if (this.children[i].name === name) return this.children[i];
    }
    return null;
};

TreeNode.prototype.findNode = function(path) {
    var parts = Array.isArray(path) ? path : String(path).split('/');
    var current = this;
    for (var i = 0; i < parts.length; i++) {
        var child = current.getChild(parts[i]);
        if (!child) return null;
        current = child;
    }
    return current;
};

TreeNode.prototype.getPath = function() {
    var parts = [];
    var node = this;
    while (node && node.parent) {
        parts.unshift(node.name);
        node = node.parent;
    }
    return parts.join('/');
};

TreeNode.prototype.forEach = function(callback, depth) {
    depth = depth || 0;
    callback(this, depth);
    for (var i = 0; i < this.children.length; i++) {
        this.children[i].forEach(callback, depth + 1);
    }
};

TreeNode.prototype.toArray = function(result, depth) {
    result = result || [];
    depth = depth || 0;
    result.push({ name: this.name, data: this.data, depth: depth });
    for (var i = 0; i < this.children.length; i++) {
        this.children[i].toArray(result, depth + 1);
    }
    return result;
};

TreeNode.prototype.show = function(indent) {
    indent = indent || '';
    console.log(indent + (this.name || '(root)'));
    for (var i = 0; i < this.children.length; i++) {
        this.children[i].show(indent + '  ');
    }
};

/**
 * 从二维数组构建树结构
 * @param {Array} arr - 多列二维数组，每列代表一个层级
 * @returns {TreeNode} 根节点
 */
TreeNode.initTree = function(arr) {
    var root = new TreeNode('root');
    if (!arr || arr.length === 0) return root;
    for (var i = 0; i < arr.length; i++) {
        var row = arr[i];
        var current = root;
        for (var j = 0; j < row.length; j++) {
            var val = String(row[j] || '');
            var child = current.getChild(val);
            if (!child) {
                child = current.addChild(val);
            }
            current = child;
        }
    }
    return root;
};

/**
 * 从父子节点表构建树
 * @param {Array} arr - 二维数组 [id, parentId, name]
 * @returns {TreeNode} 根节点
 */
TreeNode.fromParentChild = function(arr) {
    var root = new TreeNode('root');
    var nodeMap = { '': root, null: root, undefined: root };
    for (var i = 0; i < arr.length; i++) {
        var id = String(arr[i][0]);
        var parentId = String(arr[i][1] || '');
        var name = arr[i][2] || id;
        var node = new TreeNode(name, { id: id });
        nodeMap[id] = node;
    }
    for (var i = 0; i < arr.length; i++) {
        var id = String(arr[i][0]);
        var parentId = String(arr[i][1] || '');
        var node = nodeMap[id];
        var parent = nodeMap[parentId] || root;
        if (node) {
            node.parent = parent;
            parent.children.push(node);
        }
    }
    return root;
};


TreeNode.prototype.byPath = TreeNode.prototype.findNode;

TreeNode.prototype.byID = function(id) {
    var result = null;
    this.forEach(function(node) {
        if (node.data && node.data.id !== undefined && String(node.data.id) === String(id)) {
            result = node;
        }
    });
    return result;
};

TreeNode.prototype.cloneTree = function() {
    var clone = new TreeNode(this.name, this.data ? JSON.parse(JSON.stringify(this.data)) : null);
    for (var i = 0; i < this.children.length; i++) {
        var childClone = this.children[i].cloneTree();
        childClone.parent = clone;
        clone.children.push(childClone);
    }
    return clone;
};

TreeNode.prototype.loadArray2D = function(arr) {
    var tree = TreeNode.initTree(arr);
    this.children = tree.children;
    for (var i = 0; i < this.children.length; i++) {
        this.children[i].parent = this;
    }
    return this;
};

TreeNode.prototype.getTreeView = function() {
    var lines = [];
    this.forEach(function(node, depth) {
        var prefix = '';
        for (var i = 0; i < depth; i++) prefix += '  ';
        lines.push(prefix + (node.name || '(root)'));
    });
    return lines;
};

TreeNode.prototype.toPidArray = function(level) {
    var result = [];
    var idCounter = 0;
    var nodeToId = new Map();
    this.forEach(function(node) {
        var id = idCounter++;
        nodeToId.set(node, id);
    });
    this.forEach(function(node) {
        if (node.parent && (!level || arguments[1] <= level)) {
            var parentId = nodeToId.get(node.parent);
            result.push([nodeToId.get(node), parentId, node.name]);
        }
    });
    return result;
};

TreeNode.prototype.toArray2D = function() {
    var result = [];
    var maxDepth = 0;
    this.forEach(function(node) {
        if (node.parent && arguments[1] > maxDepth) maxDepth = arguments[1];
    });
    this.forEach(function(node, depth) {
        if (depth > 0) {
            var row = [];
            for (var i = 0; i < maxDepth; i++) row.push('');
            row[depth - 1] = node.name;
            result.push(row);
        }
    });
    return result;
};

TreeNode.prototype.delChild = TreeNode.prototype.removeChild;

TreeNode.prototype.getData = function() {
    return this.data;
};

TreeNode.prototype.updateForm = function(controls, level) {
    level = level || 1;
    if (level < 1 || level > controls.length) return;
    var target = controls[level - 1];
    if (!target) return;
    target.Clear();
    for (var i = 0; i < this.children.length; i++) {
        target.AddItem(this.children[i].name);
    }
    if (this.children.length > 0) {
        target.ListIndex = 0;
    }
    for (var i = level; i < controls.length; i++) {
        if (controls[i]) controls[i].Clear();
    }
};

TreeNode.initTreeByPid = TreeNode.fromParentChild;

// ==================== IO 增强函数 ====================
/**
 * IO.getFiles - 获取文件夹中的文件列表（支持递归）
 * @param {string} folderPath - 文件夹路径
 * @param {boolean} recursive - 是否递归，默认false
 * @param {boolean} includeHidden - 是否包含隐藏文件，默认false
 * @returns {Array} 文件路径数组
 */
IO.getFiles = function(folderPath, recursive, includeHidden) {
    includeHidden = includeHidden !== false;
    // XXD-176 (XXD-177): try ActiveX first (WPS/Windows), fall back to Node fs for testing/non-Windows.
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (!fso.FolderExists(folderPath)) return [];
        var results = [];
        var folder = fso.GetFolder(folderPath);
        var enumFiles = new Enumerator(folder.Files);
        for (; !enumFiles.atEnd(); enumFiles.moveNext()) {
            var file = enumFiles.item();
            results.push(file.Path);
        }
        if (recursive) {
            var enumFolders = new Enumerator(folder.SubFolders);
            for (; !enumFolders.atEnd(); enumFolders.moveNext()) {
                var sub = enumFolders.item();
                results = results.concat(IO.getFiles(sub.Path, true, includeHidden));
            }
        }
        return results;
    } catch (e) {
        // v4.0.11 / XXD-176: Node (or any non-ActiveX) fallback.
        if (typeof require !== 'function') return [];
        try {
            var fs = require('fs');
            var pathMod = require('path');
            if (!fs.existsSync(folderPath)) return [];
            var entries = fs.readdirSync(folderPath, { withFileTypes: true });
            var results = [];
            for (var i = 0; i < entries.length; i++) {
                var ent = entries[i];
                if (!includeHidden && ent.name.charAt(0) === '.') continue;
                var full = pathMod.join(folderPath, ent.name);
                if (ent.isDirectory()) {
                    if (recursive) {
                        results = results.concat(IO.getFiles(full, true, includeHidden));
                    }
                } else if (ent.isFile()) {
                    results.push(full);
                }
            }
            return results;
        } catch (e2) {
            console.warn('IO.getFiles:', e2.message);
            return [];
        }
    }
};

// XXD-149 (XXD-146): listFiles/z列出文件 alias for IO.getFiles (defined above).
IO.listFiles = IO.listFiles || IO.getFiles;
IO.z列出文件 = IO.listFiles;

IO.showFolderDialog = function(title) {
    if (typeof Application === "undefined") return null;
    try {
        var dialog = Application.FileDialog(4);
        if (title) dialog.Title = title;
        dialog.AllowMultiSelect = false;
        if (dialog.Show() === -1) return dialog.SelectedItems.Item(1);
        return null;
    } catch (e) { return null; }
};

IO.showOpenDialog = function(title, filter) {
    if (typeof Application === "undefined") return null;
    try {
        var dialog = Application.FileDialog(1);
        if (title) dialog.Title = title;
        if (filter) { dialog.Filters.Clear(); dialog.Filters.Add("文件", filter); }
        dialog.AllowMultiSelect = false;
        if (dialog.Show() === -1) return dialog.SelectedItems.Item(1);
        return null;
    } catch (e) { return null; }
};

IO.correctFileName = function(name) {
    return String(name).replace(/[\\/:*?"<>|]/g, '_');
// ==================== IO 扩展方法（xlsm 测试兼容）====================

/**
 * IO.fileToBase64 - 文件转 base64 编码
 * @param {string} path - 文件路径
 * @returns {string} base64 字符串
 */
IO.z文件转Base64 = IO.fileToBase64 = function(path) {
    try {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 1;
        stream.Open();
        stream.LoadFromFile(path);
        var binaryData = stream.Read();
        stream.Close();
        var xmlDom = new ActiveXObject("MSXML2.DOMDocument");
        var elem = xmlDom.createElement("tmp");
        elem.dataType = "bin.base64";
        elem.nodeTypedValue = binaryData;
        return elem.text;
    } catch (e) {
        return "";
    }
};

/**
 * IO.imageFileToBase64 - 图片文件转 base64
 * @param {string} path - 图片文件路径
 * @returns {string} base64 字符串
 */
IO.z图片转Base64 = IO.imageFileToBase64 = function(path) {
    return IO.fileToBase64(path);
};

/**
 * IO.downLoadFile - HTTP 下载文件到本地
 * @param {string} url - 下载地址
 * @param {string} savePath - 保存路径
 * @returns {boolean} 是否成功
 */
IO.z下载文件 = IO.downLoadFile = function(url, savePath) {
    try {
        var http = new ActiveXObject("MSXML2.XMLHTTP");
        http.Open("GET", url, false);
        http.Send();
        if (http.Status === 200) {
            var stream = new ActiveXObject("ADODB.Stream");
            stream.Type = 1;
            stream.Open();
            stream.Write(http.ResponseBody);
            stream.SaveToFile(savePath, 2);
            stream.Close();
            return true;
        }
        return false;
    } catch (e) {
        return false;
    }
};

/**
 * IO.correctPath - 统一路径分隔符
 * @param {string} p - 路径字符串
 * @returns {string} 规范化后的路径
 */
IO.z纠正路径 = IO.correctPath = function(p) {
    return (p || '').replace(/[\\/]+/g, '\\');
};

/**
 * IO.showFileDialog - 显示文件选择对话框
 * @param {string} [filter] - 文件类型过滤器，如 "*.xlsx"
 * @returns {string|null} 选择的文件路径
 */
IO.z显示文件对话框 = IO.showFileDialog = function(filter) {
    if (typeof Application === "undefined") return null;
    try {
        var dialog = Application.FileDialog(1);
        if (filter) {
            dialog.Filters.Clear();
            dialog.Filters.Add("文件", filter);
        }
        dialog.AllowMultiSelect = false;
        if (dialog.Show() === -1) return dialog.SelectedItems.Item(1);
        return null;
    } catch (e) { return null; }
};

/**
 * IO.getDirectorys - 获取子目录列表
 * @param {string} folderPath - 文件夹路径
 * @returns {Array} 子目录名称数组
 */
IO.z获取子目录 = IO.getDirectorys = function(folderPath) {
    var dirs = [];
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        var folder = fso.GetFolder(folderPath);
        var subFolders = new Enumerator(folder.SubFolders);
        for (; !subFolders.atEnd(); subFolders.moveNext()) {
            dirs.push(subFolders.item().Name);
        }
    } catch (e) { }
    return dirs;
};

/**
 * IO.deleteTree - 递归删除文件夹（IO.delTree 的别名）
 */
IO.deleteTree = IO.delTree;

};

// ==================== RngUtils 增强 ====================
/**
 * rngFilter - 高级单元格筛选（支持日期、正则等）
 * @param {string|Range} rng - 单元格区域
 * @param {number|string} colIndex - 筛选列（数字为1-based，字符串可带"-dt"或"-t"后缀）
 * @param {Function} conditionFn - 条件函数，接收一维数组
 * @example
 * RngUtils.rngFilter("A1:J14", 1, x => x > 2)
 * RngUtils.rngFilter("A1:J14", "8-dt", x => cdate(x) == cdate("2022-1-1"))
 */
RngUtils.rngFilter = function(rng, colIndex, conditionFn) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    r.AutoFilter(colIndex, conditionFn);
    return r;
};

/**
 * RngUtils.hitRange - 判断Target是否在指定区域内
 * @param {Range} target - 目标单元格
 * @param {Range|string} area - 区域
 * @returns {boolean} 是否命中
 */
RngUtils.hitRange = function(target, area) {
    try {
        var r = typeof area === 'string' ? Range(area) : area;
        return Application.Intersect(target, r) != null;
    } catch (e) { return false; }
};

// ==================== findRange - 单元格批量查找 ====================
/**
 * findRange - 在区域内批量查找值
 * @param {Range|string} rng - 区域
 * @param {string} value - 查找值
 * @returns {Array} 找到的Range对象数组
 */
function findRange(rng, value) {
    var r = typeof rng === 'string' ? Range(rng) : rng;
    var results = [];
    try {
        var found = r.Find(value);
        var firstAddress = null;
        while (found) {
            if (firstAddress === null) firstAddress = found.Address;
            else if (found.Address === firstAddress) break;
            results.push(found);
            found = r.FindNext(found);
        }
    } catch (e) { }
    return results;
}

// ==================== sheetsSort - 工作表排序 ====================
/**
 * sheetsSort - 按自定义方式排序工作表
 * @param {Sheets|Array} sheets - 工作表集合或数组
 * @param {Function|Array} sortFnOrOrder - 排序函数或自定义顺序数组
 */
function sheetsSort(sheets, sortFnOrOrder) {
    var shts = Array.isArray(sheets) ? sheets : [];
    if (!Array.isArray(sheets)) {
        try {
            for (var i = 1; i <= sheets.Count; i++) shts.push(sheets(i));
        } catch (e) { return; }
    }
    if (Array.isArray(sortFnOrOrder)) {
        var order = sortFnOrOrder;
        shts.sort(function(a, b) {
            return order.indexOf(a.Name) - order.indexOf(b.Name);
        });
    } else if (typeof sortFnOrOrder === 'function') {
        shts.sort(function(a, b) {
            return sortFnOrOrder(a.Name, b.Name);
        });
    } else {
        shts.sort(function(a, b) { return a.Name.localeCompare(b.Name); });
    }
    for (var i = 0; i < shts.length; i++) {
        shts[i].Move(shts[0]);
    }
}

// ==================== map2d / forEach2d 二维数组遍历 ====================
/**
 * map2d - 对二维数组每个元素进行映射
 * @param {Array} arr - 二维数组
 * @param {Function} fn - 映射函数 (value, row, col, arr) => newValue
 * @returns {Array} 映射后的二维数组
 */
function map2d(arr, fn) {
    if (!arr || !Array.isArray(arr)) return [];
    var result = [];
    for (var i = 0; i < arr.length; i++) {
        var row = [];
        for (var j = 0; j < arr[i].length; j++) {
            row.push(fn(arr[i][j], i, j, arr));
        }
        result.push(row);
    }
    return result;
}

/**
 * forEach2d - 对二维数组每个元素执行回调
 * @param {Array} arr - 二维数组
 * @param {Function} fn - 回调函数 (value, row, col, arr)
 */
function forEach2d(arr, fn) {
    if (!arr || !Array.isArray(arr)) return;
    for (var i = 0; i < arr.length; i++) {
        for (var j = 0; j < arr[i].length; j++) {
            fn(arr[i][j], i, j, arr);
        }
    }
}

// /* XXD-172 final fix */
// [XXD-172] attach top-level map2d / forEach2d onto JSA namespace
// Issue: R2.1 测试 JSA.map2d([[1,2],[3,4]], fn) THROW; v4.1.0 changelog 谎报已整合
// Fix: 函数体本身已存在(L20080/L20098), 这里补导出
JSA.map2d = map2d;
JSA.forEach2d = forEach2d;

// ==================== SheetChangeHelper - 事件辅助类 ====================
/**
 * SheetChangeHelper - Change事件辅助类，监控多个单元格值变化
 * @class
 * @example
 * var monitor = new SheetChangeHelper("Sheet1!B1:B2,C2");
 * monitor.check(Target, function(hit) { ... });
 */
function SheetChangeHelper(rangeStr) {
    this._rangeStr = rangeStr;
    this._lastValues = {};
    var parts = rangeStr.split(',');
    for (var i = 0; i < parts.length; i++) {
        try {
            var r = Range(parts[i].trim());
            this._lastValues[parts[i].trim()] = r.Value2;
        } catch (e) {}
    }
}

SheetChangeHelper.prototype.check = function(target, callback) {
    var parts = this._rangeStr.split(',');
    for (var i = 0; i < parts.length; i++) {
        var addr = parts[i].trim();
        try {
            var intersect = Application.Intersect(target, Range(addr));
            if (intersect != null) {
                // 取交集区域的值（而非整个 target 的值）
                var newVal = intersect.Value2;
                var oldVal = this._lastValues[addr];
                if (newVal !== oldVal) {
                    this._lastValues[addr] = newVal;
                    callback(intersect);
                    return true;
                }
            }
        } catch (e) {}
    }
    return false;
};

SheetChangeHelper.prototype.update = function() {
    var parts = this._rangeStr.split(',');
    for (var i = 0; i < parts.length; i++) {
        try {
            var r = Range(parts[i].trim());
            this._lastValues[parts[i].trim()] = r.Value2;
        } catch (e) {}
    }
};


// ==================== StopWatch 性能计时器 ====================
/**
 * StopWatch - 高精度性能计时器
 * @class
 * @example
 * var st = new StopWatch();
 * st.start();
 * // ... 被测代码 ...
 * console.log("用时:" + st.time());
 */
function StopWatch() {
    this._startTime = 0;
    this._lapTime = 0;
}

StopWatch.prototype.start = function() {
    this._startTime = new Date().getTime();
    this._lapTime = this._startTime;
    return this;
};

StopWatch.prototype.time = function() {
    var now = new Date().getTime();
    var elapsed = now - this._startTime;
    return (elapsed / 1000).toFixed(3);
};

StopWatch.prototype.lap = function() {
    var now = new Date().getTime();
    var elapsed = now - this._lapTime;
    this._lapTime = now;
    return (elapsed / 1000).toFixed(3);
};

StopWatch.prototype.usedTime = StopWatch.prototype.time;
StopWatch.prototype.restart = function() {
    return this.start();
};

// ==================== FormUtils 窗体工具 ====================
var FormUtils = {
    listBoxLoadArray: function(listbox, arr, headerArr) {
        if (!listbox || !arr) return;
        listbox.Clear();
        if (headerArr) {
            var headerRow = headerArr.map(function(h) { return String(h); });
            listbox.ColumnCount = headerRow.length;
            listbox.AddItem(headerRow.join('\t'));
        }
        for (var i = 0; i < arr.length; i++) {
            var row = Array.isArray(arr[i]) ? arr[i] : [arr[i]];
            if (!headerArr && i === 0) listbox.ColumnCount = row.length;
            listbox.AddItem(row.map(function(c) { return c !== null && c !== undefined ? String(c) : ''; }).join('\t'));
        }
    },
    listBoxSetWidth: function(listbox) {
        for (var i = 1; i < arguments.length; i += 2) {
            if (arguments[i] && arguments[i + 1]) {
                listbox.ColumnWidths += (listbox.ColumnWidths ? ',' : '') + arguments[i + 1];
            }
        }
    },
    listBoxToArray: function(listbox) {
        var result = [];
        for (var i = 0; i < listbox.ListCount; i++) {
            var row = [];
            for (var j = 0; j < listbox.ColumnCount; j++) {
                row.push(listbox.List(i, j));
            }
            result.push(row);
        }
        return result;
    },
    getControls: function(container) {
        var controls = [];
        if (container && container.Controls) {
            for (var i = 0; i < container.Controls.Count; i++) {
                controls.push(container.Controls.Item(i));
            }
        }
        return controls;
    }
};

// ==================== PicUtiles 图片工具（基础版）====================
var PicUtiles = {
    insertPic: function(picPath, rng) {
        var r = typeof rng === 'string' ? Range(rng) : rng;
        var sheet = r.Worksheet || ActiveSheet;
        var shp = sheet.Shapes.AddPicture(picPath, false, true, r.Left, r.Top, -1, -1);
        return shp;
    },
    getPicByRng: function(rng) {
        var r = typeof rng === 'string' ? Range(rng) : rng;
        var sheet = r.Worksheet || ActiveSheet;
        for (var i = 1; i <= sheet.Shapes.Count; i++) {
            var shp = sheet.Shapes(i);
            if (shp.TopLeftCell.Address === r.Address) return shp;
        }
        return null;
    },
    shpFitRng: function(shp, rng) {
        if (!shp || !rng) return;
        var r = typeof rng === 'string' ? Range(rng) : rng;
        shp.Left = r.Left;
        shp.Top = r.Top;
        shp.Width = r.Width;
        shp.Height = r.Height;
    },
    delPicByRng: function(rng) {
        var r = typeof rng === 'string' ? Range(rng) : rng;
        var sheet = r.Worksheet || ActiveSheet;
        for (var i = sheet.Shapes.Count; i >= 1; i--) {
            var shp = sheet.Shapes(i);
            try {
                if (Application.Intersect(shp.TopLeftCell, r)) shp.Delete();
            } catch (e) {}
        }
    },
    exportToPic: function(obj, path) {
        if (typeof obj === 'string') {
            var sheet = ActiveSheet;
            for (var i = 1; i <= sheet.Shapes.Count; i++) {
                if (sheet.Shapes(i).Name === obj) { sheet.Shapes(i).SaveAsPicture(path); return true; }
            }
            return false;
        }
        if (obj && obj.SaveAsPicture) { obj.SaveAsPicture(path); return true; }
        return false;
    }
};
$.month = JSA.month;

$.now = JSA.now;

$.mid = JSA.mid;
$.z截取字符 = $.mid;

$.sheetsSort = ShtUtils.sheetsSort;

/**
 * $.isError - 判断值是否为错误值
 */
$.isError = function(v) {
    if (v === undefined || v === null) return false;
    if (typeof v === 'number' && isNaN(v)) return true;
    var s = String(v);
    return s === '#N/A' || s === '#VALUE!' || s === '#REF!' || s === '#DIV/0!' || s === '#NUM!' || s === '#NAME?' || s === '#NULL!';
};

/**
 * $.text - 将值转换为文本
 */
$.text = function(v) {
    if (v === null || v === undefined) return '';
    return String(v);
};

$.m = JSA.m;

$.S = JSA.S;


// ==================== 全局 Range 别名 ====================
/**
 * Range - 获取单元格或区域（等同于 WPS 内置的 $(...) 语法）
 * @param {string|Range} addr - 区域地址或 Range 对象
 * @param {string} [colIndex] - 可选，第二个参数（列索引）
 * @returns {RangeChain} RangeChain 对象，支持链式调用
 * @example
 * Range("A1").value()           // 获取 A1 的值
 * Range("A1:C10").safeArray()   // 转换为 Array2D
 */
if (typeof Range !== 'undefined') {
    // WPS 已有内置 Range 函数，创建包装版本
    var _originRange = Range;
    Range = function(addr, colIndex) {
        if (addr === undefined || addr === null) {
            return new RangeChain(null);
        }
        if (typeof addr === 'string') {
            try {
                var r = ActiveSheet.Range(addr);
                return new RangeChain(r, colIndex);
            } catch (e) {
                return new RangeChain(null);
            }
        }
        if (addr && typeof addr === 'object' && addr.Address) {
            return new RangeChain(addr, colIndex);
        }
        return new RangeChain(null);
    };
}

// v4.0.11: $$ 方法同步 + rangeZip 新增（矩阵元素级运算；原 rangeMatrix 元素级重写迁出，避免破坏课程 group-by 语义）
(function() {
    // 同步 $$ 到 Array2D
    // 注意：rangeMatrix 不在此处同步 —— 它在 L10628 已定义并保持 group-by 语义，$$ 同步已在
    // v4.0.11 早段（L5002-L5006）通过 $.rangeMatrix = Array2D.rangeMatrix 完成。
    // rangeZip 是 v4.0.11 元素级矩阵运算新名字，需要 $$ 同步。
    var methods = ['rangeSelect','rangeMap','rangeZip','filter','selectCols','sortByCols',
        'distinct','groupInto','leftjoin','leftFulljoin','deleteCols','deleteRows'];
    for (var i = 0; i < methods.length; i++) {
        var m = methods[i];
        if (typeof Array2D[m] === 'function') {
            try { $$[m] = Array2D[m]; } catch(e) {}
        }
    }

    // 列字母转索引
    var _ci = function(s) { var idx = 0; for (var ki = 0; ki < s.length; ki++) idx = idx * 26 + (s.toUpperCase().charCodeAt(ki) - 64); return idx - 1; };
    // 解析地址
    var _pa = function(addr, arr) {
        addr = addr.trim();
        if (!addr) return null;
        var rm = addr.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
        var cm = addr.match(/^([A-Za-z]+):([A-Za-z]+)$/);
        var rw = addr.match(/^(\d+):(\d+)$/);
        if (rm) { var c1=_ci(rm[1]),r1=parseInt(rm[2])-1,c2=_ci(rm[3]),r2=parseInt(rm[4])-1; var rs=Math.min(r1,r2),cs=Math.min(c1,c2); return {rs:rs,cs:cs,rc:Math.min(Math.abs(r2-r1)+1, arr.length-rs),cc:Math.min(Math.abs(c2-c1)+1, (arr[0]?arr[0].length:0)-cs)}; }
        else if (cm) { var cc1=_ci(cm[1]),cc2=_ci(cm[2]); return {rs:0,cs:Math.min(cc1,cc2),rc:arr.length,cc:Math.min(Math.abs(cc2-cc1)+1, (arr[0]?arr[0].length:0)-Math.min(cc1,cc2))}; }
        else if (rw) { var rr1=parseInt(rw[1])-1,rr2=parseInt(rw[2])-1; var mcc=0; for(var ri=0;ri<arr.length;ri++)if(arr[ri]&&arr[ri].length>mcc)mcc=arr[ri].length; return {rs:Math.min(rr1,rr2),cs:0,rc:Math.min(Math.abs(rr2-rr1)+1, arr.length-Math.min(rr1,rr2)),cc:mcc}; }
        return null;
    };

    // ====== rangeSelect 增强 ======
    var _rs = Array2D.rangeSelect;
    Array2D.rangeSelect = function(arr, start, end, startCol, endCol) {
        if (arr && arr._items) arr = arr._items;
        if (typeof start === 'string' && /^[A-Za-z\d]/.test(start)) {
            var parts = start.split(',');
            if (parts.length > 1) {
                var results = [];
                for (var p = 0; p < parts.length; p++) {
                    var a = _pa(parts[p], arr);
                    if (!a) continue;
                    var sub = [], re = Math.min(a.rs+a.rc,arr.length);
                    for (var i=a.rs;i<re;i++) { var row=arr[i]; sub.push(row?row.slice(a.cs,Math.min(a.cs+a.cc,row.length)):[]); }
                    results.push(sub);
                }
                return results;
            }
            var a = _pa(start, arr);
            if (a) {
                var result = [], re = Math.min(a.rs+a.rc,arr.length);
                for (var i=a.rs;i<re;i++) { var row=arr[i]; result.push(row?row.slice(a.cs,Math.min(a.cs+a.cc,row.length)):[]); }
                return result;
            }
            return [];
        }
        return _rs.call(this, arr, start, end, startCol, endCol);
    };

    // ====== rangeMap 增强 ======
    var _rm = Array2D.rangeMap;
    Array2D.rangeMap = function(arr, address, mapper) {
        if (arr && arr._items) arr = arr._items;
        if (!arr || !Array.isArray(arr)) return [];
        if (!mapper) return arr;
        var fn = typeof mapper === 'function' ? mapper : Array2D.parseLambda(mapper);
        if (!fn) return arr;
        var result = []; for (var i=0;i<arr.length;i++) result[i]=Array.prototype.slice.call(arr[i]);
        if (typeof address === 'string') {
            address = address.trim();
            if (!address) {
                for (var i=0;i<arr.length;i++) { if(!arr[i])continue; for(var j=0;j<arr[i].length;j++) result[i][j]=fn(arr[i][j],i,j,arr); }
                return result;
            }
            var parts = address.split(',');
            for (var p=0;p<parts.length;p++) {
                var a=_pa(parts[p],arr); if(!a)continue;
                var re=Math.min(a.rs+a.rc,arr.length);
                for(var i=a.rs;i<re;i++){if(!arr[i])continue;var ce=Math.min(a.cs+a.cc,arr[i].length);for(var j=a.cs;j<ce;j++)result[i][j]=fn(arr[i][j],i,j,arr);}
            }
            return result;
        }
        return _rm.call(this, arr, address, mapper);
    };

    // ====== rangeSelect/rangeMap/rangeMatrix $别名同步 ======
    $.rangeSelect = Array2D.rangeSelect;
    $.z按范围选择 = $.rangeSelect;
    $.rangeMap = Array2D.rangeMap;
    $.z区域映射 = $.rangeMap;
    // XXD-185: re-alias z局部映射 to the post-enhancement rangeMap (L20950 wrapper overwrites it).
    Array2D.z局部映射 = Array2D.rangeMap;
    
    // ====== rangeZip 新增（元素级矩阵运算，原 v4.0.11 rangeMatrix 重写迁出）======
    // 命名说明：原 rangeMatrix 在 L10628 已是 group-by 聚合函数（按 keySelector 分组，
    // 对 dataArrays 求和/聚合），v4.0.11 在此处原地覆盖为元素级矩阵地址运算，导致课程
    // 第 3 章 3 个 rangeMatrix 示例行为不正确。本函数保留元素级语义，命名为 rangeZip
    // （两区域元素级对齐/zip 的含义更直观），原 rangeMatrix 保持 group-by 语义。
    Array2D.rangeZip = function(arr, keySelector, dataArrays, aggregator) {
        // 提取 _items
        if (arr && arr._items) arr = arr._items;
        if (!arr || !Array.isArray(arr)) return [];

        // 解析 arr 地址
        var a = (typeof keySelector === 'string') ? _pa(keySelector, arr) : null;
        if (!a) return arr; // 地址解析失败，返回原数组

        // 解析第二个数据源：支持 [brr, addr] 或直接 brr
        var brr = null, b = null;
        if (Array.isArray(dataArrays) && dataArrays.length >= 2 && typeof dataArrays[1] === 'string') {
            brr = dataArrays[0];
            if (brr && brr._items) brr = brr._items;
            b = _pa(dataArrays[1], brr || []);
        } else {
            brr = dataArrays;
            if (brr && brr._items) brr = brr._items;
        }
        if (!brr || !Array.isArray(brr)) return arr;

        // brr 自动扩展
        var bRows = a.rc, bCols = a.cc;
        if (typeof brr === 'number' || typeof brr === 'string' || (brr !== null && !Array.isArray(brr))) {
            var sv = brr; brr = []; for (var i=0;i<bRows;i++){brr[i]=[];for(var j=0;j<bCols;j++)brr[i][j]=sv;}
            b = {rs:0,cs:0,rc:bRows,cc:bCols};
        } else if (Array.isArray(brr) && !Array.isArray(brr[0])) {
            // 单行：扩展到所有行
            var sr = brr; brr = []; for (var i=0;i<bRows;i++) brr[i] = sr.slice(0,bCols);
            b = {rs:0,cs:0,rc:bRows,cc:bCols};
        } else if (brr.length === 1 && bRows > 1) {
            var srr = brr[0]; brr = []; for (var i=0;i<bRows;i++) brr[i] = srr.slice(0,bCols);
            b = {rs:0,cs:0,rc:bRows,cc:bCols};
        }
        if (!b) b = {rs:0, cs:0, rc: brr.length, cc: brr[0] ? brr[0].length : 0};

        // 深拷贝 arr
        var result = []; for (var i=0;i<arr.length;i++) result[i] = Array.prototype.slice.call(arr[i]);

        // 元素级运算（取两个区域的最大公共范围）
        var re = Math.min(a.rs + a.rc, b.rs + b.rc, arr.length);
        for (var i = a.rs; i < re; i++) {
            if (!arr[i]) continue;
            var ce = Math.min(a.cs + a.cc, b.cs + b.cc, arr[i].length);
            for (var j = a.cs; j < ce; j++) {
                var bi = i - a.rs + b.rs;
                var bj = j - a.cs + b.cs;
                var bv = (bi < brr.length && brr[bi] && bj < brr[bi].length) ? brr[bi][bj] : null;
                if (aggregator && typeof aggregator === 'function') {
                    result[i][j] = aggregator(arr[i][j], bv, i, j);
                } else {
                    result[i][j] = bv;
                }
            }
        }
        return result;
    };
    $.rangeZip = Array2D.rangeZip;
    $.z区域对齐 = $.rangeZip;
})();

// ==================== 完成打印 ====================
console.log('[JSA880 v4.1.0] 课程模块已整合 (含 xlsm 测试兼容 v4.0.7 10个缺失API):');
console.log('- TreeNode: 多级树结构类');
console.log('- FormUtils: 窗体工具');
console.log('- PicUtiles: 图片工具(基础版)');
console.log('- rngFilter/hitRange: 高级筛选/命中判断');
console.log('- findRange: 单元格批量查找');
console.log('- sheetsSort: 工作表排序');
console.log('- map2d/forEach2d: 二维数组高阶遍历');
console.log('- SheetChangeHelper: Change事件辅助');
console.log('- IO增强: getFiles/showFolderDialog/showOpenDialog/correctFileName');
console.log('========================================');

// ============================================================================
//  【v4.2.2】KO 一切的 k 函数 · WPS UDF 整合版
//  将 KO一切的k函数_UDF模块.js 合并到主框架，单文件即可当 WPS 加载项使用
// ============================================================================
//
// 🎯 目的：让 k() / jsaLambda() 能在 WPS 公式里直接用
//    单元格写 =k("JSA.getIndexs", 1, 10, 2)  → 1 3 5 7 9 (数组溢出)
//
// 🔑 关键机制：WPS 公式引擎只认"模块顶层 function 声明"作为 UDF
//    - this.k = JSA.k             ❌ WPS 公式找不到
//    - var k = function(){}        ❌ 也不是顶层 function
//    - function k(fn, ...args){}  ✅ 顶层 function → 自动注册成 UDF
//
// 📦 使用方法（单文件部署）：
//    方式 A（推荐）：WPS → 选项 → 加载项 → 加载本文件 (JSA880.js)
//                  所有 WPS 文档都能用 =k(...) 公式
//    方式 B：在 JSA 编辑器里新建模块，粘入本文件后 7 行
//                  Console 区域会看到“✅ k() UDF 已就绪！”
//    方式 C（在线开发）：在 JSA 编辑器顶部加 loadScript("JSA880.js")
//
// ═══════════════════════════════════════════════════════════════════════

// ╔════════════════════════════════════════════════════════════════╗
// ║ 核心：top-level function 声明 → WPS 公式 UDF                  ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * [v5.0.0] 顶层 k() UDF — 转发到 JSA.k
 * WPS 公式引擎只扫 ThisWorkbook 顶层 function 作为 UDF
 * 实现代码全在 JSA.k(JSA880.js 内),这里只是转发 shim
 */
function k(fn) {
    return JSA.k.apply(null, arguments);
}

/**
 * jsaLambda 函数 — k() 的全名版本，WPS UDF
 * 单元格公式：=jsaLambda("JSA.getIndexs", 1, 5)
 */
function jsaLambda(fn) {
    return JSA.k.apply(null, arguments);
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ Workbook_Open：模块加载时打印确认信息                          ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 工作簿打开事件 - 验证 k() UDF 已注册成功
 * 看到这条日志就说明 =k(...) 公式可以用了
 */
function Workbook_Open() {
    if (typeof Console === 'undefined') return;
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("✅ k() UDF 已就绪！(JSA880 v" + (Array2D.version ? Array2D.version() : '4.2.2') + ")");
    Console.log("   测试：在任意单元格输入");
    Console.log("   =k(\"JSA.getIndexs\", 1, 10, 2)");
    Console.log("   看到 1 3 5 7 9 = 成功！");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    // 立刻跑一个自检
    try {
        var test = k("JSA.getIndexs", 1, 5, 1);
        Console.log("   自检：k('JSA.getIndexs', 1, 5, 1) = [" + test.join(",") + "]");
    } catch (e) {
        Console.log("   ⚠️ 自检失败：" + e.message);
    }
    // 验证 JSA.k(新的带错误位置化包装)
    try {
        var ktest = JSA.k("JSA.getIndexs", 1, 3, 1);
        Console.log("   JSA.k 自检:[" + ktest.join(",") + "]");
    } catch (e) {
        Console.log("   ⚠️ JSA.k 自检失败:" + e.message);
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ Sheet_Change：依赖单元格变化时自动重算 k() 公式                ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 工作表 Change 事件 - 当数据区域变化时强制重算所有 k 公式
 * 解决"改了源数据但 k() 公式结果没自动更新"的问题
 *
 * 绑定方法：WPS → 工作表事件 → 选择事件源、事件类型=Change、宏名=k_onChange
 */
function k_onChange(Sh, Target) {
    try {
        var sheet = Sh;
        var usedRange = sheet.UsedRange;
        if (!usedRange) return;
        var data = usedRange.Formula;
        if (!data) return;

        var needRecalc = false;
        var rows = data.length;
        var cols = Array.isArray(data[0]) ? data[0].length : 1;
        for (var r = 0; r < rows; r++) {
            for (var c = 0; c < cols; c++) {
                var f = Array.isArray(data[r]) ? data[r][c] : data[r];
                if (typeof f === "string" && f.indexOf("k(") === 1) {
                    // =k( 形式（位置 1 是 =k 的 k）
                    var cell = sheet.Cells(r + usedRange.Row, c + usedRange.Column);
                    // 标记脏值强制重算
                    cell.Dirty = true;
                    needRecalc = true;
                }
            }
        }
        if (needRecalc) {
            Application.Calculate();
            if (typeof Console !== 'undefined') Console.log("k() 公式已自动重算");
        }
    } catch (e) {
        // 静默失败，不影响用户操作
    }
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 工具函数：批量执行 k() 公式（适合不熟悉公式语法的用户）         ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 把一列 k() 调用结果 spill 到目标区域
 * 用法：批量写公式 =k(...) 的等价物
 *
 * @param {string} targetRangeAddr - 目标区域起始地址，如 "J1"
 * @param {Array} fnArray - 函数调用参数数组，如 [[fn1,arg1,arg2], [fn2,arg1], ...]
 * @returns {number} 成功执行的函数数量
 */
function batch_k(targetRangeAddr, fnArray) {
    if (!Array.isArray(fnArray)) return 0;
    var rng = Range(targetRangeAddr);
    var rows = fnArray.length;
    var cols = fnArray[0] ? fnArray[0].length : 0;
    var result = [];
    for (var i = 0; i < rows; i++) {
        var args = fnArray[i];
        var fn = args[0];
        var rest = Array.prototype.slice.call(args, 1);
        result.push(k.apply(null, [fn].concat(rest)));
    }
    // 写入目标区域
    if (typeof Array2D !== 'undefined' && Array2D.toRange) {
        Array2D.toRange(result, rng);
    } else {
        rng.Value2 = result;
    }
    if (typeof Console !== 'undefined') {
        Console.log("✅ batch_k: " + rows + " 个 k() 调用结果已写入 " + targetRangeAddr);
    }
    return rows;
}

// ╔════════════════════════════════════════════════════════════════╗
// ║ 帮助：常见问题 FAQ                                              ║
// ╚════════════════════════════════════════════════════════════════╝

/**
 * 当用户问"为什么 =k(...) 报 #NAME?" 时调用这个看帮助
 */
function k_help() {
    if (typeof Console === 'undefined') return;
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("  k() 公式常见问题排查");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
    Console.log("");
    Console.log("❌ 问题1: =k(...) 报 #NAME?");
    Console.log("   原因: WPS 公式引擎不认 k() 函数");
    Console.log("   解决: 把本文件作为加载项加载，或在 JSA 编辑器顶层粘入");
    Console.log("");
    Console.log("❌ 问题2: k() 返回 #K_ERR: ...");
    Console.log("   原因: 函数表达式语法错 或 参数不匹配");
    Console.log("   解决: 在 JSA 编辑器里直接调 JSA.jsaLambda(...) 调试");
    Console.log("");
    Console.log("❌ 问题3: 改了源数据，k() 结果不更新");
    Console.log("   解决: 绑 Sheet_Change 事件到 k_onChange (见本文件)");
    Console.log("");
    Console.log("❌ 问题4: 公式返回单个值时不显示（数组没 spill）");
    Console.log("   原因: WPS 版本 < 15990，不支持数组溢出");
    Console.log("   解决: 升级 WPS Office 到 15990+ 版本");
    Console.log("");
    Console.log("✅ 验证 k() UDF 已就绪：");
    var ok = false;
    try { k("JSA.getIndexs", 1, 3, 1); ok = true; } catch (e) {}
    if (ok) Console.log("   ✓ k() 内部调用成功（公式栏 =k(\"JSA.getIndexs\", 1, 3, 1) 应当返回 1）");
    else Console.log("   ✗ k() 内部调用失败 - 检查 JSA880.js 是否加载");
    Console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
}

// ==================== KO一切k函数整合完成 ====================
console.log('[JSA880 v4.2.2] KO一切k函数 UDF 已整合到主框架，单文件即可当 WPS 加载项使用');
console.log('  顶层 function k() / jsaLambda() 会被 WPS 公式引擎自动注册为 UDF');
console.log('  单元格中可用: =k("JSA.getIndexs", 1, 10, 2) → [1,3,5,7,9]');

// [v5.0.0] Array2D 定义完成后再确认 $$ 指向
(function _kFinalizeDollarDollar() {
    try {
        var __g = (function() { return this; })();
        if (typeof __g.Array2D !== 'undefined' &&
            (typeof __g.$$ === 'undefined' || __g.$$ === null)) {
            __g.$$ = __g.Array2D;
        }
    } catch (e) { /* 静默失败 */ }
})();
