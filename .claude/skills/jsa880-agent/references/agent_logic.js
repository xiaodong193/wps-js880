/**
 * @file agent_logic.js
 * @description JSA880智能助手代码生成引擎参考实现
 * @version 1.0.0
 * @date 2026-05-14
 */

/**
 * 意图分类器
 * 将用户自然语言描述分类为具体意图
 */
var clsIntentClassifier = {
    /**
     * 意图模式定义
     */
    _patterns: {
        'array_distinct': {
            keywords: ['去重', '唯一', '不重复', '去重复', '去除重复'],
            description: '数组去重'
        },
        'array_filter': {
            keywords: ['筛选', '过滤', '只保留', '找出', '查找'],
            description: '数据筛选'
        },
        'array_sort': {
            keywords: ['排序', '升序', '降序', '由小到大', '由大到小'],
            description: '数据排序'
        },
        'array_group': {
            keywords: ['分组', '汇总', '统计', '分类', '小计'],
            description: '分组汇总'
        },
        'array_pivot': {
            keywords: ['透视', '交叉', '行转列', '列转行', '多维'],
            description: '数据透视'
        },
        'array_join': {
            keywords: ['连接', '关联', '匹配', '合并', '左连接'],
            description: '表连接'
        },
        'file_traverse': {
            keywords: ['遍历', '遍历文件', '列出', '查找文件'],
            description: '文件遍历'
        },
        'date_calc': {
            keywords: ['日期', '天数', '间隔', '计算日期'],
            description: '日期计算'
        },
        'cell_write': {
            keywords: ['写入', '输出', '填入', '显示'],
            description: '写入单元格'
        },
        'cell_format': {
            keywords: ['格式', '样式', '设置', '颜色', '字体'],
            description: '单元格格式'
        }
    },

    /**
     * 分类用户输入
     * @param {string} input - 用户输入
     * @returns {Object} 意图结果
     */
    classify: function(input) {
        var text = input.toLowerCase().replace(/[，。！？；：、]/g, ',');
        var scores = {};

        // 遍历所有模式，计算匹配分数
        for (var intentType in this._patterns) {
            var pattern = this._patterns[intentType];
            var score = 0;

            // 关键词匹配
            for (var i = 0; i < pattern.keywords.length; i++) {
                if (text.indexOf(pattern.keywords[i]) >= 0) {
                    score += 2;
                }
            }

            if (score > 0) {
                scores[intentType] = {
                    type: intentType,
                    score: score,
                    description: pattern.description
                };
            }
        }

        // 选择得分最高的意图
        if (Object.keys(scores).length === 0) {
            return { type: 'unknown', score: 0, description: '未知意图' };
        }

        var best = null;
        for (var type in scores) {
            if (!best || scores[type].score > best.score) {
                best = scores[type];
            }
        }

        return {
            type: best.type,
            score: best.score,
            description: best.description,
            confidence: Math.min(best.score / 10, 1)
        };
    }
};

/**
 * 实体提取器
 * 从用户输入中提取关键实体（范围、字段、条件等）
 */
var clsEntityExtractor = {
    /**
     * 提取数据范围
     * @param {string} input - 用户输入
     * @returns {string} 范围描述
     */
    extractRange: function(input) {
        // 匹配 A1:D100 等范围格式
        var rangeMatch = input.match(/[A-Z]+\d*:[A-Z]+\d*/i);
        if (rangeMatch) {
            return rangeMatch[0];
        }

        // 匹配 "A列" 或 "第1列" 等描述
        var colMatch = input.match(/第?\s*([A-Z]|\d+)\s*[列行]/i);
        if (colMatch) {
            return colMatch[0];
        }

        return 'A1';  // 默认
    },

    /**
     * 提取字段选择器
     * @param {string} input - 用户输入
     * @returns {string} 字段选择器
     */
    extractField: function(input) {
        // 匹配 "按第1列" "按A列" "按姓名"
        var fieldMatch = input.match(/按\s*(第?\s*\d+|第?\s*[A-Z])\s*列?/i);
        if (fieldMatch) {
            var field = fieldMatch[1];
            // 如果是数字，转换为 f1 格式
            if (/^\d+$/.test(field)) {
                return 'f' + field;
            }
            // 如果是字母，转换为列号
            if (/^[A-Z]$/i.test(field)) {
                var colNum = field.toUpperCase().charCodeAt(0) - 64;
                return 'f' + colNum;
            }
            return field;
        }

        // 匹配 "第1列" "第一列"
        var ordinalMatch = input.match(/第[一二三四五六七八九十百千万\d]+列/i);
        if (ordinalMatch) {
            var num = this._chineseToNumber(ordinalMatch[0]);
            if (num > 0) {
                return 'f' + num;
            }
        }

        return 'f1';  // 默认第1列
    },

    /**
     * 提取条件表达式
     * @param {string} input - 用户输入
     * @returns {string} Lambda条件表达式
     */
    extractCondition: function(input) {
        // 提取比较条件
        var gtMatch = input.match(/(大于|超过|多于|>|>)\s*(\d+)/i);
        if (gtMatch) {
            return 'f1 > ' + gtMatch[2];
        }

        var ltMatch = input.match(/(小于|少于|不足|<|<)\s*(\d+)/i);
        if (ltMatch) {
            return 'f1 < ' + ltMatch[2];
        }

        var eqMatch = input.match(/(等于|为|是|=)\s*["']?([^"'\s,]+)["']?/i);
        if (eqMatch) {
            return 'f1 === "' + eqMatch[2] + '"';
        }

        return '';
    },

    /**
     * 中文数字转阿拉伯数字
     * @param {string} chinese - 中文数字
     * @returns {number}
     */
    _chineseToNumber: function(chinese) {
        var map = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
            '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
        };

        var num = 0;
        for (var i = 0; i < chinese.length; i++) {
            var char = chinese[i];
            if (map[char]) {
                num = num * 10 + map[char];
            }
        }

        return num > 0 ? num : 1;
    }
};

/**
 * 代码模板管理器
 */
var clsCodeTemplateManager = {
    /**
     * 获取数组处理模板
     * @returns {string} 代码模板
     */
    getArrayProcessingTemplate: function() {
        return [
            '/**',
            ' * @description {description}',
            ' * @date {date}',
            ' */',
            'function {functionName}() {',
            '    // 1. 读取数据到二维数组',
            '    var arr = $.maxArray("{dataRange}");',
            '',
            '    if (!arr || arr.length === 0) {',
            '        console.log("数据为空");',
            '        return;',
            '    }',
            '',
            '    // 2. 数据处理',
            '    var result = {processCode};',
            '',
            '    // 3. 输出结果',
            '    if (result && result.toRange) {',
            '        result.toRange("{outputRange}", true);',
            '    } else {',
            '        new Array2D(result).toRange("{outputRange}", true);',
            '    }',
            '',
            '    console.log("处理完成，共" + result.length + "行");',
            '}'
        ].join('\n');
    },

    /**
     * 获取文件操作模板
     * @returns {string} 代码模板
     */
    getFileOperationTemplate: function() {
        return [
            '/**',
            ' * @description {description}',
            ' * @date {date}',
            ' */',
            'function {functionName}() {',
            '    // 1. 选择文件夹',
            '    var folderPath = IO.showFolderDialog();',
            '    if (!folderPath) {',
            '        MsgBox("未选择文件夹");',
            '        return;',
            '    }',
            '',
            '    // 2. 遍历文件',
            '    var files = IO.getFiles(folderPath, true, false);',
            '    var result = [];',
            '',
            '    files.forEach(function(file) {',
            '        // 文件处理逻辑',
            '        {fileProcessCode}',
            '    });',
            '',
            '    // 3. 输出结果',
            '    result.toRange("A1", true);',
            '    MsgBox("处理完成，共" + files.length + "个文件");',
            '}'
        ].join('\n');
    },

    /**
     * 获取透视表模板
     * @returns {string} 代码模板
     */
    getPivotTemplate: function() {
        return [
            '/**',
            ' * @description {description}',
            ' * @date {date}',
            ' */',
            'function {functionName}() {',
            '    // 1. 读取数据',
            '    var data = $.maxArray("A1");',
            '',
            '    if (!data || data.length === 0) {',
            '        console.log("数据为空");',
            '        return;',
            '    }',
            '',
            '    // 2. 超级透视',
            '    var pivot = Array2D.z超级透视(',
            '        data,',
            '        {rowField},     // 行字段',
            '        {colField},      // 列字段',
            '        {dataField}     // 数据字段',
            '    );',
            '',
            '    // 3. 输出结果',
            '    pivot.toRange("{outputRange}", true);',
            '    console.log("透视完成");',
            '}'
        ].join('\n');
    }
};

/**
 * 代码生成器
 * 根据意图和实体生成JSA代码
 */
var clsCodeGenerator = {
    _templates: new clsCodeTemplateManager(),

    /**
     * 生成代码
     * @param {Object} intent - 意图对象
     * @param {Object} entities - 实体对象
     * @returns {string} 生成的代码
     */
    generate: function(intent, entities) {
        var template = '';
        var processCode = '';

        switch (intent.type) {
            case 'array_distinct':
                processCode = 'Array2D(data).z去重("' + entities.field + '")';
                template = this._templates.getArrayProcessingTemplate();
                break;

            case 'array_filter':
                var condition = entities.condition || entities.field + ' > 0';
                processCode = 'Array2D(data).z筛选(\'' + condition + '\')';
                template = this._templates.getArrayProcessingTemplate();
                break;

            case 'array_sort':
                var sortExpr = entities.field + '+';
                if (intent.description.indexOf('降序') >= 0) {
                    sortExpr = entities.field + '-';
                }
                processCode = 'Array2D(data).z多列排序(\'' + sortExpr + '\')';
                template = this._templates.getArrayProcessingTemplate();
                break;

            case 'array_group':
                var aggExpr = 'sum("' + entities.field + '")';
                if (intent.description.indexOf('统计') >= 0) {
                    aggExpr = 'count()';
                }
                processCode = 'Array2D.groupInto(data, "' + entities.field + '", "' + aggExpr + '")';
                template = this._templates.getArrayProcessingTemplate();
                break;

            case 'array_pivot':
                template = this._templates.getPivotTemplate();
                break;

            case 'file_traverse':
                template = this._templates.getFileOperationTemplate();
                break;

            default:
                // 默认生成通用处理模板
                processCode = 'Array2D(data).z结果()';
                template = this._templates.getArrayProcessingTemplate();
        }

        // 填充模板
        var date = new Date().toISOString().split('T')[0];
        var functionName = this._generateFunctionName(intent.description);

        var code = template
            .replace('{description}', intent.description)
            .replace('{date}', date)
            .replace('{functionName}', functionName)
            .replace('{dataRange}', entities.range || 'A1')
            .replace('{processCode}', processCode)
            .replace('{outputRange}', 'H1')
            .replace('{rowField}', "['f1', '行标题']")
            .replace('{colField}', "['f2', '列标题']")
            .replace('{dataField}', "['sum(\"f3\")', '数据']")
            .replace('{fileProcessCode}', 'result.push([file]);');

        return code;
    },

    /**
     * 生成函数名
     * @param {string} description - 功能描述
     * @returns {string} 函数名
     */
    _generateFunctionName: function(description) {
        var keywords = description
            .replace(/[^\w\u4e00-\u9fa5]/g, ' ')
            .split(/\s+/)
            .filter(function(word) { return word.length > 1; })
            .slice(0, 3);

        var functionName = keywords.map(function(word, index) {
            if (index === 0) {
                return word.toLowerCase();
            }
            return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
        }).join('');

        return functionName || 'processData';
    }
};

/**
 * 主Agent类
 * 整合所有组件，提供统一的接口
 */
var clsJSA880Agent = {
    _classifier: clsIntentClassifier,
    _extractor: clsEntityExtractor,
    _generator: clsCodeGenerator,

    /**
     * 处理用户输入，生成代码
     * @param {string} userInput - 用户自然语言输入
     * @returns {Object} 处理结果
     */
    process: function(userInput) {
        // 1. 意图分类
        var intent = this._classifier.classify(userInput);

        // 2. 实体提取
        var entities = {
            range: this._extractor.extractRange(userInput),
            field: this._extractor.extractField(userInput),
            condition: this._extractor.extractCondition(userInput)
        };

        // 3. 生成代码
        var code = this._generator.generate(intent, entities);

        return {
            success: true,
            intent: intent,
            entities: entities,
            code: code
        };
    }
};

// 导出（兼容WPS JSA环境）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        clsIntentClassifier: clsIntentClassifier,
        clsEntityExtractor: clsEntityExtractor,
        clsCodeTemplateManager: clsCodeTemplateManager,
        clsCodeGenerator: clsCodeGenerator,
        clsJSA880Agent: clsJSA880Agent
    };
}