/**
 * JSA880 API 提取脚本
 *
 * 功能：从 JSA880.js 源代码中提取 API 信息，生成结构化 JSON
 *
 * 使用方法：node scripts/extract-api.js
 *
 * 输出：
 *   - dist/api-data.json (完整的 API 数据)
 *   - dist/jsa-functions.json (JSA 全局函数)
 *   - dist/array2d-methods.json (Array2D 方法)
 */

import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// 配置 - JSA880.js 在父目录的 js880 文件夹中
const SOURCE_FILE = path.resolve(__dirname, '../../js880/JSA880.js')
const OUTPUT_DIR = path.resolve(__dirname, '../docs/.vitepress/dist')
const DIST_DIR = path.resolve(__dirname, '../dist')

// 确保输出目录存在
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true })
}
if (!fs.existsSync(DIST_DIR)) {
    fs.mkdirSync(DIST_DIR, { recursive: true })
}

// 读取源代码
console.log('📖 读取源代码:', SOURCE_FILE)
const sourceCode = fs.readFileSync(SOURCE_FILE, 'utf-8')
console.log(`   文件大小: ${(sourceCode.length / 1024).toFixed(1)} KB`)
console.log(`   行数: ${sourceCode.split('\n').length}`)

// 提取 JSA 全局函数
function extractJSAFunctions(code) {
    const functions = []
    // 匹配 JSA.xxx = function 或 JSA.xxx = JSA.yyy 模式
    const regex = /JSA\.([z]?\w+)\s*=\s*(function\s*\([^)]*\)|JSA\.[z]?\w+)/g
    let match

    while ((match = regex.exec(code)) !== null) {
        const name = match[1]
        const isAlias = match[2].startsWith('JSA.')
        const aliasOf = isAlias ? match[2].replace('JSA.', '') : null

        // 跳过内部属性
        if (['prototype', '__proto__', 'constructor'].includes(name)) continue

        // 获取函数的 JSDoc 注释
        const jsdoc = extractJSDoc(code, match.index)

        // 提取函数签名
        const signature = isAlias ? null : extractFunctionSignature(code, match.index)

        functions.push({
            name,
            aliasOf,
            isAlias,
            signature,
            description: jsdoc.description,
            params: jsdoc.params,
            returns: jsdoc.returns,
            examples: jsdoc.examples,
            category: categorizeJSAFunction(name)
        })
    }

    return functions
}

// 提取 Array2D.prototype 方法
function extractArray2DMethods(code) {
    const methods = []
    // 匹配 Array2D.prototype.zxxx = function 或 Array2D.prototype.xxx = Array2D.prototype.zxxx 模式
    const regex = /Array2D\.prototype\.([z]?\w+)\s*=\s*(function\s*\([^)]*\)|Array2D\.prototype\.[z]?\w+)/g
    let match

    while ((match = regex.exec(code)) !== null) {
        const name = match[1]
        const isAlias = match[2].includes('Array2D.prototype.')
        const aliasOf = isAlias ? match[2].replace('Array2D.prototype.', '') : null

        // 跳过内部属性
        if (['prototype', '__proto__', 'constructor'].includes(name)) continue

        // 获取函数的 JSDoc 注释
        const jsdoc = extractJSDoc(code, match.index)

        // 提取函数签名
        const signature = isAlias ? null : extractFunctionSignature(code, match.index)

        methods.push({
            name,
            aliasOf,
            isAlias,
            signature,
            description: jsdoc.description,
            params: jsdoc.params,
            returns: jsdoc.returns,
            examples: jsdoc.examples,
            category: categorizeArray2DMethod(name)
        })
    }

    return methods
}

// 提取函数签名
function extractFunctionSignature(code, startIndex) {
    // 向前查找 function 关键字
    const beforeCode = code.substring(0, startIndex)
    const lastFuncIndex = beforeCode.lastIndexOf('function')
    if (lastFuncIndex === -1) return null

    // 提取 function 到匹配位置之间的内容
    const funcBlock = code.substring(lastFuncIndex, startIndex + 100)
    const match = funcBlock.match(/function\s*\(([^)]*)\)/)
    if (match) {
        return {
            params: match[1].split(',').map(p => p.trim()).filter(p => p)
        }
    }
    return null
}

// 提取 JSDoc 注释
function extractJSDoc(code, startIndex) {
    const beforeCode = code.substring(0, startIndex)
    const lines = beforeCode.split('\n')

    const result = {
        description: '',
        params: [],
        returns: null,
        examples: []
    }

    // 查找最近的多行注释
    for (let i = lines.length - 1; i >= Math.max(0, lines.length - 30); i--) {
        const line = lines[i]
        if (line.trim().startsWith('/**')) {
            // 找到 JSDoc 开始，提取到 */ 之间的内容
            const jsdocLines = lines.slice(i)
            const jsdocText = jsdocLines.join('\n')
            const endMatch = jsdocText.match(/\*\/([\s\S]*)/)

            if (endMatch) {
                const content = endMatch[1]
                // 解析 @param
                const paramMatches = content.matchAll(/@param\s+\{([^}]+)\}\s+(\w+)\s*[-–—]\s*([^\n]+)/g)
                for (const pm of paramMatches) {
                    result.params.push({
                        type: pm[1],
                        name: pm[2],
                        description: pm[3].trim()
                    })
                }

                // 解析 @returns
                const returnsMatch = content.match(/@returns?\s*\{([^}]+)\}/)
                if (returnsMatch) {
                    result.returns = { type: returnsMatch[1] }
                }

                // 解析 @example
                const exampleMatches = content.matchAll(/@example\s*\n\s*```(?:javascript)?\n([\s\S]*?)```/g)
                for (const em of exampleMatches) {
                    result.examples.push(em[1].trim())
                }

                // 提取描述（第一条 * 后面的内容）
                const descMatch = content.match(/^\s*\*\s*([^\n@](?:[^\n]*[^\n])?)/m)
                if (descMatch) {
                    result.description = descMatch[1].replace(/^\*/, '').trim()
                }
            }
            break
        } else if (!line.trim().startsWith('*') && !line.trim().startsWith('//') && line.trim() !== '') {
            break
        }
    }

    return result
}

// 函数分类
function categorizeJSAFunction(name) {
    if (['z转置', 'z选择列', 'z矩阵分布'].includes(name)) return '数组操作'
    if (['z求和', 'z最大值', 'z最小值', 'z平均值', 'z表达式求值', 'z生成数字序列', 'z取整数', 'z取小数'].includes(name)) return '数学计算'
    if (['z今天', 'z转日期数值', 'z日期间隔'].includes(name)) return '日期时间'
    if (['z写入单元格'].includes(name)) return 'WPS操作'
    if (['z查找索引', 'z左侧查找', 'z增强查找'].includes(name)) return '查找'
    if (['z转数值', 'z转文本', 'z替换', 'z截取字符', 'z左取字符', 'z右取字符', 'z模糊匹配', 'z转公式数组'].includes(name)) return '字符串处理'
    if (['z人民币大写', 'z随机整数', 'z随机打乱', 'z延时', 'z统一路径分隔符'].includes(name)) return '工具'
    if (['jsaLambda', 'z解析函数表达式'].includes(name)) return '函数式'
    return '其他'
}

function categorizeArray2DMethod(name) {
    if (['val', 'z是否为空', 'z数量', 'z克隆'].includes(name)) return '基础操作'
    if (['z输出HTML', 'z连接', 'z文本连接', 'z转JSON', 'z转字符串', 'z转矩阵', 'z矩阵排版', 'z矩阵运算'].includes(name)) return '输出转换'
    if (['z填充', 'z补齐空位', 'z补齐数组', 'z重设大小'].includes(name)) return '填充调整'
    if (['z扁平化', 'z反转', 'z转置', 'z矩阵信息'].includes(name)) return '矩阵操作'
    if (['z求和', 'z平均值', 'z最大值', 'z最小值', 'z中位数'].includes(name)) return '统计计算'
    if (['z第一个', 'z最后一个', 'z单元格', 'z设置单元格', 'toRange'].includes(name)) return '元素访问'
    if (['z分块', 'z挑选', 'z跳过', 'z取前N个', 'z行切片', 'z行切片删除行', 'z间隔取数', 'z重复N次'].includes(name)) return '数组切片'
    if (['z筛选', 'z跳过前面连续满足', 'z取前面连续满足', 'where'].includes(name)) return '条件筛选'
    if (['z遍历执行', 'z倒序遍历执行', 'z映射', 'z归约', 'z倒序归约', 'z全部满足', 'z有满足', 'z按范围遍历', 'z区域映射'].includes(name)) return '遍历迭代'
    if (['z行数', 'z列数', 'z获取行', 'z获取列', 'z首行', 'z末行', 'z首列', 'z末列', 'z插入行号'].includes(name)) return '行列信息'
    if (['z添加行', 'z提取列', 'z添加列', 'z删除行', 'z删除列', 'z尾部弹出一项', 'z追加一项', 'z删除第一个', 'z批量删除列', 'z批量删除行', 'z批量插入列', 'z批量插入行'].includes(name)) return '行列增删'
    if (['z按规则升序', 'z按规则降序', 'z降序排序', 'z升序排序', 'z行排序', 'z列排序', 'z多列排序', 'z自定义排序', 'z智能排序'].includes(name)) return '排序'
    if (['z去重'].includes(name)) return '去重'
    if (['z分组', 'z透视', 'z超级透视', 'z超级查找'].includes(name)) return '分组透视'
    if (['z上下连接', 'z左连接', 'z左右全连接', 'z一对多连接', 'z左右连接'].includes(name)) return '连接操作'
    if (['z排除', 'z取交集', 'z去重并集', 'z按范围选择'].includes(name)) return '集合操作'
    if (['z查找单个', 'z查找所有下标', 'z查找所有行下标', 'z查找所有列下标', 'z查找元素下标', 'z值位置', 'z从后往前值位置'].includes(name)) return '查找'
    if (['z按页数分页', 'z按行数分页', 'z按下标分页'].includes(name)) return '分页'
    if (['z处理空值', 'z选择列', 'z选择行', 'z结果', 'z版本', 'z错误值', 'z空结果', 'z是否包含值', 'z随机打乱', 'z随机一项', 'z随机整数', 'z随机整数数组', 'z随机小数', 'z随机小数数组', 'z随机乱序数字序列', 'z聚合', 'z分组排名', 'z分组汇总', 'z分组汇总到字典', 'z区域矩阵'].includes(name)) return '工具'
    return '其他'
}

// 主执行
console.log('\n🚀 开始提取 API...\n')

try {
    // 提取 JSA 全局函数
    console.log('📦 提取 JSA 全局函数...')
    const jsaFunctions = extractJSAFunctions(sourceCode)
    const jsaFunctionsNoAlias = jsaFunctions.filter(f => !f.isAlias)
    console.log(`   找到 ${jsaFunctions.length} 个条目 (含 ${jsaFunctionsNoAlias.length} 个独立函数)`)

    // 提取 Array2D 方法
    console.log('📦 提取 Array2D 方法...')
    const array2dMethods = extractArray2DMethods(sourceCode)
    const array2dMethodsNoAlias = array2dMethods.filter(f => !f.isAlias)
    console.log(`   找到 ${array2dMethods.length} 个条目 (含 ${array2dMethodsNoAlias.length} 个独立方法)`)

    // 按分类分组
    const jsaByCategory = {}
    jsaFunctionsNoAlias.forEach(f => {
        if (!jsaByCategory[f.category]) jsaByCategory[f.category] = []
        jsaByCategory[f.category].push(f)
    })

    const array2dByCategory = {}
    array2dMethodsNoAlias.forEach(f => {
        if (!array2dByCategory[f.category]) array2dByCategory[f.category] = []
        array2dByCategory[f.category].push(f)
    })

    // 输出 JSON 文件
    const apiData = {
        version: '3.9.3',
        generatedAt: new Date().toISOString(),
        jsa: {
            functions: jsaFunctionsNoAlias,
            byCategory: jsaByCategory,
            total: jsaFunctionsNoAlias.length
        },
        array2d: {
            methods: array2dMethodsNoAlias,
            byCategory: array2dByCategory,
            total: array2dMethodsNoAlias.length
        }
    }

    // 保存完整数据
    fs.writeFileSync(path.join(DIST_DIR, 'api-data.json'), JSON.stringify(apiData, null, 2))
    console.log(`\n✅ 已保存: ${path.join(DIST_DIR, 'api-data.json')}`)

    // 保存简化版本
    fs.writeFileSync(path.join(DIST_DIR, 'jsa-functions.json'), JSON.stringify(jsaFunctionsNoAlias, null, 2))
    fs.writeFileSync(path.join(DIST_DIR, 'array2d-methods.json'), JSON.stringify(array2dMethodsNoAlias, null, 2))
    console.log(`✅ 已保存: ${path.join(DIST_DIR, 'jsa-functions.json')}`)
    console.log(`✅ 已保存: ${path.join(DIST_DIR, 'array2d-methods.json')}`)

    // 打印摘要
    console.log('\n📊 JSA 全局函数分类:')
    Object.keys(jsaByCategory).sort().forEach(cat => {
        console.log(`   ${cat}: ${jsaByCategory[cat].length} 个`)
    })

    console.log('\n📊 Array2D 方法分类:')
    Object.keys(array2dByCategory).sort().forEach(cat => {
        console.log(`   ${cat}: ${array2dByCategory[cat].length} 个`)
    })

    console.log('\n✨ 提取完成!\n')

} catch (error) {
    console.error('❌ 提取失败:', error.message)
    process.exit(1)
}