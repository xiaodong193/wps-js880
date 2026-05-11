/**
 * JSA880 WPS专用版清理脚本
 * 移除 Node.js 和浏览器兼容代码
 */

const fs = require('fs');
const path = require('path');

const inputFile = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js';
const outputFile = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/dist/JSA880_WPS专用版.js';

console.log('开始清理 JSA880.js...');

let content = fs.readFileSync(inputFile, 'utf-8');
const originalLines = content.split('\n').length;

// 1. 移除环境检测变量定义
content = content.replace(
    /\/\/ ==================== \[ENV_DETECTION\] 环境检测 ====================[\s\S]*?\/\/ ==================== \[LAMBDA_PARSER\]/,
    '// ==================== [WPS_ENV] WPS专用环境 ====================\nconst isWPS = true;\n\n// ==================== [LAMBDA_PARSER]'
);

// 2. 移除 isNodeJS 和 isBrowser 的使用
content = content.replace(/const isNodeJS =[^;]+;/g, '// 已移除: isNodeJS检测');
content = content.replace(/const isBrowser =[^;]+;/g, '// 已移除: isBrowser检测');

// 3. 简化 IO 模块中的环境判断
// 替换: if (!isWPS && !isNodeJS) return false;
content = content.replace(/if \(!isWPS && !isNodeJS\) return false;/g, '// WPS专用: 跳过Node.js检测');

// 4. 移除 Node.js 分支代码
// 替换: if (isNodeJS) { ... } 块
content = content.replace(
    /if \(isNodeJS\) \{\s*try \{\s*var fs = require\('fs'\);[\s\S]*?\}\s*return false;\s*\}/g,
    '// WPS专用: 移除Node.js文件操作代码'
);

// 5. 移除文件末尾的 module.exports
content = content.replace(
    /if \(isNodeJS\) \{[\s\S]*?module\.exports\.[^}]+\}/g,
    '// WPS专用: 移除Node.js导出代码'
);

// 6. 清理空行
content = content.replace(/\n{3,}/g, '\n\n');

// 写入文件
fs.writeFileSync(outputFile, content, 'utf-8');

const newLines = content.split('\n').length;

console.log('✓ 清理完成!');
console.log(`  原始行数: ${originalLines}`);
console.log(`  清理后行数: ${newLines}`);
console.log(`  减少: ${originalLines - newLines} 行 (${((originalLines - newLines) / originalLines * 100).toFixed(1)}%)`);
console.log(`  输出文件: ${outputFile}`);
