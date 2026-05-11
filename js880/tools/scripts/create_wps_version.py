#!/usr/bin/env python3
"""
创建 JSA880 WPS专用版
移除所有 Node.js 和浏览器兼容代码
"""

import re

input_file = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js'
output_file = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/dist/JSA880_WPS.js'

with open(input_file, 'r', encoding='utf-8') as f:
    content = f.read()

original_lines = len(content.split('\n'))

# 1. 替换环境检测变量
content = re.sub(
    r'const isWPS = typeof Application !== [^;]+;',
    'const isWPS = true;  // WPS专用版',
    content
)
content = re.sub(
    r'const isNodeJS = [^;]+;',
    'const isNodeJS = false;  // WPS专用版: 禁用Node.js',
    content
)
content = re.sub(
    r'const isBrowser = [^;]+;',
    'const isBrowser = false;  // WPS专用版: 禁用浏览器',
    content
)

# 2. 移除文件末尾的 module.exports 块 (约12302行开始)
lines = content.split('\n')
new_lines = []
skip_module_exports = False

for i, line in enumerate(lines):
    # 检测 module.exports 块开始
    if 'if (isNodeJS) {' in line and i > 12000:
        skip_module_exports = True
        new_lines.append('// WPS专用版: 移除Node.js导出代码')
        continue
    
    if skip_module_exports:
        if line.strip() == '}':
            skip_module_exports = False
        continue
    
    new_lines.append(line)

lines = new_lines

# 3. 简化 IO 模块中的环境判断
# 将 if (!isWPS && !isNodeJS) 改为 if (!isWPS)
new_lines = []
for line in lines:
    line = re.sub(r'if \(!isWPS && !isNodeJS\)', 'if (!isWPS)', line)
    line = re.sub(r'if \(!isWPS && !false\)', 'if (!isWPS)', line)
    new_lines.append(line)

lines = new_lines

# 4. 移除 Node.js 特定注释
new_lines = []
for line in lines:
    if '// Node.js 环境：' in line:
        continue
    new_lines.append(line)

content = '\n'.join(new_lines)

# 5. 移除 require('fs') 相关代码
content = re.sub(r"var fs = require\('fs'\);", '// require已禁用', content)
content = re.sub(r'fs\.existsSync\([^)]+\)', 'false', content)
content = re.sub(r'fs\.statSync\([^)]+\)', '{isFile: function() {return false;}, isDirectory: function() {return false;}}', content)

# 6. 添加文件头
header = '''/**
 * JSA880.js - WPS专用版
 * 
 * 版本: v3.9.1
 * 说明: 此版本仅用于WPS Office JavaScript API环境
 *       移除了Node.js和浏览器兼容代码，精简体积
 * 
 * 使用方式:
 *   1. 在WPS中按 Alt+F11 打开宏编辑器
 *   2. 导入此文件或复制粘贴代码
 *   3. 直接调用 Array2D.z超级透视() 等功能
 */

'''

content = header + content

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(content)

new_lines = len(content.split('\n'))
print(f'✓ WPS专用版创建完成!')
print(f'  原始行数: {original_lines}')
print(f'  清理后行数: {new_lines}')
print(f'  减少: {original_lines - new_lines} 行 ({((original_lines - new_lines) / original_lines * 100):.1f}%)')
print(f'  输出文件: {output_file}')

# 验证
with open(output_file, 'r') as f:
    check = f.read()
    if 'module.exports' in check:
        print('  ⚠ 警告: 仍包含 module.exports')
    if "require('fs')" in check:
        print('  ⚠ 警告: 仍包含 require')
    else:
        print('  ✓ 已清理所有Node.js代码')
