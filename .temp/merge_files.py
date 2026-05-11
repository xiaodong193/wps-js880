#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合并 JSA880.js 文件的两个版本
保留新版本的 superPivot 功能，使用老版本的 ES6 语法
"""

def read_file_lines(file_path):
    """读取文件的所有行"""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.readlines()

def write_file_lines(file_path, lines):
    """写入行到文件"""
    with open(file_path, 'w', encoding='utf-8') as f:
        f.writelines(lines)

def merge_files():
    old_file = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880 copy.js'
    new_file = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js'
    output_file = '/Users/daidai193/Library/CloudStorage/SynologyDrive-WWJZ/temp/js880/JSA880.js'

    old_lines = read_file_lines(old_file)
    new_lines = read_file_lines(new_file)

    # 老版本需要替换的代码行范围（0-based，所以减1）
    # Array2D.z超级透视: 5877-7198行 -> 5876-7197 (0-based)
    old_superpivot_start = 5876  # 行5877
    old_superpivot_end = 7197    # 行7198

    # 新版本需要提取的代码行范围
    # Array2D.z超级透视: 5908-7227行 -> 5907-7226 (0-based)
    new_superpivot_start = 5907  # 行5908
    new_superpivot_end = 7226    # 行7227

    # 提取新版本的 superPivot 代码
    new_superpivot_code = new_lines[new_superpivot_start:new_superpivot_end + 1]

    # 构建合并后的文件
    merged_lines = []

    # 1. 添加老版本 1-5876行
    merged_lines.extend(old_lines[:old_superpivot_start])

    # 2. 添加新版本的 superPivot 代码
    merged_lines.extend(new_superpivot_code)

    # 3. 添加老版本 7198行之后的所有内容
    merged_lines.extend(old_lines[old_superpivot_end + 1:])

    # 写入合并后的文件
    write_file_lines(output_file, merged_lines)

    print(f"合并完成！")
    print(f"- 读取老版本: {len(old_lines)} 行")
    print(f"- 读取新版本: {len(new_lines)} 行")
    print(f"- 合并后文件: {len(merged_lines)} 行")
    print(f"- 替换了老版本第 {old_superpivot_start + 1}-{old_superpivot_end + 1} 行")
    print(f"- 使用了新版本第 {new_superpivot_start + 1}-{new_superpivot_end + 1} 行")

if __name__ == '__main__':
    merge_files()
