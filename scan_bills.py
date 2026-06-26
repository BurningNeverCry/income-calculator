#!/usr/bin/env python3
"""扫描文件夹中的富途年度账单文件，识别各年份。"""
import sys
import os
import glob
import re

def scan_bills(folder_path):
    """扫描文件夹，返回 {year: filename} 的字典。"""
    if not os.path.isdir(folder_path):
        print(f"错误：文件夹不存在: {folder_path}", file=sys.stderr)
        sys.exit(1)

    xlsx_files = glob.glob(os.path.join(folder_path, '*.xlsx'))
    if not xlsx_files:
        print(f"错误：文件夹中没有找到 xlsx 文件: {folder_path}", file=sys.stderr)
        sys.exit(1)

    year_files = {}
    for fpath in sorted(xlsx_files):
        fname = os.path.basename(fpath)
        # 匹配年份：文件名中的4位数字（2020-2030范围）
        matches = re.findall(r'(20[2-3]\d)', fname)
        for year_str in matches:
            year = int(year_str)
            if year not in year_files:
                year_files[year] = fpath

    return year_files

def main():
    if len(sys.argv) < 2:
        print("用法: python3 scan_bills.py <folder_path>", file=sys.stderr)
        sys.exit(1)

    folder_path = sys.argv[1]
    year_files = scan_bills(folder_path)

    if not year_files:
        print("未找到任何年度账单文件。", file=sys.stderr)
        sys.exit(1)

    print(f"共找到 {len(year_files)} 个年度账单文件：")
    print()
    for year in sorted(year_files.keys()):
        print(f"  {year} 年: {os.path.basename(year_files[year])}")

    print()
    print(f"年份范围: {min(year_files.keys())} - {max(year_files.keys())}")

if __name__ == '__main__':
    main()
