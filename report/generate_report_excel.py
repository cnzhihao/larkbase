#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析报告Excel生成器
读取文件夹下所有分析报告markdown文件，生成包含公司名称和完整markdown内容的Excel表格
"""

import os
import re
import pandas as pd
from pathlib import Path
import argparse
from datetime import datetime

def extract_company_name(filename):
    """
    从文件名中提取公司名称
    文件名格式：公司名称-分析报告.md
    """
    # 去掉"-分析报告.md"后缀，获取公司名称
    company_name = re.sub(r'-分析报告\.md$', '', filename)
    return company_name

def read_markdown_file(file_path):
    """
    读取markdown文件的完整内容
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return content
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {e}")
        return None

def scan_reports(directory):
    """
    扫描指定目录下的所有分析报告文件
    """
    reports = []

    # 确保目录存在
    if not os.path.exists(directory):
        print(f"错误：目录 {directory} 不存在")
        return reports

    # 扫描目录下所有.md文件
    for filename in os.listdir(directory):
        if filename.endswith('-分析报告.md'):
            file_path = os.path.join(directory, filename)
            company_name = extract_company_name(filename)
            content = read_markdown_file(file_path)

            if content is not None:
                reports.append({
                    'company_name': company_name,
                    'filename': filename,
                    'content': content
                })
                print(f"已读取：{company_name}")
            else:
                print(f"跳过文件：{filename}（读取失败）")

    return reports

def generate_excel(reports, output_file):
    """
    生成Excel文件
    """
    if not reports:
        print("没有找到任何报告数据")
        return False

    # 准备数据
    data = []
    for report in reports:
        data.append({
            '公司名称': report['company_name'],
            '报告内容': report['content']
        })

    # 创建DataFrame
    df = pd.DataFrame(data)

    try:
        # 创建Excel writer对象
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 写入数据
            df.to_excel(writer, sheet_name='分析报告汇总', index=False)

            # 获取工作表对象
            worksheet = writer.sheets['分析报告汇总']

            # 调整列宽
            worksheet.column_dimensions['A'].width = 30  # 公司名称列宽
            worksheet.column_dimensions['B'].width = 100  # 报告内容列宽

            # 设置文本格式，确保长文本正确显示
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.value:
                        cell.alignment = cell.alignment.copy(wrap_text=True)

        print(f"Excel文件已生成：{output_file}")
        print(f"共处理 {len(reports)} 份报告")
        return True

    except Exception as e:
        print(f"生成Excel文件时出错: {e}")
        return False

def main():
    """
    主函数
    """
    parser = argparse.ArgumentParser(description='分析报告Excel生成器')
    parser.add_argument('--input-dir', '-i', default='.',
                       help='分析报告所在目录（默认为当前目录）')
    parser.add_argument('--output-file', '-o',
                       default=f'分析报告汇总_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                       help='输出Excel文件名')

    args = parser.parse_args()

    print("=== 分析报告Excel生成器 ===")
    print(f"扫描目录：{os.path.abspath(args.input_dir)}")
    print(f"输出文件：{os.path.abspath(args.output_file)}")
    print()

    # 扫描报告文件
    print("开始扫描分析报告...")
    reports = scan_reports(args.input_dir)

    if not reports:
        print("未找到任何分析报告文件")
        return

    print(f"找到 {len(reports)} 份报告")
    print()

    # 生成Excel文件
    print("开始生成Excel文件...")
    success = generate_excel(reports, args.output_file)

    if success:
        print()
        print("✅ 处理完成！")
        print(f"📊 输出文件：{os.path.abspath(args.output_file)}")
    else:
        print()
        print("❌ 处理失败！")

if __name__ == "__main__":
    main()