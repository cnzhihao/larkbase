#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书客户案例Excel生成脚本
从report文件夹的md文件中提取公司名称、分析报告内容和客户链接，生成Excel文件
"""

import os
import re
import glob
import pandas as pd
from pathlib import Path


def extract_company_name(filename):
    """
    从文件名中提取公司名称
    文件名格式：[公司名称]-分析报告.md
    """
    # 去掉"-分析报告.md"后缀
    if filename.endswith('-分析报告.md'):
        company_name = filename[:-7]  # 去掉"-分析报告.md"（7个字符）
        # 去掉可能存在的末尾"-"
        if company_name.endswith('-'):
            company_name = company_name[:-1]
        return company_name
    return filename


def extract_customer_links(content):
    """
    从文档内容中提取客户链接
    优先提取飞书客户案例链接
    """
    # 匹配HTTP/HTTPS链接的正则表达式
    url_pattern = r'https?://[^\s\*\)]+'

    # 查找所有链接
    all_links = re.findall(url_pattern, content)

    # 优先提取飞书客户案例链接 (feishu.cn/customers/)
    feishu_links = []
    other_links = []

    for link in all_links:
        if 'feishu.cn/customers' in link:
            feishu_links.append(link)
        else:
            other_links.append(link)

    # 返回飞书链接，如果没有则返回其他链接，都没有则返回空字符串
    if feishu_links:
        return feishu_links[0]  # 返回第一个飞书客户案例链接
    elif other_links:
        return other_links[0]  # 返回第一个其他链接
    else:
        return ""  # 没有找到链接


def read_file_content(file_path):
    """
    读取文件内容，处理编码问题
    """
    try:
        # 尝试用UTF-8编码读取
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            # 如果UTF-8失败，尝试GBK编码
            with open(file_path, 'r', encoding='gbk') as f:
                return f.read()
        except UnicodeDecodeError:
            # 如果都失败，尝试自动检测编码
            try:
                import chardet
                with open(file_path, 'rb') as f:
                    raw_data = f.read()
                    result = chardet.detect(raw_data)
                    encoding = result['encoding']
                    return raw_data.decode(encoding)
            except ImportError:
                # 如果没有chardet库，使用默认编码并忽略错误
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()


def process_reports(report_folder='report'):
    """
    处理report文件夹中的所有md文件
    """
    # 获取所有md文件
    md_files = glob.glob(os.path.join(report_folder, '*.md'))

    if not md_files:
        print(f"在 {report_folder} 文件夹中没有找到md文件")
        return []

    print(f"找到 {len(md_files)} 个md文件")

    results = []

    for i, file_path in enumerate(md_files, 1):
        print(f"处理文件 {i}/{len(md_files)}: {os.path.basename(file_path)}")

        # 获取文件名（不含路径）
        filename = os.path.basename(file_path)

        # 提取公司名称
        company_name = extract_company_name(filename)

        # 读取文件内容
        content = read_file_content(file_path)

        # 提取客户链接
        customer_link = extract_customer_links(content)

        # 添加到结果列表
        results.append({
            '公司名称': company_name,
            '分析报告': content,
            '客户链接': customer_link
        })

    return results


def generate_excel(results, output_file='飞书客户案例汇总表.xlsx'):
    """
    生成Excel文件
    """
    if not results:
        print("没有数据可以生成Excel文件")
        return

    # 创建DataFrame
    df = pd.DataFrame(results)

    # 创建Excel writer对象，使用openpyxl引擎
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='客户案例汇总', index=False)

        # 获取工作表对象
        worksheet = writer.sheets['客户案例汇总']

        # 调整列宽
        worksheet.column_dimensions['A'].width = 20  # 公司名称列
        worksheet.column_dimensions['B'].width = 100  # 分析报告列
        worksheet.column_dimensions['C'].width = 50   # 客户链接列

        # 设置行高自动调整
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value:
                    # 如果内容很长，设置自动换行
                    if len(str(cell.value)) > 50:
                        cell.alignment = cell.alignment.copy(wrap_text=True)

    print(f"Excel文件已生成: {output_file}")
    print(f"共处理 {len(results)} 个公司案例")


def main():
    """
    主函数
    """
    print("开始处理飞书客户案例报告...")

    # 检查report文件夹是否存在
    if not os.path.exists('report'):
        print("错误: report 文件夹不存在")
        return

    # 处理报告文件
    results = process_reports('report')

    if results:
        # 生成Excel文件
        generate_excel(results)

        # 显示统计信息
        companies_with_links = sum(1 for r in results if r['客户链接'])
        print(f"\n统计信息:")
        print(f"总公司数: {len(results)}")
        print(f"有客户链接的公司数: {companies_with_links}")
        print(f"无客户链接的公司数: {len(results) - companies_with_links}")

        # 显示前几个有链接的公司示例
        if companies_with_links > 0:
            print(f"\n有客户链接的公司示例:")
            for i, result in enumerate(results[:5]):
                if result['客户链接']:
                    print(f"  - {result['公司名称']}: {result['客户链接']}")
                    if i >= 4:  # 只显示前5个
                        break


if __name__ == "__main__":
    main()