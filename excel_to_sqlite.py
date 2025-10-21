#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel to SQLite Converter
将Excel文件转换为SQLite数据库
支持多工作表，自动清理列名，处理空值
"""

import pandas as pd
import sqlite3
import os
import sys
import re
from pathlib import Path
from datetime import datetime

def excel_to_sqlite(excel_file, sqlite_file=None):
    """
    将Excel文件转换为SQLite数据库
    
    Args:
        excel_file (str): Excel文件路径
        sqlite_file (str, optional): 输出SQLite文件路径，默认为None时自动生成
    
    Returns:
        str: 生成的SQLite文件路径
    """
    try:
        # 读取Excel文件
        print(f"正在读取Excel文件: {excel_file}")
        
        # 读取所有工作表
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        # 如果没有指定SQLite文件名，则自动生成
        if sqlite_file is None:
            excel_name = Path(excel_file).stem
            sqlite_file = f"output/{excel_name}.db"
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(sqlite_file), exist_ok=True)
        
        # 创建SQLite连接
        conn = sqlite3.connect(sqlite_file)
        
        print(f"正在创建SQLite数据库: {sqlite_file}")
        
        # 获取数据库文件名（不含扩展名）作为默认表名
        db_name = Path(sqlite_file).stem
        
        # 遍历所有工作表
        for sheet_name, df in excel_data.items():
            # 如果只有一个工作表，使用数据库文件名作为表名
            if len(excel_data) == 1:
                clean_sheet_name = db_name
                print(f"  单工作表模式: 使用数据库名作为表名 -> {clean_sheet_name}")
            else:
                # 多工作表模式：保持原始表名，只做必要的清理
                clean_sheet_name = sheet_name.strip()
                # 只替换SQLite不支持的字符，保留更多原始字符
                clean_sheet_name = re.sub(r'[^\w\s-]', '_', clean_sheet_name)
                clean_sheet_name = clean_sheet_name.replace(' ', '_')
                # 确保表名不为空且以字母或下划线开头
                if not clean_sheet_name or not clean_sheet_name[0].isalpha():
                    clean_sheet_name = f"table_{clean_sheet_name}" if clean_sheet_name else f"sheet_{list(excel_data.keys()).index(sheet_name) + 1}"
            
            print(f"  处理工作表: {sheet_name} -> {clean_sheet_name}")
            
            # 检查数据框是否为空
            if df.empty:
                print(f"    警告: 工作表 {sheet_name} 为空，跳过")
                continue
            
            # 保持原始列名，只做SQLite兼容性处理
            original_columns = df.columns.tolist()
            print(f"    原始列名: {original_columns[:5]}...")
            
            # 只处理SQLite不支持的列名，保持原始数据不变
            new_columns = []
            for i, col in enumerate(original_columns):
                original_col = str(col).strip()
                # 如果列名为空或包含SQLite不支持的字符，才进行修改
                if not original_col or not original_col.replace('_', '').replace('-', '').isalnum():
                    # 使用原始列名，只替换空格为下划线
                    clean_col = original_col.replace(' ', '_')
                    if not clean_col:
                        clean_col = f"col_{i}"
                    new_columns.append(clean_col)
                else:
                    # 保持原始列名
                    new_columns.append(original_col)
            
            df.columns = new_columns
            
            print(f"    列名: {list(df.columns)[:5]}...")  # 显示前5个列名
            print(f"    行数: {len(df)}")
            
            # 保持原始数据，不修改任何内容
            # 只处理SQLite不兼容的数据类型，但保持原始值
            print(f"    数据完整性: 保持原始Excel数据不变")
            
            # 将数据写入SQLite
            df.to_sql(clean_sheet_name, conn, if_exists='replace', index=False)
        
        conn.close()
        print(f"转换完成！SQLite文件已保存到: {sqlite_file}")
        return sqlite_file
        
    except Exception as e:
        print(f"转换过程中出现错误: {str(e)}")
        return None

def main():
    """主函数"""
    # 检查input目录中的Excel文件
    input_dir = "input"
    output_dir = "output"
    
    if not os.path.exists(input_dir):
        print(f"错误: {input_dir} 目录不存在")
        return
    
    # 查找Excel文件
    excel_files = []
    for file in os.listdir(input_dir):
        if file.lower().endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(input_dir, file))
    
    if not excel_files:
        print(f"在 {input_dir} 目录中没有找到Excel文件")
        print("请将Excel文件放入input目录中")
        return
    
    # 确保output目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 转换每个Excel文件
    success_count = 0
    total_count = len(excel_files)
    
    print(f"\n找到 {total_count} 个Excel文件，开始批量转换...")
    print("=" * 60)
    
    for i, excel_file in enumerate(excel_files, 1):
        print(f"\n[{i}/{total_count}] 开始处理: {excel_file}")
        result = excel_to_sqlite(excel_file)
        if result:
            print(f"成功转换: {result}")
            success_count += 1
        else:
            print(f"转换失败: {excel_file}")
    
    print("\n" + "=" * 60)
    print(f"批量转换完成！")
    print(f"成功: {success_count}/{total_count} 个文件")
    if success_count < total_count:
        print(f"失败: {total_count - success_count} 个文件")
    
    # 为每个成功转换的数据库生成报告
    if success_count > 0:
        print(f"\n正在生成数据库报告...")
        generate_database_reports()

def generate_database_reports():
    """为所有数据库生成报告"""
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        print("output目录不存在")
        return
    
    # 查找所有SQLite文件
    db_files = []
    for file in os.listdir(output_dir):
        if file.lower().endswith('.db'):
            db_files.append(os.path.join(output_dir, file))
    
    if not db_files:
        print("没有找到SQLite数据库文件")
        return
    
    # 为每个数据库生成报告
    for db_file in db_files:
        generate_single_report(db_file)

def generate_single_report(db_file):
    """为单个数据库生成报告"""
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 获取文件信息
        file_size = os.path.getsize(db_file)
        
        # 获取所有表
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        
        if not tables:
            print(f"数据库 {db_file} 中没有表")
            conn.close()
            return
        
        # 生成报告内容
        report_content = []
        report_content.append("=" * 80)
        report_content.append("SQLite数据库详细报告")
        report_content.append("=" * 80)
        report_content.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report_content.append(f"数据库文件: {db_file}")
        report_content.append(f"文件大小: {file_size:,} 字节")
        report_content.append(f"表数量: {len(tables)}")
        report_content.append("")
        
        total_rows = 0
        
        # 为每个表生成详细信息
        for table in tables:
            table_name = table[0]
            
            # 获取表结构
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            
            # 获取行数
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            row_count = cursor.fetchone()[0]
            total_rows += row_count
            
            report_content.append(f"表名: {table_name}")
            report_content.append("-" * 50)
            report_content.append(f"行数: {row_count:,}")
            report_content.append(f"列数: {len(columns)}")
            report_content.append("")
            report_content.append("列信息:")
            for i, col in enumerate(columns, 1):
                report_content.append(f"   {i}. {col[1]:<30} ({col[2]})")
            report_content.append("")
            
            # 获取示例数据（前3行）
            if row_count > 0:
                cursor.execute(f"SELECT * FROM {table_name} LIMIT 3")
                sample_rows = cursor.fetchall()
                
                report_content.append("示例数据 (前3行):")
                report_content.append("-" * 30)
                
                for i, row in enumerate(sample_rows, 1):
                    report_content.append(f"第{i}行:")
                    for j, (col, value) in enumerate(zip(columns, row)):
                        report_content.append(f"  {col[1]}: {value}")
                    report_content.append("")
                
                if row_count > 3:
                    report_content.append(f"... 还有 {row_count - 3:,} 行数据")
                    report_content.append("")
        
        report_content.append("=" * 80)
        report_content.append(f"总计记录数: {total_rows:,}")
        report_content.append("=" * 80)
        
        # 保存报告到文件
        report_file = db_file.replace('.db', '_report.txt')
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_content))
        
        print(f"报告已生成: {report_file}")
        
        conn.close()
        
    except Exception as e:
        print(f"生成报告时出错: {str(e)}")

if __name__ == "__main__":
    main()
