#!/usr/bin/env python3
"""
Excel文件读取脚本
使用pandas库读取Excel文件并显示内容
"""

import pandas as pd
import sys
import os

def read_excel_file(file_path):
    """
    读取Excel文件并显示内容
    
    Args:
        file_path (str): Excel文件路径
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return
        
        # 读取Excel文件
        print(f"正在读取Excel文件: {file_path}")
        
        # 读取所有工作表
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        print(f"\nExcel文件包含 {len(sheet_names)} 个工作表:")
        for i, sheet in enumerate(sheet_names, 1):
            print(f"  {i}. {sheet}")
        
        # 读取每个工作表的内容
        for sheet_name in sheet_names:
            print(f"\n{'='*50}")
            print(f"工作表: {sheet_name}")
            print(f"{'='*50}")
            
            # 读取工作表数据
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # 显示基本信息
            print(f"数据形状: {df.shape}")
            print(f"列名: {list(df.columns)}")
            
            # 显示前几行数据
            print("\n前5行数据:")
            print(df.head())
            
            # 显示数据类型信息
            print("\n数据类型:")
            print(df.dtypes)
            
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")

def main():
    """主函数"""
    if len(sys.argv) != 2:
        print("用法: python read_excel.py <excel_file_path>")
        print("示例: python read_excel.py data.xlsx")
        sys.exit(1)
    
    file_path = sys.argv[1]
    read_excel_file(file_path)

if __name__ == "__main__":
    main()