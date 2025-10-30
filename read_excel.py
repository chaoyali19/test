#!/usr/bin/env python3
"""
Excel文件读取脚本
使用pandas库读取Excel文件并显示内容
"""

import pandas as pd
import sys

def read_excel_file(file_path):
    """
    读取Excel文件并返回DataFrame
    
    Args:
        file_path (str): Excel文件路径
        
    Returns:
        pandas.DataFrame: 包含Excel数据的DataFrame
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        print(f"成功读取文件: {file_path}")
        print(f"数据形状: {df.shape}")
        print("\n前5行数据:")
        print(df.head())
        return df
    except FileNotFoundError:
        print(f"错误: 文件 {file_path} 不存在")
        return None
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

def main():
    """主函数"""
    if len(sys.argv) != 2:
        print("用法: python read_excel.py <excel文件路径>")
        print("示例: python read_excel.py data.xlsx")
        return
    
    file_path = sys.argv[1]
    df = read_excel_file(file_path)
    
    if df is not None:
        # 显示列信息
        print("\n列信息:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")
        
        # 显示数据类型
        print("\n数据类型:")
        print(df.dtypes)

if __name__ == "__main__":
    main()