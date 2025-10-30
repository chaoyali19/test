#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel文件读取脚本
使用pandas库读取和处理Excel文件
"""

import pandas as pd
import sys
import os


def read_excel_file(file_path, sheet_name=0):
    """
    读取Excel文件
    
    参数:
        file_path (str): Excel文件路径
        sheet_name (str/int): 工作表名戗索引，默认第一个工作表
    
    返回:
        DataFrame: 包含Excel数据的DataFrame对象
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return None
        
        # 读取Excel文件
        print(f"正在读取Excel文件: {file_path}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # 显示基本信息
        print(f"\n文件读取成功!")
        print(f"数据形状: {df.shape}")
        print(f"列名: {list(df.columns)}")
        print(f"\n前5行数据:")
        print(df.head())
        
        return df
        
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return None


def save_to_csv(df, output_path):
    """
    将DataFrame保存为CSV文件
    
    参数:
        df (DataFrame): 要保存的数据
        output_path (str): 输出CSV文件路径
    """
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\n数据已保存到: {output_path}")
    except Exception as e:
        print(f"保存CSV文件时出错: {e}")


def main():
    """
    主函数
    """
    # 读取命令行参数
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        # 如果没有提供参数，请用户输入
        excel_file = input("请输入Excel文件路径: ")
    
    # 读取Excel文件
    df = read_excel_file(excel_file)
    
    if df is not None:
        # 生成输出文件名
        base_name = os.path.splitext(excel_file)[0]
        csv_file = f"{base_name}.csv"
        
        # 保存为CSV
        save_to_csv(df, csv_file)
        
        print(f"\n处理完成!")


if __name__ == "__main__":
    main()