import pandas as pd
import sys

def read_excel_file(file_path):
    """
    读取Excel文件并输出内容，不显示行数和列数
    """
    try:
        # 读取Excel文件，header=None确保不将第一行作为列名
        df = pd.read_excel(file_path, header=None)
        
        # 输出所有数据，不包含列名和索引
        for index, row in df.iterrows():
            # 将每行数据转换为列表并输出
            row_data = [str(cell) for cell in row]
            print('\t'.join(row_data))
            
    except FileNotFoundError:
        print(f"错误: 文件 '{file_path}' 未找到")
        sys.exit(1)
    except Exception as e:
        print(f"错误: 读取Excel文件时出错 - {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("用法: python excel_processor.py <excel文件路径>")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    read_excel_file(excel_file)