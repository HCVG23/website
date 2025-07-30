import pandas as pd
import json

def excel_to_json_by_sheet(excel_file_path, sheet_name, json_file_path):
    """
    讀取Excel文件中的特定工作表，並將其轉換為JSON格式。

    Args:
        excel_file_path (str): Excel文件的路徑。
        sheet_name (str): 要讀取的Excel工作表的名稱。
        json_file_path (str): 儲存JSON文件的路徑。
    """
    try:
        # 使用pandas讀取指定的Excel工作表
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

        # 將DataFrame轉換為JSON
        json_data = df.to_json(orient="records", force_ascii=False, indent=4)

        # 將JSON數據寫入文件
        with open(json_file_path, 'w', encoding='utf-8') as f:
            f.write(json_data)

        print(f"成功將工作表 '{sheet_name}' 從 '{excel_file_path}' 轉換為 '{json_file_path}'")

    except FileNotFoundError:
        print(f"錯誤: 文件 '{excel_file_path}' 未找到。")
    except KeyError:
        print(f"錯誤: 工作表 '{sheet_name}' 在 '{excel_file_path}' 中不存在。")
    except Exception as e:
        print(f"轉換過程中發生錯誤: {e}")

# 範例用法
excel_file = "甄選分析.xlsx"  # 替換為您的Excel文件路徑
sheet_name_to_convert = "工作表2"  # 替換為您要轉換的工作表名稱
json_output_file = "data_107024_111.json" # 替換為您要儲存JSON文件的路徑

excel_to_json_by_sheet(excel_file, sheet_name_to_convert, json_output_file)