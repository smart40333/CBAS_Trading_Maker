import os
import re
import shutil
from pathlib import Path
from datetime import datetime

def extract_date_from_filename(filename):
    """
    從檔名中提取日期
    格式: CBAS部位表與交易明細表_2025-05-08_呂○璇_mich.xlsx
    """
    # 使用正則表達式匹配日期格式 YYYY-MM-DD
    date_pattern = r'(\d{4}-\d{2}-\d{2})'
    match = re.search(date_pattern, filename)
    
    if match:
        date_str = match.group(1)
        try:
            # 驗證日期格式是否正確
            datetime.strptime(date_str, '%Y-%m-%d')
            return date_str
        except ValueError:
            return None
    return None

def convert_date_to_folder_name(date_str):
    """
    將日期字符串 YYYY-MM-DD 轉換為文件夾名稱 YYYYMMDD
    """
    return date_str.replace('-', '')

def organize_cbas_files(source_path, destination_path):
    """
    整理CBAS檔案從來源路徑到目的地路徑的對應日期文件夾
    """
    source_path = Path(source_path)
    destination_path = Path(destination_path)
    
    if not source_path.exists():
        print(f"錯誤：來源路徑不存在 - {source_path}")
        return
    
    if not destination_path.exists():
        print(f"錯誤：目的地路徑不存在 - {destination_path}")
        return
    
    # 記錄處理結果
    processed_files = []
    error_files = []
    
    # 掃描來源路徑中的所有Excel檔案
    excel_files = list(source_path.glob('*.xlsx')) + list(source_path.glob('*.xls'))
    
    print(f"在來源路徑找到 {len(excel_files)} 個Excel檔案")
    
    for file_path in excel_files:
        filename = file_path.name
        
        # 提取日期
        date_str = extract_date_from_filename(filename)
        
        if date_str:
            # 轉換為文件夾名稱
            folder_name = convert_date_to_folder_name(date_str)
            target_folder = destination_path / folder_name
            
            try:
                # 創建目標文件夾（如果不存在）
                target_folder.mkdir(exist_ok=True)
                
                # 目標檔案路徑
                target_file = target_folder / filename
                
                # 如果目標檔案已存在，先備份
                if target_file.exists():
                    backup_name = f"{target_file.stem}_backup_{datetime.now().strftime('%H%M%S')}{target_file.suffix}"
                    backup_path = target_folder / backup_name
                    print(f"檔案已存在，創建備份：{backup_name}")
                    shutil.move(str(target_file), str(backup_path))
                
                # 移動檔案
                shutil.move(str(file_path), str(target_file))
                processed_files.append((filename, folder_name))
                print(f"✓ 已移動：{filename} → 客戶部位表/{folder_name}/")
                
            except Exception as e:
                error_files.append((filename, str(e)))
                print(f"✗ 錯誤：無法移動 {filename} - {e}")
        else:
            error_files.append((filename, "無法從檔名中提取日期"))
            print(f"✗ 跳過：{filename} - 無法從檔名中提取日期")
    
    # 輸出總結
    print("\n" + "="*50)
    print("處理完成總結：")
    print(f"成功處理：{len(processed_files)} 個檔案")
    print(f"錯誤/跳過：{len(error_files)} 個檔案")
    
    if processed_files:
        print("\n成功處理的檔案：")
        for filename, folder_name in processed_files:
            print(f"  • {filename} → 客戶部位表/{folder_name}/")
    
    if error_files:
        print("\n錯誤/跳過的檔案：")
        for filename, error in error_files:
            print(f"  • {filename} - {error}")

def preview_organization(source_path, destination_path):
    """
    預覽整理結果，不實際移動檔案
    """
    source_path = Path(source_path)
    destination_path = Path(destination_path)
    
    if not source_path.exists():
        print(f"錯誤：來源路徑不存在 - {source_path}")
        return
    
    if not destination_path.exists():
        print(f"錯誤：目的地路徑不存在 - {destination_path}")
        return
    
    # 掃描來源路徑中的所有Excel檔案
    excel_files = list(source_path.glob('*.xlsx')) + list(source_path.glob('*.xls'))
    
    print(f"預覽模式：在來源路徑找到 {len(excel_files)} 個Excel檔案")
    print("\n將執行以下移動操作：")
    print("-" * 50)
    
    valid_files = []
    invalid_files = []
    
    for file_path in excel_files:
        filename = file_path.name
        date_str = extract_date_from_filename(filename)
        
        if date_str:
            folder_name = convert_date_to_folder_name(date_str)
            target_path = f"客戶部位表/{folder_name}/"
            valid_files.append((filename, target_path))
            print(f"✓ {filename} → {target_path}")
        else:
            invalid_files.append(filename)
            print(f"✗ {filename} - 無法提取日期")
    
    print("\n" + "="*50)
    print(f"可處理檔案：{len(valid_files)} 個")
    print(f"無法處理檔案：{len(invalid_files)} 個")

def main():
    # CBAS檔案路徑
    source_path = r"\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\客戶部位表(對帳單)\庫存部位\歷史CBAS部位表與交易明細表"
    destination_path = r"\\10.72.228.112\cbas業務公用區\!!!交易作業區!!!\客戶部位表(對帳單)\庫存部位\客戶部位表"
    
    print("CBAS檔案整理工具")
    print("="*50)
    print(f"來源路徑：{source_path}")
    print(f"目的地路徑：{destination_path}")
    
    # 選擇操作模式
    print("\n請選擇操作模式：")
    print("1. 預覽模式（只查看不移動）")
    print("2. 執行移動")
    
    choice = input("請輸入選項 (1/2): ")
    
    if choice == "1":
        preview_organization(source_path, destination_path)
    elif choice == "2":
        user_input = input("\n確認執行檔案移動操作？(y/N): ")
        if user_input.lower() in ['y', 'yes']:
            organize_cbas_files(source_path, destination_path)
        else:
            print("已取消操作")
    else:
        print("無效選項，已退出")

if __name__ == "__main__":
    main() 