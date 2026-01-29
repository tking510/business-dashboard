#!/usr/bin/env python3
"""
Excelファイルからダッシュボードのdata.jsを自動生成するスクリプト
"""
import openpyxl
import pandas as pd
import json
import glob
from datetime import datetime
from pathlib import Path

# データ構造を初期化
MOTO_AMUSE_DATA = {}
DSC_DATA = {}
SLOTEN_DATA = {}
KONIBET_DATA = {}

print("=== データ抽出開始 ===\n")

# dataディレクトリのパス
data_dir = Path(__file__).parent.parent / "data"
data_dir.mkdir(exist_ok=True)

# 1. SAMファイル（元amuseとDSC）を処理
sam_files = list(data_dir.glob("*SAM*.xlsx")) + list(data_dir.glob("収益管理*.xlsx"))
for sam_file in sam_files:
    try:
        print(f"処理中: {sam_file.name}")
        wb = openpyxl.load_workbook(sam_file, data_only=True)
        for sheet_name in wb.sheetnames:
            if '収益' in sheet_name:
                is_dsc = sheet_name.endswith('D')
                period_name = sheet_name.replace('収益', '').replace('D', '')
                
                sheet = wb[sheet_name]
                data = {}
                
                for row in sheet.iter_rows(min_row=1, values_only=True):
                    if len(row) > 2 and row[1] and row[2] is not None:
                        key = str(row[1]).strip()
                        try:
                            val = float(row[2]) if isinstance(row[2], (int, float)) else None
                            if val is not None:
                                data[key] = val
                        except:
                            pass
                
                if data:
                    if is_dsc:
                        DSC_DATA[period_name] = {'サマリー': data}
                    else:
                        MOTO_AMUSE_DATA[period_name] = {'サマリー': data}
        
        print(f"  ✓ 元amuse: {len(MOTO_AMUSE_DATA)}件, DSC: {len(DSC_DATA)}件")
    except Exception as e:
        print(f"  ✗ エラー: {e}")

# 2. スロ天KPIシートを処理
sloten_files = list(data_dir.glob("*スロ天*.xlsx")) + list(data_dir.glob("*KPI*.xlsx"))
for sloten_file in sloten_files:
    try:
        print(f"処理中: {sloten_file.name}")
        wb = openpyxl.load_workbook(sloten_file, data_only=True)
        for sheet_name in wb.sheetnames:
            if len(sheet_name) == 6 and sheet_name.isdigit():
                year = sheet_name[:4]
                month = int(sheet_name[4:6])
                period_name = f"{year}年{month}月"
                
                sheet = wb[sheet_name]
                kpi = {}
                seen_keys = set()
                
                for row in sheet.iter_rows(min_row=1, max_row=200, values_only=True):
                    if row[0] and row[1] is not None:
                        key = str(row[0]).strip().split('/')[0].strip()
                        try:
                            val = float(row[1]) if isinstance(row[1], (int, float)) else None
                            if val is not None and key and not key.startswith('0.'):
                                if key not in seen_keys:
                                    kpi[key] = val
                                    seen_keys.add(key)
                        except:
                            pass
                
                if kpi:
                    SLOTEN_DATA[period_name] = {'サマリー': kpi}
        
        print(f"  ✓ スロ天: {len(SLOTEN_DATA)}件")
    except Exception as e:
        print(f"  ✗ エラー: {e}")

# 3. KonibetCSVを処理
konibet_files = list(data_dir.glob("*日报*.csv")) + list(data_dir.glob("*データ総和*.csv"))
for konibet_file in konibet_files:
    try:
        print(f"処理中: {konibet_file.name}")
        df = pd.read_csv(konibet_file)
        for _, row in df.iterrows():
            date_str = str(row.iloc[0])
            try:
                date_obj = pd.to_datetime(date_str)
                period_name = f"{date_obj.year}年{date_obj.month}月"
                
                if period_name not in KONIBET_DATA:
                    KONIBET_DATA[period_name] = {'サマリー': {}}
                
                for col in df.columns[1:]:
                    try:
                        val = float(row[col])
                        if pd.notna(val):
                            if col not in KONIBET_DATA[period_name]['サマリー']:
                                KONIBET_DATA[period_name]['サマリー'][col] = 0
                            KONIBET_DATA[period_name]['サマリー'][col] += val
                    except:
                        pass
            except:
                pass
        
        print(f"  ✓ Konibet: {len(KONIBET_DATA)}件")
    except Exception as e:
        print(f"  ✗ エラー: {e}")

# 既存のdata.jsから欠落データを補完
output_file = Path(__file__).parent.parent / "data.js"
if output_file.exists():
    print("\n既存のdata.jsからデータを補完中...")
    import re
    
    with open(output_file, 'r', encoding='utf-8') as f:
        existing_content = f.read()
    
    def extract_data_object(content, var_name):
        pattern = rf'const {var_name} = ({{.*?}});'
        match = re.search(pattern, content, re.DOTALL)
        if match:
            js_obj = match.group(1)
            json_str = js_obj.replace("'", '"')
            try:
                return json.loads(json_str)
            except:
                return None
        return None
    
    existing_konibet = extract_data_object(existing_content, 'KONIBET_DATA')
    existing_dsc = extract_data_object(existing_content, 'DSC_DATA')
    existing_moto = extract_data_object(existing_content, 'MOTO_AMUSE_DATA')
    existing_sloten = extract_data_object(existing_content, 'SLOTEN_DATA')
    
    # 既存データで補完（新しいデータで上書き）
    if existing_konibet:
        for key, val in existing_konibet.items():
            if key not in KONIBET_DATA:
                KONIBET_DATA[key] = val
    
    if existing_dsc:
        for key, val in existing_dsc.items():
            if key not in DSC_DATA:
                DSC_DATA[key] = val
    
    if existing_moto:
        for key, val in existing_moto.items():
            if key not in MOTO_AMUSE_DATA:
                MOTO_AMUSE_DATA[key] = val
    
    if existing_sloten:
        for key, val in existing_sloten.items():
            if key not in SLOTEN_DATA:
                SLOTEN_DATA[key] = val

print(f"\n=== 最終データ件数 ===")
print(f"元amuse: {len(MOTO_AMUSE_DATA)}件")
print(f"DSC: {len(DSC_DATA)}件")
print(f"スロ天: {len(SLOTEN_DATA)}件")
print(f"Konibet: {len(KONIBET_DATA)}件")

# data.jsファイルを生成
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(f"// ダッシュボードデータファイル\n")
    f.write(f"// 最終更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    
    def to_js(data):
        return json.dumps(data, ensure_ascii=False, indent=2).replace('"', "'")
    
    f.write(f"const MOTO_AMUSE_DATA = {to_js(MOTO_AMUSE_DATA)};\n\n")
    f.write(f"const DSC_DATA = {to_js(DSC_DATA)};\n\n")
    f.write(f"const SLOTEN_DATA = {to_js(SLOTEN_DATA)};\n\n")
    f.write(f"const KONIBET_DATA = {to_js(KONIBET_DATA)};\n")

print(f"\n✓ data.js生成完了: {output_file}")
print(f"  ファイルサイズ: {output_file.stat().st_size} bytes")
