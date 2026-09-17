import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(page_title="エクセルリスト生成ツール", page_icon="⚡", layout="centered")

# カスタムCSS（ワインレッド＆クリームデザイン）
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #1A0507 0%, #3B0A11 40%, #801220 100%);
    }
    .main-card {
        background-color: rgba(255, 255, 255, 0.96);
        border-radius: 24px;
        padding: 40px;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.5);
        color: #222222;
        text-align: center;
    }
    .badge {
        background-color: #801220;
        color: #FFFFFF;
        padding: 6px 16px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 2px;
        text-transform: uppercase;
    }
    h1 {
        color: #1A0507;
        font-size: 28px;
        font-weight: 900;
        margin-top: 15px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-card"><span class="badge">Professional Tool</span><h1>エクセルリスト生成</h1><p style="color: #665555; font-size: 13px;">ファイルをアップロードするだけで、自動デザイン整形・文字装飾ルール適用・Scope別タブ分割を行います。</p></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["xlsx", "csv"])

def process_excel(file):
    # Excel / CSV の読み込み
    if file.name.endswith('.csv'):
        df_raw = pd.read_csv(file, header=None)
    else:
        df_raw = pd.read_excel(file, header=None)
        
    # ヘッダー行の自動検索
    header_idx = -1
    col_map = {}
    
    for r in range(min(15, len(df_raw))):
        row_vals = [str(val).strip().lower() for val in df_raw.iloc[r].tolist()]
        temp_map = {}
        found = 0
        for c, val in enumerate(row_vals):
            if val == "name": temp_map["Name"] = c; found += 1
            elif val in ["expert type", "scope"]: temp_map["Expert Type"] = c; found += 1
            elif val == "relevant experience": temp_map["Relevant experience"] = c; found += 1
            elif val == "status": temp_map["Status"] = c; found += 1
            elif val == "hourly rate": temp_map["Hourly Rate"] = c; found += 1
            elif val == "number": temp_map["Number"] = c; found += 1
            elif val == "relevant titles": temp_map["Relevant Titles"] = c; found += 1
            elif val == "employment history": temp_map["Employment History"] = c; found += 1
            elif val == "location": temp_map["Location"] = c; found += 1
            elif "identity verification" in val or "id verification" in val: temp_map["Identity Verification Status"] = c; found += 1
            
        if found >= 3:
            header_idx = r
            col_map = temp_map;
            break
            
    if header_idx == -1 or "Name" not in col_map:
        raise ValueError("ヘッダー項目（Name, Expert Type など）が正しく認識できませんでした。")

    # タイトルを一律「ThirdBridge」に固定
    title_val = "ThirdBridge"
    has_id = "Identity Verification Status" in col_map
    
    # データの抽出と整理
    data_rows = []
    scope_groups = {}
    
    for r in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[r]
        name_val = str(row[col_map["Name"]]).strip() if "Name" in col_map and pd.notna(row[col_map["Name"]]) else ""
        scope_val = str(row[col_map["Expert Type"]]).strip() if "Expert Type" in col_map and pd.notna(row[col_map["Expert Type"]]) else ""
        
        if not name_val and not scope_val:
            continue
        if not scope_val or scope_val.lower() == "nan":
            scope_val = "その他"
            
        raw_status = str(row[col_map["Status"]]).strip() if "Status" in col_map and pd.notna(row[col_map["Status"]]) else ""
        is_consulted = (raw_status.lower() == "consulted")
        
        number_val = row[col_map["Number"]] if "Number" in col_map and pd.notna(row[col_map["Number"]]) else ""
        titles_val = row[col_map["Relevant Titles"]] if "Relevant Titles" in col_map and pd.notna(row[col_map["Relevant Titles"]]) else ""
        exp_val = row[col_map["Relevant experience"]] if "Relevant experience" in col_map and pd.notna(row[col_map["Relevant experience"]]) else ""
        loc_val = row[col_map["Location"]] if "Location" in col_map and pd.notna(row[col_map["Location"]]) else ""
        
        # Rate
        rate_val = row[col_map["Hourly Rate"]] if "Hourly Rate" in col_map and pd.notna(row[col_map["Hourly Rate"]]) else ""
        if isinstance(rate_val, str):
            match = re.search(r'[\d\.]+', rate_val)
            rate_val = float(match.group(0)) if match else rate_val
            
        # Employment History
        emp_val = str(row[col_map["Employment History"]]) if "Employment History" in col_map and pd.notna(row[col_map["Employment History"]]) else ""
        cleaned_emp = "\n".join([line.strip() for line in emp_val.split("\n") if line.strip()])
        
        formatted_row = [scope_val, number_val, name_val, titles_val, exp_val, rate_val, cleaned_emp, loc_val]
        
        if has_id:
            raw_id = str(row[col_map["Identity Verification Status"]]).strip()
            if "verified" in raw_id.lower() and "not" not in raw_id.lower(): id_sym = "◯"
            elif "not verified" in raw_id.lower(): id_sym = "×"
            else: id_sym = raw_id
            formatted_row.append(id_sym)
            
        item = {"data": formatted_row, "is_consulted": is_consulted}
        data_rows.append(item)
        scope_groups.setdefault(scope_val, []).append(item)
        
    # openpyxl によるブック生成と書式設定
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # デフォルトシート削除
    
    headers = ["Scope", "Number", "Name", "Relevant Titles", "Relevant experience", "Hourly Rate", "Employment History", "Location"]
    if has_id: headers.append("ID Verification")
    
    def build_sheet(ws, title, items):
        ws.freeze_panes = "D4" # 3行目・C列まで固定
        
        # タイトル & 凡例
        ws["A1"] = title; ws["A1"].font = Font(name="Meiryo UI", size=9, bold=True)
        ws["D1"] = "✅ > screening済み"; ws["D1"].font = Font(name="Meiryo UI", size=9, bold=True)
        ws["D2"] = "黄色セル > コンサル済み"; ws["D2"].font = Font(name="Meiryo UI", size=9, bold=True)
        ws["D2"].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        # ヘッダー
        for col_num, h_text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num, value=h_text)
            cell.font = Font(name="Meiryo UI", size=9, bold=True)
            cell.fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
            cell.alignment = Alignment(vertical="top")
            
        # データ行
        thin_border = Border(
            left=Side(style='thin', color='D0D0D0'),
            right=Side(style='thin', color='D0D0D0'),
            top=Side(style='thin', color='D0D0D0'),
            bottom=Side(style='thin', color='D0D0D0')
        )
        
        for r_idx, item in enumerate(items, 4):
            row_data = item["data"]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = Font(name="Meiryo UI", size=9)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                
                # 配置とフォーマット
                if c_idx in [2, 8, 9]: # Number, Location, ID
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                elif c_idx == 6 and isinstance(val, (int, float)): # Rate
                    cell.number_format = '$#,##0'
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                    
                if item["is_consulted"]:
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    
        # 列幅設定
        col_widths = [22, 8, 14, 38, 60, 12, 50, 9, 12]
        for c_i, w in enumerate(col_widths[:len(headers)], 1):
            ws.column_dimensions[get_column_letter(c_i)].width = w

    # 1. 全員一覧
    ws_all = wb.create_sheet(title="全員一覧")
    build_sheet(ws_all, title_val, data_rows)
    
    # 2. Scope別タブ
    for scope_name, items in scope_groups.items():
        clean_name = re.sub(r'[\\View/*?\[\]]', '', scope_name)[:30]
        if clean_name:
            ws_scope = wb.create_sheet(title=clean_name)
            build_sheet(ws_scope, f"{title_val} ({scope_name})", items)
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

if uploaded_file is not None:
    if st.button("変換してダウンロード", type="primary"):
        with st.spinner("データを正確に処理中..."):
            try:
                processed_data = process_excel(uploaded_file)
                st.success("✨ 変換が完了しました！")
                st.download_button(
                    label="📥 エクセルファイルをダウンロード",
                    data=processed_data,
                    file_name=f"Formatted_{uploaded_file.name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
