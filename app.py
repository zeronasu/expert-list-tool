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
        
    header_idx = -1
    col_map = {}
    
    # あいまいヘッダー自動マッチングロジック
    for r in range(min(15, len(df_raw))):
        row_vals = [str(val).replace('\n', ' ').strip().lower() for val in df_raw.iloc[r].tolist()]
        temp_map = {}
        
        for c, val in enumerate(row_vals):
            if not val or val == "nan": continue
            
            if "name" in val and "Name" not in temp_map:
                temp_map["Name"] = c
            elif ("expert type" in val or "scope" in val or "type" in val) and "Expert Type" not in temp_map:
                temp_map["Expert Type"] = c
            elif ("experience" in val or "screening" in val) and "Relevant experience" not in temp_map:
                temp_map["Relevant experience"] = c
            elif "status" in val and "Status" not in temp_map:
                temp_map["Status"] = c
            elif ("rate" in val or "hourly" in val or "fee" in val) and "Hourly Rate" not in temp_map:
                temp_map["Hourly Rate"] = c
            elif ("number" in val or val == "no" or val == "no.") and "Number" not in temp_map:
                temp_map["Number"] = c
            elif "title" in val and "Relevant Titles" not in temp_map:
                temp_map["Relevant Titles"] = c
            elif ("employment" in val or "history" in val) and "Employment History" not in temp_map:
                temp_map["Employment History"] = c
            elif ("location" in val or "country" in val or "city" in val) and "Location" not in temp_map:
                temp_map["Location"] = c
            elif ("identity" in val or "id" in val or "verification" in val) and "Identity Verification Status" not in temp_map:
                temp_map["Identity Verification Status"] = c

        if "Name" in temp_map and len(temp_map) >= 3:
            header_idx = r
            col_map = temp_map
            break
            
    if header_idx == -1 or "Name" not in col_map:
        raise ValueError("ヘッダー項目（Name, Scope など）が検出できませんでした。データ形式を確認してください。")

    title_val = "ThirdBridge"
    has_id = "Identity Verification Status" in col_map
    
    data_rows = []
    scope_groups = {}
    
    for r in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[r]
        
        def get_val(key):
            if key in col_map:
                c_idx = col_map[key]
                val = row[c_idx]
                if pd.notna(val) and str(val).strip().lower() != "nan":
                    return row[c_idx]
            return ""

        name_val = str(get_val("Name")).strip()
        scope_val = str(get_val("Expert Type")).strip()
        
        if not name_val and not scope_val:
            continue
        if not scope_val:
            scope_val = "その他"
            
        raw_status = str(get_val("Status")).strip()
        is_consulted = (raw_status.lower() == "consulted")
        
        number_val = get_val("Number")
        titles_val = get_val("Relevant Titles")
        exp_val = get_val("Relevant experience")
        loc_val = get_val("Location")
        
        # Rate処理
        rate_val = get_val("Hourly Rate")
        if isinstance(rate_val, str):
            match = re.search(r'[\d\.]+', rate_val)
            rate_val = float(match.group(0)) if match else rate_val
            
        # Employment History処理
        emp_val = str(get_val("Employment History"))
        cleaned_emp = "\n".join([line.strip() for line in emp_val.split("\n") if line.strip() and line.strip().lower() != "nan"])
        
        formatted_row = [scope_val, number_val, name_val, titles_val, exp_val, rate_val, cleaned_emp, loc_val]
        
        if has_id:
            raw_id = str(get_val("Identity Verification Status")).strip()
            if "verified" in raw_id.lower() and "not" not in raw_id.lower(): id_sym = "◯"
            elif "not verified" in raw_id.lower(): id_sym = "×"
            else: id_sym = raw_id
            formatted_row.append(id_sym)
            
        item = {"data": formatted_row, "is_consulted": is_consulted}
        data_rows.append(item)
        scope_groups.setdefault(scope_val, []).append(item)
        
    # openpyxl によるブック生成
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    headers = ["Scope", "Number", "Name", "Relevant Titles", "Relevant experience", "Hourly Rate", "Employment History", "Location"]
    if has_id: headers.append("ID Verification")
    
    def build_sheet(ws, title, items):
        ws.freeze_panes = "D4" # 3行目・C列固定
        
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
                    
        # 列幅調整
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
