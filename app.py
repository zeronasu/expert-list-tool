import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont
from openpyxl.text.richtext import TextBlock, CellRichText
import streamlit as st

st.set_page_config(page_title="エクセルリスト生成ツール", page_icon="⚡", layout="wide")

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
        max-width: 650px;
        margin: 0 auto 30px auto;
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

def normalize_text(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()

def find_header_and_map(df_raw):
    # 最初の20行からヘッダー行と全列インデックスを精密検出
    for r in range(min(20, len(df_raw))):
        row_cells = [normalize_text(val).lower() for val in df_raw.iloc[r].tolist()]
        
        col_map = {}
        for c, cell_text in enumerate(row_cells):
            if not cell_text:
                continue
            
            if "number" in cell_text or cell_text == "no" or cell_text == "no.":
                col_map["Number"] = c
            elif "name" in cell_text and "filename" not in cell_text:
                col_map["Name"] = c
            elif "title" in cell_text:
                col_map["Relevant Titles"] = c
            elif "experience" in cell_text or "screening" in cell_text:
                col_map["Relevant experience"] = c
            elif "rate" in cell_text or "hourly" in cell_text or "fee" in cell_text or "price" in cell_text:
                col_map["Hourly Rate"] = c
            elif "employment" in cell_text or "history" in cell_text or "career" in cell_text:
                col_map["Employment History"] = c
            elif "location" in cell_text or "country" in cell_text or "city" in cell_text:
                col_map["Location"] = c
            elif "expert type" in cell_text or "scope" in cell_text or "category" in cell_text or "type" in cell_text:
                col_map["Scope"] = c
            elif "status" in cell_text:
                col_map["Status"] = c
            elif "identity" in cell_text or "verification" in cell_text or "id" in cell_text:
                col_map["Identity Verification Status"] = c

        if "Name" in col_map and len(col_map) >= 3:
            return r, col_map
            
    return -1, {}

def process_excel(file):
    if file.name.endswith('.csv'):
        df_raw = pd.read_csv(file, header=None)
    else:
        df_raw = pd.read_excel(file, header=None)

    header_idx, col_map = find_header_and_map(df_raw)
    
    if header_idx == -1 or "Name" not in col_map:
        raise ValueError("ヘッダー項目（Name, Number, Relevant Titles など）が検出できませんでした。ファイル形式を確認してください。")

    title_val = "ThirdBridge"
    has_id = "Identity Verification Status" in col_map
    
    data_rows = []
    scope_groups = {}
    
    for r in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[r]
        
        def safe_get(key):
            if key in col_map:
                idx = col_map[key]
                if idx < len(row):
                    val = row[idx]
                    if pd.notna(val) and str(val).strip().lower() != "nan":
                        return val
            return ""

        name_val = normalize_text(safe_get("Name"))
        scope_val = normalize_text(safe_get("Scope"))
        
        if not name_val and not scope_val:
            continue
            
        if not scope_val:
            scope_val = "その他"
            
        raw_status = normalize_text(safe_get("Status"))
        is_consulted = (raw_status.lower() == "consulted")
        
        number_val = safe_get("Number")
        titles_val = safe_get("Relevant Titles")
        exp_val = safe_get("Relevant experience")
        loc_val = safe_get("Location")
        
        # Hourly Rateの解析
        rate_val = safe_get("Hourly Rate")
        if isinstance(rate_val, str):
            match = re.search(r'[\d\.]+', rate_val.replace(',', ''))
            rate_val = float(match.group(0)) if match else rate_val
            
        # Employment Historyの整形
        emp_raw = safe_get("Employment History")
        cleaned_emp_lines = []
        if emp_raw:
            lines = str(emp_raw).split('\n')
            for line in lines:
                l_str = line.strip()
                if l_str and l_str.lower() != "nan":
                    cleaned_emp_lines.append(l_str)
        cleaned_emp = "\n".join(cleaned_emp_lines)
        
        formatted_row = [
            scope_val,
            number_val,
            name_val,
            titles_val,
            exp_val,
            rate_val,
            cleaned_emp,
            loc_val
        ]
        
        if has_id:
            raw_id = normalize_text(safe_get("Identity Verification Status"))
            if "verified" in raw_id.lower() and "not" not in raw_id.lower():
                id_sym = "◯"
            elif "not verified" in raw_id.lower():
                id_sym = "×"
            else:
                id_sym = raw_id
            formatted_row.append(id_sym)
            
        item = {"data": formatted_row, "is_consulted": is_consulted}
        data_rows.append(item)
        scope_groups.setdefault(scope_val, []).append(item)
        
    # openpyxl ワークブック構築
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    headers = ["Scope", "Number", "Name", "Relevant Titles", "Relevant experience", "Hourly Rate", "Employment History", "Location"]
    if has_id:
        headers.append("ID Verification")
    
    font_regular = Font(name="Meiryo UI", size=9)
    font_bold = Font(name="Meiryo UI", size=9, bold=True)
    
    inline_regular = InlineFont(rFont="Meiryo UI", sz=9.0)
    inline_bold_red = InlineFont(rFont="Meiryo UI", sz=9.0, b=True, color="CC0000")
    
    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )
    
    fill_consulted = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_header = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    def build_sheet(ws, title, items):
        ws.freeze_panes = "D4" # 3行目・C列まで固定
        
        # タイトル行（A1）とステータス凡例（D1, D2）
        ws["A1"] = title
        ws["A1"].font = font_bold
        
        ws["D1"] = "✅ > screening済み"
        ws["D1"].font = font_bold
        
        ws["D2"] = "黄色セル > コンサル済み"
        ws["D2"].font = font_bold
        ws["D2"].fill = fill_consulted
        
        # ヘッダー行 (3行目)
        for col_num, h_text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num, value=h_text)
            cell.font = font_bold
            cell.fill = fill_header
            cell.alignment = Alignment(vertical="top")
            
        # データ行 (4行目以降)
        for r_idx, item in enumerate(items, 4):
            row_data = item["data"]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                
                if item["is_consulted"]:
                    cell.fill = fill_consulted
                else:
                    cell.fill = fill_white

                col_name = headers[c_idx - 1]
                
                if col_name == "Hourly Rate" and isinstance(val, (int, float)):
                    cell.value = val
                    cell.font = font_regular
                    cell.number_format = '$#,##0'
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                elif col_name in ["Number", "Location", "ID Verification"]:
                    cell.value = val
                    cell.font = font_regular
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                elif col_name in ["Relevant experience", "Employment History"] and isinstance(val, str) and "Present" in val:
                    # "Present" を含む行の太字・赤字化（リッチテキスト）
                    lines = val.split('\n')
                    rich_blocks = []
                    for idx, line in enumerate(lines):
                        if "Present" in line:
                            rich_blocks.append(TextBlock(inlineFont=inline_bold_red, text=line))
                        else:
                            rich_blocks.append(TextBlock(inlineFont=inline_regular, text=line))
                        if idx < len(lines) - 1:
                            rich_blocks.append(TextBlock(inlineFont=inline_regular, text='\n'))
                    
                    cell.value = CellRichText(rich_blocks)
                else:
                    cell.value = val
                    cell.font = font_regular
                    
        # 列幅の設定
        col_widths = [22, 10, 16, 38, 60, 14, 55, 10, 14]
        for c_i, w in enumerate(col_widths[:len(headers)], 1):
            ws.column_dimensions[get_column_letter(c_i)].width = w

    # 1. 全員一覧シート
    ws_all = wb.create_sheet(title="全員一覧")
    build_sheet(ws_all, title_val, data_rows)
    
    # 2. Scope別シート
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
