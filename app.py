import io
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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

uploaded_file = st.file_uploader("", type=["xlsx", "xls", "csv"])

def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", "null"]:
        return ""
    return s

def inspect_and_parse(file_bytes, file_name):
    if file_name.lower().endswith('.csv'):
        try:
            df_raw = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str)
        except Exception:
            df_raw = pd.read_csv(io.BytesIO(file_bytes), header=None, encoding='cp932', dtype=str)
    elif file_name.lower().endswith('.xls'):
        try:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine='xlrd')
        except Exception:
            df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
    else:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)

    # 1. ヘッダー検出＆列インデックス判定
    header_row_idx = -1
    col_map = {}

    for r in range(min(25, len(df_raw))):
        row_vals = [clean_str(v).lower() for v in df_raw.iloc[r].tolist()]
        temp_map = {}

        for c, text in enumerate(row_vals):
            if not text:
                continue

            # Scope
            if any(k in text for k in ["expert type", "scope", "part", "category", "タイプ", "スコープ"]):
                if "Scope" not in temp_map: temp_map["Scope"] = c
            # Number ("thirdbridge" などの誤一致を排除し、E列や "number", "no." を優先指定)
            elif any(k in text for k in ["number", "no", "no.", "id", "番号", "エキスパート番号"]) and "name" not in text and "thirdbridge" not in text:
                if "Number" not in temp_map: temp_map["Number"] = c
            # Name
            elif any(k in text for k in ["name", "expert name", "名前", "氏名", "エキスパート名"]) and "filename" not in text and "thirdbridge" not in text:
                if "Name" not in temp_map: temp_map["Name"] = c
            # Relevant Titles
            elif any(k in text for k in ["title", "titles", "役職", "タイトル", "要職"]):
                if "Relevant Titles" not in temp_map: temp_map["Relevant Titles"] = c
            # Relevant Experience
            elif any(k in text for k in ["experience", "screening", "経歴要約", "要約", "スクリーニング"]):
                if "Relevant experience" not in temp_map: temp_map["Relevant experience"] = c
            # Hourly Rate
            elif any(k in text for k in ["rate", "hourly", "fee", "price", "cost", "単価", "時給", "料金"]):
                if "Hourly Rate" not in temp_map: temp_map["Hourly Rate"] = c
            # Employment History
            elif any(k in text for k in ["employment", "history", "career", "経歴", "職歴", "職務経歴"]):
                if "Employment History" not in temp_map: temp_map["Employment History"] = c
            # Location
            elif any(k in text for k in ["location", "country", "city", "place", "所在地", "国", "場所"]):
                if "Location" not in temp_map: temp_map["Location"] = c
            # Status
            elif "status" in text or "ステータス" in text:
                if "Status" not in temp_map: temp_map["Status"] = c
            # Identity Verification
            elif any(k in text for k in ["identity", "verification", "id verification", "本人確認"]):
                if "Identity Verification Status" not in temp_map: temp_map["Identity Verification Status"] = c

        if ("Name" in temp_map or "Relevant experience" in temp_map) and len(temp_map) >= 2:
            header_row_idx = r
            col_map = temp_map
            break

    if header_row_idx == -1:
        header_row_idx = 0
        col_map = {"Scope": 0, "Number": 4, "Name": 1, "Relevant Titles": 2, "Relevant experience": 3, "Hourly Rate": 5, "Employment History": 6, "Location": 7}

    # E列 (index 4) が Number 列のデフォルト位置
    if "Number" not in col_map or col_map.get("Number") == col_map.get("Name"):
        col_map["Number"] = 4 if len(df_raw.columns) > 4 else 1

    has_id = "Identity Verification Status" in col_map

    parsed_rows = []
    
    for r in range(header_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[r]
        
        def get_field(key):
            if key in col_map:
                c = col_map[key]
                if c < len(row):
                    return clean_str(row[c])
            return ""

        name = get_field("Name")
        scope = get_field("Scope")
        number = get_field("Number")
        titles = get_field("Relevant Titles")
        exp = get_field("Relevant experience")
        rate_raw = get_field("Hourly Rate")
        emp = get_field("Employment History")
        loc = get_field("Location")
        status = get_field("Status")
        id_ver = get_field("Identity Verification Status")

        # Number列に "ThirdBridge" の文字列が紛れ込んだ場合のフィルタリング保護
        if "thirdbridge" in str(number).lower() or "third bridge" in str(number).lower():
            number = ""

        if not any([name, scope, number, titles, exp, rate_raw, emp, loc]):
            continue

        if parsed_rows and not name and not scope and not rate_raw and not number:
            prev = parsed_rows[-1]
            if exp:
                prev["exp"] = (prev["exp"] + "\n" + exp).strip()
            if titles:
                prev["titles"] = (prev["titles"] + "\n" + titles).strip()
            if emp:
                prev["emp"] = (prev["emp"] + "\n" + emp).strip()
            continue

        if not scope:
            scope = "その他"

        is_consulted = (status.lower() == "consulted")

        rate_val = rate_raw
        if rate_raw:
            num_match = re.search(r'[\d\.]+', rate_raw.replace(',', ''))
            if num_match:
                try:
                    rate_val = float(num_match.group(0))
                except ValueError:
                    rate_val = rate_raw

        id_sym = ""
        if has_id and id_ver:
            if "verified" in id_ver.lower() and "not" not in id_ver.lower():
                id_sym = "◯"
            elif "not verified" in id_ver.lower():
                id_sym = "×"
            else:
                id_sym = id_ver

        parsed_rows.append({
            "scope": scope,
            "number": number,
            "name": name,
            "titles": titles,
            "exp": exp,
            "rate": rate_val,
            "emp": emp,
            "loc": loc,
            "id_sym": id_sym,
            "is_consulted": is_consulted
        })

    return parsed_rows, has_id

def process_excel(file):
    file_bytes = file.read()
    data_items, has_id = inspect_and_parse(file_bytes, file.name)

    if not data_items:
        raise ValueError("有効なデータ行が見つかりませんでした。ファイルの内容を確認してください。")

    title_val = "ThirdBridge"
    scope_groups = {}

    formatted_data_list = []
    for item in data_items:
        formatted_row = [
            item["scope"],
            item["number"],
            item["name"],
            item["titles"],
            item["exp"],
            item["rate"],
            item["emp"],
            item["loc"]
        ]
        if has_id:
            formatted_row.append(item["id_sym"])

        entry = {"data": formatted_row, "is_consulted": item["is_consulted"]}
        formatted_data_list.append(entry)
        scope_groups.setdefault(item["scope"], []).append(entry)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    headers = ["Scope", "Number", "Name", "Relevant Titles", "Relevant experience", "Hourly Rate", "Employment History", "Location"]
    if has_id:
        headers.append("ID Verification")

    font_regular = Font(name="Meiryo UI", size=9)
    font_bold = Font(name="Meiryo UI", size=9, bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )

    fill_consulted = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_header = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    def build_sheet(ws, sheet_title, items):
        ws.freeze_panes = "D4" # 3行目・C列まで固定

        ws["A1"] = sheet_title
        ws["A1"].font = font_bold

        ws["D1"] = "✅ > screening済み"
        ws["D1"].font = font_bold

        ws["D2"] = "黄色セル > コンサル済み"
        ws["D2"].font = font_bold
        ws["D2"].fill = fill_consulted

        for col_num, h_text in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num, value=h_text)
            cell.font = font_bold
            cell.fill = fill_header
            cell.alignment = Alignment(vertical="top")

        for r_idx, item in enumerate(items, 4):
            row_data = item["data"]
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = font_regular

                if item["is_consulted"]:
                    cell.fill = fill_consulted
                else:
                    cell.fill = fill_white

                col_name = headers[c_idx - 1]
                cell.value = val

                if col_name == "Hourly Rate" and isinstance(val, (int, float)):
                    cell.number_format = '$#,##0'
                    cell.alignment = Alignment(horizontal="center", vertical="top")
                elif col_name in ["Number", "Location", "ID Verification"]:
                    cell.alignment = Alignment(horizontal="center", vertical="top")

        col_widths = [22, 10, 16, 38, 60, 14, 55, 10, 14]
        for c_i, w in enumerate(col_widths[:len(headers)], 1):
            ws.column_dimensions[get_column_letter(c_i)].width = w

    # 1. 全員一覧
    ws_all = wb.create_sheet(title="全員一覧")
    build_sheet(ws_all, title_val, formatted_data_list)

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
