import io
import re
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
import xlrd
from openpyxl import Workbook
import streamlit as st

st.set_page_config(page_title="전기사용량 50 미만 세대 추출기", layout="wide")

st.title("⚡ 전기사용량 50 미만 세대 자동 추출 시스템")
st.markdown("한일베라체, 연동드림아이 전용 파서 및 세로형·분할형 표를 지원하는 범용 키워드 탐색 구조를 제공합니다.")

def clean_num(val):
    if val is None or val == '' or val == '-':
        return None
    val_str = str(val).strip().replace(",", "")
    if re.fullmatch(r"[-+]?\d*\.?\d+", val_str):
        try:
            return float(val_str)
        except ValueError:
            return None
    return None

@st.cache_data(show_spinner=False)
def load_any_excel_to_openpyxl(file_bytes, filename):
    if file_bytes.startswith(b'<?xml') or b'urn:schemas-microsoft-com:office:spreadsheet' in file_bytes:
        try:
            root = ET.fromstring(file_bytes)
            pyxl_wb = Workbook()
            pyxl_wb.remove(pyxl_wb.active)
            namespaces = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
            for ws_elem in root.findall('.//ss:Worksheet', namespaces):
                ws_name = ws_elem.attrib.get('{urn:schemas-microsoft-com:office:spreadsheet}Name', 'Sheet1')
                target_ws = pyxl_wb.create_sheet(title=ws_name)
                table = ws_elem.find('ss:Table', namespaces)
                if table is not None:
                    for row_elem in table.findall('ss:Row', namespaces):
                        row_vals = []
                        for cell_elem in row_elem.findall('ss:Cell', namespaces):
                            index_attr = cell_elem.attrib.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                            if index_attr:
                                target_col_idx = int(index_attr) - 1
                                while len(row_vals) < target_col_idx:
                                    row_vals.append(None)
                            data_elem = cell_elem.find('ss:Data', namespaces)
                            row_vals.append(data_elem.text if data_elem is not None else None)
                        target_ws.append(row_vals)
            if pyxl_wb.sheetnames:
                return pyxl_wb
        except Exception:
            pass

    pyxl_wb = Workbook()
    pyxl_wb.remove(pyxl_wb.active)

    try:
        xls_wb = xlrd.open_workbook(file_contents=file_bytes)
        for sheet_name in xls_wb.sheet_names():
            xls_sheet = xls_wb.sheet_by_name(sheet_name)
            target_ws = pyxl_wb.create_sheet(title=sheet_name)
            for r in range(xls_sheet.nrows):
                target_ws.append(xls_sheet.row_values(r))
        if pyxl_wb.sheetnames:
            return pyxl_wb
    except Exception:
        pass

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            target_ws = pyxl_wb.create_sheet(title=sheetname)
            for r in range(1, ws.max_row + 1):
                target_ws.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        if pyxl_wb.sheetnames:
            return pyxl_wb
    except Exception:
        pass

    return pyxl_wb

def finalize_dataframe(df):
    if not df.empty:
        df = df.drop_duplicates().reset_index(drop=True)
        if '동' in df.columns and df['동'].astype(str).str.strip().ne('').any() and not df['동'].isna().all():
            df = df.assign(
                sort_dong=pd.to_numeric(df['동'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0),
                sort_ho=pd.to_numeric(df['구분/호수'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            ).sort_values(by=['sort_dong', 'sort_ho']).drop(columns=['sort_dong', 'sort_ho']).reset_index(drop=True)
        else:
            df = df.assign(
                sort_ho=pd.to_numeric(df['구분/호수'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            ).sort_values(by=['sort_ho']).drop(columns=['sort_ho']).reset_index(drop=True)
        df.index = df.index + 1
    return df

@st.cache_data(show_spinner=False)
def parse_excel_universal_sorted(file_bytes, filename):
    wb = load_any_excel_to_openpyxl(file_bytes, filename)
    EXCLUDE_KEYWORDS = {'합계', '소계', '합  계', '전월', '당월', '구분', '호', '동', '청구월', 'None', '', '평균', '차액', 'NO', '시작지침', '종료지침', '동계', '하계', '난방', '온수'}

    if not wb or not wb.sheetnames:
        return pd.DataFrame()

    scored_sheets = []
    for s_name in wb.sheetnames:
        m = re.search(r'(\d{1,2})\s*월', s_name)
        y_m = re.search(r'(20\d{2})', s_name)
        if m:
            month_num = int(m.group(1))
            year_val = int(y_m.group(1)) if y_m else 2026
            scored_sheets.append((year_val * 12 + month_num, s_name))
        else:
            scored_sheets.append((0, s_name))

    scored_sheets.sort(key=lambda x: x[0], reverse=True)
    best_sheet = wb[scored_sheets[0][1]]
    
    if scored_sheets[0][0] == 0:
        max_score = -1
        for sheetname in wb.sheetnames:
            s = wb[sheetname]
            score = s.max_row * s.max_column
            if any(k in sheetname for k in ['전기', '당월', '금월', '검침', '관리비', '전체']) or len(wb.sheetnames) == 1:
                score *= 2
            if score > max_score:
                max_score = score
                best_sheet = s

    ws = best_sheet
    max_r = ws.max_row
    max_c = ws.max_column
    if max_r < 1 or max_c < 1:
        return pd.DataFrame()

    sheet_records = []

    # ==========================================
    # [1] 연동드림아이 전용 파서
    # ==========================================
    is_yeonmok_format = False
    for r in range(1, min(5, max_r + 1)):
        row_str = " ".join([str(ws.cell(r, c).value or '') for c in range(1, min(5, max_c + 1))])
        if '검침표' in row_str or '연동드림아이' in row_str:
            is_yeonmok_format = True
            break

    if is_yeonmok_format or ('검침표' in filename or '연동드림아이' in filename):
        for r in range(1, max_r + 1):
            ho_val = ws.cell(r, 1).value
            use_val = ws.cell(r, 4).value
            if ho_val is not None and use_val is not None:
                str_ho = str(ho_val).strip()
                if str_ho in EXCLUDE_KEYWORDS or any(k in str_ho for k in ['합계', '소계', '호수', '검침표']):
                    continue
                use_num = clean_num(use_val)
                if use_num is not None:
                    sheet_records.append({
                        '동': '',
                        '구분/호수': str_ho.replace('호', ''),
                        '사용량(kWh)': use_num
                    })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records))

    # ==========================================
    # [2] 한일베라체 전용 파서
    # ==========================================
    if '한일베라체' in filename or '한일' in filename:
        for r in range(5, max_r + 1):
            dong_val = ws.cell(r, 2).value
            ho_val = ws.cell(r, 3).value
            use_val = ws.cell(r, 7).value
            if use_val is None and max_c >= 8:
                use_val = ws.cell(r, 8).value

            if dong_val is not None and ho_val is not None and use_val is not None:
                str_dong = str(int(float(dong_val))) if str(dong_val).replace('.','',1).isdigit() else str(dong_val).strip()
                str_ho = str(int(float(ho_val))) if str(ho_val).replace('.','',1).isdigit() else str(ho_val).strip()
                if str_ho in EXCLUDE_KEYWORDS or any(k in str_ho for k in ['동계', '하계', '합계', '소계', '난방', '온수']):
                    continue
                use_num = clean_num(use_val)
                if use_num is not None:
                    sheet_records.append({
                        '동': str_dong.replace('동', ''),
                        '구분/호수': str_ho.replace('호', ''),
                        '사용량(kWh)': use_num
                    })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records))

    # ==========================================
    # [3] 기타 나머지 파일들: 범용 키워드 탐색 + 세로형/분할형(힐튼 등) 구조 탐색
    # ==========================================
    header_texts = {}
    for c in range(1, max_c + 1):
        col_text_list = []
        for r in range(1, min(15, max_r + 1)):
            val = ws.cell(r, c).value
            if val is not None:
                s_val = str(val).strip()
                if s_val and s_val not in col_text_list:
                    col_text_list.append(s_val)
        header_texts[c] = " ".join(col_text_list)

    dong_cols, ho_cols, use_cols = [], [], []
    for c, h_text in header_texts.items():
        h_clean = h_text.replace(" ", "")
        if '동' in h_clean and not any(k in h_clean for k in ['동계', '하계', '부하', '사용량', '지역', '지침']):
            dong_cols.append(c)
        if any(k in h_clean for k in ['호실', '호수', '세대', '구분', '호(구분)']) and '사용량' not in h_clean:
            ho_cols.append(c)
        if any(k in h_clean for k in ['사용량', '금월사용', '계기사용량', '검침합계', '당월계', '부하사용량', '합계사용량', '전기사용량']) and '지침' not in h_clean:
            use_cols.append(c)

    # A. 세로형 / 분할형 표 구조 탐지 (예: 여러 개의 호실/사용량 쌍이 가로 블록 단위로 나뉘어 있는 경우)
    if len(ho_cols) >= 2 or len(use_cols) >= 2:
        # 각 호실 컬럼과 대응되는 사용량 컬럼을 짝지어 탐색
        for h_c in ho_cols:
            # 해당 호실 컬럼과 가장 가까운 우측의 사용량 컬럼 매칭
            valid_u_cols = [u for u in use_cols if u >= h_c and u <= h_c + 3]
            u_c = valid_u_cols[0] if valid_u_cols else (use_cols[0] if use_cols else h_c + 1)
            
            for r in range(1, max_r + 1):
                raw_ho = ws.cell(r, h_c).value
                raw_use = ws.cell(r, u_c).value
                if raw_ho is not None and raw_use is not None:
                    str_ho = str(raw_ho).strip()
                    if str_ho in EXCLUDE_KEYWORDS or any(k in str_ho for k in ['합계', '소계', '구분', '호실', '호수', '동계', '하계', '난방', '온수']):
                        continue
                    use_num = clean_num(raw_use)
                    if use_num is not None:
                        sheet_records.append({
                            '동': '',
                            '구분/호수': str_ho.replace('.0', '').replace('호', ''),
                            '사용량(kWh)': use_num
                        })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records))

    # B. 일반적인 단일 표 구조 탐색
    has_dong_col = len(dong_cols) > 0
    d_col = dong_cols[0] if has_dong_col else None
    h_col = ho_cols[0] if ho_cols else (2 if has_dong_col else 1)
    u_col = use_cols[0] if use_cols else max_c

    current_dong = ""
    for r in range(1, max_r + 1):
        raw_dong = ws.cell(r, d_col).value if (has_dong_col and d_col) else None
        raw_ho = ws.cell(r, h_col).value if h_col else None
        raw_use = ws.cell(r, u_col).value if u_col else None

        if has_dong_col and d_col and raw_dong is not None and str(raw_dong).strip() != '':
            str_d = str(raw_dong).strip()
            if not any(k in str_d for k in ['동계', '하계', '합계', '소계', 'EV', '난방', '온수', '기타', '동', '구분', '호실']):
                clean_d_num = re.sub(r'[^0-9]', '', str_d)
                if clean_d_num != '' or '동' in str_d:
                    current_dong = str_d
            elif '동' in str_d:
                current_dong = str_d

        if raw_ho is not None and raw_use is not None:
            str_ho = str(raw_ho).strip()
            if str_ho in EXCLUDE_KEYWORDS or any(k in str_ho for k in ['동계', '하계', 'EV', '합계', '소계', '전기검침', '호 수', '난방', '온수', '구분']):
                continue

            use_num = clean_num(raw_use)
            if use_num is not None:
                target_dong = str(current_dong).replace('동', '').strip() if has_dong_col else ''
                if has_dong_col and (not target_dong or any(k in target_dong for k in ['계', '동계', '하계', '난방', '온수'])):
                    continue

                sheet_records.append({
                    '동': target_dong,
                    '구분/호수': str_ho,
                    '사용량(kWh)': use_num
                })

    return finalize_dataframe(pd.DataFrame(sheet_records))

uploaded_files = st.file_uploader("엑셀 파일(.xlsx, .xls)들을 선택하세요 (다중 선택 가능)", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    results = {}
    selected_sheets_info = {}

    with st.spinner(f"총 {len(uploaded_files)}개 파일 분석 중..."):
        for file in uploaded_files:
            file_bytes = file.read()
            wb_preview = load_any_excel_to_openpyxl(file_bytes, file.name)
            if wb_preview and wb_preview.sheetnames:
                scored_sheets = []
                for s_name in wb_preview.sheetnames:
                    m = re.search(r'(\d{1,2})\s*월', s_name)
                    y_m = re.search(r'(20\d{2})', s_name)
                    if m:
                        scored_sheets.append((int(y_m.group(1) if y_m else 2026) * 12 + int(m.group(1)), s_name))
                    else:
                        scored_sheets.append((0, s_name))
                scored_sheets.sort(key=lambda x: x[0], reverse=True)
                selected_sheets_info[file.name] = scored_sheets[0][1]

            df_parsed = parse_excel_universal_sorted(file_bytes, file.name)

            if not df_parsed.empty:
                df_under_50 = df_parsed[df_parsed['사용량(kWh)'] < 50].copy()
                df_under_50.reset_index(drop=True, inplace=True)
                df_under_50.index = df_under_50.index + 1
                results[file.name] = df_under_50
            else:
                results[file.name] = pd.DataFrame()

    st.subheader("🎯 50 미만 세대(구분) 추출 결과")
    tab_names = [f"📄 {name}" for name in results.keys()]
    tabs = st.tabs(tab_names)

    for i, (file_name, df_under_50) in enumerate(results.items()):
        with tabs[i]:
            detected_s = selected_sheets_info.get(file_name, '기본 시트')
            st.info(f"📌 **자동 선택된 분석 시트 (최근 월)**: `{detected_s}`")
            st.metric("⚠️ 50 미만 세대수", f"{len(df_under_50)} 건")
            if not df_under_50.empty:
                st.dataframe(df_under_50, use_container_width=True)

                df_excel = df_under_50.copy()
                df_excel.insert(0, 'No', range(1, len(df_excel) + 1))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='50미만세대')
                st.download_button(
                    label=f"📥 {file_name} 결과 엑셀 다운로드",
                    data=output.getvalue(),
                    file_name=f"50미만세대_{file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.success("해당 파일에는 50 미만 세대가 없습니다.")
