import io
import re
import hashlib
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
import xlrd
from openpyxl import Workbook
import streamlit as st
import os
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

st.set_page_config(page_title="전기사용량 50 미만 세대 추출기", layout="wide")

def local_css(file_name):
    if os.path.exists(file_name):
        with open(file_name, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

local_css("assets/style.css")

with st.sidebar:
    st.markdown("### 📂 파일 업로드")
    uploaded_files = st.file_uploader(
        "엑셀 파일(.xlsx, .xls) 선택", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    st.markdown("---")
    st.markdown("### 📌 시스템 안내")
    st.markdown("본 프로그램은 공동주택의 전기사용량 검침표 엑셀 파일을 분석하여 **50 kWh 미만 세대**를 자동으로 추출합니다.")
    st.markdown("---")
    st.markdown("🛠 **지원 양식**\n- 제이하임\n- 한일베라체\n- 벨라시티\n- 연동드림아이\n- 힐튼 / 엠제이벤처\n- **한셀 지침 자동 계산 및 정밀 보정 양식**")

st.markdown('<p class="main-title" style="font-size: 40px;">⚡ 전기사용량 50 미만 세대 자동 검색 시스템</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">제이하임, 한일베라체, 벨라시티, 연동드림아이, 힐튼, 엠제이벤처 및 한셀 파일 양식을 고속으로 분석합니다.</p>', unsafe_allow_html=True)

def clean_num(val):
    if val is None or pd.isna(val):
        return None
    val_str = str(val).strip()
    if val_str == '' or val_str == '-' or val_str.upper() == 'NONE' or val_str.upper() == 'NAN':
        return None
    
    val_str = val_str.replace(",", "").replace("원", "").replace("kWh", "").replace("KW", "").strip()
    
    match = re.search(r"[-+]?\d*\.?\d+", val_str)
    if match:
        try:
            return float(match.group())
        except ValueError:
            return None
    return None

def is_valid_ho(ho_str):
    if ho_str is None or pd.isna(ho_str):
        return False
    s = str(ho_str).strip().replace('.0', '')

    EXCLUDE_EXACT = {'합계', '소계', '합  계', '전월', '당월', '구분', '호', '동', '청구월', 'None', 'nan', '', '평균', '차액', 'NO', '시작지침', '종료지침', '동계', '하계', '난방', '온수', '총계', '계', '호실', '사용량'}
    if s in EXCLUDE_EXACT:
        return False

    clean_digits = re.sub(r'[^0-9]', '', s)
    if not clean_digits:
        return False

    val_num = int(clean_digits)
    if val_num > 5000 or val_num == 0:
        return False

    return True

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

@st.cache_data(show_spinner=False)
def load_any_excel_to_openpyxl_cached(file_hash, file_bytes, filename):
    pyxl_wb = Workbook()
    pyxl_wb.remove(pyxl_wb.active)

    # 1. XML 기반 포맷
    if file_bytes.startswith(b'<?xml') or b'urn:schemas-microsoft-com:office:spreadsheet' in file_bytes:
        try:
            root = ET.fromstring(file_bytes)
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

    # 2. 구형 엑셀 (.xls)
    if filename.lower().endswith('.xls'):
        try:
            xls_wb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=False)
            for sheet_name in xls_wb.sheet_names():
                xls_sheet = xls_wb.sheet_by_name(sheet_name)
                target_ws = pyxl_wb.create_sheet(title=sheet_name)
                for r in range(xls_sheet.nrows):
                    target_ws.append(xls_sheet.row_values(r))
            if pyxl_wb.sheetnames:
                return pyxl_wb
        except Exception:
            pass

    # 3. Pandas를 활용한 한셀 수식/스타일 안전 로딩
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            target_ws = pyxl_wb.create_sheet(title=sheet_name)
            for _, row in df.iterrows():
                target_ws.append(row.tolist())
        if pyxl_wb.sheetnames:
            return pyxl_wb
    except Exception:
        pass

    # 4. 표준 openpyxl Fallback
    if not pyxl_wb.sheetnames:
        pyxl_wb = Workbook()
        pyxl_wb.remove(pyxl_wb.active)
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                target_ws = pyxl_wb.create_sheet(title=sheetname)
                for r in range(1, ws.max_row + 1):
                    row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                    target_ws.append(row_vals)
        except Exception:
            try:
                wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    target_ws = pyxl_wb.create_sheet(title=sheetname)
                    for r in range(1, ws.max_row + 1):
                        row_vals = []
                        for c in range(1, ws.max_column + 1):
                            val = ws.cell(r, c).value
                            row_vals.append(None if str(val).startswith('=') else val)
                        target_ws.append(row_vals)
            except Exception:
                pass

    if not pyxl_wb.sheetnames:
        fallback_ws = pyxl_wb.create_sheet(title="Sheet1")
        fallback_ws.append(["파일 읽기 실패"])

    return pyxl_wb

def finalize_dataframe(df):
    if not df.empty:
        df = df.drop_duplicates().reset_index(drop=True)
        
        if '동' in df.columns and df['동'].astype(str).str.strip().ne('').any() and not df['동'].isna().all():
            df = df.assign(
                sort_dong=pd.to_numeric(df['동'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0),
                sort_ho=pd.to_numeric(df['호수'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            ).sort_values(by=['sort_dong', 'sort_ho']).drop(columns=['sort_dong', 'sort_ho']).reset_index(drop=True)
        else:
            df = df.assign(
                sort_ho=pd.to_numeric(df['호수'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            ).sort_values(by=['sort_ho']).drop(columns=['sort_ho']).reset_index(drop=True)
            
        df.index = df.index + 1
    return df

@st.cache_data(show_spinner=False)
def parse_excel_cached(file_hash, file_bytes, filename):
    wb = load_any_excel_to_openpyxl_cached(file_hash, file_bytes, filename)
    if not wb or not wb.sheetnames:
        return pd.DataFrame(), '기본 시트'

    scored_sheets = []
    for s_name in wb.sheetnames:
        m_month = re.search(r'(\d{1,2})\s*월', s_name)
        m_year = re.search(r'(\d{2,4})\s*[\.\s\-_년]*\s*(\d{1,2})\s*월', s_name)
        
        if m_year:
            y_val = int(m_year.group(1))
            if y_val < 100:
                y_val += 2000
            m_val = int(m_year.group(2))
            score = y_val * 12 + m_val
        elif m_month:
            score = 2024 * 12 + int(m_month.group(1))
        else:
            score = 0
        scored_sheets.append((score, s_name))

    scored_sheets.sort(key=lambda x: x[0], reverse=True)
    best_sheet_name = scored_sheets[0][1]
    best_sheet = wb[best_sheet_name]

    if scored_sheets[0][0] == 0:
        max_score = -1
        for sheetname in wb.sheetnames:
            s = wb[sheetname]
            valid_cells = sum(1 for row in s.iter_rows(values_only=True) if any(x is not None for x in row))
            score = valid_cells
            if any(k in sheetname for k in ['전기', '당월', '금월', '검침', '관리비', '전체', 'sheet1', 'Sheet1']) or len(wb.sheetnames) == 1:
                score *= 2
            if score > max_score:
                max_score = score
                best_sheet = s
                best_sheet_name = sheetname

    ws = best_sheet
    max_r = ws.max_row
    max_c = ws.max_column
    if max_r < 1 or max_c < 1:
        return pd.DataFrame(), best_sheet_name

    sheet_records = []

    # === 헤더 텍스트 매핑 분석 (최대 20행까지 정밀 스캔) ===
    header_texts = {}
    for c in range(1, max_c + 1):
        col_text_list = []
        for r in range(1, min(20, max_r + 1)):
            val = ws.cell(r, c).value
            if val is not None:
                s_val = str(val).strip()
                if s_val and s_val not in col_text_list:
                    col_text_list.append(s_val)
        header_texts[c] = " ".join(col_text_list)

    dong_cols, ho_cols, use_cols, prev_cols, curr_cols = [], [], [], [], []
    for c, h_text in header_texts.items():
        h_clean = h_text.replace(" ", "")
        if '동' in h_clean and not any(k in h_clean for k in ['동계', '하계', '부하', '사용량', '지역', '지침', '전월']):
            dong_cols.append(c)
        if any(k in h_clean for k in ['호실', '호수', '세대', '구분', '호(구분)', '호']) and not any(k in h_clean for k in ['사용량', '전월', '지침', '단가']):
            ho_cols.append(c)
        if any(k in h_clean for k in ['사용량', '금월사용', '계기사용량', '검침합계', '당월계', '부하사용량', '합계사용량', '전기사용량']):
            use_cols.append(c)
        if any(k in h_clean for k in ['전월지침', '기초지침', '전월', '시작지침']):
            prev_cols.append(c)
        if any(k in h_clean for k in ['당월지침', '현재지침', '당월', '종료지침', '금월지침']):
            curr_cols.append(c)

    has_dong_col = len(dong_cols) > 0
    d_col = dong_cols[0] if has_dong_col else None
    h_col = ho_cols[0] if ho_cols else (2 if has_dong_col else 1)
    
    # 사용량 열 확정 (만약 키워드로 못 찾았다면 전월지침/당월지침 차이 계산 또는 숫자 분포 분석)
    u_col = use_cols[0] if use_cols else None
    p_col = prev_cols[0] if prev_cols else None
    c_col = curr_cols[0] if curr_cols else None

    # 만약 사용량 열이 명확하지 않고 당월/전월 지침 열이 있다면 직접 지침 간 차이를 계산하기 위한 후보 설정
    if u_col is None and p_col and c_col:
        # 지침 간의 차이를 사용량으로 간주
        pass
    elif u_col is None:
        # 숫자가 가장 많이 들어있는 열 중 호수 열 다음의 적절한 열 탐색 (사용량 값은 보통 0~1000 사이)
        best_u_col = None
        min_avg_val = float('inf')
        for c in range(h_col + 1, max_c + 1):
            vals = []
            for r in range(1, min(50, max_r + 1)):
                v = clean_num(ws.cell(r, c).value)
                if v is not None and v < 5000: # 지침처럼 너무 큰 숫자가 아닌 것
                    vals.append(v)
            if len(vals) > 5:
                avg_v = sum(vals) / len(vals)
                if avg_v < min_avg_val: # 평균값이 가장 작은 열이 보통 '사용량' 열임 (지침은 1만 단위로 큼)
                    min_avg_val = avg_v
                    best_u_col = c
        u_col = best_u_col if best_u_col else (6 if max_c >= 6 else max_c)

    current_dong = ""
    for r in range(1, max_r + 1):
        raw_dong = ws.cell(r, d_col).value if (has_dong_col and d_col) else None
        raw_ho = ws.cell(r, h_col).value if (h_col <= max_c) else None
        
        # 사용량 값 추출 (u_col 직접 읽기 또는 당월지침 - 전월지침 계산)
        raw_use = None
        if u_col and u_col <= max_c:
            raw_use = ws.cell(r, u_col).value
            
        use_num = clean_num(raw_use)
        
        # 만약 직접 읽은 사용량이 None이거나 너무 크다면 (지침을 잘못 읽었을 경우), 당월 - 전월로 계산 시도
        if (use_num is None or use_num > 5000) and p_col and c_col and p_col <= max_c and c_col <= max_c:
            p_val = clean_num(ws.cell(r, p_col).value)
            c_val = clean_num(ws.cell(r, c_col).value)
            if p_val is not None and c_val is not None and c_val >= p_val:
                use_num = c_val - p_val

        if has_dong_col and d_col and raw_dong is not None and str(raw_dong).strip() != '':
            str_d = str(raw_dong).strip()
            if not any(k in str_d for k in ['동계', '하계', '합계', '소계', 'EV', '난방', '온수', '기타', '동', '구분', '호실']):
                clean_d_num = re.sub(r'[^0-9]', '', str_d)
                if clean_d_num != '' or '동' in str_d:
                    current_dong = str_d
            elif '동' in str_d:
                current_dong = str_d

        if raw_ho is not None and use_num is not None:
            str_ho = str(raw_ho).strip()
            if not is_valid_ho(str_ho):
                continue

            target_dong = str(current_dong).replace('동', '').strip() if has_dong_col else ''
            if has_dong_col and (not target_dong or any(k in target_dong for k in ['계', '동계', '하계', '난방', '온수'])):
                continue

            sheet_records.append({
                '동': target_dong,
                '호수': str_ho.replace('호', ''),
                '사용량(kw)': use_num
            })

    return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

if uploaded_files:
    results = {}
    selected_sheets_info = {}

    with st.spinner("🚀 고속으로 엑셀 파일들을 분석 중입니다..."):
        for file in uploaded_files:
            file_bytes = file.read()
            f_hash = get_file_hash(file_bytes)

            df_parsed, detected_sheet = parse_excel_cached(f_hash, file_bytes, file.name)
            selected_sheets_info[file.name] = detected_sheet

            if not df_parsed.empty:
                # 50 미만 세대 추출 (0 포함, 단 0보다 크고 50 미만인 세대들도 정확히 포착)
                df_under_50 = df_parsed[df_parsed['사용량(kw)'] < 50].copy()
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
            st.info(f"📌 **자동 선택된 분석 시트**: `{detected_s}`")
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
else:
    st.info("👈 왼쪽 사이드바에서 분석할 엑셀 파일들을 업로드해 주세요.")
