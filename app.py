import io
import re
import hashlib
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
import xlrd
from openpyxl import Workbook
import streamlit as st

# --- 페이지 기본 설정 ---
st.set_page_config(
    page_title="전기사용량 50 미만 세대 추출기",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 외부 CSS 실시간 로드 함수 ---
def local_css(file_name):
    try:
        with open(file_name, encoding='utf-8') as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"CSS 파일을 찾을 수 없습니다: {file_name}")

# 스타일 적용
local_css("assets/style.css")


# --- 데이터 전처리 및 파싱 함수들 ---
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


def is_valid_ho(ho_str):
    if ho_str is None:
        return False
    s = str(ho_str).strip().replace('.0', '')
    
    EXCLUDE_EXACT = {'합계', '소계', '합  계', '전월', '당월', '구분', '호', '동', '청구월', 'None', '', '평균', '차액', 'NO', '시작지침', '종료지침', '동계', '하계', '난방', '온수', '총계', '계', '호실', '사용량'}
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
def parse_excel_cached(file_hash, file_bytes, filename):
    wb = load_any_excel_to_openpyxl_cached(file_hash, file_bytes, filename)
    if not wb or not wb.sheetnames:
        return pd.DataFrame(), '기본 시트'

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
    best_sheet_name = scored_sheets[0][1]
    best_sheet = wb[best_sheet_name]
    
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
                best_sheet_name = sheetname

    ws = best_sheet
    max_r = ws.max_row
    max_c = ws.max_column
    if max_r < 1 or max_c < 1:
        return pd.DataFrame(), best_sheet_name

    sheet_records = []

    # [1] 힐튼 전용 파서 (병렬 블록 구조 완벽 대응)
    if '힐튼' in filename:
        for c in range(1, max_c + 1, 4):
            for r in range(3, max_r + 1):
                ho_val = ws.cell(row=r, column=c).value
                use_val = ws.cell(row=r, column=c + 3).value
                if ho_val is not None and use_val is not None:
                    str_ho = str(ho_val).strip().replace('호', '').replace('.0', '')
                    if is_valid_ho(str_ho):
                        use_num = clean_num(use_val)
                        if use_num is not None:
                            sheet_records.append({
                                '동': '',
                                '구분/호수': str_ho,
                                '사용량(kWh)': use_num
                            })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

    # [2] 연동드림아이 전용 파서
    is_yeonmok_format = False
    for r in range(1, min(5, max_r + 1)):
        row_str = " ".join([str(ws.cell(r, c).value or '') for c in range(1, min(5, max_c + 1))])
        if '검침표' in row_str or '연동드림아이' in row_str:
            is_yeonmok_format = True
            break

    if is_yeonmok_format or ('검침표' in filename and '힐튼' not in filename):
        for r in range(1, max_r + 1):
            ho_val = ws.cell(r, 1).value
            use_val = ws.cell(r, 4).value
            if ho_val is not None and use_val is not None:
                str_ho = str(ho_val).strip()
                if not is_valid_ho(str_ho):
                    continue
                use_num = clean_num(use_val)
                if use_num is not None:
                    sheet_records.append({
                        '동': '',
                        '구분/호수': str_ho.replace('호', ''),
                        '사용량(kWh)': use_num
                    })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

    # [3] 한일베라체 전용 파서
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
                if not is_valid_ho(str_ho):
                    continue
                use_num = clean_num(use_val)
                if use_num is not None:
                    sheet_records.append({
                        '동': str_dong.replace('동', ''),
                        '구분/호수': str_ho.replace('호', ''),
                        '사용량(kWh)': use_num
                    })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

    # [4] 벨라시티 전용 파서
    if '벨라시티' in filename:
        current_dong = ""
        for r in range(1, max_r + 1):
            val_d = ws.cell(row=r, column=4).value
            if val_d is None:
                val_d = ws.cell(row=r, column=3).value
            
            val_h = ws.cell(row=r, column=5).value
            if val_h is None:
                val_h = ws.cell(row=r, column=4).value
                
            val_u = ws.cell(row=r, column=9).value
            if val_u is None and max_c >= 9:
                for c_alt in range(6, max_c + 1):
                    cand = ws.cell(row=r, column=c_alt).value
                    if clean_num(cand) is not None:
                        val_u = cand
                        break

            if val_d is not None:
                s_d = str(val_d).strip().replace('.0', '').replace('동', '')
                if s_d.isdigit() and int(s_d) < 100:
                    current_dong = s_d

            if val_h is not None and val_u is not None:
                str_h = str(val_h).strip().replace('.0', '').replace('호', '')
                if is_valid_ho(str_h):
                    u_num = clean_num(val_u)
                    if u_num is not None:
                        sheet_records.append({
                            '동': current_dong,
                            '구분/호수': str_h,
                            '사용량(kWh)': u_num
                        })
        if sheet_records:
            return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

    # [5] 좌우 병렬 반복형 구조 파서
    for c in range(1, max_c + 1):
        h_val = ws.cell(row=1, column=c).value
        if h_val is not None and ('동' in str(h_val)):
            d_col = c
            h_col = c + 1
            u_col = c + 2
            
            if u_col <= max_c:
                h_name = str(ws.cell(row=1, column=h_col).value or '')
                u_name = str(ws.cell(row=1, column=u_col).value or '')
                
                if '호' in h_name and '사용량' in u_name:
                    for r in range(2, max_r + 1):
                        d_val = ws.cell(row=r, column=d_col).value
                        h_val_cell = ws.cell(row=r, column=h_col).value
                        u_val_cell = ws.cell(row=r, column=u_col).value
                        
                        if d_val is not None and h_val_cell is not None and u_val_cell is not None:
                            str_d = str(d_val).strip().replace('.0', '').replace('동', '')
                            str_h = str(h_val_cell).strip().replace('.0', '').replace('호', '')
                            
                            if is_valid_ho(str_h):
                                u_num = clean_num(u_val_cell)
                                if u_num is not None:
                                    sheet_records.append({
                                        '동': str_d,
                                        '구분/호수': str_h,
                                        '사용량(kWh)': u_num
                                    })
    if sheet_records:
        return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name

    # [6] 일반 단일 표 구조 파서
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
        if '동' in h_clean and not any(k in h_clean for k in ['동계', '하계', '부하', '사용량', '지역', '지침', '전월']):
            dong_cols.append(c)
        if any(k in h_clean for k in ['호실', '호수', '세대', '구분', '호(구분)']) and not any(k in h_clean for k in ['사용량', '전월', '지침']):
            ho_cols.append(c)
        if any(k in h_clean for k in ['사용량', '금월사용', '계기사용량', '검침합계', '당월계', '부하사용량', '합계사용량', '전기사용량']) and not any(k in h_clean for k in ['전월사용', '전월지침', '시작지침']):
            use_cols.append(c)

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
            if not is_valid_ho(str_ho):
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

    return finalize_dataframe(pd.DataFrame(sheet_records)), best_sheet_name


# --- [UI 메인 레이아웃 구성] ---

# 사이드바 영역
with st.sidebar:
    st.markdown("### 📂 파일 업로드 센터")
    uploaded_files = st.file_uploader(
        "엑셀 파일(.xlsx, .xls) 다중 선택 가능", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.markdown("### 💡 지원 양식 안내")
    st.info(
        "• 한일베라체 / 연동드림아이  \n"
        "• 벨라시티 / 힐튼 검침표  \n"
        "• 좌우 병렬형 및 일반 단일 표"
    )

# 메인 헤더 영역 (CSS 클래스 활용)
st.markdown('<p class="main-title">⚡ 전기사용량 50 미만 세대 자동 추출 시스템</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">관리비 및 검침표 엑셀 파일을 업로드하면 50kWh 미만 세대를 자동으로 분석하고 추출해 드립니다.</p>', unsafe_allow_html=True)

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
                df_under_50 = df_parsed[df_parsed['사용량(kWh)'] < 50].copy()
                df_under_50.reset_index(drop=True, inplace=True)
                df_under_50.index = df_under_50.index + 1
                results[file.name] = df_under_50
            else:
                results[file.name] = pd.DataFrame()

    st.markdown("### 🎯 분석 결과 대시보드")
    tab_names = [f"📄 {name}" for name in results.keys()]
    tabs = st.tabs(tab_names)

    for i, (file_name, df_under_50) in enumerate(results.items()):
        with tabs[i]:
            detected_s = selected_sheets_info.get(file_name, '기본 시트')
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.success(f"📌 **감지된 분석 시트**: `{detected_s}`")
            with col2:
                st.metric(label="⚠️ 50 미만 세대수", value=f"{len(df_under_50)} 건")
            
            st.markdown("---")
            
            if not df_under_50.empty:
                st.dataframe(df_under_50, use_container_width=True, height=400)

                df_excel = df_under_50.copy()
                df_excel.insert(0, 'No', range(1, len(df_excel) + 1))

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel.to_excel(writer, index=False, sheet_name='50미만세대')
                
                st.download_button(
                    label=f"📥 [{file_name}] 결과 엑셀 다운로드받기",
                    data=output.getvalue(),
                    file_name=f"50미만세대_{file_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.success("🎉 해당 파일에는 사용량 50 미만인 세대가 존재하지 않습니다!")
else:
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; padding: 40px; background-color: #FFFFFF; border-radius: 12px; border: 2px dashed #CBD5E1;'>
            <h3>👈 좌측 사이드바에서 분석할 엑셀 파일을 업로드해주세요.</h3>
            <p style='color: #64748B;'>여러 개의 파일을 동시에 드래그 & 드롭하여 간편하게 비교·분석할 수 있습니다.</p>
        </div>
        """, 
        unsafe_allow_html=True
    )
