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
st.markdown("한일베라체, 연동드림아이, 힐튼 등 다양한 아파트 전기 검침표 양식을 완벽 지원합니다.")

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
    """세대 호수로 볼 수 있는 값인지 보수적으로 검증합니다."""
    if ho_str is None:
        return False

    s = str(ho_str).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = s.replace("호", "").strip()

    exclude_exact = {
        "", "None", "nan", "합계", "총계", "소계", "평균", "통계",
        "전월", "당월", "구분", "호", "동", "청구월", "차액", "NO",
        "시작지침", "종료지침", "동계", "하계", "난방", "온수", "계", "호실"
    }
    if s in exclude_exact:
        return False

    # 숫자로만 된 호수만 허용합니다. 숫자가 섞인 통계/지침 텍스트를 호수로 오인하지 않습니다.
    if not re.fullmatch(r"\d+", s):
        return False

    value = int(s)
    return 0 < value <= 5000


def is_summary_row(row_values):
    """행 전체를 확인해 통계/합계/설명 행을 제외합니다."""
    exclude_keywords = (
        "합계", "총계", "소계", "평균", "통계", "동계", "하계",
        "난방", "온수", "전체합계", "세대합계", "페이지", "총사용량"
    )
    for value in row_values:
        if value is None:
            continue
        text = re.sub(r"\s+", "", str(value)).strip()
        if any(keyword in text for keyword in exclude_keywords):
            return True
    return False


def normalize_identifier(value):
    """엑셀 숫자 셀의 1.0을 1로 정규화합니다."""
    if value is None:
        return ""
    text = str(value).strip()
    if re.fullmatch(r"[-+]?\d+\.0", text):
        return text[:-2]
    return text.replace("동", "").replace("호", "").strip()

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
            row_values = [ws.cell(r, c).value for c in range(1, max_c + 1)]
            if is_summary_row(row_values):
                continue
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
            return finalize_dataframe(pd.DataFrame(sheet_records))

    # ==========================================
    # [2] 한일베라체 전용 파서
    # ==========================================
    if '한일베라체' in filename or '한일' in filename:
        for r in range(5, max_r + 1):
            row_values = [ws.cell(r, c).value for c in range(1, max_c + 1)]
            if is_summary_row(row_values):
                continue
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
            return finalize_dataframe(pd.DataFrame(sheet_records))

    # ==========================================
    # [3] 힐튼 등 가로 블록 반복형 구조 파서 (4열 단위: 구분, 전월, 당월, 사용량)
    # ==========================================
    for header_row in range(1, max_r + 1):
        for c_start in range(1, max_c + 1, 4):
            val = ws.cell(row=header_row, column=c_start).value
            if val is not None and str(val).strip() == '구분':
                # 데이터는 헤더 행 바로 다음 또는 다다음 행부터 시작
                for r in range(header_row + 1, max_r + 1):
                    row_values = [ws.cell(r, c).value for c in range(1, max_c + 1)]
                    if is_summary_row(row_values):
                        continue
                    ho_cell = ws.cell(row=r, column=c_start).value
                    use_cell = ws.cell(row=r, column=c_start + 3).value # 4번째 열이 사용량
                    if ho_cell is not None and use_cell is not None:
                        str_ho = str(ho_cell).strip().replace('.0', '').replace('호', '')
                        if not is_valid_ho(str_ho):
                            continue
                        use_num = clean_num(use_cell)
                        if use_num is not None:
                            sheet_records.append({
                                '동': '',
                                '구분/호수': str_ho,
                                '사용량(kWh)': use_num
                            })
    if sheet_records:
        return finalize_dataframe(pd.DataFrame(sheet_records))

    # ==========================================
    # [4] 일반 단일 표 파서: 실제 헤더 행을 찾아 열을 명시적으로 매핑
    # ==========================================
    # 열 전체의 텍스트를 합쳐 추정하지 않고, 한 행에 있는 헤더명을 기준으로 탐색합니다.
    header_candidates = []
    for r in range(1, min(30, max_r) + 1):
        values = [ws.cell(r, c).value for c in range(1, max_c + 1)]
        texts = [re.sub(r"\s+", "", str(v)).lower() if v is not None else "" for v in values]

        dong_idx = next((i for i, t in enumerate(texts) if t in {"동", "동번호", "동명"}), None)
        ho_idx = next((i for i, t in enumerate(texts) if t in {"호", "호수", "호실", "세대", "세대번호", "구분/호수", "호(구분)"}), None)
        use_idx = next((i for i, t in enumerate(texts) if any(k in t for k in (
            "사용량", "금월사용", "계기사용량", "검침합계", "당월계", "부하사용량", "합계사용량", "전기사용량"
        )) and not any(k in t for k in ("전월사용", "전월지침", "시작지침", "종료지침"))), None)

        # 호수와 사용량 헤더가 같은 행에서 확인되어야 신뢰할 수 있는 표 헤더로 인정합니다.
        if ho_idx is not None and use_idx is not None:
            header_candidates.append((r, dong_idx, ho_idx, use_idx))

    if not header_candidates:
        # 헤더를 확실히 찾지 못하면 임의로 마지막 열을 사용하지 않습니다.
        return finalize_dataframe(pd.DataFrame(columns=["동", "구분/호수", "사용량(kWh)"]))

    # 가장 위쪽의 유효한 헤더를 사용합니다.
    header_row, dong_idx, ho_idx, use_idx = header_candidates[0]
    d_col = dong_idx + 1 if dong_idx is not None else None
    h_col = ho_idx + 1
    u_col = use_idx + 1
    has_dong_col = d_col is not None

    current_dong = ""
    for r in range(header_row + 1, max_r + 1):
        row_values = [ws.cell(r, c).value for c in range(1, max_c + 1)]

        # 표 아래쪽의 합계/통계/페이지 표기 등은 행 전체를 검사해 제외합니다.
        if is_summary_row(row_values):
            continue

        raw_dong = ws.cell(r, d_col).value if d_col else None
        raw_ho = ws.cell(r, h_col).value
        raw_use = ws.cell(r, u_col).value

        if raw_ho is None or raw_use is None:
            continue

        str_ho = normalize_identifier(raw_ho)
        if not is_valid_ho(str_ho):
            continue

        use_num = clean_num(raw_use)
        if use_num is None or use_num < 0:
            continue

        # 동 값이 있는 행에서만 현재 동을 갱신합니다. '동계' 같은 문자는 행 검사에서 제외됩니다.
        if has_dong_col and raw_dong is not None and str(raw_dong).strip() != "":
            dong_text = normalize_identifier(raw_dong)
            if re.fullmatch(r"\d+", dong_text):
                current_dong = dong_text
            else:
                # 숫자 동 번호가 아닌 텍스트를 만나면 이전 동 번호를 이어 쓰지 않습니다.
                current_dong = ""
                continue

        target_dong = current_dong if has_dong_col else ""
        if has_dong_col and not target_dong:
            continue

        sheet_records.append({
            "동": target_dong,
            "구분/호수": str_ho,
            "사용량(kWh)": use_num
        })

    # 동일 동/호가 중복 추출되면 첫 결과만 남기되, 다른 사용량이 있는 중복은 UI에서 확인할 수 있게 처리합니다.
    if sheet_records:
        parsed = pd.DataFrame(sheet_records)
        parsed = parsed.drop_duplicates(subset=["동", "구분/호수"], keep="first")
        return finalize_dataframe(parsed)

    return finalize_dataframe(pd.DataFrame(columns=["동", "구분/호수", "사용량(kWh)"]))

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
                st.info("50kWh 미만 세대가 없거나, 검침표의 헤더/열 구조를 자동으로 확인하지 못했습니다. 원본 시트의 동·호수·사용량 헤더를 확인해 주세요.")
