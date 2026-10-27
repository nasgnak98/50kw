import io
import re
import xml.etree.ElementTree as ET
from datetime import datetime

import pandas as pd
import openpyxl
import xlrd
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="전기사용량 미만 세대 추출기", layout="wide")
st.title("⚡ 전기사용량 기준 미만 세대 자동 추출 시스템")
st.markdown("다양한 아파트 전기 검침표 양식을 자동 인식합니다. 인식에 실패하면 열을 직접 선택할 수 있습니다.")

MAX_HO = 9999  # 이 값을 넘는 숫자는 지침(계기 숫자)으로 보고 호수에서 제외
EXCLUDE_EXACT = {
    '합계', '소계', '합계', '전월', '당월', '구분', '호', '동', '청구월', 'None', '', '평균',
    '차액', 'NO', 'No', '시작지침', '종료지침', '동계', '하계', '난방', '온수', '총계', '계', '호실',
}
HO_KEYS = ('구분', '호', '호수', '세대', '호실')


# ───────────────────────── 공통 유틸 ─────────────────────────
def norm_str(v):
    """셀 값을 문자열로. 101.0 -> '101'"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def clean_num(val):
    if val is None:
        return None
    s = str(val).strip().replace(',', '')
    if s in ('', '-'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def is_valid_ho(ho_str):
    s = norm_str(ho_str).replace(' ', '')
    if s in EXCLUDE_EXACT:
        return False
    digits = re.sub(r'[^0-9]', '', s)
    if not digits:
        return False
    n = int(digits)
    return 0 < n <= MAX_HO


def cell(grid, r, c):
    """1-based 좌표로 안전하게 셀 읽기"""
    if r < 1 or c < 1 or r > len(grid):
        return None
    row = grid[r - 1]
    return row[c - 1] if c <= len(row) else None


def grid_dims(grid):
    return len(grid), max((len(r) for r in grid), default=0)


def make_record(dong, ho, use):
    return {'동': str(dong).replace('동', '').strip() if dong else '',
            '구분/호수': norm_str(ho).replace('호', ''),
            '사용량(kWh)': use}


# ───────────────────────── 파일 로딩 ─────────────────────────
@st.cache_data(show_spinner=False)
def load_grids(file_bytes):
    """어떤 엑셀이든 {시트명: list[list]} 로 변환 (XML 2003 / xls / xlsx)"""
    # 1) XML Spreadsheet 2003
    if file_bytes[:200].lstrip().startswith(b'<?xml') or b'urn:schemas-microsoft-com:office:spreadsheet' in file_bytes[:2000]:
        try:
            ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
            key = '{urn:schemas-microsoft-com:office:spreadsheet}'
            root = ET.fromstring(file_bytes)
            grids = {}
            for ws in root.findall('.//ss:Worksheet', ns):
                name = ws.attrib.get(key + 'Name', 'Sheet1')
                rows = []
                table = ws.find('ss:Table', ns)
                if table is not None:
                    for row in table.findall('ss:Row', ns):
                        vals = []
                        for c in row.findall('ss:Cell', ns):
                            idx = c.attrib.get(key + 'Index')
                            if idx:
                                while len(vals) < int(idx) - 1:
                                    vals.append(None)
                            d = c.find('ss:Data', ns)
                            vals.append(d.text if d is not None else None)
                        rows.append(vals)
                grids[name] = rows
            if grids:
                return grids
        except Exception:
            pass
    # 2) xls
    try:
        wb = xlrd.open_workbook(file_contents=file_bytes)
        grids = {s.name: [s.row_values(r) for r in range(s.nrows)] for s in wb.sheets()}
        if grids:
            return grids
    except Exception:
        pass
    # 3) xlsx
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    except Exception:
        return {}


def month_score(sheet_name):
    m = re.search(r'(\d{1,2})\s*월', sheet_name)
    y = re.search(r'(20\d{2})', sheet_name)
    if not m:
        return 0
    return (int(y.group(1)) if y else datetime.now().year) * 12 + int(m.group(1))


def order_sheets(grids):
    """파싱 시도 순서: 최근 월 시트 → 그 외 큰 시트 순"""
    def size_score(name):
        r, c = grid_dims(grids[name])
        s = r * c
        if any(k in name for k in ('전기', '당월', '금월', '검침', '관리비', '전체')) or len(grids) == 1:
            s *= 2
        return s
    return sorted(grids, key=lambda n: (month_score(n), size_score(n)), reverse=True)


# ───────────────────────── 결과 정리 ─────────────────────────
def finalize(records):
    df = pd.DataFrame(records)
    if df.empty:
        return df
    df = df.drop_duplicates().reset_index(drop=True)
    df['_d'] = pd.to_numeric(df['동'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
    df['_h'] = pd.to_numeric(df['구분/호수'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
    df = df.sort_values(['_d', '_h']).drop(columns=['_d', '_h']).reset_index(drop=True)
    if not df['동'].astype(str).str.strip().ne('').any():
        pass  # 동 정보가 없으면 빈 열로 둠
    df.index = df.index + 1
    return df


# ───────────────────────── 양식별 파서 (grid 기반) ─────────────────────────
def parser_yeondong(grid, filename):
    max_r, max_c = grid_dims(grid)
    head = ' '.join(norm_str(cell(grid, r, c)) for r in range(1, min(5, max_r) + 1) for c in range(1, min(5, max_c) + 1))
    if not ('검침표' in head or '연동드림아이' in head or '검침표' in filename or '연동드림아이' in filename):
        return []
    out = []
    for r in range(1, max_r + 1):
        ho, use = cell(grid, r, 1), cell(grid, r, 4)
        if ho is not None and use is not None and is_valid_ho(ho):
            u = clean_num(use)
            if u is not None:
                out.append(make_record('', ho, u))
    return out


def parser_hanil(grid, filename):
    if '한일' not in filename:
        return []
    max_r, max_c = grid_dims(grid)
    out = []
    for r in range(5, max_r + 1):
        dong, ho, use = cell(grid, r, 2), cell(grid, r, 3), cell(grid, r, 7)
        if use is None and max_c >= 8:
            use = cell(grid, r, 8)
        if dong is None or ho is None or use is None or not is_valid_ho(ho):
            continue
        u = clean_num(use)
        if u is not None:
            out.append(make_record(norm_str(dong), ho, u))
    return out


def parser_hilton(grid, filename):
    """가로 반복 블록: [구분, 전월, 당월, 사용량] 4열 단위"""
    max_r, max_c = grid_dims(grid)
    out = []
    for hr in range(1, max_r + 1):
        for c0 in range(1, max_c + 1, 4):
            if norm_str(cell(grid, hr, c0)) == '구분':
                for r in range(hr + 1, max_r + 1):
                    ho, use = cell(grid, r, c0), cell(grid, r, c0 + 3)
                    if ho is not None and use is not None and is_valid_ho(norm_str(ho).replace('호', '')):
                        u = clean_num(use)
                        if u is not None:
                            out.append(make_record('', ho, u))
    return out


def parser_single_table(grid, filename):
    """일반 단일 표 (헤더 키워드로 열 찾기)"""
    max_r, max_c = grid_dims(grid)
    headers = {}
    for c in range(1, max_c + 1):
        seen = []
        for r in range(1, min(15, max_r) + 1):
            s = norm_str(cell(grid, r, c))
            if s and s not in seen:
                seen.append(s)
        headers[c] = ''.join(seen).replace(' ', '')
    dong_cols, ho_cols, use_cols = [], [], []
    for c, h in headers.items():
        if '동' in h and not any(k in h for k in ['동계', '하계', '부하', '사용량', '지역', '지침', '전월']):
            dong_cols.append(c)
        if any(k in h for k in ['호실', '호수', '세대', '구분', '호(구분)']) and not any(k in h for k in ['사용량', '전월', '지침']):
            ho_cols.append(c)
        if any(k in h for k in ['사용량', '금월사용', '계기사용량', '검침합계', '당월계', '부하사용량', '합계사용량', '전기사용량']) \
                and not any(k in h for k in ['전월사용', '전월지침', '시작지침']):
            use_cols.append(c)
    if not use_cols and not ho_cols:
        return []
    d_col = dong_cols[0] if dong_cols else None
    h_col = ho_cols[0] if ho_cols else (2 if d_col else 1)
    u_col = use_cols[0] if use_cols else max_c
    out, cur = [], ''
    for r in range(1, max_r + 1):
        raw_d = cell(grid, r, d_col) if d_col else None
        if d_col and raw_d not in (None, ''):
            sd = norm_str(raw_d)
            if not any(k in sd for k in ['동계', '하계', '합계', '소계', 'EV', '난방', '온수', '기타', '구분', '호실']):
                if re.sub(r'[^0-9]', '', sd) or '동' in sd:
                    cur = sd
        ho, use = cell(grid, r, h_col), cell(grid, r, u_col)
        if ho is None or use is None or not is_valid_ho(ho):
            continue
        u = clean_num(use)
        if u is None:
            continue
        dong = cur.replace('동', '').strip() if d_col else ''
        if d_col and (not dong or any(k in dong for k in ['계', '난방', '온수'])):
            continue
        out.append(make_record(dong, ho, u))
    return out


def parser_auto(grid, filename):
    """범용: 헤더 행 탐지 → 블록 분리 → 전월/당월/사용량 열 판정 (사용량 열이 없으면 당월-전월 계산)"""
    max_r, max_c = grid_dims(grid)
    keys = ['구분', '호', '전월', '당월', '사용량', '지침', '동']
    best, best_s = 0, 0
    for i, row in enumerate(grid[:20]):
        t = ''.join(norm_str(v) for v in row)
        s = sum(k in t for k in keys)
        if s > best_s:
            best, best_s = i, s
    if best_s < 2:
        return []
    hdr = grid[best] + [None] * (max_c - len(grid[best]))
    starts = [i for i, v in enumerate(hdr) if re.sub(r'\s', '', norm_str(v)) in HO_KEYS]
    if not starts:
        return []
    ends = starts[1:] + [max_c]
    out = []
    for c0, c1 in zip(starts, ends):
        names = {}
        for c in range(c0, c1):
            t = ''
            for r in range(max(0, best - 1), min(len(grid), best + 3)):
                t += norm_str(cell(grid, r + 1, c + 1)) if not (r > best and clean_num(cell(grid, r + 1, c + 1)) is not None) else ''
            names[c] = t.replace(' ', '')
        prev_c = cur_c = use_c = None
        for c, h in names.items():
            if '전월' in h and prev_c is None:
                prev_c = c
            elif '당월' in h and cur_c is None and '사용' not in h:
                cur_c = c
            elif '사용량' in h and use_c is None and '전월' not in h:
                use_c = c
        if use_c is None and (prev_c is None or cur_c is None):
            continue
        for r in range(best + 2, len(grid) + 1):
            ho = cell(grid, r, c0 + 1)
            if ho is None or not is_valid_ho(ho):
                continue
            u = clean_num(cell(grid, r, use_c + 1)) if use_c is not None else None
            if u is None and prev_c is not None and cur_c is not None:
                p, cc = clean_num(cell(grid, r, prev_c + 1)), clean_num(cell(grid, r, cur_c + 1))
                if p is not None and cc is not None:
                    u = cc - p
            if u is not None:
                out.append(make_record('', ho, u))
    return out


# 순서대로 시도하여 결과가 나오는 첫 파서 사용 (기존 검증된 파서 우선, 범용 파서가 마지막 보루)
PARSERS = [
    ('연동드림아이', parser_yeondong),
    ('한일베라체', parser_hanil),
    ('가로 반복 블록', parser_hilton),
    ('일반 단일 표', parser_single_table),
    ('범용 자동 인식', parser_auto),
]


def parse_grid(grid, filename):
    for name, fn in PARSERS:
        try:
            recs = fn(grid, filename)
        except Exception:
            recs = []
        if recs:
            return name, finalize(recs)
    return None, pd.DataFrame()


def auto_parse_file(grids, filename):
    """시트를 우선순위대로 시도 → (시트명, 파서명, df)"""
    for sheet in order_sheets(grids):
        pname, df = parse_grid(grids[sheet], filename)
        if not df.empty:
            return sheet, pname, df
    return None, None, pd.DataFrame()


# ───────────────────────── 수동 설정 ─────────────────────────
def grid_to_preview(grid, max_rows=40):
    _, mc = grid_dims(grid)
    df = pd.DataFrame([r + [None] * (mc - len(r)) for r in grid[:max_rows]])
    df.columns = [f"{get_column_letter(i + 1)}열" for i in range(mc)]
    df.index = df.index + 1
    return df.astype(str).replace({'None': '', 'nan': ''})


def parse_with_config(grid, cfg):
    _, mc = grid_dims(grid)
    offsets = range(0, mc, cfg['step']) if cfg['step'] else [0]
    out, cur = [], ''
    for off in offsets:
        for r in grid[cfg['start']:]:
            def get(c):
                if c is None:
                    return None
                i = c + off
                return r[i] if i < len(r) else None
            if cfg['dong'] is not None and get(cfg['dong']) not in (None, ''):
                cur = norm_str(get(cfg['dong'])).replace('동', '')
            ho, use = get(cfg['ho']), clean_num(get(cfg['use']))
            if ho is None or use is None or not is_valid_ho(ho):
                continue
            out.append(make_record(cur if cfg['dong'] is not None else '', ho, use))
    return finalize(out)


def manual_ui(fkey, grids, default_sheet):
    st.warning("자동 인식에 실패했습니다. 미리보기를 보고 열을 직접 선택하세요.")
    names = list(grids)
    sheet = st.selectbox("시트", names, index=names.index(default_sheet) if default_sheet in names else 0, key=f"sh_{fkey}")
    grid = grids[sheet]
    prev = grid_to_preview(grid)
    st.dataframe(prev, use_container_width=True)
    cols = list(prev.columns)
    c1, c2, c3, c4 = st.columns(4)
    ho = c1.selectbox("호수(구분) 열", cols, key=f"ho_{fkey}")
    use = c2.selectbox("당월 사용량 열", cols, index=min(3, len(cols) - 1), key=f"use_{fkey}")
    dong = c3.selectbox("동 열 (없으면 '없음')", ["없음"] + cols, key=f"dong_{fkey}")
    start = c4.number_input("데이터 시작 행", min_value=1, value=1, key=f"start_{fkey}")
    repeat = st.checkbox("같은 표가 가로로 반복됨", key=f"rep_{fkey}")
    step = st.number_input("한 블록의 열 개수", min_value=2, value=4, key=f"step_{fkey}") if repeat else None
    if st.button("이 설정으로 추출", key=f"go_{fkey}"):
        return {'sheet': sheet, 'ho': cols.index(ho), 'use': cols.index(use),
                'dong': None if dong == "없음" else cols.index(dong),
                'start': int(start) - 1, 'step': int(step) if step else None}
    return None


# ───────────────────────── 메인 화면 ─────────────────────────
threshold = st.sidebar.number_input("기준 사용량 (kWh) 미만", min_value=0.0, value=50.0, step=1.0)
include_zero = st.sidebar.checkbox("0 kWh 세대 포함", value=True)

files = st.file_uploader("엑셀 파일(.xlsx, .xls)을 선택하세요 (다중 선택 가능)",
                         type=["xlsx", "xls"], accept_multiple_files=True)

if files:
    tabs = st.tabs([f"📄 {f.name}" for f in files])
    for idx, (f, tab) in enumerate(zip(files, tabs)):
        with tab:
            fkey = f"{idx}_{f.name}"
            grids = load_grids(f.getvalue())
            if not grids:
                st.error("파일을 읽을 수 없습니다. 손상되었거나 지원하지 않는 형식입니다.")
                continue

            cfg = st.session_state.get(f"cfg_{fkey}")
            sheet = pname = None
            if cfg:
                sheet, pname = cfg['sheet'], '수동 설정'
                df_all = parse_with_config(grids[sheet], cfg)
                if st.button("수동 설정 초기화", key=f"reset_{fkey}"):
                    del st.session_state[f"cfg_{fkey}"]
                    st.rerun()
            else:
                sheet, pname, df_all = auto_parse_file(grids, f.name)
                if df_all.empty:
                    new_cfg = manual_ui(fkey, grids, order_sheets(grids)[0])
                    if new_cfg:
                        st.session_state[f"cfg_{fkey}"] = new_cfg
                        st.rerun()
                    continue

            if df_all.empty:
                st.error("선택한 설정으로 추출된 세대가 없습니다. 열/시작 행을 다시 확인하세요.")
                continue

            st.info(f"📌 분석 시트: `{sheet}`  |  사용된 파서: **{pname}**")
            lo = df_all['사용량(kWh)'] >= 0 if include_zero else df_all['사용량(kWh)'] > 0
            df_under = df_all[(df_all['사용량(kWh)'] < threshold) & lo].reset_index(drop=True)
            df_under.index = df_under.index + 1

            m1, m2 = st.columns(2)
            m1.metric("전체 인식 세대수", f"{len(df_all)} 세대")
            m2.metric(f"⚠️ {threshold:g} kWh 미만 세대수", f"{len(df_under)} 세대")

            with st.expander("인식 결과 검증 (전체 세대 앞/뒤 5건)"):
                st.dataframe(pd.concat([df_all.head(5), df_all.tail(5)]), use_container_width=True)
                st.caption("호수·사용량이 실제 엑셀과 맞는지 확인하세요. 틀리면 파일을 다시 올리지 말고 '수동 설정'을 사용하세요.")
                if not cfg and st.button("이 파일을 수동 설정으로 다시 지정", key=f"force_{fkey}"):
                    st.session_state[f"force_{fkey}"] = True

            if st.session_state.get(f"force_{fkey}") and not cfg:
                new_cfg = manual_ui(fkey, grids, sheet)
                if new_cfg:
                    st.session_state[f"cfg_{fkey}"] = new_cfg
                    st.session_state.pop(f"force_{fkey}", None)
                    st.rerun()

            if df_under.empty:
                st.success("기준 미만 세대가 없습니다.")
            else:
                st.dataframe(df_under, use_container_width=True)
                buf = io.BytesIO()
                out = df_under.copy()
                out.insert(0, 'No', range(1, len(out) + 1))
                with pd.ExcelWriter(buf, engine='openpyxl') as w:
                    out.to_excel(w, index=False, sheet_name='미만세대')
                base = re.sub(r'\.xlsx?$', '', f.name)
                st.download_button(f"📥 {f.name} 결과 엑셀 다운로드", buf.getvalue(),
                                   file_name=f"{threshold:g}미만세대_{base}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key=f"dl_{fkey}")
