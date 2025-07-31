# 필요한 라이브러리 import
import streamlit as st                    # Streamlit 웹앱 프레임워크
import pandas as pd                      # 데이터프레임 처리용 pandas
import requests                          # HTTP 요청을 위한 라이브러리
import xml.etree.ElementTree as ET       # XML 파싱을 위한 모듈
from collections import defaultdict       # 중복 키 병합용 defaultdict
from openpyxl import Workbook             # 엑셀 파일 생성용 openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment  # 스타일 관련 모듈
from openpyxl.utils import get_column_letter  # 열 너비 계산용
from io import BytesIO                   # 메모리 내 파일 처리를 위한 모듈

# 앱 제목 출력
st.title("유해물질 CAS 번호 조회 및 유해성 정보 수집기_0731")

# 파일 업로드 컴포넌트
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (예: data-회사명.xlsx)", type="xlsx")

# 사용자가 파일을 업로드한 경우
if uploaded_file:
    # 파일 이름에서 회사명 추출 (파일명 형식: data-회사명.xlsx)
    file_name = uploaded_file.name
    company_name = file_name.split('-')[-1].split('.')[0]

    # 엑셀 파일의 첫 번째 시트를 읽기 (헤더 없음)
    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    # 유효한(비어 있지 않은) 열만 유지
    valid_cols = [col for col in df.columns if df[col].notna().sum() > 0]
    df = df[valid_cols]

    # 첫 번째 행을 헤더로 사용하고 데이터는 그 아래 행부터 사용
    header_row = df.iloc[0]
    data_rows = df[1:].copy()
    data_rows.columns = header_row

    # 열 이름 자동 매핑 (필요한 컬럼 찾기)
    col_map = {}
    for col in data_rows.columns:
        if 'cas' in str(col).lower():
            col_map['CAS No.'] = col
        elif '물질' in str(col):
            col_map['물질명칭'] = col
        elif '#' in str(col) or '번호' in str(col):
            col_map['#'] = col

    # 필요한 열만 추출하고 이름 통일
    try:
        data_rows = data_rows[[col_map['#'], col_map['물질명칭'], col_map['CAS No.']]].copy()
        data_rows.columns = ['#', '물질명칭', 'CAS No.']
    except:
        # 필수 열이 없으면 에러 출력 후 중단
        st.error("필수 열을 찾을 수 없습니다. #, 물질명칭, CAS No. 컬럼이 필요합니다.")
        st.stop()

    # CAS 번호 문자열로 처리하고 공백 제거
    data_rows['CAS No.'] = data_rows['CAS No.'].astype(str).str.strip()

    # 공단 인증키 (URL 인코딩된 상태)
    SERVICE_KEY = 'MJFEGDzjkGr4Rg4pQtOxcYT%2BxteNCe0HuK0PUWKt%2B4hZHqYk%2BpNIf3RwocbhI1twsbNknwMur9m0fcPZir9jyg%3D%3D'

    # 결과 저장용 리스트와 컬럼 집합
    all_results = []
    all_new_columns = set()

    # 진행률 표시
    progress = st.progress(0)

    # 각 행에 대해 CAS 번호로 조회 시작
    for idx, row in enumerate(data_rows.itertuples(index=False), start=1):
        cas = row[2]              # CAS 번호
        id_num = row[0]           # 순번
        name = row[1]             # 물질명칭
        result = {'#': id_num, '물질명칭': name, 'CAS No.': cas}  # 결과 딕셔너리

        try:
            # chemId 조회 API 호출
            url_id = 'https://msds.kosha.or.kr/openapi/service/msdschem/chemlist'
            params_id = {'serviceKey': SERVICE_KEY, 'searchWrd': cas, 'searchCnd': 1}
            res_id = requests.get(url_id, params=params_id, timeout=10)
            res_id.encoding = 'utf-8'
            root_id = ET.fromstring(res_id.text)

            chem_id = root_id.findtext('.//chemId')  # chemId 추출

            if not chem_id:
                # chemId 없으면 결과 없음 처리
                result['결과없음'] = '공단 MSDS 없음'
                all_results.append(result)
                progress.progress(idx / len(data_rows))
                continue

            # 상세 정보 조회 API 호출
            url_detail = 'https://msds.kosha.or.kr/openapi/service/msdschem/chemdetail02'
            params_detail = {'serviceKey': SERVICE_KEY, 'chemId': chem_id}
            res_detail = requests.get(url_detail, params=params_detail, timeout=10)
            res_detail.encoding = 'utf-8'
            root_detail = ET.fromstring(res_detail.text)

            # B02 항목 추출 (유해성 정보)
            b02_detail = None
            for item in root_detail.findall('.//item'):
                if item.findtext('msdsItemCode') == 'B02':
                    b02_detail = item.findtext('itemDetail')
                    break

            # 유해성 정보가 없거나 "자료없음"이면 결과 없음 처리
            if b02_detail is None or b02_detail.strip() in ['', '자료없음']:
                result['결과없음'] = '자료 없음'
            else:
                # |로 구분된 항목들을 파싱하여 key-value로 병합
                merged = defaultdict(list)
                for entry in b02_detail.split('|'):
                    if ':' in entry and '자료없음' not in entry:
                        k, v = map(str.strip, entry.rsplit(':', 1))
                        v = v.replace('구분', '').strip()
                        if v not in merged[k]:
                            merged[k].append(v)

                # 결과에 유해성 정보 추가
                if merged:
                    for k, v_list in merged.items():
                        result[k] = ', '.join(v_list)
                        all_new_columns.add(k)

        except Exception as e:
            # 요청 실패 등 예외 처리
            result['결과없음'] = f'조회 오류: {str(e)}'

        # 결과 저장
        all_results.append(result)
        # 진행률 업데이트
        progress.progress(idx / len(data_rows))

    # 결과를 데이터프레임으로 정리
    result_df = pd.DataFrame(all_results)

    # '결과없음' 컬럼의 결측치를 빈 문자열로 채움
    if '결과없음' in result_df.columns:
        result_df['결과없음'] = result_df['결과없음'].fillna('')

    # 열 순서 정리
    first_cols = ['#', '물질명칭', 'CAS No.', '결과없음']
    remaining_cols = sorted([col for col in result_df.columns if col not in first_cols])
    ordered_cols = first_cols + remaining_cols
    for col in ordered_cols:
        if col not in result_df.columns:
            result_df[col] = ''  # 누락된 열 채움
    result_df = result_df[ordered_cols]

    # 엑셀 파일 생성 시작
    wb = Workbook()
    ws = wb.active
    ws.title = '결과'

    # 스타일 정의
    header_font = Font(bold=True, size=10)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))

    # 왼쪽에 빈 열 삽입 (여백용)
    ws.insert_cols(1)
    ws.column_dimensions['A'].width = 2

    # 회사명 출력
    ws['B1'] = company_name
    ws['B1'].font = Font(bold=True, size=10)
    ws['B1'].alignment = Alignment(horizontal='left')

    # 헤더 작성
    ws.append([''] + result_df.columns.tolist())
    for col_idx, cell in enumerate(ws[2][1:], start=2):
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    # 데이터 행 추가
    for row_idx, row in enumerate(result_df.itertuples(index=False), start=3):
        row_values = [''] + list(row)
        ws.append(row_values)
        for col_idx, value in enumerate(row_values[1:], start=2):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(size=10)
            cell.border = thin_border
            if col_idx == result_df.columns.get_loc("CAS No.") + 2:
                cell.number_format = '@'  # CAS 번호는 문자열로 표시

    # 열 너비 자동 조정
    for column_cells in ws.iter_cols(min_col=2):
        max_len = 0
        col = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col].width = max_len + 2

    # 엑셀 파일을 메모리 버퍼에 저장
    file_buffer = BytesIO()
    wb.save(file_buffer)
    file_buffer.seek(0)

    # 성공 메시지 출력 및 다운로드 버튼 제공
    st.success("처리가 완료되었습니다. 아래 버튼을 눌러 결과를 다운로드하세요.")
    st.download_button(
        label="결과 엑셀 다운로드",
        data=file_buffer,
        file_name=f"유해물질정보-{company_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
