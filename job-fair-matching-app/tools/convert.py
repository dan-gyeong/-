#!/usr/bin/env python3
"""
엑셀(.xlsx) 참여기업 목록을 data/companies.csv 형식으로 변환하는 스크립트.

- 별도 설치(pip install) 없이 파이썬 표준 라이브러리만 사용합니다.
  (엑셀 파일(.xlsx)은 사실 zip 압축 파일이라, zipfile/xml 모듈로 직접 읽습니다.)
- 원본 엑셀의 첫 번째 시트, 첫 번째 행을 "열 제목(헤더)"로 인식합니다.
- 헤더 이름이 companies.csv 표준 열 이름과 정확히 같지 않아도, 자주 쓰는 표현은
  아래 HEADER_ALIASES 표에서 자동으로 매칭합니다. 매칭되지 않는 열은 무시됩니다.
- 이 스크립트는 "값을 추측해서 채우지" 않습니다. 엑셀 셀이 비어 있으면 결과도 빈칸입니다.

사용법 (터미널/명령 프롬프트에서):
    python3 convert.py 원본파일.xlsx [출력파일.csv]

출력파일을 생략하면 ../data/companies.csv 로 저장됩니다(기존 파일을 덮어씁니다. 미리 백업하세요).

빈 템플릿만 만들고 싶다면:
    python3 convert.py --template [출력파일.csv]
"""

import sys
import csv
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

COLUMNS = ['job_id', 'booth_no', 'company', 'industry', 'job_title', 'job_family', 'headcount',
           'work_area', 'employment_type', 'work_schedule', 'pay_text', 'pay_monthly_min',
           'required_license', 'preferred_license', 'required_exp_years', 'physical_load',
           'age_note', 'etc_note', 'status']

# 원본 엑셀 헤더가 아래 표현이면 표준 열 이름으로 자동 매칭합니다.
# 왼쪽(표준 열 이름)에 맞는 표현이 더 있다면 이 리스트에 추가해서 재사용하세요.
HEADER_ALIASES = {
    'job_id': ['job_id', '직무id', '직무번호'],
    'booth_no': ['booth_no', '부스번호', '부스'],
    'company': ['company', '기업명', '기업', '회사명', '회사'],
    'industry': ['industry', '업종'],
    'job_title': ['job_title', '직무', '직무명', '채용직무'],
    'job_family': ['job_family', '직무군', '직무군코드'],
    'headcount': ['headcount', '채용인원', '인원'],
    'work_area': ['work_area', '근무지', '근무지역'],
    'employment_type': ['employment_type', '고용형태'],
    'work_schedule': ['work_schedule', '근무형태', '근무시간'],
    'pay_text': ['pay_text', '급여', '급여조건'],
    'pay_monthly_min': ['pay_monthly_min', '월급하한', '월급환산하한'],
    'required_license': ['required_license', '필수자격', '필수자격면허'],
    'preferred_license': ['preferred_license', '우대자격'],
    'required_exp_years': ['required_exp_years', '필수경력연수', '필수경력'],
    'physical_load': ['physical_load', '신체강도'],
    'age_note': ['age_note', '정년', '연령관련문구'],
    'etc_note': ['etc_note', '기타', '기타특이사항', '비고'],
    'status': ['status', '상태', '모집상태'],
}


def col_letter_to_index(letter):
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1


def read_xlsx_first_sheet(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('m:si', NS):
                text = ''.join(t.text or '' for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'))
                shared.append(text)

        # 첫 번째 워크시트 찾기 (workbook.xml의 sheets 순서 기준)
        sheet_path = 'xl/worksheets/sheet1.xml'
        names = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet')]
        if names:
            sheet_path = sorted(names)[0]

        root = ET.fromstring(z.read(sheet_path))
        rows_xml = root.find('m:sheetData', NS)
        rows = []
        for row_xml in rows_xml.findall('m:row', NS):
            row = {}
            max_idx = -1
            for c in row_xml.findall('m:c', NS):
                ref = c.get('r', '')
                letters = ''.join(ch for ch in ref if ch.isalpha())
                idx = col_letter_to_index(letters) if letters else 0
                v = c.find('m:v', NS)
                t = c.get('t')
                if v is None:
                    value = ''
                elif t == 's':
                    value = shared[int(v.text)] if v.text and v.text.isdigit() else ''
                elif t == 'inlineStr':
                    is_el = c.find('m:is', NS)
                    value = ''.join(x.text or '' for x in is_el.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')) if is_el is not None else ''
                else:
                    value = v.text or ''
                row[idx] = value
                max_idx = max(max_idx, idx)
            rows.append([row.get(i, '') for i in range(max_idx + 1)])
        return rows


def build_header_map(header_row):
    normalized = [h.strip().lower().replace(' ', '') for h in header_row]
    mapping = {}
    for std_col, aliases in HEADER_ALIASES.items():
        for i, h in enumerate(normalized):
            if h in [a.lower().replace(' ', '') for a in aliases]:
                mapping[std_col] = i
                break
    return mapping


def convert(xlsx_path, out_path):
    rows = read_xlsx_first_sheet(xlsx_path)
    if not rows:
        print('시트에서 데이터를 찾지 못했습니다.')
        return
    header, data_rows = rows[0], rows[1:]
    mapping = build_header_map(header)

    missing = [c for c in COLUMNS if c not in mapping]
    if missing:
        print('※ 다음 표준 열은 원본 엑셀에서 자동으로 찾지 못해 빈 칸으로 채워집니다:')
        print('  ' + ', '.join(missing))
        print('  (헤더 이름을 맞추거나 HEADER_ALIASES 표에 표현을 추가하면 자동 매칭됩니다)')

    out_rows = []
    for i, r in enumerate(data_rows, start=1):
        if all((c or '').strip() == '' for c in r):
            continue
        out = {}
        for std_col in COLUMNS:
            idx = mapping.get(std_col)
            out[std_col] = (r[idx].strip() if idx is not None and idx < len(r) and r[idx] else '')
        if not out['job_id']:
            out['job_id'] = f"{out['booth_no'] or 'ROW'}-{i}"
        out_rows.append(out)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(out_rows)
    print(f'완료: {len(out_rows)}건을 {out_path} 에 저장했습니다.')


def write_template(out_path):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(COLUMNS)
    print(f'빈 템플릿을 {out_path} 에 만들었습니다.')


def main():
    args = sys.argv[1:]
    default_out = Path(__file__).resolve().parent.parent / 'data' / 'companies.csv'

    if not args:
        print(__doc__)
        return

    if args[0] == '--template':
        out = args[1] if len(args) > 1 else default_out
        write_template(out)
        return

    xlsx_path = args[0]
    out_path = args[1] if len(args) > 1 else default_out
    convert(xlsx_path, out_path)


if __name__ == '__main__':
    main()
