#!/usr/bin/env python3
"""
행사 종료 후, 8대(L1~L6, D1~D2)에서 각각 내보낸
"신청자명부_*.csv"와 "상담요약_*.csv" 파일들을 한 폴더에 모은 뒤,
이 스크립트로 통합 파일 2개를 만듭니다.

기능:
- 기기코드 누락/알 수 없는 코드 검사
- 세션번호 중복 검사 (있으면 안 되는 이상 상황이므로 경고만 하고 그대로 둠 - 사람이 판단)
- 신청자명부: 연락처 뒤 4자리 + 성명이 같은 "중복 신청 후보"를 표시만 함 (자동 삭제하지 않음)
- 같은 중복 후보 그룹이 서로 다른 컨설턴트에 배정된 경우 "배정 충돌"로 표시
- 병합 결과 요약 출력 (총 상담건수, 기기별 건수, 신청건수, 중복 후보 건수)

사용법:
    python3 merge.py <신청자명부·상담요약 CSV들이 모인 폴더> [출력 폴더]

출력 폴더를 생략하면 입력 폴더 안에 결과 파일을 만듭니다.
"""

import sys
import csv
from pathlib import Path
from datetime import date

VALID_DEVICES = {'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'D1', 'D2'}


def load_csv(path):
    with open(path, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def device_of(session_id):
    return session_id.split('-')[0] if session_id and '-' in session_id else ''


def phone_last4(phone):
    digits = ''.join(c for c in (phone or '') if c.isdigit())
    return digits[-4:] if len(digits) >= 4 else ''


def find_files(folder, prefix):
    return sorted(Path(folder).glob(prefix + '*.csv'))


def merge_sessions(files):
    rows = []
    seen_ids = {}
    device_counts = {}
    warnings = []
    for f in files:
        for r in load_csv(f):
            sid = r.get('session_id', '')
            dev = device_of(sid)
            device_counts[dev] = device_counts.get(dev, 0) + 1
            if dev not in VALID_DEVICES:
                warnings.append(f'알 수 없는 기기코드: 세션번호="{sid}" (파일: {f.name})')
            if sid in seen_ids and seen_ids[sid] != f.name:
                warnings.append(f'세션번호 중복: "{sid}" (파일 {seen_ids[sid]} 및 {f.name}) - 사람이 직접 확인하세요')
            else:
                seen_ids[sid] = f.name
            rows.append(r)
    return rows, device_counts, warnings


def merge_applicants(files):
    rows = []
    device_counts = {}
    for f in files:
        for r in load_csv(f):
            sid = r.get('session_id', '')
            dev = device_of(sid)
            device_counts[dev] = device_counts.get(dev, 0) + 1
            r = dict(r)
            r['source_file'] = f.name
            rows.append(r)

    groups = {}
    for i, r in enumerate(rows):
        key = (r.get('name', '').strip(), phone_last4(r.get('phone', '')))
        if key == ('', ''):
            continue
        groups.setdefault(key, []).append(i)
    dup_groups = {k: v for k, v in groups.items() if len(v) > 1}

    for r in rows:
        r['dup_candidate_group'] = ''
    for gi, (key, idxs) in enumerate(dup_groups.items(), start=1):
        for i in idxs:
            rows[i]['dup_candidate_group'] = f'DUP-{gi}'

    conflicts = []
    for key, idxs in dup_groups.items():
        consultants = set(
            rows[i].get('assigned_consultant', '').strip()
            for i in idxs if rows[i].get('assigned_consultant', '').strip()
        )
        if len(consultants) > 1:
            conflicts.append((key, idxs, consultants))
            for i in idxs:
                rows[i]['assignment_conflict'] = 'Y'
    for r in rows:
        r.setdefault('assignment_conflict', '')

    return rows, device_counts, dup_groups, conflicts


def write_csv(path, rows, fieldnames):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, '') for k in fieldnames})


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return
    in_folder = Path(args[0])
    out_folder = Path(args[1]) if len(args) > 1 else in_folder
    today = date.today().strftime('%Y%m%d')

    session_files = find_files(in_folder, '상담요약_')
    applicant_files = find_files(in_folder, '신청자명부_')

    print(f'상담요약 파일 {len(session_files)}개, 신청자명부 파일 {len(applicant_files)}개 발견')

    session_rows, sess_dev_counts, sess_warnings = merge_sessions(session_files)
    app_rows, app_dev_counts, dup_groups, conflicts = merge_applicants(applicant_files)

    session_headers = ['session_id', 'timestamp', 'intake_summary', 'recommended_job_ids', 'counselor_memo']
    applicant_headers = ['session_id', 'name', 'phone', 'contact_time', 'region', 'interest_families',
                          'consent', 'assigned_consultant', 'assignment_status',
                          'dup_candidate_group', 'assignment_conflict', 'source_file']

    out_session_path = out_folder / f'상담요약_병합_{today}.csv'
    out_applicant_path = out_folder / f'신청자명부_병합_{today}.csv'
    write_csv(out_session_path, session_rows, session_headers)
    write_csv(out_applicant_path, app_rows, applicant_headers)

    print('\n=== 병합 결과 요약 ===')
    print(f'총 상담건수: {len(session_rows)}')
    print('기기별 상담건수:', dict(sorted(sess_dev_counts.items())))
    print(f'총 신청건수: {len(app_rows)}')
    print('기기별 신청건수:', dict(sorted(app_dev_counts.items())))
    print(f'중복 신청 후보 그룹: {len(dup_groups)}개 (관련 행 {sum(len(v) for v in dup_groups.values())}건) - 자동 삭제하지 않았습니다. 사람이 직접 확인하세요.')
    print(f'배정 충돌(같은 후보가 서로 다른 컨설턴트에 배정됨): {len(conflicts)}건')
    if sess_warnings:
        print('\n=== 세션 데이터 경고 ===')
        for w in sess_warnings:
            print(' -', w)
    print(f'\n저장됨: {out_session_path}')
    print(f'저장됨: {out_applicant_path}')


if __name__ == '__main__':
    main()
