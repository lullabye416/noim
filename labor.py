import pandas as pd
from typing import Dict, Callable, Optional
from excel_template import ExcelTemplate


# ── 데이터 계산 ───────────────────────────────────────────────────

def run_migration(settings: Dict, base_file: str, progress_callback: Callable = None):
    """
    입력 엑셀에서 근로자·출근 데이터를 읽어 날짜별로 집계하고,
    ExcelTemplate을 통해 노임일보 xlsx를 생성한다.
    """

    df = pd.read_excel(
        settings['input_file'],
        sheet_name=settings['input_sheet'],
        skiprows=settings['header_row'],
        header=None
    )

    workers = []
    for i in range(0, len(df) - 1, 2):
        row1 = df.iloc[i]
        row2 = df.iloc[i + 1]
        if pd.isna(row1[settings['name_col']]):
            continue
        days = {}
        for day in range(1, 16):
            days[day] = row1[settings['first_day_col1'] + (day - 1)]
        for day in range(16, 32):
            idx = settings['first_day_col2'] + (day - 16)
            days[day] = row2[idx] if idx < len(row2) else 0
        work_desc = ''
        if settings.get('work_desc_enabled'):
            col = settings['work_desc_col']
            val = row1[col] if col < len(row1) else ''
            if pd.notna(val):
                work_desc = str(val).strip()
                if work_desc:
                    work_desc_map      = settings.get('work_desc_map', [])
                    work_desc_fallback = settings.get('work_desc_fallback', '현장정리')
                    matched = False
                    for entry in work_desc_map:
                        if entry['keyword'] and entry['keyword'] in work_desc:
                            work_desc = entry['value']
                            matched = True
                            break
                    if not matched:
                        work_desc = work_desc_fallback

        worker = {
            'name':      str(row1[settings['name_col']]).strip(),
            'jumin':     str(row1[settings['jumin_col']]).strip(),
            'cost':      row1[settings['cost_col']],
            'days':      days,
            'work_desc': work_desc,
        }
        worker['certi'] = f"{worker['jumin']}_{worker['cost']}"
        workers.append(worker)

    cumulative_dict = {w['certi']: 0 for w in workers}
    day_data = []

    for day in range(1, 32):
        # 그날 일한 사람 집합
        worked_today = {
            w['certi'] for w in workers
            if pd.notna(w['days'].get(day)) and w['days'].get(day) != 0
        }
        # 아무도 일하지 않은 날은 시트 생성 안 함
        if not worked_today:
            if progress_callback:
                progress_callback(int(day / 31 * 100))
            continue

        # 누계 업데이트 — 실제 공수 반영 (0.5, 0.75, 1, 1.25 등)
        for w in workers:
            if w['certi'] in worked_today:
                day_val = w['days'].get(day, 0)
                cumulative_dict[w['certi']] += float(day_val) if pd.notna(day_val) else 0

        # 전원 포함, 금일은 실제 공수 / 안 한 사람은 0
        day_workers = []
        for w in workers:
            if w['certi'] in worked_today:
                day_val = w['days'].get(day, 0)
                today = float(day_val) if pd.notna(day_val) else 0
            else:
                today = 0
            day_workers.append({**w, 'today': today, '_cumulative': cumulative_dict[w['certi']]})

        day_str  = f"{day:02d}"
        date_str = f"{settings['year']}-{settings['month']:02d}-{day_str}"
        day_data.append({'day_str': day_str, 'date_str': date_str, 'workers': day_workers})

        if progress_callback:
            progress_callback(int(day / 31 * 100))

    template = ExcelTemplate(base_file)
    output_file = template.write_output(day_data, settings)

    # ── 작업 결과 요약 ─────────────────────────────────────────────
    workers_detail = []
    for w in workers:
        cost = int(w['cost']) if pd.notna(w['cost']) else 0
        days_worked = cumulative_dict[w['certi']]
        workers_detail.append({
            'name': w['name'], 'cost': cost,
            'days_worked': days_worked, 'amount': days_worked * cost,
        })

    return {
        'output_file':     output_file,
        'total_workers':   len(workers),
        'total_mandays':   sum(cumulative_dict.values()),
        'total_amount':    sum(wd['amount'] for wd in workers_detail),
        'work_days_count': len(day_data),
        'workers_detail':  workers_detail,
    }