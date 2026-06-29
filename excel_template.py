import html
import os
import re
import shutil
import tempfile
import uuid
import zipfile


def _parse_rows(sheet_xml: str) -> tuple[str, dict[int, str], str]:
    """sheet XML을 before / row_dict / after 세 부분으로 분리."""
    m_before = re.search(r'^(.*?<sheetData>)', sheet_xml, re.DOTALL)
    m_after  = re.search(r'(</sheetData>.*?)$', sheet_xml, re.DOTALL)
    assert m_before and m_after, "sheetData 태그를 찾을 수 없습니다."
    before = m_before.group(1)
    after  = m_after.group(1)
    row_dict = {}
    for m in re.finditer(r'(<row r="(\d+)"[^>]*>.*?</row>)', sheet_xml, re.DOTALL):
        row_dict[int(m.group(2))] = m.group(1)
    return before, row_dict, after


class ExcelTemplate:
    """
    sample.xlsx의 모든 디자인 요소를 보존하면서 output xlsx를 생성한다.

    보존 항목:
    - 행 높이 (ht, customHeight, x14ac:dyDescent)
    - 셀 스타일 인덱스 (s=) — 폰트·테두리·정렬·숫자형식
    - 셀 병합 (결재란 K37:K40 등 포함)
    - 결재란 행(34행~) 원본 그대로 유지
    - 카메라 이미지 소스 범위 ($K$37:$O$40)
    - 인쇄 영역 (Print_Area definedName)
    - 페이지 설정 (pageSetup, pageMargins)
    """

    def __init__(self, path: str):
        self.path = path

        # zip 전체 내용을 메모리에 캐싱
        self._files: dict[str, bytes] = {}
        with zipfile.ZipFile(path, 'r') as z:
            for name in z.namelist():
                self._files[name] = z.read(name)

        # 워크시트 경로 파악 (workbook.xml.rels 기반)
        wb_rels = self._files['xl/_rels/workbook.xml.rels'].decode('utf-8')
        m = re.search(r'Type="[^"]*worksheet"[^>]*Target="([^"]*)"', wb_rels)
        assert m, "workbook.xml.rels에서 worksheet 관계를 찾을 수 없습니다."
        rel_target      = m.group(1)                # e.g. 'worksheets/sheet1.xml'
        self._sheet_key = f'xl/{rel_target}'
        self._rels_key  = (
            f'xl/worksheets/_rels/'
            f'{os.path.basename(rel_target)}.rels'
        )

        # 시트 XML 파싱
        sheet_xml = self._files[self._sheet_key].decode('utf-8')
        self._before: str
        self._row_dict: dict[int, str]
        self._after: str
        self._before, self._row_dict, self._after = _parse_rows(sheet_xml)

        # 행 여는 태그 캐싱 — ht, customHeight, x14ac:dyDescent 등 보존
        self._row_open_tags: dict[int, str] = {}
        for rn, rx in self._row_dict.items():
            tm = re.match(r'(<row [^>]*>)', rx)
            if tm:
                self._row_open_tags[rn] = tm.group(1)

        # 셀 스타일 인덱스 캐싱 — key: "B3", value: "8"
        self._cell_s: dict[str, str] = {}
        for rn, rx in self._row_dict.items():
            for cm in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/|>)', rx):
                addr = cm.group(1)
                sm   = re.search(r'\bs="(\d+)"', cm.group(2))
                if sm:
                    self._cell_s[addr] = sm.group(1)

        # 합계 행 감지 — start_row(8) 이후 SUM 수식이 있는 첫 번째 행
        self._sum_row: int | None = None
        for rn in sorted(self._row_dict.keys()):
            if rn <= self._DATA_REF_ROW:
                continue
            if re.search(r'<f[^>]*>SUM\(', self._row_dict[rn]):
                self._sum_row = rn
                break

        # ── 결재란 행 범위를 VML FmlaPict에서 파싱 ──────────────────
        self._keoljairan_cells: dict[int, list[str]] = {}
        self._keoljairan_open:  dict[int, str]       = {}
        self._keoljairan_new_start: int | None       = None
        self._kj_row_start: int | None               = None
        self._kj_row_end:   int | None               = None

        for _fname, _fbytes in self._files.items():
            if _fname.endswith('.vml'):
                _fmla = re.search(
                    r'<x:FmlaPict>\$[A-Z]+\$(\d+):\$[A-Z]+\$(\d+)</x:FmlaPict>',
                    _fbytes.decode('utf-8')
                )
                if _fmla:
                    self._kj_row_start = int(_fmla.group(1))
                    self._kj_row_end   = int(_fmla.group(2))
                break

        if self._kj_row_start is not None:
            n_kj = self._kj_row_end - self._kj_row_start + 1
            for offset in range(n_kj):
                rn       = self._kj_row_start + offset
                row_xml  = self._row_dict.get(rn, '')
                open_tag = self._row_open_tags.get(rn, '')
                kp_cells = [
                    m.group(0)
                    for m in re.finditer(
                        r'<c r="([A-Z]+)\d+"[^/>]*(?:/>|>.*?</c>)', row_xml, re.DOTALL
                    )
                    if m.group(1) not in self._DATA_COLS
                ]
                self._keoljairan_cells[offset] = kp_cells
                self._keoljairan_open[offset]  = open_tag

        # K+ 열 관련 merge 파싱 (after 섹션 갱신용)
        self._keoljairan_merges: list[str] = [
            ref for ref in re.findall(r'<mergeCell ref="([^"]+)"', self._after)
            if any(c not in self._DATA_COLS for c in re.findall(r'[A-Z]+', ref))
        ]

    # ── 내부 헬퍼 ─────────────────────────────────────────────────

    # 데이터 행 스타일 기준이 되는 샘플의 첫 번째 데이터 행
    _DATA_REF_ROW = 8

    def prepare(self, n_workers: int) -> None:
        """
        템플릿을 n_workers 데이터 행에 맞게 한 번만 준비한다.

        1. RAW 인원(n_workers) vs SAMPLE NO~합계 사이 행 수 비교
        2. 부족한 행 삽입 (row 8 스타일·높이 기준)
        3. 디자인 통일 — 모든 데이터 행 open-tag을 row 8 기준으로 정규화
        4. 합계 행을 마지막 데이터 행 바로 다음으로 이동 + SUM 범위 갱신
        """
        if self._sum_row is None:
            return

        start_row       = self._DATA_REF_ROW          # 8
        orig_data_count = self._sum_row - start_row    # 샘플의 기존 데이터 행 수

        ref_open = self._row_open_tags.get(
            start_row, f'<row r="{start_row}" spans="1:16">'
        )

        # ── 1·2. 부족한 행 삽입 ─────────────────────────────────
        for i in range(orig_data_count, n_workers):
            new_rn = start_row + i
            if new_rn in self._row_dict:
                continue                              # 이미 존재하면 건너뜀
            open_tag = re.sub(r'r="\d+"', f'r="{new_rn}"', ref_open)
            if 'spans=' in open_tag:
                open_tag = re.sub(r'spans="[^"]*"', 'spans="1:16"', open_tag)
            else:
                open_tag = open_tag.replace('<row ', '<row spans="1:16" ')
            # A~I 스타일 셀만 삽입 (K-O 결재란은 원래 행 위치에 고정)
            cells = ''
            for col in 'ABCDEFGHI':
                s = self._cell_s.get(f'{col}{start_row}')
                cells += (f'<c r="{col}{new_rn}" s="{s}"/>' if s
                          else f'<c r="{col}{new_rn}"/>')
            self._row_dict[new_rn]      = f'{open_tag}{cells}</row>'
            self._row_open_tags[new_rn] = open_tag

        # ── 3. 디자인 통일 — 기존 데이터 행 open-tag 정규화 ────
        for i in range(orig_data_count):
            rn = start_row + i
            if rn == start_row or rn not in self._row_dict:
                continue
            rx      = self._row_dict[rn]
            old_tag = re.match(r'(<row [^>]*>)', rx)
            if old_tag:
                new_tag = re.sub(r'r="\d+"', f'r="{rn}"', ref_open)
                if 'spans=' in new_tag:
                    new_tag = re.sub(r'spans="[^"]*"', 'spans="1:16"', new_tag)
                else:
                    new_tag = new_tag.replace('<row ', '<row spans="1:16" ')
                self._row_dict[rn]      = rx.replace(old_tag.group(1), new_tag, 1)
                self._row_open_tags[rn] = new_tag

        # ── 4. 합계 행 이동 + SUM 수식 범위 갱신 ───────────────
        if self._sum_row in self._row_dict:
            sum_xml      = self._row_dict.pop(self._sum_row)
            new_sum_row  = start_row + n_workers
            last_data_rn = start_row + n_workers - 1

            sum_xml = re.sub(r'(<row r=")\d+"',
                             rf'\g<1>{new_sum_row}"', sum_xml)
            sum_xml = re.sub(
                r'(<c r=")([A-Z]+)\d+"',
                lambda m: f'{m.group(1)}{m.group(2)}{new_sum_row}"',
                sum_xml
            )
            sum_xml = re.sub(
                r'SUM\(([A-Z]+)(\d+):([A-Z]+)\d+\)',
                lambda m: f'SUM({m.group(1)}{m.group(2)}:{m.group(3)}{last_data_rn})',
                sum_xml
            )
            self._row_dict[new_sum_row] = sum_xml
            self._sum_row               = new_sum_row

            # ── 5. 결재란 행 재배치 (합계행 바로 아래) ──────────────
            if self._keoljairan_cells and self._kj_row_start is not None:
                n_kj         = len(self._keoljairan_cells)
                new_kj_start = new_sum_row + 1

                # (a) 원래 결재란 행들에서 K+ 셀 제거
                for offset in range(n_kj):
                    rn = self._kj_row_start + offset
                    if rn in self._row_dict:
                        self._row_dict[rn] = self._strip_kp_cells(self._row_dict[rn])

                # (b) 새 결재란 행 추가 (원본 높이·스타일 보존)
                for offset in range(n_kj):
                    new_rn   = new_kj_start + offset
                    orig_tag = self._keoljairan_open.get(offset, '')
                    if orig_tag:
                        open_tag = re.sub(r'r="\d+"', f'r="{new_rn}"', orig_tag)
                    else:
                        open_tag = f'<row r="{new_rn}" spans="11:16">'
                    cells_xml = ''.join(
                        re.sub(r'r="([A-Z]+)\d+"', rf'r="\g<1>{new_rn}"', cell)
                        for cell in self._keoljairan_cells.get(offset, [])
                    )
                    self._row_dict[new_rn]      = f'{open_tag}{cells_xml}</row>'
                    self._row_open_tags[new_rn] = open_tag

                # (c) after 섹션의 mergeCells 행 번호 갱신
                for old_ref in self._keoljairan_merges:
                    new_ref = re.sub(
                        r'([A-Z]+)(\d+)',
                        lambda m: (
                            f'{m.group(1)}'
                            f'{int(m.group(2)) - self._kj_row_start + new_kj_start}'
                        ),
                        old_ref
                    )
                    self._after = self._after.replace(
                        f'<mergeCell ref="{old_ref}"',
                        f'<mergeCell ref="{new_ref}"',
                        1
                    )

                self._keoljairan_new_start = new_kj_start

    def _row_open(self, row_num: int, override_spans: str | None = None) -> str:
        """원본 행 여는 태그 반환.
        해당 행이 샘플에 없으면 기준 행(row 8)의 태그를 row 번호만 교체해 재사용.
        override_spans 지정 시 spans 속성만 교체."""
        if row_num in self._row_open_tags:
            tag = self._row_open_tags[row_num]
        else:
            ref = self._row_open_tags.get(self._DATA_REF_ROW, f'<row r="{row_num}" spans="1:9">')
            tag = re.sub(r'r="\d+"', f'r="{row_num}"', ref)
        if override_spans:
            if 'spans=' in tag:
                tag = re.sub(r'spans="[^"]*"', f'spans="{override_spans}"', tag)
            else:
                tag = tag.replace('<row ', f'<row spans="{override_spans}" ')
        return tag

    def _s_attr(self, addr: str) -> str:
        """셀 주소 → ' s="N"' 속성 문자열.
        해당 셀이 샘플에 없으면 기준 행(row 8)의 같은 컬럼 스타일로 폴백."""
        v = self._cell_s.get(addr)
        if v is None:
            col = re.match(r'([A-Z]+)', addr).group(1)
            v = self._cell_s.get(f'{col}{self._DATA_REF_ROW}')
        return f' s="{v}"' if v else ''

    def _set_cell_inlineStr(self, row_xml: str, col: str, value: str) -> str:
        """기존 셀 스타일(s=)을 보존하면서 inlineStr 값으로 교체."""
        m_rn = re.search(r'<row r="(\d+)"', row_xml)
        assert m_rn, "row_xml에서 행 번호를 찾을 수 없습니다."
        rn      = m_rn.group(1)
        addr    = f'{col}{rn}'
        s_attr  = self._s_attr(addr)
        escaped = html.escape(str(value))
        new_c   = f'<c r="{addr}"{s_attr} t="inlineStr"><is><t>{escaped}</t></is></c>'
        pattern = rf'<c r="{re.escape(addr)}"[^/]*/?>(?:.*?</c>)?'
        if re.search(pattern, row_xml, re.DOTALL):
            return re.sub(pattern, new_c, row_xml, flags=re.DOTALL)
        return row_xml.replace('</row>', new_c + '</row>')

    # 데이터 열 범위 (A~I 만 덮어씀, 그 외는 원본 보존)
    _DATA_COLS = set('ABCDEFGHI')

    def _strip_kp_cells(self, row_xml: str) -> str:
        """row XML에서 J열 이후(A-I 외) 셀을 제거한다."""
        matches = [
            m for m in re.finditer(
                r'<c r="([A-Z]+)\d+"[^/>]*(?:/>|>.*?</c>)', row_xml, re.DOTALL
            )
            if m.group(1) not in self._DATA_COLS
        ]
        result = row_xml
        for m in reversed(matches):
            result = result[:m.start()] + result[m.end():]
        return result

    def _make_data_row(self, row_num: int, idx: int, w: dict) -> str:
        """
        데이터 행 XML 생성.
        - A~I 열: 근무자 데이터로 채움
        - J열 이후: 원본 샘플 셀 그대로 보존 (결재란은 prepare()에서 합계 아래로 이동됨)
        - 행 여는 태그·셀 스타일은 항상 기준 행(row 8) 기준
        """
        gender = '남' if w['jumin'].split('-')[1][0] in '1357' else '여'
        today  = w.get('today', 1)
        prev   = w['_cumulative'] - today          # 전회 = 누계 - 금일
        amount = w['_cumulative'] * int(w['cost'])

        # 항상 기준 행(row 8) 스타일 사용
        def s(col): return self._s_attr(f'{col}{self._DATA_REF_ROW}')

        cells = (
            f'<c r="A{row_num}"{s("A")}><v>{idx + 1}</v></c>'
            f'<c r="B{row_num}"{s("B")} t="inlineStr"><is><t>{html.escape(w["name"])}</t></is></c>'
            f'<c r="C{row_num}"{s("C")} t="inlineStr"><is><t>{gender}</t></is></c>'
            f'<c r="D{row_num}"{s("D")}><v>{today}</v></c>'
            f'<c r="E{row_num}"{s("E")}><v>{prev}</v></c>'
            f'<c r="F{row_num}"{s("F")}><v>{w["_cumulative"]}</v></c>'
            f'<c r="G{row_num}"{s("G")}><v>{int(w["cost"])}</v></c>'
            f'<c r="H{row_num}"{s("H")}><v>{amount}</v></c>'
            + (f'<c r="I{row_num}"{s("I")} t="inlineStr"><is><t>{html.escape(w["work_desc"])}</t></is></c>'
               if w.get('work_desc') and today > 0 else f'<c r="I{row_num}"{s("I")}/>')
        )

        # J열 이후 원본 셀 보존 (결재란 등 A~I 범위 밖 내용 유지)
        original = self._row_dict.get(row_num, '')
        if original:
            for cm in re.finditer(r'<c r="([A-Z]+)\d+"[^/>]*(?:/>|>.*?</c>)', original, re.DOTALL):
                if cm.group(1) not in self._DATA_COLS:
                    cells += cm.group(0)

        # 행 높이/스타일도 기준 행 태그에서 r= 만 교체
        ref_tag  = self._row_open_tags.get(self._DATA_REF_ROW, f'<row r="{row_num}" spans="1:9">')
        open_tag = re.sub(r'r="\d+"', f'r="{row_num}"', ref_tag)
        if 'spans=' in open_tag:
            open_tag = re.sub(r'spans="[^"]*"', 'spans="1:16"', open_tag)
        else:
            open_tag = open_tag.replace('<row ', '<row spans="1:16" ')
        return f'{open_tag}{cells}</row>'

    # ── 시트 XML 생성 ─────────────────────────────────────────────

    def build_sheet_xml(self, day_info: dict, settings: dict) -> str:
        """
        템플릿 기반으로 하루치 시트 XML 생성.
        after 섹션(병합·결재란·pageSetup·drawing 참조 등)은 일절 수정하지 않음.
        """
        before: str   = self._before
        row_dict: dict[int, str] = {**self._row_dict}  # 원본 보존을 위한 shallow copy
        after: str    = self._after                    # 결재란·병합·pageSetup 전부 원본 그대로

        # sheetView 정규화 — pageBreakPreview 오작동 방지
        before = re.sub(
            r'<sheetViews>.*?</sheetViews>',
            '<sheetViews><sheetView workbookViewId="0">'
            '<selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>',
            before, flags=re.DOTALL
        )

        # xr:uid 고유화 — 모든 시트가 같은 UID를 공유하면 Excel이 손상으로 인식
        before = re.sub(
            r'xr:uid="\{[^}]+\}"',
            f'xr:uid="{{{str(uuid.uuid4()).upper()}}}"',
            before
        )

        # B3(현장명), B4(날짜) — 원본 s= 스타일 보존하며 값 교체
        if 3 in row_dict:
            row_dict[3] = self._set_cell_inlineStr(row_dict[3], 'B', settings['site_name'])
        if 4 in row_dict:
            row_dict[4] = self._set_cell_inlineStr(row_dict[4], 'B', day_info['date_str'])

        # 데이터 행 교체 (prepare()에서 최대 인원 기준 행 확보·합계 행 이동됨)
        start_row = 8
        workers   = day_info['workers']

        for i, w in enumerate(workers):
            row_dict[start_row + i] = self._make_data_row(start_row + i, i, w)

        rows_xml = ''.join(row_dict[k] for k in sorted(row_dict))
        # 수식 셀의 캐시된 <v> 제거 → Excel 열 때 강제 재계산 (합계 행 등)
        rows_xml = re.sub(r'(<f>[^<]*</f>)<v>[^<]*</v>', r'\1', rows_xml)
        return before + rows_xml + after

    # ── xlsx 파일 출력 ────────────────────────────────────────────

    def write_output(self, day_data: list, settings: dict) -> str:
        """날짜별 시트를 생성해 xlsx 파일로 저장."""
        # ── 템플릿 준비 (1회) ──────────────────────────────────────
        # 1) RAW 인원 파악 → 2) NO~합계 사이 행 수 비교 → 3) 행 삽입 → 4) 디자인 통일
        if day_data:
            self.prepare(len(day_data[0]['workers']))

        tmp_dir = tempfile.mkdtemp()
        save_path = ''
        try:
            with zipfile.ZipFile(self.path, 'r') as zin:
                zin.extractall(tmp_dir)

            wb_rels_path = os.path.join(tmp_dir, 'xl', '_rels', 'workbook.xml.rels')
            wb_xml_path  = os.path.join(tmp_dir, 'xl', 'workbook.xml')
            ct_path      = os.path.join(tmp_dir, '[Content_Types].xml')

            wb_rels_text = open(wb_rels_path, encoding='utf-8').read()
            wb_xml       = open(wb_xml_path,  encoding='utf-8').read()
            ct_xml       = open(ct_path,      encoding='utf-8').read()

            sample_sheet_path = os.path.join(
                tmp_dir, self._sheet_key.replace('/', os.sep)
            )
            sample_rels_path  = os.path.join(
                tmp_dir, self._rels_key.replace('/', os.sep)
            )

            worksheets_dir = os.path.join(tmp_dir, 'xl', 'worksheets')
            ws_rels_dir    = os.path.join(worksheets_dir, '_rels')
            drawings_dir   = os.path.join(tmp_dir, 'xl', 'drawings')
            dr_rels_dir    = os.path.join(drawings_dir, '_rels')
            os.makedirs(ws_rels_dir, exist_ok=True)

            # rId 충돌 방지: worksheet 관계 제거 후 남은 최댓값 기준
            wb_rels_no_ws = re.sub(
                r'<Relationship[^>]*Type="[^"]*worksheet"[^>]*/>', '', wb_rels_text
            )
            existing_ids  = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb_xml)]
            existing_rids = [int(n) for n in re.findall(r'Id="rId(\d+)"', wb_rels_no_ws)]
            next_sheet_id = max(existing_ids,  default=1) + 1
            next_rid      = max(existing_rids, default=5) + 1

            # 샘플 .rels 에서 drawing 참조 파싱
            sample_drawing_rels: list[tuple[str, str, str]] = []  # (rId, type_suffix, old_target)
            if os.path.exists(sample_rels_path):
                rels_text = open(sample_rels_path, encoding='utf-8').read()
                for m in re.finditer(
                    r'<Relationship Id="([^"]+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]+)"',
                    rels_text
                ):
                    sample_drawing_rels.append((m.group(1), m.group(2), m.group(3)))

            new_sheets:      list[str] = []
            new_rels:        list[str] = []
            new_ct:          list[str] = []
            new_print_areas: list[str] = []

            for i, day_info in enumerate(day_data):
                day_str      = day_info['day_str']
                fname        = f'sheet_day{day_str}'
                new_xml_path = os.path.join(worksheets_dir, f'{fname}.xml')
                new_rel_path = os.path.join(ws_rels_dir,    f'{fname}.xml.rels')

                open(new_xml_path, 'w', encoding='utf-8').write(
                    self.build_sheet_xml(day_info, settings)
                )

                # 시트별 drawing 파일 복사 + 고유 .rels 생성
                # VML o:spid 고유화: 블록 번호 = i+1, spid = block*1024+1
                vml_block = i + 1
                new_spid  = f'_x0000_s{vml_block * 1024 + 1}'

                if sample_drawing_rels:
                    sheet_rels_entries: list[str] = []
                    for rid, rel_type, old_target in sample_drawing_rels:
                        # drawing / vmlDrawing 은 시트별 고유 복사본 생성
                        if 'drawing' in rel_type.lower() or 'vml' in old_target.lower():
                            # old_target: ../drawings/drawing1.xml 형태
                            old_basename = os.path.basename(old_target)
                            stem, ext    = os.path.splitext(old_basename)
                            new_basename = f'{stem}_{day_str}{ext}'
                            new_target   = f'../drawings/{new_basename}'

                            src = os.path.join(drawings_dir, old_basename)
                            dst = os.path.join(drawings_dir, new_basename)
                            if os.path.exists(src) and not os.path.exists(dst):
                                shutil.copy2(src, dst)
                                content = open(dst, encoding='utf-8').read()
                                if ext == '.vml':
                                    # o:idmap 블록 번호 + 모든 o:spid 교체
                                    content = re.sub(
                                        r'(<o:idmap[^>]*data=")[^"]*(")',
                                        rf'\g<1>{vml_block}\g<2>', content
                                    )
                                    content = re.sub(r'o:spid="[^"]*"', f'o:spid="{new_spid}"', content)
                                    # 결재란 카메라 소스 범위 갱신
                                    if (self._keoljairan_new_start is not None
                                            and self._kj_row_start is not None):
                                        kj_end = (self._keoljairan_new_start
                                                  + self._kj_row_end - self._kj_row_start)
                                        content = re.sub(
                                            r'<x:FmlaPict>[^<]*</x:FmlaPict>',
                                            f'<x:FmlaPict>$K${self._keoljairan_new_start}:$O${kj_end}</x:FmlaPict>',
                                            content
                                        )
                                elif ext == '.xml':
                                    # DrawingML cameraTool spid 교체
                                    content = re.sub(r'(spid=")[^"]*(")', rf'\g<1>{new_spid}\g<2>', content)
                                    # cNvPr id 고유화 (워크북 전체에서 고유해야 함)
                                    new_cnvpr_id = str(vml_block * 10 + 2)
                                    content = re.sub(r'(cNvPr id=")[^"]*(")', rf'\g<1>{new_cnvpr_id}\g<2>', content)
                                    # 결재란 카메라 소스 범위 갱신 (DrawingML cellRange)
                                    if (self._keoljairan_new_start is not None
                                            and self._kj_row_start is not None):
                                        kj_end = (self._keoljairan_new_start
                                                  + self._kj_row_end - self._kj_row_start)
                                        content = re.sub(
                                            r'(cellRange="\$[A-Z]+\$)\d+(:\$[A-Z]+\$)\d+(")',
                                            rf'\g<1>{self._keoljairan_new_start}\g<2>{kj_end}\g<3>',
                                            content
                                        )
                                open(dst, 'w', encoding='utf-8').write(content)

                            # drawing .rels 도 복사
                            src_drels = os.path.join(dr_rels_dir, f'{old_basename}.rels')
                            dst_drels = os.path.join(dr_rels_dir, f'{new_basename}.rels')
                            if os.path.exists(src_drels) and not os.path.exists(dst_drels):
                                shutil.copy2(src_drels, dst_drels)

                            # Content_Types 에 등록 (drawing xml 만 — vml 은 Default로 처리됨)
                            if ext == '.xml':
                                new_ct.append(
                                    f'<Override PartName="/xl/drawings/{new_basename}"'
                                    f' ContentType="application/vnd.openxmlformats-officedocument'
                                    f'.drawing+xml"/>'
                                )
                        elif 'printerSettings' in old_target:
                            # printerSettings 도 시트별 고유 복사본 생성 (OPC 소유권 위반 방지)
                            old_basename = os.path.basename(old_target)
                            stem, ext    = os.path.splitext(old_basename)
                            new_basename = f'{stem}_{day_str}{ext}'
                            new_target   = f'../printerSettings/{new_basename}'
                            ps_dir = os.path.join(tmp_dir, 'xl', 'printerSettings')
                            src = os.path.join(ps_dir, old_basename)
                            dst = os.path.join(ps_dir, new_basename)
                            if os.path.exists(src) and not os.path.exists(dst):
                                shutil.copy2(src, dst)
                        else:
                            new_target = old_target

                        sheet_rels_entries.append(
                            f'<Relationship Id="{rid}" Type="{rel_type}" Target="{new_target}"/>'
                        )

                    rels_xml = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        + ''.join(sheet_rels_entries)
                        + '</Relationships>'
                    )
                    open(new_rel_path, 'w', encoding='utf-8').write(rels_xml)

                wb_rid = f'rId{next_rid}'
                new_sheets.append(
                    f'<sheet name="{day_str}" sheetId="{next_sheet_id}" r:id="{wb_rid}"/>'
                )
                new_rels.append(
                    f'<Relationship Id="{wb_rid}"'
                    f' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
                    f' Target="worksheets/{fname}.xml"/>'
                )
                new_ct.append(
                    f'<Override PartName="/xl/worksheets/{fname}.xml"'
                    f' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                )
                # 인쇄 영역: localSheetId는 새로 추가되는 시트 순서 기반 0-indexed
                # 숫자로만 된 시트명은 따옴표 필요 (OOXML 규격)
                quoted_day = f"'{day_str}'" if day_str.isdigit() else day_str
                print_last_row = self._sum_row if self._sum_row else 33
                new_print_areas.append(
                    f'<definedName name="_xlnm.Print_Area" localSheetId="{i}">'
                    f'{quoted_day}!$A$1:$I${print_last_row}</definedName>'
                )
                next_sheet_id += 1
                next_rid      += 1

            # workbook.xml 수정
            # xr:revisionPtr 제거 — sample documentId가 그대로 남으면 손상으로 인식
            wb_xml = re.sub(r'<xr:revisionPtr[^/]*/>', '', wb_xml)
            wb_xml = re.sub(r'<sheet [^/]*/>', '', wb_xml)
            wb_xml = wb_xml.replace('<sheets>', '<sheets>' + ''.join(new_sheets))
            # 기존 definedNames 블록 제거 후 </sheets> 직후에 삽입
            # (OOXML workbook.xml 요소 순서: sheets → definedNames → calcPr → extLst)
            wb_xml = re.sub(r'<definedNames>.*?</definedNames>', '', wb_xml, flags=re.DOTALL)
            wb_xml = wb_xml.replace(
                '</sheets>',
                '</sheets><definedNames>' + ''.join(new_print_areas) + '</definedNames>'
            )
            open(wb_xml_path, 'w', encoding='utf-8').write(wb_xml)

            # workbook.xml.rels 수정
            wb_rels_text = wb_rels_no_ws.replace(
                '</Relationships>', ''.join(new_rels) + '</Relationships>'
            )
            wb_rels_text = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', '', wb_rels_text)
            open(wb_rels_path, 'w', encoding='utf-8').write(wb_rels_text)

            # calcChain.xml 삭제 (수식 재계산 강제)
            calc_path = os.path.join(tmp_dir, 'xl', 'calcChain.xml')
            if os.path.exists(calc_path):
                os.remove(calc_path)

            # Content_Types.xml 수정
            ct_xml = re.sub(r'<Override[^>]*worksheets/sheet\d+\.xml"[^>]*/>', '', ct_xml)
            ct_xml = re.sub(r'<Override[^>]*calcChain[^>]*/>', '', ct_xml)
            # 원본 drawing1.xml CT 항목 제거 (시트별 복사본으로 대체)
            ct_xml = re.sub(r'<Override[^>]*drawings/drawing\d+\.xml"[^>]*/>', '', ct_xml)
            ct_xml = ct_xml.replace('</Types>', ''.join(new_ct) + '</Types>')
            open(ct_path, 'w', encoding='utf-8').write(ct_xml)

            # 원본 샘플 시트 + 원본 drawing/printerSettings 삭제 (시트별 복사본으로 대체됨)
            for p in [sample_sheet_path, sample_rels_path]:
                if os.path.exists(p):
                    os.remove(p)
            ps_dir = os.path.join(tmp_dir, 'xl', 'printerSettings')
            for rid, rel_type, old_target in sample_drawing_rels:
                old_basename = os.path.basename(old_target)
                if 'drawing' in rel_type.lower() or 'vml' in old_target.lower():
                    for p in [
                        os.path.join(drawings_dir, old_basename),
                        os.path.join(dr_rels_dir,  f'{old_basename}.rels'),
                    ]:
                        if os.path.exists(p):
                            os.remove(p)
                elif 'printerSettings' in old_target:
                    p = os.path.join(ps_dir, old_basename)
                    if os.path.exists(p):
                        os.remove(p)

            # zip 재압축
            save_path = settings['output_file']
            if os.path.exists(save_path):
                os.remove(save_path)

            with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for root, _, files in os.walk(tmp_dir):
                    for f in files:
                        full = os.path.join(root, f)
                        zout.write(full, os.path.relpath(full, tmp_dir))

        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        return save_path
