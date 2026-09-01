import sys, os, datetime
import pandas as pd
try:
    import win32com.client as win32
    import pythoncom
    _HAS_WIN32 = True
except Exception:
    win32 = None
    pythoncom = None
    _HAS_WIN32 = False
from PyQt6.QtWidgets import *
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QPalette, QFont
from labor import run_migration
from work_status import fill_work_status

# ── 표준 로거 ─────────────────────────────────────────────────────
# 런처가 setup_logging을 이미 호출했으면 그 설정을 재사용,
# 단독 실행(개발) 시에는 get_logger가 자동 기본 설정한다.
try:
    from app_logging import get_logger, capture_exception, setup_logging, install_excepthook
except Exception:
    # app_logging이 없을 때도 최소 동작하도록 폴백
    import logging as _logging
    def get_logger(name="gui"): return _logging.getLogger(name)
    def capture_exception(exc=None, **ctx):
        _logging.getLogger("gui").error("예외 %s", ctx, exc_info=exc or True)
    def setup_logging(*a, **k): pass
    def install_excepthook(): pass

_flog = get_logger("gui")


def _app_dir() -> str:
    """작업 기준 폴더. 런처가 심어준 NOIM_BASE_DIR 우선."""
    env = os.environ.get("NOIM_BASE_DIR")
    if env:
        return env
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _sub_dir(name: str) -> str:
    """input/output/sample 하위 폴더 경로 반환 (없으면 생성)."""
    path = os.path.join(_app_dir(), name)
    os.makedirs(path, exist_ok=True)
    return path


# ── 고정 매핑 상수 (회사 프로그램 출력 형식 기준) ────────────────
COL_START_ROW = 7   # 데이터 시작행 (0-based df 인덱스)
COL_NAME      = 1   # 이름 열
COL_JUMIN     = 2   # 주민번호 열
COL_COST      = 21  # 단가 열
COL_DAY1      = 4   # 1~15일 시작 열
COL_DAY2      = 4   # 16~31일 시작 열


# ══════════════════════════════════════════════════════════════════
# 탭 1 — 노임일보 자동화
# ══════════════════════════════════════════════════════════════════

class Tab1Widget(QWidget):

    def __init__(self, log_fn):
        super().__init__()
        self.log = log_fn
        self.input_file  = ""
        self.sample_file = ""
        self.temp_files  = []
        self.df_raw      = None
        self._build_ui()

    # ── UI 구성 ───────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout()
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 12)

        # 1. 현장 정보
        info_group = QGroupBox("1. 현장 정보")
        i_lay = QGridLayout(); i_lay.setVerticalSpacing(3)

        _today = datetime.date.today()
        _prev = _today.replace(day=1) - datetime.timedelta(days=1)
        self.spin_year  = QSpinBox(); self.spin_year.setRange(2024, 2030); self.spin_year.setValue(_prev.year);  self.spin_year.setFixedHeight(24)
        self.spin_month = QSpinBox(); self.spin_month.setRange(1, 12);    self.spin_month.setValue(_prev.month); self.spin_month.setFixedHeight(24)
        self.edit_site  = QLineEdit("송도RC11"); self.edit_site.setFixedHeight(24)

        i_lay.addWidget(QLabel("년 / 월:"), 0, 0)
        i_lay.addWidget(self.spin_year,      0, 1)
        i_lay.addWidget(self.spin_month,     0, 2)
        i_lay.addWidget(QLabel("현장명:"),  1, 0)
        i_lay.addWidget(self.edit_site,      1, 1, 1, 2)

        self.chk_work_desc = QCheckBox("작업내용 열 채우기")
        self.chk_work_desc.setToolTip("체크 시 입력 파일의 지정 열 값을 작업내용(I열)에 채웁니다.")
        self.chk_work_desc.stateChanged.connect(self._on_work_desc_toggled)
        self.spin_work_desc_col = QSpinBox()
        self.spin_work_desc_col.setRange(0, 100)
        self.spin_work_desc_col.setValue(28)
        self.spin_work_desc_col.setFixedHeight(24)
        self.spin_work_desc_col.setEnabled(False)
        self.spin_work_desc_col.setToolTip("작업내용이 있는 열 번호 (0부터 시작)")
        i_lay.addWidget(self.chk_work_desc,      2, 0, 1, 2)
        i_lay.addWidget(self.spin_work_desc_col, 2, 2)

        info_group.setLayout(i_lay)
        root.addWidget(info_group)

        # 2. 파일 설정
        file_group = QGroupBox("2. 파일 설정")
        f_lay = QVBoxLayout(); f_lay.setSpacing(4)

        self.btn_load = QPushButton("📂  Raw 엑셀 파일 불러오기")
        self.btn_load.setFixedHeight(36)
        self.btn_load.clicked.connect(self.select_excel_file)
        f_lay.addWidget(self.btn_load)

        self.lbl_file = QLabel("파일을 선택하세요.")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setStyleSheet("color: #888; font-size: 11px; padding: 2px 4px;")
        f_lay.addWidget(self.lbl_file)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setToolTip("데이터가 담긴 시트를 선택하세요.")
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        f_lay.addWidget(self.sheet_combo)

        self.btn_load_sample = QPushButton("📂  샘플 파일 불러오기 (노임일보 서식)")
        self.btn_load_sample.setFixedHeight(30)
        self.btn_load_sample.clicked.connect(self._load_sample)
        f_lay.addWidget(self.btn_load_sample)

        self.lbl_sample = QLabel("파일을 선택하세요.")
        self.lbl_sample.setWordWrap(True)
        self.lbl_sample.setStyleSheet("color: #888; font-size: 11px; padding: 2px 4px;")
        f_lay.addWidget(self.lbl_sample)

        file_group.setLayout(f_lay)
        root.addWidget(file_group)

        # 3. 작업결과
        res_group = QGroupBox("3. 작업결과")
        r_lay = QVBoxLayout(); r_lay.setSpacing(4)

        self.lbl_summary = QLabel("실행 후 결과가 표시됩니다.")
        self.lbl_summary.setWordWrap(True)
        self.lbl_summary.setStyleSheet(
            "color: #888; font-size: 11px; padding: 6px;"
            "background: #fafafa; border: 1px solid #ddd; border-radius: 4px;"
        )
        r_lay.addWidget(self.lbl_summary)

        self.result_table = QTableWidget()
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["이름", "단가", "출근", "금액"])
        self.result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.result_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.result_table.verticalHeader().setVisible(False)
        self.result_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.result_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setFixedHeight(180)
        self.result_table.setStyleSheet("font-size: 11px;")
        self.result_table.setVisible(False)
        r_lay.addWidget(self.result_table)

        res_group.setLayout(r_lay)
        root.addWidget(res_group)

        # 실행 버튼
        self.btn_run = QPushButton("▶  노임일보 생성 시작")
        self.btn_run.setFixedHeight(52)
        self.btn_run.setStyleSheet(
            "background-color: #2c3e50; color: white; font-weight: bold;"
            "font-size: 13px; border-radius: 5px;"
        )
        self.btn_run.clicked.connect(self.execute)
        root.addWidget(self.btn_run)

        # 프로그레스바
        self.p_bar = QProgressBar(); self.p_bar.setFixedHeight(14)
        root.addWidget(self.p_bar)

        self.setLayout(root)

    def _on_work_desc_toggled(self, state: int):
        self.spin_work_desc_col.setEnabled(bool(state))

    # ── 파일 로드 ─────────────────────────────────────────────────

    def select_excel_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "파일 선택", _sub_dir("input"), "Excel (*.xlsx *.xlsm *.xls)")
        if not fname:
            return

        target_fname = fname

        if fname.lower().endswith('.xls'):
            self.log("구버전(.xls) 감지 → 변환 중...")
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            try:
                from xls_convert import convert_xls_to_xlsx
                target_fname = convert_xls_to_xlsx(fname)
                if target_fname not in self.temp_files:
                    self.temp_files.append(target_fname)
                self.log(f"✅ 변환 완료 → {os.path.basename(target_fname)}")
            except Exception as e:
                capture_exception(e, tab="tab1", step="xls_convert",
                                  file=os.path.basename(fname))
                QMessageBox.critical(self, "변환 실패", str(e))
                return
            finally:
                QApplication.restoreOverrideCursor()

        try:
            with pd.ExcelFile(target_fname) as xls:
                sheets = xls.sheet_names

            self.input_file = target_fname
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)
            self.sheet_combo.blockSignals(False)

            size_kb = os.path.getsize(target_fname) // 1024
            self.lbl_file.setText(
                f"📄 {os.path.basename(target_fname)}  ({size_kb} KB)\n"
                f"시트 {len(sheets)}개: {', '.join(sheets[:5])}"
                + (" ..." if len(sheets) > 5 else "")
            )
            self.lbl_file.setStyleSheet("color: #1a6e1a; font-size: 11px; padding: 2px 4px;")
            self.log(f"✅ 로드 성공: {os.path.basename(target_fname)}  |  시트 {len(sheets)}개")

            target_sheet = f"일용노무비지급명세서({self.spin_month.value():02d}월)_전체"
            if target_sheet not in sheets:
                QMessageBox.warning(self, "시트 불일치",
                    f"'{target_sheet}' 시트를 찾을 수 없습니다.\n\n"
                    f"첫 번째 시트 '{sheets[0]}'로 불러옵니다.\n"
                    f"월 설정 또는 시트 선택을 검토해 주세요.")
                target_sheet = sheets[0]
            self.sheet_combo.setCurrentText(target_sheet)
            self.load_sheet(target_sheet)

        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))

    def _load_sample(self):
        fname, _ = QFileDialog.getOpenFileName(self, "샘플 파일 선택", _sub_dir("sample"), "Excel (*.xlsx)")
        if not fname:
            return
        self.sample_file = fname
        self.lbl_sample.setText(f"📄 {os.path.basename(fname)}")
        self.lbl_sample.setStyleSheet("color: #1a6e1a; font-size: 11px; padding: 2px 4px;")
        self.log(f"샘플 파일 로드: {os.path.basename(fname)}")

    def on_sheet_changed(self):
        sheet = self.sheet_combo.currentText()
        if sheet and self.input_file:
            self.load_sheet(sheet)

    def load_sheet(self, sheet_name: str):
        try:
            self.df_raw = pd.read_excel(self.input_file, sheet_name=sheet_name, header=None)
            self.log(f"📋 시트 '{sheet_name}'  ({self.df_raw.shape[0]}행 × {self.df_raw.shape[1]}열)")
        except Exception as e:
            self.log(f"❌ 시트 로드 실패: {e}")
            self.df_raw = None

    # ── 작업결과 표시 ───────────────────────────────────────────────

    def _show_result(self, result: dict, settings: dict):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        self.lbl_summary.setText(
            f"📋 {settings['year']}년 {settings['month']:02d}월 {settings['site_name']} 노임일보\n"
            f"─────────────────────────────\n"
            f"■ 총 인원:  {result['total_workers']}명\n"
            f"■ 총 공수:  {result['total_mandays']:.1f}공\n"
            f"■ 총 금액:  {result['total_amount']:,.0f}원\n"
            f"■ 출력 파일: {os.path.basename(result['output_file'])}\n"
            f"■ 생성일시: {now}"
        )
        self.lbl_summary.setStyleSheet(
            "color: #1a3a1a; font-size: 11px; padding: 6px;"
            "background: #eaf5ea; border: 1px solid #b5d6b5; border-radius: 4px;"
        )

        details = result.get('workers_detail', [])
        self.result_table.setRowCount(len(details))
        for i, wd in enumerate(details):
            self.result_table.setItem(i, 0, QTableWidgetItem(wd['name']))
            self.result_table.setItem(i, 1, QTableWidgetItem(f"{wd['cost']:,}"))
            self.result_table.setItem(i, 2, QTableWidgetItem(f"{wd['days_worked']:.1f}"))
            self.result_table.setItem(i, 3, QTableWidgetItem(f"{wd['amount']:,.0f}"))
            self.result_table.setRowHeight(i, 24)
        self.result_table.setVisible(True)

    # ── 실행 ─────────────────────────────────────────────────────

    def execute(self):
        if not self.input_file:
            QMessageBox.warning(self, "파일 없음", "먼저 Raw 엑셀 파일을 불러오세요.")
            return
        if not self.sample_file:
            QMessageBox.warning(self, "파일 없음", "샘플 파일(노임일보 서식)을 불러오세요.")
            return

        _work_map, _work_fallback = [], '현장정리'
        if callable(getattr(self, 'get_work_desc_map', None)):
            _work_map, _work_fallback = self.get_work_desc_map()

        settings = {
            "input_file":     self.input_file,
            "input_sheet":    self.sheet_combo.currentText(),
            "header_row":     COL_START_ROW,
            "output_file":    os.path.join(_sub_dir("output"), f"{self.edit_site.text()}_직영일일출력일보_{str(self.spin_year.value())[2:]}.{self.spin_month.value():02d}.xlsx"),
            "year":           self.spin_year.value(),
            "month":          self.spin_month.value(),
            "site_name":      self.edit_site.text(),
            "name_col":       COL_NAME,
            "jumin_col":      COL_JUMIN,
            "cost_col":       COL_COST,
            "first_day_col1": COL_DAY1,
            "first_day_col2":   COL_DAY2,
            "limit_count":      15,
            "work_desc_enabled":  self.chk_work_desc.isChecked(),
            "work_desc_col":      self.spin_work_desc_col.value(),
            "work_desc_map":      _work_map,
            "work_desc_fallback": _work_fallback,
        }
        try:
            self.btn_run.setEnabled(False)
            self.p_bar.setValue(0)
            self.log(f"▶ 생성 시작: {settings['input_sheet']} / {settings['year']}-{settings['month']:02d}")
            result = run_migration(settings, self.sample_file, lambda v: self.p_bar.setValue(v))
            self._show_result(result, settings)
            self.log(f"✅ 완료 → {settings['output_file']}")
            QMessageBox.information(self, "완료", f"노임일보가 생성되었습니다.\n\n파일명: [{settings['output_file']}]")
        except Exception as e:
            capture_exception(e, tab="tab1", step="run_migration",
                              sheet=settings.get('input_sheet'),
                              year=settings.get('year'), month=settings.get('month'))
            self.log(f"❌ 오류: {e}")
            QMessageBox.critical(self, "에러", str(e))
        finally:
            self.btn_run.setEnabled(True)

    def cleanup_temp_files(self):
        for f in self.temp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except:
                pass


# ══════════════════════════════════════════════════════════════════
# 탭 2 — 작업현황 채우기
# ══════════════════════════════════════════════════════════════════

class Tab2Widget(QWidget):

    def __init__(self, log_fn):
        super().__init__()
        self.log = log_fn
        self.noim_file    = ""   # 노임일보 xlsx (입력)
        self.target_file  = ""   # 샘플 파일 원본 경로
        self.target_xlsx  = ""   # .xls→.xlsx 변환 경로 (임시)
        self.target_sheet = ""   # 자동 감지된 작업현황 시트명
        self.temp_files   = []
        self._build_ui()

    # ── UI 구성 ───────────────────────────────────────────────────

    def _build_ui(self):
        root = QVBoxLayout()
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 12)

        # 1. 년/월 선택
        g3 = QGroupBox("1. 대상 년/월")
        l3 = QHBoxLayout()
        _today = datetime.date.today()
        _prev = _today.replace(day=1) - datetime.timedelta(days=1)
        self.spin2_year  = QSpinBox(); self.spin2_year.setRange(2024, 2030); self.spin2_year.setValue(_prev.year);  self.spin2_year.setFixedHeight(24)
        self.spin2_month = QSpinBox(); self.spin2_month.setRange(1, 12);    self.spin2_month.setValue(_prev.month); self.spin2_month.setFixedHeight(24)
        l3.addWidget(QLabel("년:"))
        l3.addWidget(self.spin2_year)
        l3.addWidget(QLabel("월:"))
        l3.addWidget(self.spin2_month)
        l3.addStretch()
        g3.setLayout(l3)
        root.addWidget(g3)

        # 2. 노임일보 (입력)
        g1 = QGroupBox("2. 노임일보 파일 (생성된 출력물)")
        l1 = QVBoxLayout(); l1.setSpacing(4)

        self.btn_load_noim = QPushButton("📂  노임일보 xlsx 불러오기")
        self.btn_load_noim.setFixedHeight(34)
        self.btn_load_noim.clicked.connect(self._load_noim)
        l1.addWidget(self.btn_load_noim)

        self.lbl_noim = QLabel("파일을 선택하세요.")
        self.lbl_noim.setWordWrap(True)
        self.lbl_noim.setStyleSheet("color: #888; font-size: 11px; padding: 2px 4px;")
        l1.addWidget(self.lbl_noim)

        g1.setLayout(l1)
        root.addWidget(g1)

        # 3. 작업현황 샘플 파일 (출력 대상)
        g2 = QGroupBox("3. 샘플데이터 (작업현황 시트가 있는 파일)")
        l2 = QVBoxLayout(); l2.setSpacing(4)

        self.btn_load_target = QPushButton("📂  샘플데이터 불러오기")
        self.btn_load_target.setFixedHeight(34)
        self.btn_load_target.clicked.connect(self._load_target)
        l2.addWidget(self.btn_load_target)

        self.lbl_target = QLabel("파일을 선택하세요.")
        self.lbl_target.setWordWrap(True)
        self.lbl_target.setStyleSheet("color: #888; font-size: 11px; padding: 2px 4px;")
        l2.addWidget(self.lbl_target)

        g2.setLayout(l2)
        root.addWidget(g2)

        # 4. 작업내용 매핑 규칙
        g4 = QGroupBox("4. 작업내용 매핑 규칙  (노임일보 생성 탭 + 작업현황 분류에 적용)")
        l4 = QVBoxLayout(); l4.setSpacing(4)

        self.map_table = QTableWidget()
        self.map_table.setColumnCount(2)
        self.map_table.setHorizontalHeaderLabels(["키워드 (포함 시)", "치환값"])
        self.map_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.map_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.map_table.verticalHeader().setVisible(False)
        self.map_table.setFixedHeight(175)
        self.map_table.setStyleSheet("font-size: 11px;")

        # (키워드, 치환값, 카테고리키) — 카테고리키는 work_status.py 분류 열(신호수/
        # 안전자재/세륜기/동절기)과 1:1 대응. 탭2 실행 시 이 키워드로 분류 규칙을 재구성한다.
        _defaults = [
            ("신호수",   "3번게이트 신호수", "shin"),
            ("안전자재", "안전자재 정리",    "anjeon"),
            ("세륜기",   "세륜기 관리",      "selyun"),
            ("동절기",   "동절기",           "dong"),
        ]
        self.map_table.setRowCount(len(_defaults) + 1)
        for i, (kw, val, cat) in enumerate(_defaults):
            kw_item = QTableWidgetItem(kw)
            kw_item.setData(Qt.ItemDataRole.UserRole, cat)
            self.map_table.setItem(i, 0, kw_item)
            self.map_table.setItem(i, 1, QTableWidgetItem(val))

        # else 행 — 키워드 열은 편집 불가
        _else_kw = QTableWidgetItem("(기타 — else)")
        _else_kw.setFlags(_else_kw.flags() & ~Qt.ItemFlag.ItemIsEditable)
        _else_kw.setForeground(QColor("#999"))
        self.map_table.setItem(len(_defaults), 0, _else_kw)
        self.map_table.setItem(len(_defaults), 1, QTableWidgetItem("현장정리"))

        l4.addWidget(self.map_table)

        _btn_row = QHBoxLayout()
        _btn_add = QPushButton("+ 행 추가"); _btn_add.setFixedHeight(24)
        _btn_add.clicked.connect(self._add_map_row)
        _btn_del = QPushButton("- 행 삭제"); _btn_del.setFixedHeight(24)
        _btn_del.clicked.connect(self._del_map_row)
        _btn_row.addWidget(_btn_add)
        _btn_row.addWidget(_btn_del)
        _btn_row.addStretch()
        l4.addLayout(_btn_row)

        g4.setLayout(l4)
        root.addWidget(g4)

        # 실행 버튼
        self.btn_run2 = QPushButton("▶  작업현황 채우기 실행")
        self.btn_run2.setFixedHeight(52)
        self.btn_run2.setStyleSheet(
            "background-color: #1a4e6e; color: white; font-weight: bold;"
            "font-size: 13px; border-radius: 5px;"
        )
        self.btn_run2.clicked.connect(self._execute)
        root.addWidget(self.btn_run2)

        self.p_bar2 = QProgressBar(); self.p_bar2.setFixedHeight(14)
        root.addWidget(self.p_bar2)

        root.addStretch()
        self.setLayout(root)

    # ── 노임일보 로드 ─────────────────────────────────────────────

    def _load_noim(self):
        fname, _ = QFileDialog.getOpenFileName(self, "노임일보 파일 선택", _sub_dir("output"), "Excel (*.xlsx)")
        if not fname:
            return
        self.noim_file = fname
        size_kb = os.path.getsize(fname) // 1024
        self.lbl_noim.setText(f"📄 {os.path.basename(fname)}  ({size_kb} KB)")
        self.lbl_noim.setStyleSheet("color: #1a6e1a; font-size: 11px; padding: 2px 4px;")
        self.log(f"[탭2] 노임일보 로드: {os.path.basename(fname)}")

    # ── 원본 파일 로드 ────────────────────────────────────────────

    def _load_target(self):
        fname, _ = QFileDialog.getOpenFileName(self, "원본 파일 선택", _sub_dir("sample"), "Excel (*.xlsx *.xlsm *.xls)")
        if not fname:
            return

        target = fname
        if fname.lower().endswith('.xls'):
            self.log("[탭2] 구버전(.xls) 감지 → 변환 중...")
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            try:
                from xls_convert import convert_xls_to_xlsx
                target = convert_xls_to_xlsx(fname)
                if target not in self.temp_files:
                    self.temp_files.append(target)
                self.log(f"[탭2] 변환 완료 → {os.path.basename(target)}")
            except Exception as e:
                capture_exception(e, tab="tab2", step="xls_convert",
                                  file=os.path.basename(fname))
                QMessageBox.critical(self, "변환 실패", str(e))
                return
            finally:
                QApplication.restoreOverrideCursor()

        try:
            with pd.ExcelFile(target) as xls:
                sheets = xls.sheet_names
            status_sheet = next((s for s in sheets if '작업현황' in s), None)
            if not status_sheet:
                QMessageBox.critical(self, "시트 없음",
                    "파일에서 '작업현황' 시트를 찾을 수 없습니다.")
                return
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
            return

        self.target_file   = fname
        self.target_xlsx   = target
        self.target_sheet  = status_sheet
        size_kb = os.path.getsize(target) // 1024
        self.lbl_target.setText(f"📄 {os.path.basename(target)}  ({size_kb} KB)  →  시트: {status_sheet}")
        self.lbl_target.setStyleSheet("color: #1a6e1a; font-size: 11px; padding: 2px 4px;")
        self.log(f"[탭2] 샘플데이터 로드: {os.path.basename(target)}  |  시트: {status_sheet}")

    # ── 실행 ─────────────────────────────────────────────────────

    def _execute(self):
        import shutil
        if not self.noim_file:
            QMessageBox.warning(self, "파일 없음", "노임일보 파일을 먼저 불러오세요.")
            return
        if not self.target_xlsx:
            QMessageBox.warning(self, "파일 없음", "샘플데이터 파일을 먼저 불러오세요.")
            return

        year  = self.spin2_year.value()
        month = self.spin2_month.value()

        output_dir  = _sub_dir("output")
        output_name = f"작업현황_{year}년_{month:02d}월.xlsx"
        output_path = os.path.join(output_dir, output_name)

        try:
            self.btn_run2.setEnabled(False)
            self.p_bar2.setValue(0)
            self.log(f"[탭2] ▶ 시작: {year}-{month:02d} / 시트: {self.target_sheet}")

            # 1. 샘플 → output 복사 (샘플 원본 보존)
            shutil.copy(self.target_xlsx, output_path)
            self.p_bar2.setValue(20)

            # 2. output 파일에 데이터 채우기
            fill_work_status(self.noim_file, output_path, self.target_sheet, year, month,
                              category_keywords=self.get_category_keywords())
            self.p_bar2.setValue(100)

            self.log(f"[탭2] ✅ 완료 → {output_name}")
            QMessageBox.information(self, "완료",
                f"작업현황 시트 채우기가 완료되었습니다.\n\n파일명: [{output_name}]")
        except Exception as e:
            capture_exception(e, tab="tab2", step="fill_work_status",
                              sheet=self.target_sheet, year=year, month=month)
            self.p_bar2.setValue(0)
            self.log(f"[탭2] ❌ 오류: {e}")
            QMessageBox.critical(self, "에러", str(e))
        finally:
            self.btn_run2.setEnabled(True)

    def _add_map_row(self):
        last = self.map_table.rowCount() - 1   # else 행 앞에 삽입
        self.map_table.insertRow(last)
        self.map_table.setItem(last, 0, QTableWidgetItem(""))
        self.map_table.setItem(last, 1, QTableWidgetItem(""))

    def _del_map_row(self):
        sel = self.map_table.currentRow()
        last = self.map_table.rowCount() - 1
        if sel < 0 or sel >= last:             # else 행은 삭제 불가
            return
        self.map_table.removeRow(sel)

    def get_work_desc_map(self):
        """Tab 1 execute()에서 호출 — (mapping_list, fallback_str) 반환."""
        mapping = []
        last = self.map_table.rowCount() - 1
        for i in range(last):
            kw_item  = self.map_table.item(i, 0)
            val_item = self.map_table.item(i, 1)
            kw  = kw_item.text().strip()  if kw_item  else ''
            val = val_item.text().strip() if val_item else ''
            if kw:
                mapping.append({'keyword': kw, 'value': val})
        fallback_item = self.map_table.item(last, 1)
        fallback = fallback_item.text().strip() if fallback_item else '현장정리'
        return mapping, fallback

    def get_category_keywords(self):
        """탭2 실행() 호출 — work_status.py 분류 규칙에 쓰일 {category: keyword} 반환."""
        result = {}
        last = self.map_table.rowCount() - 1
        for i in range(last):
            kw_item = self.map_table.item(i, 0)
            if not kw_item:
                continue
            cat = kw_item.data(Qt.ItemDataRole.UserRole)
            if cat:
                result[cat] = kw_item.text().strip()
        return result

    def _save_back_as_xls(self, src_xlsx: str, dst_xls: str):
        """편집된 xlsx를 COM으로 열어 원본 xls 경로에 덮어씀."""
        if not _HAS_WIN32:
            raise RuntimeError(
                "이 PC에는 Excel(win32com)이 없어 .xls로 재저장할 수 없습니다. "
                ".xlsx 형식으로 사용하세요."
            )
        excel = wb = None
        try:
            pythoncom.CoInitialize()
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(os.path.abspath(src_xlsx))
            wb.SaveAs(os.path.abspath(dst_xls), FileFormat=56)  # 56 = xls
        finally:
            if wb:    wb.Close(SaveChanges=False)
            if excel: excel.Quit()
            pythoncom.CoUninitialize()

    def cleanup_temp_files(self):
        for f in self.temp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except:
                pass


# ══════════════════════════════════════════════════════════════════
# 메인 윈도우
# ══════════════════════════════════════════════════════════════════

class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("현장 노임일보 자동화 시스템")
        self.resize(460, 640)

        root = QVBoxLayout()
        root.setSpacing(4)
        root.setContentsMargins(0, 0, 0, 0)

        # 탭
        self.tabs = QTabWidget()
        self.tab1 = Tab1Widget(self.log)
        self.tab2 = Tab2Widget(self.log)
        self.tab1.get_work_desc_map = self.tab2.get_work_desc_map  # 매핑 공유
        self.tabs.addTab(self.tab1, "노임일보 생성")
        self.tabs.addTab(self.tab2, "작업현황 작성")
        root.addWidget(self.tabs)

        # 공용 로그창
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFixedHeight(100)
        self.log_box.setStyleSheet("font-size: 11px; background: #fafafa;")
        root.addWidget(self.log_box)

        self.setLayout(root)

    def log(self, msg: str):
        # 화면 출력
        self.log_box.append(msg)
        self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())
        # 파일 기록 (휘발 방지) — 이모지/장식 제거 없이 그대로 정규화 로그에 남김
        try:
            _flog.info(msg)
        except Exception:
            pass

    def closeEvent(self, event):
        self.tab1.cleanup_temp_files()
        super().closeEvent(event)


def run():
    # 런처를 거치지 않고 단독 실행되는 경우에도 로깅 보장
    try:
        setup_logging(component="gui")
        install_excepthook()
    except Exception:
        pass

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(245, 245, 245))
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.black)
    app.setPalette(palette)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    run()