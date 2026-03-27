import sys
import os
import threading
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QMessageBox, QGroupBox, QComboBox, QCheckBox,
    QPlainTextEdit, QProgressBar, QSystemTrayIcon, QMenu, QDialog,
    QCalendarWidget, QSpinBox, QTimeEdit, QDateEdit, QDateTimeEdit,
    QSplitter, QFrame, QLineEdit, QScrollArea, QAbstractItemView, QTextEdit, # [추가] QTextEdit
    QSizePolicy, QGridLayout # [추가] QSizePolicy, QGridLayout
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QTime, QDate, QDateTime, QSize, QObject, QRect # [추가] QRect
from PySide6.QtGui import QIcon, QFont, QAction, QColor, QDesktopServices, QIntValidator, QPalette, QPainter, QPen # [추가] QPen
from sheet_manager import GoogleSheetManager
from navercafe_auto import NaverCafeBot

class TaskCountCalendar(QCalendarWidget):
    """날짜별 예약 건수를 표시하는 커스텀 달력 위젯"""
    def __init__(self, tasks, parent=None):
        super().__init__(parent)
        self.tasks = tasks
        self.task_counts = self._calculate_counts()

    def _calculate_counts(self):
        counts = {}
        for t in self.tasks:
            nr = t.get('next_run', '')
            if nr:
                date_str = nr.split()[0] if " " in nr else nr
                counts[date_str] = counts.get(date_str, 0) + 1
        return counts

    def paintCell(self, painter, rect, date):
        is_selected = (date == self.selectedDate())
        
        if is_selected:
            painter.save()
            # 1. 배경 채우기 (선명한 옅은 파란색)
            painter.setBrush(QColor("#a0c8f0")) # 더 선명한 파란색
            # [추가] 테두리도 그려서 더 명확하게 표시
            painter.setPen(QPen(QColor("#3a86ff"), 2)) # 진한 파란색 테두리
            painter.drawRect(rect.adjusted(1, 1, -1, -1))
            
            # 2. 날짜 텍스트 (검은색으로 명확하게)
            painter.setPen(Qt.black)
            font = painter.font()
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(rect, Qt.AlignCenter, str(date.day()))
            painter.restore()
        else:
            super().paintCell(painter, rect, date)
            
        # 3. 태스크 건수 표시
        date_str = date.toString("yyyy-MM-dd")
        count = self.task_counts.get(date_str, 0)
        
        if count > 0:
            painter.save()
            painter.setPen(QColor("#ffffff"))
            font = painter.font()
            font.setPointSize(8)
            font.setBold(True)
            painter.setFont(font)
            
            count_text = f"{count}건"
            bg_rect = QRect(rect.right() - 25, rect.top() + 2, 22, 14)
            painter.setBrush(QColor("#67c23a"))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(bg_rect, 4, 4)
            
            painter.setPen(QColor("#ffffff"))
            painter.drawText(bg_rect, Qt.AlignCenter, count_text)
            painter.restore()

class DatePickerDialog(QDialog):
    def __init__(self, current_text="", tasks=None, parent=None):
        super().__init__(parent)
        self.tasks = tasks if tasks else []
        self.setWindowTitle("첫 실행 일시 설정")
        self.resize(600, 400)
        self.setStyleSheet("background-color: white;")
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        lbl_info = QLabel("게시글을 처음으로 실행할 날짜와 시간을 선택하세요.")
        lbl_info.setStyleSheet("font-size: 14pt; font-weight: bold; color: #333; margin-bottom: 20px;")
        lbl_info.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(lbl_info)
        
        # 컨텐츠 영역 (달력 + 시간)
        content_layout = QHBoxLayout()
        content_layout.setSpacing(30)
        
        # 1. 왼쪽: 달력 (커스텀 스타일링 강화)
        self.calendar = TaskCountCalendar(self.tasks)
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        
        # [수정] Palette 강제 적용 (시스템 테마 오버라이드)
        p = self.calendar.palette()
        p.setColor(QPalette.Highlight, QColor("#add8e6")) # 옅은 파란색 (선택 시)
        p.setColor(QPalette.HighlightedText, Qt.black)
        p.setColor(QPalette.Inactive, QPalette.Highlight, QColor("#add8e6")) # 비활성 상태에서도 색상 유지
        p.setColor(QPalette.Inactive, QPalette.HighlightedText, Qt.black)
        self.calendar.setPalette(p)
        
        self.calendar.setStyleSheet("""
            QCalendarWidget QAbstractItemView {
                selection-background-color: #a0c8f0;
                selection-color: black;
                background-color: white;
                outline: none;
            }
            QCalendarWidget QAbstractItemView:disabled {
                selection-background-color: #a0c8f0;
                selection-color: black;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar { 
                background-color: #409eff; 
                min-height: 40px; 
            }
            QCalendarWidget QToolButton {
                color: white;
                background-color: transparent;
                border: none;
                font-weight: bold;
                font-size: 14px;
                margin: 5px;
                padding: 5px;
            }
            QCalendarWidget QToolButton:hover {
                background-color: rgba(255, 255, 255, 0.2);
                border-radius: 4px;
            }
            QCalendarWidget QToolButton#qt_calendar_monthbutton {
                padding-right: 20px;
                width: 100px;
            }
            QCalendarWidget QToolButton#qt_calendar_monthbutton::menu-indicator {
                subcontrol-position: center right;
                image: none; 
                width: 10px;
            }
            QCalendarWidget QSpinBox {
                width: 50px;
                font-size: 14px;
                background-color: white;
                selection-background-color: #409eff;
            }
        """)
        # [추가] 선택 변경 시 달력 강제 업데이트 (그리기 갱신)
        self.calendar.selectionChanged.connect(self.calendar.update)
        content_layout.addWidget(self.calendar)

        # 2. 시간 설정 (우측 패널)
        grp_time = QFrame() # Changed to QFrame
        grp_time.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #ebeef5;
                border-radius: 8px;
            }
        """)
        vbox_time = QVBoxLayout(grp_time)
        vbox_time.setContentsMargins(15, 15, 15, 15) # Adjusted margins
        vbox_time.setSpacing(15) # Adjusted spacing

        lbl_title = QLabel("시간 설정")
        lbl_title.setStyleSheet("font-size: 12pt; color: #606266; border: none;") # Adjusted style
        lbl_title.setAlignment(Qt.AlignCenter)
        vbox_time.addWidget(lbl_title)

        vbox_time.addStretch()

        # 시간 (Hour)
        hbox_h = QHBoxLayout()
        self.btn_h_minus = QPushButton("－")
        self.hour_edit = QLineEdit("12") # Replaced QLabel with QLineEdit
        self.hour_edit.setValidator(QIntValidator(0, 23)) # Added validator
        self.hour_edit.setAlignment(Qt.AlignCenter)
        self.hour_edit.setStyleSheet("font-size: 24pt; font-weight: bold; color: #303133; min-width: 60px; background: white; border: 1px solid #dcdfe6; border-radius: 4px;")
        self.btn_h_plus = QPushButton("＋")

        self.setup_time_widgets(self.btn_h_minus, self.hour_edit, self.btn_h_plus)
        hbox_h.addWidget(self.btn_h_minus)
        hbox_h.addWidget(self.hour_edit)
        hbox_h.addWidget(self.btn_h_plus)

        # 분 (Minute)
        hbox_m = QHBoxLayout()
        self.btn_m_minus = QPushButton("－")
        self.min_edit = QLineEdit("00") # Replaced QLabel with QLineEdit
        self.min_edit.setValidator(QIntValidator(0, 59)) # Added validator
        self.min_edit.setAlignment(Qt.AlignCenter)
        self.min_edit.setStyleSheet("font-size: 24pt; font-weight: bold; color: #303133; min-width: 60px; background: white; border: 1px solid #dcdfe6; border-radius: 4px;")
        self.btn_m_plus = QPushButton("＋")

        self.setup_time_widgets(self.btn_m_minus, self.min_edit, self.btn_m_plus)
        hbox_m.addWidget(self.btn_m_minus)
        hbox_m.addWidget(self.min_edit)
        hbox_m.addWidget(self.btn_m_plus)

        vbox_time.addLayout(self.create_labeled_layout("시 (Hour)", hbox_h))
        vbox_time.addSpacing(10) # Added spacing
        vbox_time.addLayout(self.create_labeled_layout("분 (Min)", hbox_m))
        vbox_time.addStretch()

        content_layout.addWidget(grp_time)
        content_layout.setStretch(0, 3) # 달력 넓게
        content_layout.setStretch(1, 2) # 시간 좁게

        main_layout.addLayout(content_layout)

        # 하단: 선택 정보 및 버튼
        bottom_layout = QVBoxLayout()

        self.lbl_selected_display = QLabel("")
        self.lbl_selected_display.setStyleSheet("""
            background-color: #ecf5ff; border: 1px solid #d9ecff; border-radius: 4px;
            color: #409eff; font-size: 13pt; font-weight: bold; padding: 10px;
        """)
        self.lbl_selected_display.setAlignment(Qt.AlignCenter)
        bottom_layout.addWidget(self.lbl_selected_display)
        bottom_layout.addSpacing(10)

        btn_box = QHBoxLayout()
        btn_box.addStretch()

        self.btn_ok = QPushButton("설정 완료 (OK)")
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setStyleSheet("""
            padding: 12px 30px; font-size: 11pt; font-weight: bold; color: white;
            background-color: #409eff; border-radius: 4px; border: none;
        """)
        self.btn_ok.clicked.connect(self.accept)

        self.btn_cancel = QPushButton("취소")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.setStyleSheet("""
            padding: 12px 20px; font-size: 11pt; color: #606266;
            background-color: white; border: 1px solid #dcdfe6; border-radius: 4px;
        """)
        self.btn_cancel.clicked.connect(self.reject)

        btn_box.addWidget(self.btn_cancel)
        btn_box.addWidget(self.btn_ok)
        bottom_layout.addLayout(btn_box)

        main_layout.addLayout(bottom_layout)

        # 로직 연결
        self.cur_hour = 12
        self.cur_min = 0

        self.btn_h_minus.clicked.connect(lambda: self.adjust_time('h', -1))
        self.btn_h_plus.clicked.connect(lambda: self.adjust_time('h', 1))
        self.btn_m_minus.clicked.connect(lambda: self.adjust_time('m', -1)) # 1분 단위
        self.btn_m_plus.clicked.connect(lambda: self.adjust_time('m', 1))

        # [추가] 텍스트 입력 시 업데이트 연결
        self.hour_edit.editingFinished.connect(lambda: self.update_from_input('h'))
        self.min_edit.editingFinished.connect(lambda: self.update_from_input('m'))

        self.calendar.selectionChanged.connect(self.update_display)

        # 초기값 로드
        self.load_initial_datetime(current_text)
        self.update_display()

    def setup_time_widgets(self, btn_minus, widget, btn_plus): # Changed lbl to widget
        btn_style = """
            QPushButton {
                background-color: #f0f2f5; border: 1px solid #dcdfe6; border-radius: 4px;
                font-size: 16px; font-weight: bold; min-width: 40px; min-height: 40px;
            }
            QPushButton:hover { background-color: #e6e8eb; color: #409eff; }
            QPushButton:pressed { background-color: #dcdfe6; }
        """
        btn_minus.setStyleSheet(btn_style)
        btn_plus.setStyleSheet(btn_style)

        # widget style already set in init

    def create_labeled_layout(self, title, control_layout):
        v = QVBoxLayout()
        l = QLabel(title)
        l.setStyleSheet("color: #909399; font-size: 10pt; font-weight: bold;")
        l.setAlignment(Qt.AlignCenter)
        v.addWidget(l)
        v.addLayout(control_layout)
        return v

    def adjust_time(self, type_, delta):
        if type_ == 'h':
            self.cur_hour = (self.cur_hour + delta) % 24
            self.hour_edit.setText(f"{self.cur_hour:02d}") # Updated to hour_edit
        else:
            self.cur_min = (self.cur_min + delta) % 60
            self.min_edit.setText(f"{self.cur_min:02d}") # Updated to min_edit
        self.update_display()

    def update_from_input(self, type_):
        if type_ == 'h':
            try:
                val = int(self.hour_edit.text())
                if 0 <= val <= 23:
                    self.cur_hour = val
                else:
                    self.hour_edit.setText(f"{self.cur_hour:02d}") # Revert if invalid
            except ValueError:
                self.hour_edit.setText(f"{self.cur_hour:02d}") # Revert if not a number
        else:
            try:
                val = int(self.min_edit.text())
                if 0 <= val <= 59:
                    self.cur_min = val
                else:
                    self.min_edit.setText(f"{self.cur_min:02d}") # Revert if invalid
            except ValueError:
                self.min_edit.setText(f"{self.cur_min:02d}") # Revert if not a number
        self.update_display()

    def load_initial_datetime(self, text):
        dt = QDateTime.currentDateTime()
        if text:
            try:
                if len(text) > 10:
                    dt = QDateTime.fromString(text, "yyyy-MM-dd HH:mm")
                else:
                    dt = QDateTime.fromString(text, "yyyy-MM-dd")
            except:
                pass

        self.calendar.setSelectedDate(dt.date())
        self.cur_hour = dt.time().hour()
        self.cur_min = dt.time().minute()

        self.hour_edit.setText(f"{self.cur_hour:02d}") # Updated to hour_edit
        self.min_edit.setText(f"{self.cur_min:02d}") # Updated to min_edit

    def update_display(self):
        d_str = self.calendar.selectedDate().toString("yyyy-MM-dd")
        t_str = f"{self.cur_hour:02d}:{self.cur_min:02d}"
        self.lbl_selected_display.setText(f"선택된 일시: {d_str} {t_str}")

    def get_datetime_str(self):
        d_str = self.calendar.selectedDate().toString("yyyy-MM-dd")
        t_str = f"{self.cur_hour:02d}:{self.cur_min:02d}"
        return f"{d_str} {t_str}"

class CalendarDialog(QDialog):
    def __init__(self, tasks, parent=None):
        super().__init__(parent)
        self.tasks = tasks
        self.setWindowTitle("일정 달력 (Calendar)")
        self.setFixedSize(700, 650)
        self.setStyleSheet("background-color: white;")
        
        self.current_date = QDate.currentDate()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Header (Month/Year with Prev/Next buttons)
        header_layout = QHBoxLayout()
        btn_prev = QPushButton("◀ 이전달")
        btn_prev.setCursor(Qt.PointingHandCursor)
        btn_prev.setFixedSize(100, 40)
        btn_prev.setStyleSheet("background-color: #f0f2f5; border: 1px solid #dcdfe6; border-radius: 4px; font-weight: bold;")
        btn_prev.clicked.connect(self.prev_month)
        
        self.lbl_month = QLabel()
        self.lbl_month.setAlignment(Qt.AlignCenter)
        self.lbl_month.setStyleSheet("font-size: 20pt; font-weight: bold; color: #303133;")
        
        btn_next = QPushButton("다음달 ▶")
        btn_next.setCursor(Qt.PointingHandCursor)
        btn_next.setFixedSize(100, 40)
        btn_next.setStyleSheet("background-color: #f0f2f5; border: 1px solid #dcdfe6; border-radius: 4px; font-weight: bold;")
        btn_next.clicked.connect(self.next_month)
        
        today_btn = QPushButton("오늘")
        today_btn.setCursor(Qt.PointingHandCursor)
        today_btn.setFixedSize(60, 40)
        today_btn.setStyleSheet("background-color: #e6f7ff; color: #1890ff; border: 1px solid #91d5ff; border-radius: 4px; font-weight: bold;")
        today_btn.clicked.connect(self.go_today)
        
        header_layout.addWidget(btn_prev)
        header_layout.addStretch()
        header_layout.addWidget(self.lbl_month)
        header_layout.addStretch()
        header_layout.addWidget(today_btn)
        header_layout.addWidget(btn_next)
        
        layout.addLayout(header_layout)
        
        # Calendar Grid
        self.grid_layout = QGridLayout()
        self.grid_layout.setSpacing(5)
        layout.addLayout(self.grid_layout, 1)
        
        self.update_calendar()
        
    def prev_month(self):
        self.current_date = self.current_date.addMonths(-1)
        self.update_calendar()
        
    def next_month(self):
        self.current_date = self.current_date.addMonths(1)
        self.update_calendar()
        
    def go_today(self):
        self.current_date = QDate.currentDate()
        self.update_calendar()
        
    def update_calendar(self):
        # Clear old grid
        for i in reversed(range(self.grid_layout.count())): 
            widget = self.grid_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
                
        self.lbl_month.setText(self.current_date.toString("yyyy년 M월"))
        
        days_of_week = ["일", "월", "화", "수", "목", "금", "토"]
        for i, day in enumerate(days_of_week):
            lbl = QLabel(day)
            lbl.setAlignment(Qt.AlignCenter)
            style = "font-weight: bold; font-size: 13pt; padding: 5px; background-color: #f8f9fa; border-top: 2px solid #e4e7ed; border-bottom: 2px solid #e4e7ed;"
            if i == 0: style += " color: #f56c6c;"
            elif i == 6: style += " color: #409eff;"
            else: style += " color: #606266;"
            lbl.setStyleSheet(style)
            self.grid_layout.addWidget(lbl, 0, i)
            
        # Group tasks by date
        task_dict = {}
        for t in self.tasks:
            nr = t.get('next_run', '')
            if not nr: continue
            
            # extract "YYYY-MM-DD"
            date_str = nr.split()[0] if " " in nr else nr
            if date_str not in task_dict:
                task_dict[date_str] = []
            task_dict[date_str].append(t)
            
        first_day = QDate(self.current_date.year(), self.current_date.month(), 1)
        first_day_of_week = int(first_day.dayOfWeek()) % 7 # 0 is Sunday
        
        row: int = 1
        col: int = first_day_of_week
        days_in_month = int(self.current_date.daysInMonth())
        
        # Draw blanks before first day
        for c in range(col):
            b_frame = QFrame()
            b_frame.setStyleSheet("background-color: #fafbfc; border: 1px solid #ebeef5; border-radius: 4px;")
            self.grid_layout.addWidget(b_frame, row, c)
            
        # Draw days
        for day in range(1, days_in_month + 1):
            cell_frame = QFrame()
            cell_frame.setFrameShape(QFrame.StyledPanel)
            cell_frame.setStyleSheet("""
                QFrame { background-color: white; border: 1px solid #ebeef5; border-radius: 4px; }
                QFrame:hover { background-color: #f0f9eb; border: 1px solid #c2e7b0; }
            """)
            
            cell_layout = QVBoxLayout(cell_frame)
            cell_layout.setContentsMargins(5, 5, 5, 5)
            cell_layout.setSpacing(2)
            
            # 날짜 라벨
            lbl_day = QLabel(str(day))
            lbl_day.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            day_style = "font-size: 13pt; font-weight: bold; border: none; background: transparent;"
            
            if self.current_date.year() == QDate.currentDate().year() and self.current_date.month() == QDate.currentDate().month() and day == QDate.currentDate().day():
                # Today highlight
                day_style += " color: white; background-color: #409eff; border-radius: 12px; min-width: 24px; max-width: 24px; min-height: 24px; max-height: 24px;"
                lbl_day.setAlignment(Qt.AlignCenter)
            elif col == 0: day_style += " color: #f56c6c;"
            elif col == 6: day_style += " color: #409eff;"
            else: day_style += " color: #606266;"
            
            lbl_day.setStyleSheet(day_style)
            cell_layout.addWidget(lbl_day, 0, Qt.AlignLeft | Qt.AlignTop)
            
            date_str = f"{self.current_date.year()}-{self.current_date.month():02d}-{day:02d}"
            day_tasks = task_dict.get(date_str, [])
            
            # 태스크 건수 마커
            if day_tasks:
                cnt_lbl = QLabel(f"{len(day_tasks)} 건")
                cnt_lbl.setAlignment(Qt.AlignCenter)
                cnt_lbl.setStyleSheet("color: white; background-color: #67c23a; font-weight: bold; font-size: 11pt; border-radius: 4px; padding: 3px;")
                cell_layout.addWidget(cnt_lbl, 0, Qt.AlignCenter)
                
                # 툴팁 구성 (이름 / 카페명 / 주차)
                tooltip_lines = [f"【 {date_str} 예약 현황 】"]
                for t in day_tasks:
                    stage = t.get('stage_name', "")
                    remain = t.get('remain_count', t.get('upload_count', ''))
                    disp = stage if stage else f"{remain} 남음"
                    tooltip_lines.append(f"• {t['name']} / {t['cafe_name']} / {disp}")
                cell_frame.setToolTip("\n".join(tooltip_lines))
                # Add hover hand
                cell_frame.setCursor(Qt.PointingHandCursor)
            else:
                cell_layout.addStretch()
                
            self.grid_layout.addWidget(cell_frame, row, col)
            
            col += 1
            if col > 6:
                col = 0
                row += 1
                
        # Fill remaining with blanks
        while row < 7:
            if col > 6:
                col = 0
                row += 1
                if row > 6: break
            
            dummy_frame = QFrame()
            dummy_frame.setStyleSheet("background-color: #fafbfc; border: 1px solid #ebeef5; border-radius: 4px;")
            self.grid_layout.addWidget(dummy_frame, row, col)
            col += 1
            
        for r in range(1, 7):
            self.grid_layout.setRowStretch(r, 1)
        for c in range(7):
            self.grid_layout.setColumnStretch(c, 1)

class WorkerSignals(QObject):
    log = Signal(str)
    finished = Signal()

class TaskLoaderThread(QThread):
    tasksLoaded = Signal(object) # list or None
    errorOccurred = Signal(str)

    def __init__(self, sheet_mgr):
        super().__init__()
        self.sheet_mgr = sheet_mgr

    def run(self):
        try:
            tasks = self.sheet_mgr.get_tasks()
            self.tasksLoaded.emit(tasks)
        except Exception as e:
            self.errorOccurred.emit(str(e))


__version__ = "1.1.4"

class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # App Info
        self.setWindowTitle(f'네이버 카페 자동 포스팅 (Naver Cafe Auto) v{__version__}')
        self.resize(1500, 850) # 넓이 조정 (1300 -> 1500)
        
        self.sheet_mgr = GoogleSheetManager()
        self.tasks = []
        self.task_layouts = {} # 작업별 컨텐츠 순서 저장
        self.default_order = ["before_img", "after_img", "body"]
        
        self.init_ui()
        self.apply_stylesheet()
        
        # 상태바에 버전 및 상태 표시
        self.statusBar().showMessage(f"버전: {__version__} - 대기 중")
        self.statusBar().setStyleSheet("color: #606266; font-weight: bold;")
        
        # 프로그램 시작 시 자동으로 시트 불러오기 (0.1초 후 실행)
        QTimer.singleShot(100, self.load_tasks)
        self.first_load_completed = False
        
        
    def apply_stylesheet(self):
        # 모던한 스타일시트 적용
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', 'Malgun Gothic', sans-serif;
                font-size: 10pt;
                background-color: #f5f7fa;
                color: #333333;
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #dcdfe6;
                border-radius: 4px;
                padding: 8px 16px;
                color: #606266;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ecf5ff;
                color: #409eff;
                border: 1px solid #c6e2ff;
            }
            QPushButton:pressed {
                background-color: #ecf5ff;
                border: 1px solid #b3d8ff;
            }
            QPushButton:disabled {
                background-color: #f5f7fa;
                border: 1px solid #e4e7ed;
                color: #c0c4cc;
            }
            
            /* 특정 버튼 스타일링 */
            QPushButton#btn_load {
                background-color: #409eff;
                border: 1px solid #409eff;
                color: white;
            }
            QPushButton#btn_load:hover {
                background-color: #66b1ff;
                border: 1px solid #66b1ff;
            }
            
            QPushButton#btn_start {
                background-color: #67c23a;
                border: 1px solid #67c23a;
                color: white;
            }
            QPushButton#btn_start:hover {
                background-color: #85ce61;
                border: 1px solid #85ce61;
            }
            
            QPushButton#btn_cancel {
                background-color: #f56c6c;
                border: 1px solid #f56c6c;
                color: white;
            }
            QPushButton#btn_cancel:hover {
                background-color: #f78989;
                border: 1px solid #f78989;
            }

            QTableWidget {
                background-color: white;
                border: 1px solid #e4e7ed;
                border-radius: 4px;
                gridline-color: #ebeef5;
                selection-background-color: #ecf5ff;
                selection-color: #606266;
            }
            QHeaderView::section {
                background-color: #f5f7fa;
                border: none;
                border-bottom: 1px solid #ebeef5;
                border-right: 1px solid #ebeef5;
                padding: 6px;
                font-weight: bold;
                color: #909399;
            }
            
            QTextEdit {
                background-color: white;
                border: 1px solid #dcdfe6;
                border-radius: 4px;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
            
            QLabel {
                font-weight: bold;
                color: #303133;
                margin-top: 10px;
                margin-bottom: 5px;
            }
        """)

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 상단: 불러오기 버튼 및 설정
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)
        
        # 1. 시트 불러오기
        self.btn_load = QPushButton(" 시트 불러오기 (Load)")
        self.btn_load.setObjectName("btn_load") 
        self.btn_load.setCursor(Qt.PointingHandCursor)
        self.btn_load.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # [수정] 확장 정책
        self.btn_load.setFixedHeight(40) # 높이 통일
        self.btn_load.clicked.connect(self.load_tasks)
        top_layout.addWidget(self.btn_load, 1) # [수정] Stretch 1
        
        # 2. 작업 시작
        self.btn_start = QPushButton(" 작업 시작 (Start Ready)")
        self.btn_start.setObjectName("btn_start")
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # [수정] 확장 정책
        self.btn_start.setFixedHeight(40) # 높이 통일
        self.btn_start.clicked.connect(self.start_automation)
        self.btn_start.setEnabled(False)
        top_layout.addWidget(self.btn_start, 1) # [수정] Stretch 1
        
        # 3. 예약 모드 (컨테이너)
        sched_frame = QFrame()
        sched_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # [수정] 확장 정책
        sched_frame.setFixedHeight(40) # 높이 통일
        sched_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dcdfe6;
                border-radius: 4px;
                padding: 0px; 
            }
            QLabel {
                border: none;
                margin-left: 0px; 
                font-weight: bold;
                font-size: 10pt;
                margin-top: 0px;
                margin-bottom: 0px;
                color: #333333;
            }
        """)
        sched_layout = QHBoxLayout(sched_frame)
        sched_layout.setContentsMargins(0, 0, 0, 0) # 마진 0
        sched_layout.setSpacing(15) # [수정] 간격 넓힘 (5 -> 15)
        
        # 내부 컨텐츠 가운데 정렬
        sched_layout.addStretch(1)
        
        lbl_mode = QLabel("예약 모드") 
        sched_layout.addWidget(lbl_mode)
        
        # 버튼 영역
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(5) # [수정] 버튼 사이 간격 (0 -> 5)
        
        self.btn_sched_on = QPushButton("ON")
        self.btn_sched_on.setCursor(Qt.PointingHandCursor)
        self.btn_sched_on.setCheckable(True)
        self.btn_sched_on.setFixedSize(70, 30) # [수정] 너비 확대 (50 -> 70)
        self.btn_sched_on.clicked.connect(lambda: self.set_scheduler_mode(True))
        btn_layout.addWidget(self.btn_sched_on)
        
        self.btn_sched_off = QPushButton("OFF")
        self.btn_sched_off.setCursor(Qt.PointingHandCursor)
        self.btn_sched_off.setCheckable(True)
        self.btn_sched_off.setFixedSize(70, 30) # [수정] 너비 확대 (50 -> 70)
        self.btn_sched_off.clicked.connect(lambda: self.set_scheduler_mode(False))
        btn_layout.addWidget(self.btn_sched_off)
        
        sched_layout.addLayout(btn_layout)
        sched_layout.addStretch(1) 
        
        top_layout.addWidget(sched_frame, 1) # [수정] Stretch 1
        
        # 초기 상태 (비활성화)
        self.btn_sched_on.setEnabled(False)
        self.btn_sched_off.setEnabled(False)

        # 4. 예약 취소
        self.btn_cancel_res = QPushButton(" 예약 취소 (Cancel)")
        self.btn_cancel_res.setObjectName("btn_cancel")
        self.btn_cancel_res.setCursor(Qt.PointingHandCursor)
        self.btn_cancel_res.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # [수정] 확장 정책
        self.btn_cancel_res.setFixedHeight(40) # 높이 통일
        self.btn_cancel_res.clicked.connect(self.cancel_reservation)
        self.btn_cancel_res.setEnabled(True) 
        top_layout.addWidget(self.btn_cancel_res, 1) # [수정] Stretch 1
        
        # 5. 일정 달력
        self.btn_calendar = QPushButton(" 📅 일정 달력 (Calendar)")
        self.btn_calendar.setObjectName("btn_calendar")
        self.btn_calendar.setCursor(Qt.PointingHandCursor)
        self.btn_calendar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) 
        self.btn_calendar.setFixedHeight(40) 
        self.btn_calendar.setStyleSheet("background-color: #f39c12; color: white; font-weight: bold; border-radius: 4px; border: 1px solid #e67e22;")
        self.btn_calendar.clicked.connect(self.show_calendar)
        top_layout.addWidget(self.btn_calendar, 1) 
        
        layout.addLayout(top_layout)

        # 6. 사전 로그인 (브라우저 열기) 포트 설정
        port_layout = QHBoxLayout()
        port_layout.setSpacing(10)
        
        lbl_prelogin = QLabel("🔑 사전 로그인 (수동 브라우저 열기):")
        port_layout.addWidget(lbl_prelogin)
        
        lbl_id = QLabel("아이디:")
        port_layout.addWidget(lbl_id)
        
        self.combo_prep_id = QComboBox()
        self.combo_prep_id.setMinimumWidth(250)
        self.combo_prep_id.setEditable(True)
        self.combo_prep_id.currentIndexChanged.connect(self.on_prep_id_changed)
        port_layout.addWidget(self.combo_prep_id)
        
        lbl_port = QLabel("포트:")
        port_layout.addWidget(lbl_port)
        
        self.line_prep_port = QLineEdit("9222")
        self.line_prep_port.setFixedWidth(60)
        port_layout.addWidget(self.line_prep_port)
        
        self.btn_open_browser = QPushButton(" 브라우저 열기")
        self.btn_open_browser.setCursor(Qt.PointingHandCursor)
        self.btn_open_browser.setStyleSheet("background-color: #8e44ad; color: white; font-weight: bold; border-radius: 4px; padding: 5px 15px;")
        self.btn_open_browser.clicked.connect(self.open_prep_browser)
        port_layout.addWidget(self.btn_open_browser)
        
        port_layout.addStretch()
        layout.addLayout(port_layout)
        
        # 메인: 스플리터
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(2)
        splitter.setStyleSheet("QSplitter::handle { background-color: #dcdfe6; }")
        
        # 좌측 패널: 위(대기중), 아래(예약됨)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)
        
        # 대기중인 작업 테이블
        ready_header_layout = QHBoxLayout()
        ready_header_layout.setContentsMargins(0, 0, 0, 0)
        lbl_ready = QLabel("⏳ 대기중인 작업 (Ready)")
        lbl_ready.setStyleSheet("font-size: 11pt; color: #409eff;")
        ready_header_layout.addWidget(lbl_ready)
        
        self.chk_all_ready = QCheckBox("전체선택")
        self.chk_all_ready.stateChanged.connect(lambda state: self.toggle_select_all(self.table_ready, state))
        ready_header_layout.addWidget(self.chk_all_ready)
        ready_header_layout.addStretch()
        left_layout.addLayout(ready_header_layout)
        
        self.table_ready = QTableWidget()
        self.table_ready.setAlternatingRowColors(True)
        self.setup_table(self.table_ready, table_type="ready")
        left_layout.addWidget(self.table_ready, 1) # 스트레치 추가 (비율 1)
        
        # 예약된 작업 테이블
        res_header_layout = QHBoxLayout()
        res_header_layout.setContentsMargins(0, 15, 0, 0)
        lbl_res = QLabel("📅 예약된 작업 (Reserved)")
        lbl_res.setStyleSheet("font-size: 11pt; color: #67c23a;")
        res_header_layout.addWidget(lbl_res)
        
        self.chk_all_scheduled = QCheckBox("전체선택")
        self.chk_all_scheduled.stateChanged.connect(lambda state: self.toggle_select_all(self.table_scheduled, state))
        res_header_layout.addWidget(self.chk_all_scheduled)
        res_header_layout.addStretch()
        left_layout.addLayout(res_header_layout)
        
        self.table_scheduled = QTableWidget()
        self.table_scheduled.setAlternatingRowColors(True)
        self.setup_table(self.table_scheduled, table_type="scheduled")
        left_layout.addWidget(self.table_scheduled, 1) # 스트레치 추가 (비율 1 -> 위아래 동일 높이)
        
        splitter.addWidget(left_panel)
        
        # 우측: 스플리터 (로그 + 완료목록)
        right_splitter = QSplitter(Qt.Vertical)
        
        # 1. 로그창
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)
        
        lbl_log = QLabel("📝 작업 로그 (Logs)")
        log_layout.addWidget(lbl_log)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.document().setMaximumBlockCount(5000) # [추가] 5000줄까지만 보관하여 메모리 누수 방지
        self.log_text.setStyleSheet("background-color: #2b2b2b; color: #f0f0f0; border-radius: 4px;") 
        log_layout.addWidget(self.log_text)
        
        right_splitter.addWidget(log_widget)
        
        # 2. 완료된 작업 목록 & 리셋
        comp_widget = QWidget()
        comp_layout = QVBoxLayout(comp_widget)
        comp_layout.setContentsMargins(0, 10, 0, 0)
        
        comp_header_layout = QHBoxLayout()
        comp_header_layout.setContentsMargins(0, 0, 0, 0)
        lbl_comp = QLabel("✅ 완료된 작업 (Completed - Reset to Ready)")
        lbl_comp.setStyleSheet("font-size: 11pt; color: #606266; font-weight: bold;")
        comp_header_layout.addWidget(lbl_comp)
        
        self.chk_all_comp = QCheckBox("전체선택")
        self.chk_all_comp.stateChanged.connect(lambda state: self.toggle_select_all(self.table_completed, state))
        comp_header_layout.addWidget(self.chk_all_comp)
        comp_header_layout.addStretch()
        comp_layout.addLayout(comp_header_layout)
        
        self.table_completed = QTableWidget()
        self.table_completed.setAlternatingRowColors(True)
        self.setup_table(self.table_completed, table_type="completed")
        comp_layout.addWidget(self.table_completed)
        
        self.btn_reset = QPushButton(" 선택한 작업 리셋 (Reset Task)")
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.setStyleSheet("""
            padding: 8px; font-weight: bold; color: white;
            background-color: #e6a23c; border: 1px solid #e6a23c; border-radius: 4px;
        """)
        self.btn_reset.clicked.connect(self.reset_completed_tasks)
        comp_layout.addWidget(self.btn_reset)
        
        right_splitter.addWidget(comp_widget)
        
        # 비율 조정 (로그 : 완료 = 1:1 -> 좌측 패널과 높이 비슷하게)
        right_splitter.setSizes([450, 450]) 
        
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(10, 0, 0, 0)
        right_layout.addWidget(right_splitter)
        
        splitter.addWidget(right_panel)
        splitter.setSizes([900, 500]) # 좌우 너비 조정
        
        layout.addWidget(splitter, 1) # 스트레치 1 추가 (하단 공백 제거)

        # 하단: 카페 세팅 바로가기
        link_label = QLabel('<a href="https://docs.google.com/spreadsheets/d/1LNB7mhszGpWRPrIIh7YZz0Rcdmx9EgFp7SRRFB2A87o/edit?gid=1044519412#gid=1044519412" style="color: #409eff; text-decoration: none; font-weight: bold;">카페세팅 바로가기 (Google Sheet)</a>')
        link_label.setOpenExternalLinks(True)
        link_label.setAlignment(Qt.AlignCenter)
        link_label.setStyleSheet("font-size: 11pt; margin-top: 5px; margin-bottom: 5px; padding: 5px;")
        layout.addWidget(link_label)
        
        self.scheduler_running = False

    def log(self, msg):
        # 로그 필터링: Traceback 등 불필요한 시스템 에러 메시지 숨기기
        msg_str = str(msg)
        
        # 만약 여러 줄로 분리된 에러라면 첫 줄 내지는 핵심만 남기기
        lines = msg_str.split('\n')
        filtered_lines = []
        
        # 차단할 키워드/패턴
        block_keywords = [
            "Traceback", "AttributeError", "selenium.common.exceptions",
            "urllib3.exceptions", "Stacktrace:", "GetHandleVerifier"
            # "Quota exceeded", "429", "RESOURCE_EXHAUSTED" -> 제한 원인 확인을 위해 주석 처리
        ]
        
        for raw_line in lines:
            line = raw_line.strip()
            if not line:
                continue
            
            # 원래 들여쓰기가 있는 줄 (Traceback의 코드 라인 등)은 필터링
            if raw_line.startswith("  ") or raw_line.startswith("\t"):
                continue
            
            # File "..." 형태나 ^ 가 시작되는 라인 등은 무조건 무시
            if line.startswith("File \"") or line.startswith("^") or line.startswith("line "):
                continue

            # 블록 키워드가 포함되어 있는지 확인
            is_blocked = any(bk in line for bk in block_keywords)
            if is_blocked:
                continue
                
            filtered_lines.append(line)
            
        if not filtered_lines:
            return
            
        # 여러 줄이면 합쳐서 출력 (필터링 된 결과만)
        final_msg = " ".join(filtered_lines)
        
        current_time = QDateTime.currentDateTime().toString("HH:mm:ss")
        self.log_text.append(f"[{current_time}] {final_msg}")
        
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())


    def update_log_signal(self, msg):
        # 메인 스레드로 전달해주기 위해 connect된 signal emit
        if hasattr(self, 'update_log_signal_emit'):
             self.update_log_signal_emit(msg)
        pass 
        
    def on_table_click(self, row, col):
        # 순서 변경 로직 제거됨, 단순 선택만 처리 (필요시 사용)
        pass 

    def setup_table(self, table, table_type="ready"):
        if table_type == "completed":
            table.setColumnCount(6)
            table.setHorizontalHeaderLabels(["선택", "이름", "아이디", "포트", "카페명", "게시판"])
        else:
            table.setColumnCount(8)
            table.setHorizontalHeaderLabels(["선택", "이름", "아이디", "포트", "카페명", "게시판", "업로드", "다음예약"])
            
        table.horizontalHeader().setStretchLastSection(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.cellClicked.connect(self.on_table_click) 
        table.cellDoubleClicked.connect(self.on_table_double_click)
        
        # Context Menu 설정
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos, t=table: self.on_table_context_menu(pos, t))
        
        # [추가] 정렬 기능 활성화
        table.setSortingEnabled(True)
        
        # 컬럼 너비 조정 (1500px 창에 맞춰 넉넉하게, 오른쪽 다음예약 잘 보이게)
        table.setColumnWidth(0, 40)   # 선택
        table.setColumnWidth(1, 80)   # 이름
        table.setColumnWidth(2, 100)  # 아이디
        table.setColumnWidth(3, 50)   # 포트
        table.setColumnWidth(4, 120)  # 카페명
        table.setColumnWidth(5, 140)  # 게시판
        if table_type != "completed":
            table.setColumnWidth(6, 170)  # 업로드 폭 늘림 (120 -> 170) ('(8주 파일없음!)' 표시 공간 확보)
            table.setColumnWidth(7, 200)  # 다음예약 (모두 표시)

    def on_table_context_menu(self, pos, table):
        from PySide6.QtWidgets import QMenu
        
        item = table.itemAt(pos)
        if not item: return
        
        row = item.row()
        real_idx = table.item(row, 0).data(Qt.UserRole)
        
        menu = QMenu(self)
        
        action_cancel = None
        has_next_run = self.tasks[real_idx].get('next_run')
        if table == self.table_scheduled or (table == self.table_ready and has_next_run):
            action_cancel = menu.addAction("예약 취소 (Cancel Reservation)")
            
        action_complete = menu.addAction("강제 완료 처리 (Force Complete)")
        
        action = menu.exec(table.mapToGlobal(pos))
        
        if action_cancel and action == action_cancel:
            self.cancel_single_reservation(real_idx)
        elif action == action_complete:
            self.mark_task_completed(real_idx)

    def cancel_single_reservation(self, task_idx):
        task = self.tasks[task_idx]
        confirm = QMessageBox.question(self, "확인", f"[{task['name']}] 작업을 예약 취소하시겠습니까?\n(예약 날짜가 지워지고 대기 목록으로 이동합니다)")
        if confirm == QMessageBox.Yes:
            self.log(f">>> [{task['name']}] 예약 취소 진행 중...")
            # [수정] stage_index 전달
            self.sheet_mgr.update_date_manual(task['row_index'], "", task.get('id'), stage_index=task.get('current_stage_idx'), task_data=task)
            self.log(f"   완료. 목록을 갱신합니다.")
            self.load_tasks()

    def mark_task_completed(self, task_idx):
        task = self.tasks[task_idx]
        confirm = QMessageBox.question(self, "확인", f"[{task['name']}] 작업을 강제로 완료 처리하시겠습니까?\n(목록에서 사라지고 완료 목록으로 이동합니다)")
        if confirm == QMessageBox.Yes:
            self.log(f">>> [{task['name']}] 강제 완료 처리 중...")
            if self.sheet_mgr.force_complete_task(task['row_index'], task.get('id')):
                self.log(f"   성공. 목록을 갱신합니다.")
                self.load_tasks()
            else:
                self.log(f"   실패. 시트 접근 오류 등.")

    def refresh_order_list(self, items):
        self.order_list.clear()
        for item_text in items:
            self.order_list.addItem(QListWidgetItem(item_text))
            
    # on_table_click 중복 제거 (위에서 정의함)

    def on_table_double_click(self, row, col):
        sender = self.sender()
        if not sender: return
        
        # 7번 컬럼(다음예약) 체크 (인덱스 변경됨 8 -> 7)
        if col == 7:
            # 해당 테이블의 row에 매핑된 task 찾기
            task_idx = sender.item(row, 0).data(Qt.UserRole)
            if task_idx is not None and task_idx < len(self.tasks):
                task = self.tasks[task_idx]
                current_val = task.get('next_run', '')
                
                dlg = DatePickerDialog(current_val, self.tasks, self)
                if dlg.exec():
                    new_date = dlg.get_datetime_str()
                    
                    # 1. 메모리(Task) 업데이트
                    task['next_run'] = new_date
                    
                    # 2. 구글 시트 업데이트
                    self.log(f"[{task['name']}] 예약 시간 변경 중... ({new_date})")
                    # [수정] stage_index 및 task_data 전달하여 올바른 컬럼 업데이트 (API 호출 감소)
                    self.sheet_mgr.update_date_manual(task['row_index'], new_date, task.get('id'), stage_index=task.get('current_stage_idx'), task_data=task)
                    self.log(">>> 시트 업데이트 완료. (목록이 자동 이동됩니다)")
                    
                    # 리로드 (가장 확실)
                    self.load_tasks()

    def load_tasks(self):
        """구글 시트에서 작업 목록 불러오기 (비동기)"""
        self.log("구글 시트에서 작업을 불러오는 중입니다...")
        self.btn_load.setEnabled(False)
        
        # 스레드 생성 및 실행
        self.loader_thread = TaskLoaderThread(self.sheet_mgr)
        self.loader_thread.tasksLoaded.connect(self.on_tasks_loaded)
        self.loader_thread.errorOccurred.connect(self.on_load_error)
        self.loader_thread.start()

    def on_load_error(self, err_msg):
        self.log(f"작업 불러오기 실패: {err_msg}")
        self.btn_load.setEnabled(True)

    def on_tasks_loaded(self, tasks):
        from datetime import datetime # Moved import here for clarity, assuming it's not global
        try:
            self.tasks = tasks
            self.log(f"총 {len(self.tasks)}개의 작업을 불러왔습니다.")
            
            ready_rows = []
            scheduled_rows = []
            completed_rows = []
            
            now = datetime.now()
            
            for i, task in enumerate(self.tasks):
                # 1. 완료된 작업 우선 분류
                if task.get('is_completed'):
                    completed_rows.append((i, task))
                    continue
                    
                # 2. 예약 vs 대기 분류
                is_ready = True
                nr_val = str(task.get('next_run', '')).strip()
                if nr_val:
                    try:
                        # 포맷 시도
                        try:
                            res_time = datetime.strptime(nr_val, "%Y-%m-%d %H:%M")
                        except:
                            res_time = datetime.strptime(nr_val, "%Y-%m-%d")
                        
                        if res_time > now:
                            is_ready = False
                    except:
                        pass # 포맷 에러는 Ready로 간주
                
                if is_ready:
                    ready_rows.append((i, task))
                else:
                    scheduled_rows.append((i, task))
            
            # ID 정렬 리스트 생성 (포트 매핑용 기준)
            self.all_prep_ids = sorted(list(set([t['id'] for t in self.tasks if t['id']])))
            
            # ID를 이름에 매핑 및 A열 번호(no) 매핑
            id_to_info = {}
            for t in self.tasks:
                if t['id'] and t['id'] not in id_to_info:
                    no_val = 0
                    try:
                        no_val = int(str(t.get('no', '0')).strip())
                    except:
                        no_val = 0
                    id_to_info[t['id']] = {'name': t.get('name', '이름없음'), 'no': no_val}

            # 사전 로그인 콤보박스 업데이트
            current_prep_data = self.combo_prep_id.currentData()
            self.combo_prep_id.blockSignals(True)
            self.combo_prep_id.clear()
            
            # 콤보박스 아이템용 리스트 생성 및 A열 기준 오름차순 정렬
            combo_items = []
            for uid in self.all_prep_ids:
                info = id_to_info[uid]
                display_text = f"[{info['no']}] {info['name']} ({uid})"
                combo_items.append((info['no'], display_text, uid))
                
            combo_items.sort(key=lambda x: x[0], reverse=False)
            
            for _, display_text, uid in combo_items:
                self.combo_prep_id.addItem(display_text, userData=uid)
                
            if current_prep_data in self.all_prep_ids:
                idx = self.combo_prep_id.findData(current_prep_data)
                if idx >= 0:
                     self.combo_prep_id.setCurrentIndex(idx)
            elif self.all_prep_ids:
                self.combo_prep_id.setCurrentIndex(0)
                self.on_prep_id_changed(0)
                
            self.combo_prep_id.blockSignals(False)

            # 테이블 채우기
            self.fill_table(self.table_ready, ready_rows, self.all_prep_ids)
            self.fill_table(self.table_scheduled, scheduled_rows, self.all_prep_ids)
            self.fill_table(self.table_completed, completed_rows, self.all_prep_ids)
                
            self.log(f"대기중: {len(ready_rows)}개, 예약됨: {len(scheduled_rows)}개, 완료됨: {len(completed_rows)}개 불러오기 완료.")
            self.btn_start.setEnabled(True)
            self.btn_load.setEnabled(True)
            
            # [수정] 스케줄러 버튼 활성화 및 기본값 ON (자동 시작)
            self.btn_sched_on.setEnabled(True)
            self.btn_sched_off.setEnabled(True)
            
            # 스케줄러 기본값 ON (자동 시작 - 사용자 요청 복구)
            self.set_scheduler_mode(True)

            # 포트 번호 일괄 시트 업데이트 (백그라운드)
            port_updates = {}
            for t in self.tasks:
                if t['id'] in self.all_prep_ids:
                    p_idx = self.all_prep_ids.index(t['id'])
                    port = 9222 + p_idx
                    port_updates[t['row_index']] = port
                    
            if port_updates:
                threading.Thread(target=self.sheet_mgr.update_ports_bulk, args=(port_updates,), daemon=True).start()

            
        except Exception as e:
            self.log(f"작업 목록 처리 중 오류: {e}")
            import traceback
            traceback.print_exc()
            self.btn_load.setEnabled(True)

    def reset_completed_tasks(self):
        """완료된 작업을 선택하여 리셋"""
        selected_real_indices = []
        
        for r in range(self.table_completed.rowCount()):
            if self._is_row_checked(self.table_completed, r):
                real_idx = self.table_completed.item(r, 0).data(Qt.UserRole)
                selected_real_indices.append(real_idx)
        
        if not selected_real_indices:
            QMessageBox.warning(self, "알림", "리셋할 완료된 작업을 선택해주세요.")
            return
            
        confirm = QMessageBox.question(self, "확인", f"선택한 {len(selected_real_indices)}개 작업을 리셋하시겠습니까?\n(업로드 횟수가 초기화됩니다)")
        if confirm != QMessageBox.Yes:
            return
            
        self.log(">>> 작업 리셋 중...")
        for idx in selected_real_indices:
            task = self.tasks[idx]
            # ID를 함께 전달하여 행 검증 및 재검색 유도
            if self.sheet_mgr.reset_task(task['row_index'], task.get('id')):
                self.log(f"   [{task['name']}] 리셋 완료")
            else:
                self.log(f"   [{task['name']}] 리셋 실패")
                
        # 자동 등 새로고침
        self.load_tasks()

    def toggle_select_all(self, table, state):
        for r in range(table.rowCount()):
            widget = table.cellWidget(r, 0)
            if widget:
                chk = widget.findChild(QCheckBox)
                if chk:
                    # Qt.Checked is 2, Qt.Unchecked is 0
                    chk.setChecked(state == 2)

    def _is_row_checked(self, table, row):
        """셀 위젯에서 체크박스 상태 읽기"""
        widget = table.cellWidget(row, 0)
        if widget:
            cb = widget.findChild(QCheckBox)
            if cb:
                return cb.isChecked()
        return False

    def fill_table(self, table, rows_with_index, all_ids=None):
        table.setSortingEnabled(False)
        table.clearContents()
        table.setRowCount(0)
        table.setRowCount(len(rows_with_index))
        # [추가] 수동으로 수직 헤더 라벨 리셋 (1, 2, 3...)
        table.setVerticalHeaderLabels([str(i+1) for i in range(len(rows_with_index))])
        
        for r, (real_idx, task) in enumerate(rows_with_index):
            try:
                # 중앙 정렬된 체크박스 위젯 생성
                chk_widget = QWidget()
                chk_layout = QHBoxLayout(chk_widget)
                chk_layout.setContentsMargins(0, 0, 0, 0)
                chk_layout.setAlignment(Qt.AlignCenter)
                chk_box = QCheckBox()
                chk_layout.addWidget(chk_box)
                table.setCellWidget(r, 0, chk_widget)
                
                # 데이터 저장용 hidden item (UserRole에 real_idx 저장)
                data_item = QTableWidgetItem()
                data_item.setData(Qt.UserRole, real_idx)
                table.setItem(r, 0, data_item)
                
                item_name = QTableWidgetItem(task['name'])
                item_name.setTextAlignment(Qt.AlignCenter)
                table.setItem(r, 1, item_name)
                
                # ID 및 Port 계산
                uid = task['id']
                
                item_id = QTableWidgetItem(uid)
                item_id.setTextAlignment(Qt.AlignCenter)
                table.setItem(r, 2, item_id)
                
                port_str = "-"
                if all_ids and uid in all_ids:
                    try:
                        p_idx = all_ids.index(uid)
                        port_str = str(9222 + p_idx)
                    except:
                        pass
                item_port = QTableWidgetItem(port_str)
                item_port.setTextAlignment(Qt.AlignCenter)
                table.setItem(r, 3, item_port)
                
                item_cafe = QTableWidgetItem(task['cafe_name'])
                item_cafe.setTextAlignment(Qt.AlignCenter)
                table.setItem(r, 4, item_cafe)

                item_board = QTableWidgetItem(task['board_name'])
                item_board.setTextAlignment(Qt.AlignCenter)
                table.setItem(r, 5, item_board)

                if table == self.table_completed:
                    continue

                # 6. 업로드 (남은 주기 표시)
                p_str = task['period']
                display_period = "-"
                
                if table == self.table_ready:
                    # 대기중 -> 전체 표시
                    display_period = p_str
                else:
                    # 예약됨(또는 완료) -> 현재 스테이지 표시
                    period_list = []
                    if "," in p_str:
                        period_list = [p.strip() for p in p_str.split(',')]
                    else:
                        period_list = [p_str.strip()]
                        
                    stage_idx = task.get('current_stage_idx', 0)
                    
                    if 0 <= stage_idx < len(period_list):
                         display_period = period_list[stage_idx]
                    else:
                         display_period = p_str if p_str else "완료/빈값"
                
                if not display_period: display_period = "-"
                
                # [추가] 파일 존재 여부 체크 (빨간색 표시)
                file_exists = task.get('file_exists', True)
                missing_files_str = task.get('missing_files_str', '')
                
                if not file_exists:
                    if missing_files_str:
                        display_period += f" ({missing_files_str} 없음!)"
                    else:
                        display_period += " (파일없음!)"
                
                item_period = QTableWidgetItem(display_period)
                item_period.setTextAlignment(Qt.AlignCenter)
                
                if not file_exists:
                    item_period.setForeground(QColor("#ffffff")) # 글자색 흰색
                    item_period.setBackground(QColor("#e74c3c")) # 배경색 빨간색으로 강력하게 표시
                    font = item_period.font()
                    font.setBold(True)
                    item_period.setFont(font)
                    
                table.setItem(r, 6, item_period)

                # 7. 다음 예약
                next_run_str = task['next_run']
                item_res = QTableWidgetItem(next_run_str)
                item_res.setTextAlignment(Qt.AlignCenter)
                
                # 예약 테이블인 경우 꾸미기
                if table == self.table_scheduled and next_run_str:
                    item_res.setText(f"[예약됨] {next_run_str}")
                    # 형광색 대신 편안한 초록색으로 변경 (#27ae60)
                    item_res.setForeground(QColor("#27ae60")) 
                    font = item_res.font()
                    font.setBold(True)
                    item_res.setFont(font)
                    
                table.setItem(r, 7, item_res)
            except Exception as e:
                print(f"Error filling row {r}: {e}")
                continue
            
        table.setSortingEnabled(True)

    def set_scheduler_mode(self, is_on):
        """예약 모드 ON/OFF 설정"""
        if is_on:
            if not self.scheduler_running:
                self.log(">>> [예약 모드 ON] 스케줄러를 활성화합니다. (설정된 PC 한대에서만 사용하세요)")
                self.scheduler_running = True
                
                # 기존 스레드가 살아있다면 새로 띄우지 않음 (중복 방지)
                if not hasattr(self, 'scheduler_thread') or not self.scheduler_thread.is_alive():
                    self.scheduler_thread = threading.Thread(target=self.run_scheduler)
                    self.scheduler_thread.daemon = True # 데몬 스레드로 설정하여 메인 프로세스 종료시 함께 종료
                    self.scheduler_thread.start()
            
            # 스타일: ON=Green, OFF=Default
            self.btn_sched_on.setChecked(True)
            self.btn_sched_off.setChecked(False)
            self.btn_sched_on.setStyleSheet("background-color: #2ecc71; color: white; border: 1px solid #2ecc71; font-weight: bold;")
            self.btn_sched_off.setStyleSheet("background-color: #f0f2f5; color: #606266; border: 1px solid #dcdfe6;")
            
        else:
            if self.scheduler_running:
                self.log(">>> [예약 모드 OFF] 스케줄러를 중지합니다.")
                self.scheduler_running = False
                # 스레드는 루프 플래그 확인 후 종료됨
            
            # 스타일: ON=Default, OFF=Red
            self.btn_sched_on.setChecked(False)
            self.btn_sched_off.setChecked(True)
            self.btn_sched_on.setStyleSheet("background-color: #f0f2f5; color: #606266; border: 1px solid #dcdfe6;")
            self.btn_sched_off.setStyleSheet("background-color: #ff4d4f; color: white; border: 1px solid #ff4d4f; font-weight: bold;")

    # def toggle_scheduler(self): ... (Removed)

    def cancel_reservation(self):
        """선택한 예약된 작업의 날짜를 지워서 예약을 취소함"""
        selected_real_indices = []
        
        # 예약된 작업 테이블에서 체크된 것만 대상
        for r in range(self.table_scheduled.rowCount()):
            if self._is_row_checked(self.table_scheduled, r):
                real_idx = self.table_scheduled.item(r, 0).data(Qt.UserRole)
                selected_real_indices.append(real_idx)
                
        # 대기 작업 테이블에서도 체크된 것 중 예약 정보가 있는 것 포함 (시간이 지나 대기로 넘어온 예약들)
        for r in range(self.table_ready.rowCount()):
            if self._is_row_checked(self.table_ready, r):
                real_idx = self.table_ready.item(r, 0).data(Qt.UserRole)
                if self.tasks[real_idx].get('next_run'):
                    selected_real_indices.append(real_idx)
        
        if not selected_real_indices:
            QMessageBox.warning(self, "알림", "취소할 예약 작업을 선택해주세요.")
            return

        idx_str = ", ".join([str(self.tasks[i]['row_index']) for i in selected_real_indices])
        confirm = QMessageBox.question(self, "확인", f"선택한 {len(selected_real_indices)}개 작업의 예약을 취소하시겠습니까?\n(시트의 날짜가 지워집니다)")
        
        if confirm == QMessageBox.Yes:
            self.log(f">>> 예약 취소 시작 ({len(selected_real_indices)}개)...")
            
            # 역순으로 지워야 인덱스가 꼬이지 않음 (테이블에서 제거할 때)
            # 하지만 여기서는 selected_real_indices를 task index로 가지고 있으므로,
            # 테이블의 row를 찾아서 지워야 함.
            # 간편하게: 여기서는 Task 메모리/시트 업데이트 후, '화면'만 reload 하는게 제일 안전하지만
            # "Reload without API Call"을 구현하거나, 
            # 그냥 간단히 reload_tasks() 호출 (API 비용이 들지만 안전함)
            # 사용자 요청: "예약 취소버튼도 있음 좋을듯 해" -> 기능 구현은 되었으나
            # "reload"는 깜빡임이 있음.
            # 여기서는 UI 업데이트 로직을 추가하지 않고 load_tasks()를 호출하되, 
            # load_tasks()가 너무 느리다면 최적화 고려.
             
            for i in selected_real_indices:
                task = self.tasks[i]
                task['next_run'] = "" 
                # [수정] stage_index 전달
                self.sheet_mgr.update_date_manual(task['row_index'], "", task.get('id'), stage_index=task.get('current_stage_idx'), task_data=task)
                self.log(f"   [{task['name']}] 예약 취소 완료")
                QApplication.processEvents() # Prevent GUI freeze and rendering glitches
            
            # 리스트 갱신 (전체 다시 로드)
            self.load_tasks() 
            self.log(">>> 모든 취소 완료 및 목록 갱신됨")

    def run_scheduler(self):
        import time, random
        from datetime import datetime
        
        current_bot = None
        current_user_id = None
        
        while self.scheduler_running:
            # 매 루프마다 시트에서 최신 데이터 로드 (GUI와 별개로 백그라운드 데이터)
            try:
                # self.update_log_signal("스케줄러: 백그라운드 확인 중...")
                latest_tasks = self.sheet_mgr.get_tasks()
            except Exception as e:
                self.update_log_signal(f"스케줄러 데이터 로드 실패: {e}")
                time.sleep(60)
                continue

            now = datetime.now()
            target_tasks = [] 
            
            # 예약 체크
            for task in latest_tasks:
                if not task['next_run']: continue
                
                try:
                    try:
                        res_time = datetime.strptime(task['next_run'], "%Y-%m-%d %H:%M")
                    except:
                        res_time = datetime.strptime(task['next_run'], "%Y-%m-%d")
                        
                    # time_diff = (now - res_time).total_seconds()
                    # self.update_log_signal(f"DEBUG: [{task['name']}] 예정: {res_time}, 현재: {now}, 차이: {time_diff:.1f}초")

                    if res_time <= now:
                        target_tasks.append(task)
                        self.update_log_signal(f"-> 실행 대상 발견: {task['name']} ({task['next_run']})")
                except Exception as e:
                    self.update_log_signal(f"[{task['name']}] 날짜 파싱 오류: {task['next_run']} - {e}")
                    pass

            if target_tasks:
                self.update_log_signal(f"예약된 작업 {len(target_tasks)}개를 발견했습니다. 시트 최신 데이터 기반으로 처리합니다.")
                
                # ID별 포트 할당을 위해 전체 ID 목록 생성
                all_ids = sorted(list(set([t['id'] for t in latest_tasks if t['id']])))
                
                for task in target_tasks:
                    if not self.scheduler_running: break
                    
                    uid = task['id']
                    
                    try:
                        uid = task['id']
                        
                        # 브라우저 전환 필요 여부 확인
                        if current_bot is None or current_user_id != uid:
                            if current_bot:
                                try:
                                    current_bot.close_browser()
                                except:
                                    pass
                                current_bot = None
                                current_user_id = None
                            
                            # 새 브라우저 설정
                            try:
                                port_idx = all_ids.index(uid)
                                port = 9222 + port_idx
                            except:
                                port = 9222
                                
                            # 공유 폴더 충돌 방지를 위해 로컬 사용자 폴더 내에 크롬 프로필 생성
                            profile_dir = os.path.abspath(os.path.join(os.path.expanduser("~"), "navercafe_profiles", uid))
                            
                            self.update_log_signal(f"브라우저 실행 (ID: {uid}, Port: {port})")
                            current_bot = NaverCafeBot()
                            current_bot.start_browser(port=port, profile_dir=profile_dir)
                            current_user_id = uid
                        
                        # 랜덤 지연
                        delay_sec = random.randint(10, 60) 
                        self.update_log_signal(f"[{task['name']}] 작업 대기 중... ({delay_sec}초)")
                        time.sleep(delay_sec)
                        
                        if not self.scheduler_running: break

                        # 태스크 실행 (객체 전달)
                        self.process_single_task(current_bot, task)

                    except Exception as e:
                        self.update_log_signal(f"작업 중 오류 발생: {e}")
                        # 에러 발생 시 현재 브라우저 세션 초기화 (좀비 세션 방지)
                        if current_bot:
                            try:
                                current_bot.close_browser()
                            except:
                                pass
                            current_bot = None
                            current_user_id = None

                if current_bot:
                    try:
                        current_bot.close_browser()
                    except:
                        pass
                    current_bot = None
                    current_user_id = None
                
                # 작업 완료 후 UI 갱신 요청
                from PySide6.QtCore import QMetaObject, Qt
                QMetaObject.invokeMethod(self, "load_tasks", Qt.QueuedConnection)

            # 주기적 대기 (60초 - 구글 API 할당량 초과 방지)
            time.sleep(60)
                


    def start_automation(self):
        selected_real_indices = []
        
        # ready 테이블 체크 확인
        for r in range(self.table_ready.rowCount()):
            if self._is_row_checked(self.table_ready, r):
                real_idx = self.table_ready.item(r, 0).data(Qt.UserRole)
                selected_real_indices.append(real_idx)
                
        # scheduled 테이블 체크 확인 (원하면 예약된 것도 강제 실행 가능)
        for r in range(self.table_scheduled.rowCount()):
            if self._is_row_checked(self.table_scheduled, r):
                real_idx = self.table_scheduled.item(r, 0).data(Qt.UserRole)
                selected_real_indices.append(real_idx)
                
        if not selected_real_indices:
            QMessageBox.warning(self, "경고", "선택된 작업이 없습니다.")
            return
        
        # 중복제거 (혹시 모르니)
        selected_real_indices = list(set(selected_real_indices))
        
        # 스레드 실행
        self.worker_thread = threading.Thread(target=self.run_process, args=(selected_real_indices,))
        self.worker_thread.start()

    def show_calendar(self):
        if not hasattr(self, 'tasks') or not self.tasks:
            QMessageBox.information(self, "알림", "작업 목록을 먼저 불러와주세요.")
            return
        dlg = CalendarDialog(self.tasks, self)
        dlg.exec()

    def on_prep_id_changed(self, index):
        if not hasattr(self, 'all_prep_ids') or index < 0:
            return
        uid = self.combo_prep_id.itemData(index)
        if not uid:
            return
            
        try:
            port_idx = self.all_prep_ids.index(uid)
            port = 9222 + port_idx
        except:
            port = 9222
        self.line_prep_port.setText(str(port))
        
    def open_prep_browser(self):
        uid = self.combo_prep_id.currentData()
        port_str = self.line_prep_port.text().strip()
        if not uid or not port_str.isdigit():
            QMessageBox.warning(self, "경고", "아이디를 확인하고 숫자 형태의 포트 번호를 입력하세요.")
            return
            
        port = int(port_str)
        profile_dir = os.path.abspath(os.path.join(os.path.expanduser("~"), "navercafe_profiles", uid))
        
        self.log(f">>> 사전 로그인용 브라우저를 엽니다 (ID: {uid}, Port: {port})")
        self.log(f"    브라우저가 열리면 수동으로 [로그인 상태 유지]를 꼭 체크하고 로그인해주세요.")
        self.log(f"    로그인 후 브라우저를 그대로 두셔도 자동 지원되며, 닫아도 세션이 유지됩니다.")
        
        # 백그라운드 스레드에서 브라우저 열기 (GUI 멈춤 방지)
        threading.Thread(target=self._run_prep_browser, args=(port, profile_dir)).start()
        
    def _run_prep_browser(self, port, profile_dir):
        try:
            bot = NaverCafeBot()
            bot.start_browser(port=port, profile_dir=profile_dir)
            bot.driver.get("https://nid.naver.com/nidlogin.login")
            self.update_log_signal(">>> 브라우저 실행 완료. 로그인 완료 후 창을 닫거나 그대로 두시면 됩니다.")
        except Exception as e:
            self.update_log_signal(f"수동 브라우저 실행 중 에러 (크롬 프로세스가 이미 실행중인지 확인하세요): {e}")

    def run_process(self, indices):
        """수동 실행용"""
        # ID별 포트 할당을 위해 전체 ID 목록 생성
        all_ids = sorted(list(set([t['id'] for t in self.tasks if t['id']])))
        
        current_bot = None
        current_user_id = None
        
        try:
            for idx in indices:
                task = self.tasks[idx]
                uid = task['id']
                
                # 브라우저 전환 필요 여부 확인
                if current_bot is None or current_user_id != uid:
                    if current_bot:
                        current_bot.close_browser()
                        current_bot = None
                    
                    # 새 브라우저 설정
                    try:
                        port_idx = all_ids.index(uid)
                        port = 9222 + port_idx
                    except:
                        port = 9222
                        
                    # 공유 폴더 충돌 방지를 위해 로컬 사용자 폴더 내에 크롬 프로필 생성
                    profile_dir = os.path.abspath(os.path.join(os.path.expanduser("~"), "navercafe_profiles", uid))
                    
                    self.update_log_signal(f"브라우저 실행 (ID: {uid}, Port: {port})")
                    current_bot = NaverCafeBot()
                    current_bot.start_browser(port=port, profile_dir=profile_dir)
                    current_user_id = uid
                
                
                self.process_single_task(current_bot, task)
                
        except Exception as e:
            self.update_log_signal(f"작업 중 오류 발생: {e}")
            
        finally:
            if current_bot:
                current_bot.close_browser()
                
        self.update_log_signal("모든 작업이 완료되었습니다.")
        # 작업 완료 후 자동 새로고침
        # 메인 스레드에서 실행되어야 함 -> Signal 이용하거나, load_tasks 내부에서 invokeMethod 등 처리 필요
        # 하지만 간단하게 메인 스레드가 아니면 invokeMethod 사용
        # 여기서는 self.load_tasks()가 메인 스레드 UI 접근이 포함되어 있으므로 주의.
        # QMetaObject.invokeMethod를 사용하여 메인 스레드에서 실행되도록 함.
        from PySide6.QtCore import QMetaObject, Qt, Q_ARG
        QMetaObject.invokeMethod(self, "load_tasks", Qt.QueuedConnection)
        
    def process_single_task(self, bot, task):
        """단일 작업 수행 로직"""
        # task는 딕셔너리
        self.update_log_signal(f"=== 작업 시작: {task['cafe_name']} - {task['board_name']} ===")
        
        def cancel_task_on_error(msg):
            self.update_log_signal(msg)
            
            # 실패 횟수 추적 로직 추가
            if not hasattr(self, 'task_retry_counts'):
                self.task_retry_counts = {}
                
            task_id_key = f"{task.get('id', '')}_{task.get('row_index', '')}"
            current_retries = self.task_retry_counts.get(task_id_key, 0)
            
            if current_retries >= 2:
                self.update_log_signal(f"[{task['name']}] 2회 이상 실패하여 해당 일정 칸에 에러 사유를 기록합니다.")
                
                # J:M 일정 칸에 에러 사유 쓰기
                error_label = f"에러: {msg}"
                if len(error_label) > 15:
                    error_label = error_label[:15] + ".."
                    
                task['next_run'] = error_label
                self.sheet_mgr.update_date_manual(task['row_index'], error_label, task.get('id'), stage_index=task.get('current_stage_idx'), task_data=task)
                self.task_retry_counts[task_id_key] = 0 # 리셋
            else:
                self.task_retry_counts[task_id_key] = current_retries + 1
                self.update_log_signal(f"[{task['name']}] 예약 일정은 유지됩니다. ({self.task_retry_counts[task_id_key]}/2회 재시도 실패)")
                self.update_log_signal(f"  -> 일시적 오류일 경우 다음 실행 시 재시도됩니다.")
        
        # 1. 로그인
        if not bot.login(task['id'], task['pw']):
             cancel_task_on_error("작업 실패: 로그인 불가. 계정 정보나 보안 입력을 확인하세요.")
             return

        # 2. URL 및 접속
        cafe_url = self.sheet_mgr.get_cafe_url(task['cafe_name'])
        if not cafe_url:
            cancel_task_on_error(f"작업 실패: '{task['cafe_name']}'의 URL을 찾지 못했습니다. 카페 정보를 확인하세요.")
            return
            
        if not bot.navigate_to_cafe(cafe_url):
             cancel_task_on_error(f"작업 실패: '{task['cafe_name']}' 카페 접속 실패. 주소나 상태를 확인하세요.")
             return

        # 1-1. 기존 글 삭제 로직 (2주차 이상인 경우)
        # 카페 접속 후에 실행해야 '내가 쓴 게시글' 메뉴가 보임
        # Total Count 계산
        p_str = task.get('period', '')
        if "," in p_str:
            total_cnt = len(p_str.split(','))
        elif p_str:
            total_cnt = 1
        else:
            total_cnt = 1
            
        # [수정] 스케줄러가 지정한 stage_name이 있으면 그것을 우선 사용
        # (만약 4주차 예약이면, 2주차가 안끝났어도 4주차 내용을 가져와야 함)
        if 'stage_name' in task and task['stage_name']:
            stage_name = task['stage_name']
            # stage index도 task에서 가져옴 (없으면 계산)
            current_idx = task.get('current_stage_idx', -1)
            is_first_stage = (current_idx == 0)
        else:
            # 기존 로직 (수동 실행 등)
            stage_name = self.sheet_mgr.get_current_period_name(p_str, task.get('remain_count', task.get('upload_count')), str(total_cnt))
            current_idx = -1 # 알수없음
            is_first_stage = False # 보수적 접근
            
        # [수정] 첫 번째 단계(인덱스 0)이거나, 이름에 '2주'가 포함된 경우 기존 글 삭제
        # 사용자의 요청: "꼭 2주 시작이 아니라도 맨 처음 게시글을 쓸때는 지우는 기능"
        if current_idx != -1:
             is_first_stage = (current_idx == 0)
        
        if (stage_name and "2주" in stage_name) or is_first_stage:
            self.update_log_signal(f"[{task['name']}] 첫 게시글({stage_name}) 또는 2주차 작업 감지: 기존 게시글 전체 삭제를 시도합니다.")
            if bot.delete_all_my_posts():
                self.update_log_signal(">>> 기존 게시글 삭제 완료.")
                # 삭제 후 다시 카페 메인으로 돌아와야 글쓰기 진입 가능
                bot.navigate_to_cafe(cafe_url)
            else:
                self.update_log_signal(">>> 게시글 삭제 실패 또는 글 없음.")

        if not bot.enter_board(task['board_name']):
             cancel_task_on_error(f"작업 실패: 게시판 '{task['board_name']}' 진입 실패. 카페에 해당 게시판이 있는지 확인하세요.")
             return
        
        # 3. 컨텐츠 로드
        # total_cnt already calculated above
        
        # [수정] stage 인덱스 확정
        if 'current_stage_idx' in task:
             stage = task['current_stage_idx'] + 1 # 1-based for get_body_for_stage
        else:
             stage = self.sheet_mgr.get_stage_index(p_str, task.get('remain_count', task.get('upload_count')), str(total_cnt))
        
        # stage_name은 위에서 이미 결정됨
        
        self.update_log_signal(f"[{task['name']}] 단계: {stage} ({stage_name})")
        
        # [DEBUG] 인덱스 확인
        self.update_log_signal(f"[DEBUG] Stage Index: {stage} (Logic: {task.get('current_stage_idx', -99)} + 1)")

        # 본문 텍스트 (제목/내용)
        # 1. 컬럼에서 본문 가져오기 (우선순위)
        stage_body = self.sheet_mgr.get_body_for_stage(task, stage)
        if stage_body:
            task['body'] = stage_body
            
        # 2. 폴더에서 가져오기 (보조/없을 경우, 또는 고급 포맷)
        folder_title, folder_data, is_advanced = bot.load_text_from_folder(task['file_path'], stage=stage, stage_name=stage_name)
        
        if is_advanced:
            # 고급 포맷이면 folder_data가 content_list (list of dict)
            post_content = folder_data
            
            # 여기서 제목 결정 (폴더 제목 우선, 없으면 시트 제목)
            final_title = folder_title if folder_title else task['title']
            
            # 만약 시트 본문(J열 등)이 있다면? -> 맨 뒤에 추가?
            # 사용자 의도는 텍스트 파일이 메인.
            # 하지만 J열에 뭔가를 적었다면 무시하기 아까우니 맨 뒤에 추가해줌.
            if stage_body:
                 post_content.append({'type': 'text', 'value': stage_body})
                 
        else:
            # 기존 로직 (folder_data is body string)
            folder_body = folder_data
            final_title = folder_title if folder_title else task.get('title', "(제목없음)")
            
            # 폴더 본문이 있고, 컬럼 본문이 없으면 폴더 본문 사용
            if folder_body and not stage_body:
                task['body'] = folder_body
            
            # 전문구/후문구
            before_txt = bot.load_simple_text(task['file_path'], "전문구", stage=stage, stage_name=stage_name)
            after_txt = bot.load_simple_text(task['file_path'], "후문구", stage=stage, stage_name=stage_name)
            
            # 4. 컨텐츠 구성 (Legacy)
            content_order = self.get_content_order_for_task_idx(task['row_index'])
            post_content = []
            
            # 이미지 찾기
            before_imgs = bot.find_images(task['file_path'], "전", stage=stage, stage_name=stage_name)
            after_imgs = bot.find_images(task['file_path'], "후", stage=stage, stage_name=stage_name)

            for item in content_order:
                if item == "before_img":
                    for img in before_imgs: post_content.append({'type': 'image', 'value': img})
                elif item == "after_img":
                    for img in after_imgs: post_content.append({'type': 'image', 'value': img})
                elif item == "before_txt":
                    val = before_txt if before_txt else task.get('before_txt', '')
                    if val: post_content.append({'type': 'text', 'value': val})
                elif item == "after_txt":
                    val = after_txt if after_txt else task.get('after_txt', '')
                    if val: post_content.append({'type': 'text', 'value': val})
                elif item == "body":
                    if task['body']: 
                        # [수정] 본문 시작 전에 명시적 줄바꿈(빈 줄) 추가
                        # 이전 요소(이미지나 텍스트)와 본문 사이에 공백 라인을 두기 위함
                        post_content.append({'type': 'text', 'value': '\n'})
                        post_content.append({'type': 'text', 'value': task['body']})
                
        # 5. 글쓰기
        result_url = bot.write_post(final_title, post_content)
        
        if result_url:
             self.update_log_signal(f"성공: {result_url}")
             
             # 카운트 감소 및 상태 업데이트 (먼저 수행해야 로그에 반영됨)
             # 카운트 감소 및 상태 업데이트 (먼저 수행해야 로그에 반영됨)
             # [수정] stage_index 전달
             stage_idx = task.get('current_stage_idx')
             success, new_count_str = self.sheet_mgr.decrement_upload_count(task['row_index'], task['upload_count'], task.get('id'), stage_index=stage_idx)
             
             if success and new_count_str:
                 # 메모리 상의 task 정보 업데이트 (로그 기록용)
                 task['upload_count'] = new_count_str
             
             # 로그 기록 (업데이트된 상태로 기록)
             self.sheet_mgr.log_result(task, result_url)
             
             # 다음 예약 설정 (업데이트된 카운트 기반)
             try:
                 # 만약 완료되었다면('완료'), update_next_run은 동작하지 않아야 함 (이미 빈칸처리됨 decrement에서?)
                 # decrement는 NEXT_RUN을 건드리지 않음. reset_task가 건드림.
                 # update_next_run은 "다음 예정일"을 계산함.
                 # 완료 상태가 아니라면 계산.
                 if new_count_str != "완료":
                     self.sheet_mgr.update_next_run(task['row_index'], task['period'], new_count_str, task.get('id'))
                 else:
                     # 완료되면 Next Run 지우기 (이미 지워져 있을 수 있지만 확실히)
                     self.sheet_mgr.update_date_manual(task['row_index'], "", task.get('id'))
             except:
                 pass
        else:
             cancel_task_on_error("작업 실패: 글쓰기 오류 (요소 찾기, 시간 초과 등)")

    def get_content_order_for_task_idx(self, task_idx):
        # 사용자 요청: 전, 후, 본문 순서 (이미지 -> 텍스트)
        # 전 사진 -> 후 사진 -> 본문 
        return ["before_img", "after_img", "body"]

class EmittingStream(QObject):
    textWritten = Signal(str)

    def write(self, text):
        self.textWritten.emit(str(text))
        
    def flush(self):
        pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    
    # stdout 리다이렉션 설정
    # WorkerSignals는 쓰레드용, 메인에서도 쓸 수 있게
    # 여기서는 간단히 window.log에 직접 연결하거나 시그널 사용
    
    sys.stdout = EmittingStream()
    sys.stdout.textWritten.connect(window.log)
    sys.stderr = EmittingStream()
    sys.stderr.textWritten.connect(window.log)

    signals = WorkerSignals()
    signals.log.connect(window.log)
    window.update_log_signal = signals.log.emit
    
    window.show()
    sys.exit(app.exec())