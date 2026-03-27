import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import random
import glob
import os
import sys

# Constants removed, will be instance variables

SHEET_URL = "https://docs.google.com/spreadsheets/d/1LNB7mhszGpWRPrIIh7YZz0Rcdmx9EgFp7SRRFB2A87o/edit?gid=1044519412#gid=1044519412"

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class GoogleSheetManager:
    def __init__(self, key_file="service_account.json"):
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        
        # PyInstaller로 패킹된 리소스 경로 찾기
        key_path = get_resource_path(key_file)
        
        creds = Credentials.from_service_account_file(key_path, scopes=scopes)
        self.client = gspread.authorize(creds)
        self.doc = self.client.open_by_url(SHEET_URL)
        self.task_sheet = self.doc.worksheet("카페작업리스트")
        self.board_sheet = self.doc.worksheet("게시판")

        # Column Mapping (Default)
        self.COL_NO = 0
        self.COL_NAME = 1
        self.COL_ID = 2
        self.COL_PW = 3
        self.COL_CAFE_NAME = 4
        self.COL_BOARD_NAME = 5
        self.COL_PRESET = 6
        self.COL_PRESET = 6
        self.COL_UPLOAD_CNT = 7
        self.COL_REMAIN_CNT = 5 # F (Default) -> 13 (N)
        self.COL_REMAIN_CNT = 13 # [수정] N열 (Index 13)
        self.COL_BODY_1 = 8  # 1차 (J -> index 9? No 0-based. H=7, I=8??)
        # User said J is 1st.
        # A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7 (Upload Count)
        # I=8 (File Path? Previous was 8). User didn't mention I.
        # J=9 (Body 1)
        # K=10 (Body 2)
        # L=11 (Body 3)
        # M=12 (Body 4)
        # N=13 ??
        # Let's assume defaults:
        self.COL_FILE_PATH = 8 # I
        self.COL_BODY_1 = 9    # J
        self.COL_BODY_2 = 10   # K
        self.COL_BODY_3 = 11   # L
        self.COL_BODY_4 = 12   # M
        self.COL_TITLE = 13    # N (Title? or I=8?)
        # Let's assume Title is I (8) based on user report (Next Run showing Title)
        # But Path was 8. Maybe Path is gone?
        self.COL_FILE_PATH = 8 
        # Title was 13, but now Remain is 13?
        # Let's adjust Title/Remain default indices
        # If N (13) is Remain, maybe Title is O (14)? Or previously I (8)?
        # User didn't specify Title column, but existing code used 13 for Title.
        # Assuming mapping will fix this, but setting safer defaults.
        # Let's assume Title is moved or mapping will find it.
        # self.COL_TITLE = 13 
        self.COL_PORT = 14     # O
        self.COL_URL = 15      # P
        self.COL_NEXT_RUN = 16 # Q
        self.COL_UPLOAD_TIME = 17 # R
        
        # Auto Map Columns
        self._map_columns()

        # Override COL_BODY_X with SCHED Columns
        # J, K, L, M -> 9, 10, 11, 12
        if not hasattr(self, 'COL_SCHED_1'): self.COL_SCHED_1 = 9
        
        # Period Map
        self.PERIOD_MAP = {
            "1주": 7,
            "2주": 14,
            "3주": 21,
            "4주": 28,
            "5주": 35,
            "6주": 42,
            "1달": 30,
            "한달": 30,
            "1개월": 30,
            "3달": 90,
            "세달": 90,
            "3개월": 90,
            "4달": 120,
            "네달": 120,
            "4개월": 120,
            "5달": 150,
            "다섯달": 150,
            "5개월": 150,
            "60일": 60,
            "2달": 60,
            "두달": 60,
            "2개월": 60,
            "8주": 56,
            "6달": 180,
            "여섯달": 180,
            "6개월": 180,
            "1년": 365
        }
        if not hasattr(self, 'COL_SCHED_2'): self.COL_SCHED_2 = 10
        if not hasattr(self, 'COL_SCHED_3'): self.COL_SCHED_3 = 11
        if not hasattr(self, 'COL_SCHED_4'): self.COL_SCHED_4 = 12

        # '작업로그' 시트 가져오기 (없으면 생성 시도)
        try:
            self.log_sheet = self.doc.worksheet("작업로그")
        except:
            try:
                self.log_sheet = self.doc.add_worksheet(title="작업로그", rows=1000, cols=10)
                self.log_sheet.append_row(["No", "이름", "아이디", "카페명", "게시판명", "남은 업로드 수", "URL", "업로드 날짜"])
            except:
                print("작업로그 시트 생성 실패")

        # 프리셋 로드
        self.load_presets()

    def _map_columns(self, all_rows=None):
        """헤더를 읽어 컬럼 인덱스를 동적으로 매핑 (상위 5행 검색)"""
        try:
            # get_all_values로 전체 데이터 가져오기 (이미 로드된게 있으면 사용)
            if all_rows is None:
                all_rows = self.task_sheet.get_all_values()
            if not all_rows:
                print("DEBUG: Sheet is completely empty.")
                return

            headers = []
            header_row_idx = 0
            
            # 매핑 규칙 (키워드 포함 여부)
            mapping = {
                'COL_NO': ['No', '번호'],
                'COL_NAME': ['이름', 'Name'],
                'COL_ID': ['아이디', 'ID'],
                'COL_PW': ['비번', '비밀번호', 'Pass'],
                'COL_CAFE_NAME': ['카페명', 'Cafe'],
                'COL_BOARD_NAME': ['게시판', 'Board'],
                'COL_PRESET': ['단계', 'Preset', 'Period', '업로드'], # '업로드' included here (Period)
                'COL_UPLOAD_CNT': ['업로드 수', 'Count', '횟수'], # '업로드' removed (ambiguous)

                'COL_FILE_PATH': ['파일', 'Path', '위치'],
                'COL_BODY_1': ['1차', 'Body1', 'J열'],
                'COL_BODY_2': ['2차', 'Body2', 'K열'],
                'COL_BODY_3': ['3차', 'Body3', 'L열'],
                'COL_BODY_4': ['4차', 'Body4', 'M열'],
                'COL_TITLE': ['제목', 'Title'],
                'COL_TITLE': ['제목', 'Title'],
                'COL_REMAIN_CNT': ['남은', 'Remain'], # 'Count' removed to avoid conflict with Upload Count
                'COL_URL': ['URL', '링크', '주소'],
                'COL_NEXT_RUN': ['다음예약', 'Next Run', '예약일'],
                'COL_PORT': ['포트', 'Port']
            }

            # 상위 5행 중 헤더(키워드가 많이 포함된 행) 찾기
            best_score = 0
            best_row = None
            
            for r_idx in range(min(5, len(all_rows))):
                row = all_rows[r_idx]
                score = 0
                for kw_list in mapping.values():
                    for cell in row:
                        if any(k in str(cell) for k in kw_list):
                            score += 1
                            break
                
                # 키워드가 3개 이상 포함되면 헤더로 간주
                # [수정] N열(남은 업로드 수) 매핑 강화
                # 만약 헤더에 '남은'이 있으면 COL_REMAIN_CNT 업데이트
                if score >= 3 and score > best_score:
                    best_score = score
                    best_row = row
                    header_row_idx = r_idx
            
            if best_row:
                headers = best_row
                print(f"DEBUG: Found Header at Row {header_row_idx+1}: {headers}")
                self.HEADER_ROW_INDEX = header_row_idx
            else:
                print("DEBUG: Could not find a valid header row. Using default (Row 1).")
                headers = all_rows[0]
                self.HEADER_ROW_INDEX = 0

            # 매핑 수행
            for attr_name, keywords in mapping.items():
                found = False
                for idx, h_text in enumerate(headers):
                    cell_val = str(h_text)
                    
                    # 특수 예외 처리: 업로드 카운트에 '날짜'가 들어가면 안됨
                    if attr_name == 'COL_UPLOAD_CNT':
                        if any(x in cell_val for x in ['날짜', 'Date', 'Time']):
                            continue

                    for kw in keywords:
                        if kw in cell_val:
                            setattr(self, attr_name, idx)
                            found = True
                            break
                    if found: break
            
            # [수정] COL_NEXT_RUN을 J열(index 9)로 강제? 혹은 확실히 확인
            # 사용자 요청: J열에 다음 일정 기록. 만약 자동매핑이 다른곳(P열)으로 잡혔다면 J열로 수정.
            if self.COL_NEXT_RUN != 9:
                # J열 인덱스는 9 (A=0, ..., J=9)
                # 혹시 J열에 다른게 매핑되어있는지 확인? 
                # 일단 사용자가 J열을 원하므로 강제 설정 고려.
                # 하지만 기존 헤더 명칭이 '다음예약'이라면 그쪽으로 매핑됨.
                # J열 헤더가 비어있거나 다른 이름일 수 있음.
                pass
        except Exception as e:
            print(f"Map Columns Error: {e}")

    def _load_log_counts(self):
        """작업로그 시트에서 각 이름별 작업 횟수 카운트"""
        log_counts = {}
        try:
            # B열(이름) 전체 가져오기
            names = self.log_sheet.col_values(2)
            # 첫 줄(헤더) 제외
            if len(names) > 1:
                for name in names[1:]:
                    if name:
                        log_counts[name] = log_counts.get(name, 0) + 1
        except Exception as e:
            print(f"Log Count Load Error: {e}")
        return log_counts

    def load_presets(self):
        import json
        try:
            preset_path = get_resource_path('presets.json')
            with open(preset_path, 'r', encoding='utf-8') as f:
                self.presets = json.load(f)
        except:
            self.presets = {}

    def get_cafe_url(self, cafe_name):
        """Looks up the cafe URL from the '게시판' sheet."""
        try:
            # Column B is Cafe Name (Index 2 in gspread 1-based)
            cell = self.board_sheet.find(cafe_name, in_column=2)
            if cell:
                return self.board_sheet.cell(cell.row, 1).value
        except Exception as e:
            print(f"Error finding cafe URL for {cafe_name}: {e}")
        return None

    def _parse_date_robust(self, date_str):
        """다양한 포맷의 날짜 문자열을 파싱하여 정규화된 문자열(YYYY-MM-DD HH:MM 또는 YYYY-MM-DD)로 반환"""
        if not date_str: return ""
        date_str = str(date_str).strip()
        if not date_str: return ""

        formats = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
            "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y.%m.%d",
            "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"
        ]
        
        for fmt in formats:
            try:
                dt = datetime.strptime(date_str, fmt)
                if "H" in fmt: return dt.strftime("%Y-%m-%d %H:%M")
                else: return dt.strftime("%Y-%m-%d")
            except ValueError: continue
            
        try:
             norm_str = date_str.replace('.', '-').replace('/', '-')
             parts = norm_str.split('-')
             if len(parts) == 2:
                 now = datetime.now()
                 dt = datetime.strptime(f"{now.year}-{norm_str}", "%Y-%m-%d")
                 return dt.strftime("%Y-%m-%d")
        except: pass
        return ""

    def _get_verified_row_index(self, row_index, task_id):
        """주어진 row_index의 ID가 task_id와 일치하는지 확인하고, 다르면 다시 검색하여 올바른 행을 반환"""
        if not task_id:
            return row_index
            
        try:
            # 1. 현재 행의 ID 확인
            try:
                current_id_val = self.task_sheet.cell(row_index, self.COL_ID + 1).value
            except:
                current_id_val = None
            
            # 2. 일치하면 그대로 반환
            if current_id_val == task_id:
                return row_index
                
            # 3. 불일치하면 전체 검색
            print(f"DEBUG: Row mismatch! Expected {task_id}, found {current_id_val}. Searching for correct row...")
            id_list = self.task_sheet.col_values(self.COL_ID + 1)
            for i, val in enumerate(id_list):
                if val == task_id:
                    found_idx = i + 1
                    print(f"DEBUG: Found correct row at {found_idx}")
                    return found_idx
                    
            print(f"CRITICAL: Could not find task ID {task_id} in sheet.")
            return None # 찾지 못함
            
        except Exception as e:
            print(f"Error verifying row index: {e}")
            return row_index # 에러 발생 시 원래 인덱스 반환 (혹은 None?) -> 안전하게 원래값
    def get_tasks(self):
        """Reads all tasks from '카페작업리스트'."""
        try:
            rows = self.task_sheet.get_all_values()

            # 헤더가 없으면 다시 매핑 시도 (로드한 rows 재사용)
            if not hasattr(self, 'HEADER_ROW_INDEX') or self.HEADER_ROW_INDEX is None:
                self._map_columns(all_rows=rows)
        except Exception as e:
            print(f"Error fetching sheet data or mapping columns: {e}")
            return []

        tasks = []
        
        # Skip up to header row
        # HEADER_ROW_INDEX가 None이면 0으로.
        start_row_idx = getattr(self, 'HEADER_ROW_INDEX', 0)
        
        # 헤더 행이 존재하면, 해당 행의 컬럼 인덱스를 기반으로 데이터 읽기
        # 헤더 행이 없으면, 기본 컬럼 인덱스 사용 (0부터 시작)
        
        # 실제 데이터 시작 행 인덱스 (헤더 행 다음 줄)
        data_start_row = start_row_idx + 1
        if data_start_row >= len(rows):
             return tasks # No data rows

        # [최적화] 불필요한 로그 카운트 로드 제거 (사용되지 않음)
        # log_counts = self._load_log_counts()
        
        # [최적화] 폴더 리스팅 캐시
        dir_cache = {}

        for idx, row in enumerate(rows[data_start_row:], start=data_start_row): 
            # Pad row if too short
            if len(row) < 16: 
                row += [""] * (16 - len(row))

            # No와 Name, ID 체크 (유효성)
            # [수정] ID가 있으면 로드하도록 조건 완화
            if not row[self.COL_NO] and not row[self.COL_NAME] and not row[self.COL_ID]:
                # print(f"DEBUG: Skipping row {idx} (Empty No, Name, ID)")
                continue
            
            name = row[self.COL_NAME]
            
            # Preset 파싱 (총 횟수)
            # [수정] G열(Preset)의 개수와 H열(Upload Count) 중 큰 값을 총 횟수로 사용
            preset_str = row[self.COL_PRESET]
            total_stages = 1
            preset_cnt = 1
            if preset_str:
                 if "," in str(preset_str):
                     preset_cnt = len(str(preset_str).split(','))
                 elif str(preset_str).strip():
                     preset_cnt = 1
            
            # H열 확인
            try:
                h_val = str(row[self.COL_UPLOAD_CNT]).strip()
                if h_val:
                    if h_val.isdigit():
                        total_stages = max(preset_cnt, int(h_val))
                    elif '/' in h_val:
                        parts = h_val.split('/')
                        if len(parts) == 2 and parts[1].strip().isdigit():
                            total_stages = max(preset_cnt, int(parts[1].strip()))
                    else:
                        total_stages = preset_cnt
                else:
                    total_stages = preset_cnt
            except:
                total_stages = preset_cnt

            # Debug Log for Row 4 (specifically or all)
            # print(f"Processing Row {idx}: Name={name}, TotalStages={total_stages}, H_Val={h_val}")
            
            # [수정] 일정 완료 컬럼(I,J,K,L..) 확인하여 진행 상태 카운트
            sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
            
            # 1. 완료 상태 카운트 및 미완료 일정 수집
            completed_stages_count = 0
            incomplete_stages = [] # (index, date_str or empty)
            
            for i in range(total_stages):
                if i >= len(sched_cols): break
                
                col_idx = sched_cols[i]
                try:
                    val = str(row[col_idx])
                except:
                    val = ""
                    
                cell_val_str = val
                cleaned_val = cell_val_str.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")
                
                # [수정] 강력한 완료 체크
                if "완료" in cleaned_val or "완료" in cell_val_str:
                     completed_stages_count += 1
                     continue

                date_str = ""
                if len(row) > col_idx:
                    raw_date = row[col_idx]
                    date_str = self._parse_date_robust(raw_date)

                # 미완료 - 디버그 로그 (속도를 위해 주석 처리)
                # print(f"  [GET_TASKS DEBUG] Row {idx+1}, Stage {i+1}, Col {col_idx}: RawDate='{val}', ParsedDate='{date_str}'")
                incomplete_stages.append((i, date_str))

            is_completed_total = (completed_stages_count >= total_stages)
            remain_cnt_total = total_stages - completed_stages_count
            if remain_cnt_total < 0: remain_cnt_total = 0

            
            # Task 생성 로직
            # 1. 모두 완료되었으면 -> 완료된 Task 1개 생성 (대표)
            # [추가] N열(남은 업로드 수) 체크
            # N열이 0이면 (이름이 있는데) -> 완료 처리
            is_completed_by_n = False
            try:
                remain_n_val = row[self.COL_REMAIN_CNT]
                if str(remain_n_val).strip() == '0':
                    is_completed_by_n = True
            except:
                pass

            if is_completed_total or is_completed_by_n:
                 # 완료된 작업으로 처리
                 tasks.append({
                    'row_index': idx + 1,
                    'no': row[self.COL_NO],
                    'name': name,
                    'id': row[self.COL_ID],
                    'pw': row[self.COL_PW],
                    'cafe_name': row[self.COL_CAFE_NAME],
                    'board_name': row[self.COL_BOARD_NAME],
                    'period': row[self.COL_PRESET], 
                    'upload_count': str(total_stages),
                    'remain_count': str(remain_n_val) if is_completed_by_n else "0",
                    'file_path': row[self.COL_FILE_PATH],
                    'next_run': "", # 완료됨
                    'is_completed': True,
                    'title': row[self.COL_TITLE],
                    'body': "",
                    'current_stage_idx': total_stages - 1 # Last one
                 })
            else:
                # [재수정] 다시 1행 표시로 복구 (사용자 요청: "1행에 2주,1달,3달 같이 있어야지")
                # - 날짜가 있는 일정(Reserved)은 각각 생성
                # - 날짜가 없는 일정(Ready)은 "첫 번째" 것만 생성하되, 보여주는 텍스트를 "2주,1달,3달" 처럼 전체 표시하도록 함
                
                # [수정] 단일 스테이지인데 날짜가 없는 경우도 처리하기 위해 로직 보강
                # 1. Loop through incomplete stages
                # [수정] 파일 존재 여부 확인 (전체 미완료 스테이지 검사하여 세밀하게 안내)
                # 이 로직은 단일 행(row)에 대해 한 번만 전체를 검사함
                global_missing_files_str = ""
                global_file_exists = True
                folder_path = row[self.COL_FILE_PATH]
                
                stages_to_check = []
                # 파일 검사는 모든 미완료 스테이지(incomplete_stages)에 대해 수행
                s_name_base = str(preset_str).strip()
                s_arr_all = [s.strip() for s in s_name_base.split(',')] if "," in s_name_base else [s_name_base]
                
                for inc_idx, _ in incomplete_stages:
                    if inc_idx < len(s_arr_all):
                        s_name = s_arr_all[inc_idx]
                        if s_name not in stages_to_check:
                            stages_to_check.append(s_name)
                            
                global_missing_stages = []
                if folder_path and os.path.exists(folder_path):
                    # 캐시 확인
                    if folder_path in dir_cache:
                        folder_files = dir_cache[folder_path]
                    else:
                        try:
                            folder_files = [f for f in os.listdir(folder_path) if f.endswith('.txt')]
                            dir_cache[folder_path] = folder_files
                        except:
                            folder_files = []
                            dir_cache[folder_path] = []

                    for s_name in stages_to_check:
                        s_exists = False
                        for fname in folder_files:
                            if f"[{s_name}]" in fname or s_name in fname:
                                s_exists = True
                                break
                        
                        if not s_exists:
                            global_missing_stages.append(s_name)
                else:
                    global_missing_stages = stages_to_check
                    
                global_file_exists = (len(global_missing_stages) == 0)
                global_missing_files_str = ",".join(global_missing_stages)

                # [수정] 대기중 중복 제거: 해당 행에 하나라도 예약 날짜가 잡힌 스테이지가 있으면 대기목록(Ready)에 표시 안 함
                has_any_scheduled = any(ds for _, ds in incomplete_stages)
                
                processed_first_pending = False
                
                # [수정] 단일 스테이지 및 멀티 스테이지 처리 로직 보강
                for stage_idx, date_str in incomplete_stages:
                    # if '테스트' in name:
                    #    print(f"DEBUG Loop[{idx}]: Stage {stage_idx}, Date='{date_str}', Pending={processed_first_pending}")

                    should_create = False
                    
                    if date_str:
                         # 날짜가 있으면 예약된 작업이므로 고유 항목으로 무조건 생성 (대기열/예약열 상관없이)
                         should_create = True
                    elif not processed_first_pending:
                         # 날짜가 없으면 아직 설정되지 않은 '대기중' 상태.
                         # [수정] 만약 예약된 항목이 하나라도 있으면, '대기중' 항목은 굳이 표시하지 않음 (중복 방지)
                         if not has_any_scheduled:
                             should_create = True
                             processed_first_pending = True
                    
                    if not should_create:
                        continue
                        
                    current_period_name = str(preset_str).strip()
                    if "," in str(preset_str):
                        stages = [s.strip() for s in str(preset_str).split(',')]
                        if stage_idx < len(stages):
                            current_period_name = stages[stage_idx]
                    
                    # [Safety] If single stage (no comma), use the preset_str as name
                    if not "," in str(preset_str) and stage_idx == 0:
                        current_period_name = str(preset_str).strip()

                    # [수정] 예약된 작업은 개별 파일 검사 필요, 대기중은 글로벌 검사
                    
                    if date_str: # 예약된 작업 (현재 스테이지만 표시)
                        s_exists = False
                        if folder_path and os.path.exists(folder_path):
                            # 캐시 사용
                            if folder_path in dir_cache:
                                folder_files = dir_cache[folder_path]
                            else:
                                try:
                                    folder_files = [f for f in os.listdir(folder_path) if f.endswith('.txt')]
                                    dir_cache[folder_path] = folder_files
                                except:
                                    folder_files = []
                                    dir_cache[folder_path] = []
                            
                            for fname in folder_files:
                                if f"[{current_period_name}]" in fname or current_period_name in fname:
                                    s_exists = True
                                    break
                        missing_files_str = "" if s_exists else current_period_name
                        file_exists = s_exists
                    else: # 대기중인 작업 (전체 표시)
                        file_exists = global_file_exists
                        missing_files_str = global_missing_files_str
                    
                    task = {
                        'row_index': idx + 1,
                        'no': row[self.COL_NO],
                        'name': name,
                        'id': row[self.COL_ID],
                        'pw': row[self.COL_PW],
                        'cafe_name': row[self.COL_CAFE_NAME],
                        'board_name': row[self.COL_BOARD_NAME],
                        'period': row[self.COL_PRESET], 
                        'stage_name': current_period_name, # [추가] 현재 단계 이름 (예: '1달')
                        'upload_count': str(total_stages),
                        'remain_count': str(remain_cnt_total), 
                        'file_path': row[self.COL_FILE_PATH],
                        'next_run': date_str,
                        'is_completed': False,
                        'title': row[self.COL_TITLE],
                        'body': "", 
                        'current_stage_idx': stage_idx,
                        'file_exists': file_exists,
                        'missing_files_str': missing_files_str
                    }
                    tasks.append(task)

        return tasks

    def reset_task(self, row_index, task_id=None):
        """완료된 작업을 리셋 (초기 상태로 복구)"""
        try:
            target_row = self._get_verified_row_index(row_index, task_id)
            if not target_row: return False

            # 1. Period(주기) 값 읽기 (초기 총 횟수 계산용)
            # G열 (COL_PRESET) 읽기
            try:
                preset_str = self.task_sheet.cell(target_row, self.COL_PRESET + 1).value
            except:
                preset_str = ""
            
            # [수정] 리셋 요청 시 J, K, L, M 열 (예약 일정들) 모두 지우기
            # 범위: COL_SCHED_1(J) ~ COL_SCHED_4(M)
            sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
            for col_idx in sched_cols:
                self.task_sheet.update_cell(target_row, col_idx + 1, "")
            
            # 다음 예약일 지우기 중단 (사용자 보호)
            # self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, "")
            
            print(f"Reset task row {target_row}: Cleared schedule columns.")
            return True
            
            # 2. 업데이트 수행
            
            # F열(남은 횟수): 총 횟수로 초기화 (지우지 않고 명시적으로 입력)
            self.task_sheet.update_cell(target_row, self.COL_REMAIN_CNT + 1, str(total_cnt))
            
            # H열(Total): 계산된 총 횟수로 복구
            self.task_sheet.update_cell(target_row, self.COL_UPLOAD_CNT + 1, str(total_cnt))
            
            # 다음 예약 지우기 (단, PRESET 컬럼과 다를 경우에만 - G열 보호)
            # if self.COL_NEXT_RUN != self.COL_PRESET:
            #     self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, "")
            
            return True
        except Exception as e:
            print(f"Error resetting task row {row_index}: {e}")
            return False

    def update_ports_bulk(self, port_updates):
        """포트 번호를 일괄 업데이트 (API 1회 호출)
        port_updates: dict {row_index: port_number}
        """
        try:
            cells_to_update = []
            for row_idx, port_num in port_updates.items():
                a1_range = gspread.utils.rowcol_to_a1(row_idx, self.COL_PORT + 1)
                cells_to_update.append({
                    'range': a1_range,
                    'values': [[str(port_num)]]
                })
            
            if cells_to_update:
                self.task_sheet.batch_update(cells_to_update)
                print(f"Updated {len(cells_to_update)} port numbers in Google Sheet.")
        except Exception as e:
            print(f"Error updating ports: {e}")

    def force_complete_task(self, row_index, task_id=None):
        """작업을 강제로 완료 처리"""
        try:
            target_row = self._get_verified_row_index(row_index, task_id)
            if not target_row: return False
            
            # 남은 횟수(F열)를 '완료'로 변경
            # [사용자 요청] F열 건드리지 않음
            # self.task_sheet.update_cell(target_row, self.COL_REMAIN_CNT + 1, "완료")
            
            # 다음 예약일도 지움 (이건 유지?) -> 유지 안함 (사용자 보호)
            # self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, "")
            return True
        except Exception as e:
            print(f"Error forcing complete task row {row_index}: {e}")
            return False

    def log_result(self, task, url):
        """작업로그 시트에 기록 추가 (A열 공란, B열부터 입력)"""
        try:
            if not getattr(self, 'log_sheet', None):
                return
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 남은 업로드 수 계산 (시트에서 다시 읽거나, task 정보 활용)
            # decrement_upload_count가 호출된 직후라면 task['upload_count']에 최신 값이 있을 수 있음
            # 하지만 정확성을 위해 여기서 간단히 계산하거나 task 값을 신뢰
            # task['upload_count']는 decrement 함수에서 업데이트 됨 (new_count_str)
            
            remain_str = task.get('upload_count', '')
            
            # A열은 건드리지 않음 (B열부터 시작)
            # B열(이름)에 데이터가 있는 마지막 행 찾기
            try:
                # get_all_values는 느릴 수 있으니 col_values 사용?
                # col_values(2) -> B열
                b_col = self.log_sheet.col_values(2)
                next_row = len(b_col) + 1
            except:
                next_row = 1

            row_data = [
                task.get('name', ''),      # B (이름)
                task.get('id', ''),        # C (아이디)
                task.get('cafe_name', ''), # D (카페명)
                task.get('board_name', ''), # E (게시판)
                remain_str,                # F (남은 업로드 수)
                url,                       # G (URL)
                now                        # H (날짜)
            ]
            
            # B{row}:H{row} 업데이트
            # gspread cell range update
            # self.log_sheet.update(f"B{next_row}:H{next_row}", [row_data]) # gspread 6.0+ update
            # gspread 구버전일 경우 range string 지원 확인 필요. 
            # 안전하게 cell list update? 혹은 append_row는 못씀.
            
            # gspread 최신버전: worksheet.update(range_name, values)
            # update_cells는 Cell 객체 리스트 필요.
            
            # range 문자열: B5:H5
            range_str = f"B{next_row}:H{next_row}"
            print(f"DEBUG: Logging to {range_str}: {row_data}")
            self.log_sheet.update(range_str, [row_data])
            
        except Exception as e:
            print(f"Log Error: {e}")

    # ... (other methods) ...

    # ... (log_result modified previously) ...

    def get_stage_index(self, preset_name, current_remain_str, total_cnt_str):
        # current_remain_str: 남은 횟수 (ex: 2)
        # total_cnt_str: 총 횟수 (ex: 3)
        # return: 현재 진행해야 할 단계 (1, 2, 3...)
        try:
             remain = int(current_remain_str)
             total = int(total_cnt_str)
        except:
             return 1
             
        # 총 3회 중 3이 남았으면 1단계
        # 총 3회 중 2가 남았으면 2단계
        # 총 3회 중 1이 남았으면 3단계
        current_stage = total - remain + 1
        return current_stage

    def get_current_period_name(self, preset_name, current_remain_str, total_cnt_str):
        if not preset_name: return None
        
        idx = self.get_stage_index(preset_name, current_remain_str, total_cnt_str) - 1
        # idx는 0, 1, 2...
        
        # preset_name이 '2주,한달,세달' 형태라면
        periods = []
        if "," in preset_name:
             periods = [p.strip() for p in preset_name.split(',')]
        else:
             periods = [preset_name.strip()]
             
        if idx < 0: idx = 0
        if idx >= len(periods): idx = len(periods) - 1
        
        return periods[idx]

    def get_remaining_periods(self, preset_name, current_remain_str, total_cnt_str):
        """남은 주기 문자열 반환 (예: '2주,한달' -> 1단계 완료 시 '한달' 반환)"""
        if not preset_name: return ""
        
        idx = self.get_stage_index(preset_name, current_remain_str, total_cnt_str) - 1
        # idx: 현재 진행해야 할 단계의 인덱스 (0, 1, 2...)
        # 즉, 0~idx-1 까지는 완료된 것.
        # 따라서 periods[idx:] 가 남은 것들임.
        
        periods = []
        if "," in preset_name:
             periods = [p.strip() for p in preset_name.split(',')]
        else:
             periods = [preset_name.strip()]
             
        if idx < 0: idx = 0
        
        # 남은 것들만 join
        remaining = periods[idx:]
        return ",".join(remaining)

    def get_body_for_stage(self, task, stage):
        """단계별 본문 텍스트 반환 (1 -> body_1, 2 -> body_2 ...)"""
        # stage: 1-based index
        # task: 딕셔너리
        
        # stage가 1이면 body_1, 2이면 body_2...
        # 최대 4차까지 지원
        
        try:
            stage_idx = int(stage)
        except:
             stage_idx = 1
             
        key = f"body_{stage_idx}"
        return task.get(key, "")

    def get_days_from_period(self, p_name):
        """Parse period string dynamically (e.g., '7개월', '1년') to return offset days."""
        if not p_name:
            return 14
        
        p_str = str(p_name).strip()
        
        # Check standard map first
        if p_str in getattr(self, 'PERIOD_MAP', {}):
            return self.PERIOD_MAP[p_str]
            
        # Parse dynamically for N달, N개월, N년, N일, N주
        import re
        match = re.search(r'(\d+)(주|일|개월|달|년)', p_str)
        if match:
            num = int(match.group(1))
            unit = match.group(2)
            if unit == '주': return num * 7
            elif unit == '일': return num
            elif unit in ['개월', '달']: return num * 30
            elif unit == '년': return num * 365
            
        return 14

    def update_date_manual(self, row_index, date_str, task_id=None, stage_index=None, task_data=None):
        """사용자가 수동으로 예약 날짜를 지정했을 때 호출됨. 향후 단계 연쇄 자동 계산 포함"""
        try:
            target_row = self._get_verified_row_index(row_index, task_id)
            if not target_row: return

            # [수정] stage_index가 있으면 해당 일정 컬럼 업데이트 및 이후 일정 연쇄 계산
            if stage_index is not None:
                sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
                if stage_index < len(sched_cols):
                    col_idx = sched_cols[stage_index]
                    
                    # 1. 대상 스테이지 명확 업데이트
                    print(f"\n[AUTO-SCHEDULER] Setting Stage {stage_index+1} for Row {target_row} to {date_str} (Col {col_idx+1})")
                    try:
                        self.task_sheet.update_cell(target_row, col_idx + 1, date_str)
                    except Exception as e:
                        print(f"[AUTO-SCHEDULER] Error writing first base date! : {e}")
                    
                    if date_str:
                        # [복구] 필요한 정보 가져오기 (task_data가 불충분할 수 있으므로 시트 직접 참조)
                        row_vals = self.task_sheet.row_values(target_row)
                        preset_str = row_vals[self.COL_PRESET] if len(row_vals) > self.COL_PRESET else ""
                        
                        # total_stages 계산
                        preset_cnt = 1
                        if preset_str:
                            if "," in str(preset_str):
                                preset_cnt = len(str(preset_str).split(','))
                        
                        total_stages = preset_cnt
                        try:
                            h_val = str(row_vals[self.COL_UPLOAD_CNT]).strip() if len(row_vals) > self.COL_UPLOAD_CNT else ""
                            if h_val:
                                if h_val.isdigit():
                                    total_stages = max(preset_cnt, int(h_val))
                                elif '/' in h_val:
                                    parts = h_val.split('/')
                                    if len(parts) == 2 and parts[1].strip().isdigit():
                                        total_stages = max(preset_cnt, int(parts[1].strip()))
                        except:
                            pass
                        periods = []
                        if preset_str:
                            if "," in preset_str:
                                periods = [p.strip() for p in preset_str.split(',')]
                            else:
                                periods = [preset_str.strip()]
                        
                        norm_date = self._parse_date_robust(date_str)
                        current_date_obj = datetime.strptime(norm_date, "%Y-%m-%d %H:%M") if len(norm_date) > 10 else datetime.strptime(norm_date, "%Y-%m-%d")
                        
                        current_p_name = periods[stage_index] if stage_index < len(periods) else "2주"
                        current_days = self.get_days_from_period(current_p_name)
                        base_surgery_date = current_date_obj - timedelta(days=current_days)
                        
                        # 일괄 업데이트를 위한 리스트 준비 (J, K, L, M 열 값들)
                        # [수정] 기존 시트의 값들을 먼저 복사하여 '완료' 등의 상태 보존
                        update_values = [""] * 4
                        for i in range(len(sched_cols)):
                            c_idx = sched_cols[i]
                            if c_idx < len(row_vals):
                                update_values[i] = str(row_vals[c_idx])

                        # 이미 완료된 스테이지들은 유지하거나, 현재 스테이지만 업데이트
                        # 여기선 '연쇄 계산'이 목적이므로 현재 스테이지 이후를 채움
                        for i in range(len(sched_cols)):
                            if i < stage_index:
                                continue # 이전 스테이지는 건드리지 않음 (상태 보존)
                            
                            # [수정] 총 횟수를 넘어가는 컬럼은 빈칸으로 처리 (정확한 동기화)
                            if i >= total_stages:
                                update_values[i] = ""
                                continue

                            # 현재 및 향후 스테이지 계산
                            p_name = periods[i] if i < len(periods) else (periods[-1] if periods else "2주")
                            days_offset = self.get_days_from_period(p_name)
                            target_date = base_surgery_date + timedelta(days=days_offset)
                            
                            if i == stage_index:
                                # 입력받은 날짜 그대로 사용 (시간 보존)
                                target_date_str = date_str
                            else:
                                # 시간 랜덤 설정
                                rand_hour = random.randint(10, 20)
                                rand_minute = random.randint(0, 59)
                                target_date = target_date.replace(hour=rand_hour, minute=rand_minute)
                                target_date_str = target_date.strftime("%Y-%m-%d %H:%M")
                                
                            update_values[i] = target_date_str

                        # 범위 지정 (J열 ~ M열) - J는 인덱스 9이므로 컬럼명 J=10, M=13
                        range_name = f"J{target_row}:M{target_row}"
                        try:
                            # 2차원 배열 형태로 전달 [[val1, val2, val3, val4]]
                            self.task_sheet.update(range_name, [update_values])
                            print(f"[AUTO-SCHEDULER] Batch update Success for range {range_name}")
                        except Exception as e:
                            print(f"[AUTO-SCHEDULER] Batch update FAILED: {e}")
                    
                    # P열 (COL_NEXT_RUN) 쓰기 중단 (사용자 참조용 보호)
                    # try:
                    #     print(f"[AUTO-SCHEDULER]   --> Also updating global NEXT_RUN (Col {self.COL_NEXT_RUN+1})")
                    #     self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, date_str)
                    # except:
                    #     pass
                        
                    return

            # gspread is 1-based, so +1
            # self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, date_str)
        except Exception as e:
            print(f"[AUTO-SCHEDULER] Critical Exception inside update_date_manual: {e}")


    def decrement_upload_count(self, row_index, current_count_str, task_id=None, stage_index=None):
        """남은 업로드 카운트 차감 (F열 업데이트, H열은 유지) -> 이제 스케줄 관리"""
        target_row = self._get_verified_row_index(row_index, task_id)
        if not target_row: return False, None
        
        # 1. Period(주기) 값 읽기 (정확한 총 횟수 계산용)
        # [수정] G열(Preset)과 H열(Upload Count) 중 큰 값을 사용 (get_tasks와 동일)
        total_stages = 1
        preset_cnt = 1
        preset_str = ""
        try:
            preset_str = self.task_sheet.cell(target_row, self.COL_PRESET + 1).value
            if preset_str:
                if "," in str(preset_str):
                    preset_cnt = len(str(preset_str).split(','))
                elif str(preset_str).strip():
                    preset_cnt = 1
            
            h_val = self.task_sheet.cell(target_row, self.COL_UPLOAD_CNT + 1).value
            if h_val:
                if str(h_val).isdigit():
                    total_stages = max(preset_cnt, int(str(h_val)))
                elif '/' in str(h_val):
                    parts = str(h_val).split('/')
                    if len(parts) == 2 and parts[1].strip().isdigit():
                        total_stages = max(preset_cnt, int(parts[1].strip()))
                else:
                    total_stages = preset_cnt
            else:
                total_stages = preset_cnt
        except:
             total_stages = preset_cnt

        # [수정] 일정 컬럼에 '완료' 마킹 및 다음 일정 예약
        # F열은 건드리지 않음.
        
        # 1. 현재 어떤 스테이지인지 확인
        sched_cols = [getattr(self, f'COL_SCHED_{i+1}', 9+i) for i in range(4)]
        
        current_stage_idx = -1
        last_date_str = ""
        
        # 행 전체 읽기 (최신 상태)
        row_vals = self.task_sheet.row_values(target_row)
        
        if stage_index is not None:
             # 명시적으로 지정된 스테이지 사용 (GUI에서 전달됨)
             current_stage_idx = stage_index
             target_col_idx = sched_cols[current_stage_idx]
             
             if target_col_idx < len(row_vals):
                 val = str(row_vals[target_col_idx])
                 last_date_str = val.split('\n')[0].strip()
        else:
            # 자동 탐색 (Deprecated logic but kept for safety)
            # 0-based col checking
            for i in range(total_stages):
                 if i >= len(sched_cols): break
                 c_idx = sched_cols[i]
                 
                 if c_idx < len(row_vals):
                     val = str(row_vals[c_idx])
                 else:
                     val = ""
                     
                 if "완료" not in val:
                     # 미완료된 첫 번째 -> 현재 스테이지
                     current_stage_idx = i
                     last_date_str = val.split('\n')[0].strip() # 날짜
                     break
            
        if current_stage_idx == -1:
             # 모두 완료됨?
             return True, "완료"

        # 2. 현재 스테이지 컬럼에 완료 마킹
        # "2026-02-05 15:30\n2주 완료"
        
        # 주기 구하기 (2주, 1달 등)
        periods = []
        if "," in str(preset_str):
            periods = [p.strip() for p in str(preset_str).split(',')]
        elif str(preset_str).strip():
            periods = [str(preset_str).strip()]
            
        current_period_name = ""
        if current_stage_idx < len(periods):
            current_period_name = periods[current_stage_idx]
            
        # [수정] 현재 시각으로 기록 ("YYYY-MM-DD HH:MM\n[Period] 완료")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        new_text = f"{now_str}\n{current_period_name} 완료"
        
        target_col_idx = sched_cols[current_stage_idx]
        self._safe_update_cell(target_row, target_col_idx + 1, new_text)
        print(f"Marked Stage {current_stage_idx+1} ({current_period_name}) as Done for row {target_row}")

        # 3. 다음 일정 예약 (연쇄 예약)
        # [수정] 1차 완료 시 2차, 3차, 4차... 모두 계산하여 입력
        
        # 기준 날짜: 방금 완료한 시점 (now)
        current_date_obj = datetime.now()
        
        # 현재 주기의 일수를 구하여 수술일(기준일) 역산
        current_days = self.get_days_from_period(current_period_name)
        base_surgery_date = current_date_obj - timedelta(days=current_days)
        
        # 다음 단계부터 끝까지 순회
        start_next_idx = current_stage_idx + 1
        
        updates_made = 0
        
        for next_idx in range(start_next_idx, total_stages):
             if next_idx >= len(sched_cols): break 
             
             # 기간 계산을 위한 Period Index
             p_idx = next_idx
             
             p_name = "2주" # default
             if p_idx < len(periods):
                 p_name = periods[p_idx]
             elif periods:
                 p_name = periods[-1] # 마지막 주기 반복 (예: 1달, 1달, 1달...)
                 
             days = self.get_days_from_period(p_name) 
             
             # 날짜 계산 (역산된 기준일 + 해당 기간의 절대값)
             next_date_obj = base_surgery_date + timedelta(days=days)
             
             # [수정] 시간 랜덤 설정 (10:00 ~ 20:00)
             rand_hour = random.randint(10, 20)
             rand_minute = random.randint(0, 59)
             next_date_obj = next_date_obj.replace(hour=rand_hour, minute=rand_minute)
             
             next_date_str = next_date_obj.strftime("%Y-%m-%d %H:%M")
             
             # 시트 업데이트
             target_col_idx = sched_cols[next_idx]
             
             # [Safe Update]
             self._safe_update_cell(target_row, target_col_idx + 1, next_date_str)
             print(f"Scheduled Stage {next_idx+1} ({p_name}) at {next_date_str} (Based on {base_surgery_date.strftime('%Y-%m-%d')})")
             
             updates_made += 1

        return True, str(total_stages - (current_stage_idx + 1))

    def _safe_update_cell(self, row, col, value):
        """안전한 셀 업데이트 (재시도 로직 포함)"""
        import time
        for attempt in range(3):
            try:
                # gspread 버전에 따라 update_cell이 deprecated일 수 있음
                # update 사용 권장 (A1 notation or coordinates)
                self.task_sheet.update_cell(row, col, value)
                return True
            except Exception as e:
                print(f"Update failed (Attempt {attempt+1}/3): {e}")
                time.sleep(1 + attempt) # Exponential backoff
        return False
        return True, str(total_stages - (current_stage_idx + 1))

    def _calculate_next_date(self, period_name):
        """주기 이름에 따라 다음 예약 날짜를 계산하여 문자열로 반환"""
        days = self.PERIOD_MAP.get(period_name, 0)
        
        if days == 0:
            print(f"Unknown period name for days map: {period_name}")
            return None

        next_date = datetime.now() + timedelta(days=days)
        return next_date.strftime("%Y-%m-%d %H:%M")

    def update_next_run(self, row_index, preset_str, current_count_str, task_id=None):
        """다음 예약일 계산 및 업데이트 - Disbaled (Handled in decrement)"""
        return
            
        try:
            # 남은 횟수 (숫자)
            remain = int(current_count_str) 
            
            # 총 횟수 계산
            periods = []
            if "," in preset_str:
                periods = [p.strip() for p in preset_str.split(',') if p.strip()]
            elif preset_str.strip():
                periods = [preset_str.strip()]
                
            total = len(periods)
            if total == 0: total = 1
            
            # 현재 완료한 단계 인덱스 계산
            # 예: Total 3, Remain 2 -> 1회차 완료 (Index 0 완료) -> 다음은 Index 1 예약
            # 예: Total 3, Remain 1 -> 2회차 완료 (Index 1 완료) -> 다음은 Index 2 예약
            
            completed_idx = total - remain - 1 # 방금 완료한 인덱스?
            # 잠깐, remain은 decrement 후의 값.
            # Total 3.
            # 1회차 완료 -> remain 2. 
            # completed_idx = 3 - 2 - 1 = 0. (첫번째 단계 '2주' 완료)
            # 다음 예약은 periods[0] ('2주') 뒤.
            
            # 2회차 완료 -> remain 1.
            # completed_idx = 3 - 1 - 1 = 1. (두번째 단계 '한달' 완료)
            # 다음 예약은 periods[1] ('한달') 뒤.
            
            if completed_idx < 0: completed_idx = 0
            if completed_idx >= len(periods): 
                # 더 이상 예약할 단계 없음
                return
            
            period_name = periods[completed_idx]
            next_date_str = self._calculate_next_date(period_name)
            
            if next_date_str:
                self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, next_date_str)
                print(f"Updated Next Run for Row {row_index}: {next_date_str} ({period_name} after)")
            
        except Exception as e:
            print(f"Update Next Run Error: {e}")

 