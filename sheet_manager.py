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
        # Constants removed, will be instance variables
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
        
        # Constants removed, will be instance variables
        key_path = get_resource_path(key_file)
        
        creds = Credentials.from_service_account_file(key_path, scopes=scopes)
        self.client = gspread.authorize(creds)
        self.doc = self.client.open_by_url(SHEET_URL)
        self.task_sheet = self.doc.worksheet("카페작업리스트")
        self.board_sheet = self.doc.worksheet("게시판")

        # Constants removed, will be instance variables
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
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        self.COL_FILE_PATH = 8 # I
        self.COL_BODY_1 = 9    # J
        self.COL_BODY_2 = 10   # K
        self.COL_BODY_3 = 11   # L
        self.COL_BODY_4 = 12   # M
        self.COL_TITLE = 13    # N (Title? or I=8?)
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        self.COL_FILE_PATH = 8 
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        self.COL_PORT = 14     # O
        self.COL_URL = 15      # P
        self.COL_NEXT_RUN = 16 # Q
        self.COL_UPLOAD_TIME = 17 # R
        
        # Constants removed, will be instance variables
        self._map_columns()

        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        if not hasattr(self, 'COL_SCHED_1'): self.COL_SCHED_1 = 9
        
        # Constants removed, will be instance variables
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
            "꽭달": 90,
            "3개월": 90,
            "4달": 120,
            "꽕달": 120,
            "4개월": 120,
            "5달": 150,
            "떎꽢달": 150,
            "5개월": 150,
            "60일": 60,
            "2달": 60,
            "몢달": 60,
            "2개월": 60,
            "8주": 56,
            "6달": 180,
            "뿬꽢달": 180,
            "6개월": 180,
            "1년": 365
        }
        if not hasattr(self, 'COL_SCHED_2'): self.COL_SCHED_2 = 10
        if not hasattr(self, 'COL_SCHED_3'): self.COL_SCHED_3 = 11
        if not hasattr(self, 'COL_SCHED_4'): self.COL_SCHED_4 = 12

        # Constants removed, will be instance variables
        try:
            self.log_sheet = self.doc.worksheet("작업로그")
        except:
            try:
                self.log_sheet = self.doc.add_worksheet(title="작업로그", rows=1000, cols=10)
                self.log_sheet.append_row(["No", "이름", "아이디", "카페명", "게시판紐", "남은 업로드 닔", "URL", "업로드 궇吏"])
            except:
                print("작업로그 시트 생성 실패")

        # Constants removed, will be instance variables
        self.load_presets()

    def _map_columns(self, all_rows=None):
        """헤더를 읽어 컬럼 인덱스를 동적으로 매핑 (상위 5행 검색)"""
        try:
            # Constants removed, will be instance variables
            if all_rows is None:
                all_rows = self.task_sheet.get_all_values()
            if not all_rows:
                print("DEBUG: Sheet is completely empty.")
                return

            headers = []
            header_row_idx = 0
            
            # Constants removed, will be instance variables
            mapping = {
                'COL_NO': ['No', '번호'],
                'COL_NAME': ['이름', 'Name'],
                'COL_ID': ['아이디', 'ID'],
                'COL_PW': ['비번', '鍮꾨번호', 'Pass'],
                'COL_CAFE_NAME': ['카페명', 'Cafe'],
                'COL_BOARD_NAME': ['게시판', 'Board'],
                'COL_PRESET': ['단계', 'Preset', 'Period', '업로드'], # '업로드' included here (Period)
                'COL_UPLOAD_CNT': ['업로드 닔', 'Count', '횟수'], # '업로드' removed (ambiguous)

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

            # Constants removed, will be instance variables
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
                
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
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

            # Constants removed, will be instance variables
            for attr_name, keywords in mapping.items():
                found = False
                for idx, h_text in enumerate(headers):
                    cell_val = str(h_text)
                    
                    # Constants removed, will be instance variables
                    if attr_name == 'COL_UPLOAD_CNT':
                        if any(x in cell_val for x in ['궇吏', 'Date', 'Time']):
                            continue

                    for kw in keywords:
                        if kw in cell_val:
                            setattr(self, attr_name, idx)
                            found = True
                            break
                    if found: break
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            if self.COL_NEXT_RUN != 9:
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                pass
        except Exception as e:
            print(f"Map Columns Error: {e}")

    def _load_log_counts(self):
        """작업로그 시트에서 각 이름별 작업 횟수 카운트"""
        log_counts = {}
        try:
            # Constants removed, will be instance variables
            names = self.log_sheet.col_values(2)
            # Constants removed, will be instance variables
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
            # Constants removed, will be instance variables
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
            # Constants removed, will be instance variables
            try:
                current_id_val = self.task_sheet.cell(row_index, self.COL_ID + 1).value
            except:
                current_id_val = None
            
            # Constants removed, will be instance variables
            if current_id_val == task_id:
                return row_index
                
            # Constants removed, will be instance variables
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

            # Constants removed, will be instance variables
            if not hasattr(self, 'HEADER_ROW_INDEX') or self.HEADER_ROW_INDEX is None:
                self._map_columns(all_rows=rows)
        except Exception as e:
            print(f"Error fetching sheet data or mapping columns: {e}")
            return []

        tasks = []
        
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        start_row_idx = getattr(self, 'HEADER_ROW_INDEX', 0)
        
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        data_start_row = start_row_idx + 1
        if data_start_row >= len(rows):
             return tasks # No data rows

        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        dir_cache = {}

        for idx, row in enumerate(rows[data_start_row:], start=data_start_row): 
            # Constants removed, will be instance variables
            if len(row) < 16: 
                row += [""] * (16 - len(row))

            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            if not row[self.COL_NO] and not row[self.COL_NAME] and not row[self.COL_ID]:
                # Constants removed, will be instance variables
                continue
            
            name = row[self.COL_NAME]
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            preset_str = row[self.COL_PRESET]
            total_stages = 1
            preset_cnt = 1
            if preset_str:
                 if "," in str(preset_str):
                     preset_cnt = len(str(preset_str).split(','))
                 elif str(preset_str).strip():
                     preset_cnt = 1
            
            # Constants removed, will be instance variables
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

            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
            
            # Constants removed, will be instance variables
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
                
                # Constants removed, will be instance variables
                if "완료" in cleaned_val or "완료" in cell_val_str:
                     completed_stages_count += 1
                     continue

                date_str = ""
                if len(row) > col_idx:
                    raw_date = row[col_idx]
                    date_str = self._parse_date_robust(raw_date)

                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                incomplete_stages.append((i, date_str))

            is_completed_total = (completed_stages_count >= total_stages)
            remain_cnt_total = total_stages - completed_stages_count
            if remain_cnt_total < 0: remain_cnt_total = 0

            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            is_completed_by_n = False
            try:
                remain_n_val = row[self.COL_REMAIN_CNT]
                if str(remain_n_val).strip() == '0':
                    is_completed_by_n = True
            except:
                pass

            if is_completed_total or is_completed_by_n:
                 # Constants removed, will be instance variables
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
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
                global_missing_files_str = ""
                global_file_exists = True
                folder_path = row[self.COL_FILE_PATH]
                
                stages_to_check = []
                # Constants removed, will be instance variables
                s_name_base = str(preset_str).strip()
                s_arr_all = [s.strip() for s in s_name_base.split(',')] if "," in s_name_base else [s_name_base]
                
                for inc_idx, _ in incomplete_stages:
                    if inc_idx < len(s_arr_all):
                        s_name = s_arr_all[inc_idx]
                        if s_name not in stages_to_check:
                            stages_to_check.append(s_name)
                            
                global_missing_stages = []
                if folder_path and os.path.exists(folder_path):
                    # Constants removed, will be instance variables
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

                # Constants removed, will be instance variables
                has_any_scheduled = any(ds for _, ds in incomplete_stages)
                
                processed_first_pending = False
                
                # Constants removed, will be instance variables
                for stage_idx, date_str in incomplete_stages:
                    
                    should_create = False
                    
                    if date_str:
                         # Constants removed, will be instance variables
                         should_create = True
                    elif not processed_first_pending:
                         # Constants removed, will be instance variables
                         should_create = True
                         processed_first_pending = True
                         
                    if not should_create: continue
                         
                    current_period_name = s_arr_all[stage_idx] if stage_idx < len(s_arr_all) else s_name_base
                    port_val = ""
                    if hasattr(self, 'COL_PORT') and self.COL_PORT is not None and self.COL_PORT < len(row):
                        port_val = str(row[self.COL_PORT]).strip()

                    if date_str: # 예약된 작업 (현재 스테이지만 표시)
                        s_exists = False
                        if folder_path and os.path.exists(folder_path):
                            # Constants removed, will be instance variables
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
                    else: # 대기중씤 작업 (쟾泥 몴떆)
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
                        'stage_name': current_period_name, # [異붽] 쁽옱 단계 이름 (삁: '1달')
                        'upload_count': str(total_stages),
                        'remain_count': str(remain_cnt_total), 
                        'file_path': row[self.COL_FILE_PATH],
                        'next_run': date_str,
                        'is_completed': False,
                        'title': row[self.COL_TITLE],
                        'body': "", 
                        'current_stage_idx': stage_idx,
                        'file_exists': file_exists,
                        'missing_files_str': missing_files_str,
                        'port': port_val
                    }
                    tasks.append(task)

        return tasks

    def reset_task(self, row_index, task_id=None):
        """완료된 작업 리셋 (초기 상태로 복구)"""
        try:
            target_row = self._get_verified_row_index(row_index, task_id)
            if not target_row: return False

            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            try:
                preset_str = self.task_sheet.cell(target_row, self.COL_PRESET + 1).value
            except:
                preset_str = ""
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
            for col_idx in sched_cols:
                self.task_sheet.update_cell(target_row, col_idx + 1, "")
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            print(f"Reset task row {target_row}: Cleared schedule columns.")
            return True
            
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            self.task_sheet.update_cell(target_row, self.COL_REMAIN_CNT + 1, str(total_cnt))
            
            # Constants removed, will be instance variables
            self.task_sheet.update_cell(target_row, self.COL_UPLOAD_CNT + 1, str(total_cnt))
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
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
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
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

            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            remain_str = task.get('upload_count', '')
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            try:
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
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
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            range_str = f"B{next_row}:H{next_row}"
            print(f"DEBUG: Logging to {range_str}: {row_data}")
            self.log_sheet.update(range_str, [row_data])
            
        except Exception as e:
            print(f"Log Error: {e}")

    # Constants removed, will be instance variables

    # Constants removed, will be instance variables

    def get_stage_index(self, preset_name, current_remain_str, total_cnt_str):
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        try:
             remain = int(current_remain_str)
             total = int(total_cnt_str)
        except:
             return 1
             
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        current_stage = total - remain + 1
        return current_stage

    def get_current_period_name(self, preset_name, current_remain_str, total_cnt_str):
        if not preset_name: return None
        
        idx = self.get_stage_index(preset_name, current_remain_str, total_cnt_str) - 1
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        periods = []
        if "," in preset_name:
             periods = [p.strip() for p in preset_name.split(',')]
        else:
             periods = [preset_name.strip()]
             
        if idx < 0: idx = 0
        if idx >= len(periods): idx = len(periods) - 1
        
        return periods[idx]

    def get_remaining_periods(self, preset_name, current_remain_str, total_cnt_str):
        """남은 주쇨린 臾몄옄뿴 諛섑솚 (삁: '2주,한달' -> 1단계 완료 떆 '한달' 諛섑솚)"""
        if not preset_name: return ""
        
        idx = self.get_stage_index(preset_name, current_remain_str, total_cnt_str) - 1
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        periods = []
        if "," in preset_name:
             periods = [p.strip() for p in preset_name.split(',')]
        else:
             periods = [preset_name.strip()]
             
        if idx < 0: idx = 0
        
        # Constants removed, will be instance variables
        remaining = periods[idx:]
        return ",".join(remaining)

    def get_body_for_stage(self, task, stage):
        """단계별 본문 텍스트 반환 (1 -> body_1, 2 -> body_2 ...)"""
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
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
        
        # Constants removed, will be instance variables
        if p_str in getattr(self, 'PERIOD_MAP', {}):
            return self.PERIOD_MAP[p_str]
            
        # Constants removed, will be instance variables
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

            # Constants removed, will be instance variables
            if stage_index is not None:
                sched_cols = [self.COL_SCHED_1, self.COL_SCHED_2, self.COL_SCHED_3, self.COL_SCHED_4]
                if stage_index < len(sched_cols):
                    col_idx = sched_cols[stage_index]
                    
                    # Constants removed, will be instance variables
                    print(f"\n[AUTO-SCHEDULER] Setting Stage {stage_index+1} for Row {target_row} to {date_str} (Col {col_idx+1})")
                    try:
                        self.task_sheet.update_cell(target_row, col_idx + 1, date_str)
                    except Exception as e:
                        print(f"[AUTO-SCHEDULER] Error writing first base date! : {e}")
                    
                    if date_str:
                        # Constants removed, will be instance variables
                        if "뿉윭" in date_str or "실패" in date_str:
                            return
                            
                        # Constants removed, will be instance variables
                        row_vals = self.task_sheet.row_values(target_row)
                        preset_str = row_vals[self.COL_PRESET] if len(row_vals) > self.COL_PRESET else ""
                        
                        # Constants removed, will be instance variables
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
                        
                        # Constants removed, will be instance variables
                        # Constants removed, will be instance variables
                        update_values = [""] * 4
                        for i in range(len(sched_cols)):
                            c_idx = sched_cols[i]
                            if c_idx < len(row_vals):
                                update_values[i] = str(row_vals[c_idx])

                        # Constants removed, will be instance variables
                        # Constants removed, will be instance variables
                        for i in range(len(sched_cols)):
                            if i < stage_index:
                                continue # 씠쟾 뒪뀒씠吏뒗 嫄대뱶由ъ 븡쓬 (긽깭 蹂댁〈)
                            
                            # Constants removed, will be instance variables
                            if i >= total_stages:
                                update_values[i] = ""
                                continue

                            # Constants removed, will be instance variables
                            p_name = periods[i] if i < len(periods) else (periods[-1] if periods else "2주")
                            days_offset = self.get_days_from_period(p_name)
                            target_date = base_surgery_date + timedelta(days=days_offset)
                            
                            if i == stage_index:
                                # Constants removed, will be instance variables
                                target_date_str = date_str
                            else:
                                # Constants removed, will be instance variables
                                rand_hour = random.randint(10, 20)
                                rand_minute = random.randint(0, 59)
                                target_date = target_date.replace(hour=rand_hour, minute=rand_minute)
                                target_date_str = target_date.strftime("%Y-%m-%d %H:%M")
                                
                            update_values[i] = target_date_str

                        # Constants removed, will be instance variables
                        range_name = f"J{target_row}:M{target_row}"
                        try:
                            # Constants removed, will be instance variables
                            self.task_sheet.update(range_name, [update_values])
                            print(f"[AUTO-SCHEDULER] Batch update Success for range {range_name}")
                        except Exception as e:
                            print(f"[AUTO-SCHEDULER] Batch update FAILED: {e}")
                    
                    # Constants removed, will be instance variables
                    # Constants removed, will be instance variables
                    # Constants removed, will be instance variables
                    # Constants removed, will be instance variables
                    # Constants removed, will be instance variables
                    # Constants removed, will be instance variables
                        
                    return

            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
        except Exception as e:
            print(f"[AUTO-SCHEDULER] Critical Exception inside update_date_manual: {e}")


    def decrement_upload_count(self, row_index, current_count_str, task_id=None, stage_index=None):
        """남은 업로드 카운트 차감 (F열 업데이트, H열 유지) -> 이제 스케줄 관리"""
        target_row = self._get_verified_row_index(row_index, task_id)
        if not target_row: return False, None
        
        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
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

        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        sched_cols = [getattr(self, f'COL_SCHED_{i+1}', 9+i) for i in range(4)]
        
        current_stage_idx = -1
        last_date_str = ""
        
        # Constants removed, will be instance variables
        row_vals = self.task_sheet.row_values(target_row)
        
        if stage_index is not None:
             # Constants removed, will be instance variables
             current_stage_idx = stage_index
             target_col_idx = sched_cols[current_stage_idx]
             
             if target_col_idx < len(row_vals):
                 val = str(row_vals[target_col_idx])
                 last_date_str = val.split('\n')[0].strip()
        else:
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            for i in range(total_stages):
                 if i >= len(sched_cols): break
                 c_idx = sched_cols[i]
                 
                 if c_idx < len(row_vals):
                     val = str(row_vals[c_idx])
                 else:
                     val = ""
                     
                 if "완료" not in val:
                     # Constants removed, will be instance variables
                     current_stage_idx = i
                     last_date_str = val.split('\n')[0].strip() # 날짜
                     break
            
        if current_stage_idx == -1:
             # Constants removed, will be instance variables
             return True, "완료"

        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        periods = []
        if "," in str(preset_str):
            periods = [p.strip() for p in str(preset_str).split(',')]
        elif str(preset_str).strip():
            periods = [str(preset_str).strip()]
            
        current_period_name = ""
        if current_stage_idx < len(periods):
            current_period_name = periods[current_stage_idx]
            
        # Constants removed, will be instance variables
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        new_text = f"{now_str}\n{current_period_name} 완료"
        
        target_col_idx = sched_cols[current_stage_idx]
        self._safe_update_cell(target_row, target_col_idx + 1, new_text)
        print(f"Marked Stage {current_stage_idx+1} ({current_period_name}) as Done for row {target_row}")

        # Constants removed, will be instance variables
        # Constants removed, will be instance variables
        
        # Constants removed, will be instance variables
        current_date_obj = datetime.now()
        
        # Constants removed, will be instance variables
        current_days = self.get_days_from_period(current_period_name)
        base_surgery_date = current_date_obj - timedelta(days=current_days)
        
        # Constants removed, will be instance variables
        start_next_idx = current_stage_idx + 1
        
        updates_made = 0
        
        for next_idx in range(start_next_idx, total_stages):
             if next_idx >= len(sched_cols): break 
             
             # Constants removed, will be instance variables
             p_idx = next_idx
             
             p_name = "2주" # default
             if p_idx < len(periods):
                 p_name = periods[p_idx]
             elif periods:
                 p_name = periods[-1] # 마지막 주기 반복 (예: 1달, 1달, 1달...)
                 
             days = self.get_days_from_period(p_name) 
             
             # Constants removed, will be instance variables
             next_date_obj = base_surgery_date + timedelta(days=days)
             
             # Constants removed, will be instance variables
             rand_hour = random.randint(10, 20)
             rand_minute = random.randint(0, 59)
             next_date_obj = next_date_obj.replace(hour=rand_hour, minute=rand_minute)
             
             next_date_str = next_date_obj.strftime("%Y-%m-%d %H:%M")
             
             # Constants removed, will be instance variables
             target_col_idx = sched_cols[next_idx]
             
             # Constants removed, will be instance variables
             self._safe_update_cell(target_row, target_col_idx + 1, next_date_str)
             print(f"Scheduled Stage {next_idx+1} ({p_name}) at {next_date_str} (Based on {base_surgery_date.strftime('%Y-%m-%d')})")
             
             updates_made += 1

        return True, str(total_stages - (current_stage_idx + 1))

    def _safe_update_cell(self, row, col, value):
        """안전한 셀 업데이트 (재시도 로직 포함)"""
        import time
        for attempt in range(3):
            try:
                # Constants removed, will be instance variables
                # Constants removed, will be instance variables
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
            # Constants removed, will be instance variables
            remain = int(current_count_str) 
            
            # Constants removed, will be instance variables
            periods = []
            if "," in preset_str:
                periods = [p.strip() for p in preset_str.split(',') if p.strip()]
            elif preset_str.strip():
                periods = [preset_str.strip()]
                
            total = len(periods)
            if total == 0: total = 1
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            completed_idx = total - remain - 1 # 방금 완료한 인덱스?
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            # Constants removed, will be instance variables
            
            if completed_idx < 0: completed_idx = 0
            if completed_idx >= len(periods): 
                # Constants removed, will be instance variables
                return
            
            period_name = periods[completed_idx]
            next_date_str = self._calculate_next_date(period_name)
            
            if next_date_str:
                self.task_sheet.update_cell(target_row, self.COL_NEXT_RUN + 1, next_date_str)
                print(f"Updated Next Run for Row {row_index}: {next_date_str} ({period_name} after)")
            
        except Exception as e:
            print(f"Update Next Run Error: {e}")

 