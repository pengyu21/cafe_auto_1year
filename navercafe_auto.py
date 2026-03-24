import subprocess
from pathlib import Path
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import glob
import pyperclip
import re # [추가] 정규표현식 사용
import shutil
# import pyautogui # [제거] 이미지 업로드용 (OS 파일 다이얼로그 제어) -> 클립보드 방식으로 변경
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

class NaverCafeBot:
    def __init__(self):
        self.driver = None

    def start_browser(self, port=None, profile_dir=None):
        """브라우저 실행"""
        chrome_options = Options()
        # chrome_options.add_argument("--headless") # 디버깅을 위해 헤드리스 끔
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        
        import socket
        port_in_use = False
        if port:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(('127.0.0.1', int(port)))
            if result == 0:
                port_in_use = True
            sock.close()

        if port_in_use:
            print(f"DEBUG: Port {port} is already open. Attaching to existing browser.")
            chrome_options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
        else:
            if port:
                chrome_options.add_argument(f"--remote-debugging-port={port}")
            if profile_dir:
                chrome_options.add_argument(f"--user-data-dir={profile_dir}")
        
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        self.driver.implicitly_wait(2) # 10 -> 2로 단축 (속도 개선)
        if not port_in_use:
            self.driver.maximize_window()

    def close_browser(self):
        if self.driver:
            self.driver.quit()

    def _paste_image_from_clipboard(self, img_path: str):
        """이미지를 클립보드에 복사 (PowerShell 사용) 및 붙여넣기"""
        if not os.path.exists(img_path):
            return

        abs_path = str(Path(img_path).absolute())

        # PowerShell을 사용하여 파일 목록을 클립보드에 설정
        ps_command = (
            f"Add-Type -AssemblyName System.Windows.Forms;"
            f"$files = New-Object System.Collections.Specialized.StringCollection;"
            f"$files.Add('{abs_path}');"
            f"[System.Windows.Forms.Clipboard]::SetFileDropList($files);"
        )

        subprocess.run(
            ["powershell", "-command", ps_command],
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )

        time.sleep(1.0)
        
        # 브라우저에 포커스가 있는지 확인 후 붙여넣기
        try:
            ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        except Exception as e:
            print(f"이미지 붙여넣기 실패: {e}")
            
        time.sleep(3.5) # 업로드 대기

    ## ... (skip to _parse_advanced_text)

    def _parse_advanced_text(self, lines, folder_path, stage, stage_name):
        """★ 구분자를 사용하는 고급 텍스트 포맷 파싱"""
        title = ""
        content_list = []
        
        current_section = None
        current_text = [] 
        
        def flush_section():
            nonlocal title
            text_val = "".join(current_text).strip()
            print(f"DEBUG: Flushing Section '{current_section}' (Text len: {len(text_val)})")

    def login(self, user_id, user_pw):
        """네이버 로그인 (세션 유효 시 스킵)"""
        assert self.driver is not None
        try:
            print(f"DEBUG: Login check started for {user_id}")
            
            # [최적화] 이미 네이버에 접속해있고 로그인 된 상태라면 스킵
            # URL이 빈값('data:,')이면 무조건 접속해야 함.
            curr_url = self.driver.current_url
            print(f"DEBUG: Current URL: {curr_url}")
            
            if "naver.com" in curr_url:
                try:
                    logout_btn = self.driver.find_element(By.CLASS_NAME, "btn_logout")
                    print(f"이미 로그인 되어 있습니다. (ID: {user_id}) - 페이지 이동 생략")
                    return True
                except:
                    print("DEBUG: Naver page but not logged in (no logout btn)")
                    pass
            else:
                print("DEBUG: Not on Naver page, navigating...")

            # 0. 현재 네이버 도메인이 아니라면 쿠키를 읽을 수 없으므로 가벼운 네이버 도메인으로 먼저 이동
            curr_url = self.driver.current_url
            if not curr_url or "naver.com" not in curr_url:
                self.driver.set_page_load_timeout(10)
                try:
                    self.driver.get("https://cafe.naver.com") # 로그인 창이 아닌 평범한 카페 메인으로 이동
                except:
                    pass
                time.sleep(1)
                
            # 1. 즉시 쿠키(NID_SES) 확인 (네이버 도메인에서만 읽기 가능)
            cookies = self.driver.get_cookies()
            has_nid_ses = any(cookie.get('name') == 'NID_SES' for cookie in cookies)
            
            if has_nid_ses:
                print(f"이미 로그인 되어 있습니다. (로직 스킵, ID: {user_id})")
                return True
                
            # 2. 쿠키가 없다면(로그인 안된 상태라면) 그때서야 로그인 페이지로 진입
            print("로그인이 필요하여 로그인 페이지로 이동합니다.")
            self.driver.get("https://nid.naver.com/nidlogin.login")
            time.sleep(2)

            try:
                id_input = self.driver.find_element(By.ID, "id")
            except:
                print("이미 로그인 되어 있거나 로그인 페이지가 아닙니다.")
                return True

            # 아이디 입력
            id_input.click()
            pyperclip.copy(user_id)
            webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            time.sleep(1)

            # 비밀번호 입력
            pw_input = self.driver.find_element(By.ID, "pw")
            pw_input.click()
            pyperclip.copy(user_pw)
            webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            time.sleep(1)

            # 로그인 상태 유지 체크 (JS 클릭으로 강제 적용하여 오작동 방지)
            try:
                keep_checkbox = self.driver.find_element(By.ID, "keep")
                if not keep_checkbox.is_selected():
                    keep_label = self.driver.find_element(By.XPATH, "//label[@for='keep']")
                    self.driver.execute_script("arguments[0].click();", keep_label)
                    time.sleep(0.5)
            except Exception as e:
                print(f"로그인 상태 유지 체크박스 처리 실패 (무시됨): {e}")

            # 로그인 버튼 클릭
            self.driver.find_element(By.ID, "log.login").click()
            time.sleep(3)
            
            # 에러 메시지 확인 (비밀번호 틀림 등)
            try:
                # 네이버의 대표적인 에러 메시지 클래스 및 ID 확인
                error_elements = self.driver.find_elements(By.CLASS_NAME, "error_message")
                for err in error_elements:
                    if err.is_displayed() and err.text.strip():
                        print(f"로그인 실패 (네이버 메시지): {err.text.strip()}")
                        return False
            except:
                pass
                
            # 추가적인 에러 처리 (로그인 페이지에 계속 머물러 있는 경우)
            if "nid.naver.com/nidlogin.login" in self.driver.current_url:
                try:
                    err_common = self.driver.find_element(By.ID, "err_common")
                    if err_common.is_displayed() and err_common.text.strip():
                         print(f"로그인 실패: {err_common.text.strip()}")
                         return False
                except:
                    pass
            
            # 캡차/2단계 인증 등 대기 (필요시 수동)
            return True
        except Exception as e:
            print(f"로그인 실패: {e}")
            return False

    def navigate_to_cafe(self, cafe_url):
        """카페 접속"""
        assert self.driver is not None
        try:
            self.driver.get(cafe_url)
            time.sleep(2)
            return True
        except Exception as e:
            print(f"카페 접속 실패: {e}")
            return False

    def enter_board(self, board_name):
        """게시판 진입 (좌측 메뉴 or 상단 메뉴 검색)"""
        try:
            # 프레임 전환이 필요할 수 있음. 네이버 카페 메인은 보통 'cafe_main' 프레임 사용하지만 메뉴는 바깥에 있을 수 있음.
            # 일단 기본 컨텐츠로 검색
            
            # 메뉴 링크 찾기 (Partial Link Text 사용)
            # 네이버 카페 메뉴는 보통 iframe 밖에 있거나 cafe_main 안에 있을 수 있음. 구조에 따라 다름.
            # 통상적으로 메뉴는 `a` 태그에 텍스트로 존재
            
            # 메뉴가 너무 많아서 검색이 힘들 경우:
            # 좌측 메뉴바에서 텍스트로 찾기.
            
            # 1. 'cafe_main' 프레임이 있다면 일단 나올 것 (메뉴는 보통 메인 프레임 바깥)
            try:
                self.driver.switch_to.default_content()
            except:
                pass

            # 게시판 이름으로 링크 찾기
            try:
                board_link = self.driver.find_element(By.PARTIAL_LINK_TEXT, board_name)
                board_link.click()
            except:
                print(f"게시판 '{board_name}' 링크를 찾을 수 없습니다. 정확한 이름을 확인해주세요.")
                return False
                
            time.sleep(2)
            return True

        except Exception as e:
            print(f"게시판 진입 에러: {e}")
            return False

    def delete_all_my_posts(self):
        """'내가 쓴 게시글'로 이동하여 전체 선택 후 삭제"""
        try:
            # 0. 프레임 초기화
            try:
                self.driver.switch_to.default_content()
            except:
                pass
                
            # 1. '나의활동' 클릭
            try:
                my_act_btn = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".tit-action"))
                )
                my_act_btn.click()
                time.sleep(1)
            except:
                print("'나의활동' 버튼을 찾을 수 없거나 이미 열려있습니다.")

            # 2. '내가 쓴 게시글' 클릭
            try:
                my_post_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "내가 쓴 게시글"))
                )
                my_post_btn.click()
            except:
                print("'내가 쓴 게시글' 메뉴를 찾을 수 없습니다.")
                return False
                
            time.sleep(3)
            
            # cafe_main 프레임 전환
            try:
                self.driver.switch_to.frame("cafe_main")
            except:
                pass

            # 2-1. '전체보기' 클릭
            try:
                # 텍스트 '전체보기'를 포함한 요소 찾기
                view_all_btn = WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '전체보기')]"))
                )
                view_all_btn.click()
                time.sleep(2)
            except:
                print("'전체보기' 버튼 없음 (이미 전체 표시 중이거나 메뉴 구조 다름)")
                
            # 3. 게시글 유무 확인
            try:
                try:
                    # '등록된 게시글이 없습니다' 체크
                    no_article = self.driver.find_elements(By.CLASS_NAME, "nodata")
                    if no_article and no_article[0].is_displayed():
                        print("삭제할 게시글이 없습니다.")
                        return True
                except:
                    pass

                # 4. 전체선택 체크박스 클릭
                # User Screenshot: div.check_all > input#chk_all + label[for='chk_all']
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.ID, "chk_all"))
                    )
                    
                    check_input = self.driver.find_element(By.ID, "chk_all")
                    
                    if not check_input.is_selected():
                        # label 클릭 시도 (가장 권장됨)
                        try:
                            check_label = self.driver.find_element(By.CSS_SELECTOR, "label[for='chk_all']")
                            check_label.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", check_input)
                    
                    print("전체선택 체크 완료")
                    time.sleep(1)
                except Exception as e:
                    print(f"체크박스 선택 실패: {e}")
                    # 리스트가 아예 없는 경우일 수 있으므로 계속 진행해봄
                
                # 5. 삭제 버튼 클릭
                # User: span.BaseButton__txt '삭제'
                try:
                    # 1순위: 특정 클래스와 텍스트 조합
                    del_btn = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, 
                            "//span[contains(@class, 'BaseButton__txt') and contains(text(), '삭제')] | " + 
                            "//button[contains(@class, 'BaseButton') and .//span[contains(text(), '삭제')]]"))
                    )
                    del_btn.click()
                except:
                    # 2순위: 좀 더 포괄적인 검색
                    try:
                        del_btn = self.driver.find_element(By.XPATH, "//button[contains(text(), '삭제')] | //*[text()='삭제']")
                        self.driver.execute_script("arguments[0].click();", del_btn)
                    except:
                        print("삭제 버튼을 도저히 찾을 수 없습니다.")
                        return False
                
                print("삭제 버튼 클릭 완료")
                time.sleep(2)
                
                # 6. 확인 팝업 처리
                try:
                    # 브라우저 기본 Alert
                    WebDriverWait(self.driver, 2).until(EC.alert_is_present())
                    alert = self.driver.switch_to.alert
                    alert.accept()
                    print("Alert 확인 클릭")
                except:
                    # 커스텀 모달
                    try:
                        # '확인' 또는 '삭제' 버튼 (긍정 버튼)
                        confirm_btn = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, 
                                "//button[contains(text(), '확인')] | " +
                                "//button[contains(@class, 'confirm')] | " +
                                "//a[contains(@class, 'confirm')] | " +
                                "//button[contains(text(), '삭제')]"))
                        )
                        confirm_btn.click()
                        print("모달 확인 클릭")
                    except:
                        print("확인 팝업을 찾을 수 없으나 진행합니다.")
                    
                time.sleep(3)
                print("게시글 삭제 완료")
                return True
                
            except Exception as e:
                print(f"게시글 삭제 과정 실패: {e}")
                return False 
                
        except Exception as e:
            print(f"delete_all_my_posts 에러: {e}")
            return False

    def write_post(self, title, content_list):
        """글쓰기 실행 (공개설정: 전체공개 추가)
        content_list: [{'type': 'text'|'image', 'value': '...'}, ...] 순서대로 작성
        """
        assert self.driver is not None
        try:
            # iframe 전환 (글쓰기 버튼 누른 후 에디터는 보통 iframe 안에 있음)
            self.driver.switch_to.default_content()
            
            # 'cafe_main' 프레임 전환 (게시글 리스트가 있는 곳)
            try:
                self.driver.switch_to.frame("cafe_main")
            except:
                pass

            # 글쓰기 버튼 찾기 및 클릭 (Retry Logic)
            # 사용자 제공 클래스: BaseButtonLink BaseButton--skinGreen size_default
            for attempt in range(3):
                try:
                    btn = None
                    
                    # 1. 사용자 지정 클래스 (가장 우선)
                    try:
                        # 복합 클래스이므로 CSS Selector 사용 (띄어쓰기는 점으로)
                        # .BaseButtonLink.BaseButton--skinGreen.size_default
                        # 혹은 그냥 '글쓰기' 텍스트 포함 확인
                        btn_selector = "a.BaseButtonLink.BaseButton--skinGreen"
                        btn = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, btn_selector))
                        )
                    except:
                        pass
                    
                    # 2. 기존 방법들 (백업)
                    if not btn:
                        try:
                            btn = self.driver.find_element(By.ID, "writeFormBtn")
                        except:
                            pass
                            
                    if not btn:
                         try:
                            btn = self.driver.find_element(By.XPATH, "//a[contains(text(), '글쓰기')]")
                         except:
                            pass
                            
                    if btn:
                        # [창 1개만 쓰는 로직으로 변경]
                        # btn.click() 시 새 창이 뜨지 않도록 href를 추출하여 현재 탭에서 직접 이동
                        href = btn.get_attribute("href")
                        
                        self.driver.switch_to.default_content() # 최상위로 빠져나옴
                        
                        if href and "javascript" not in href and href != "#":
                            print(f"새 창을 띄우지 않고 글쓰기 페이지로 직접 이동합니다: {href}")
                            self.driver.get(href)
                        else:
                            print("href가 없으므로 target='_blank'를 속성에서 제거하고 클릭합니다.")
                            try:
                                self.driver.execute_script("arguments[0].removeAttribute('target');", btn)
                            except: pass
                            btn.click()
                        
                        time.sleep(5) # 페이지 이동 대기
                        break # 성공 시 루프 탈출
                        
                    else:
                        print(f"글쓰기 버튼을 찾을 수 없습니다. (시도 {attempt+1}/3)")
                        if attempt == 2:
                            return False
                        time.sleep(2)
                        
                except Exception as e:
                    print(f"글쓰기 버튼 클릭 중 에러 (시도 {attempt+1}/3): {e}")
                    if attempt == 2:
                        return False
                    time.sleep(2)
                    # 프레임 재진입 시도 for Retry
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame("cafe_main")
                    except:
                        pass
            
            # 스마트에디터 로딩 대기
            time.sleep(5) 
            
            # [중요] 단일 창 모드에서는 에디터가 iframe이 아닌 최상위에 로드될 가능성이 높음
            self.driver.switch_to.default_content()
            
            title_area = None
            
            # 1. 1차 시도: 최상위 DOM에서 찾기
            try:
                print("최상위 DOM에서 제목 입력란을 탐색합니다...")
                title_area = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "textarea.textarea_input, .ArticleWritingTitle textarea"))
                )
            except:
                pass
                
            # 2. 2차 시도: cafe_main 프레임 안에서 찾기
            if not title_area:
                try:
                    print("최상위에서 못 찾음. cafe_main 프레임 안에서 다시 시도합니다...")
                    WebDriverWait(self.driver, 2).until(EC.frame_to_be_available_and_switch_to_it("cafe_main"))
                    title_area = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "textarea.textarea_input, .ArticleWritingTitle textarea"))
                    )
                except:
                    pass

            if title_area:
                try:
                    title_area.click()
                    title_area.clear() # 기존 텍스트(예: "제목을 입력해 주세요.")가 남아있을 수 있으므로 클리어
                    
                    # 이모지(BMP 외부 문자) 등 send_keys()가 미지원하는 문자열을 위해 클립보드 붙여넣기 사용
                    pyperclip.copy(title)
                    webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                    time.sleep(1)
                except Exception as e:
                    print(f"제목 입력 (클립보드 붙여넣기) 실패: {e}")
            else:
                print("제목 검색 1, 2차 모두 실패. 구버전 에디터(subject)로 시도합니다.")
                try:
                    self.driver.switch_to.default_content()
                    try:
                        self.driver.switch_to.frame("cafe_main")
                    except: pass
                    title_input = self.driver.find_element(By.ID, "subject")
                    title_input.click()
                    pyperclip.copy(title)
                    webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                except Exception as e:
                    print(f"구버전 제목란도 찾을 수 없습니다: {e}")

            time.sleep(1)

            # 본문 작성 (Content List 순서대로)
            
            # 본문 작성 (Content List 순서대로)
            
            # 본문 작성 (Content List 순서대로)
            
            # 중요: 본문 영역 포커스 (제목 -> 본문 이동)
            print("본문 포커스 시도...")
            try:
                # 1. 사용자가 제안한 특정 텍스트 섹션 클래스를 먼저 직접 클릭 시도 (광고/공지사항 건너뛰기)
                try:
                    editor_area = self.driver.find_element(By.CSS_SELECTOR, "div.se-section.se-section-text")
                    editor_area.click()
                    print("-> div.se-section.se-section-text 클릭 성공")
                except:
                    print("-> div.se-section.se-section-text 찾기 일차 실패, 넓은 범위로 재시도합니다.")
                    try:
                        # Fallback 1: 기존 .se-components-wrap 등 타겟팅
                        editor_area = self.driver.find_element(By.CSS_SELECTOR, "article.se-components-wrap, .se-main-container")
                        editor_area.click()
                        print("-> article.se-components-wrap 또는 .se-main-container 클릭 성공")
                    except:
                        print("-> 본문 요소 찾아 클릭 실패. 탭 키로 이동 백업 실행합니다.")
                        # Fallback 2: 탭 키를 이용해 이동 (1번 탭은 공지사항에 걸릴 수 있으므로 2연속 탭 시도 - 상황에 따라)
                        webdriver.ActionChains(self.driver).send_keys(Keys.TAB).send_keys(Keys.TAB).perform()
                        
                time.sleep(1)
                
                # 2. 기존 내용이 있을 수 있으므로 끝으로 이동 (Control + End)
                webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys(Keys.END).key_up(Keys.CONTROL).perform()
                print("-> 에디터(본문) 진입 및 끝으로 커서 이동 시도")
                
            except Exception as e:
                print(f"본문 포커스 설정 에러: {e}")
                
            time.sleep(1)

            # 이미지 업로드 로직 개선 (전역 재귀 검색 + JS unhide)
            def find_and_upload_image_global(driver, file_path, current_depth=0):
                if current_depth > 3: # 깊이 제한
                    return False
                
                # 1. 현재 프레임에서 input 찾기
                try:
                    inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
                    if inputs:
                        print(f"-> [Depth {current_depth}] file input {len(inputs)}개 발견. 시도 중...")
                        for idx, inp in enumerate(inputs):
                            try:
                                # JS로 강제 unhide (필수)
                                driver.execute_script("arguments[0].style.display = 'block';", inp)
                                inp.send_keys(file_path)
                                print(f"-> input #{idx}에 send_keys 성공")
                                return True
                            except Exception as e:
                                print(f"-> input #{idx} 실패: {e}")
                                continue
                except:
                    pass

                # 2. iframe 순회 (element 기반)
                try:
                    iframes = driver.find_elements(By.TAG_NAME, "iframe")
                    for i, frame in enumerate(iframes):
                        try:
                            # frame element로 스위치
                            driver.switch_to.frame(frame)
                            if find_and_upload_image_global(driver, file_path, current_depth + 1):
                                return True
                            driver.switch_to.parent_frame()
                        except:
                            driver.switch_to.parent_frame()
                except:
                    pass
                
                return False

            # 메인 윈도우 핸들 저장
            main_window_handle = self.driver.current_window_handle
            
            # SmartEditor ONE 툴바 버튼 찾기용 (이미지)
            toolbar_btn_selector = "button.se-image-toolbar-button"

            prev_type = None # [추가] 이전 컨텐츠 타입 추적

            for i, content in enumerate(content_list):
                ctype = content['type']
                cval = content['value']
                
                if ctype == 'text':
                    # [수정] 텍스트 입력 로직 개선 (한 줄씩 입력 + 엔터)
                    # cval은 이제 줄바꿈이 보존된 상태로 옴 (`_parse_advanced_text` 수정됨)
                    
                    if not cval:
                        continue

                    # [제거] 매 텍스트마다 윈도우/프레임 전환하는 오작동 로직 제거
                    # 이미지 업로드 등에서 프레임이 꼬였을 때만 복구하도록 수정
                    try:
                        # 현재 윈도우가 에디터 윈도우인지 확인 (새창 모드 대응)
                        if self.driver.current_window_handle != main_window_handle:
                             self.driver.switch_to.window(main_window_handle)
                    except:
                        pass

                    # [추가] 이전 컨텐츠가 이미지거나 텍스트였다면, 구분용 줄바꿈(엔터) 입력
                    # 사용자 요청: "수술후 문구 쓴다음에 엔터하고 본문을 쓰면 되는거잖아" -> "엔터 말고 \n으로"
                    # [수정] 이미지는 자동으로 줄바꿈이 될 수 있으므로, 텍스트->텍스트 전환 시에만 줄바꿈 추가
                    if prev_type == 'text':
                        print(f"-> 컨텐츠 구분 줄바꿈 추가 (Previous: {prev_type})")
                        webdriver.ActionChains(self.driver).send_keys('\n').perform()
                        time.sleep(0.1)

                    # 줄바꿈 기준으로 분리
                    lines = cval.split('\n')
                    
                    for line in lines:
                        # 빈 줄 처리
                        if not line.strip():
                            # 엔터 입력 -> \n 입력
                            webdriver.ActionChains(self.driver).send_keys('\n').perform()
                            time.sleep(0.05)
                            continue
                            
                        # 내용 입력
                        pyperclip.copy(line)
                        webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
                        
                        # 붙여넣기 후 엔터 (줄바꿈) -> \n
                        webdriver.ActionChains(self.driver).send_keys('\n').perform()
                        time.sleep(0.05)
                        
                    print("-> 텍스트 입력 완료 (Line-by-Line)")

                elif ctype == 'image':
                    if not os.path.exists(cval):
                        print(f"-> [오류] 이미지 파일이 존재하지 않음: {cval}")
                        continue
                        
                    print(f"-> 이미지 업로드 시도 (클립보드): {cval}")
                    
                    # [수정] 클립보드 복사 & 프로그래머틱 붙여넣기 (PyAutoGUI 제거)
                    try:
                        # 1. 포커스 확보 (혹시 본문이 아닐 경우 대비)
                        self.driver.switch_to.window(main_window_handle)
                        try:
                            self.driver.switch_to.frame("cafe_main")
                            body_container = self.driver.find_element(By.CSS_SELECTOR, "article.se-components-wrap")
                            body_container.click()
                        except:
                            pass
                            
                        # 2. 클립보드에 이미지(파일) 복사 후 붙여넣기
                        self._paste_image_from_clipboard(cval)
                        print(f"-> 이미지 붙여넣기 완료: {cval}")
                        
                        # 엔터 입력 (이미지 뒤 줄바꿈) -> 제거 (텍스트 입력 시 처리하거나, 자동 줄바꿈)
                        # webdriver.ActionChains(self.driver).send_keys('\n').perform()
                        # time.sleep(0.5)
                            
                    except Exception as e:
                        print(f"-> 이미지 업로드(붙여넣기) 실패: {e}")

                    # 대기
                    time.sleep(1)

                prev_type = ctype # [추가] 현재 타입 저장
            time.sleep(0.2)
            
            # 6. 등록 버튼 클릭 전 '전체공개' 설정
            # cafe_main 프레임 확보
            try:
                self.driver.switch_to.window(main_window_handle)
                try:
                    self.driver.switch_to.frame("cafe_main")
                except:
                    pass
            except:
                pass

            print("전체공개 설정 시도...")
            try:
                # 1. 펼치기 버튼 확인 (JS 클릭 시도)
                try:
                    btn_open = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "div.open_set .btn_open_set"))
                    )
                    self.driver.execute_script("arguments[0].click();", btn_open)
                    time.sleep(0.2)
                except:
                    # 없으면 패스 (이미 열려있거나)
                    pass 
                
                # 2. 전체공개 라벨 클릭 (JS 클릭 시도)
                all_label = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "label[for='all']"))
                )
                self.driver.execute_script("arguments[0].click();", all_label)
                print("-> 전체공개(label for='all') 클릭 성공 (JS)")
            except Exception as e:
                print(f"-> 전체공개 설정 실패: {e}") 

            # 등록 버튼 클릭
            print("등록 버튼 클릭 시도...")
            try:
                # 사용자 지정: a class='BaseButton BaseButton--skinGreen size_default'
                pub_btn = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a.BaseButton--skinGreen"))
                )
                
                # 화면에 보이도록 스크롤
                self.driver.execute_script("arguments[0].scrollIntoView(true);", pub_btn)
                time.sleep(0.5)
                
                # JS 클릭 (가장 확실)
                self.driver.execute_script("arguments[0].click();", pub_btn)
                print("-> 등록 버튼(BaseButton--skinGreen) 클릭 성공 (JS)")
                
                # 등록 후 처리 대기
                time.sleep(2)
                
            except Exception as e:
                print(f"-> 등록 버튼 처리 중 에러: {e}")
                

                
                # 백업: 기존 방식
                try:
                    print("-> 백업 등록 버튼 시도")
                    btn = self.driver.find_element(By.XPATH, "//button[contains(text(), '등록')]")
                    btn.click()
                except:
                    return None

            # 등록 완료 후 URL 가져오기


            time.sleep(3)
            
            # 등록 완료 후 URL 가져오기
            # 등록 완료 후 처리 대기
            time.sleep(2)
            
            final_url = self.driver.current_url
            
            # 7. URL 복사 버튼 클릭 시도 (Clean URL)
            print("URL 복사 버튼 찾기 시도...")
            try:
                # 등록 후 프레임이 새로고침되거나 빠져나올 수 있으므로 다시 명시적으로 진입
                self.driver.switch_to.default_content()
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.frame_to_be_available_and_switch_to_it((By.ID, "cafe_main"))
                    )
                except:
                    pass
                
                # 복사 전 클립보드 비우기
                pyperclip.copy("")
                
                # a 태그에 button_url 클래스 우선
                copy_btn = None
                try:
                    copy_btn = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.button_url"))
                    )
                except:
                    # 혹시 button 태그일 수도 있음
                    try:
                        copy_btn = WebDriverWait(self.driver, 2).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, ".button_url"))
                        )
                    except:
                        pass
                
                if copy_btn:
                    # 복사 전 클립보드 비우기 (방심 차원)
                    pyperclip.copy("") 
                    
                    try:
                        copy_btn.click()
                    except:
                        self.driver.execute_script("arguments[0].click();", copy_btn)

                    print("-> URL 복사 버튼 클릭 완료")
                    time.sleep(2) # 클립보드 대기
                    
                    post_url = pyperclip.paste().strip()
                    
                    # 재시도 로직
                    if not post_url or "naver.me" not in post_url:
                        print("-> 클립보드 단축 URL 추출 지연/실패. 재시도 중...")
                        try:
                            copy_btn.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", copy_btn)
                        time.sleep(2)
                        post_url = pyperclip.paste().strip()
                    
                    # 결과 확인
                    if not post_url or "http" not in post_url:
                        print("-> 경고: 클립보드에서 유효한 URL을 가져오지 못했습니다. 기존 주소창 URL을 사용합니다.")
                    else:
                        print(f"-> 최종 게시글 URL 확인: {post_url}")
                        final_url = post_url
                else:
                    print("-> URL 복사 버튼을 찾을 수 없습니다.")
                    
            except Exception as e:
                print(f"URL 복사 로직 에러 (기존 URL 사용): {e}")

            return final_url


        except Exception as e:
            print(f"글쓰기 중 에러 발생: {e}")
            return None

    def find_images(self, folder_path, keyword, stage=None, stage_name=None):
        """폴더 내에서 키워드(전/후)가 포함된 이미지 파일 찾기
        stage(1, 2...) 또는 stage_name(2주, 1달...)이 포함된 파일 우선 검색
        단, 다른 StageName이 포함된 파일은 제외 (Strict Mode)
        """
        if not os.path.exists(folder_path):
            print(f"폴더가 존재하지 않습니다: {folder_path}")
            return []
            
        extensions = ['*.jpg', '*.jpeg', '*.png', '*.gif']
        images = []
        
        # 알려진 스테이지 이름 목록 (제외용)
        known_stages = ["1주", "2주", "3주", "4주", "1달", "한달", "2달", "두달", "3달", "세달", "6개월", "1년"]
        
        # 검색 패턴 설정
        search_patterns = []
        
        # 1. 정석 패턴: [Stage]*키워드
        if stage:
            # literal brackets matching for glob
            safe_stage_pat = f"[[]{stage}[]]" 
            search_patterns.append(f"{safe_stage_pat}*{keyword}*")
            
        # 2. Stage Name 패턴: *StageName* (키워드가 '후'일 때)
        if keyword == "후" and stage_name:
             search_patterns.append(f"*{stage_name}*")
             
        # 3. 일반 패턴: *키워드*
        search_patterns.append(f"*{keyword}*")

        found_files = set()
        
        for pat in search_patterns:
            for ext in extensions:
                full_pattern = os.path.join(folder_path, f"{pat}{ext[1:]}")
                matched = glob.glob(full_pattern)
                for f in matched:
                    # [수정] Strict Filtering
                    fname = os.path.basename(f)
                    
                    # 현재 찾는 stage_name이 포함되어 있으면 최우선 (통과)
                    if stage_name and stage_name in fname:
                        found_files.add(f)
                        continue
                        
                    # 다른 스테이지 이름이 포함되어 있는지 확인
                    is_other_stage = False
                    for ks in known_stages:
                        if stage_name and ks == stage_name: continue # 현재 스테이지면 통과
                        if ks in fname:
                            is_other_stage = True
                            break
                    
                    if not is_other_stage:
                        found_files.add(f)
        
        # 정렬하여 리스트 반환
        return sorted(list(found_files))

    def load_text_from_folder(self, folder_path, stage=None, stage_name=None):
        """폴더에서 텍스트 파일 읽기 (첫줄=제목, 나머지=본문)
        우선순위:
        1. [Stage]*.txt
        2. StageName.txt (예: 2주.txt)
        3. *.txt (아무거나) -> 단, 다른 StageName이 포함된 파일은 제외
        """
        if not os.path.exists(folder_path):
            print(f"폴더가 존재하지 않습니다: {folder_path}")
            return None, None, False
            
        # 알려진 스테이지 이름 목록 (제외용)
        # sheet_manager에서 가져오면 좋겠지만, 여기서 하드코딩하거나 전달받아야 함.
        # 일단 일반적인 주기 이름들 정의
        known_stages = ["1주", "2주", "3주", "4주", "1달", "한달", "2달", "두달", "3달", "세달", "6개월", "1년"]
        
        target_file = None
        
        # 0. 공통 필터링 함수
        def is_valid_file(fpath, current_stage_name):
            fname = os.path.basename(fpath)
            # 현재 스테이지 이름이 포함되어 있으면 무조건 통과
            if current_stage_name and current_stage_name in fname:
                return True
                
            # 다른 스테이지 이름이 포함되어 있으면 탈락
            curr_base = current_stage_name.replace("차", "") if current_stage_name else ""
            for ks in known_stages:
                ks_base = ks.replace("차", "")
                if curr_base and ks_base == curr_base: 
                    continue
                if ks in fname:
                    return False
            return True

        # 1. Stage Name 패턴 (예: 4주.txt) - 최우선 순위
        if stage_name:
            curr_pat = f"*{stage_name}*.txt"
            files = glob.glob(os.path.join(folder_path, curr_pat))
            if files:
                # 필터링 적용 (혹시 모르니)
                for f in files:
                    if is_valid_file(f, stage_name):
                        target_file = f
                        print(f"DEBUG: Found text file by Stage Name '{stage_name}': {os.path.basename(target_file)}")
                        break

        # 2. [Stage] 패턴 (예: [2]*.txt)
        # glob에서 [ ]는 문자셋이므로, literal [ ]를 매칭하려면 escaping 필요
        # [2] -> [[]2[]]
        if not target_file and stage:
            # literal brackets matching for glob
            safe_stage_pat = f"[[]{stage}[]]" 
            curr_pat = f"{safe_stage_pat}*.txt"
            files = glob.glob(os.path.join(folder_path, curr_pat))
            
            if files:
                for f in files:
                    # 여기도 필터링 적용 (예: [1]2주.txt 같은 혼종? 혹은 [2]가 2주와 겹칠 일은 없지만 안전하게)
                    if is_valid_file(f, stage_name):
                        target_file = f
                        print(f"DEBUG: Found text file by Stage ID [{stage}]: {os.path.basename(target_file)}")
                        break
                        
        # 3. Fallback: *.txt
        if not target_file:
            files = glob.glob(os.path.join(folder_path, "*.txt"))
            if files:
                # 필터링 및 우선순위 정렬
                valid_files = []
                for f in files:
                     if is_valid_file(f, stage_name):
                         # stage_name 포함된거 우선 (이미 1번에서 찾았겠지만 fallback에서도 챙김)
                         if stage_name and stage_name in os.path.basename(f):
                             valid_files.insert(0, f)
                         else:
                             valid_files.append(f)
                        
                if valid_files:
                    target_file = valid_files[0]
                    print(f"DEBUG: Found text file by Fallback (filtered): {os.path.basename(target_file)}")
                else:
                    print(f"DEBUG: Found text files but all seemed to belong to other stages. Skipping.")

        if not target_file:
            print("DEBUG: No suitable text file found.")
            return None, None, False
            
        try:
            lines = None
            for enc in ['utf-8', 'utf-8-sig', 'cp949', 'utf-16']:
                try:
                    with open(target_file, 'r', encoding=enc) as f:
                        lines = f.readlines()
                    break
                except UnicodeDecodeError:
                    continue

            if lines is None:
                print(f"파일 인코딩 오류 (지원하지 않는 형식): {target_file}")
                return None, None, False
                
            if not lines:
                return None, None, False
                
            title = lines[0].strip()
            
            # 고급 포맷 탐지 (★)
            is_advanced = False
            for line in lines:
                if line.strip().startswith("★"):
                    is_advanced = True
                    break
            
            if is_advanced:
                return self._parse_advanced_text(lines, folder_path, stage, stage_name)
            else:
                # [수정] 본문 원본 유지 (줄바꿈 제거 안함) - write_post에서 처리
                body = "".join(lines[1:])
                return title, body, False
        except Exception as e:
            print(f"파일 읽기 실패 ({target_file}): {e}")
            return None, None, False

    def _parse_advanced_text(self, lines, folder_path, stage, stage_name):
        """★ 구분자를 사용하는 고급 텍스트 포맷 파싱"""
        title = ""
        content_list = []
        
        current_section = None
        current_text = [] 
        
        def flush_section():
            nonlocal title
            # [수정] 줄바꿈 보존: 리스트를 join만 하고, 불필요한 공백만 제거
            # 다만, 섹션의 앞뒤 공백은 제거하되, 내부 줄바꿈은 유지
            text_val = "".join(current_text).strip()
            
            if current_section == "제목":
                title = text_val
            elif current_section in ["전사진", "후사진", "사진"]:
                # 키워드 검색 (전/후)
                if current_section == "전사진": keyword = "전"
                elif current_section == "후사진": keyword = "후"
                else: keyword = "" # 그냥 사진
                
                # 이미지 찾기
                imgs = self.find_images(folder_path, keyword, stage, stage_name)
                for img in imgs:
                    content_list.append({'type': 'image', 'value': img})
                    
            elif text_val:
                # 일반 텍스트 섹션 (본문, 수술전 문구 등)
                content_list.append({'type': 'text', 'value': text_val})
                
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("★"):
                # 이전 섹션 저장
                if current_section is not None:
                    flush_section()
                
                # 새 섹션 시작
                current_section = stripped[1:].strip() 
                current_text = []
            else:
                # [중요] 줄바꿈 문자를 포함해서 저장해야 함.
                # readlines()는 \n을 포함함.
                current_text.append(line)
        
        # 마지막 섹션 저장
        if current_section is not None:
            flush_section()
        
        return title, content_list, True

    def load_simple_text(self, folder_path, keyword, stage=None, stage_name=None):
        """특정 키워드가 포함된 텍스트 파일 내용을 읽어서 반환
        예: '전문구'
        1. [Stage]*전문구*.txt
        2. *전문구*.txt
        3. *StageName*전문구*.txt? (필요시)
        """
        if not os.path.exists(folder_path):
            return None
            
        patterns = []
        if stage:
            patterns.append(f"[{stage}]*{keyword}*.txt")
        
        # 키워드가 포함된 파일
        patterns.append(f"*{keyword}*.txt")
        
        target_file = None
        for pat in patterns:
            full_pat = os.path.join(folder_path, pat)
            files = glob.glob(full_pat)
            if files:
                target_file = files[0]
                break
                
        if not target_file:
            return None
            
        try:
            for enc in ['utf-8', 'utf-8-sig', 'cp949', 'utf-16']:
                try:
                    with open(target_file, 'r', encoding=enc) as f:
                        return f.read().strip()
                except UnicodeDecodeError:
                    continue
            print(f"텍스트 파일 인코딩 오류: {target_file}")
            return None
        except Exception as e:
            print(f"텍스트 파일 로드 실패: {e}")
            return None