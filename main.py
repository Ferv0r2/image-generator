import shutil
import time
import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import pygetwindow as gw
import pyautogui

# ✅ .env 로드
load_dotenv()

# ✅ Google Sheets API 인증
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_CONSOLE_KEY_PATH")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# ✅ Google Sheets 문서 정보
SHEET_ID = os.getenv("PROMPT_SHEET_ID")  # 시트 ID
SHEET_NAME = os.getenv("PROMPT_SHEET_NAME")  # 시트 이름

# ✅ 시트 열기 및 데이터 가져오기
sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
data = sheet.get_all_values()
df = pd.DataFrame(data[1:], columns=data[0])  # 첫 번째 행을 컬럼으로 사용

# ✅ 시작 행 번호 입력받기 (입력이 없으면 처음부터 진행)
start_row_input = input("진행을 시작할 행 번호를 입력하세요 (default: 처음부터): ").strip()
if start_row_input:
    try:
        # 사용자 입력이 1부터 시작한다고 가정 (1을 빼서 0 기반 인덱스로 변환)
        start_row = int(start_row_input) - 1
        if start_row < 0 or start_row >= len(df):
            print("입력한 행 번호가 유효하지 않습니다. 처음부터 진행합니다.")
            start_row = 0
    except ValueError:
        print("유효한 숫자를 입력하지 않았습니다. 처음부터 진행합니다.")
        start_row = 0
else:
    start_row = 0

print(f"[INFO] {start_row+1}번째 행부터 진행합니다.")

# ✅ Chrome Beta 실행
CHROME_BETA_PATH = r"C:\Program Files\Google\Chrome Beta\Application\chrome.exe"
CHATGPT_IMAGE_GENERATOR_URL = "https://chatgpt.com/g/g-pmuQfob8d-image-generator/c/67acb726-6aa4-8002-bb2d-d80890c4de97"

print("[🚀] Chrome Beta 실행 중...")
app = Application().start(f'"{CHROME_BETA_PATH}" {CHATGPT_IMAGE_GENERATOR_URL}')
time.sleep(5)  # 브라우저 로딩 대기

# ✅ Chrome 창 찾기
chrome_windows = [win for win in gw.getWindowsWithTitle("image generator") if win.isActive]
if not chrome_windows:
    print("[❌] Chrome 창을 찾을 수 없습니다.")
    exit()

chrome_window = chrome_windows[0]  # 첫 번째 Chrome 창 선택
chrome_window.activate()
time.sleep(2)

print("[✅] ChatGPT Image Generator 페이지 열림!")

# ✅ 다운로드 폴더 및 저장 폴더 설정
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")
SAVE_FOLDER = "generated_images"
os.makedirs(SAVE_FOLDER, exist_ok=True)

# ✅ DALLE 프롬프트 입력 후 이미지 생성 (입력한 행부터 시작)
for idx, (i, row) in enumerate(df.iloc[start_row:].iterrows(), start=start_row):
    item_1 = row["item_spain_01"].lower().replace(" ", "_")
    item_2 = row["item_spain_02"].lower().replace(" ", "_")
    
    prompt_1 = row["item_prompt_01"]
    prompt_2 = row["item_prompt_02"]

    for j, (item, prompt) in enumerate([(item_1, prompt_1), (item_2, prompt_2)]):
        if prompt:
            try:
                print(f"[🎨] {idx+1}/{len(df)}: '{item}' 이미지 생성 중...")
                
                # ✅ 띄어쓰기 변환 적용
                formatted_prompt = prompt.replace(" ", "{SPACE}")

                # ✅ 프롬프트 입력 및 이미지 생성 명령어 실행
                send_keys(formatted_prompt + "+{ENTER}+{ENTER}only{SPACE}response{SPACE}image.{SPACE}NO{SPACE}MESSAGE")
                time.sleep(1)
                send_keys("{ENTER}")  # 엔터 키 입력
                time.sleep(40)  # 이미지 생성 대기

                # ✅ 다운로드 버튼 클릭 전, 기존 파일 목록 저장
                existing_files = set(os.listdir(DOWNLOAD_FOLDER))
                
                # ✅ 좌표를 이용해 다운로드 버튼 클릭
                download_button_coords = (3338, 289)
                pyautogui.moveTo(download_button_coords[0], download_button_coords[1], duration=0.5)
                time.sleep(2)
                pyautogui.click()
                print("[✅] 좌표를 통한 다운로드 버튼 클릭 완료")
                
                # 다운로드 완료 대기
                time.sleep(10)
                
                # ✅ 다운로드 완료 후, 새로 추가된 파일 찾기
                new_files = set(os.listdir(DOWNLOAD_FOLDER)) - existing_files
                downloaded_file = None
                for file in new_files:
                    if file.endswith(".webp"):
                        downloaded_file = os.path.join(DOWNLOAD_FOLDER, file)
                        break

                if downloaded_file:
                    # ✅ 파일 이동 및 이름 변경
                    new_filename = os.path.join(SAVE_FOLDER, f"{idx}_{item}.webp")
                    shutil.move(downloaded_file, new_filename)
                    print(f"[✅] 이미지 저장 완료: {new_filename}")
                else:
                    print("[❌] webp 파일을 찾을 수 없습니다.")

                time.sleep(3)

            except Exception as e:
                print(f"[❌] 오류 발생 (아이템: {item}): {e}")

print("[🎉] 모든 이미지 생성 완료")
