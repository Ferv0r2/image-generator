import time
import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import pygetwindow as gw
import requests

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

# ✅ Chrome Beta 실행
CHROME_BETA_PATH = r"C:\Program Files\Google\Chrome Beta\Application\chrome.exe"
CHATGPT_IMAGE_GENERATOR_URL = "https://chatgpt.com/g/g-pmuQfob8d-image-generator"

print("[🚀] Chrome Beta 실행 중...")
app = Application().start(f'"{CHROME_BETA_PATH}" {CHATGPT_IMAGE_GENERATOR_URL}')
time.sleep(5)  # 브라우저 로딩 대기

# ✅ Chrome 창 찾기
chrome_windows = [win for win in gw.getWindowsWithTitle("ChatGPT") if win.isActive]
if not chrome_windows:
    print("[❌] Chrome 창을 찾을 수 없습니다.")
    exit()

chrome_window = chrome_windows[0]  # 첫 번째 Chrome 창 선택
chrome_window.activate()
time.sleep(2)

print("[✅] ChatGPT Image Generator 페이지 열림!")


# ✅ DALLE 프롬프트 입력 후 이미지 생성
for i, row in df.iterrows():
    item_1 = row["item_01"].lower().replace(" ", "_")
    item_2 = row["item_02"].lower().replace(" ", "_")
    
    prompt_1 = row["item_prompt_01"]
    prompt_2 = row["item_prompt_02"]

    for j, (item, prompt) in enumerate([(item_1, prompt_1), (item_2, prompt_2)]):
        if prompt:
            try:
                print(f"[🎨] {i+1}/{len(df)}: '{item}' 이미지 생성 중...")
                
                # ✅ 띄어쓰기 변환 적용
                formatted_prompt = prompt.replace(" ", "{SPACE}")

                # ✅ 프롬프트 입력
                send_keys(formatted_prompt)
                time.sleep(1)
                send_keys("{ENTER}")  # 엔터 키 입력
                time.sleep(30)  # 이미지 생성 대기

                # ✅ 생성된 이미지 다운로드 (스크린샷 방식)
                filename = f"generated_images/{i}_{item}.png"
                chrome_window.screenshot(filename)

                print(f"[✅] 이미지 생성 및 저장 완료: {filename}")
                time.sleep(3)

            except Exception as e:
                print(f"[❌] 오류 발생 (아이템: {item}): {e}")

print("[🎉] 모든 이미지 생성 완료")
