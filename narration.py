import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gtts import gTTS

def load_sheet_to_df() -> pd.DataFrame:
    """
    Google Sheets API를 이용해 지정된 시트에서 데이터를 읽어 DataFrame으로 반환합니다.
    시트의 첫 행은 컬럼명으로 사용합니다.
    """
    # .env 파일 로드
    load_dotenv()
    
    SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_CONSOLE_KEY_PATH")
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    
    # Google Sheets API 인증
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    
    # 시트 정보 가져오기
    SHEET_ID = os.getenv("PROMPT_SHEET_ID")
    SHEET_NAME = os.getenv("RESULT_SHEET_NAME")
    
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_values()
    
    # 첫 번째 행을 컬럼명으로 사용
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def generate_vs_narrations(df: pd.DataFrame, output_folder: str = "narrations_vs", lang: str = "es") -> None:
    """
    DataFrame의 각 행에서 item_01과 item_02를 결합해 "item_01 VS item_02" 형태의 텍스트를 생성하고,
    이를 gTTS를 사용해 mp3 파일로 저장합니다.
    
    Args:
        df (pd.DataFrame): Google Sheets에서 불러온 데이터 (컬럼: "item_01", "item_02")
        output_folder (str, optional): 생성된 mp3 파일을 저장할 폴더. Defaults to "narrations_vs".
        lang (str, optional): TTS에 사용할 언어 코드 (예: "es" for Spanish). Defaults to "es".
    """
    # 출력 폴더가 없으면 생성
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 각 행마다 "item_01 VS item_02" 텍스트로 나레이션 생성
    for index, row in df.iterrows():
        item1 = row["item_01"]
        item2 = row["item_02"]
        combined_text = f"{item1} VS {item2}"
        
        try:
            tts = gTTS(text=combined_text, lang=lang)
            file_path = os.path.join(output_folder, f"row{index+1}_vs.mp3")
            tts.save(file_path)
            print(f"[✅] Row {index+1} 나레이션 생성 완료: {file_path}")
        except Exception as e:
            print(f"[❌] Row {index+1} 오류: {e}")

if __name__ == "__main__":
    # Google Sheets에서 데이터 로드
    df = load_sheet_to_df()
    
    # "item_01 VS item_02" 형태의 나레이션 생성
    generate_vs_narrations(df, output_folder="narrations_vs", lang="es")
    print("모든 나레이션 생성 완료!")
