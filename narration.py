import os
from dotenv import load_dotenv
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gtts import gTTS

def load_sheet_data() -> pd.DataFrame:
    """
    .env 파일의 환경 변수를 로드한 후, Google Sheets API를 통해
    지정된 시트에서 데이터를 읽어 DataFrame으로 반환합니다.
    
    시트의 첫 번째 행은 컬럼명으로 사용되며, 반드시 "item_01"과 "item_02" 컬럼이 있어야 합니다.
    """
    load_dotenv()
    SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_CONSOLE_KEY_PATH")
    SHEET_ID = os.getenv("PROMPT_SHEET_ID")
    SHEET_NAME = os.getenv("PROMPT_SHEET_NAME")
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def generate_narrations(df: pd.DataFrame, output_folder: str = "narrations", lang: str = "es") -> None:
    """
    DataFrame의 각 행에서 item_01과 item_02를 결합하여
    "item_01 VS item_02. ¿Cuál te gusta más?" 형태의 나레이션 텍스트를 생성하고,
    이를 gTTS를 사용해 mp3 파일로 저장합니다.
    
    Args:
        df (pd.DataFrame): Google Sheets에서 읽어온 데이터 (컬럼: "item_01", "item_02")
        output_folder (str): 생성된 mp3 파일을 저장할 폴더 경로 (기본값: "narrations")
        lang (str): TTS에 사용할 언어 코드 (예: "es" for Spanish)
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for index, row in df.iterrows():
        item1 = row["item_01"]
        item2 = row["item_02"]
        narration_text = f"{item1} VS {item2}. ¿Cuál te gusta más?"
        try:
            tts = gTTS(text=narration_text, lang=lang)
            file_path = os.path.join(output_folder, f"row{index+1}_narration.mp3")
            tts.save(file_path)
            print(f"[✅] Row {index+1} 나레이션 생성 완료: {file_path}")
        except Exception as e:
            print(f"[❌] Row {index+1} 오류: {e}")

if __name__ == "__main__":
    df = load_sheet_data()
    generate_narrations(df, output_folder="narrations", lang="es")
    print("모든 나레이션 생성 완료!")
