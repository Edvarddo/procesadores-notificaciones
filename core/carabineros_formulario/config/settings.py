import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / ".env"
PAUSA_CORTA = 0.2
PAUSA_MEDIA = 0.5
PAUSA_LARGA = 1.0
#print("BASE_DIR:", BASE_DIR)
#print("ENV_PATH:", ENV_PATH)
#print("EXISTE .env?:", ENV_PATH.exists())

load_dotenv(dotenv_path=ENV_PATH)

CINJ_URL = os.getenv("CINJ_URL", "http://www.cinj.pjud/#/login")
CINJ_USER = os.getenv("CINJ_USER", "")
CINJ_PASS = os.getenv("CINJ_PASS", "")
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
TIMEOUT = int(os.getenv("TIMEOUT", "15"))

CSV_PATH = os.getenv("CSV_PATH", str(BASE_DIR / "data" / "registros.csv"))

OUTPUT_LOGS_DIR = str(BASE_DIR / "outputs" / "logs")
OUTPUT_RESULTADOS_DIR = str(BASE_DIR / "outputs" / "resultados")

