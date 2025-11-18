import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BASE_TEMPLATE_DIR = os.path.join(BASE_DIR, "Report", "template")
BASE_IMAGE_DIR = os.path.join(BASE_DIR, "Report")
OUTPUT_DIR_WITH_IMAGES = os.path.join(BASE_DIR, "Report", "completed_with_images")
OUTPUT_DIR = os.path.join(BASE_DIR, "Report", "completed_final")

GRAFANA_URL = os.getenv("GRAFANA_URL", "http://localhost:3000").rstrip('/')
API_KEY = os.getenv("GRAFANA_API_KEY", "")
VERIFY_SSL = os.getenv("GRAFANA_VERIFY_SSL", "true").lower() in ("true", "1", "yes")

DASHBOARD_MAP = {
    "kpmo": "dejkgjz0jnoqoa",
    "GIT": "aejkgkoze5nggb",
    "hansystem": "cejnb5yyuk5q8e",
    "humecca": "bejnb5db19blse",
    "klcns": "eejnb31cylreod",
    "sungwoo": "cejnb4aafury8e",
    "thepnl": "fejkgid897xtsc",
    "프리스타일": "fejkgfwux1fy8c"
}

SENTENCE_TEMPLATE = "사용량 최대 {max}%, 평균 {mean}% 입니다."
