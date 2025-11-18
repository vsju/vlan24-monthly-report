import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BASE_TEMPLATE_DIR = os.path.join(BASE_DIR, "Report", "template")
BASE_IMAGE_DIR = os.path.join(BASE_DIR, "Report")
OUTPUT_DIR_WITH_IMAGES = os.path.join(BASE_DIR, "Report", "completed_with_images")
OUTPUT_DIR = os.path.join(BASE_DIR, "Report", "completed_final")

def get_secret(key, default=""):
    """환경 변수 또는 .streamlit/secrets.toml에서 값 가져오기
    
    우선순위:
    1. 환경 변수 (os.getenv)
    2. .streamlit/secrets.toml 파일
    3. 기본값
    """
    env_value = os.getenv(key)
    if env_value is not None:
        return env_value
    
    try:
        secrets_file = os.path.join(BASE_DIR, ".streamlit", "secrets.toml")
        if os.path.exists(secrets_file):
            import toml
            secrets = toml.load(secrets_file)
            if key in secrets:
                return secrets[key]
    except Exception:
        pass
    
    try:
        import streamlit as st
        if hasattr(st, 'secrets') and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    
    return default

GRAFANA_URL = get_secret("GRAFANA_URL", "http://localhost:3000").rstrip('/')
API_KEY = get_secret("GRAFANA_API_KEY", "")
VERIFY_SSL = get_secret("GRAFANA_VERIFY_SSL", "true").lower() in ("true", "1", "yes")

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
