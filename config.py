import os
import json

REPORT_BASE_DIR = os.getenv("REPORT_BASE_DIR", "/root/Report")

BASE_TEMPLATE_DIR = os.path.join(REPORT_BASE_DIR, "template")
BASE_IMAGE_DIR = REPORT_BASE_DIR
OUTPUT_DIR_WITH_IMAGES = os.path.join(REPORT_BASE_DIR, "completed_with_images")
OUTPUT_DIR = os.path.join(REPORT_BASE_DIR, "completed_final")

GRAFANA_URL = os.getenv("GRAFANA_URL", "http://localhost:3000").rstrip('/')
API_KEY = os.getenv("GRAFANA_API_KEY", "glsa_8RKuPs2USwIwxdafke3r4bcs93zkGO4E_462d232d")
VERIFY_SSL = os.getenv("GRAFANA_VERIFY_SSL", "true").lower() in ("true", "1", "yes")

DASHBOARD_MAPPING_FILE = os.path.join(os.path.dirname(__file__), "dashboard_mapping.json")

def load_dashboard_map():
    """JSON 파일에서 대시보드 매핑 로드
    
    구조 (신규):
    {
        "customer_name": {
            "dashboard_uid": "xxx",
            "display_name": "표시명",
            "contact_name": "담당자 이름",
            "contact_phone": "연락처",
            "contact_email": "이메일"
        }
    }
    
    구조 (레거시 - 자동 마이그레이션):
    {
        "customer_name": "dashboard_uid"
    }
    """
    if os.path.exists(DASHBOARD_MAPPING_FILE):
        try:
            with open(DASHBOARD_MAPPING_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                migrated = False
                for key, value in list(data.items()):
                    if isinstance(value, str):
                        data[key] = {
                            "dashboard_uid": value, 
                            "display_name": "",
                            "contact_name": "",
                            "contact_phone": "",
                            "contact_email": ""
                        }
                        migrated = True
                    elif isinstance(value, dict):
                        if "contact_name" not in value:
                            value["contact_name"] = ""
                            migrated = True
                        if "contact_phone" not in value:
                            value["contact_phone"] = ""
                            migrated = True
                        if "contact_email" not in value:
                            value["contact_email"] = ""
                            migrated = True
                if migrated:
                    save_dashboard_map(data)
                return data
        except Exception:
            pass
    return {}

def save_dashboard_map(mapping):
    """대시보드 매핑을 JSON 파일에 저장"""
    with open(DASHBOARD_MAPPING_FILE, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=4)


def get_all_customers():
    """모든 고객사 목록 조회 (폴더 + dashboard_map 병합)"""
    customers = []
    dashboard_map = load_dashboard_map()
    
    if os.path.exists(BASE_IMAGE_DIR):
        for item in os.listdir(BASE_IMAGE_DIR):
            item_path = os.path.join(BASE_IMAGE_DIR, item)
            if os.path.isdir(item_path) and item not in ['template', 'completed_with_images', 'completed_final']:
                customer_data = dashboard_map.get(item, {})
                if isinstance(customer_data, str):
                    customer_data = {"dashboard_uid": customer_data, "display_name": "", "contact_name": "", "contact_phone": "", "contact_email": ""}
                customers.append({
                    "name": item,
                    "display_name": customer_data.get('display_name', '') or '',
                    "dashboard_uid": customer_data.get('dashboard_uid', '') or '',
                    "contact_name": customer_data.get('contact_name', '') or '',
                    "contact_phone": customer_data.get('contact_phone', '') or '',
                    "contact_email": customer_data.get('contact_email', '') or ''
                })
    
    for customer_name, customer_data in dashboard_map.items():
        if not any(c["name"] == customer_name for c in customers):
            if isinstance(customer_data, str):
                customer_data = {"dashboard_uid": customer_data, "display_name": "", "contact_name": "", "contact_phone": "", "contact_email": ""}
            customers.append({
                "name": customer_name,
                "display_name": customer_data.get('display_name', '') or '',
                "dashboard_uid": customer_data.get('dashboard_uid', '') or '',
                "contact_name": customer_data.get('contact_name', '') or '',
                "contact_phone": customer_data.get('contact_phone', '') or '',
                "contact_email": customer_data.get('contact_email', '') or ''
            })
    
    return sorted(customers, key=lambda x: x["name"])


def get_customer_data(customer_name):
    """고객사 전체 데이터 조회"""
    mapping = load_dashboard_map()
    data = mapping.get(customer_name, {})
    if isinstance(data, str):
        return {
            "dashboard_uid": data, 
            "display_name": "",
            "contact_name": "",
            "contact_phone": "",
            "contact_email": ""
        }
    result = {
        "dashboard_uid": data.get("dashboard_uid", ""),
        "display_name": data.get("display_name", ""),
        "contact_name": data.get("contact_name", ""),
        "contact_phone": data.get("contact_phone", ""),
        "contact_email": data.get("contact_email", "")
    }
    return result

def get_dashboard_uid(customer_name):
    """고객사 이름으로 대시보드 UID 조회"""
    data = get_customer_data(customer_name)
    return data.get('dashboard_uid', '')

def get_display_name(customer_name):
    """고객사의 표시명 조회"""
    data = get_customer_data(customer_name)
    return data.get('display_name', '')

def set_dashboard_uid(customer_name, dashboard_uid):
    """고객사의 대시보드 UID 설정"""
    mapping = load_dashboard_map()
    if customer_name not in mapping:
        mapping[customer_name] = {"dashboard_uid": "", "display_name": ""}
    elif isinstance(mapping[customer_name], str):
        mapping[customer_name] = {"dashboard_uid": mapping[customer_name], "display_name": ""}
    mapping[customer_name]["dashboard_uid"] = dashboard_uid
    save_dashboard_map(mapping)

def set_display_name(customer_name, display_name):
    """고객사의 표시명 설정"""
    mapping = load_dashboard_map()
    if customer_name not in mapping:
        mapping[customer_name] = {"dashboard_uid": "", "display_name": ""}
    elif isinstance(mapping[customer_name], str):
        mapping[customer_name] = {"dashboard_uid": mapping[customer_name], "display_name": ""}
    mapping[customer_name]["display_name"] = display_name
    save_dashboard_map(mapping)

def set_customer_data(customer_name, dashboard_uid=None, display_name=None, 
                      contact_name=None, contact_phone=None, contact_email=None):
    """고객사 데이터 일괄 설정"""
    mapping = load_dashboard_map()
    if customer_name not in mapping:
        mapping[customer_name] = {
            "dashboard_uid": "", 
            "display_name": "",
            "contact_name": "",
            "contact_phone": "",
            "contact_email": ""
        }
    elif isinstance(mapping[customer_name], str):
        mapping[customer_name] = {
            "dashboard_uid": mapping[customer_name], 
            "display_name": "",
            "contact_name": "",
            "contact_phone": "",
            "contact_email": ""
        }
    
    if dashboard_uid is not None:
        mapping[customer_name]["dashboard_uid"] = dashboard_uid
    if display_name is not None:
        mapping[customer_name]["display_name"] = display_name
    if contact_name is not None:
        mapping[customer_name]["contact_name"] = contact_name
    if contact_phone is not None:
        mapping[customer_name]["contact_phone"] = contact_phone
    if contact_email is not None:
        mapping[customer_name]["contact_email"] = contact_email
    save_dashboard_map(mapping)

def delete_dashboard_mapping(customer_name):
    """고객사의 대시보드 매핑 삭제"""
    mapping = load_dashboard_map()
    if customer_name in mapping:
        del mapping[customer_name]
        save_dashboard_map(mapping)
        return True
    return False

def delete_customer_metadata(customer_name):
    """고객사 메타데이터 삭제 (delete_dashboard_mapping과 동일)"""
    return delete_dashboard_mapping(customer_name)

def get_customer_info(customer_name):
    """고객사 전체 정보 조회"""
    data = get_customer_data(customer_name)
    return {
        'name': customer_name,
        'dashboard_uid': data.get('dashboard_uid', ''),
        'display_name': data.get('display_name', ''),
        'contact_name': data.get('contact_name', ''),
        'contact_phone': data.get('contact_phone', ''),
        'contact_email': data.get('contact_email', '')
    }

def find_customer_by_name_or_display(search_name):
    """name 또는 display_name으로 고객사 검색"""
    if os.path.exists(BASE_IMAGE_DIR):
        dirs = [d for d in os.listdir(BASE_IMAGE_DIR) 
                if os.path.isdir(os.path.join(BASE_IMAGE_DIR, d)) 
                and d not in ['template', 'completed_with_images', 'completed_final']]
        if search_name in dirs:
            return search_name
    
    mapping = load_dashboard_map()
    for name, data in mapping.items():
        if isinstance(data, dict) and data.get('display_name') == search_name:
            return name
    
    return None

def load_customer_metadata():
    """고객사 메타데이터 로드 (dashboard_map과 통합됨)"""
    mapping = load_dashboard_map()
    result = {}
    for name, data in mapping.items():
        if isinstance(data, dict):
            result[name] = {
                "dashboard_uid": data.get("dashboard_uid", ""),
                "display_name": data.get("display_name", ""),
                "contact_name": data.get("contact_name", ""),
                "contact_phone": data.get("contact_phone", ""),
                "contact_email": data.get("contact_email", "")
            }
        else:
            result[name] = {
                "dashboard_uid": data, 
                "display_name": "",
                "contact_name": "",
                "contact_phone": "",
                "contact_email": ""
            }
    return result

DASHBOARD_MAP = load_dashboard_map()

SENTENCE_TEMPLATE = "사용량 최대 {max}%, 평균 {mean}% 입니다."
