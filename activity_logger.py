import os
import json
import uuid
import threading
from datetime import datetime

LOGS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
LOGS_FILE = os.path.join(LOGS_DIR, 'activity_logs.json')
MAX_LOGS = 200

_lock = threading.Lock()


def _ensure_dir():
    os.makedirs(LOGS_DIR, exist_ok=True)


def _load_logs():
    _ensure_dir()
    if not os.path.exists(LOGS_FILE):
        return []
    try:
        with open(LOGS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return []


def _save_logs(logs):
    _ensure_dir()
    if len(logs) > MAX_LOGS:
        logs = logs[-MAX_LOGS:]
    with open(LOGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)


def create_log(action_type, customer_name=None):
    log_entry = {
        "id": str(uuid.uuid4())[:8],
        "action_type": action_type,
        "customer": customer_name or "전체",
        "status": "running",
        "start_time": datetime.now().isoformat(),
        "end_time": None,
        "duration_seconds": None,
        "summary": {},
        "detail_logs": [],
        "errors": []
    }
    with _lock:
        logs = _load_logs()
        logs.append(log_entry)
        _save_logs(logs)
    return log_entry["id"]


def add_detail(log_id, message):
    with _lock:
        logs = _load_logs()
        for entry in logs:
            if entry["id"] == log_id:
                entry["detail_logs"].append({
                    "time": datetime.now().isoformat(),
                    "message": message
                })
                _save_logs(logs)
                return True
    return False


def add_detail_batch(log_id, messages):
    if not messages:
        return False
    with _lock:
        logs = _load_logs()
        for entry in logs:
            if entry["id"] == log_id:
                now = datetime.now().isoformat()
                for msg in messages:
                    entry["detail_logs"].append({
                        "time": now,
                        "message": msg
                    })
                _save_logs(logs)
                return True
    return False


def add_error(log_id, error_message):
    with _lock:
        logs = _load_logs()
        for entry in logs:
            if entry["id"] == log_id:
                entry["errors"].append({
                    "time": datetime.now().isoformat(),
                    "message": error_message
                })
                _save_logs(logs)
                return True
    return False


def complete_log(log_id, success=True, summary=None):
    with _lock:
        logs = _load_logs()
        for entry in logs:
            if entry["id"] == log_id:
                entry["status"] = "success" if success else "failed"
                entry["end_time"] = datetime.now().isoformat()
                start = datetime.fromisoformat(entry["start_time"])
                end = datetime.fromisoformat(entry["end_time"])
                entry["duration_seconds"] = round((end - start).total_seconds(), 1)
                if summary:
                    entry["summary"] = summary
                _save_logs(logs)
                return True
    return False


def get_logs(action_type=None, status=None, limit=50):
    logs = _load_logs()
    logs = list(reversed(logs))

    if action_type:
        logs = [l for l in logs if l.get("action_type") == action_type]
    if status:
        logs = [l for l in logs if l.get("status") == status]

    return logs[:limit]


def get_log_detail(log_id):
    logs = _load_logs()
    for entry in logs:
        if entry["id"] == log_id:
            return entry
    return None


def delete_log(log_id):
    with _lock:
        logs = _load_logs()
        original_len = len(logs)
        logs = [l for l in logs if l["id"] != log_id]
        if len(logs) < original_len:
            _save_logs(logs)
            return True
    return False


def clear_logs():
    with _lock:
        _save_logs([])
    return True
