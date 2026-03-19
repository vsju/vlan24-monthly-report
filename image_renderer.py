import os
import re
import requests
from datetime import datetime, timezone, timedelta
import logging

try:
    from . import config
except ImportError:
    import config

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

KST = timezone(timedelta(hours=9))
DEFAULT_WIDTH = 640
DEFAULT_HEIGHT = 300

_WIN_INVALID = re.compile(r'[\\/:\*\?"<>\|]')


def safe_filename(name):
    return _WIN_INVALID.sub('_', name)


def get_previous_month_timestamps():
    today = datetime.now(KST)
    first_day_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_last_month = first_day_this_month - timedelta(seconds=1)
    first_day_last_month = last_day_last_month.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
    from_ts = int(first_day_last_month.timestamp() * 1000)
    to_ts = int(last_day_last_month.timestamp() * 1000)
    return from_ts, to_ts


def get_dashboard_info(dashboard_uid):
    url = f"{config.GRAFANA_URL}/api/dashboards/uid/{dashboard_uid}"
    headers = {"Authorization": f"Bearer {config.API_KEY}"}
    
    try:
        resp = requests.get(url, headers=headers, timeout=10, verify=config.VERIFY_SSL)
        resp.raise_for_status()
        data = resp.json()
        dashboard = data.get("dashboard", {})
        meta = data.get("meta", {})
        return {
            "title": dashboard.get("title", ""),
            "uid": dashboard_uid,
            "slug": meta.get("slug", dashboard_uid),
            "panels": dashboard.get("panels", [])
        }
    except requests.exceptions.RequestException as e:
        logging.error(f"대시보드 정보 조회 실패: {e}")
        return None


def extract_panels_from_dashboard(panels, row_map=None, current_row=None):
    if row_map is None:
        row_map = {}
    
    for p in panels:
        panel_type = p.get("type")
        panel_title = p.get("title", f"panel_{p.get('id', 'unknown')}")
        
        if panel_type == "row":
            row_title = safe_filename(p.get("title", "Unnamed_Row")).replace(" ", "_")
            row_map[row_title] = row_map.get(row_title, [])
            current_row = row_title
            
            sub_panels = p.get("panels", [])
            if sub_panels:
                extract_panels_from_dashboard(sub_panels, row_map, current_row)
        elif "id" in p:
            panel_info = {
                "id": p["id"],
                "title": panel_title,
                "title_safe": safe_filename(panel_title).replace(" ", "_"),
                "type": panel_type
            }
            if current_row:
                row_map[current_row].append(panel_info)
            else:
                row_map.setdefault("기타", []).append(panel_info)
    
    return row_map


def get_panels_for_customer(customer_name):
    dashboard_uid = config.get_dashboard_uid(customer_name)
    if not dashboard_uid:
        return {"success": False, "error": f"고객사 '{customer_name}'의 대시보드 UID가 설정되지 않았습니다."}
    
    dashboard_info = get_dashboard_info(dashboard_uid)
    if not dashboard_info:
        return {"success": False, "error": f"대시보드 정보를 가져올 수 없습니다. (UID: {dashboard_uid})"}
    
    panel_rows = extract_panels_from_dashboard(dashboard_info.get("panels", []))
    
    all_panels = []
    for row_title, panels in panel_rows.items():
        for panel in panels:
            all_panels.append({
                "id": panel["id"],
                "title": panel["title"],
                "row": row_title
            })
    
    return {
        "success": True,
        "customer_name": customer_name,
        "dashboard_uid": dashboard_uid,
        "dashboard_title": dashboard_info.get("title", ""),
        "dashboard_slug": dashboard_info.get("slug", ""),
        "panel_rows": panel_rows,
        "all_panels": all_panels,
        "total_panels": len(all_panels)
    }


def render_panel_image(dashboard_uid, dashboard_slug, panel_id, save_path, 
                        from_ts=None, to_ts=None, width=DEFAULT_WIDTH, height=DEFAULT_HEIGHT):
    if from_ts is None or to_ts is None:
        from_ts, to_ts = get_previous_month_timestamps()
    
    headers = {"Authorization": f"Bearer {config.API_KEY}"}
    params = {
        "from": from_ts,
        "to": to_ts,
        "panelId": panel_id,
        "width": width,
        "height": height,
        "tz": "Asia/Seoul"
    }
    
    url = f"{config.GRAFANA_URL}/render/d-solo/{dashboard_uid}/{dashboard_slug}"
    
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=60, verify=config.VERIFY_SSL)
        if resp.status_code == 200:
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            with open(save_path, "wb") as f:
                f.write(resp.content)
            return {"success": True, "path": save_path}
        else:
            return {"success": False, "error": f"렌더링 실패 (status: {resp.status_code})"}
    except requests.exceptions.RequestException as e:
        return {"success": False, "error": str(e)}


def render_all_panels(customer_name, log_func=None):
    if log_func is None:
        log_func = logging.info
    
    result = {
        "success": False,
        "customer_name": customer_name,
        "rendered": [],
        "failed": [],
        "total": 0
    }
    
    panels_info = get_panels_for_customer(customer_name)
    if not panels_info.get("success"):
        result["error"] = panels_info.get("error", "패널 정보 조회 실패")
        return result
    
    dashboard_uid = panels_info["dashboard_uid"]
    dashboard_slug = panels_info["dashboard_slug"]
    panel_rows = panels_info["panel_rows"]
    
    save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
    from_ts, to_ts = get_previous_month_timestamps()
    
    total_count = 0
    for row_title, panels in panel_rows.items():
        row_dir = os.path.join(save_base_dir, row_title)
        os.makedirs(row_dir, exist_ok=True)
        
        for panel in panels:
            panel_id = panel["id"]
            panel_title_safe = panel["title_safe"]
            filename = f"{panel_title_safe}_{panel_id}.png"
            save_path = os.path.join(row_dir, filename)
            
            log_func(f"렌더링 중: {row_title}/{panel['title']} (ID: {panel_id})")
            
            render_result = render_panel_image(
                dashboard_uid, dashboard_slug, panel_id, save_path,
                from_ts=from_ts, to_ts=to_ts
            )
            
            total_count += 1
            if render_result.get("success"):
                result["rendered"].append({
                    "panel_id": panel_id,
                    "title": panel["title"],
                    "row": row_title,
                    "path": save_path
                })
                log_func(f"  저장 완료: {save_path}")
            else:
                result["failed"].append({
                    "panel_id": panel_id,
                    "title": panel["title"],
                    "row": row_title,
                    "error": render_result.get("error", "알 수 없는 오류")
                })
                log_func(f"  실패: {render_result.get('error')}")
    
    result["total"] = total_count
    result["success"] = len(result["failed"]) == 0
    result["save_dir"] = save_base_dir
    return result


def render_selected_panels(customer_name, panel_ids, log_func=None):
    if log_func is None:
        log_func = logging.info
    
    result = {
        "success": False,
        "customer_name": customer_name,
        "rendered": [],
        "failed": [],
        "total": len(panel_ids)
    }
    
    if not panel_ids:
        result["error"] = "선택된 패널이 없습니다."
        return result
    
    panels_info = get_panels_for_customer(customer_name)
    if not panels_info.get("success"):
        result["error"] = panels_info.get("error", "패널 정보 조회 실패")
        return result
    
    dashboard_uid = panels_info["dashboard_uid"]
    dashboard_slug = panels_info["dashboard_slug"]
    panel_rows = panels_info["panel_rows"]
    
    panel_id_set = set(int(pid) for pid in panel_ids)
    panel_map = {}
    for row_title, panels in panel_rows.items():
        for panel in panels:
            if panel["id"] in panel_id_set:
                panel_map[panel["id"]] = {**panel, "row": row_title}
    
    save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
    from_ts, to_ts = get_previous_month_timestamps()
    
    for panel_id in panel_ids:
        panel_id_int = int(panel_id)
        panel_info = panel_map.get(panel_id_int)
        
        if not panel_info:
            result["failed"].append({
                "panel_id": panel_id_int,
                "error": "패널을 찾을 수 없습니다."
            })
            continue
        
        row_title = panel_info["row"]
        row_dir = os.path.join(save_base_dir, row_title)
        os.makedirs(row_dir, exist_ok=True)
        
        filename = f"{panel_info['title_safe']}_{panel_id_int}.png"
        save_path = os.path.join(row_dir, filename)
        
        log_func(f"렌더링 중: {row_title}/{panel_info['title']} (ID: {panel_id_int})")
        
        render_result = render_panel_image(
            dashboard_uid, dashboard_slug, panel_id_int, save_path,
            from_ts=from_ts, to_ts=to_ts
        )
        
        if render_result.get("success"):
            result["rendered"].append({
                "panel_id": panel_id_int,
                "title": panel_info["title"],
                "row": row_title,
                "path": save_path
            })
            log_func(f"  저장 완료: {save_path}")
        else:
            result["failed"].append({
                "panel_id": panel_id_int,
                "title": panel_info["title"],
                "row": row_title,
                "error": render_result.get("error", "알 수 없는 오류")
            })
            log_func(f"  실패: {render_result.get('error')}")
    
    result["success"] = len(result["failed"]) == 0
    result["save_dir"] = save_base_dir
    return result


def render_all_panels_stream(customer_name):
    panels_info = get_panels_for_customer(customer_name)
    if not panels_info.get("success"):
        yield {"type": "error", "error": panels_info.get("error", "패널 정보 조회 실패")}
        return

    dashboard_uid = panels_info["dashboard_uid"]
    dashboard_slug = panels_info["dashboard_slug"]
    panel_rows = panels_info["panel_rows"]

    all_panels_list = []
    for row_title, panels in panel_rows.items():
        for panel in panels:
            all_panels_list.append({**panel, "row": row_title})

    total = len(all_panels_list)
    yield {"type": "init", "total": total, "customer": customer_name}

    if total == 0:
        save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
        yield {"type": "complete", "rendered": 0, "failed": 0,
               "total": 0, "save_dir": save_base_dir, "success": True,
               "message": "렌더링할 패널이 없습니다."}
        return

    save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
    from_ts, to_ts = get_previous_month_timestamps()

    rendered_count = 0
    failed_count = 0
    rendered_files = []

    for idx, panel in enumerate(all_panels_list):
        panel_id = panel["id"]
        row_title = panel["row"]
        row_dir = os.path.join(save_base_dir, row_title)
        os.makedirs(row_dir, exist_ok=True)

        filename = f"{panel['title_safe']}_{panel_id}.png"
        save_path = os.path.join(row_dir, filename)

        yield {"type": "progress", "current": idx + 1, "total": total,
               "panel": panel["title"], "row": row_title, "status": "rendering"}

        try:
            render_result = render_panel_image(
                dashboard_uid, dashboard_slug, panel_id, save_path,
                from_ts=from_ts, to_ts=to_ts
            )
        except Exception as e:
            render_result = {"success": False, "error": str(e)}

        if render_result.get("success"):
            rendered_count += 1
            rendered_files.append(save_path)
            yield {"type": "panel_done", "current": idx + 1, "total": total,
                   "panel_id": panel_id, "panel": panel["title"], "row": row_title,
                   "success": True, "rendered": rendered_count, "failed": failed_count}
        else:
            failed_count += 1
            yield {"type": "panel_done", "current": idx + 1, "total": total,
                   "panel_id": panel_id, "panel": panel["title"], "row": row_title,
                   "success": False, "error": render_result.get("error", "알 수 없는 오류"),
                   "rendered": rendered_count, "failed": failed_count}

    yield {"type": "complete", "rendered": rendered_count, "failed": failed_count,
           "total": total, "save_dir": save_base_dir,
           "rendered_files": rendered_files,
           "success": failed_count == 0}


def render_selected_panels_stream(customer_name, panel_ids):
    panels_info = get_panels_for_customer(customer_name)
    if not panels_info.get("success"):
        yield {"type": "error", "error": panels_info.get("error", "패널 정보 조회 실패")}
        return

    dashboard_uid = panels_info["dashboard_uid"]
    dashboard_slug = panels_info["dashboard_slug"]
    panel_rows = panels_info["panel_rows"]

    panel_id_set = set(int(pid) for pid in panel_ids)
    panel_map = {}
    for row_title, panels in panel_rows.items():
        for panel in panels:
            if panel["id"] in panel_id_set:
                panel_map[panel["id"]] = {**panel, "row": row_title}

    total = len(panel_ids)
    yield {"type": "init", "total": total, "customer": customer_name}

    if total == 0:
        save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
        yield {"type": "complete", "rendered": 0, "failed": 0,
               "total": 0, "save_dir": save_base_dir, "success": True,
               "message": "렌더링할 패널이 없습니다."}
        return

    save_base_dir = os.path.join(config.BASE_IMAGE_DIR, customer_name)
    from_ts, to_ts = get_previous_month_timestamps()

    rendered_count = 0
    failed_count = 0
    rendered_files = []

    for idx, pid in enumerate(panel_ids):
        panel_id_int = int(pid)
        panel_info = panel_map.get(panel_id_int)

        if not panel_info:
            failed_count += 1
            yield {"type": "panel_done", "current": idx + 1, "total": total,
                   "panel_id": panel_id_int, "panel": f"ID {panel_id_int}",
                   "success": False, "error": "패널을 찾을 수 없습니다.",
                   "rendered": rendered_count, "failed": failed_count}
            continue

        row_title = panel_info["row"]
        row_dir = os.path.join(save_base_dir, row_title)
        os.makedirs(row_dir, exist_ok=True)

        filename = f"{panel_info['title_safe']}_{panel_id_int}.png"
        save_path = os.path.join(row_dir, filename)

        yield {"type": "progress", "current": idx + 1, "total": total,
               "panel": panel_info["title"], "row": row_title, "status": "rendering"}

        try:
            render_result = render_panel_image(
                dashboard_uid, dashboard_slug, panel_id_int, save_path,
                from_ts=from_ts, to_ts=to_ts
            )
        except Exception as e:
            render_result = {"success": False, "error": str(e)}

        if render_result.get("success"):
            rendered_count += 1
            rendered_files.append(save_path)
            yield {"type": "panel_done", "current": idx + 1, "total": total,
                   "panel_id": panel_id_int, "panel": panel_info["title"], "row": row_title,
                   "success": True, "rendered": rendered_count, "failed": failed_count}
        else:
            failed_count += 1
            yield {"type": "panel_done", "current": idx + 1, "total": total,
                   "panel_id": panel_id_int, "panel": panel_info["title"], "row": row_title,
                   "success": False, "error": render_result.get("error", "알 수 없는 오류"),
                   "rendered": rendered_count, "failed": failed_count}

    yield {"type": "complete", "rendered": rendered_count, "failed": failed_count,
           "total": total, "save_dir": save_base_dir,
           "rendered_files": rendered_files,
           "success": failed_count == 0}


def render_all_customers_stream():
    customers = config.get_all_customers()
    customers_with_uid = [c for c in customers if c.get('dashboard_uid')]

    if not customers_with_uid:
        yield {"type": "error", "error": "대시보드 UID가 설정된 고객사가 없습니다."}
        return

    total_customers = len(customers_with_uid)
    yield {"type": "init", "total_customers": total_customers}

    total_rendered = 0
    total_failed = 0
    customers_rendered = []
    customers_failed = []

    for c_idx, customer in enumerate(customers_with_uid):
        customer_name = customer['name']
        yield {"type": "customer_start", "customer": customer_name,
               "customer_index": c_idx + 1, "total_customers": total_customers}

        try:
            panel_count = 0
            c_rendered = 0
            c_failed = 0
            c_files = []
            for event in render_all_panels_stream(customer_name):
                if event["type"] == "init":
                    panel_count = event["total"]
                    yield {"type": "customer_progress", "customer": customer_name,
                           "customer_index": c_idx + 1, "total_customers": total_customers,
                           "panel_total": panel_count, "panel_current": 0}
                elif event["type"] == "progress":
                    yield {"type": "customer_progress", "customer": customer_name,
                           "customer_index": c_idx + 1, "total_customers": total_customers,
                           "panel_total": panel_count,
                           "panel_current": event["current"],
                           "panel": event.get("panel", ""), "row": event.get("row", "")}
                elif event["type"] == "panel_done":
                    c_rendered = event.get("rendered", 0)
                    c_failed = event.get("failed", 0)
                elif event["type"] == "complete":
                    c_rendered = event.get("rendered", 0)
                    c_failed = event.get("failed", 0)
                    c_files = event.get("rendered_files", [])
                elif event["type"] == "error":
                    customers_failed.append({"name": customer_name, "error": event["error"]})
                    yield {"type": "customer_done", "customer": customer_name,
                           "customer_index": c_idx + 1, "total_customers": total_customers,
                           "success": False, "error": event["error"]}
                    continue

            total_rendered += c_rendered
            total_failed += c_failed
            customers_rendered.append({"name": customer_name, "rendered_count": c_rendered,
                                       "failed_count": c_failed, "rendered_files": c_files})
            yield {"type": "customer_done", "customer": customer_name,
                   "customer_index": c_idx + 1, "total_customers": total_customers,
                   "success": True, "rendered": c_rendered, "failed": c_failed,
                   "rendered_files": c_files}

        except Exception as e:
            customers_failed.append({"name": customer_name, "error": str(e)})
            yield {"type": "customer_done", "customer": customer_name,
                   "customer_index": c_idx + 1, "total_customers": total_customers,
                   "success": False, "error": str(e)}

    yield {"type": "complete", "total_rendered": total_rendered, "total_failed": total_failed,
           "customers_rendered": customers_rendered, "customers_failed": customers_failed,
           "success": len(customers_failed) == 0}


def render_all_customers(log_func=None):
    if log_func is None:
        log_func = logging.info
    
    result = {
        "success": False,
        "customers_rendered": [],
        "customers_failed": [],
        "total_panels_rendered": 0,
        "total_panels_failed": 0
    }
    
    customers = config.get_all_customers()
    customers_with_uid = [c for c in customers if c.get('dashboard_uid')]
    
    if not customers_with_uid:
        result["error"] = "대시보드 UID가 설정된 고객사가 없습니다."
        return result
    
    log_func(f"전체 고객사 렌더링 시작: {len(customers_with_uid)}개 고객사")
    
    for customer in customers_with_uid:
        customer_name = customer['name']
        log_func(f"\n=== {customer_name} 렌더링 시작 ===")
        
        try:
            customer_result = render_all_panels(customer_name, log_func=log_func)
            
            if customer_result.get("success") or customer_result.get("rendered"):
                result["customers_rendered"].append({
                    "name": customer_name,
                    "rendered_count": len(customer_result.get("rendered", [])),
                    "failed_count": len(customer_result.get("failed", []))
                })
                result["total_panels_rendered"] += len(customer_result.get("rendered", []))
                result["total_panels_failed"] += len(customer_result.get("failed", []))
                log_func(f"=== {customer_name} 완료: {len(customer_result.get('rendered', []))}개 성공 ===")
            else:
                result["customers_failed"].append({
                    "name": customer_name,
                    "error": customer_result.get("error", "알 수 없는 오류")
                })
                log_func(f"=== {customer_name} 실패: {customer_result.get('error')} ===")
        except Exception as e:
            result["customers_failed"].append({
                "name": customer_name,
                "error": str(e)
            })
            log_func(f"=== {customer_name} 예외 발생: {str(e)} ===")
    
    result["success"] = len(result["customers_failed"]) == 0
    return result
