import os
import re
import json
import time
import datetime
import requests
import gc
from dateutil.relativedelta import relativedelta
from pptx import Presentation
from pptx.dml.color import RGBColor
import config

def normalize_title(title):
    if not title:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', title).lower()

def find_all_templates(root_dir):
    templates = []
    if not os.path.isdir(root_dir):
        return templates
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.pptx') and not filename.startswith('~$'):
                full_path = os.path.join(dirpath, filename)
                templates.append(full_path)
    return templates

def find_templates_for_customer(template_base_dir, customer_path):
    customer_template_dir = os.path.join(template_base_dir, customer_path)
    return find_all_templates(customer_template_dir)

def calculate_previous_month_dates():
    today = datetime.date.today()
    end_date = today.replace(day=1) - relativedelta(days=1)
    start_date = end_date.replace(day=1)
    
    start_date_str_kr = start_date.strftime("%Y년 %m월 %d일")
    end_date_str_kr = end_date.strftime("%Y년 %m월 %d일")
    start_date_str_hyphen = start_date.strftime("%Y-%m-%d")
    end_date_str_hyphen = end_date.strftime("%Y-%m-%d")
    
    y_str = end_date.strftime("%Y")
    m_str = end_date.strftime("%m")
    target_date_str = f"{y_str}.{int(m_str)}월"
    
    return {
        "placeholders": {
            "{{START_DATE}}": start_date_str_kr,
            "{{END_DATE}}": end_date_str_kr,
            "{{YEAR}}": y_str,
            "{{MONTH}}": m_str,
            "{{DATE_RANGE}}": f"{start_date_str_kr} ~ {end_date_str_kr}",
            "{{DATE_RANGE_HYPHEN}}": f"{start_date_str_hyphen} ~ {end_date_str_hyphen}"
        },
        "filename_date": end_date.strftime("%Y-%m"),
        "target_date_str": target_date_str,
        "start_ts": int(datetime.datetime.combine(start_date, datetime.time.min).timestamp() * 1000),
        "end_ts": int(datetime.datetime.combine(end_date, datetime.time.max).timestamp() * 1000)
    }

def find_all_panels_recursively(panel_list):
    all_panels = []
    for panel in panel_list:
        all_panels.append(panel)
        if panel.get("type") == "row" and "panels" in panel:
            all_panels.extend(find_all_panels_recursively(panel["panels"]))
    return all_panels

def get_dashboard_definition(dashboard_uid, retries=3, delay=1):
    url = f"{config.GRAFANA_URL.rstrip('/')}/api/dashboards/uid/{dashboard_uid}"
    headers = {"Authorization": f"Bearer {config.API_KEY}"}
    for i in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=20, verify=config.VERIFY_SSL)
            response.raise_for_status()
            dashboard = response.json().get('dashboard', {})
            dashboard['all_panels'] = find_all_panels_recursively(dashboard.get('panels', []))
            return dashboard
        except requests.exceptions.RequestException:
            if i < retries - 1:
                time.sleep(delay)
    return None

def find_panel_by_title(all_panels, title_from_placeholder):
    normalized_placeholder_title = normalize_title(title_from_placeholder)
    for panel in all_panels:
        panel_title_from_grafana = panel.get('title', '')
        normalized_grafana_title = normalize_title(panel_title_from_grafana)
        if normalized_grafana_title and normalized_grafana_title == normalized_placeholder_title:
            return panel
    return None

def get_grafana_stats_by_panel(panel, query_letter, start_ts, end_ts):
    headers = {"Authorization": f"Bearer {config.API_KEY}", "Content-Type": "application/json"}
    
    all_queries = panel.get('targets', [])
    selected = [q.copy() for q in all_queries if q.get('refId') == query_letter]
    if not selected:
        return None
    
    panel_datasource = panel.get('datasource')
    if panel_datasource:
        for q in selected:
            if 'datasource' not in q or q['datasource'] is None:
                q['datasource'] = panel_datasource
    
    for q in selected:
        q.pop('real_hosts', None)
    
    for q in selected:
        q.setdefault('maxDataPoints', 720)
        q.setdefault('intervalMs', 3600000)
    
    query_payload = {
        "queries": selected,
        "from": str(start_ts),
        "to": str(end_ts)
    }
    query_url = f"{config.GRAFANA_URL.rstrip('/')}/api/ds/query"
    
    try:
        response = requests.post(query_url, headers=headers,
                                 data=json.dumps(query_payload),
                                 timeout=120, verify=config.VERIFY_SSL)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException:
        return None

def get_all_placeholders(prs):
    placeholders = set()
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    matches = re.findall(r'(\{\{.*?\}\})', p.text)
                    for match in matches:
                        placeholders.add(match)
    return list(placeholders)

def replace_text_in_presentation(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p in shape.text_frame.paragraphs:
                if "{{" not in p.text:
                    continue
                original_run = p.runs[0] if p.runs else None
                original_font = original_run.font if original_run else None
                temp_text = p.text
                text_changed = False
                for placeholder, value in replacements.items():
                    if placeholder in temp_text:
                        temp_text = temp_text.replace(placeholder, str(value))
                        text_changed = True
                if text_changed:
                    p.text = temp_text
                    if original_font:
                        for run in p.runs:
                            font = run.font
                            font.name = original_font.name
                            if original_font.size:
                                font.size = original_font.size
                            if original_font.bold is not None:
                                font.bold = original_font.bold
                            if original_font.italic is not None:
                                font.italic = original_font.italic
                            try:
                                if original_font.color.rgb:
                                    font.color.rgb = RGBColor.from_string(str(original_font.color.rgb))
                            except:
                                pass

def process_statistics(customer_name=None):
    results = {
        "success": True,
        "processed_files": [],
        "errors": [],
        "failed_placeholders": [],
        "logs": [],
        "summary": {}
    }
    
    def log(message):
        results["logs"].append(message)
    
    try:
        date_info = calculate_previous_month_dates()
        date_placeholders = date_info["placeholders"]
        target_date_str = date_info["target_date_str"]
        start_ts = date_info["start_ts"]
        end_ts = date_info["end_ts"]
        log("날짜 정보 계산 완료")
        
        if customer_name:
            templates = find_templates_for_customer(config.OUTPUT_DIR_WITH_IMAGES, customer_name)
            results["summary"]["mode"] = f"특정 고객사: {customer_name}"
            log(f"고객사 '{customer_name}' 템플릿 검색 중...")
        else:
            templates = find_all_templates(config.OUTPUT_DIR_WITH_IMAGES)
            results["summary"]["mode"] = "전체 고객사"
            log("전체 고객사 템플릿 검색 중...")
        
        if not templates:
            results["success"] = False
            results["errors"].append("처리할 템플릿이 없습니다.")
            log("❌ 처리할 템플릿이 없습니다.")
            return results
        
        results["summary"]["total_templates"] = len(templates)
        log(f"총 {len(templates)}개 템플릿 발견")
        
        dashboard_defs_cache = {}
        panel_data_cache = {}
        
        for idx, template_path in enumerate(templates, 1):
            filename = os.path.basename(template_path)
            log(f"[{idx}/{len(templates)}] {filename} 처리 중...")
            try:
                final_replacements = date_placeholders.copy()
                relative_path = os.path.relpath(os.path.dirname(template_path), config.OUTPUT_DIR_WITH_IMAGES)
                first_level_folder = relative_path.split(os.sep)[0] if relative_path != '.' else None
                
                dashboard_def = dashboard_defs_cache.get(first_level_folder)
                if not dashboard_def and first_level_folder:
                    dashboard_uid = config.get_dashboard_uid(first_level_folder)
                    if not dashboard_uid:
                        log(f"  ⚠️ 고객사 '{first_level_folder}' 대시보드 매핑 없음")
                        results["errors"].append(f"고객사 '{first_level_folder}' 대시보드 매핑 없음")
                    else:
                        log(f"  → Grafana 대시보드 로딩: {dashboard_uid}")
                        dashboard_def = get_dashboard_definition(dashboard_uid)
                        dashboard_defs_cache[first_level_folder] = dashboard_def
                
                prs = Presentation(template_path)
                
                grafana_failures = []
                if dashboard_def:
                    all_panels_flat_list = dashboard_def.get('all_panels', [])
                    all_ph = get_all_placeholders(prs)
                    
                    for ph in all_ph:
                        if ph in final_replacements:
                            continue
                        try:
                            inner_text = ph.replace("{{", "").replace("}}", "")
                            parts = inner_text.rsplit('_', 1)
                            if len(parts) != 2:
                                continue
                            
                            panel_title_slug, query_letter = parts
                            cache_key = (panel_title_slug, query_letter)
                            
                            if cache_key not in panel_data_cache:
                                panel_title = panel_title_slug.replace("-", " ")
                                log(f"  → Grafana 조회: '{panel_title}' - {query_letter} 쿼리")
                                panel = find_panel_by_title(all_panels_flat_list, panel_title)
                                if panel:
                                    panel_data_cache[cache_key] = get_grafana_stats_by_panel(panel, query_letter, start_ts, end_ts)
                                    time.sleep(0.5)
                                else:
                                    log(f"  ⚠️ 패널 '{panel_title}' 없음")
                                    panel_data_cache[cache_key] = None
                            
                            stats_data = panel_data_cache.get(cache_key)
                            if not stats_data:
                                final_replacements[ph] = "N/A"
                                grafana_failures.append(ph)
                                continue
                            
                            metric_values = None
                            frames = stats_data.get('results', {}).get(query_letter, {}).get('frames', [])
                            
                            if frames:
                                fields = frames[0].get('schema', {}).get('fields', [])
                                for idx, field in enumerate(fields):
                                    f_type = field.get('type', '')
                                    if f_type in ['number', 'float', 'int']:
                                        metric_values = frames[0].get('data', {}).get('values', [])[idx]
                                        break
                            
                            if metric_values:
                                valid_numbers = [v for v in metric_values if v is not None]
                                if valid_numbers:
                                    max_val = max(valid_numbers)
                                    mean_val = sum(valid_numbers) / len(valid_numbers)
                                    final_replacements[ph] = config.SENTENCE_TEMPLATE.format(
                                        max=f"{max_val:.1f}", mean=f"{mean_val:.1f}"
                                    )
                                else:
                                    final_replacements[ph] = "N/A"
                                    grafana_failures.append(ph)
                            else:
                                final_replacements[ph] = "N/A"
                                grafana_failures.append(ph)
                                
                        except (KeyError, IndexError, TypeError, ValueError):
                            final_replacements[ph] = "N/A"
                            grafana_failures.append(ph)
                
                replace_text_in_presentation(prs, final_replacements)
                
                if grafana_failures:
                    for name in grafana_failures:
                        results["failed_placeholders"].append({
                            "placeholder": name,
                            "reason": "Grafana 데이터 조회 실패"
                        })
                
                original_filename = os.path.basename(template_path)
                new_filename = re.sub(r'\d{4}[.년]\s*\d{1,2}월', target_date_str, original_filename)
                new_filename = new_filename.replace(" - 복사본", "").replace("-복사본", "")
                
                output_subdir = os.path.join(config.OUTPUT_DIR, relative_path)
                os.makedirs(output_subdir, exist_ok=True)
                final_output_path = os.path.join(output_subdir, new_filename)
                prs.save(final_output_path)
                grafana_query_count = len([p for p in final_replacements if p not in date_placeholders])
                log(f"  → {grafana_query_count}개 통계 삽입 완료")
                log(f"  ✅ 저장 완료: {new_filename}")
                
                results["processed_files"].append({
                    "template": original_filename,
                    "customer": first_level_folder,
                    "output_path": final_output_path,
                    "new_filename": new_filename,
                    "grafana_queries": grafana_query_count
                })
                
                del prs
                gc.collect()
                time.sleep(1)
                
            except Exception as e:
                log(f"  ❌ 오류: {str(e)}")
                results["errors"].append(f"템플릿 처리 실패 ({os.path.basename(template_path)}): {str(e)}")
        
        results["summary"]["processed_count"] = len(results["processed_files"])
        results["summary"]["error_count"] = len(results["errors"])
        results["summary"]["failed_placeholder_count"] = len(results["failed_placeholders"])
        
        log(f"===== 작업 완료 =====")
        log(f"처리된 파일: {results['summary']['processed_count']}개")
        log(f"오류: {results['summary']['error_count']}개")
        log(f"실패한 플레이스홀더: {results['summary']['failed_placeholder_count']}개")
        
        if results["summary"]["processed_count"] == 0:
            results["success"] = False
        
    except Exception as e:
        results["success"] = False
        results["errors"].append(f"전체 프로세스 실패: {str(e)}")
        log(f"❌ 전체 프로세스 실패: {str(e)}")
    
    return results
