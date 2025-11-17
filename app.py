import streamlit as st
import os
import sys
import subprocess
import config
from datetime import datetime

st.set_page_config(
    page_title="PowerPoint 자동화 도구",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PowerPoint 자동화 도구")
st.markdown("이미지와 Grafana 통계를 자동으로 PowerPoint에 삽입하는 도구입니다.")

tab1, tab2, tab3, tab4 = st.tabs(["🏠 홈", "🖼️ 이미지 삽입", "📈 통계 삽입", "⚙️ 설정"])

with tab1:
    st.header("환영합니다!")
    st.markdown("""
    이 도구는 두 단계로 PowerPoint 보고서를 자동 생성합니다:
    
    **Step 1: 이미지 삽입**
    - PowerPoint 템플릿의 도형 이름과 일치하는 이미지를 자동으로 삽입합니다
    - 날짜 플레이스홀더를 자동으로 채웁니다
    
    **Step 2: Grafana 통계 삽입**
    - Grafana 대시보드에서 통계를 가져옵니다
    - 플레이스홀더를 실제 데이터로 교체합니다
    
    왼쪽 탭에서 각 기능을 사용하실 수 있습니다.
    """)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("설정된 고객사", len(config.DASHBOARD_MAP))
    
    with col2:
        template_count = 0
        if os.path.exists(config.BASE_TEMPLATE_DIR):
            for root, dirs, files in os.walk(config.BASE_TEMPLATE_DIR):
                template_count += len([f for f in files if f.endswith('.pptx')])
        st.metric("템플릿 파일", template_count)
    
    with col3:
        api_status = "✅ 설정됨" if config.API_KEY else "❌ 미설정"
        st.metric("Grafana API", api_status)

with tab2:
    st.header("🖼️ 이미지 삽입 (Step 1)")
    st.markdown("PowerPoint 템플릿에 이미지를 삽입합니다.")
    
    st.info(f"**템플릿 폴더:** `{config.BASE_TEMPLATE_DIR}`")
    
    if not os.path.exists(config.BASE_TEMPLATE_DIR):
        st.warning("템플릿 폴더가 없습니다. 아래 버튼을 눌러 폴더를 생성하세요.")
        if st.button("📁 폴더 생성", key="create_dirs_tab2"):
            dirs = [config.BASE_TEMPLATE_DIR, config.OUTPUT_DIR_WITH_IMAGES, config.OUTPUT_DIR]
            for d in dirs:
                os.makedirs(d, exist_ok=True)
            st.success("폴더가 생성되었습니다!")
            st.rerun()
    else:
        st.subheader("📤 템플릿 업로드")
        uploaded_files = st.file_uploader(
            "PowerPoint 템플릿 파일을 선택하세요 (.pptx)",
            type=['pptx'],
            accept_multiple_files=True,
            key="template_uploader"
        )
        
        if uploaded_files:
            existing_files = []
            new_files = []
            for uploaded_file in uploaded_files:
                file_path = os.path.join(config.BASE_TEMPLATE_DIR, uploaded_file.name)
                if os.path.exists(file_path):
                    existing_files.append(uploaded_file.name)
                else:
                    new_files.append(uploaded_file.name)
            
            if existing_files:
                st.warning(f"⚠️ 다음 파일은 이미 존재합니다. 저장하면 덮어씁니다:\n" + "\n".join([f"- {f}" for f in existing_files]))
            
            if new_files:
                st.info(f"📝 새로 저장될 파일:\n" + "\n".join([f"- {f}" for f in new_files]))
            
            if st.button("💾 업로드된 파일 저장", type="primary", key="save_templates"):
                success_count = 0
                overwritten_count = len(existing_files)
                
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(config.BASE_TEMPLATE_DIR, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    success_count += 1
                
                if overwritten_count > 0:
                    st.success(f"✅ {success_count}개의 파일이 저장되었습니다! ({overwritten_count}개 덮어쓰기)")
                else:
                    st.success(f"✅ {success_count}개의 템플릿 파일이 저장되었습니다!")
                
                import time
                time.sleep(1)
                st.rerun()
        
        st.divider()
        
        st.subheader("🖼️ 이미지 업로드")
        st.markdown("템플릿에 삽입할 이미지를 고객사별로 업로드합니다.")
        
        upload_method = st.radio(
            "업로드 방식 선택",
            ["개별 파일 업로드", "폴더 업로드 (ZIP)"],
            key="upload_method",
            horizontal=True
        )
        
        st.divider()
        
        if upload_method == "폴더 업로드 (ZIP)":
            st.info("💡 **폴더 업로드 방법**: 고객사별 이미지 폴더를 ZIP으로 압축하여 업로드하세요. ZIP 파일 내부의 폴더 구조가 그대로 유지됩니다.")
            st.markdown("""
            **예시 ZIP 구조:**
            ```
            images.zip
            ├── 고객사A/
            │   ├── image1.png
            │   └── image2.jpg
            └── 고객사B/
                ├── logo.png
                └── chart.png
            ```
            """)
            
            uploaded_zip = st.file_uploader(
                "ZIP 파일을 선택하세요",
                type=['zip'],
                key="zip_uploader"
            )
            
            if uploaded_zip:
                import zipfile
                import io
                
                if st.button("📦 ZIP 압축 해제 및 저장", type="primary", key="extract_zip"):
                    try:
                        zip_buffer = io.BytesIO(uploaded_zip.getbuffer())
                        base_dir_abs = os.path.abspath(config.BASE_IMAGE_DIR)
                        
                        with zipfile.ZipFile(zip_buffer, 'r') as zip_ref:
                            file_list = zip_ref.namelist()
                            image_extensions = ('.png', '.jpg', '.jpeg', '.gif')
                            
                            extracted_count = 0
                            skipped_count = 0
                            created_folders = set()
                            
                            with st.spinner("압축을 풀고 있습니다..."):
                                for file_path in file_list:
                                    if file_path.endswith('/') or not file_path.lower().endswith(image_extensions):
                                        continue
                                    
                                    if os.path.isabs(file_path):
                                        skipped_count += 1
                                        continue
                                    
                                    parts = file_path.split('/')
                                    safe_parts = []
                                    for part in parts:
                                        safe_part = os.path.basename(part.strip())
                                        if safe_part and safe_part not in ['.', '..'] and not safe_part.startswith('.'):
                                            safe_parts.append(safe_part)
                                        else:
                                            safe_parts = []
                                            break
                                    
                                    if not safe_parts:
                                        skipped_count += 1
                                        continue
                                    
                                    target_path = os.path.join(config.BASE_IMAGE_DIR, *safe_parts)
                                    target_path_abs = os.path.abspath(os.path.realpath(target_path))
                                    base_dir_real = os.path.abspath(os.path.realpath(config.BASE_IMAGE_DIR))
                                    
                                    try:
                                        common = os.path.commonpath([base_dir_real, target_path_abs])
                                        if common != base_dir_real:
                                            skipped_count += 1
                                            continue
                                    except (ValueError, TypeError):
                                        skipped_count += 1
                                        continue
                                    
                                    target_dir = os.path.dirname(target_path)
                                    os.makedirs(target_dir, exist_ok=True)
                                    created_folders.add(os.path.relpath(target_dir, config.BASE_IMAGE_DIR))
                                    
                                    with zip_ref.open(file_path) as source:
                                        with open(target_path, 'wb') as target:
                                            target.write(source.read())
                                    
                                    extracted_count += 1
                            
                            if extracted_count > 0:
                                st.success(f"✅ {extracted_count}개의 이미지가 저장되었습니다!")
                                if created_folders:
                                    st.info(f"📁 생성/업데이트된 폴더:\n" + "\n".join([f"- {f}" for f in sorted(created_folders)]))
                            else:
                                st.warning("⚠️ 저장할 수 있는 이미지가 없습니다.")
                            
                            if skipped_count > 0:
                                st.warning(f"⚠️ {skipped_count}개의 파일은 건너뛰었습니다. (이미지 파일이 아니거나 올바르지 않은 경로)")
                            
                            import time
                            time.sleep(1)
                            st.rerun()
                    
                    except zipfile.BadZipFile:
                        st.error("❌ 올바른 ZIP 파일이 아닙니다.")
                    except Exception as e:
                        st.error(f"❌ 오류가 발생했습니다: {str(e)}")
        
        else:
            col1, col2 = st.columns([1, 2])
            
            with col1:
                customer_folders = []
                if os.path.exists(config.BASE_IMAGE_DIR):
                    for item in os.listdir(config.BASE_IMAGE_DIR):
                        item_path = os.path.join(config.BASE_IMAGE_DIR, item)
                        if os.path.isdir(item_path) and item not in ['template', 'completed_with_images', 'completed_final']:
                            customer_folders.append(item)
                
                customer_input_mode = st.radio(
                    "고객사 폴더",
                    ["기존 폴더 선택", "새 폴더 생성"],
                    key="customer_mode"
                )
                
                if customer_input_mode == "기존 폴더 선택":
                    if customer_folders:
                        selected_customer = st.selectbox("고객사 선택", customer_folders, key="customer_folder_select")
                    else:
                        st.warning("고객사 폴더가 없습니다. 새 폴더를 생성하세요.")
                        selected_customer = None
                else:
                    customer_input = st.text_input("고객사 이름 입력 (영문/숫자/한글만 사용)", key="new_customer_name")
                    if customer_input:
                        sanitized = os.path.basename(customer_input.strip())
                        if sanitized and sanitized not in ['.', '..'] and '/' not in customer_input and '\\' not in customer_input:
                            selected_customer = sanitized
                        else:
                            st.error("⚠️ 올바르지 않은 폴더 이름입니다. 특수문자(/, \\)는 사용할 수 없습니다.")
                            selected_customer = None
                    else:
                        selected_customer = None
            
            with col2:
                if selected_customer:
                    uploaded_images = st.file_uploader(
                        f"이미지 파일을 선택하세요 ({selected_customer})",
                        type=['png', 'jpg', 'jpeg', 'gif'],
                        accept_multiple_files=True,
                        key="image_uploader"
                    )
                    
                    if uploaded_images:
                        customer_dir = os.path.join(config.BASE_IMAGE_DIR, selected_customer)
                        customer_dir_abs = os.path.abspath(customer_dir)
                        base_dir_abs = os.path.abspath(config.BASE_IMAGE_DIR)
                        
                        existing_images = []
                        new_images = []
                        invalid_images = []
                        
                        for uploaded_image in uploaded_images:
                            safe_filename = os.path.basename(uploaded_image.name)
                            
                            if not safe_filename or safe_filename in ['.', '..'] or '/' in uploaded_image.name or '\\' in uploaded_image.name:
                                invalid_images.append(uploaded_image.name)
                                continue
                            
                            image_path = os.path.join(customer_dir, safe_filename)
                            image_path_abs = os.path.abspath(image_path)
                            
                            if not image_path_abs.startswith(customer_dir_abs):
                                invalid_images.append(uploaded_image.name)
                                continue
                            
                            if os.path.exists(image_path):
                                existing_images.append(safe_filename)
                            else:
                                new_images.append(safe_filename)
                        
                        if invalid_images:
                            st.error(f"⚠️ 다음 파일은 올바르지 않은 이름입니다:\n" + "\n".join([f"- {f}" for f in invalid_images]))
                        
                        if existing_images:
                            st.warning(f"⚠️ 다음 이미지는 이미 존재합니다:\n" + "\n".join([f"- {f}" for f in existing_images]))
                        
                        if new_images:
                            st.info(f"📝 새로 저장될 이미지:\n" + "\n".join([f"- {f}" for f in new_images]))
                        
                        if st.button("💾 이미지 저장", type="primary", key="save_images"):
                            os.makedirs(customer_dir, exist_ok=True)
                            success_count = 0
                            overwritten_count = len(existing_images)
                            
                            for uploaded_image in uploaded_images:
                                safe_filename = os.path.basename(uploaded_image.name)
                                
                                if not safe_filename or safe_filename in ['.', '..'] or '/' in uploaded_image.name or '\\' in uploaded_image.name:
                                    continue
                                
                                image_path = os.path.join(customer_dir, safe_filename)
                                image_path_abs = os.path.abspath(image_path)
                                
                                if not image_path_abs.startswith(customer_dir_abs):
                                    continue
                                
                                with open(image_path, "wb") as f:
                                    f.write(uploaded_image.getbuffer())
                                success_count += 1
                            
                            if success_count == 0:
                                st.error("❌ 저장할 수 있는 이미지가 없습니다.")
                            elif overwritten_count > 0:
                                st.success(f"✅ {success_count}개의 이미지가 저장되었습니다! ({overwritten_count}개 덮어쓰기)")
                            else:
                                st.success(f"✅ {success_count}개의 이미지가 저장되었습니다!")
                            
                            import time
                            time.sleep(1)
                            st.rerun()
            
            if customer_folders:
                with st.expander("📁 고객사별 이미지 현황"):
                    for folder in customer_folders:
                        folder_path = os.path.join(config.BASE_IMAGE_DIR, folder)
                        image_files = [f for f in os.listdir(folder_path) 
                                       if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]
                        st.markdown(f"**{folder}**: {len(image_files)}개 이미지")
                        if image_files:
                            for img in image_files:
                                st.text(f"  📷 {img}")
        
        st.divider()
        
        templates = []
        for root, dirs, files in os.walk(config.BASE_TEMPLATE_DIR):
            for f in files:
                if f.endswith('.pptx') and not f.startswith('~$'):
                    templates.append(f)
        
        if templates:
            st.success(f"✅ {len(templates)}개의 템플릿 파일을 찾았습니다")
            with st.expander("템플릿 파일 목록 보기"):
                for t in templates:
                    st.text(f"📄 {t}")
            
            if st.button("▶️ 이미지 삽입 실행", type="primary", key="run_image"):
                with st.spinner("이미지를 삽입하는 중..."):
                    result = subprocess.run(
                        ["python", "insert_images.py"],
                        capture_output=True,
                        text=True
                    )
                    
                    st.subheader("실행 결과")
                    if result.returncode == 0:
                        st.success("✅ 이미지 삽입이 완료되었습니다!")
                    else:
                        st.error("❌ 오류가 발생했습니다.")
                    
                    with st.expander("상세 로그 보기"):
                        st.code(result.stdout + result.stderr)
        else:
            st.warning("⚠️ 템플릿 파일이 없습니다. PowerPoint 파일(.pptx)을 템플릿 폴더에 업로드하세요.")

with tab3:
    st.header("📈 Grafana 통계 삽입 (Step 2)")
    st.markdown("이미지가 삽입된 파일에 Grafana 통계를 추가합니다.")
    
    st.info(f"**입력 폴더:** `{config.OUTPUT_DIR_WITH_IMAGES}`")
    
    col1, col2 = st.columns(2)
    
    with col1:
        process_all = st.radio(
            "처리 범위",
            ["전체 고객사", "특정 고객사"],
            key="process_range"
        )
    
    with col2:
        customer_name = ""
        if process_all == "특정 고객사":
            customer_list = [""] + list(config.DASHBOARD_MAP.keys())
            customer_name = st.selectbox("고객사 선택", customer_list, key="customer_select")
    
    if not config.API_KEY:
        st.error("⚠️ Grafana API 키가 설정되지 않았습니다. 설정 탭에서 환경 변수를 확인하세요.")
    
    if st.button("▶️ 통계 삽입 실행", type="primary", key="run_stats"):
        cmd = ["python", "numinsert3.py"]
        if process_all == "특정 고객사" and customer_name:
            cmd.append(customer_name)
        
        with st.spinner("Grafana 통계를 삽입하는 중..."):
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True
            )
            
            st.subheader("실행 결과")
            if result.returncode == 0:
                st.success("✅ 통계 삽입이 완료되었습니다!")
            else:
                st.error("❌ 오류가 발생했습니다.")
            
            with st.expander("상세 로그 보기"):
                st.code(result.stdout + result.stderr)

def update_config_file(base_dir):
    """config.py 파일 업데이트"""
    config_content = f'''import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BASE_TEMPLATE_DIR = os.path.join(BASE_DIR, "{base_dir}", "template")
BASE_IMAGE_DIR = os.path.join(BASE_DIR, "{base_dir}")
OUTPUT_DIR_WITH_IMAGES = os.path.join(BASE_DIR, "{base_dir}", "completed_with_images")
OUTPUT_DIR = os.path.join(BASE_DIR, "{base_dir}", "completed_final")

GRAFANA_URL = os.getenv("GRAFANA_URL", "http://localhost:3000")
API_KEY = os.getenv("GRAFANA_API_KEY", "")
VERIFY_SSL = os.getenv("GRAFANA_VERIFY_SSL", "true").lower() in ("true", "1", "yes")

DASHBOARD_MAP = {{
    "kpmo": "dejkgjz0jnoqoa",
    "GIT": "aejkgkoze5nggb",
    "hansystem": "cejnb5yyuk5q8e",
    "humecca": "bejnb5db19blse",
    "klcns": "eejnb31cylreod",
    "sungwoo": "cejnb4aafury8e",
    "thepnl": "fejkgid897xtsc",
    "프리스타일": "fejkgfwux1fy8c"
}}

SENTENCE_TEMPLATE = "사용량 최대 {{max}}%, 평균 {{mean}}% 입니다."
'''
    with open("config.py", "w", encoding="utf-8") as f:
        f.write(config_content)

with tab4:
    st.header("⚙️ 설정")
    
    st.subheader("📁 디렉토리 경로 설정")
    
    with st.expander("디렉토리 경로 편집", expanded=False):
        st.info("기본 디렉토리 이름을 변경할 수 있습니다. 변경 후 페이지를 새로고침해야 적용됩니다.")
        
        current_base = "Report"
        base_dir_name = st.text_input(
            "기본 디렉토리 이름",
            value=current_base,
            help="모든 파일이 저장될 기본 폴더 이름입니다.",
            key="base_dir_input"
        )
        
        st.markdown("**변경 후 디렉토리 구조:**")
        st.code(f"""
{base_dir_name}/
├── template/                  # 템플릿 파일
├── completed_with_images/     # 이미지 삽입 결과
└── completed_final/           # 최종 결과
        """)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 경로 저장", type="primary", key="save_paths"):
                try:
                    update_config_file(base_dir_name)
                    
                    new_dirs = [
                        os.path.join(base_dir_name, "template"),
                        os.path.join(base_dir_name, "completed_with_images"),
                        os.path.join(base_dir_name, "completed_final")
                    ]
                    for d in new_dirs:
                        os.makedirs(d, exist_ok=True)
                    
                    st.success("✅ 설정이 저장되었습니다! 페이지를 새로고침하세요.")
                    st.info("⚠️ 변경사항을 적용하려면 브라우저를 새로고침하세요 (Ctrl+R 또는 F5)")
                except Exception as e:
                    st.error(f"오류 발생: {e}")
        
        with col2:
            if st.button("🔄 기본값으로 복원", key="reset_paths"):
                try:
                    update_config_file("Report")
                    st.success("✅ 기본 설정으로 복원되었습니다! 페이지를 새로고침하세요.")
                except Exception as e:
                    st.error(f"오류 발생: {e}")
    
    st.divider()
    
    st.subheader("현재 디렉토리 상태")
    dirs_info = {
        "템플릿 폴더": config.BASE_TEMPLATE_DIR,
        "이미지 폴더": config.BASE_IMAGE_DIR,
        "중간 출력 폴더": config.OUTPUT_DIR_WITH_IMAGES,
        "최종 출력 폴더": config.OUTPUT_DIR
    }
    
    for name, path in dirs_info.items():
        col1, col2 = st.columns([3, 1])
        with col1:
            st.text(f"{name}: {path}")
        with col2:
            if os.path.exists(path):
                st.success("✅ 존재")
            else:
                st.error("❌ 없음")
    
    if st.button("📁 모든 폴더 생성", key="create_all_dirs"):
        for path in dirs_info.values():
            os.makedirs(path, exist_ok=True)
        st.success("모든 폴더가 생성되었습니다!")
    
    st.divider()
    
    st.subheader("🔐 Grafana API 설정")
    
    with st.expander("Grafana API 설정 편집", expanded=False):
        st.info("Grafana API 연결 정보를 설정합니다. 설정 후 페이지를 새로고침해야 적용됩니다.")
        
        try:
            secrets_file = ".streamlit/secrets.toml"
            current_secrets = {}
            if os.path.exists(secrets_file):
                import toml
                with open(secrets_file, "r") as f:
                    current_secrets = toml.load(f)
        except:
            current_secrets = {}
        
        grafana_url = st.text_input(
            "Grafana URL",
            value=current_secrets.get("GRAFANA_URL", "http://localhost:3000"),
            help="Grafana 서버 주소 (예: http://localhost:3000 또는 https://your-grafana.com)",
            key="grafana_url_input"
        )
        
        grafana_api_key = st.text_input(
            "Grafana API Key",
            value=current_secrets.get("GRAFANA_API_KEY", ""),
            type="password",
            help="Grafana에서 생성한 API 키",
            key="grafana_api_input"
        )
        
        verify_ssl = st.selectbox(
            "SSL 인증서 검증",
            options=["true", "false"],
            index=0 if current_secrets.get("GRAFANA_VERIFY_SSL", "true") == "true" else 1,
            help="HTTPS 사용 시 SSL 인증서 검증 여부 (프로덕션에서는 true 권장)",
            key="verify_ssl_select"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 API 설정 저장", type="primary", key="save_grafana"):
                try:
                    os.makedirs(".streamlit", exist_ok=True)
                    
                    secrets_content = f"""# Grafana API 설정
GRAFANA_URL = "{grafana_url}"
GRAFANA_API_KEY = "{grafana_api_key}"
GRAFANA_VERIFY_SSL = "{verify_ssl}"
"""
                    
                    with open(secrets_file, "w", encoding="utf-8") as f:
                        f.write(secrets_content)
                    
                    st.success("✅ Grafana API 설정이 저장되었습니다!")
                    st.info("⚠️ 변경사항을 적용하려면 페이지를 새로고침하세요 (Ctrl+R 또는 F5)")
                except Exception as e:
                    st.error(f"오류 발생: {e}")
        
        with col2:
            if st.button("🗑️ API 설정 삭제", key="clear_grafana"):
                try:
                    if os.path.exists(secrets_file):
                        os.remove(secrets_file)
                        st.success("✅ Grafana API 설정이 삭제되었습니다!")
                        st.info("⚠️ 변경사항을 적용하려면 페이지를 새로고침하세요")
                except Exception as e:
                    st.error(f"오류 발생: {e}")
    
    st.divider()
    
    st.subheader("현재 Grafana 설정")
    st.text(f"Grafana URL: {config.GRAFANA_URL}")
    st.text(f"API 키: {'*' * 20 if config.API_KEY else '❌ 미설정'}")
    st.text(f"SSL 검증: {'✅ 활성화' if config.VERIFY_SSL else '⚠️ 비활성화'}")
    
    if not config.API_KEY:
        st.warning("⚠️ Grafana API 키가 설정되지 않았습니다. 위의 'Grafana API 설정 편집'에서 설정하세요.")
    
    st.divider()
    
    st.subheader("고객사 대시보드 매핑")
    st.text(f"총 {len(config.DASHBOARD_MAP)}개의 고객사가 설정되어 있습니다.")
    
    with st.expander("고객사 목록 보기"):
        for customer, uid in config.DASHBOARD_MAP.items():
            st.text(f"• {customer}: {uid}")
    
    st.info("고객사 매핑을 변경하려면 `config.py` 파일을 수정하세요.")

st.divider()
st.caption("PowerPoint 자동화 도구 v1.0 | Replit에서 실행 중")
