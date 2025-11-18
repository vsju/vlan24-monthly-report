import streamlit as st
import os
import sys
import subprocess
import config
from datetime import datetime
import db_utils

st.set_page_config(
    page_title="PowerPoint 자동화 도구",
    page_icon="📊",
    layout="wide"
)

# Initialize session state
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False
if 'user_id' not in st.session_state:
    st.session_state['user_id'] = None
if 'username' not in st.session_state:
    st.session_state['username'] = None
if 'role' not in st.session_state:
    st.session_state['role'] = None
if 'full_name' not in st.session_state:
    st.session_state['full_name'] = None

# Login/Logout functions
def login(username, password):
    """Authenticate user and set session state"""
    user = db_utils.get_user_by_username(username)
    
    if user and user.is_active and db_utils.verify_password(password, user.password_hash):
        st.session_state['authenticated'] = True
        st.session_state['user_id'] = user.id
        st.session_state['username'] = user.username
        st.session_state['role'] = user.role
        st.session_state['full_name'] = user.full_name
        db_utils.update_last_login(username)
        return True
    return False

def logout():
    """Clear session state and log out"""
    st.session_state['authenticated'] = False
    st.session_state['user_id'] = None
    st.session_state['username'] = None
    st.session_state['role'] = None
    st.session_state['full_name'] = None

# Main application
if not st.session_state['authenticated']:
    # Show login form
    st.title("🔐 로그인")
    st.markdown("PowerPoint 자동화 도구에 접속하려면 로그인하세요.")
    
    with st.form("login_form"):
        username = st.text_input("사용자명", key="login_username")
        password = st.text_input("비밀번호", type="password", key="login_password")
        submit = st.form_submit_button("로그인", type="primary")
        
        if submit:
            if not username or not password:
                st.error("사용자명과 비밀번호를 모두 입력하세요.")
            elif login(username, password):
                st.success("로그인 성공!")
                st.rerun()
            else:
                st.error("로그인 실패. 사용자명 또는 비밀번호를 확인하세요.")

else:
    # User is logged in - show main application
    
    # Header with user info and logout button
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("📊 PowerPoint 자동화 도구")
        st.markdown(f"안녕하세요, **{st.session_state['full_name']}**님! ({st.session_state['role']})")
    with col2:
        if st.button("🚪 로그아웃", type="secondary"):
            logout()
            st.rerun()
    
    st.markdown("이미지와 Grafana 통계를 자동으로 PowerPoint에 삽입하는 도구입니다.")
    
    # Create tabs - add "사용자 관리" tab for admins
    if st.session_state['role'] == 'admin':
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "🏠 홈", 
            "🖼️ 이미지 삽입", 
            "📈 통계 삽입", 
            "📋 작업 이력",
            "👥 사용자 관리",
            "⚙️ 설정"
        ])
    else:
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🏠 홈", 
            "🖼️ 이미지 삽입", 
            "📈 통계 삽입", 
            "📋 작업 이력",
            "⚙️ 설정"
        ])
        tab6 = None  # No user management tab for regular users
    
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
                            
                            completed_files = []
                            if os.path.exists(config.OUTPUT_DIR_WITH_IMAGES):
                                for f in os.listdir(config.OUTPUT_DIR_WITH_IMAGES):
                                    if f.endswith('.pptx') and not f.startswith('~$'):
                                        completed_files.append(f)
                            
                            if completed_files:
                                st.subheader("📥 완료된 파일 다운로드")
                                st.info(f"총 {len(completed_files)}개의 파일이 생성되었습니다.")
                                
                                for file_name in completed_files:
                                    file_path = os.path.join(config.OUTPUT_DIR_WITH_IMAGES, file_name)
                                    file_size = os.path.getsize(file_path) / 1024
                                    
                                    col_file, col_download = st.columns([3, 1])
                                    with col_file:
                                        st.text(f"📄 {file_name} ({file_size:.1f} KB)")
                                    
                                    with col_download:
                                        with open(file_path, "rb") as f:
                                            st.download_button(
                                                label="다운로드",
                                                data=f,
                                                file_name=file_name,
                                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                                key=f"download_img_{file_name}"
                                            )
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
        
        if not os.path.exists(config.OUTPUT_DIR_WITH_IMAGES):
            st.warning("입력 폴더가 없습니다. 아래 버튼을 눌러 폴더를 생성하세요.")
            if st.button("📁 폴더 생성", key="create_dirs_tab3"):
                os.makedirs(config.OUTPUT_DIR_WITH_IMAGES, exist_ok=True)
                os.makedirs(config.OUTPUT_DIR, exist_ok=True)
                st.success("폴더가 생성되었습니다!")
                st.rerun()
        else:
            st.subheader("📤 템플릿 업로드")
            st.markdown("통계를 삽입할 PowerPoint 파일을 업로드하세요. (이미 이미지가 삽입된 파일이거나 원본 템플릿)")
            
            customer_for_upload = st.selectbox(
                "고객사 선택",
                options=list(config.DASHBOARD_MAP.keys()),
                help="파일이 저장될 고객사 폴더를 선택하세요",
                key="customer_for_upload"
            )
            
            uploaded_stats_files = st.file_uploader(
                "PowerPoint 파일을 선택하세요 (.pptx)",
                type=['pptx'],
                accept_multiple_files=True,
                key="stats_template_uploader"
            )
            
            if uploaded_stats_files and customer_for_upload:
                customer_folder = os.path.join(config.OUTPUT_DIR_WITH_IMAGES, customer_for_upload)
                
                existing_files = []
                new_files = []
                for uploaded_file in uploaded_stats_files:
                    file_path = os.path.join(customer_folder, uploaded_file.name)
                    if os.path.exists(file_path):
                        existing_files.append(uploaded_file.name)
                    else:
                        new_files.append(uploaded_file.name)
                
                st.info(f"📁 저장 경로: `{customer_folder}/`")
                
                if existing_files:
                    st.warning(f"⚠️ 다음 파일은 이미 존재합니다. 저장하면 덮어씁니다:\n" + "\n".join([f"- {f}" for f in existing_files]))
                
                if new_files:
                    st.info(f"📝 새로 저장될 파일:\n" + "\n".join([f"- {f}" for f in new_files]))
                
                col_btn1, col_btn2 = st.columns(2)
                
                with col_btn1:
                    if st.button("💾 저장만 하기", type="secondary", key="save_only_stats_templates"):
                        success_count = 0
                        overwritten_count = len(existing_files)
                        
                        os.makedirs(customer_folder, exist_ok=True)
                        
                        for uploaded_file in uploaded_stats_files:
                            file_path = os.path.join(customer_folder, uploaded_file.name)
                            with open(file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            success_count += 1
                        
                        if overwritten_count > 0:
                            st.success(f"✅ {success_count}개의 파일이 `{customer_for_upload}/` 폴더에 저장되었습니다! ({overwritten_count}개 덮어쓰기)")
                        else:
                            st.success(f"✅ {success_count}개의 파일이 `{customer_for_upload}/` 폴더에 저장되었습니다!")
                        
                        import time
                        time.sleep(1)
                        st.rerun()
                
                with col_btn2:
                    if st.button("💾 저장 후 바로 통계 삽입", type="primary", key="save_and_run_stats"):
                        import time as time_module
                        
                        # First, save files
                        success_count = 0
                        overwritten_count = len(existing_files)
                        
                        os.makedirs(customer_folder, exist_ok=True)
                        
                        for uploaded_file in uploaded_stats_files:
                            file_path = os.path.join(customer_folder, uploaded_file.name)
                            with open(file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            success_count += 1
                        
                        if overwritten_count > 0:
                            st.success(f"✅ {success_count}개의 파일이 `{customer_for_upload}/` 폴더에 저장되었습니다! ({overwritten_count}개 덮어쓰기)")
                        else:
                            st.success(f"✅ {success_count}개의 파일이 `{customer_for_upload}/` 폴더에 저장되었습니다!")
                        
                        time_module.sleep(0.5)
                        
                        # Then run statistics insertion for this customer
                        st.info(f"🚀 {customer_for_upload} 고객사에 대한 통계 삽입을 시작합니다...")
                        
                        if not config.API_KEY:
                            st.error("⚠️ Grafana API 키가 설정되지 않았습니다. 설정 탭에서 환경 변수를 확인하세요.")
                        else:
                            cmd = ["python", "numinsert3.py", customer_for_upload]
                            
                            with st.spinner("Grafana 통계를 삽입하는 중..."):
                                result = subprocess.run(
                                    cmd,
                                    capture_output=True,
                                    text=True
                                )
                                
                                if result.returncode == 0:
                                    st.success("✅ 통계 삽입이 완료되었습니다!")
                                    
                                    # Show generated files
                                    customer_output_dir = os.path.join(config.OUTPUT_DIR, customer_for_upload)
                                    if os.path.exists(customer_output_dir):
                                        generated_files = [f for f in os.listdir(customer_output_dir) 
                                                          if f.endswith('.pptx') and not f.startswith('~$')]
                                        
                                        if generated_files:
                                            st.subheader("📥 생성된 파일")
                                            for file_name in generated_files:
                                                file_path = os.path.join(customer_output_dir, file_name)
                                                file_size = os.path.getsize(file_path) / 1024
                                                
                                                col_file, col_download = st.columns([3, 1])
                                                with col_file:
                                                    st.text(f"📄 {file_name} ({file_size:.1f} KB)")
                                                
                                                with col_download:
                                                    with open(file_path, "rb") as f:
                                                        st.download_button(
                                                            label="다운로드",
                                                            data=f,
                                                            file_name=file_name,
                                                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                                            key=f"download_quick_{file_name}"
                                                        )
                                else:
                                    st.error("❌ 통계 삽입 중 오류가 발생했습니다.")
                                
                                with st.expander("상세 로그 보기"):
                                    st.code(result.stdout + result.stderr)
            
            st.divider()
            
            st.subheader("📂 고객사별 파일 현황")
            
            customer_files = {}
            if os.path.exists(config.OUTPUT_DIR_WITH_IMAGES):
                for customer in config.DASHBOARD_MAP.keys():
                    customer_folder = os.path.join(config.OUTPUT_DIR_WITH_IMAGES, customer)
                    if os.path.exists(customer_folder):
                        files = [f for f in os.listdir(customer_folder) 
                                if f.endswith('.pptx') and not f.startswith('~$')]
                        if files:
                            customer_files[customer] = files
            
            total_files = sum(len(files) for files in customer_files.values())
            
            if customer_files:
                st.success(f"✅ 총 {total_files}개의 입력 파일을 찾았습니다 ({len(customer_files)}개 고객사)")
                
                with st.expander("고객사별 파일 목록 보기"):
                    for customer, files in sorted(customer_files.items()):
                        st.markdown(f"**{customer}** ({len(files)}개)")
                        for f in files:
                            st.text(f"  📄 {f}")
            else:
                st.warning("⚠️ 입력 파일이 없습니다. 위에서 파일을 업로드하거나 Step 1을 먼저 실행하세요.")
        
        st.divider()
        st.subheader("🚀 통계 삽입 실행")
        
        st.info("💡 팁: 위에서 템플릿 업로드 후 '저장 후 바로 통계 삽입' 버튼을 누르면 한 번에 처리됩니다.")
        
        process_all = st.radio(
            "처리 범위",
            ["전체 고객사 처리", "특정 고객사만 처리"],
            index=0,
            key="process_range",
            help="전체 고객사를 선택하면 completed_with_images 폴더 내 모든 고객사 파일을 처리합니다"
        )
        
        customer_name = ""
        if process_all == "특정 고객사만 처리":
            customer_list = list(config.DASHBOARD_MAP.keys())
            customer_name = st.selectbox("고객사 선택", customer_list, key="customer_select")
        
        if not config.API_KEY:
            st.error("⚠️ Grafana API 키가 설정되지 않았습니다. 설정 탭에서 환경 변수를 확인하세요.")
        
        if st.button("▶️ 통계 삽입 실행", type="primary", key="run_stats"):
            cmd = ["python", "numinsert3.py"]
            if process_all == "특정 고객사만 처리" and customer_name:
                cmd.append(customer_name)
                st.info(f"📊 {customer_name} 고객사에 대해서만 통계를 삽입합니다.")
            else:
                st.info("📊 전체 고객사에 대해 통계를 삽입합니다.")
            
            with st.spinner("Grafana 통계를 삽입하는 중..."):
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True
                )
                
                st.subheader("실행 결과")
                
                # Parse output for better display
                output_text = result.stdout + result.stderr
                
                if result.returncode == 0:
                    st.success("✅ 통계 삽입이 완료되었습니다!")
                    
                    # Extract key information from output
                    import re
                    processed_files = re.findall(r"작업 \d+/\d+: '(.+?)'", output_text)
                    failed_placeholders = re.findall(r"- (\{\{.+?\}\})", output_text)
                    saved_files = re.findall(r"최종 보고서 저장 완료: (.+)", output_text)
                    
                    if processed_files:
                        st.info(f"📊 처리된 파일 수: {len(processed_files)}개")
                    
                    if failed_placeholders:
                        with st.expander(f"⚠️ Grafana 조회 실패 플레이스홀더 ({len(failed_placeholders)}개)", expanded=True):
                            for ph in failed_placeholders:
                                st.text(f"  • {ph}")
                            st.caption("이 플레이스홀더는 'N/A'로 대체되었습니다. Grafana 대시보드에 해당 패널이 있는지 확인하세요.")
                    
                    if saved_files:
                        st.success(f"✅ {len(saved_files)}개의 최종 보고서가 생성되었습니다.")
                        with st.expander("생성된 파일 경로"):
                            for file_path in saved_files:
                                st.code(file_path, language=None)
                    
                    completed_files = []
                    if os.path.exists(config.OUTPUT_DIR):
                        for f in os.listdir(config.OUTPUT_DIR):
                            if f.endswith('.pptx') and not f.startswith('~$'):
                                completed_files.append(f)
                    
                    if completed_files or os.path.exists(config.OUTPUT_DIR):
                        st.subheader("📥 최종 완료된 파일 다운로드")
                        
                        # Recursively find all files in OUTPUT_DIR, preserving full folder structure
                        folder_files = {}
                        for root, dirs, files in os.walk(config.OUTPUT_DIR):
                            for file_name in files:
                                if file_name.endswith('.pptx') and not file_name.startswith('~$'):
                                    rel_path = os.path.relpath(root, config.OUTPUT_DIR)
                                    # Use full relative path as the folder key (e.g., "GIT", "GIT2", "GIT3")
                                    folder_key = rel_path if rel_path != '.' else '루트'
                                    
                                    if folder_key not in folder_files:
                                        folder_files[folder_key] = []
                                    
                                    file_path = os.path.join(root, file_name)
                                    file_size = os.path.getsize(file_path) / 1024
                                    folder_files[folder_key].append({
                                        'name': file_name,
                                        'path': file_path,
                                        'size': file_size
                                    })
                        
                        if folder_files:
                            total_files = sum(len(files) for files in folder_files.values())
                            st.info(f"총 {total_files}개의 최종 보고서가 생성되었습니다.")
                            
                            for folder_path in sorted(folder_files.keys()):
                                files = folder_files[folder_path]
                                st.markdown(f"### 📁 {folder_path} ({len(files)}개)")
                                
                                for file_info in files:
                                    col_file, col_download = st.columns([3, 1])
                                    with col_file:
                                        st.text(f"📄 {file_info['name']} ({file_info['size']:.1f} KB)")
                                    
                                    with col_download:
                                        with open(file_info['path'], "rb") as f:
                                            # Use folder path in key to ensure uniqueness
                                            safe_key = folder_path.replace(os.sep, '_')
                                            st.download_button(
                                                label="다운로드",
                                                data=f,
                                                file_name=file_info['name'],
                                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                                key=f"download_final_{safe_key}_{file_info['name']}"
                                            )
                        else:
                            st.warning("생성된 파일이 없습니다.")
                else:
                    st.error("❌ 오류가 발생했습니다.")
                
                with st.expander("상세 로그 보기"):
                    st.code(result.stdout + result.stderr)

    with tab4:
        st.header("📋 작업 이력")
        st.markdown("최근 실행된 보고서 생성 작업을 확인할 수 있습니다.")
        
        # Get user's report runs
        if st.session_state['role'] == 'admin':
            # Admin can see all runs
            view_mode = st.radio("보기 모드", ["내 작업만", "전체 작업"], horizontal=True, key="view_mode")
            
            if view_mode == "전체 작업":
                runs = db_utils.get_all_report_runs(limit=100)
                st.info(f"전체 사용자의 작업 이력 ({len(runs)}개)")
            else:
                runs = db_utils.get_user_report_runs(st.session_state['user_id'], limit=50)
                st.info(f"내 작업 이력 ({len(runs)}개)")
        else:
            # Regular users see only their runs
            runs = db_utils.get_user_report_runs(st.session_state['user_id'], limit=50)
            st.info(f"내 작업 이력 ({len(runs)}개)")
        
        if runs:
            # Display runs in a table-like format
            for run in runs:
                with st.expander(f"📊 {run.created_at.strftime('%Y-%m-%d %H:%M:%S')} - {run.customer_name or '전체'} - {run.report_type}"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown(f"**작업 ID:** {run.id}")
                        if st.session_state['role'] == 'admin':
                            # Show username for admin
                            user = db_utils.get_user_by_username(run.user.username) if run.user else None
                            if user:
                                st.markdown(f"**사용자:** {user.full_name} ({user.username})")
                        st.markdown(f"**고객사:** {run.customer_name or '전체'}")
                    
                    with col2:
                        st.markdown(f"**보고서 유형:** {run.report_type}")
                        st.markdown(f"**템플릿:** {run.template_name or 'N/A'}")
                        st.markdown(f"**상태:** {run.status}")
                    
                    with col3:
                        st.markdown(f"**실행 시간:** {run.created_at.strftime('%Y-%m-%d %H:%M:%S')}")
                        if run.duration_seconds:
                            st.markdown(f"**소요 시간:** {run.duration_seconds:.1f}초")
                    
                    # Show generated files
                    files = db_utils.get_report_files_by_run_id(run.id)
                    if files:
                        st.markdown("**생성된 파일:**")
                        for file in files:
                            file_col1, file_col2 = st.columns([3, 1])
                            with file_col1:
                                st.text(f"📄 {file.filename} ({file.file_size / 1024:.1f} KB) - {file.step}")
                            with file_col2:
                                # Check if file still exists
                                if os.path.exists(file.file_path):
                                    with open(file.file_path, "rb") as f:
                                        st.download_button(
                                            label="다운로드",
                                            data=f,
                                            file_name=file.filename,
                                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                            key=f"download_history_{file.id}"
                                        )
                                else:
                                    st.text("파일 없음")
                    
                    # Show logs if available
                    if run.log_data:
                        with st.expander("로그 보기"):
                            st.code(run.log_data)
        else:
            st.info("아직 실행된 작업이 없습니다.")

    # User Management Tab (Admin only)
    if tab6 is not None and st.session_state['role'] == 'admin':
        with tab6:
            st.header("👥 사용자 관리")
            st.markdown("시스템의 사용자를 관리합니다.")
            
            # Get all users
            users = db_utils.get_all_users()
            
            if users:
                st.subheader(f"전체 사용자 ({len(users)}명)")
                
                for user in users:
                    with st.expander(f"{'🔴' if not user.is_active else '🟢'} {user.full_name} ({user.username}) - {user.role}"):
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.markdown(f"**사용자 ID:** {user.id}")
                            st.markdown(f"**사용자명:** {user.username}")
                            st.markdown(f"**이름:** {user.full_name}")
                        
                        with col2:
                            st.markdown(f"**이메일:** {user.email}")
                            st.markdown(f"**권한:** {user.role}")
                            st.markdown(f"**상태:** {'활성' if user.is_active else '비활성'}")
                        
                        with col3:
                            st.markdown(f"**가입일:** {user.created_at.strftime('%Y-%m-%d %H:%M:%S')}")
                            if user.last_login:
                                st.markdown(f"**마지막 로그인:** {user.last_login.strftime('%Y-%m-%d %H:%M:%S')}")
                            else:
                                st.markdown(f"**마지막 로그인:** 없음")
                        
                        st.divider()
                        
                        # User actions
                        action_col1, action_col2 = st.columns(2)
                        
                        with action_col1:
                            if user.username != st.session_state['username']:  # Can't deactivate self
                                if user.is_active:
                                    if st.button(f"🔴 비활성화", key=f"deactivate_{user.id}"):
                                        if db_utils.update_user_active_status(user.username, False):
                                            st.success(f"{user.username} 사용자가 비활성화되었습니다.")
                                            st.rerun()
                                        else:
                                            st.error("오류가 발생했습니다.")
                                else:
                                    if st.button(f"🟢 활성화", key=f"activate_{user.id}"):
                                        if db_utils.update_user_active_status(user.username, True):
                                            st.success(f"{user.username} 사용자가 활성화되었습니다.")
                                            st.rerun()
                                        else:
                                            st.error("오류가 발생했습니다.")
                        
                        with action_col2:
                            # Password reset
                            with st.form(f"reset_password_{user.id}"):
                                new_password = st.text_input("새 비밀번호", type="password", key=f"new_pass_{user.id}")
                                if st.form_submit_button("🔑 비밀번호 재설정"):
                                    if new_password:
                                        if db_utils.update_user_password(user.username, new_password):
                                            st.success(f"{user.username}의 비밀번호가 변경되었습니다.")
                                        else:
                                            st.error("오류가 발생했습니다.")
                                    else:
                                        st.error("비밀번호를 입력하세요.")
            else:
                st.info("등록된 사용자가 없습니다.")
            
            st.divider()
            
            # Create new user
            st.subheader("➕ 새 사용자 추가")
            with st.form("create_user_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    new_username = st.text_input("사용자명*")
                    new_email = st.text_input("이메일*")
                    new_password = st.text_input("비밀번호*", type="password")
                
                with col2:
                    new_full_name = st.text_input("이름*")
                    new_role = st.selectbox("권한", ["user", "admin"])
                
                if st.form_submit_button("사용자 생성", type="primary"):
                    if not all([new_username, new_email, new_password, new_full_name]):
                        st.error("모든 필드를 입력하세요.")
                    else:
                        try:
                            user_id = db_utils.create_user(
                                username=new_username,
                                email=new_email,
                                password=new_password,
                                full_name=new_full_name,
                                role=new_role
                            )
                            st.success(f"✅ 사용자 '{new_username}'가 생성되었습니다! (ID: {user_id})")
                            import time
                            time.sleep(1)
                            st.rerun()
                        except ValueError as e:
                            st.error(f"❌ {str(e)}")
                        except Exception as e:
                            st.error(f"❌ 오류가 발생했습니다: {str(e)}")

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

    with tab5:
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
            st.success("✅ 모든 폴더가 생성되었습니다!")
            st.rerun()
        
        st.divider()
        
        st.subheader("🔐 환경 변수")
        env_vars = {
            "GRAFANA_URL": os.getenv("GRAFANA_URL", "미설정"),
            "GRAFANA_API_KEY": "설정됨" if os.getenv("GRAFANA_API_KEY") else "미설정",
            "GRAFANA_VERIFY_SSL": os.getenv("GRAFANA_VERIFY_SSL", "미설정"),
            "DATABASE_URL": "설정됨" if os.getenv("DATABASE_URL") else "미설정"
        }
        
        for var, value in env_vars.items():
            col1, col2 = st.columns([2, 1])
            with col1:
                st.text(f"{var}")
            with col2:
                if value in ["미설정", "미설정됨"]:
                    st.error(f"❌ {value}")
                else:
                    st.success(f"✅ {value}")
        
        st.info("💡 환경 변수는 Replit Secrets 또는 아래 Grafana 설정에서 관리할 수 있습니다.")
        
        st.divider()
        
        st.subheader("🔧 Grafana 설정")
        
        with st.expander("Grafana 연결 설정 편집", expanded=False):
            st.markdown("""
            Grafana 서버 연결 정보를 설정합니다. 설정은 `.streamlit/secrets.toml` 파일에 저장됩니다.
            """)
            
            current_grafana_url = config.GRAFANA_URL or "http://localhost:3000"
            current_grafana_key = config.API_KEY
            current_verify_ssl = config.VERIFY_SSL
            
            new_grafana_url = st.text_input(
                "Grafana URL",
                value=current_grafana_url,
                help="Grafana 서버 주소 (예: http://grafana.example.com:3000)",
                key="grafana_url_input"
            )
            
            new_grafana_key = st.text_input(
                "Grafana API Key",
                value=current_grafana_key if current_grafana_key else "",
                type="password",
                help="Grafana API 토큰 (Service Account Token 권장)",
                key="grafana_key_input"
            )
            
            new_verify_ssl = st.checkbox(
                "SSL 인증서 검증",
                value=current_verify_ssl,
                help="프로덕션 환경에서는 항상 활성화하세요. 테스트 환경의 자체 서명 인증서인 경우만 비활성화하세요.",
                key="grafana_ssl_input"
            )
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("💾 Grafana 설정 저장", type="primary", key="save_grafana"):
                    try:
                        secrets_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit")
                        secrets_file = os.path.join(secrets_dir, "secrets.toml")
                        
                        os.makedirs(secrets_dir, exist_ok=True)
                        
                        existing_secrets = {}
                        if os.path.exists(secrets_file):
                            try:
                                import toml
                                existing_secrets = toml.load(secrets_file)
                            except:
                                pass
                        
                        existing_secrets["GRAFANA_URL"] = new_grafana_url
                        existing_secrets["GRAFANA_API_KEY"] = new_grafana_key
                        existing_secrets["GRAFANA_VERIFY_SSL"] = "true" if new_verify_ssl else "false"
                        
                        import toml
                        with open(secrets_file, "w") as f:
                            toml.dump(existing_secrets, f)
                        
                        st.success("✅ Grafana 설정이 저장되었습니다!")
                        st.info("⚠️ 변경사항을 적용하려면 브라우저를 새로고침하세요 (Ctrl+R 또는 F5)")
                    except Exception as e:
                        st.error(f"❌ 저장 중 오류 발생: {e}")
            
            with col2:
                if st.button("🧪 연결 테스트", key="test_grafana"):
                    with st.spinner("Grafana 서버에 연결 중..."):
                        try:
                            import requests
                            headers = {"Authorization": f"Bearer {new_grafana_key}"}
                            response = requests.get(
                                f"{new_grafana_url}/api/org",
                                headers=headers,
                                verify=new_verify_ssl,
                                timeout=5
                            )
                            if response.status_code == 200:
                                org_data = response.json()
                                st.success(f"✅ 연결 성공! 조직: {org_data.get('name', 'N/A')}")
                            else:
                                st.error(f"❌ 연결 실패 (HTTP {response.status_code})")
                        except requests.exceptions.SSLError:
                            st.error("❌ SSL 인증서 오류. 자체 서명 인증서인 경우 'SSL 인증서 검증'을 해제하세요.")
                        except requests.exceptions.ConnectionError:
                            st.error("❌ 연결 실패. Grafana URL을 확인하세요.")
                        except Exception as e:
                            st.error(f"❌ 오류: {str(e)}")
        
        st.divider()
        
        st.subheader("📊 고객사 대시보드 매핑")
        st.markdown("config.py에 정의된 고객사와 Grafana 대시보드 UID 매핑:")
        
        for customer, dashboard_uid in config.DASHBOARD_MAP.items():
            st.text(f"  • {customer}: {dashboard_uid}")
