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
    
    st.subheader("Grafana 설정")
    st.text(f"Grafana URL: {config.GRAFANA_URL}")
    st.text(f"API 키: {'*' * 20 if config.API_KEY else '미설정'}")
    st.text(f"SSL 검증: {'활성화' if config.VERIFY_SSL else '비활성화'}")
    
    st.info("""
    **환경 변수 설정 방법:**
    1. 왼쪽 사이드바에서 "Tools" 클릭
    2. "Secrets" 선택
    3. 다음 키를 추가:
       - `GRAFANA_URL`: Grafana 서버 URL
       - `GRAFANA_API_KEY`: API 키
       - `GRAFANA_VERIFY_SSL`: SSL 검증 (true/false)
    """)
    
    st.divider()
    
    st.subheader("고객사 대시보드 매핑")
    st.text(f"총 {len(config.DASHBOARD_MAP)}개의 고객사가 설정되어 있습니다.")
    
    with st.expander("고객사 목록 보기"):
        for customer, uid in config.DASHBOARD_MAP.items():
            st.text(f"• {customer}: {uid}")
    
    st.info("고객사 매핑을 변경하려면 `config.py` 파일을 수정하세요.")

st.divider()
st.caption("PowerPoint 자동화 도구 v1.0 | Replit에서 실행 중")
