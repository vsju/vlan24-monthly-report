import streamlit as st
import streamlit.components.v1 as components
import requests
import os
import json
import time
import urllib.parse

API_URL = os.getenv("BACKEND_API_URL", "http://192.168.10.30:5001")
BACKEND_API_PUBLIC_URL = os.getenv("BACKEND_API_PUBLIC_URL", "")
ONLYOFFICE_URL = os.getenv("ONLYOFFICE_URL", "http://121.78.82.22:8502")

if st.query_params.get('reset') == '1':
    st.query_params.clear()
    if 'template_list_select' in st.session_state:
        del st.session_state['template_list_select']
    st.rerun()


@st.cache_data(ttl=60, show_spinner=False)
def fetch_customers_cached(api_url):
    """고객사 목록을 캐싱하여 반복 API 호출 방지"""
    try:
        resp = requests.get(f"{api_url}/api/customers", timeout=5)
        if resp.ok:
            return resp.json().get('customers', [])
    except Exception:
        pass
    return []

@st.cache_data(ttl=120, show_spinner=False)
def fetch_templates_cached(api_url, template_type, refresh_ts=None):
    """템플릿 목록을 캐싱하여 반복 API 호출 방지"""
    try:
        resp = requests.get(f"{api_url}/api/templates", params={'type': template_type}, timeout=10)
        if resp.ok:
            return resp.json().get('templates', [])
    except Exception:
        pass
    return []

@st.cache_data(ttl=300, show_spinner=False)
def fetch_template_info_cached(api_url, encoded_template, template_type_key, refresh_ts=None):
    """템플릿 정보를 캐싱하여 반복 API 호출 방지"""
    try:
        resp = requests.get(f"{api_url}/api/templates/{encoded_template}/info?type={template_type_key}", timeout=30)
        if resp.ok:
            return resp.json().get('template', {})
    except Exception:
        pass
    return None

@st.cache_data(ttl=300, show_spinner=False)
def fetch_tables_cached(api_url, encoded_template, template_type_key, refresh_ts):
    """테이블 데이터를 캐싱하여 반복 API 호출 방지"""
    resp = requests.get(f"{api_url}/api/templates/{encoded_template}/tables?type={template_type_key}", timeout=30)
    if resp.ok:
        return {"success": True, "tables": resp.json().get('tables', [])}
    return {"success": False, "error": f"HTTP {resp.status_code}"}


st.set_page_config(
    page_title="PowerPoint 자동화",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
    .stTextInput > div > div > input {
        padding: 0.3rem 0.5rem;
        font-size: 0.85rem;
    }
    .stSelectbox > div > div {
        padding: 0.2rem 0.4rem;
        font-size: 0.85rem;
        min-height: 2rem;
    }
    .stButton > button {
        padding: 0.25rem 0.75rem;
        font-size: 0.85rem;
    }
    div[data-testid="stFormSubmitButton"] > button {
        padding: 0.25rem 0.5rem;
        font-size: 0.8rem;
    }
    .stExpander {
        font-size: 0.9rem;
    }
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.5rem;
    }
    .stCaption {
        font-size: 0.75rem;
    }
</style>
""", unsafe_allow_html=True)


if not BACKEND_API_PUBLIC_URL or "localhost" in BACKEND_API_PUBLIC_URL:
    st.warning(
        "BACKEND_API_PUBLIC_URL 환경 변수가 설정되지 않았거나 localhost를 사용합니다. "
        "PPT 에디터가 제대로 작동하려면 브라우저에서 접근 가능한 백엔드 URL을 설정하세요. "
        "(예: http://192.168.10.30:5001)"
    )

st.title("📊 PowerPoint 자동화 도구")
st.markdown("이미지와 Grafana 통계를 자동으로 PowerPoint에 삽입하는 도구입니다.")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🏠 홈",
    "🎨 이미지 렌더링",
    "🖼️ 이미지 삽입",
    "📈 통계 삽입",
    "👥 고객사 관리",
    "📄 템플릿 관리",
    "📥 다운로드",
    "⚙️ 설정"
])

with tab1:
    st.header("📊 PowerPoint 자동화 도구")
    
    try:
        response = requests.get(f"{API_URL}/health", timeout=5)
        if response.ok:
            data = response.json()
            col1, col2, col3 = st.columns(3)
            with col1:
                st.success("✅ 서버 연결됨")
            with col2:
                st.info(f"🌐 Grafana: {data.get('grafana_url', 'N/A')}")
            with col3:
                status = "✅ 설정됨" if data.get("grafana_configured") else "❌ 미설정"
                st.info(f"🔑 API: {status}")
        else:
            st.error("❌ 백엔드 서버 연결 실패")
    except Exception as e:
        st.error(f"❌ 서버 연결 불가: {str(e)}")
    
    st.divider()
    
    st.subheader("📋 작업 순서 가이드")
    st.markdown("""
    보고서 생성은 아래 순서로 진행합니다:
    
    | 순서 | 탭 | 작업 내용 |
    |:---:|:---|:---|
    | 1️⃣ | **이미지 렌더링** | Grafana 대시보드 패널을 PNG 이미지로 저장 |
    | 2️⃣ | **이미지 삽입** | 템플릿에 이미지 삽입 → 이미지삽입 템플릿 생성 |
    | 3️⃣ | **통계 삽입** | Grafana 통계 데이터 삽입 → 최종 보고서 완성 |
    | 4️⃣ | **다운로드** | 완성된 보고서 다운로드 |
    """)
    
    st.divider()
    
    st.subheader("📖 각 탭 사용 설명서")
    
    with st.expander("🎨 이미지 렌더링", expanded=False):
        st.markdown("""
        **목적:** Grafana 대시보드의 패널들을 PNG 이미지로 렌더링하여 고객사 폴더에 저장
        
        **사용 방법:**
        1. 고객사 선택
        2. 대시보드 UID 확인 (고객사 관리에서 설정)
        3. 렌더링할 패널 선택 또는 전체 선택
        4. `이미지 렌더링` 버튼 클릭
        5. 완료 후 고객사 폴더에 이미지 파일 생성됨
        
        **주의:** 렌더링 전 Grafana 연결 상태 확인 필요
        """)
    
    with st.expander("🖼️ 이미지 삽입", expanded=False):
        st.markdown("""
        **목적:** PowerPoint 템플릿의 도형에 이미지를 자동 삽입
        
        **사용 방법:**
        1. **이미지 업로드:** 고객사 선택 → 파일 업로드 (선택사항)
        2. **이미지 삽입 실행:**
           - 고객사 선택 (전체 선택 시 일괄 처리)
           - 기준 날짜 설정
           - `이미지 삽입 실행` 버튼 클릭
        3. 결과: `이미지삽입` 템플릿 폴더에 결과물 저장
        
        **원리:** 템플릿 내 도형 이름과 동일한 파일명의 이미지를 자동 매칭
        """)
    
    with st.expander("📈 통계 삽입", expanded=False):
        st.markdown("""
        **목적:** Grafana에서 통계 데이터를 가져와 플레이스홀더에 삽입
        
        **사용 방법:**
        1. 고객사 선택
        2. 날짜 범위 설정 (시작일 ~ 종료일)
        3. `통계 삽입 실행` 버튼 클릭
        4. 결과: `통계삽입` 템플릿 폴더에 최종 보고서 저장
        
        **플레이스홀더 형식:** `{{패널명_쿼리문자}}` (예: `{{CPU사용률_A}}`)
        """)
    
    with st.expander("👥 고객사 관리", expanded=False):
        st.markdown("""
        **목적:** 고객사별 정보 및 Grafana 대시보드 매핑 관리
        
        **주요 기능:**
        - **고객사 추가/삭제:** 새 고객사 폴더 생성 또는 삭제
        - **대시보드 UID 설정:** 고객사별 Grafana 대시보드 UID 연결
        - **폴더 구조 확인:** 고객사 이미지 폴더 내용 조회
        """)
    
    with st.expander("📄 템플릿 관리", expanded=False):
        st.markdown("""
        **목적:** PowerPoint 템플릿 조회, 편집, 생성 관리
        
        **주요 기능:**
        - **템플릿 조회:** 원본/이미지삽입/통계삽입 템플릿 목록
        - **슬라이드 정보:** 각 슬라이드의 도형, 플레이스홀더 확인
        - **Shape 편집:** 도형 이름 변경으로 이미지 매핑 설정
        - **VM 슬라이드 복제:** 패턴 슬라이드 복제 및 템플릿 생성
        - **PPT 에디터:** 온라인에서 직접 템플릿 편집
        """)
    
    with st.expander("📥 다운로드", expanded=False):
        st.markdown("""
        **목적:** 생성된 보고서 및 템플릿 파일 다운로드
        
        **사용 방법:**
        1. 템플릿 유형 선택 (원본/이미지삽입/통계삽입)
        2. 파일 목록에서 원하는 파일 선택
        3. 다운로드 버튼 클릭
        """)
    
    with st.expander("⚙️ 설정", expanded=False):
        st.markdown("""
        **목적:** 시스템 설정 및 연결 정보 관리
        
        **주요 설정:**
        - Grafana API URL 및 인증 정보
        - 백엔드 서버 연결 상태 확인
        - 기타 시스템 설정
        """)

with tab2:
    st.header("🎨 이미지 렌더링")
    st.info("💡 Grafana 대시보드 패널을 PNG 이미지로 렌더링하여 고객사 폴더에 저장합니다.")
    
    if 'render_mode' not in st.session_state:
        st.session_state['render_mode'] = 'all'
    
    render_col1, render_col2 = st.columns(2)
    with render_col1:
        if st.button("🚀 전체 렌더링", use_container_width=True, type="primary" if st.session_state.get('render_mode') == 'all' else "secondary"):
            st.session_state['render_mode'] = 'all'
            st.rerun()
    with render_col2:
        if st.button("🎯 개별 렌더링", use_container_width=True, type="primary" if st.session_state.get('render_mode') == 'individual' else "secondary"):
            st.session_state['render_mode'] = 'individual'
            st.rerun()
    
    st.divider()
    
    if st.session_state.get('render_mode') == 'all':
        st.subheader("🚀 전체 고객사 렌더링")
        st.markdown("모든 고객사의 대시보드 패널을 한 번에 렌더링합니다.")
        if st.button("▶️ 전체 고객사 렌더링 실행", type="primary"):
            with st.spinner("모든 고객사 패널 렌더링 중... (시간이 오래 걸릴 수 있습니다)"):
                try:
                    resp = requests.post(
                        f"{API_URL}/api/render/all-customers",
                        timeout=600
                    )
                    if resp.ok:
                        result = resp.json()
                        if result.get('success'):
                            st.success(f"전체 렌더링 완료: {len(result.get('customers_rendered', []))}개 고객사, 총 {result.get('total_panels_rendered', 0)}개 패널")
                        else:
                            st.warning(f"부분 완료: {len(result.get('customers_rendered', []))}개 성공, {len(result.get('customers_failed', []))}개 실패")
                        
                        if result.get('customers_rendered'):
                            with st.expander("성공한 고객사"):
                                for c in result.get('customers_rendered', []):
                                    st.text(f"{c['name']}: {c['rendered_count']}개 성공, {c['failed_count']}개 실패")
                        
                        if result.get('customers_failed'):
                            with st.expander("실패한 고객사"):
                                for c in result.get('customers_failed', []):
                                    st.text(f"{c['name']}: {c['error']}")
                        
                        with st.expander("상세 로그"):
                            for log in result.get('logs', []):
                                st.text(log)
                    else:
                        st.error(f"렌더링 실패: {resp.json().get('error', 'Unknown error')}")
                except Exception as e:
                    st.error(f"오류: {str(e)}")
    
    elif st.session_state.get('render_mode') == 'individual':
        st.subheader("🎯 개별 고객사 렌더링")
        try:
            response = requests.get(f"{API_URL}/api/customers", timeout=5)
            if response.ok:
                customers_data = response.json().get('customers', [])
                customer_options = ["선택하세요"] + [c['name'] for c in customers_data if c.get('dashboard_uid')]
                
                render_customer = st.selectbox(
                    "고객사 선택 (개별 렌더링)",
                    customer_options,
                    key="render_customer_select"
                )
                
                if render_customer and render_customer != "선택하세요":
                    if st.session_state.get('last_render_customer') != render_customer:
                        st.session_state['last_render_customer'] = render_customer
                        st.session_state['selected_panels'] = set()
                        st.session_state['render_panels_loaded'] = False
                    
                    customer_info = next((c for c in customers_data if c['name'] == render_customer), None)
                    if customer_info:
                        st.info(f"대시보드 UID: `{customer_info.get('dashboard_uid', 'N/A')}`")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("🔄 패널 목록 조회", type="primary"):
                            st.session_state['render_panels_loaded'] = True
                    with col2:
                        if st.button("🎨 전체 렌더링", type="secondary"):
                            with st.spinner("전체 패널 렌더링 중..."):
                                try:
                                    resp = requests.post(
                                        f"{API_URL}/api/render/all",
                                        json={"customer_name": render_customer},
                                        timeout=300
                                    )
                                    if resp.ok:
                                        result = resp.json()
                                        if result.get('success'):
                                            st.success(f"렌더링 완료: {len(result.get('rendered', []))}개 성공")
                                            st.info(f"저장 위치: `{result.get('save_dir', '')}`")
                                        else:
                                            st.warning(f"부분 실패: {len(result.get('rendered', []))}개 성공, {len(result.get('failed', []))}개 실패")
                                        
                                        with st.expander("상세 로그"):
                                            for log in result.get('logs', []):
                                                st.text(log)
                                    else:
                                        st.error(f"렌더링 실패: {resp.json().get('error', 'Unknown error')}")
                                except Exception as e:
                                    st.error(f"오류: {str(e)}")
                    
                    if st.session_state.get('render_panels_loaded'):
                        try:
                            panels_resp = requests.get(f"{API_URL}/api/render/panels/{render_customer}", timeout=30)
                            if panels_resp.ok:
                                panels_data = panels_resp.json()
                                if panels_data.get('success'):
                                    panel_rows = panels_data.get('panel_rows', {})
                                    all_panels = panels_data.get('all_panels', [])
                                    
                                    st.markdown(f"**대시보드:** {panels_data.get('dashboard_title', '')} | **총 패널 수:** {len(all_panels)}")
                                    
                                    search_term = st.text_input("🔍 패널 검색 (title)", key="panel_search")
                                    
                                    if 'selected_panels' not in st.session_state:
                                        st.session_state['selected_panels'] = set()
                                    
                                    row_options = ["전체"] + list(panel_rows.keys())
                                    selected_row = st.selectbox(
                                        "📁 Row 선택",
                                        row_options,
                                        key="render_row_select"
                                    )
                                    
                                    if selected_row == "전체":
                                        display_panels = all_panels
                                    else:
                                        display_panels = panel_rows.get(selected_row, [])
                                    
                                    if search_term:
                                        display_panels = [p for p in display_panels if search_term.lower() in p['title'].lower()]
                                    
                                    col_all, col_none = st.columns(2)
                                    with col_all:
                                        if st.button("✅ 전체 선택"):
                                            st.session_state['selected_panels'] = set(p['id'] for p in display_panels)
                                            st.rerun()
                                    with col_none:
                                        if st.button("❌ 전체 해제"):
                                            st.session_state['selected_panels'] = set()
                                            st.rerun()
                                    
                                    st.markdown(f"**표시 중인 패널: {len(display_panels)}개**")
                                    for panel in display_panels:
                                        panel_id = panel['id']
                                        is_selected = panel_id in st.session_state['selected_panels']
                                        
                                        if st.checkbox(
                                            f"[{panel_id}] {panel['title']}",
                                            value=is_selected,
                                            key=f"panel_cb_{render_customer}_{panel_id}"
                                        ):
                                            st.session_state['selected_panels'].add(panel_id)
                                        else:
                                            st.session_state['selected_panels'].discard(panel_id)
                                    
                                    selected_count = len(st.session_state['selected_panels'])
                                    st.divider()
                                    st.markdown(f"**선택된 패널: {selected_count}개**")
                                    
                                    if selected_count > 0:
                                        if st.button(f"🎨 선택한 {selected_count}개 패널 렌더링", type="primary"):
                                            with st.spinner("선택한 패널 렌더링 중..."):
                                                try:
                                                    resp = requests.post(
                                                        f"{API_URL}/api/render/selected",
                                                        json={
                                                            "customer_name": render_customer,
                                                            "panel_ids": list(st.session_state['selected_panels'])
                                                        },
                                                        timeout=300
                                                    )
                                                    if resp.ok:
                                                        result = resp.json()
                                                        if result.get('success'):
                                                            st.success(f"렌더링 완료: {len(result.get('rendered', []))}개 성공")
                                                            st.info(f"저장 위치: `{result.get('save_dir', '')}`")
                                                        else:
                                                            st.warning(f"부분 실패: {len(result.get('rendered', []))}개 성공, {len(result.get('failed', []))}개 실패")
                                                        
                                                        with st.expander("상세 로그"):
                                                            for log in result.get('logs', []):
                                                                st.text(log)
                                                        
                                                        if result.get('rendered'):
                                                            with st.expander("저장된 파일 목록"):
                                                                for item in result.get('rendered', []):
                                                                    st.text(f"[{item['panel_id']}] {item['title']} → {item['path']}")
                                                    else:
                                                        st.error(f"렌더링 실패: {resp.json().get('error', 'Unknown error')}")
                                                except Exception as e:
                                                    st.error(f"오류: {str(e)}")
                                else:
                                    st.error(panels_data.get('error', '패널 정보를 가져올 수 없습니다.'))
                            else:
                                st.error("패널 목록 조회 실패")
                        except Exception as e:
                            st.error(f"패널 조회 오류: {str(e)}")
                else:
                    st.info("대시보드 UID가 설정된 고객사를 선택하세요.")
            else:
                st.error("고객사 목록을 불러올 수 없습니다.")
        except Exception as e:
            st.error(f"오류: {str(e)}")

with tab3:
    st.header("🖼️ 이미지 삽입 (Step 1)")
    st.info("💡 템플릿 도형 이름과 일치하는 이미지를 자동 삽입합니다. 이미지 업로드 후 삽입을 실행하세요.")
    
    if 'img_insert_mode' not in st.session_state:
        st.session_state['img_insert_mode'] = 'upload'
    
    img_col1, img_col2 = st.columns(2)
    with img_col1:
        if st.button("📤 이미지 업로드", use_container_width=True, type="primary" if st.session_state.get('img_insert_mode') == 'upload' else "secondary"):
            st.session_state['img_insert_mode'] = 'upload'
            st.rerun()
    with img_col2:
        if st.button("▶️ 삽입 실행", use_container_width=True, type="primary" if st.session_state.get('img_insert_mode') == 'execute' else "secondary"):
            st.session_state['img_insert_mode'] = 'execute'
            st.rerun()
    
    st.divider()
    
    if st.session_state.get('img_insert_mode') == 'upload':
        st.subheader("📤 이미지 업로드")
        ROOT_REPORT_OPTION = "📁 루트 (/root/Report)"
        img_customer_options = [ROOT_REPORT_OPTION] + [c['name'] for c in fetch_customers_cached(API_URL)]
        
        if img_customer_options:
            if 'image_customer_select' in st.session_state:
                if st.session_state['image_customer_select'] not in img_customer_options:
                    del st.session_state['image_customer_select']
        
            image_customer_selection = st.selectbox(
                "고객사 선택 (필수)",
                options=img_customer_options,
                key="image_customer_select"
            )
            
            is_root_report = (image_customer_selection == ROOT_REPORT_OPTION)
            image_customer = "" if is_root_report else image_customer_selection
            
            existing_subdirs = []
            if is_root_report:
                try:
                    subdir_resp = requests.get(f"{API_URL}/api/report-root/subdirs", timeout=10)
                    if subdir_resp.ok:
                        existing_subdirs = subdir_resp.json().get('subdirs', [])
                except Exception:
                    pass
            elif image_customer:
                try:
                    subdir_resp = requests.get(f"{API_URL}/api/customers/{image_customer}/subdirs", timeout=10)
                    if subdir_resp.ok:
                        existing_subdirs = subdir_resp.json().get('subdirs', [])
                except Exception:
                    pass
            
            NEW_FOLDER_OPTION = "➕ 새 폴더 생성..."
            ROOT_OPTION = "📁 루트 (고객사 폴더)"
            
            subdir_options = [ROOT_OPTION]
            for sd in existing_subdirs:
                subdir_options.append(f"📂 {sd['name']} ({sd['image_count']}개 이미지)")
            subdir_options.append(NEW_FOLDER_OPTION)
            
            subdir_map = {ROOT_OPTION: ""}
            for sd in existing_subdirs:
                subdir_map[f"📂 {sd['name']} ({sd['image_count']}개 이미지)"] = sd['name']
            
            selected_subdir_option = st.selectbox(
                "저장할 폴더 선택",
                options=subdir_options,
                key="image_subdir_select"
            )
            
            if selected_subdir_option == NEW_FOLDER_OPTION:
                new_folder_name = st.text_input(
                    "새 폴더 이름",
                    key="new_folder_name",
                    placeholder="예: 2024-12"
                )
                image_subdir = new_folder_name
            else:
                image_subdir = subdir_map.get(selected_subdir_option, "")
            
            if image_subdir and image_customer:
                try:
                    img_resp = requests.get(f"{API_URL}/api/customers/{image_customer}/images", 
                                           params={"subdir": image_subdir}, timeout=10)
                    if img_resp.ok:
                        existing_images = img_resp.json().get('images', [])
                        if existing_images:
                            with st.expander(f"📷 기존 이미지 ({len(existing_images)}개)", expanded=False):
                                img_names = [img['name'] for img in existing_images]
                                st.write(", ".join(img_names[:20]))
                                if len(img_names) > 20:
                                    st.caption(f"... 외 {len(img_names) - 20}개")
                except Exception:
                    pass
            
            uploaded_images = st.file_uploader(
                "이미지 파일을 선택하세요",
                type=['png', 'jpg', 'jpeg', 'gif'],
                accept_multiple_files=True,
                key="image_uploader"
            )
            
            can_upload = uploaded_images and (is_root_report or image_customer)
            if can_upload and st.button("📤 이미지 업로드", type="primary"):
                try:
                    files = [('files', (f.name, f, f'image/{f.type.split("/")[1]}')) for f in uploaded_images]
                    data = {'customer': image_customer, 'root_report': 'true' if is_root_report else 'false'}
                    if image_subdir:
                        data['subdir'] = image_subdir
                    
                    response = requests.post(f"{API_URL}/api/upload/images", files=files, data=data, timeout=30)
                    
                    if response.ok:
                        result = response.json()
                        if result.get('success'):
                            st.success(f"✅ {len(result.get('uploaded_files', []))}개 이미지 업로드 완료!")
                            if result.get('errors'):
                                st.warning("⚠️ 일부 이미지 업로드 실패:\n" + "\n".join(result['errors']))
                            st.rerun()
                        else:
                            st.error(f"❌ 업로드 실패: {result.get('error', '알 수 없는 오류')}")
                    else:
                        st.error(f"❌ 서버 오류: {response.status_code}")
                except Exception as e:
                    st.error(f"❌ 업로드 실패: {str(e)}")
        else:
            st.warning("등록된 고객사가 없습니다. 먼저 고객사를 추가하세요.")
    
    elif st.session_state.get('img_insert_mode') == 'execute':
        st.subheader("▶️ 이미지 삽입 실행")
        
        proc_customer_options = ["전체"] + [c['name'] for c in fetch_customers_cached(API_URL)]
        
        if 'process_customer_select' in st.session_state:
            if st.session_state['process_customer_select'] not in proc_customer_options:
                del st.session_state['process_customer_select']
        
        col1, col2 = st.columns([2, 1])
        with col1:
            process_customer_select = st.selectbox(
                "고객사 선택",
                options=proc_customer_options,
                key="process_customer_select"
            )
            process_customer = process_customer_select if process_customer_select != "전체" else ""
        
        with col2:
            st.write("")
            st.write("")
            run_images = st.button("🚀 이미지 삽입 실행", type="primary", use_container_width=True)
        
        if run_images:
            with st.spinner("처리 중..."):
                try:
                    payload = {"customer_name": process_customer if process_customer else None}
                    response = requests.post(f"{API_URL}/api/process/images", json=payload, timeout=300)
                    
                    if response.ok:
                        result = response.json()
                        
                        logs = result.get('logs', [])
                        if logs:
                            st.markdown("**📋 진행 로그**")
                            log_text = "\n".join(logs)
                            st.markdown(
                                f'<div style="background-color:#1e1e1e; color:#d4d4d4; padding:12px; '
                                f'border-radius:6px; height:300px; overflow-y:auto; font-family:monospace; '
                                f'font-size:13px; white-space:pre-wrap;">{log_text}</div>',
                                unsafe_allow_html=True
                            )
                        
                        if result.get('success'):
                            st.success(f"✅ 이미지 삽입 완료! ({result['summary']['processed_count']}개 파일 처리)")
                            
                            with st.expander("처리 결과 상세", expanded=False):
                                for file_info in result.get('processed_files', []):
                                    st.write(f"📄 {file_info['template']} - {file_info['images_inserted']}개 이미지 삽입")
                            
                            if result.get('errors'):
                                with st.expander("⚠️ 오류 내역"):
                                    for error in result['errors']:
                                        st.error(error)
                        else:
                            st.error("❌ 이미지 삽입 실패")
                            for error in result.get('errors', []):
                                st.error(error)
                    else:
                        st.error(f"❌ 서버 오류: {response.status_code}")
                except Exception as e:
                    st.error(f"❌ 실행 실패: {str(e)}")

with tab4:
    st.header("📈 통계 삽입 (Step 2)")
    st.info("💡 Grafana에서 통계 데이터를 가져와 플레이스홀더({{패널명_쿼리}})에 삽입합니다.")
    
    st.subheader("▶️ 통계 삽입 실행")
    st.info("💡 이미지 삽입이 완료된 파일에 Grafana 통계를 삽입합니다.")
    
    stats_customer_options = ["전체"] + [c['name'] for c in fetch_customers_cached(API_URL)]
    
    if 'stats_customer_select' in st.session_state:
        if st.session_state['stats_customer_select'] not in stats_customer_options:
            del st.session_state['stats_customer_select']
    
    col1, col2 = st.columns([2, 1])
    with col1:
        stats_customer_select = st.selectbox(
            "고객사 선택",
            options=stats_customer_options,
            key="stats_customer_select"
        )
        stats_customer = stats_customer_select if stats_customer_select != "전체" else ""
    
    with col2:
        st.write("")
        st.write("")
        run_stats = st.button("🚀 통계 삽입 실행", type="primary", use_container_width=True)
    
    if run_stats:
        with st.spinner("처리 중... (Grafana 조회 시간이 소요될 수 있습니다)"):
            try:
                payload = {"customer_name": stats_customer if stats_customer else None}
                response = requests.post(f"{API_URL}/api/process/statistics", json=payload, timeout=600)
                
                if response.ok:
                    result = response.json()
                    
                    logs = result.get('logs', [])
                    if logs:
                        st.markdown("**📋 진행 로그**")
                        log_text = "\n".join(logs)
                        st.markdown(
                            f'<div style="background-color:#1e1e1e; color:#d4d4d4; padding:12px; '
                            f'border-radius:6px; height:300px; overflow-y:auto; font-family:monospace; '
                            f'font-size:13px; white-space:pre-wrap;">{log_text}</div>',
                            unsafe_allow_html=True
                        )
                    
                    if result.get('success'):
                        st.success(f"✅ 통계 삽입 완료! ({result['summary']['processed_count']}개 파일 처리)")
                        
                        with st.expander("처리 결과 상세", expanded=False):
                            for file_info in result.get('processed_files', []):
                                st.write(f"📄 {file_info['template']} - {file_info.get('grafana_queries', 0)}개 통계 조회")
                        
                        if result.get('failed_placeholders'):
                            with st.expander(f"⚠️ 실패한 플레이스홀더 ({len(result['failed_placeholders'])}개)"):
                                for failed in result['failed_placeholders']:
                                    st.warning(f"{failed['placeholder']}: {failed['reason']}")
                        
                        if result.get('errors'):
                            with st.expander("⚠️ 오류 내역"):
                                for error in result['errors']:
                                    st.error(error)
                    else:
                        st.error("❌ 통계 삽입 실패")
                        for error in result.get('errors', []):
                            st.error(error)
                else:
                    st.error(f"❌ 서버 오류: {response.status_code}")
            except Exception as e:
                st.error(f"❌ 실행 실패: {str(e)}")

with tab5:
    st.header("👥 고객사 관리")
    st.info("💡 고객사 추가/삭제 및 Grafana 대시보드 UID를 매핑합니다.")
    
    if 'customer_mgmt_mode' not in st.session_state:
        st.session_state['customer_mgmt_mode'] = 'add'
    
    cust_col1, cust_col2 = st.columns(2)
    with cust_col1:
        if st.button("➕ 신규 추가", use_container_width=True, type="primary" if st.session_state.get('customer_mgmt_mode') == 'add' else "secondary"):
            st.session_state['customer_mgmt_mode'] = 'add'
            st.rerun()
    with cust_col2:
        if st.button("📋 목록/수정", use_container_width=True, type="primary" if st.session_state.get('customer_mgmt_mode') == 'manage' else "secondary"):
            st.session_state['customer_mgmt_mode'] = 'manage'
            st.rerun()
    
    st.divider()
    
    if st.session_state.get('customer_mgmt_mode') == 'add':
        st.subheader("➕ 신규 고객사 추가")
        
        col1, col2 = st.columns(2)
        with col1:
            new_customer_name = st.text_input("폴더명 (name)", key="new_customer_name", help="이미지/템플릿 폴더에 사용되는 이름 (영문/숫자 권장)")
            new_display_name = st.text_input("표시명 (display_name)", key="new_display_name", help="보고서에 표시될 이름 (한글 가능)")
            new_dashboard_uid = st.text_input("Grafana 대시보드 UID", key="new_dashboard_uid", help="Grafana 대시보드 URL에서 확인 가능")
        with col2:
            st.markdown("**담당자 정보**")
            new_contact_name = st.text_input("담당자 이름", key="new_contact_name")
            new_contact_phone = st.text_input("연락처", key="new_contact_phone")
            new_contact_email = st.text_input("이메일", key="new_contact_email")
        
        if st.button("➕ 고객사 추가", type="primary"):
            if new_customer_name:
                try:
                    response = requests.post(
                        f"{API_URL}/api/customers",
                        json={
                            "name": new_customer_name, 
                            "display_name": new_display_name,
                            "dashboard_uid": new_dashboard_uid,
                            "contact_name": new_contact_name,
                            "contact_phone": new_contact_phone,
                            "contact_email": new_contact_email
                        },
                        timeout=10
                    )
                    if response.ok:
                        result = response.json()
                        st.success(result.get('message', '고객사 추가 완료'))
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"추가 실패: {response.json().get('error', '알 수 없는 오류')}")
                except Exception as e:
                    st.error(f"오류 발생: {str(e)}")
            else:
                st.warning("폴더명(name)을 입력하세요.")
    
    else:
        st.subheader("📋 고객사 관리")
        
        try:
            customers = fetch_customers_cached(API_URL)
            if customers:
                customer_names = [c['name'] for c in customers]
                customer_map = {c['name']: c for c in customers}
                
                if 'selected_customer_manage' in st.session_state:
                    if st.session_state['selected_customer_manage'] not in customer_names:
                        del st.session_state['selected_customer_manage']
                
                def format_customer_option(name):
                    c = customer_map.get(name, {})
                    disp = c.get('display_name', '')
                    if disp:
                        return f"🏢 {name} ({disp})"
                    return f"🏢 {name}"
                
                selected_customer_name = st.selectbox(
                    "고객사 선택",
                    options=customer_names,
                    key="selected_customer_manage",
                    format_func=format_customer_option
                )
                
                if selected_customer_name:
                    customer = customer_map[selected_customer_name]
                    
                    st.markdown("---")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("##### 기본 정보")
                        st.write(f"**폴더명 (name):** {customer.get('name', '')}")
                        st.write(f"**표시명 (display_name):** {customer.get('display_name', '') or '미설정'}")
                        st.write(f"**대시보드 UID:** {customer.get('dashboard_uid', '') or '미설정'}")
                        st.write(f"**이미지 폴더:** {'✅ 있음' if customer.get('has_images') else '❌ 없음'}")
                        
                        st.markdown("##### 담당자 정보")
                        st.write(f"**담당자:** {customer.get('contact_name', '') or '미설정'}")
                        st.write(f"**연락처:** {customer.get('contact_phone', '') or '미설정'}")
                        st.write(f"**이메일:** {customer.get('contact_email', '') or '미설정'}")
                    
                    with col2:
                        st.markdown("##### 기본 정보 수정")
                        new_display = st.text_input(
                            "표시명",
                            value=customer.get('display_name', ''),
                            key=f"display_{selected_customer_name}"
                        )
                        new_uid = st.text_input(
                            "대시보드 UID",
                            value=customer.get('dashboard_uid', ''),
                            key=f"uid_{selected_customer_name}"
                        )
                        
                        st.markdown("##### 담당자 정보 수정")
                        edit_contact_name = st.text_input(
                            "담당자 이름",
                            value=customer.get('contact_name', ''),
                            key=f"contact_name_{selected_customer_name}"
                        )
                        edit_contact_phone = st.text_input(
                            "연락처",
                            value=customer.get('contact_phone', ''),
                            key=f"contact_phone_{selected_customer_name}"
                        )
                        edit_contact_email = st.text_input(
                            "이메일",
                            value=customer.get('contact_email', ''),
                            key=f"contact_email_{selected_customer_name}"
                        )
                        
                        if st.button("💾 저장", key=f"save_{selected_customer_name}"):
                            try:
                                resp = requests.put(
                                    f"{API_URL}/api/customers/{selected_customer_name}",
                                    json={
                                        "dashboard_uid": new_uid, 
                                        "display_name": new_display,
                                        "contact_name": edit_contact_name,
                                        "contact_phone": edit_contact_phone,
                                        "contact_email": edit_contact_email
                                    },
                                    timeout=10
                                )
                                if resp.ok:
                                    st.success("고객사 정보 업데이트 완료")
                                    st.cache_data.clear()
                                    st.rerun()
                                else:
                                    st.error("업데이트 실패")
                            except Exception as e:
                                st.error(str(e))
                    
                    st.markdown("---")
                    
                    st.markdown("##### ⚠️ 고객사 삭제")
                    delete_files = st.checkbox("관련 파일도 함께 삭제", key=f"del_files_{selected_customer_name}")
                    if st.button("🗑️ 고객사 삭제", key=f"delete_{selected_customer_name}", type="secondary"):
                        try:
                            resp = requests.delete(
                                f"{API_URL}/api/customers/{selected_customer_name}",
                                json={"delete_files": delete_files},
                                timeout=10
                            )
                            if resp.ok:
                                st.success(f"고객사 '{selected_customer_name}' 삭제 완료")
                                st.cache_data.clear()
                                st.rerun()
                            else:
                                st.error("삭제 실패")
                        except Exception as e:
                            st.error(str(e))
            else:
                st.info("등록된 고객사가 없습니다.")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")

with tab6:
    st.header("📄 템플릿 관리")
    st.info("💡 템플릿 조회, 슬라이드/도형 편집, VM 슬라이드 복제, PPT 에디터 등을 사용할 수 있습니다.")
    
    if 'last_tab5_access' not in st.session_state:
        st.session_state['last_tab5_access'] = time.time()
    else:
        elapsed = time.time() - st.session_state['last_tab5_access']
        if elapsed > 300:
            keys_to_remove = [k for k in st.session_state.keys() if k.startswith('vm_table_select_') or k.startswith('table_') or k.startswith('form_table_')]
            for k in keys_to_remove:
                del st.session_state[k]
            st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
        st.session_state['last_tab5_access'] = time.time()
    
    if 'template_mgmt_mode' not in st.session_state:
        st.session_state['template_mgmt_mode'] = 'upload'
    
    tmpl_col1, tmpl_col2, tmpl_col3, tmpl_col4 = st.columns(4)
    with tmpl_col1:
        if st.button("📤 업로드", use_container_width=True, type="primary" if st.session_state.get('template_mgmt_mode') == 'upload' else "secondary"):
            st.session_state['template_mgmt_mode'] = 'upload'
            st.rerun()
    with tmpl_col2:
        if st.button("➕ 자동 생성", use_container_width=True, type="primary" if st.session_state.get('template_mgmt_mode') == 'generate' else "secondary"):
            st.session_state['template_mgmt_mode'] = 'generate'
            st.rerun()
    with tmpl_col3:
        if st.button("🔧 VM 추가", use_container_width=True, type="primary" if st.session_state.get('template_mgmt_mode') == 'add_vm' else "secondary"):
            st.session_state['template_mgmt_mode'] = 'add_vm'
            st.rerun()
    with tmpl_col4:
        if st.button("📝 편집기", use_container_width=True, type="primary" if st.session_state.get('template_mgmt_mode') == 'editor' else "secondary"):
            st.session_state['template_mgmt_mode'] = 'editor'
            st.rerun()
    
    st.divider()
    
    if st.session_state.get('template_mgmt_mode') == 'upload':
        st.subheader("📤 템플릿 업로드")
        uploaded_templates = st.file_uploader(
            "PowerPoint 템플릿 파일을 선택하세요 (.pptx)",
            type=['pptx'],
            accept_multiple_files=True,
            key="template_uploader"
        )
        
        template_customer = st.text_input(
            "고객사 이름 (선택사항)",
            key="template_customer",
            help="비워두면 루트 디렉토리에 저장됩니다."
        )
        
        if uploaded_templates and st.button("📤 템플릿 업로드", type="primary"):
            try:
                files = [('files', (f.name, f, 'application/vnd.openxmlformats-officedocument.presentationml.presentation')) 
                         for f in uploaded_templates]
                data = {'customer': template_customer}
                
                response = requests.post(f"{API_URL}/api/upload/templates", files=files, data=data, timeout=30)
                
                if response.ok:
                    result = response.json()
                    if result.get('success'):
                        st.success(f"✅ {len(result.get('uploaded_files', []))}개 파일 업로드 완료!")
                        if result.get('errors'):
                            st.warning("⚠️ 일부 파일 업로드 실패:\n" + "\n".join(result['errors']))
                    else:
                        st.error(f"❌ 업로드 실패: {result.get('error', '알 수 없는 오류')}")
                else:
                    st.error(f"❌ 서버 오류: {response.status_code}")
            except Exception as e:
                st.error(f"❌ 업로드 실패: {str(e)}")
    
    elif st.session_state.get('template_mgmt_mode') == 'generate':
        st.subheader("📄 신규 템플릿 자동 생성")
        st.info("마스터 템플릿 + 이미지 폴더 분석 → VM별 슬라이드 자동 생성")
        
        if 'template_gen_result' in st.session_state:
            gen_msg = st.session_state['template_gen_result']
            st.success(f"✅ 템플릿 생성 완료! ({gen_msg.get('customer', '')})")
            st.info(f"VM {gen_msg.get('vm_count', 0)}개, 슬라이드 {gen_msg.get('slide_count', 0)}개")
            if gen_msg.get('logs'):
                with st.expander("📋 생성 로그"):
                    for log in gen_msg['logs']:
                        st.text(log)
            if st.button("✖️ 메시지 닫기", key="close_gen_msg"):
                del st.session_state['template_gen_result']
                st.rerun()
        
        try:
            target_customers = [c['name'] for c in fetch_customers_cached(API_URL)]
            
            if target_customers:
                gen_col1, gen_col2 = st.columns(2)
                
                with gen_col1:
                    if 'gen_target_customer' in st.session_state:
                        if st.session_state['gen_target_customer'] not in target_customers:
                            del st.session_state['gen_target_customer']
                    
                    gen_customer = st.selectbox(
                        "고객사 선택",
                        options=target_customers,
                        key="gen_target_customer"
                    )
                
                with gen_col2:
                    if st.button("🔍 VM 분석", type="secondary"):
                        st.session_state['gen_analyze_clicked'] = True
                
                if gen_customer and st.session_state.get('gen_analyze_clicked'):
                    try:
                        analyze_resp = requests.get(f"{API_URL}/api/customers/{gen_customer}/analyze-vms", timeout=15)
                        if analyze_resp.ok:
                            analyze_data = analyze_resp.json()
                            vms = analyze_data.get('vms', [])
                            
                            if vms:
                                st.success(f"✅ VM {len(vms)}개 발견")
                                
                                with st.expander("📊 VM 분석 결과", expanded=True):
                                    for vm in vms:
                                        vm_name = vm.get('vm_name', 'Unknown')
                                        vm_ip = vm.get('ip', '')
                                        resources = vm.get('resources', [])
                                        pages = vm.get('pages_needed', 0)
                                        
                                        st.markdown(f"**{vm_name}** ({vm_ip})")
                                        st.caption(f"리소스 {len(resources)}개 → 슬라이드 {pages}페이지")
                                        
                                        if resources:
                                            res_names = [r.get('name', '') for r in resources]
                                            st.write(f"  리소스: {', '.join(res_names)}")
                                
                                if st.button("🚀 템플릿 자동 생성", type="primary", key="gen_template_btn"):
                                    with st.spinner("템플릿 생성 중..."):
                                        try:
                                            gen_resp = requests.post(
                                                f"{API_URL}/api/templates/generate",
                                                json={"customer_name": gen_customer},
                                                timeout=120
                                            )
                                            if gen_resp.ok:
                                                gen_result = gen_resp.json()
                                                if gen_result.get('success'):
                                                    st.session_state['template_gen_result'] = {
                                                        'customer': gen_customer,
                                                        'vm_count': gen_result.get('vm_count', 0),
                                                        'slide_count': gen_result.get('slide_count', 0),
                                                        'logs': gen_result.get('logs', [])
                                                    }
                                                    st.session_state['gen_analyze_clicked'] = False
                                                    st.rerun()
                                                else:
                                                    st.error(f"❌ 생성 실패: {', '.join(gen_result.get('errors', []))}")
                                            else:
                                                st.error(f"❌ 서버 오류: {gen_resp.status_code}")
                                        except Exception as e:
                                            st.error(f"❌ 오류: {str(e)}")
                            else:
                                st.warning("이미지 폴더에 VM 디렉토리가 없습니다.")
                        else:
                            st.error("VM 분석 실패")
                    except Exception as e:
                        st.error(f"분석 오류: {str(e)}")
            else:
                st.warning("등록된 고객사가 없습니다. 먼저 고객사를 추가하세요.")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")
    
    elif st.session_state.get('template_mgmt_mode') == 'add_vm':
        st.subheader("➕ 기존 템플릿에 VM 추가")
        st.info("기존 고객사 템플릿에 새로운 VM 슬라이드를 추가합니다.")
        
        if 'vm_add_result' in st.session_state:
            add_msg = st.session_state['vm_add_result']
            st.success(f"✅ VM 추가 완료! ({add_msg.get('vm_name', '')} - SEQ: {add_msg.get('seq', '')})")
            st.info(f"슬라이드 {add_msg.get('slides_added', 0)}개 추가됨")
            if add_msg.get('logs'):
                with st.expander("📋 추가 로그"):
                    for log in add_msg['logs']:
                        st.text(log)
            if st.button("✖️ 메시지 닫기", key="close_vm_add_msg"):
                del st.session_state['vm_add_result']
                st.rerun()
        
        try:
            vm_add_customers = [c['name'] for c in fetch_customers_cached(API_URL)]
            
            if vm_add_customers:
                vm_add_col1, vm_add_col2 = st.columns(2)
                
                with vm_add_col1:
                    if 'vm_add_customer' in st.session_state:
                        if st.session_state['vm_add_customer'] not in vm_add_customers:
                            del st.session_state['vm_add_customer']
                    
                    vm_add_customer = st.selectbox(
                        "고객사 선택",
                        options=vm_add_customers,
                        key="vm_add_customer"
                    )
                
                vm_folders = []
                if vm_add_customer:
                    try:
                        subdir_resp = requests.get(f"{API_URL}/api/customers/{vm_add_customer}/subdirs", timeout=10)
                        if subdir_resp.ok:
                            vm_folders = [sd['name'] for sd in subdir_resp.json().get('subdirs', [])]
                    except Exception:
                        pass
                
                with vm_add_col2:
                    if vm_folders:
                        vm_folder_select = st.selectbox(
                            "추가할 VM 폴더 선택",
                            options=vm_folders,
                            key="vm_add_folder"
                        )
                    else:
                        st.warning("VM 폴더가 없습니다.")
                        vm_folder_select = None
                
                if vm_add_customer and vm_folder_select:
                    if st.button("➕ VM 슬라이드 추가", type="primary", key="add_vm_btn"):
                        with st.spinner("VM 슬라이드 추가 중..."):
                            try:
                                add_resp = requests.post(
                                    f"{API_URL}/api/templates/{vm_add_customer}/add-vm",
                                    json={"vm_dir_name": vm_folder_select},
                                    timeout=120
                                )
                                if add_resp.ok:
                                    add_result = add_resp.json()
                                    if add_result.get('success'):
                                        st.session_state['vm_add_result'] = {
                                            'vm_name': add_result.get('vm_name', ''),
                                            'vm_ip': add_result.get('vm_ip', ''),
                                            'seq': add_result.get('seq', ''),
                                            'slides_added': add_result.get('slides_added', 0),
                                            'logs': add_result.get('logs', [])
                                        }
                                        st.rerun()
                                    else:
                                        st.error(f"❌ 추가 실패: {', '.join(add_result.get('errors', []))}")
                                else:
                                    error_msg = add_resp.json().get('error', '알 수 없는 오류') if add_resp.content else f"HTTP {add_resp.status_code}"
                                    st.error(f"❌ 서버 오류: {error_msg}")
                            except Exception as e:
                                st.error(f"❌ 오류: {str(e)}")
            else:
                st.warning("등록된 고객사가 없습니다.")
        except Exception as e:
            st.error(f"오류 발생: {str(e)}")
    
    elif st.session_state.get('template_mgmt_mode') == 'editor':
        st.subheader("📝 템플릿 편집기")
        st.info("💡 PPT 에디터에서 직접 편집하거나, 하단의 도구로 플레이스홀더/Shape를 수정할 수 있습니다.")
        
        try:
            all_templates_resp = requests.get(f"{API_URL}/api/templates/all", timeout=10)
            if all_templates_resp.ok:
                all_templates_data = all_templates_resp.json()
                
                source_templates = all_templates_data.get('templates', [])
                with_images_templates = all_templates_data.get('with_images', [])
                final_templates = all_templates_data.get('final', [])
                
                if not source_templates and not with_images_templates and not final_templates:
                    st.info("편집할 템플릿이 없습니다. 먼저 템플릿을 업로드하세요.")
                else:
                    if 'current_edit_template' not in st.session_state:
                        st.session_state['current_edit_template'] = None
                    if 'current_template_type' not in st.session_state:
                        st.session_state['current_template_type'] = None
                    
                    editor_type_tabs = st.tabs(["📁 원본", "🖼️ 이미지삽입", "✅ 통계삽입(최종)"])
                    
                    def on_template_change(template_type, session_key):
                        """템플릿 선택 변경 시 캐시 클리어 및 상태 업데이트"""
                        selected = st.session_state.get(session_key)
                        if selected and not selected.startswith("—"):
                            current = st.session_state.get('current_edit_template')
                            new_template = f"{template_type}:{selected}"
                            if current == new_template:
                                return
                            
                            keys_to_clear = [k for k in list(st.session_state.keys()) 
                                            if k.startswith('previews_') or k.startswith('info_') 
                                            or k.startswith('loading_') or k.startswith('load_previews_')
                                            or k.startswith('slide_') or k.startswith('shape_')
                                            or k.startswith('placeholder_') or k.startswith('vm_')
                                            or k.startswith('edit_vm_') or k.startswith('selected_vm_')]
                            for k in keys_to_clear:
                                st.session_state.pop(k, None)
                            
                            st.session_state['current_edit_template'] = new_template
                            st.session_state['current_template_type'] = template_type
                    
                    EDITOR_SENTINEL = "— 템플릿을 선택하세요 —"
                    
                    with editor_type_tabs[0]:
                        if source_templates:
                            source_paths = [t.get('path', str(t)) if isinstance(t, dict) else str(t) for t in source_templates]
                            source_path_to_info = {t.get('path', str(t)) if isinstance(t, dict) else str(t): t for t in source_templates}
                            source_options = [EDITOR_SENTINEL] + source_paths
                            
                            if 'edit_template_source' in st.session_state:
                                if st.session_state['edit_template_source'] not in source_options:
                                    del st.session_state['edit_template_source']
                            
                            col_select, col_refresh = st.columns([4, 1])
                            with col_select:
                                selected_source = st.selectbox(
                                    "원본 템플릿 선택",
                                    options=source_options,
                                    key="edit_template_source",
                                    on_change=on_template_change,
                                    args=("source", "edit_template_source")
                                )
                                if selected_source and selected_source != EDITOR_SENTINEL and not st.session_state.get('current_edit_template'):
                                    st.session_state['current_edit_template'] = f"source:{selected_source}"
                                    st.session_state['current_template_type'] = "source"
                            with col_refresh:
                                st.write("")
                                if st.button("🔄", key="refresh_editor_source", help="새로고침"):
                                    keys_to_clear = [k for k in list(st.session_state.keys()) 
                                                    if k.startswith('previews_') or k.startswith('info_') 
                                                    or k.startswith('table_') or k.startswith('new_row_')]
                                    for k in keys_to_clear:
                                        del st.session_state[k]
                                    st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                    st.rerun()
                        else:
                            st.info("원본 템플릿이 없습니다.")
                    
                    with editor_type_tabs[1]:
                        if with_images_templates:
                            images_paths = [t.get('path', str(t)) if isinstance(t, dict) else str(t) for t in with_images_templates]
                            images_path_to_info = {t.get('path', str(t)) if isinstance(t, dict) else str(t): t for t in with_images_templates}
                            images_options = [EDITOR_SENTINEL] + images_paths
                            
                            if 'edit_template_with_images' in st.session_state:
                                if st.session_state['edit_template_with_images'] not in images_options:
                                    del st.session_state['edit_template_with_images']
                            
                            col_select, col_refresh = st.columns([4, 1])
                            with col_select:
                                selected_images = st.selectbox(
                                    "이미지삽입 템플릿 선택",
                                    options=images_options,
                                    key="edit_template_with_images",
                                    on_change=on_template_change,
                                    args=("with_images", "edit_template_with_images")
                                )
                            with col_refresh:
                                st.write("")
                                if st.button("🔄", key="refresh_editor_images", help="새로고침"):
                                    keys_to_clear = [k for k in list(st.session_state.keys()) 
                                                    if k.startswith('previews_') or k.startswith('info_') 
                                                    or k.startswith('table_') or k.startswith('new_row_')]
                                    for k in keys_to_clear:
                                        del st.session_state[k]
                                    st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                    st.rerun()
                        else:
                            st.info("이미지삽입 템플릿이 없습니다.")
                    
                    with editor_type_tabs[2]:
                        if final_templates:
                            final_paths = [t.get('path', str(t)) if isinstance(t, dict) else str(t) for t in final_templates]
                            final_path_to_info = {t.get('path', str(t)) if isinstance(t, dict) else str(t): t for t in final_templates}
                            final_options = [EDITOR_SENTINEL] + final_paths
                            
                            if 'edit_template_final' in st.session_state:
                                if st.session_state['edit_template_final'] not in final_options:
                                    del st.session_state['edit_template_final']
                            
                            col_select, col_refresh = st.columns([4, 1])
                            with col_select:
                                selected_final = st.selectbox(
                                    "최종(통계삽입) 템플릿 선택",
                                    options=final_options,
                                    key="edit_template_final",
                                    on_change=on_template_change,
                                    args=("final", "edit_template_final")
                                )
                            with col_refresh:
                                st.write("")
                                if st.button("🔄", key="refresh_editor_final", help="새로고침"):
                                    keys_to_clear = [k for k in list(st.session_state.keys()) 
                                                    if k.startswith('previews_') or k.startswith('info_') 
                                                    or k.startswith('table_') or k.startswith('new_row_')]
                                    for k in keys_to_clear:
                                        del st.session_state[k]
                                    st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                    st.rerun()
                        else:
                            st.info("최종(통계삽입) 템플릿이 없습니다.")
                    
                    selected_template = st.session_state.get('current_edit_template')
                    
                    current_source = st.session_state.get('edit_template_source', EDITOR_SENTINEL)
                    current_images = st.session_state.get('edit_template_with_images', EDITOR_SENTINEL)
                    current_final = st.session_state.get('edit_template_final', EDITOR_SENTINEL)
                    
                    if selected_template:
                        if selected_template.startswith("source:"):
                            template_type_key = "source"
                            if current_source == EDITOR_SENTINEL:
                                selected_template = None
                        elif selected_template.startswith("with_images:"):
                            template_type_key = "with_images"
                            if current_images == EDITOR_SENTINEL:
                                selected_template = None
                        elif selected_template.startswith("final:"):
                            template_type_key = "final"
                            if current_final == EDITOR_SENTINEL:
                                selected_template = None
                        else:
                            template_type_key = None
                    else:
                        template_type_key = None
                    
                    if selected_template and template_type_key:
                        actual_template_path = selected_template.split(":", 1)[1] if ":" in selected_template else selected_template
                        
                        preview_cache_key = f"previews_{selected_template}"
                        info_cache_key = f"info_{selected_template}"
                        encoded_template = urllib.parse.quote(actual_template_path, safe='')
                        
                        if info_cache_key not in st.session_state:
                            try:
                                info_resp = requests.get(
                                    f"{API_URL}/api/templates/{encoded_template}/info?type={template_type_key}",
                                    timeout=30
                                )
                                if info_resp.ok:
                                    st.session_state[info_cache_key] = info_resp.json().get('template', {})
                            except Exception:
                                st.session_state[info_cache_key] = None
                        
                        cached_info = st.session_state.get(info_cache_key)
                        
                        st.markdown("---")
                        
                        type_labels = {"source": "📁 원본", "with_images": "🖼️ 이미지삽입", "final": "✅ 최종"}
                        st.info(f"📌 현재 편집 중: **{type_labels.get(template_type_key, template_type_key)}** - `{actual_template_path}`")
                        
                        st.markdown("#### 🖥️ PPT 온라인 에디터")
                        
                        public_url = BACKEND_API_PUBLIC_URL if BACKEND_API_PUBLIC_URL else API_URL
                        editor_refresh_ts = st.session_state.get('editor_refresh_ts', '')
                        refresh_param = f"&refresh={editor_refresh_ts}" if editor_refresh_ts else ""
                        editor_page_url = f"{public_url}/api/onlyoffice/editor-page?template={encoded_template}&type={template_type_key}{refresh_param}"
                        
                        iframe_html = f"""
                        <div style="width: 100%; height: 800px; border: 1px solid #ddd; border-radius: 4px; overflow: hidden;">
                            <iframe 
                                id="onlyoffice-frame"
                                src="{editor_page_url}" 
                                style="width: 100%; height: 100%; border: none;"
                                allow="fullscreen"
                            ></iframe>
                        </div>
                        """
                        st.markdown(iframe_html, unsafe_allow_html=True)
                        st.caption("💾 변경사항은 자동으로 저장됩니다.")
                        
                        st.markdown("#### 🛠️ 상세 편집 도구")
                        
                        edit_tabs = st.tabs(["📊 슬라이드 정보", "🏷️ Shape 편집", "📝 플레이스홀더 관리", "➕ VM 슬라이드 복제", "📋 VM 표 편집"])
                        
                        with edit_tabs[0]:
                            st.write("**슬라이드 정보**")
                            if cached_info:
                                info = cached_info
                                total_slides = info.get('slide_count', 0)
                                st.write(f"총 슬라이드 수: {total_slides}")
                                
                                if total_slides > 0:
                                    slide_info_idx = st.selectbox(
                                        "슬라이드 선택",
                                        options=list(range(total_slides)),
                                        format_func=lambda x: f"슬라이드 {x + 1} ({info.get('slides', [])[x].get('shape_count', 0)}개 Shape)" if x < len(info.get('slides', [])) else f"슬라이드 {x + 1}",
                                        key="slide_info_select"
                                    )
                                    
                                    if slide_info_idx is not None:
                                        selected_slide = info.get('slides', [])[slide_info_idx] if slide_info_idx < len(info.get('slides', [])) else None
                                        if selected_slide:
                                            st.markdown(f"**슬라이드 {slide_info_idx + 1}** - Shape {selected_slide.get('shape_count', 0)}개")
                                            
                                            slide_placeholders = selected_slide.get('placeholders', [])
                                            if slide_placeholders:
                                                st.markdown("**📋 플레이스홀더:**")
                                                st.code(", ".join(slide_placeholders))
                                            
                                            for shape in selected_slide.get('shapes', []):
                                                icon = "🖼️" if shape['type'] == 'PICTURE' else "📦"
                                                with st.container(border=shape['type'] == 'PICTURE'):
                                                    st.write(f"{icon} **{shape['name']}** ({shape['type']})")
                                                    if shape.get('text'):
                                                        st.caption(shape['text'][:100] + "..." if len(shape.get('text', '')) > 100 else shape['text'])
                                else:
                                    st.info("슬라이드가 없습니다.")
                            else:
                                st.warning("템플릿 정보를 불러올 수 없습니다.")
                        
                        with edit_tabs[1]:
                            st.write("**Shape 이름 편집**")
                            st.info("💡 Shape 이름은 이미지 자동 삽입 시 파일명과 매칭됩니다. 예: Shape 이름이 'CPU_Chart'이면 'CPU_Chart.png' 파일이 삽입됩니다.")
                            
                            if cached_info:
                                info = cached_info
                                
                                filter_col1, filter_col2 = st.columns(2)
                                with filter_col1:
                                    slide_idx = st.selectbox(
                                        "슬라이드 선택",
                                        options=list(range(info.get('slide_count', 0))),
                                        format_func=lambda x: f"슬라이드 {x + 1}",
                                        key="shape_edit_slide"
                                    )
                                with filter_col2:
                                    shape_filter = st.selectbox(
                                        "Shape 타입 필터",
                                        options=["전체", "PICTURE (이미지)", "PLACEHOLDER", "TEXT_BOX", "AUTO_SHAPE", "TABLE", "기타"],
                                        key="shape_type_filter"
                                    )
                                
                                if slide_idx is not None:
                                    slide_info = info.get('slides', [])[slide_idx] if slide_idx < len(info.get('slides', [])) else None
                                    if slide_info:
                                        shapes = slide_info.get('shapes', [])
                                        
                                        if shape_filter != "전체":
                                            filter_type = shape_filter.split(" ")[0]
                                            if filter_type == "기타":
                                                main_types = ["PICTURE", "PLACEHOLDER", "TEXT_BOX", "AUTO_SHAPE", "TABLE"]
                                                shapes = [s for s in shapes if s['type'] not in main_types]
                                            else:
                                                shapes = [s for s in shapes if s['type'] == filter_type]
                                        
                                        if shapes:
                                            st.write(f"**{len(shapes)}개 Shape**")
                                            
                                            for shape in shapes:
                                                is_picture = shape['type'] == 'PICTURE'
                                                icon = "🖼️" if is_picture else "📦"
                                                
                                                with st.container(border=is_picture):
                                                    col1, col2, col3 = st.columns([2, 2, 1])
                                                    with col1:
                                                        st.write(f"{icon} **{shape['type']}** (ID: {shape['id']})")
                                                        if is_picture:
                                                            st.caption("🎯 이미지 삽입 대상")
                                                    with col2:
                                                        new_name = st.text_input(
                                                            "Shape 이름",
                                                            value=shape['name'],
                                                            key=f"shape_name_{slide_idx}_{shape['id']}",
                                                            label_visibility="collapsed"
                                                        )
                                                    with col3:
                                                        if st.button("💾 저장", key=f"save_shape_{slide_idx}_{shape['id']}"):
                                                            try:
                                                                resp = requests.put(
                                                                    f"{API_URL}/api/templates/{encoded_template}/shapes/{slide_idx}/{shape['id']}?type={template_type_key}",
                                                                    json={"name": new_name},
                                                                    timeout=10
                                                                )
                                                                if resp.ok:
                                                                    st.success("Shape 이름 변경 완료")
                                                                    if info_cache_key in st.session_state:
                                                                        del st.session_state[info_cache_key]
                                                                    if preview_cache_key in st.session_state:
                                                                        del st.session_state[preview_cache_key]
                                                                    st.rerun()
                                                                else:
                                                                    st.error("변경 실패")
                                                            except Exception as e:
                                                                st.error(str(e))
                                        else:
                                            st.info("해당 조건에 맞는 Shape가 없습니다.")
                            else:
                                st.warning("템플릿 정보를 불러올 수 없습니다.")
                        
                        with edit_tabs[2]:
                            st.write("**플레이스홀더 관리**")
                            
                            try:
                                ph_resp = requests.get(f"{API_URL}/api/templates/{encoded_template}/placeholders?type={template_type_key}", timeout=30)
                                if ph_resp.ok:
                                    ph_data = ph_resp.json()
                                    
                                    if cached_info:
                                        total_slides = cached_info.get('slide_count', 0)
                                        slide_options = ["전체 슬라이드"] + [f"슬라이드 {i+1}" for i in range(total_slides)]
                                        ph_slide_filter = st.selectbox(
                                            "슬라이드 선택",
                                            options=slide_options,
                                            key="ph_slide_filter"
                                        )
                                    else:
                                        ph_slide_filter = "전체 슬라이드"
                                    
                                    st.write("**현재 플레이스홀더:**")
                                    
                                    placeholders_to_show = []
                                    if ph_slide_filter == "전체 슬라이드":
                                        placeholders_to_show = ph_data.get('unique_placeholders', [])
                                    else:
                                        slide_num = int(ph_slide_filter.replace("슬라이드 ", "")) - 1
                                        all_phs = ph_data.get('placeholders', [])
                                        slide_phs = [p['placeholder'] for p in all_phs if p.get('slide_index') == slide_num]
                                        placeholders_to_show = list(set(slide_phs))
                                    
                                    if not placeholders_to_show:
                                        st.info("해당 슬라이드에 플레이스홀더가 없습니다.")
                                    
                                    for ph in placeholders_to_show:
                                        col1, col2, col3 = st.columns([2, 2, 1])
                                        with col1:
                                            st.code(ph)
                                        with col2:
                                            new_ph_text = st.text_input(
                                                "변경할 텍스트",
                                                value=ph,
                                                key=f"edit_ph_{ph}",
                                                label_visibility="collapsed"
                                            )
                                        with col3:
                                            btn_cols = st.columns(2)
                                            with btn_cols[0]:
                                                if st.button("💾", key=f"save_ph_{ph}", help="수정"):
                                                    if new_ph_text and new_ph_text != ph:
                                                        try:
                                                            resp = requests.put(
                                                                f"{API_URL}/api/templates/{encoded_template}/placeholders/update?type={template_type_key}",
                                                                json={"old_placeholder": ph, "new_placeholder": new_ph_text},
                                                                timeout=10
                                                            )
                                                            if resp.ok:
                                                                result = resp.json()
                                                                st.success(f"플레이스홀더 수정 완료 ({result.get('replaced_count', 0)}개)")
                                                                st.rerun()
                                                            else:
                                                                st.error("수정 실패")
                                                        except Exception as e:
                                                            st.error(str(e))
                                            with btn_cols[1]:
                                                if st.button("🗑️", key=f"del_ph_{ph}", help="삭제"):
                                                    try:
                                                        resp = requests.delete(
                                                            f"{API_URL}/api/templates/{encoded_template}/placeholders?type={template_type_key}",
                                                            json={"placeholder": ph},
                                                            timeout=10
                                                        )
                                                        if resp.ok:
                                                            st.success(f"플레이스홀더 '{ph}' 삭제 완료")
                                                            st.rerun()
                                                        else:
                                                            st.error("삭제 실패")
                                                    except Exception as e:
                                                        st.error(str(e))
                                    
                                    st.divider()
                                    st.write("**플레이스홀더 추가**")
                                    
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        new_ph = st.text_input("플레이스홀더 이름", placeholder="예: CPU-Usage_A", key="new_placeholder")
                                    with col2:
                                        add_slide_idx = st.number_input("슬라이드 번호", min_value=1, value=1, key="add_ph_slide") - 1
                                    
                                    if st.button("➕ 플레이스홀더 추가", type="primary"):
                                        if new_ph:
                                            try:
                                                resp = requests.post(
                                                    f"{API_URL}/api/templates/{encoded_template}/placeholders?type={template_type_key}",
                                                    json={"placeholder": new_ph, "slide_index": add_slide_idx},
                                                    timeout=10
                                                )
                                                if resp.ok:
                                                    st.success("플레이스홀더 추가 완료")
                                                    st.rerun()
                                                else:
                                                    st.error("추가 실패")
                                            except Exception as e:
                                                st.error(str(e))
                                        else:
                                            st.warning("플레이스홀더 이름을 입력하세요.")
                            except Exception as e:
                                st.error(str(e))
                        
                        with edit_tabs[3]:
                            st.write("**VM 슬라이드 복제**")
                            st.info("기존 VM 슬라이드를 복제하여 새 VM 슬라이드를 생성합니다. 순번이 자동으로 증가합니다 (예: 3.1 → 3.2).")
                            
                            try:
                                info_resp = requests.get(f"{API_URL}/api/templates/{encoded_template}/info?type={template_type_key}", timeout=30)
                                if info_resp.ok:
                                    info = info_resp.json().get('template', {})
                                    slides = info.get('slides', [])
                                    slide_count = info.get('slide_count', 0)
                                    
                                    slide_options = []
                                    for slide in slides:
                                        slide_idx = slide['index']
                                        slide_text = ""
                                        for shape in slide.get('shapes', []):
                                            if shape.get('text'):
                                                text = shape['text'].strip()
                                                import re
                                                if re.match(r'^\d+\.\d+\s+', text):
                                                    slide_text = text[:50]
                                                    break
                                        if slide_text:
                                            slide_options.append((slide_idx, f"슬라이드 {slide_idx + 1}: {slide_text}"))
                                        else:
                                            slide_options.append((slide_idx, f"슬라이드 {slide_idx + 1}"))
                                    
                                    if slide_options:
                                        preview_col, form_col = st.columns([1, 1])
                                        
                                        with form_col:
                                            selected_slide_option = st.selectbox(
                                                "복제할 슬라이드 선택",
                                                options=slide_options,
                                                format_func=lambda x: x[1],
                                                key="duplicate_slide_select"
                                            )
                                            
                                            st.divider()
                                            st.write("**새 VM 정보 입력**")
                                            
                                            new_vm_name = st.text_input("VM 이름 *", placeholder="예: PMO-DB3", key="new_vm_name")
                                            new_vm_ip = st.text_input("IP 주소", placeholder="예: 192.168.1.100", key="new_vm_ip")
                                            new_vm_os = st.text_input("OS", placeholder="예: Ubuntu20.04", key="new_vm_os")
                                            
                                            if st.button("➕ 슬라이드 복제", type="primary"):
                                                if new_vm_name:
                                                    try:
                                                        resp = requests.post(
                                                            f"{API_URL}/api/templates/{encoded_template}/slides/duplicate?type={template_type_key}",
                                                            json={
                                                                "slide_index": selected_slide_option[0],
                                                                "vm_name": new_vm_name,
                                                                "vm_ip": new_vm_ip,
                                                                "vm_os": new_vm_os
                                                            },
                                                            timeout=30
                                                        )
                                                        if resp.ok:
                                                            result = resp.json()
                                                            st.success(f"슬라이드 복제 완료: {result.get('new_sequence', '')} {new_vm_name}")
                                                            if result.get('vm_table_updated'):
                                                                st.info("VM 목록 표에도 행이 추가되었습니다.")
                                                            st.rerun()
                                                        else:
                                                            result = resp.json()
                                                            st.error(f"복제 실패: {result.get('error', '알 수 없는 오류')}")
                                                    except Exception as e:
                                                        st.error(str(e))
                                                else:
                                                    st.warning("VM 이름을 입력하세요.")
                                        
                                        with preview_col:
                                            st.write("**슬라이드 미리보기**")
                                            if selected_slide_option:
                                                slide_idx = selected_slide_option[0]
                                                try:
                                                    with st.spinner("미리보기 생성 중..."):
                                                        preview_resp = requests.get(
                                                            f"{API_URL}/api/templates/{encoded_template}/slides/{slide_idx}/preview?type={template_type_key}",
                                                            timeout=60
                                                        )
                                                        if preview_resp.ok:
                                                            preview_data = preview_resp.json()
                                                            if preview_data.get('image'):
                                                                st.image(preview_data['image'], use_container_width=True)
                                                            elif preview_data.get('shapes'):
                                                                st.info("이미지 미리보기를 사용할 수 없습니다. 텍스트 내용:")
                                                                for shape in preview_data['shapes'][:10]:
                                                                    if shape.get('text'):
                                                                        st.text(shape['text'][:100])
                                                            else:
                                                                st.warning("미리보기를 생성할 수 없습니다.")
                                                        else:
                                                            st.warning("미리보기 로드 실패")
                                                except requests.exceptions.Timeout:
                                                    st.warning("미리보기 생성 시간 초과")
                                                except Exception as e:
                                                    st.warning(f"미리보기 오류: {str(e)[:50]}")
                                    else:
                                        st.info("복제할 수 있는 슬라이드가 없습니다.")
                            except Exception as e:
                                st.error(str(e))
                        
                        with edit_tabs[4]:
                            st.write("**VM 표 편집**")
                            
                            try:
                                refresh_ts = st.session_state.get('editor_refresh_ts', '0')
                                with st.spinner("표 목록 불러오는 중..."):
                                    tables_result = fetch_tables_cached(API_URL, encoded_template, template_type_key, refresh_ts)
                                if not tables_result.get('success'):
                                    st.error(f"표 정보를 불러올 수 없습니다: {tables_result.get('error', '알 수 없는 오류')}")
                                else:
                                    tables = tables_result.get('tables', [])
                                    if not tables:
                                        st.info("템플릿에 표가 없습니다.")
                                    else:
                                        table_options = [f"슬라이드 {t['slide_index'] + 1}: {t['shape_name']} ({t['col_count']}열 x {t['row_count']}행)" for t in tables]
                                    
                                    vm_select_key = f"vm_table_select_{encoded_template}"
                                    saved_idx = st.session_state.get(vm_select_key, 0)
                                    if saved_idx >= len(tables):
                                        saved_idx = 0
                                    
                                    selected_table_idx = st.selectbox(
                                        "편집할 표 선택",
                                        range(len(tables)),
                                        format_func=lambda i: table_options[i],
                                        key=vm_select_key,
                                        index=saved_idx
                                    )
                                    
                                    table = tables[selected_table_idx]
                                    headers = table.get('headers', [])
                                    rows = table.get('rows', [])
                                    
                                    if headers:
                                        st.caption(f"헤더: {' | '.join(headers)}")
                                    
                                    table_form_key = f"form_table_{table['slide_index']}_{table['shape_id']}_{refresh_ts}"
                                    with st.form(key=table_form_key, enter_to_submit=False):
                                        all_edited_rows = []
                                        delete_row_submitted = {}
                                        
                                        for row_idx, row_data in enumerate(rows):
                                            if row_idx == 0:
                                                all_edited_rows.append(row_data)
                                                continue
                                            
                                            cols = st.columns([0.5] + [2] * len(row_data) + [0.5])
                                            with cols[0]:
                                                st.write(f"{row_idx}")
                                            
                                            edited_row = []
                                            for col_idx, cell_text in enumerate(row_data):
                                                with cols[col_idx + 1]:
                                                    edited_text = st.text_input(
                                                        headers[col_idx] if col_idx < len(headers) else f"열{col_idx+1}",
                                                        value=cell_text,
                                                        key=f"table_{table['slide_index']}_{table['shape_id']}_{row_idx}_{col_idx}_{refresh_ts}",
                                                        label_visibility="collapsed"
                                                    )
                                                    edited_row.append(edited_text)
                                            all_edited_rows.append(edited_row)
                                            
                                            with cols[-1]:
                                                delete_row_submitted[row_idx] = st.form_submit_button(
                                                    "🗑️", 
                                                    key=f"del_{table['slide_index']}_{table['shape_id']}_{row_idx}_{refresh_ts}",
                                                    help=f"행 {row_idx} 삭제"
                                                )
                                        
                                        btn_cols = st.columns([2, 1, 1])
                                        with btn_cols[0]:
                                            save_all_submitted = st.form_submit_button(
                                                "💾 전체 저장", 
                                                key=f"save_{table['slide_index']}_{table['shape_id']}_{refresh_ts}",
                                                type="primary", 
                                                use_container_width=True
                                            )
                                        with btn_cols[1]:
                                            row_count = len(rows)
                                            insert_positions = ["맨 아래"] + [f"{i}번 행 위에" for i in range(1, row_count)]
                                            insert_pos_key = f"insert_pos_{table['slide_index']}_{table['shape_id']}_{refresh_ts}"
                                            insert_pos = st.selectbox(
                                                "삽입 위치",
                                                insert_positions,
                                                key=insert_pos_key,
                                                label_visibility="collapsed"
                                            )
                                        with btn_cols[2]:
                                            add_row_submitted = st.form_submit_button(
                                                "➕ 행 추가", 
                                                key=f"add_{table['slide_index']}_{table['shape_id']}_{refresh_ts}",
                                                use_container_width=True
                                            )
                                        
                                        deleted_row_idx = None
                                        for r_idx, was_clicked in delete_row_submitted.items():
                                            if was_clicked:
                                                deleted_row_idx = r_idx
                                                break
                                        
                                        if deleted_row_idx is not None:
                                            with st.spinner("삭제 중..."):
                                                try:
                                                    resp = requests.delete(
                                                        f"{API_URL}/api/templates/{encoded_template}/tables/{table['slide_index']}/{table['shape_id']}/rows/{deleted_row_idx}?type={template_type_key}",
                                                        timeout=10
                                                    )
                                                    if resp.ok:
                                                        st.success(f"행 {deleted_row_idx} 삭제 완료")
                                                        fetch_tables_cached.clear()
                                                except Exception as e:
                                                    st.error(str(e))
                                            st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                            st.rerun()
                                        elif save_all_submitted:
                                            if len(all_edited_rows) > 0:
                                                with st.spinner("저장 중..."):
                                                    try:
                                                        resp = requests.put(
                                                            f"{API_URL}/api/templates/{encoded_template}/tables/{table['slide_index']}/{table['shape_id']}?type={template_type_key}",
                                                            json={"rows": all_edited_rows},
                                                            timeout=30
                                                        )
                                                        if resp.ok:
                                                            st.success("저장 완료!")
                                                            fetch_tables_cached.clear()
                                                            st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                                            st.rerun()
                                                        else:
                                                            st.error("저장 실패")
                                                    except Exception as e:
                                                        st.error(f"저장 오류: {str(e)}")
                                            else:
                                                st.warning("저장할 데이터가 없습니다.")
                                        elif add_row_submitted:
                                            col_count = len(headers) if headers else table.get('col_count', 0)
                                            if col_count > 0:
                                                with st.spinner("행 추가 중..."):
                                                    try:
                                                        empty_row = [""] * col_count
                                                        insert_at = None
                                                        if insert_pos != "맨 아래":
                                                            insert_at = int(insert_pos.replace("번 행 위에", ""))
                                                        
                                                        payload = {"row_data": empty_row}
                                                        if insert_at is not None:
                                                            payload["insert_at"] = insert_at
                                                        
                                                        resp = requests.post(
                                                            f"{API_URL}/api/templates/{encoded_template}/tables/{table['slide_index']}/{table['shape_id']}/rows?type={template_type_key}",
                                                            json=payload,
                                                            timeout=10
                                                        )
                                                        if resp.ok:
                                                            pos_msg = f"{insert_at}번 행 위에" if insert_at else "맨 아래에"
                                                            st.success(f"빈 행 추가 완료 ({pos_msg})")
                                                            fetch_tables_cached.clear()
                                                            st.session_state['editor_refresh_ts'] = str(int(time.time() * 1000))
                                                            st.rerun()
                                                        else:
                                                            st.error("추가 실패")
                                                    except Exception as e:
                                                        st.error(str(e))
                                            else:
                                                st.warning("열 정보를 확인할 수 없어 행을 추가할 수 없습니다.")
                                    
                                    st.divider()
                                    if st.button("🔄 다른 템플릿 선택", key="reset_template_from_vm_editor", help="템플릿 선택 초기화"):
                                        if 'template_list_select' in st.session_state:
                                            del st.session_state['template_list_select']
                                        st.rerun()
                            except Exception as e:
                                st.error(str(e))
            else:
                st.error(f"템플릿 목록을 불러올 수 없습니다. (상태 코드: {all_templates_resp.status_code})")
        except Exception as e:
            st.error(f"템플릿 목록 로드 실패: {str(e)}")

with tab7:
    st.header("📥 다운로드")
    
    st.subheader("📦 템플릿 다운로드")
    st.info("원본, 이미지 삽입 완료, 통계 삽입 완료 템플릿 중 선택하여 다운로드할 수 있습니다.")
    
    dl_col1, dl_col2 = st.columns(2)
    
    with dl_col1:
        dl_type_options = {
            "source": "📁 원본 템플릿",
            "with_images": "🖼️ 이미지 삽입 완료",
            "final": "✅ 통계 삽입 완료 (최종)"
        }
        dl_type_keys = list(dl_type_options.keys())
        
        def format_dl_type(x):
            return dl_type_options.get(x, x)
        
        if 'download_type_select' in st.session_state:
            if st.session_state['download_type_select'] not in dl_type_keys:
                del st.session_state['download_type_select']
        
        download_type = st.selectbox(
            "템플릿 유형 선택",
            options=dl_type_keys,
            format_func=format_dl_type,
            key="download_type_select"
        )
    
    with dl_col2:
        dl_customer_options = ["전체"] + [c['name'] for c in fetch_customers_cached(API_URL)]
        
        if 'download_customer_select' in st.session_state:
            if st.session_state['download_customer_select'] not in dl_customer_options:
                del st.session_state['download_customer_select']
        
        download_customer = st.selectbox(
            "고객사 선택",
            options=dl_customer_options,
            key="download_customer_select"
        )
    
    if st.button("📥 다운로드 요청", type="primary"):
        try:
            params = {'type': download_type}
            if download_customer != "전체":
                params['customer'] = download_customer
            
            response = requests.get(f"{API_URL}/api/download/templates", params=params, timeout=60)
            
            if response.ok:
                type_label = dl_type_options.get(download_type, download_type).replace("📁 ", "").replace("🖼️ ", "").replace("✅ ", "")
                customer_label = download_customer if download_customer != "전체" else "all"
                filename = f"{customer_label}_{download_type}_templates.zip"
                
                st.download_button(
                    label="💾 ZIP 파일 저장",
                    data=response.content,
                    file_name=filename,
                    mime="application/zip"
                )
                st.success(f"✅ 다운로드 준비 완료: {type_label} ({download_customer})")
            else:
                error_msg = "다운로드 실패"
                try:
                    error_msg = response.json().get('error', error_msg)
                except (ValueError, KeyError):
                    pass
                st.error(f"❌ {error_msg}")
        except Exception as e:
            st.error(f"❌ 다운로드 실패: {str(e)}")

with tab8:
    st.header("⚙️ 설정")
    
    st.subheader("🔗 백엔드 API 연결")
    st.code(API_URL, language=None)
    st.info("환경 변수 BACKEND_API_URL을 설정하여 변경할 수 있습니다.")
    
    if st.button("🔄 연결 테스트"):
        try:
            response = requests.get(f"{API_URL}/health", timeout=5)
            if response.ok:
                st.success("✅ 연결 성공!")
                st.json(response.json())
            else:
                st.error(f"❌ 연결 실패: {response.status_code}")
        except Exception as e:
            st.error(f"❌ 연결 실패: {str(e)}")
    
    st.divider()
    
    st.subheader("📋 대시보드 매핑 관리")
    
    try:
        response = requests.get(f"{API_URL}/api/dashboard-mapping", timeout=5)
        if response.ok:
            mapping = response.json().get('mapping', {})
            
            st.write("**현재 대시보드 매핑:**")
            
            if mapping:
                edited_mapping = {}
                for customer, uid in mapping.items():
                    col1, col2 = st.columns([1, 2])
                    with col1:
                        st.write(f"**{customer}**")
                    with col2:
                        edited_mapping[customer] = st.text_input(
                            "UID",
                            value=uid,
                            key=f"mapping_{customer}",
                            label_visibility="collapsed"
                        )
                
                if st.button("💾 매핑 저장", type="primary"):
                    try:
                        resp = requests.put(
                            f"{API_URL}/api/dashboard-mapping",
                            json={"mapping": edited_mapping},
                            timeout=10
                        )
                        if resp.ok:
                            st.success("대시보드 매핑 저장 완료")
                            st.rerun()
                        else:
                            st.error("저장 실패")
                    except Exception as e:
                        st.error(str(e))
            else:
                st.info("등록된 대시보드 매핑이 없습니다.")
            
            st.divider()
            st.write("**JSON 직접 편집:**")
            mapping_json = st.text_area(
                "대시보드 매핑 JSON",
                value=json.dumps(mapping, ensure_ascii=False, indent=2),
                height=200,
                key="mapping_json_editor"
            )
            
            if st.button("💾 JSON으로 저장"):
                try:
                    new_mapping = json.loads(mapping_json)
                    resp = requests.put(
                        f"{API_URL}/api/dashboard-mapping",
                        json={"mapping": new_mapping},
                        timeout=10
                    )
                    if resp.ok:
                        st.success("대시보드 매핑 저장 완료")
                        st.rerun()
                    else:
                        st.error("저장 실패")
                except json.JSONDecodeError as e:
                    st.error(f"JSON 형식 오류: {str(e)}")
                except Exception as e:
                    st.error(str(e))
        else:
            st.error("대시보드 매핑을 불러올 수 없습니다.")
    except Exception as e:
        st.error(f"오류 발생: {str(e)}")
    
    st.divider()
    
    st.subheader("📋 백엔드 설정 정보")
    try:
        response = requests.get(f"{API_URL}/api/config", timeout=5)
        if response.ok:
            config_data = response.json()
            st.json(config_data)
        else:
            st.error("설정 정보를 가져올 수 없습니다.")
    except Exception as e:
        st.error(f"설정 조회 실패: {str(e)}")
