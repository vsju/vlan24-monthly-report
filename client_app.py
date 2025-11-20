import streamlit as st
import requests
import os

API_URL = os.getenv("BACKEND_API_URL", "http://localhost:5001")

st.set_page_config(
    page_title="PowerPoint 자동화",
    page_icon="📊",
    layout="wide"
)

st.title("📊 PowerPoint 자동화 도구")
st.markdown("이미지와 Grafana 통계를 자동으로 PowerPoint에 삽입하는 도구입니다.")

tab1, tab2, tab3, tab4 = st.tabs([
    "🏠 홈",
    "🖼️ 이미지 삽입",
    "📈 통계 삽입",
    "⚙️ 설정"
])

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
    
    try:
        response = requests.get(f"{API_URL}/health", timeout=5)
        if response.ok:
            data = response.json()
            st.success("✅ 백엔드 서버 연결 성공")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Grafana URL", data.get("grafana_url", "N/A"))
            with col2:
                status = "✅ 설정됨" if data.get("grafana_configured") else "❌ 미설정"
                st.metric("Grafana API", status)
        else:
            st.error("❌ 백엔드 서버 연결 실패")
    except Exception as e:
        st.error(f"❌ 백엔드 서버에 연결할 수 없습니다: {str(e)}")
        st.info(f"백엔드 API URL: {API_URL}")

with tab2:
    st.header("🖼️ 이미지 삽입 (Step 1)")
    
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
    
    st.divider()
    
    st.subheader("📤 이미지 업로드")
    image_customer = st.text_input(
        "고객사 이름 (필수)",
        key="image_customer"
    )
    
    uploaded_images = st.file_uploader(
        "이미지 파일을 선택하세요",
        type=['png', 'jpg', 'jpeg', 'gif'],
        accept_multiple_files=True,
        key="image_uploader"
    )
    
    if uploaded_images and image_customer and st.button("📤 이미지 업로드", type="primary"):
        try:
            files = [('files', (f.name, f, f'image/{f.type.split("/")[1]}')) for f in uploaded_images]
            data = {'customer': image_customer}
            
            response = requests.post(f"{API_URL}/api/upload/images", files=files, data=data, timeout=30)
            
            if response.ok:
                result = response.json()
                if result.get('success'):
                    st.success(f"✅ {len(result.get('uploaded_files', []))}개 이미지 업로드 완료!")
                    if result.get('errors'):
                        st.warning("⚠️ 일부 이미지 업로드 실패:\n" + "\n".join(result['errors']))
                else:
                    st.error(f"❌ 업로드 실패: {result.get('error', '알 수 없는 오류')}")
            else:
                st.error(f"❌ 서버 오류: {response.status_code}")
        except Exception as e:
            st.error(f"❌ 업로드 실패: {str(e)}")
    
    st.divider()
    
    st.subheader("▶️ 이미지 삽입 실행")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        process_customer = st.text_input(
            "고객사 이름 (비워두면 전체 처리)",
            key="process_customer"
        )
    
    with col2:
        st.write("")
        st.write("")
        if st.button("🚀 이미지 삽입 실행", type="primary", use_container_width=True):
            with st.spinner("처리 중..."):
                try:
                    payload = {"customer_name": process_customer if process_customer else None}
                    response = requests.post(f"{API_URL}/api/process/images", json=payload, timeout=300)
                    
                    if response.ok:
                        result = response.json()
                        if result.get('success'):
                            st.success(f"✅ 이미지 삽입 완료! ({result['summary']['processed_count']}개 파일 처리)")
                            
                            with st.expander("처리 결과 상세"):
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

with tab3:
    st.header("📈 통계 삽입 (Step 2)")
    
    st.subheader("▶️ 통계 삽입 실행")
    st.info("💡 이미지 삽입이 완료된 파일에 Grafana 통계를 삽입합니다.")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        stats_customer = st.text_input(
            "고객사 이름 (비워두면 전체 처리)",
            key="stats_customer"
        )
    
    with col2:
        st.write("")
        st.write("")
        if st.button("🚀 통계 삽입 실행", type="primary", use_container_width=True):
            with st.spinner("처리 중... (Grafana 조회 시간이 소요될 수 있습니다)"):
                try:
                    payload = {"customer_name": stats_customer if stats_customer else None}
                    response = requests.post(f"{API_URL}/api/process/statistics", json=payload, timeout=600)
                    
                    if response.ok:
                        result = response.json()
                        if result.get('success'):
                            st.success(f"✅ 통계 삽입 완료! ({result['summary']['processed_count']}개 파일 처리)")
                            
                            with st.expander("처리 결과 상세"):
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
    
    st.divider()
    
    st.subheader("📥 결과 다운로드")
    
    download_customer = st.text_input(
        "고객사 이름 (비워두면 전체 다운로드)",
        key="download_customer"
    )
    
    if st.button("📥 결과 파일 다운로드"):
        try:
            params = {'customer': download_customer} if download_customer else {}
            response = requests.get(f"{API_URL}/api/download/results", params=params, timeout=60)
            
            if response.ok:
                filename = f"{download_customer}_results.zip" if download_customer else "all_results.zip"
                st.download_button(
                    label="💾 ZIP 파일 저장",
                    data=response.content,
                    file_name=filename,
                    mime="application/zip"
                )
            else:
                st.error(f"❌ 다운로드 실패: {response.status_code}")
        except Exception as e:
            st.error(f"❌ 다운로드 실패: {str(e)}")

with tab4:
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
