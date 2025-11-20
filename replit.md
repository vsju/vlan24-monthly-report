# PowerPoint Automation Tool

## Overview
This project automates the generation of PowerPoint reports by:
1. Inserting images into PowerPoint templates based on shape names
2. Fetching data from Grafana dashboards and inserting statistics into presentations

**Architecture:** Client-Server separation
- **4C4M Server**: Streamlit UI (frontend only)
- **Zabbix Server**: Flask API + All processing (backend)

## Project Structure
```
.
├── client_app.py                # Streamlit UI (4C4M 서버용)
├── backend_api/                 # Flask API (Zabbix 서버용)
│   ├── app.py                   # Flask API 메인
│   ├── config.py                # 설정
│   ├── image_processor.py       # 이미지 삽입 로직
│   ├── stats_processor.py       # 통계 삽입 로직
│   ├── requirements.txt         # Backend 의존성
│   ├── README.md                # Backend 가이드
│   └── Report/                  # 파일 저장 위치 (자동 생성)
│       ├── template/            # 템플릿 저장
│       ├── [customer]/          # 고객사별 이미지
│       ├── completed_with_images/  # Step 1 출력
│       └── completed_final/     # Step 2 출력
├── DEPLOYMENT.md                # 배포 가이드
├── app.py                       # 구버전 (사용 안 함)
├── insert_images.py             # 구버전 (사용 안 함)
├── numinsert3.py                # 구버전 (사용 안 함)
└── replit.md                    # This file
```

## Deployment

전체 배포 가이드는 `DEPLOYMENT.md` 파일을 참조하세요.

### 1. Zabbix 서버 (Backend API)

```bash
cd backend_api
pip install -r requirements.txt

export GRAFANA_URL="http://zabbix.vlan24.co.kr:3000"
export GRAFANA_API_KEY="your_api_key"
export GRAFANA_VERIFY_SSL="false"

python app.py
```

Flask API가 `http://0.0.0.0:5001`에서 실행됩니다.

### 2. 4C4M 서버 (Frontend UI)

```bash
pip install streamlit requests

export BACKEND_API_URL="http://<Zabbix_Server_IP>:5001"

streamlit run client_app.py
```

Streamlit UI가 `http://localhost:8501`에서 실행됩니다.

### 3. Customer Dashboard Mapping

`backend_api/config.py`의 `DASHBOARD_MAP`에서 고객사와 대시보드 UID 매핑을 설정합니다.

## How to Use

### Web GUI (Streamlit)

1. **템플릿 업로드**
   - 이미지 삽입 탭에서 PowerPoint 템플릿 업로드
   - 고객사 이름 지정 (선택사항)

2. **이미지 업로드**
   - 고객사별 이미지 파일 업로드

3. **이미지 삽입 실행**
   - 특정 고객사 또는 전체 고객사 처리

4. **통계 삽입 실행**
   - Grafana 데이터 조회 및 플레이스홀더 치환

5. **결과 다운로드**
   - 완성된 보고서 ZIP 파일 다운로드

## Workflow

### Step 1: Image Insertion
1. Places PowerPoint templates in `Report/template/`
2. Organizes images in customer-specific folders
3. Script matches shape names to image filenames
4. Outputs to `Report/completed_with_images/`

### Step 2: Statistics Insertion
1. Takes files from `Report/completed_with_images/`
2. Queries Grafana dashboards for metrics
3. Replaces placeholders like `{{panel-name_A}}` with statistics
4. Replaces date placeholders ({{START_DATE}}, {{END_DATE}}, etc.)
5. Outputs final reports to `Report/completed_final/`

## Placeholder Format

### Date Placeholders
- `{{START_DATE}}` - Start date in Korean format
- `{{END_DATE}}` - End date in Korean format
- `{{MONTH}}` - Month number
- `{{DATE_RANGE}}` - Full date range in Korean format
- `{{DATE_RANGE_HYPHEN}}` - Full date range with hyphens

### Grafana Statistics Placeholders
Format: `{{panel-name_QueryLetter}}`
Example: `{{CPU-Usage_A}}`

The script will:
1. Find the panel with matching title
2. Query the specified query letter (A, B, C, etc.)
3. Calculate max and mean values
4. Replace with: "사용량 최대 X%, 평균 Y% 입니다."

## Configuration

### config.py
Key settings:
- `BASE_TEMPLATE_DIR` - Template directory
- `OUTPUT_DIR_WITH_IMAGES` - Intermediate output
- `OUTPUT_DIR` - Final output
- `GRAFANA_URL` - Grafana server URL
- `API_KEY` - Grafana API key
- `DASHBOARD_MAP` - Customer to dashboard UID mapping
- `SENTENCE_TEMPLATE` - Output format for statistics

## Dependencies

### Backend (Zabbix 서버)
- flask==3.0.0 - Web API framework
- flask-cors==4.0.0 - CORS support
- python-pptx==0.6.23 - PowerPoint manipulation
- python-dateutil==2.8.2 - Date calculations
- requests==2.31.0 - HTTP requests for Grafana API
- urllib3==2.1.0 - HTTP client

### Frontend (4C4M 서버)
- streamlit - Web GUI framework
- requests - HTTP client for API calls

## Troubleshooting

### No templates found
- Ensure .pptx files are in `Report/template/`
- Check file permissions
- Run option 4 to create directories

### Images not inserting
- Verify image filenames match shape names (case-insensitive, special characters ignored)
- Ensure images are in the correct customer folder
- Check image file extensions (.png, .jpg, .jpeg, .gif)

### Grafana queries failing
- Verify GRAFANA_URL is correct
- Check GRAFANA_API_KEY is valid
- Ensure dashboard UIDs in DASHBOARD_MAP are correct
- Verify panel titles match placeholder names

### Environment Variables
Set these in Replit Secrets if needed:
- `GRAFANA_URL` - Grafana server URL
- `GRAFANA_API_KEY` - Grafana API token
- `GRAFANA_VERIFY_SSL` - SSL verification (default: true, set to "false" only for testing)

## Architecture

### Client-Server Separation (옵션 A)

```
┌─────────────────────────┐      ┌──────────────────────────┐
│  4C4M 서버              │      │  Zabbix 서버             │
│  ─────────────────      │ HTTP │  ─────────────────       │
│  • Streamlit UI만       │ ───> │  • Flask API             │
│  • client_app.py        │      │  • 스크립트 실행          │
│                         │      │  • Grafana 연동          │
└─────────────────────────┘      │  • 파일 저장/처리         │
                                 └──────────────────────────┘
```

**특징:**
- 4C4M 서버: 가벼운 UI만 (사용자 접근 포인트)
- Zabbix 서버: 모든 비즈니스 로직 및 파일 처리
- 네트워크: 같은 PVE node 내부 통신
- DB: 현재 미사용 (추후 추가 가능)

## Recent Changes

- 2025-11-20: **아키텍처 재설계 - 클라이언트/서버 분리 (옵션 A)**
  - **Backend (Zabbix 서버)**:
    - Flask API 백엔드 구축 (backend_api/)
    - 핵심 엔드포인트만 유지: /api/process/images, /api/process/statistics, /api/process/all
    - 파일 업로드/다운로드 API 추가
    - 기존 스크립트 모듈화 (image_processor.py, stats_processor.py)
  - **Frontend (4C4M 서버)**:
    - Streamlit UI 단순화 (client_app.py)
    - DB 및 인증 완전 제거 (추후 추가 가능하도록 설계)
    - API 클라이언트로만 작동
  - **배포**:
    - 상세 배포 가이드 작성 (DEPLOYMENT.md)
    - Systemd 서비스 설정 예시 포함
    - 네트워크 및 방화벽 설정 가이드
  - **제거된 기능** (MVP 범위 외):
    - 자동 템플릿 생성
    - 지원 내역 추가
    - 권고사항 자동 생성
    - 사용자 인증/관리
    - 작업 이력 DB

- 2025-11-18: Grafana URL 정규화 - 이중 슬래시 404 에러 수정
  - config.py, numinsert3.py에서 URL 끝 슬래시 자동 제거
  - 탭 이름 수정 (사용자 관리/설정 순서)

- 2025-11-17: Initial Replit setup
  - 경로 설정 가능하도록 변경
  - CLI 및 Streamlit GUI 추가
  - 디렉토리 구조 설정
