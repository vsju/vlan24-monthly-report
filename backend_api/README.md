# PowerPoint Automation Backend API

Zabbix 서버에서 실행되는 Flask API 백엔드입니다. 모든 PowerPoint 처리 작업을 담당합니다.

## 주요 기능

- PowerPoint 템플릿에 이미지 자동 삽입
- Grafana 통계 데이터 조회 및 삽입
- 날짜 플레이스홀더 자동 치환
- 파일 업로드/다운로드
- 고객사별 처리 지원

## 설치 방법

자세한 설치 및 배포 가이드는 상위 디렉토리의 `DEPLOYMENT.md` 파일을 참조하세요.

### 빠른 시작

```bash
cd backend_api
pip install -r requirements.txt

export GRAFANA_URL="http://zabbix.vlan24.co.kr:3000"
export GRAFANA_API_KEY="your_api_key"
export GRAFANA_VERIFY_SSL="false"

python app.py
```

## 디렉토리 구조

```
backend_api/
├── Report/
│   ├── template/              # PowerPoint 템플릿
│   ├── [customer]/            # 고객사별 이미지 폴더
│   ├── completed_with_images/ # Step 1 출력 (이미지 삽입 완료)
│   └── completed_final/       # Step 2 출력 (최종 완성)
├── app.py                     # Flask API 메인
├── config.py                  # 설정
├── image_processor.py         # 이미지 삽입 로직
├── stats_processor.py         # 통계 삽입 로직
└── requirements.txt           # Python 의존성
```

## API 엔드포인트

### 시스템

- `GET /health` - 서버 상태 및 설정 확인
- `GET /api/config` - 현재 설정 조회
- `GET /api/customers` - 등록된 고객사 목록

### 처리

- `POST /api/process/images` - 이미지 삽입 실행
  - Body: `{"customer_name": "GIT"}` (선택사항)
  
- `POST /api/process/statistics` - 통계 삽입 실행
  - Body: `{"customer_name": "GIT"}` (선택사항)
  
- `POST /api/process/all` - 전체 프로세스 실행 (이미지 + 통계)

### 파일 관리

- `POST /api/upload/templates` - 템플릿 업로드
  - Form: `files`, `customer` (선택사항)
  
- `POST /api/upload/images` - 이미지 업로드
  - Form: `files`, `customer` (필수)
  
- `GET /api/download/results` - 결과 다운로드
  - Query: `customer` (선택사항)
  
- `GET /api/files/list` - 파일 목록 조회
  - Query: `type=results|templates`, `customer` (선택사항)

## 환경 변수

| 변수 | 설명 | 기본값 |
|------|------|--------|
| `GRAFANA_URL` | Grafana 서버 URL | `http://localhost:3000` |
| `GRAFANA_API_KEY` | Grafana API 키 | (없음) |
| `GRAFANA_VERIFY_SSL` | SSL 인증서 검증 | `true` |
| `PORT` | Flask 서버 포트 | `5001` |

## 에러 처리

모든 엔드포인트는 다음 형식으로 응답합니다:

**성공:**
```json
{
  "success": true,
  "processed_files": [...],
  "summary": {...}
}
```

**실패:**
```json
{
  "success": false,
  "errors": ["error message"],
  "traceback": "..."
}
```
