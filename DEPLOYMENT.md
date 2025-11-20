# PowerPoint 자동화 배포 가이드

## 아키텍처 개요

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

---

## 1. Zabbix 서버 설정 (백엔드)

### 1.1 사전 요구사항
- Python 3.10 이상
- Grafana API 접근 권한
- 충분한 디스크 공간 (보고서 파일 저장용)

### 1.2 디렉토리 구조 생성

```bash
# backend_api 폴더로 이동
cd /path/to/backend_api

# 필요한 디렉토리 생성
mkdir -p Report/template
mkdir -p Report/completed_with_images
mkdir -p Report/completed_final
```

### 1.3 Python 패키지 설치

```bash
pip install -r requirements.txt
```

**requirements.txt 내용:**
```
flask==3.0.0
flask-cors==4.0.0
python-pptx==0.6.23
python-dateutil==2.8.2
requests==2.31.0
urllib3==2.1.0
```

### 1.4 환경 변수 설정

Zabbix 서버에서 다음 환경 변수를 설정하세요:

```bash
export GRAFANA_URL="http://zabbix.vlan24.co.kr:3000"
export GRAFANA_API_KEY="your_grafana_api_key_here"
export GRAFANA_VERIFY_SSL="false"  # 자체 서명 인증서 사용 시
export PORT="5001"  # Flask API 포트 (기본값: 5001)
```

**영구 설정 (systemd 사용 시):**

`/etc/systemd/system/ppt-automation-backend.service` 파일 생성:

```ini
[Unit]
Description=PowerPoint Automation Backend API
After=network.target

[Service]
Type=simple
User=your_user
WorkingDirectory=/path/to/backend_api
Environment="GRAFANA_URL=http://zabbix.vlan24.co.kr:3000"
Environment="GRAFANA_API_KEY=your_api_key"
Environment="GRAFANA_VERIFY_SSL=false"
Environment="PORT=5001"
ExecStart=/usr/bin/python3 /path/to/backend_api/app.py
Restart=always

[Install]
WantedBy=multi-user.target
```

서비스 시작:

```bash
sudo systemctl daemon-reload
sudo systemctl enable ppt-automation-backend
sudo systemctl start ppt-automation-backend
sudo systemctl status ppt-automation-backend
```

### 1.5 수동 실행 (개발/테스트용)

```bash
cd /path/to/backend_api
python app.py
```

서버가 `http://0.0.0.0:5001`에서 실행됩니다.

### 1.6 방화벽 설정

4C4M 서버에서 접근할 수 있도록 포트 5001을 개방하세요:

```bash
sudo firewall-cmd --permanent --add-port=5001/tcp
sudo firewall-cmd --reload
```

### 1.7 동작 확인

```bash
curl http://localhost:5001/health
```

정상 응답 예시:
```json
{
  "status": "healthy",
  "grafana_url": "http://zabbix.vlan24.co.kr:3000",
  "grafana_configured": true,
  "directories": {
    "templates": true,
    "images": true,
    "output_images": true,
    "output_final": true
  }
}
```

---

## 2. 4C4M 서버 설정 (프론트엔드)

### 2.1 사전 요구사항
- Python 3.10 이상
- Zabbix 서버와 네트워크 연결

### 2.2 Python 패키지 설치

```bash
pip install streamlit requests
```

### 2.3 환경 변수 설정

```bash
export BACKEND_API_URL="http://<Zabbix_Server_IP>:5001"
```

예시:
```bash
export BACKEND_API_URL="http://192.168.1.100:5001"
```

**영구 설정 (systemd 사용 시):**

`/etc/systemd/system/ppt-automation-frontend.service` 파일 생성:

```ini
[Unit]
Description=PowerPoint Automation Frontend UI
After=network.target

[Service]
Type=simple
User=your_user
WorkingDirectory=/path/to/project
Environment="BACKEND_API_URL=http://192.168.1.100:5001"
ExecStart=/usr/bin/streamlit run client_app.py --server.port=8501 --server.address=0.0.0.0
Restart=always

[Install]
WantedBy=multi-user.target
```

서비스 시작:

```bash
sudo systemctl daemon-reload
sudo systemctl enable ppt-automation-frontend
sudo systemctl start ppt-automation-frontend
sudo systemctl status ppt-automation-frontend
```

### 2.4 수동 실행 (개발/테스트용)

```bash
export BACKEND_API_URL="http://192.168.1.100:5001"
streamlit run client_app.py
```

UI가 `http://localhost:8501`에서 실행됩니다.

### 2.5 외부 접근 설정

```bash
streamlit run client_app.py --server.port=8501 --server.address=0.0.0.0
```

### 2.6 방화벽 설정

```bash
sudo firewall-cmd --permanent --add-port=8501/tcp
sudo firewall-cmd --reload
```

### 2.7 동작 확인

브라우저에서 `http://<4C4M_Server_IP>:8501` 접속

---

## 3. 네트워크 설정

### 3.1 PVE Node 내부 통신

두 서버가 같은 PVE node에 있으므로 내부 IP로 통신합니다:

- 4C4M 서버 → Zabbix 서버: `http://<Zabbix_Internal_IP>:5001`

### 3.2 DNS 설정 (선택사항)

`/etc/hosts` 파일에 추가:

```
192.168.1.100    zabbix-server
192.168.1.101    frontend-server
```

그러면:
```bash
export BACKEND_API_URL="http://zabbix-server:5001"
```

---

## 4. 사용 방법

### 4.1 기본 워크플로우

1. **템플릿 업로드**
   - 이미지 삽입 탭에서 PowerPoint 템플릿 업로드
   - 고객사 이름 지정 (선택사항)

2. **이미지 업로드**
   - 고객사별 이미지 파일 업로드

3. **이미지 삽입 실행**
   - 특정 고객사 또는 전체 실행

4. **통계 삽입 실행**
   - Grafana 데이터 조회 및 삽입

5. **결과 다운로드**
   - 완성된 보고서 다운로드

---

## 5. 트러블슈팅

### 5.1 백엔드 연결 실패

**증상:** "백엔드 서버에 연결할 수 없습니다"

**해결:**
1. Zabbix 서버에서 Flask API 실행 확인:
   ```bash
   sudo systemctl status ppt-automation-backend
   ```

2. 네트워크 연결 확인:
   ```bash
   curl http://<Zabbix_Server_IP>:5001/health
   ```

3. 방화벽 확인:
   ```bash
   sudo firewall-cmd --list-all
   ```

### 5.2 Grafana API 오류

**증상:** "Grafana 대시보드 조회 실패"

**해결:**
1. Grafana URL 확인:
   ```bash
   curl http://zabbix.vlan24.co.kr:3000
   ```

2. API 키 유효성 확인:
   ```bash
   curl -H "Authorization: Bearer YOUR_API_KEY" \
        http://zabbix.vlan24.co.kr:3000/api/dashboards/uid/aejkgkoze5nggb
   ```

3. 환경 변수 재확인

### 5.3 파일 업로드 실패

**증상:** "업로드 실패"

**해결:**
1. 디렉토리 권한 확인:
   ```bash
   ls -la /path/to/backend_api/Report/
   ```

2. 디스크 공간 확인:
   ```bash
   df -h
   ```

---

## 6. 유지보수

### 6.1 로그 확인

**백엔드 로그:**
```bash
sudo journalctl -u ppt-automation-backend -f
```

**프론트엔드 로그:**
```bash
sudo journalctl -u ppt-automation-frontend -f
```

### 6.2 업데이트

```bash
# 코드 업데이트
cd /path/to/project
git pull  # (Git 사용 시)

# 서비스 재시작
sudo systemctl restart ppt-automation-backend
sudo systemctl restart ppt-automation-frontend
```

### 6.3 백업

```bash
# 보고서 파일 백업
tar -czf backup-$(date +%Y%m%d).tar.gz /path/to/backend_api/Report/
```

---

## 7. 보안 권장사항

1. **API 키 보호**: 환경 변수 사용, 파일에 직접 저장 금지
2. **HTTPS 사용**: 프로덕션 환경에서는 리버스 프록시(nginx) 사용
3. **방화벽**: 필요한 포트만 개방
4. **정기 업데이트**: Python 패키지 보안 업데이트

---

## 8. 문의

문제 발생 시:
1. 로그 확인
2. 네트워크 연결 테스트
3. 환경 변수 재확인
