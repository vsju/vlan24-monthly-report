# OnlyOffice Document Server 배포 가이드

## 4C4M 서버 (192.168.10.77)에 배포

### 1. Docker 설치 (이미 설치되어 있다면 건너뛰기)

```bash
# Docker 설치
curl -fsSL https://get.docker.com -o get-docker.sh
sudo sh get-docker.sh

# Docker Compose 설치
sudo apt-get install docker-compose-plugin

# 현재 사용자를 docker 그룹에 추가
sudo usermod -aG docker $USER
newgrp docker
```

### 2. OnlyOffice 배포

```bash
# 디렉토리 생성
mkdir -p /home/vlan24/onlyoffice
cd /home/vlan24/onlyoffice

# docker-compose.yml 파일 복사 후 실행
docker compose up -d

# 상태 확인 (약 1-2분 소요)
docker compose logs -f
```

### 3. JWT Secret 설정

`docker-compose.yml`에서 `JWT_SECRET`을 안전한 값으로 변경하세요:

```yaml
environment:
  - JWT_SECRET=your-actual-secret-key-here
```

이 값은 Flask 백엔드의 `config.py`와 동일해야 합니다.

### 4. 방화벽 설정

```bash
# 8080 포트 열기
sudo ufw allow 8080/tcp
```

### 5. 동작 확인

브라우저에서 접속: `http://192.168.10.77:8080`

"Document Server is running" 메시지가 표시되면 성공입니다.

## 시스템 요구사항

- RAM: 최소 2GB (권장 4GB)
- 디스크: 10GB 이상
- CPU: 2코어 이상

## Systemd 서비스 등록 (선택)

Docker 컨테이너가 시스템 부팅 시 자동 시작되도록 설정:

```bash
# /etc/systemd/system/onlyoffice.service
[Unit]
Description=OnlyOffice Document Server
After=docker.service
Requires=docker.service

[Service]
Type=oneshot
RemainAfterExit=yes
WorkingDirectory=/home/vlan24/onlyoffice
ExecStart=/usr/bin/docker compose up -d
ExecStop=/usr/bin/docker compose down

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable onlyoffice
sudo systemctl start onlyoffice
```

## 문제 해결

### 컨테이너가 시작되지 않는 경우

```bash
# 로그 확인
docker compose logs onlyoffice

# 컨테이너 재시작
docker compose restart
```

### 메모리 부족

```bash
# 스왑 추가
sudo fallocate -l 2G /swapfile
sudo chmod 600 /swapfile
sudo mkswap /swapfile
sudo swapon /swapfile
```
