# 유작(遺作) 공연 예매 현황

네이버폼에서 내려받은 엑셀을 관리자 화면에 업로드하면, 관객이 첫 화면에서 **지인 이름**을 입력해 본인과 연결된 예매 내역만 확인하는 모바일 최적화 웹사이트입니다.

## 공개 화면
- 지인 이름 검색
- 예매자 성함은 마스킹하지 않음
- 연락처/응답일시/참여자/예매경로는 공개 API와 화면에 포함하지 않음
- 날짜 → 시간순 자동 정렬
- 같은 날짜 + 같은 시간 묶음
- 전체 총 예매 건수 / 총 예매 인원
- 날짜·시간별 건수 / 인원
- 마지막 엑셀 업로드 일시(KST)

## 관리자
`/#admin`에서 관리자 비밀번호와 `.xlsx/.xls` 파일을 입력하면 기존 데이터를 새 엑셀 데이터로 교체합니다.

## 로컬 실행
Node.js 18 이상 필요.

```bash
npm install
```

Windows PowerShell:
```powershell
$env:ADMIN_PASSWORD="원하는강한비밀번호"
$env:NODE_ENV="development"
npm start
```

브라우저: http://localhost:3000

## Render 배포
`render.yaml`은 **1GB Persistent Disk가 붙은 Render Starter Web Service** 기준입니다. Render의 Persistent Disk는 유료 웹 서비스에서 사용할 수 있고, 디스크가 없으면 로컬 파일 변경사항이 재시작/재배포 때 사라집니다.

1. 이 프로젝트를 GitHub에 올립니다.
2. Render에서 New → Web Service → GitHub 저장소를 연결합니다.
3. `render.yaml` Blueprint로 생성하거나 다음 값을 사용합니다.
   - Build: `npm install`
   - Start: `npm start`
   - Health Check: `/health`
   - `ADMIN_PASSWORD`: 원하는 강한 비밀번호
   - `DATA_DIR`: `/var/data`
   - Persistent Disk mount: `/var/data`
4. 배포 후 `https://서비스이름.onrender.com` 주소를 사용합니다.

### 중요
공개 링크를 실제 운영하기 전, 관리자 비밀번호를 반드시 설정하세요. `ADMIN_PASSWORD`가 없으면 production 모드에서는 서버가 시작되지 않습니다.
