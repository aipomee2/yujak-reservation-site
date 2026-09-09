# 유작 예매 현황 사이트

## GitHub
프로젝트 폴더 안의 파일을 GitHub repository 최상단에 업로드합니다.

## Render
- Build Command: `npm install`
- Start Command: `npm start`
- Health Check Path: `/health`
- Environment Variable: `ADMIN_PASSWORD`
- `render.yaml`을 사용할 경우 DATA_DIR은 `/var/data`입니다.

## 관리자
배포 후 `/admin`이 아니라 메인 주소 뒤에 `#admin`을 붙여 관리자 화면을 엽니다.
예: `https://사이트주소.onrender.com/#admin`

네이버폼 엑셀 컬럼은 다음을 지원합니다.
`응답일시 / 참여자 / 예매자 성함 / 예매 인원 / 예매자 연락처 / 예매 일자 / 예매 경로 / 지인`

연락처는 화면에 전송하지 않습니다.
