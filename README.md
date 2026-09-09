# 유작 예매 현황 — 무료 Render 버전

관리자 페이지나 업로드 기능 없이 `data/reservations.xlsx`를 서버가 읽어 표시합니다.

## Render 설정
- Build Command: `npm install`
- Start Command: `npm start`
- Root Directory: 비워둠
- Plan: Free

## 데이터 업데이트
새 네이버폼 엑셀을 내려받아 GitHub의 `data/reservations.xlsx`를 교체하고 Commit하면 Render가 자동 재배포됩니다.

## 집계 제외
`김맹훈`, `정해진` 예매자 행은 전체 집계 및 배우별 검색 결과에서 제외됩니다.

## 마지막 업데이트
무료/정적 파일 구조에서는 실제 GitHub 업로드 시각을 서버가 알 수 없으므로 UI에는 엑셀의 가장 최근 `응답일시`를 `엑셀 마지막 응답`으로 표시하고, 보조적으로 파일의 배포 시각을 제공합니다.
