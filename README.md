# 애슐리 인건비 모니터

주간 인건비 모니터링 대시보드. 데이터는 Supabase Storage의 JSON에서 실시간으로 불러옵니다.

## 매주 데이터 갱신
1. 새 주차 JSON을 Supabase Storage `snapshots` 버킷에 업로드
2. `components/LaborApp.jsx` 상단의 `DATA_URL`을 새 파일 주소로 변경 (또는 파일명을 동일하게 덮어쓰기)

## 테스트 통합: 스케줄 진단
상단 메뉴에 `스케줄 진단`이 추가되어 있습니다.
현재 단계에서는 기존 스케줄 진단 HTML을 `public/schedule-diagnosis.html`에 두고,
React 컴포넌트 `ScheduleDiagnosisFrame.jsx`가 같은 Vercel 프로젝트 안에서 iframe으로 로드합니다.

이 방식은 두 코드를 서로 건드리지 않고 공존시키는 1차 통합 테스트용입니다.
정상 동작 확인 후 스케줄 진단을 React 컴포넌트로 단계적으로 변환할 수 있습니다.
