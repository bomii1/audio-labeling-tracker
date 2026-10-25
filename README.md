# Audio Labeling Tracker — MVP 0.1

React + TypeScript + Vite + Dexie(IndexedDB) 기반 로컬 우선 Label Studio JSON 작업량 분석 앱입니다.

## 실행

```bash
npm install
npm run dev
```

## 이번 버전에 포함

- 프로젝트 추가 / 프로젝트 목록 / 프로젝트별 대시보드
- 전체 프로젝트 JSON 업로드 및 프로젝트 ID 검증
- Draft Task ID 검출 및 정식 분석 차단
- 다중 Annotation 수동 선택, 이전 선택 자동 재적용, 새 Annotation 발견 시 확인
- 최초 제출자(completed_by) 귀속
- 프로젝트별 스냅샷 저장, 날짜별 최신 기록과 비교, 기록 조회

## 다음 단계

- Draft 목록 XLSX 및 상세 분석 XLSX 다운로드
- 동일 파일 해시 중복 감지 및 업로드 버전 UI
- 선택된 Annotation 이력 스냅샷 보존 및 백업/복원
- 파일 검증 강화 및 자동화 테스트

**주의:** 현재 버전은 학습용 초기 구현입니다. IndexedDB는 현재 브라우저에만 저장됩니다. 실제 업무 정산용으로 사용하기 전 추가 검증이 필요합니다.
