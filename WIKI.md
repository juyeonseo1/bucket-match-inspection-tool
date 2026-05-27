# 버킷플레이스 매칭 검수 툴 — 기능 위키

> **배포 URL**: Google Apps Script Web App  
> **ScriptId**: `1yL-OZRc0BuTjYvBC0N3yk94Px3qT2XL_AHag7iS9Rn7qn4MqhGiG-0DD`  
> **DeploymentId**: `AKfycbx3LuSPl0mUP63ZwD4wRX-Ueq49glzCenZ03imRV1Nww70wL8w1XhU6-l0SCBfn4n8j`  
> **구성 파일**: `Code.gs` (백엔드) + `index.html` (프론트엔드)  
> **배포 방법**: `npx clasp push --force && npx clasp deploy --deploymentId <id>`

---

## 기능 목록

### 1. 자유 모드 데이터 업로드
- 엑셀/CSV 파일을 업로드해 Data 탭에 저장
- **추가 모드**: 기존 데이터 뒤에 이어 붙이기
- **덮어쓰기 모드**: 기존 데이터 전체 교체
- 파일 헤더 자동 감지 (`detectFileHeaderRowIdx`):
  - 1행 헤더 파일 (singleHeader) vs 2행 헤더 파일 (그룹헤더+컬럼명) 자동 구분
  - 스키마 컬럼과 fuzzy match 점수 비교로 판별

### 2. 컬럼 매핑 모달
- 업로드 파일의 컬럼명이 기존 스키마와 다를 때 자동으로 팝업
- 파일 컬럼(직접 선택 가능) vs 스키마 컬럼 1:1 매핑
- `__skip__` 선택 시 해당 컬럼 비워둠

### 3. 📊 통계 시트 업데이트
- 프로젝트 Google Sheet의 **'통계'** 탭과 **'처리주체_일별'** 탭을 현재 데이터로 업데이트
- 다운로드 없이 브라우저에서 바로 확인 가능
- **'통계' 탭 내용**:
  - 전체 진행 현황 (전체/완료/일치/불일치/정확도)
  - 작업자별 성과 (처리 건수, 일치율, 시간당 처리량)
  - 처리주체별 정확도 (human / system)
- **'처리주체_일별' 탭 내용**:
  - 단일 헤더: `날짜 | 처리주체 | 전체 | 검수완료 | 일치 | 불일치 | 정확도 (%)`
  - Collected date 기준 날짜별 집계
  - 다른 시트에서 IMPORTRANGE로 땡겨 쓸 수 있는 깔끔한 flat 테이블

### 4. 🔧 Data 시트 스키마 변경 (singleHeader 전환)
- Data 탭의 컬럼 구조를 변경하는 기능
- **기존 포맷** (free mode): 1행=그룹헤더(병합), 2행=컬럼명, 3행~=데이터
- **새 포맷** (singleHeader): 1행=컬럼명, 2행~=데이터 (병합 없음)
- 기존 데이터는 같은 이름의 컬럼으로 자동 이전
- 사용 방법:
  1. 관리자 패널 → 프로젝트 선택 → 맨 아래 스크롤
  2. **🔧 Data 시트 스키마 변경** 버튼 클릭
  3. 컬럼 목록 확인 (기본값 11개 pre-fill)
  4. **왼쪽 카드 컬럼 수** 지정 (기본값 8 = 오늘의집 정보)
  5. **변경 적용** 클릭

#### 현재 운영 컬럼 (11개)
| # | 컬럼명 | 카드 |
|---|--------|------|
| 1 | 우선순위 | 왼쪽 (오늘의집) |
| 2 | 처리주체 | 왼쪽 |
| 3 | 고유키 id | 왼쪽 |
| 4 | 상품명 | 왼쪽 |
| 5 | 오늘의집_url | 왼쪽 |
| 6 | 브랜드명 | 왼쪽 |
| 7 | 오늘의집_옵션1 | 왼쪽 |
| 8 | 오늘의집_옵션2 | 왼쪽 |
| 9 | 매칭 셀러 | 오른쪽 (매칭) |
| 10 | 매칭 상품명 | 오른쪽 |
| 11 | 매칭 url | 오른쪽 |

---

## 데이터 구조

### Projects 시트 (관리 스프레드시트)
| 열 | 내용 |
|----|------|
| A (col 1) | projectId |
| B (col 2) | 프로젝트명 |
| C (col 3) | 생성일 |
| D (col 4) | spreadsheetId (프로젝트 전용 시트) |
| E (col 5) | mode (`free` / `standard`) |
| F (col 6) | schemaJson |

#### schemaJson 포맷 (singleHeader)
```json
{
  "mode": "free",
  "singleHeader": true,
  "columns": ["우선순위", "처리주체", "고유키 id", "상품명", "오늘의집_url", "브랜드명", "오늘의집_옵션1", "오늘의집_옵션2", "매칭 셀러", "매칭 상품명", "매칭 url"],
  "leftCount": 8
}
```

#### schemaJson 포맷 (기존 free mode)
```json
{
  "mode": "free",
  "leftGroup": { "name": "오늘의집", "columns": [...] },
  "rightGroup": { "name": "매칭상품", "columns": [...] }
}
```

### 프로젝트 전용 스프레드시트
- **Data 탭**: 검수 대상 데이터
- **Results 탭**: 검수 결과 (`키 | 식별값1 | 식별값2 | 검수결과 | 오매핑유형 | 메모 | 작업자 | 작업시간`)
- **Assignments 탭**: 작업자별 행 범위 할당
- **Logs 탭**: 작업 이력
- **Reasons 탭**: 오매핑 사유 목록
- **통계 탭**: 📊 버튼으로 업데이트
- **처리주체_일별 탭**: IMPORTRANGE용 flat 테이블

---

## 주요 백엔드 함수

| 함수 | 설명 |
|------|------|
| `uploadData(payload)` | 파일 데이터 업로드 (추가/덮어쓰기) |
| `getProjectState(payload)` | 프로젝트 전체 상태 반환 |
| `saveResult(payload)` | 검수 결과 저장 |
| `updateStatsSheet(payload)` | 통계/처리주체_일별 시트 업데이트 |
| `updateDataSchema(payload)` | Data 시트 스키마 변경 (singleHeader 전환) |
| `deleteUnworkedRows(payload)` | 미작업 행 삭제 + 키 재정렬 |
| `getWorkerStats(payload)` | 작업자별 성과 통계 |
| `isSingleHeader_(proj)` | singleHeader 모드 여부 판별 헬퍼 |

---

## 개발 규칙

- `.clasp.json`, `appsscript.json` → **git에 올리지 않음**
- `Code.gs`, `index.html`만 커밋
- 배포는 항상 기존 DeploymentId로 (`--deploymentId` 플래그 사용)
