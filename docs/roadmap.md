# 향후 개선 로드맵

> 마지막 갱신: 2026-03-16

---

## 현재 완료

| 항목 | 완료일 | 문서 |
|------|--------|------|
| 4단계 파이프라인 아키텍처 | 2026-03-03 | architecture.md |
| Phase 1: 전력판매 정확화 (SMP→MP, 열제약매출, 용량요금) | 2026-03-03 | revenue.md |
| 시간별 수요 데이터 전면 활용 | 2026-03-14 | engine.md |
| 입력 데이터 분리 (sales_hourly + temperature_hourly) | 2026-03-14 | architecture.md |
| PLB 수요 피크 기반 배치 | 2026-03-14 | engine.md |
| 기동비용 경제성 (DP + 휴리스틱) | 2026-03-15 | engine.md, logic-changelog.md |
| 야간 단시간 정지 합리화 | 2026-03-15 | logic-changelog.md |

---

## Phase 2: 로직 최적화 (진행 중)

| 과제 | 설명 | 우선순위 | 상태 |
|------|------|----------|------|
| 전환기 로직 | 봄/가을 수요 급감기 부하 조절 정밀화 | 중간 | 대기 |
| 다부하 조합 최적화 | 하루 중 부하를 시간대별로 달리하는 전략 | 낮음 | 대기 |

---

## Phase 3: 가스비 정확화

| 과제 | 설명 |
|------|------|
| NG 단가 3종 구분 | 발전용 / 집단용 / 연료전지용 분리 |
| 안산도시개발 단가 | NG 3종 구분으로 산출 |
| 인테코 단가 | NG 3종 구분으로 산출 |
| CHP/PLB 별도 집계 검증 | 현재 로직 정합성 확인 |

---

## Phase 4: 전체 손익

| 과제 | 설명 |
|------|------|
| 열판매 완성 | 기본료 + 지역사용 + 인테코 (현재 일부 구현) |
| 전력비 | 시간별 TOU 요금 |
| 용수비/약품비 | 보일러 변동비용 반영 |
| 수열비 | 외부수열 × 구매단가 |
| 공헌이익 완성 | 매출 - 재료비 전체 |

---

## Phase 5: 기능 확장

| 과제 | 설명 | 비고 |
|------|------|------|
| 연료전지 열원 추가 | CHP/PLB에 더해 연료전지를 열원으로 포함 | 새로운 열원 경쟁 |
| 외부수열량 변수화 | 현재 고정값 → 단가 포함 변수로 변경 | 자체열원과 단가 경쟁 |
| 용수비/약품비/수전비 | 보일러 가동에 따른 변동비 반영 | |
| 기동비용 포함 경제성 | ~~기동 비용까지 포함한 생산단가 경쟁~~ → Phase 2에서 완료 | 완료 |
| 급전매출 (데이터 연동) | 급전 데이터 있으면 QPC×MW²+LPC×MW+NLPC 적용 | Phase 4 이후 |

---

## Phase 6: 다년도 확장

| 과제 | 설명 |
|------|------|
| 1년 → 중장기 | 수요 성장률, 설비 노후화, 신규 열원 투입 등 반영 |
| 시나리오 분석 | 연료가격 변동, 수요 변화 등 다양한 시나리오 비교 |

---

## 배포 계획

| 단계 | 방법 | 상태 |
|------|------|------|
| 개발 | 로컬 HTML+JS + pywebview 데스크탑 앱 | 진행 중 |
| 배포 A | PyInstaller exe → 사내 배포 (오프라인) | 앱 완성 후 |
| 배포 B | SharePoint 호스팅 (E3) → Teams 탭 / URL 접속 | 앱 완성 후 |
| 보안 | Azure AD + SharePoint 권한 | 자동 |
| 모바일 | PWA 설정 추가 (선택) | |

---

## 보안 및 코드 보호 계획

### 1. 배포 경로별 보안

| 배포 방식 | 코드 보호 수준 | 설명 |
|-----------|---------------|------|
| **PyInstaller exe** | ★★★ 높음 | JS/HTML이 exe 내부 번들에 포함 → 일반 사용자는 접근 불가 |
| **SharePoint 웹** | ★★ 중간 | 브라우저 DevTools로 소스 볼 수 있음 → 난독화 필요 |
| **Teams 탭** | ★★ 중간 | SharePoint 호스팅과 동일, iframe으로 로드 |

### 2. 코드 난독화 (웹 배포 시)

| 방법 | 도구 | 효과 |
|------|------|------|
| JS 난독화 | `javascript-obfuscator` | 변수명·로직 난독화, 문자열 암호화 |
| JS 압축 | `terser` / `uglify-js` | 공백·주석 제거, 변수명 축약 |
| CSS 압축 | `cssnano` | 가독성 제거 |
| HTML 압축 | `html-minifier` | 인라인 최소화 |
| Source Map 제거 | 빌드 시 `.map` 생성 안 함 | DevTools에서 원본 추적 불가 |

**적용 명령 예시 (빌드 스크립트):**
```bash
# JS 난독화 + 압축
npx javascript-obfuscator js/app.js --output dist/js/app.js \
  --compact true --string-array true --string-array-encoding rc4
# CSS 압축
npx cssnano css/style.css dist/css/style.css
# HTML 압축
npx html-minifier index.html -o dist/index.html --collapse-whitespace --remove-comments
```

### 3. PyInstaller exe 보안 강화

| 항목 | 방법 |
|------|------|
| 리소스 번들링 | `--add-data` 로 html/js/css를 exe 내부에 포함 (이미 적용) |
| 임시폴더 보호 | `sys._MEIPASS` 경로는 실행 종료 시 자동 삭제 |
| PyArmor | Python 코드 자체를 암호화 (선택, 유료) |
| 코드 서명 | Windows Authenticode 서명 → 변조 감지 |

### 4. Teams 배포 상세

| 단계 | 작업 |
|------|------|
| 1. 빌드 | JS/CSS 난독화 → `dist/` 폴더 생성 |
| 2. 업로드 | SharePoint 사이트 > 문서 라이브러리에 `dist/` 내용 업로드 |
| 3. Teams 앱 등록 | Teams 관리센터 > 앱 > "웹사이트 탭" 추가 → SharePoint URL 지정 |
| 4. 권한 설정 | SharePoint 사이트 권한으로 접근 제어 (부서/팀 단위) |
| 5. SSO | Azure AD 로그인이 자동 적용 (E3 라이선스) |

### 5. 데이터 보안

| 항목 | 현재 | 개선 방안 |
|------|------|-----------|
| 프로젝트 데이터 | 로컬 `data/` 폴더에 CSV 저장 | exe: 앱 종료 시 암호화 저장 (AES) |
| 웹 배포 시 | - | SharePoint 문서 라이브러리에 저장 (Azure AD 권한) |
| 엑셀 내보내기 | 평문 xlsx | 비밀번호 보호 옵션 추가 (openpyxl 지원) |

### 6. 우선순위

| 순서 | 항목 | 시점 |
|------|------|------|
| ① | JS/CSS 난독화 빌드 스크립트 | 배포 직전 |
| ② | PyInstaller exe 빌드 + 코드 서명 | 배포 직전 |
| ③ | SharePoint 업로드 + Teams 탭 등록 | 배포 시 |
| ④ | 데이터 암호화 (선택) | 보안 요구 시 |
| ⑤ | PyArmor Python 암호화 (선택) | 고보안 요구 시 |
