# 시스템 아키텍처

> 시스템 개요, 파이프라인 구조, 입력 데이터, 설정 파라미터
> 마지막 갱신: 2026-03-15

---

## 1. 개요

### 목표
**연간 생산단가(원/Gcal) 최소화** — CHP와 PLB의 최적 운전 패턴을 결정한다.

### 기본 원칙
- CHP 우선 운전 (전기 판매수익이 있어 PLB보다 경제적)
- 축열조를 버퍼로 활용하여 피크 수요 대응
- PLB는 최후 수단 (고비용 → 최소 가동)
- 시뮬레이션과 차트가 **동일한 hourly 데이터**를 사용 (이중 계산 없음)

### 핵심 아키텍처: 4단계 파이프라인

```
Stage 1: 연간 최적화     → 일별 CHP 시간/부하 결정 (DP 또는 휴리스틱)
Stage 2: 시간별 배치     → CHP 연속블록 시작시각 결정 (MP 스코어링)
Stage 3: 시간별 시뮬레이션 → 24h 순차 실행 (축열조 잔량 추적, 시간별 수요 반영)
Stage 4: 집계           → hourly → daily → monthly
차트: results[].hourly 직접 표시 (별도 계산 없음)
```

---

## 2. 데이터 흐름 (파이프라인)

```
[CSV 데이터] → buildDailyData() → days[365]
                                      │
                      ┌────────────────┴────────────────┐
                      │                                 │
              runDailySimulation()               solveDP()
              (휴리스틱 2-Pass)              (DP 역방향 최적화)
                      │                                 │
                      │  ┌─ Stage 1: 일별 CHP h/부하    │
                      │  ├─ Stage 2: CHP 블록 배치      │
                      │  ├─ Stage 3: 24h 시뮬레이션     │
                      │  └─ Stage 4: 일별 집계          │
                      │                                 │
                      └────────────┬────────────────────┘
                                   │
                          results[365] 배열
                          ├─ 일별 요약 (chpHours, chpHeat, plbHeat, storageLevel...)
                          ├─ hourly[24] (hour, demand, ext, chp, plb, storage, smp, mp)
                          └─ smpH[24]
                                   │
                      ┌────────────┼────────────────┐
                      │            │                │
               월별 차트      일별 차트      일 상세 모달
            (aggregateMonthly) (renderDailyChart) (showDayDetailModal)
```

---

## 3. 입력 데이터

| 구분 | 데이터 | 키 | 출처 |
|------|--------|-----|------|
| 열수요(시간별) | 시간별 수요량 (Gcal) | `sales_hourly` | 주택난방/업무난방·냉방/공공난방·냉방 (8,760행) |
| 열수요(일별) | 인테코 판매 (Gcal) | `sales_daily` | 인테코판매만 일별 |
| 외부수열 | 일별 외부 공급량 (Gcal) | `operations_daily` | 남부소각장, ERG, SRF, 안산 |
| 인테코 | 수열/판매 (Gcal) | `operations_daily` / `sales_daily` | |
| SMP | 시간별 전력거래가격 (원/kWh) | `smp_hourly` | 24시간 × 365일 |
| 기온 | 시간별 기온 (℃) | `temperature_hourly` | 8,760행 |
| 공휴일 | 날짜별 공휴일 여부 | `holidays` | |
| NG단가 | 월별 CHP/PLB 열량단가 | `ng_price_monthly` | 원/MJ → 원/N㎥ 환산 |

---

## 4. 설정 파라미터

### CHP 설정
| 파라미터 | ID | 기본값 | 설명 |
|---------|-----|--------|------|
| 최소부하 | `chpMinLoad` | 60% | 이하 효율 급락, 운전 불가 |
| 최소운전시간 | `chpMinRunHours` | 8h | 한번 켜면 최소 운전 |
| 최소정지시간 | `chpMinStopHours` | 4h | 한번 끄면 최소 정지 |
| 부하 테이블 | `loadTable` | — | 부하별 열/전기/NG |

### PLB 설정
| 파라미터 | ID | 기본값 | 설명 |
|---------|-----|--------|------|
| 보유대수 | `plbCount` | 3 | |
| 최소운전시간 | `plbMinRunHours` | 2h | 한번 켜면 최소 운전 (정수 시간) |
| PLB 회복목표 | `plbRecoveryTarget` | 200 | (현재 storageMin 기준으로 변경됨) |

### 축열조 설정
| 파라미터 | ID | 기본값 | 설명 |
|---------|-----|--------|------|
| 용량 | `storageCapacity` | — | Gcal |
| 초기잔량 | `storageInitial` | — | Gcal |
| 최소잔량 | `storageMinLevel` | — | Gcal, 이하 시 PLB 투입 |

### 기동조건
| 분류 | 기동시간 ID | 가스사용량 ID | 기본값 | 적용 조건 |
|------|-----------|-------------|--------|----------|
| Hot | `hotTime` | `hotGasRate` | 4h / 300 Nm³/h | 정지 ≤8h (1일) |
| Warm | `warmTime` | `warmGasRate` | 6h / 350 Nm³/h | 정지 ≤48h (2일) |
| Cold | `coldTime` | `coldGasRate` | 8h / 400 Nm³/h | 정지 >48h (3일+) |

**기동비용** = 기동시간(h) × 가스사용량(Nm³/h) × 월별 NG단가(원/Nm³)

### 정비일정
- 최대 3건, 시작~종료 날짜

---

## 5. Stage 0: 기초 데이터 산출

**함수**: `buildDailyData()`

365일 배열 생성. 각 요소:

```javascript
{
    month, day, dow, isWeekend, isHoliday, isMaint,
    demand,           // 열수요 합계 (시간별 합산)
    demandH[24],      // ★ 시간별 열수요 배열 (sales_hourly → 시간별 Gcal)
    adjExternal,      // 외부수열 × (1 - 손실율)
    adjIntecoNet,     // 인테코수열 × (1 - 손실율) - 인테코판매
    netRequired,      // demand - adjExternal - adjIntecoNet
    smpAvg,           // SMP 24시간 평균
    smpH[24],         // ★ SMP 시간별
    mpH[24],          // ★ MP 시간별 (SMP × TLF + 무부하비정산)
    ngCostNm3,        // CHP NG단가 (원/N㎥)
    ngPLBNm3,         // PLB NG단가 (원/N㎥)
}
```

**핵심**:
- `demandH[24]`: 시간별 수요 패턴 → Stage 1~3 전체에서 축열조 시뮬레이션에 사용
- `mpH[24]`: Stage 2 CHP 블록 배치 + 열제약매출 계산에 사용
- 시간별 데이터가 없는 경우 `demand / 24`로 균등 분배 (폴백)

---

## 6. 주요 파일

| 파일 | 역할 |
|------|------|
| `js/app.js` | 메인 로직 (buildDailyData, runDailySimulation, solveDP) |
| `index.html` | UI (설정, 데이터 입력, 결과 표시) |
| `css/style.css` | 차트 크기 제한, 레이아웃 |
