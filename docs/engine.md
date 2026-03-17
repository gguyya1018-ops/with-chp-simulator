# 시뮬레이션 엔진 로직

> Stage 1~4 시뮬레이션 엔진의 상세 로직
> 마지막 갱신: 2026-03-15

---

## Stage 1: 연간 최적화

**목적**: 각 날의 CHP 가동시간과 부하율을 결정

### 1-1. 휴리스틱 (runDailySimulation)

**2-Pass 방식**:

**Pass 1 — SMP 기반 글로벌 우선순위 배정**
```
1. 정비일 제외, SMP 수익이 높은 순으로 정렬
2. 각 일에 초기 CHP 가동시간 배정:
   - CHP 이익 (SMP수익 > NG비용) → 24시간
   - CHP가 PLB보다 쌈 → 열필요 기반 시간
   - PLB가 더 쌈 → 0시간
```

**Pass 2 — 시간순 시뮬레이션 + 축열조 제약 (시간별 수요 패턴 반영)**
```
For d = 1일 ~ 365일:
   1. 축열조 기반 가동시간 조정 ★ 시간별 수요 패턴 반영
      - findMaxHours(): 24h 시뮬레이션으로 축열조 오버플로 없는 최대 가동시간 산출
      - 시나리오 A (100% 부하 × 축열조 허용시간) vs B (부하↓ × 더 긴 운전) 비교
   2. 연속운전 판단
      - 전날 ON + 오늘 OFF → 정지기간 vs 최소정지시간 비교
      - 정지기간 < 최소정지시간 → 연속운전 강제 (부하 낮춤)
   3. 축열조 부족 시 CHP 투입 강제 ★ 시간별 수요 기반 축열조 최저점 검사
      - checkStorageDeficit(): CHP 없이 24h 돌렸을 때 축열조 최저점 계산
   4. 최종 제약 적용 (최소운전시간, 최소정지시간, 오버플로 방지)
   → chpHours, chpLoad 확정
   → Stage 2~4 실행
```

### 1-2. DP (solveDP)

**Bellman Equation 기반 역방향 동적 프로그래밍**

```
상태: [축열조 잔량, 정지레벨] (2차원)
  - 축열조: 10Gcal 단위 이산화 (예: 100~850 → 76 상태)
  - 정지레벨: 0=가동중, 1=Hot(1일정지), 2=Warm(2일정지), 3=Cold(3일+정지)
  - 총 상태 수 = 축열조상태 × 4

결정: CHP 시간 × 부하율
방향: 역방향 (365일 → 1일)

dp[j][k] = 축열조상태 j, 정지레벨 k에서 마지막 날까지의 최소비용

역방향 ★ 기동비용 + 시간별 수요 패턴 기반:
  For i = 365 → 1:
    For j = 0 → numStates:
      For k = 0 → 3 (정지레벨):
        baseFlow[24] = hourlyExt + intecoH - demandH[t]  ← 시간별 수요 반영
        For 부하 × 시간 조합:
          startupCost = (h>0 && k>0) ? 기동시간×가스사용량×월별NG단가 : 0
          h=0: CHP 없이 24h 축열조 시뮬 → PLB 필요량 산출, nextK = min(k+1, 3)
          h>0: 최적 시작시간 탐색 + 기동비용 포함, nextK = 0
          chpCost + startupCost + plbCost + dp_next[newState][nextK] → 최소 선택

기동비용 산출:
  stopLevel=1 (Hot, ≤1일): hotTime(h) × hotGasRate(Nm³/h) × ngCostNm3
  stopLevel=2 (Warm, 2일): warmTime(h) × warmGasRate(Nm³/h) × ngCostNm3
  stopLevel=3 (Cold, 3일+): coldTime(h) × coldGasRate(Nm³/h) × ngCostNm3

순방향 복원:
  storageInit, stopLevel=3(Cold) → decisions[0][sj][sk] → ...
  → 각 날의 chpHours, chpLoad, startupCost 확정
  → Stage 2~4 실행
```

**DP 특징**:
- 365일 전체를 역방향으로 계산 → 전역 최적해 보장
- ★ 기동비용을 상태에 포함 → 빈번한 ON/OFF 패턴 자동 억제
- ★ 시간별 수요 패턴을 DP 내부에서 직접 반영 → 수요 피크/저점에 따른 정밀 결정
- "주말에 CHP를 안 돌리고 축열조를 비워두면 월요일에 더 벌 수 있다" 자동 발견
- 계산시간: ~수초 (상태 4배 증가하나 여전히 빠름)

### 솔버 비교

| 항목 | 휴리스틱 | DP |
|------|---------|-----|
| 방식 | 2-Pass: SMP 우선순위 + 순방향 | Bellman 역방향 DP |
| 최적성 | 준최적 (로컬 판단) | 전역 최적 |
| 속도 | 즉시 | ~수초 |
| 공통점 | Stage 2~4는 완전히 동일한 로직 | |

---

## Stage 2: 시간별 배치

Stage 1에서 결정된 `chpHours`를 24시간 내 **어디에 배치**할지 결정.

### CHP 연속블록 배치 (스코어링)

CHP는 연속운전이 원칙 → `chpStart` ~ `chpStart + chpHours` 연속블록

```javascript
for (s = 0; s <= 24 - chpHours; s++) {
    // 24시간 시뮬레이션으로 축열조 변화 추적
    for (h = 0; h < 24; h++) {
        stSim += (ext + inteco + chp - demandH[h])
        overflow += max(0, stSim - storageCap)
        underflow += max(0, storageMin - stSim)
        mpSum += mp[h]  (CHP 가동시간만)
    }
    score = mpSum - overflow × 50 - underflow × 200
}
chpStart = argmax(score)
```

**스코어 구성**:
| 요소 | 가중치 | 의미 |
|------|--------|------|
| `mpSum` | +1 | CHP 가동시간대 MP 합계 (높을수록 좋음) |
| `overflow` | -50 | 축열조 상한 초과 (과생산 방지) |
| `underflow` | -200 | 축열조 하한 미달 (PLB 투입 방지) |

**일간 연속성 제약**: `prevChpEndH`로 전날 CHP 종료시각 추적

### PLB 배치 ★ 수요 피크 시간대 기반

```
1. CHP만으로 24h 시뮬 → storageMin 미달 시점 확인 (시간별 수요 기반)
2. deficit = storageMin - 최저 축열조잔량
3. plbUnits, plbHours 산출
4. ★ PLB 최적 시작시간 탐색: 모든 가능한 시작시간에 대해
   24h 축열조 시뮬 → 최저점이 가장 높은 배치를 선택
```

---

## Stage 3: 시간별 시뮬레이션

Stage 2에서 결정된 CHP/PLB 블록을 기반으로 **24시간 순차 실행**.

```javascript
hourly = [];
stH = storageLevel;  // 전날 말 축열조잔량

for (h = 0; h < 24; h++) {
    chp = (CHP 가동시간 내?) ? chpHeatPerH : 0;
    plb = (PLB 가동시간 내?) ? plbHeat/plbHours : 0;

    stH += (hourlyExt + intecoH + chp + plb - demandH[h]);
    if (stH > storageCap) stH = storageCap;
    if (stH < storageMin) stH = storageMin;

    hourly.push({ hour: h, demand, ext, chp, plb, storage, smp, mp });
}

storageLevel = stH;  // ★ 시간별 시뮬 최종값 → 다음날로 전달
```

### hourly 배열 구조
```javascript
{
    hour: 0~23,
    demand: number,  // 시간별 수요 (Gcal/h)
    ext: number,     // 외부수열 + 인테코 (Gcal/h)
    chp: number,     // CHP 열생산 (Gcal/h)
    plb: number,     // PLB 열생산 (Gcal/h)
    storage: number, // 축열조잔량 (Gcal)
    smp: number,     // SMP (원/kWh)
    mp: number,      // MP (원/kWh)
}
```

---

## Stage 4: 집계

### 일별 집계 → results[]
- 열수급: demand, adjExternal, adjIntecoNet, netRequired
- CHP: chpHours, chpHeat, chpPower, chpNg, chpLoad, chpFuelCost, chpElecRev, chpNetCost
- PLB: plbUnits, plbHours, plbHeat, plbNg, plbFuelCost
- 비용: totalFuelCost, totalNetCost
- 시간별 데이터: hourly[24], smpH[24]

### 월별 집계 (aggregateMonthly)
- 가동일수, 가동시간, 생산열, 발전량, NG소비
- PLB 가동일, 최대 대수
- 비용: 연료비, 전기수익, 순비용, 용량요금

---

## 차트 시각화

### 일 상세 모달 (시간별 차트)
```
차트 데이터셋:
- 축열조 (bar, Y1축)
- 수요 (line, 빨강, Y2축) ★ 시간별 판매실적 합산
- 외부수열 (line, 초록 점선, Y2축)
- CHP (line, 파랑, Y2축)
- PLB (line, 주황, Y2축)
```
DOM 재사용 패턴: 모달 구조 1회 생성, 이후 데이터만 갱신 (깜빡임 없음)
