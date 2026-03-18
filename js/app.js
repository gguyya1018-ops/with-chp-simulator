/**
 * CHP 보일러 기동계획 시뮬레이터
 * Phase 2 - 최적운전 계획 수립
 */

/* ═══════════════════════════════════════════════
   커스텀 타이틀바 (frameless)
   ═══════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
    const btnMin = document.getElementById('tbMin');
    const btnMax = document.getElementById('tbMax');
    const btnClose = document.getElementById('tbClose');
    if (!btnMin) return;
    btnMin.addEventListener('click', () => pywebview.api.win_minimize());
    btnMax.addEventListener('click', () => pywebview.api.win_toggle_max());
    btnClose.addEventListener('click', () => pywebview.api.win_close());
    // 더블클릭으로 최대화 토글
    document.querySelector('.titlebar-drag').addEventListener('dblclick', () => {
        pywebview.api.win_toggle_max();
    });
});

/* ═══════════════════════════════════════════════
   전역 상태
   ═══════════════════════════════════════════════ */
const DATA = {};            // { key: { headers:[], rows:[][] } }
let CURRENT_PROJECT = null;
let CURRENT_YEAR = 2024;
let YEAR_DAYS = [];
let PLAN_RESULTS = null;    // 시뮬레이션 결과

// 데이터 키 정의 (독립 관리)
const INPUT_KEYS = ['smp_hourly', 'operations_daily', 'sales_hourly', 'sales_daily', 'water_chem_daily', 'holidays',
    'capacity_hourly', 'dispatch_hourly', 'heat_constraint_hourly'];
const PRICING_KEYS = [
    'cp_monthly', 'ng_price_monthly', 'base_charge_monthly',
    'elec_rate_weekday', 'elec_rate_saturday', 'elec_rate_holiday',
    'heat_rate_table', 'water_price_monthly', 'chem_price_monthly',
    'sale_prices_daily'
];
const ALL_DATA_KEYS = [...INPUT_KEYS, ...PRICING_KEYS];

// 전기요금표 → elec_rate 탭에서 3개를 묶어 표시
const ELEC_RATE_KEYS = ['elec_rate_weekday', 'elec_rate_saturday', 'elec_rate_holiday'];

// 월별/시간별 편집 가능 테이블 키
const MONTHLY_KEYS = [
    'cp_monthly', 'ng_price_monthly', 'base_charge_monthly',
    'elec_rate_weekday', 'elec_rate_saturday', 'elec_rate_holiday',
    'heat_rate_table', 'water_price_monthly', 'chem_price_monthly'
];

const TEMPLATES = {
    smp_hourly:         { type:'hourly_wide', headers:['month','day','1h','2h','3h','4h','5h','6h','7h','8h','9h','10h','11h','12h','13h','14h','15h','16h','17h','18h','19h','20h','21h','22h','23h','24h'] },
    operations_daily:   { type:'daily', headers:['월','일','보유열량','CHP(급전)','CHP(열제약)','CHP 합산','#1 PLB','#2 PLB','#3 PLB','PLB 합산','남부소각장','ERG','SRF','인테코','안산(SP)','안산(GS E&R)','안산(CO기본)','안산(CO추가)','안산(열제약)','안산합산'] },
    sales_hourly:       { type:'hourly_daily', headers:['월','일','시','주택용난방','업무용난방','업무용냉방','공공용난방','공공용냉방'] },
    sales_daily:        { type:'daily', headers:['월','일','인테코'] },
    water_chem_daily:   { type:'daily', headers:['월','일','용수사용량(㎥)','가성소다(Kg)','탈산소제(Kg)','암모니아수(Kg)','염산(Kg)','카보하이드라자이드(Kg)','PH조절제(Kg)'] },
    holidays:           { type:'daily', headers:['month','day','weekday','is_holiday'] },
    cp_monthly:         { type:'monthly', headers:['월','평일CP','휴일CP','TLF(평일)','TLF(토요일)','TLF(공휴일)'] },
    ng_price_monthly:   { type:'monthly', headers:['월','NG단위열량','안전관리부담금(원/N㎥)','석유수입부과금(원/MJ)','CHP열량단가(원/MJ)','PLB열량단가(원/MJ)'] },
    base_charge_monthly:{ type:'monthly', headers:['월','기본료(154kV)','기본료(22kV)','기후환경요금계수(원/kWh)','연료비조정계수(원/kWh)','전력기금(%)'] },
    elec_rate_weekday:  { type:'monthly_hours', headers:['hour','m1','m2','m3','m4','m5','m6','m7','m8','m9','m10','m11','m12'] },
    elec_rate_saturday: { type:'monthly_hours', headers:['hour','m1','m2','m3','m4','m5','m6','m7','m8','m9','m10','m11','m12'] },
    elec_rate_holiday:  { type:'monthly_hours', headers:['hour','m1','m2','m3','m4','m5','m6','m7','m8','m9','m10','m11','m12'] },
    heat_rate_table:    { type:'monthly', headers:['월','기본료','주택용','업무용','공공용','공공용(수영장)','수영장비율(%)','순추절기'] },
    water_price_monthly:{ type:'monthly', headers:['월','정액요금(원)','물이용단가','상수도단가(300이하)','상수도단가(300초과)','하수도단가(50이하)','하수도단가(50~100)','하수도단가(100~300)','하수도단가(300~500)','하수도단가(500~1000)','하수도단가(1000초과)'] },
    chem_price_monthly: { type:'monthly', headers:['월','가성소다','탈산소제','암모니아수','염산','카보하이드라사이드','PH조절제'] },
    sale_prices_daily:  { type:'daily', headers:['월','일','ERG','송도소각장','SRF','인테코','그린스코','안산(SP)','안산(GS E&R)','안산(CO기본)','안산(CO추가)','안산(열제약)','인테코(판매)'] },
    capacity_hourly:    { type:'hourly_wide', headers:['month','day','1h','2h','3h','4h','5h','6h','7h','8h','9h','10h','11h','12h','13h','14h','15h','16h','17h','18h','19h','20h','21h','22h','23h','24h'] },
    dispatch_hourly:    { type:'hourly_wide', headers:['month','day','1h','2h','3h','4h','5h','6h','7h','8h','9h','10h','11h','12h','13h','14h','15h','16h','17h','18h','19h','20h','21h','22h','23h','24h'] },
    heat_constraint_hourly:{ type:'hourly_wide', headers:['month','day','1h','2h','3h','4h','5h','6h','7h','8h','9h','10h','11h','12h','13h','14h','15h','16h','17h','18h','19h','20h','21h','22h','23h','24h'] },
};

/* ═══════════════════════════════════════════════
   외부수열원 동적 레지스트리
   - type: 'supply'(단방향 수열) | 'dual'(수열+판매, 인테코)
   - enabled: false면 데이터 유지하되 시뮬에서 제외
   ═══════════════════════════════════════════════ */
const DEFAULT_EXT_SOURCES = [
    { id:'nambu',      opsName:'남부소각장',    priceName:'송도소각장',   type:'supply', enabled:true },
    { id:'erg',        opsName:'ERG',           priceName:'ERG',          type:'supply', enabled:true },
    { id:'srf',        opsName:'SRF',           priceName:'SRF',          type:'supply', enabled:true },
    { id:'inteco',     opsName:'인테코',        priceName:'인테코',       type:'dual',   enabled:true, salePriceName:'인테코(판매)' },
    { id:'ansan_sp',   opsName:'안산(SP)',       priceName:'안산(SP)',      type:'supply', enabled:true },
    { id:'ansan_gs',   opsName:'안산(GS E&R)',   priceName:'안산(GS E&R)', type:'supply', enabled:true },
    { id:'ansan_co1',  opsName:'안산(CO기본)',   priceName:'안산(CO기본)', type:'supply', enabled:true },
    { id:'ansan_co2',  opsName:'안산(CO추가)',   priceName:'안산(CO추가)', type:'supply', enabled:true },
    { id:'ansan_heat', opsName:'안산(열제약)',   priceName:'안산(열제약)', type:'supply', enabled:true },
];
let EXT_SOURCES = JSON.parse(JSON.stringify(DEFAULT_EXT_SOURCES));

function rebuildTemplateHeaders() {
    // operations_daily: 고정열 + 동적 외부열원
    const opsFixed = ['월','일','보유열량','CHP(급전)','CHP(열제약)','CHP 합산','#1 PLB','#2 PLB','#3 PLB','PLB 합산'];
    const opsDynamic = [];
    EXT_SOURCES.forEach(s => {
        opsDynamic.push(s.opsName);
        if (s.type === 'dual' && s.salePriceName) opsDynamic.push(s.opsName + '(판매)');
    });
    TEMPLATES.operations_daily.headers = [...opsFixed, ...opsDynamic];

    // sale_prices_daily: 고정열 + 동적 단가열
    const priceFixed = ['월','일'];
    const priceDynamic = [];
    EXT_SOURCES.forEach(s => {
        priceDynamic.push(s.priceName);
        if (s.type === 'dual' && s.salePriceName) priceDynamic.push(s.salePriceName);
    });
    TEMPLATES.sale_prices_daily.headers = [...priceFixed, ...priceDynamic];
}
rebuildTemplateHeaders();

/* ═══════════════════════════════════════════════
   유틸리티
   ═══════════════════════════════════════════════ */
function parseCSV(text) {
    const lines = text.trim().split('\n').filter(l => l.trim());
    if (!lines.length) return null;
    const headers = lines[0].split(',').map(h => h.trim());
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const cells = lines[i].split(',').map(c => {
            const t = c.trim();
            const n = Number(t);
            return t !== '' && !isNaN(n) ? n : t;
        });
        rows.push(cells);
    }
    return { headers, rows };
}

function csvToString(headers, rows) {
    return [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
}

function fmtNum(n) {
    if (typeof n !== 'number') return n;
    if (Number.isInteger(n)) return n.toLocaleString();
    // 소수점 원본 자릿수 보존 (최대 4자리)
    const s = String(n);
    const dec = s.includes('.') ? s.split('.')[1].length : 0;
    return n.toLocaleString(undefined, {
        minimumFractionDigits: Math.min(dec, 4),
        maximumFractionDigits: Math.min(dec, 4)
    });
}

// operations_daily 불필요 컬럼 제거 + 보유열량 컬럼 보장
function stripOperationsCols() {
    if (!DATA.operations_daily) return;
    const d = DATA.operations_daily;
    // 불필요 컬럼 제거
    const skip = new Set(['인테코(그린스코)', '인테코 합산', '안산합산']);
    if (d.headers.some(h => skip.has(h))) {
        const keepIdx = d.headers.map((h, i) => skip.has(h) ? -1 : i).filter(i => i >= 0);
        d.headers = keepIdx.map(i => d.headers[i]);
        d.rows = d.rows.map(row => keepIdx.map(i => row[i]));
    }
    // 보유열량 컬럼이 없으면 '일' 다음(인덱스 2)에 삽입
    if (!d.headers.includes('보유열량')) {
        d.headers.splice(2, 0, '보유열량');
        d.rows.forEach(row => row.splice(2, 0, 0));
    }
}

/* ═══════════════════════════════════════════════
   메인 화면 - 프로젝트 관리
   ═══════════════════════════════════════════════ */
async function loadMainScreen() {
    // 프로젝트 목록
    const projects = await pywebview.api.list_projects();
    const list = document.getElementById('projectList');
    if (!projects.length) {
        list.innerHTML = '<p class="hint">프로젝트가 없습니다</p>';
    } else {
        list.innerHTML = projects.map(p =>
            `<button class="proj-item" onclick="openProject('${p}')">${p}</button>`
        ).join('');
    }

}

async function createProject() {
    const name = document.getElementById('newProjectName').value.trim();
    const year = parseInt(document.getElementById('newProjectYear').value) || 2024;
    if (!name) return alert('프로젝트명을 입력하세요');

    const res = await pywebview.api.create_project(name);
    if (res.error) return alert(res.error);

    // 연도 설정 저장
    await pywebview.api.save_csv(name, '_settings',
        JSON.stringify({ year, created: new Date().toISOString() }));

    openProject(name);
}

async function openProject(name) {
    CURRENT_PROJECT = name;
    document.getElementById('navProjectName').textContent = name;

    // CSV 로드
    const csvData = await pywebview.api.load_project(name);
    for (const key of ALL_DATA_KEYS) {
        if (csvData[key] && !csvData[key].startsWith('ERROR:')) {
            DATA[key] = parseCSV(csvData[key]);
        } else {
            delete DATA[key];
        }
    }

    stripOperationsCols();

    // 설정 로드
    if (csvData._settings && !csvData._settings.startsWith('ERROR:')) {
        try {
            const settings = JSON.parse(csvData._settings);
            CURRENT_YEAR = settings.year || 2024;
            if (settings.ui) applySettings(settings.ui);
        } catch (e) { console.warn('설정 파싱 실패:', e); }
    }

    // 연도 데이터
    YEAR_DAYS = await pywebview.api.get_year_days(CURRENT_YEAR);

    // 열요금표 기본값 (운전최적화 동일)
    if (!DATA.heat_rate_table) {
        DATA.heat_rate_table = {
            headers: ['월','기본료','주택용','업무용','공공용','공공용(수영장)','수영장비율(%)','순추절기'],
            rows: [
                [1, 327641800, 101570, 131870, 115000, 0, 0.6663, 99510],
                [2, 327641800, 101570, 131870, 115000, 0, 0.6239, 99510],
                [3, 327641800, 101570, 131870, 115000, 0, 0.6797, 99510],
                [4, 327441600, 101570, 131870, 115000, 0, 0.0785, 99510],
                [5, 328674600, 101570, 131870, 115000, 0, 0.0325, 99510],
                [6, 329548000, 101570, 131870, 115000, 0, 0, 99510],
                [7, 329693400, 112320, 145820, 127340, 0, 0, 110040],
                [8, 330038800, 112320, 145820, 127340, 0, 0, 110040],
                [9, 330038800, 112320, 145820, 127340, 0, 0.5286, 110040],
                [10, 330038800, 112320, 145820, 127340, 0, 0.8834, 110040],
                [11, 330038800, 112320, 145820, 127340, 0, 0.7738, 110040],
                [12, 330038800, 112320, 145820, 127340, 0, 0.6647, 110040],
            ]
        };
    }

    // 화면 전환
    document.getElementById('mainScreen').style.display = 'none';
    document.getElementById('appScreen').style.display = '';

    updateDataStatus();
    renderAllPreviews();

    renderHeatBalance();
}

function goHome() {
    document.getElementById('appScreen').style.display = 'none';
    document.getElementById('mainScreen').style.display = '';
    loadMainScreen();
}

/* ═══════════════════════════════════════════════
   데이터 상태 표시
   ═══════════════════════════════════════════════ */
function updateDataStatus() {
    let loaded = 0;
    ALL_DATA_KEYS.forEach(k => { if (DATA[k]) loaded++; });
    document.getElementById('dataStatus').textContent = `${loaded} / ${ALL_DATA_KEYS.length}`;
    if (typeof checkPipelineStatus === 'function') checkPipelineStatus();
}

/* ═══════════════════════════════════════════════
   데이터 미리보기 렌더링
   ═══════════════════════════════════════════════ */
function renderPreview(key) {
    // 월별/시간별 데이터는 편집 가능 테이블로 렌더링
    const tmpl = TEMPLATES[key];
    if (tmpl && (tmpl.type === 'monthly' || tmpl.type === 'monthly_hours')) {
        renderEditableMonthly(key);
        return;
    }

    const el = document.getElementById('pv_' + key);
    if (!el) return;
    const d = DATA[key];
    if (!d) { el.innerHTML = '<p class="hint">데이터 없음 - 업로드하거나 가져오세요</p>'; return; }

    // 외부거래: CHP, PLB, 보유열량 컬럼 숨김 (시뮬 산출값이므로 입력 데이터 미리보기에서 제외)
    const hideKeywords = (key === 'operations_daily') ? ['보유열량', 'CHP', 'PLB'] : [];
    const hideCols = new Set();
    if (hideKeywords.length) {
        d.headers.forEach((h, i) => { if (hideKeywords.some(kw => h.toUpperCase().includes(kw.toUpperCase()))) hideCols.add(i); });
    }

    // 인테코(판매) 병합: operations_daily에 sales_daily 인테코 판매량 추가 표시
    const isOps = (key === 'operations_daily');
    const salesD = isOps ? DATA.sales_daily : null;
    const saleDayMap = {};
    if (salesD) {
        const sIntecoCol = salesD.headers.indexOf('인테코');
        if (sIntecoCol >= 0) {
            salesD.rows.forEach(row => {
                const k = `${Number(row[0])}-${Number(row[1])}`;
                saleDayMap[k] = Number(row[sIntecoCol]) || 0;
            });
        }
    }
    const showIntecoSale = isOps && Object.keys(saleDayMap).length > 0;

    // 외부열원 컬럼명 Set (이름변경 가능 표시용)
    const extNames = new Set(EXT_SOURCES.map(s => s.opsName));

    const maxRows = Math.min(d.rows.length, 30);
    let html = '<table class="pv-tbl"><thead><tr>';
    d.headers.forEach((h, i) => {
        if (hideCols.has(i)) return;
        if (isOps && extNames.has(h)) {
            html += `<th class="ext-rename-th" ondblclick="renameExtSource('${h.replace(/'/g,"\\'")}')\" title="더블클릭으로 이름 변경">${h}</th>`;
        } else {
            html += `<th>${h}</th>`;
        }
    });
    if (showIntecoSale) html += '<th class="calc-col">인테코(판매)</th>';
    html += '</tr></thead><tbody>';

    for (let i = 0; i < maxRows; i++) {
        html += '<tr>';
        d.rows[i].forEach((c, ci) => { if (!hideCols.has(ci)) html += `<td>${fmtNum(c)}</td>`; });
        if (showIntecoSale) {
            const k = `${Number(d.rows[i][0])}-${Number(d.rows[i][1])}`;
            const v = saleDayMap[k] || 0;
            html += `<td class="calc-col">${fmtNum(v)}</td>`;
        }
        html += '</tr>';
    }
    html += '</tbody></table>';
    if (d.rows.length > maxRows) {
        html += `<div class="pv-more">... 외 ${d.rows.length - maxRows}행</div>`;
    }
    el.innerHTML = html;
}

function renderAllPreviews() {
    ALL_DATA_KEYS.forEach(k => {
        // elec_rate 3개는 개별 pv_elec_rate_* 패널로 각각 렌더링
        if (ELEC_RATE_KEYS.includes(k)) {
            renderPreview(k); // renderEditableMonthly 로 디스패치됨
            return;
        }
        renderPreview(k);
    });
    // 데이터 산출 자동 계산
    runAllCalc();
}

/* ═══════════════════════════════════════════════
   편집 가능 월별/시간별 테이블
   ═══════════════════════════════════════════════ */
function renderEditableMonthly(key) {
    const el = document.getElementById('pv_' + key);
    if (!el) return;
    const tmpl = TEMPLATES[key];
    const d = DATA[key];

    let rowCount, firstColGen;
    if (tmpl.type === 'monthly') {
        rowCount = 12;
        firstColGen = i => i + 1;
    } else if (tmpl.type === 'monthly_hours') {
        rowCount = 24;
        firstColGen = i => (i + 1) + 'h';
    } else return;

    // 열요금표: 계산 컬럼 추가 (업무용냉방, 공공용냉방) + 수영장비율 데이터 컬럼
    const isHeatRate = (key === 'heat_rate_table');
    const heatColOrder = isHeatRate ? [
        { type:'data', idx:0, label:'월' },
        { type:'data', idx:1, label:'기본료' },
        { type:'data', idx:2, label:'주택용' },
        { type:'data', idx:3, label:'업무용' },
        { type:'calc', label:'업무용(냉방)' },
        { type:'data', idx:4, label:'공공용' },
        { type:'calc', label:'공공용(냉방)' },
        { type:'data', idx:5, label:'공공용(수영장)' },
        { type:'data', idx:6, label:'수영장비율(%)' },
        { type:'data', idx:7, label:'순추절기' },
    ] : null;

    // NG단가: 계산 컬럼 추가 (CHP단가, PLB단가)
    const isNgPrice = (key === 'ng_price_monthly');
    const ngColOrder = isNgPrice ? [
        { type:'data', idx:0, label:'월' },
        { type:'data', idx:1, label:'NG단위열량' },
        { type:'data', idx:2, label:'안전관리부담금(원/N㎥)' },
        { type:'data', idx:3, label:'석유수입부과금(원/MJ)' },
        { type:'data', idx:4, label:'CHP열량단가(원/MJ)' },
        { type:'data', idx:5, label:'PLB열량단가(원/MJ)' },
        { type:'calc', label:'CHP단가(원/N㎥)' },
        { type:'calc', label:'PLB단가(원/N㎥)' },
    ] : null;

    const customCols = heatColOrder || ngColOrder;
    let html = `<table class="pv-tbl editable-tbl" data-key="${key}"><thead><tr>`;
    if (customCols) {
        customCols.forEach(col => html += col.type === 'calc' ? `<th class="calc-col">${col.label}</th>` : `<th>${col.label}</th>`);
    } else {
        tmpl.headers.forEach(h => html += `<th>${h}</th>`);
    }
    html += '</tr></thead><tbody>';

    for (let i = 0; i < rowCount; i++) {
        html += '<tr>';
        if (isHeatRate) {
            const m = i + 1;
            const row = d && d.rows[i] ? d.rows[i] : null;
            const 업무용 = row ? (Number(row[3]) || 0) : 0;
            const 공공용 = row ? (Number(row[4]) || 0) : 0;
            const isCool = (m >= 5 && m <= 9);
            heatColOrder.forEach(col => {
                if (col.type === 'data' && col.idx === 0) {
                    html += `<td>${row ? row[0] : m}</td>`;
                } else if (col.type === 'data') {
                    const val = row && row[col.idx] != null ? row[col.idx] : '';
                    html += `<td><input type="text" value="${val}" data-r="${i}" data-c="${col.idx}"></td>`;
                } else if (col.label === '업무용(냉방)') {
                    html += `<td class="calc-col">${(isCool ? Math.round(업무용 * 0.4) : 업무용).toLocaleString()}</td>`;
                } else if (col.label === '공공용(냉방)') {
                    html += `<td class="calc-col">${(isCool ? Math.round(공공용 * 0.4) : 공공용).toLocaleString()}</td>`;
                }
            });
        } else if (isNgPrice) {
            const row = d && d.rows[i] ? d.rows[i] : null;
            const B = row ? (Number(row[1]) || 0) : 0;
            const C = row ? (Number(row[2]) || 0) : 0;
            const D = row ? (Number(row[3]) || 0) : 0;
            const E = row ? (Number(row[4]) || 0) : 0;
            const F = row ? (Number(row[5]) || 0) : 0;
            ngColOrder.forEach(col => {
                if (col.type === 'data' && col.idx === 0) {
                    html += `<td>${row ? row[0] : i + 1}</td>`;
                } else if (col.type === 'data') {
                    const val = row && row[col.idx] != null ? row[col.idx] : '';
                    html += `<td><input type="text" value="${val}" data-r="${i}" data-c="${col.idx}"></td>`;
                } else if (col.label === 'CHP단가(원/N㎥)') {
                    const chp = Math.round((B * (E - D) - C) * 100) / 100;
                    html += `<td class="calc-col">${chp.toLocaleString()}</td>`;
                } else if (col.label === 'PLB단가(원/N㎥)') {
                    const plb = Math.round((B * F) * 100) / 100;
                    html += `<td class="calc-col">${plb.toLocaleString()}</td>`;
                }
            });
        } else {
            // 일반 월별/시간별 테이블
            tmpl.headers.forEach((h, j) => {
                if (j === 0) {
                    const val = d && d.rows[i] ? d.rows[i][0] : firstColGen(i);
                    html += `<td>${val}</td>`;
                } else {
                    let val = d && d.rows[i] && d.rows[i][j] != null ? d.rows[i][j] : '';
                    if (typeof val === 'string') val = val.replace(/,/g, '');
                    html += `<td><input type="text" value="${val}" data-r="${i}" data-c="${j}"></td>`;
                }
            });
        }
        html += '</tr>';
    }
    html += '</tbody></table>';
    html += '<div class="pv-hint">셀을 직접 편집하거나 엑셀에서 복사(Ctrl+V)하여 붙여넣기 가능</div>';
    el.innerHTML = html;

    // 열요금표/NG단가: 입력 변경 시 계산 컬럼 실시간 갱신
    if (isHeatRate) {
        el.querySelectorAll('input').forEach(inp => {
            inp.addEventListener('input', () => updateHeatRateCalcCols());
        });
    }
    if (isNgPrice) {
        el.querySelectorAll('input').forEach(inp => {
            inp.addEventListener('input', () => updateNgPriceCalcCols());
        });
    }
}

/* ── 열요금표 계산 컬럼 실시간 갱신 ── */
function updateHeatRateCalcCols() {
    const el = document.getElementById('pv_heat_rate_table');
    if (!el) return;
    el.querySelectorAll('tbody tr').forEach((tr, i) => {
        const m = i + 1;
        const 업무용 = Number(tr.querySelector('input[data-c="3"]')?.value) || 0;
        const 공공용 = Number(tr.querySelector('input[data-c="4"]')?.value) || 0;
        const isCool = (m >= 5 && m <= 9);
        const calcTds = tr.querySelectorAll('td.calc-col');
        if (calcTds.length >= 2) {
            calcTds[0].textContent = (isCool ? Math.round(업무용 * 0.4) : 업무용).toLocaleString();
            calcTds[1].textContent = (isCool ? Math.round(공공용 * 0.4) : 공공용).toLocaleString();
        }
    });
}

/* ── NG단가 계산 컬럼 실시간 갱신 ── */
function updateNgPriceCalcCols() {
    const el = document.getElementById('pv_ng_price_monthly');
    if (!el) return;
    el.querySelectorAll('tbody tr').forEach(tr => {
        const B = Number(tr.querySelector('input[data-c="1"]')?.value) || 0;
        const C = Number(tr.querySelector('input[data-c="2"]')?.value) || 0;
        const D = Number(tr.querySelector('input[data-c="3"]')?.value) || 0;
        const E = Number(tr.querySelector('input[data-c="4"]')?.value) || 0;
        const F = Number(tr.querySelector('input[data-c="5"]')?.value) || 0;
        const calcTds = tr.querySelectorAll('td.calc-col');
        if (calcTds.length >= 2) {
            calcTds[0].textContent = (Math.round((B * (E - D) - C) * 100) / 100).toLocaleString();
            calcTds[1].textContent = (Math.round((B * F) * 100) / 100).toLocaleString();
        }
    });
}

/* ═══════════════════════════════════════════════
   월별 편집 테이블 ↔ CSV 자동 저장
   ═══════════════════════════════════════════════ */
function collectEditableMonthly(key) {
    const el = document.getElementById('pv_' + key);
    if (!el) return null;
    const tmpl = TEMPLATES[key];
    const table = el.querySelector('.editable-tbl');
    if (!table) return null;

    const rows = [];
    table.querySelectorAll('tbody tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('td').forEach((td, j) => {
            if (j === 0) {
                const v = td.textContent.trim();
                const n = Number(v.replace('h', ''));
                row.push(tmpl.type === 'monthly_hours' ? v : (isNaN(n) ? v : n));
            } else {
                const inp = td.querySelector('input');
                if (!inp) return; // calc-col 등 input 없는 셀 건너뜀
                const v = inp.value.trim();
                const cleaned = v.replace(/,/g, '');
                const n = Number(cleaned);
                row.push(v === '' ? '' : (isNaN(n) ? v : n));
            }
        });
        rows.push(row);
    });

    return { headers: [...tmpl.headers], rows };
}

let _monthlySaveTimers = {};
function scheduleMonthlyAutoSave(key) {
    if (!CURRENT_PROJECT) return;
    clearTimeout(_monthlySaveTimers[key]);
    _monthlySaveTimers[key] = setTimeout(() => saveEditableMonthlyNow(key), 800);
}

async function saveEditableMonthlyNow(key) {
    const data = collectEditableMonthly(key);
    if (!data || !CURRENT_PROJECT) return;
    DATA[key] = data;
    updateDataStatus();

    const lines = [data.headers.join(',')];
    data.rows.forEach(row => lines.push(row.join(',')));
    try {
        await pywebview.api.save_csv(CURRENT_PROJECT, key, lines.join('\n'));
    } catch (e) { console.error(`${key} 저장 실패:`, e); }
    // 단가 변경 시 데이터 산출 자동 갱신
    if (['water_price_monthly', 'chem_price_monthly', 'cp_monthly'].includes(key)) runAllCalc();
}

function handleMonthlyPaste(e, key) {
    const target = e.target;
    if (target.tagName !== 'INPUT') return;

    const clipData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipData) return;

    const pasteRows = clipData.trim().split(/\r?\n/).map(line => line.split('\t'));
    if (pasteRows.length <= 1 && pasteRows[0].length <= 1) return; // 단일 셀은 기본 동작

    e.preventDefault();
    const startRow = parseInt(target.dataset.r);
    const startCol = parseInt(target.dataset.c);
    const table = target.closest('.editable-tbl');
    if (!table) return;

    pasteRows.forEach((cols, ri) => {
        cols.forEach((val, ci) => {
            const inp = table.querySelector(`input[data-r="${startRow + ri}"][data-c="${startCol + ci}"]`);
            if (inp) inp.value = val.trim().replace(/,/g, '');
        });
    });

    // 계산 컬럼 갱신
    if (key === 'heat_rate_table') updateHeatRateCalcCols();
    if (key === 'ng_price_monthly') updateNgPriceCalcCols();
    scheduleMonthlyAutoSave(key);
}

/* ═══════════════════════════════════════════════
   CSV 업로드 (개별 / Excel)
   ═══════════════════════════════════════════════ */
async function uploadSingle(key) {
    const res = await pywebview.api.open_file_dialog(['CSV Files (*.csv)']);
    if (!res || res.error) return;
    const parsed = parseCSV(res.content);
    if (!parsed) return;
    DATA[key] = parsed;
    if (key === 'operations_daily') stripOperationsCols();
    await pywebview.api.save_csv(CURRENT_PROJECT, key, res.content);
    updateDataStatus();
    renderPreview(key);

    renderHeatBalance();
    // 관련 데이터 변경 시 데이터 산출 자동 갱신
    const CALC_TRIGGER_KEYS = ['smp_hourly', 'cp_monthly', 'water_chem_daily', 'water_price_monthly', 'chem_price_monthly',
        'operations_daily', 'sale_prices_daily', 'sales_hourly', 'sales_daily', 'heat_rate_table', 'holidays'];
    if (CALC_TRIGGER_KEYS.includes(key)) runAllCalc();
}

// 시트명 → 데이터키 매핑 (한글 + 영문 키 모두 지원)
const KEY_DISPLAY = {
    smp_hourly:'SMP', operations_daily:'외부거래', sales_hourly:'판매실적(시간별)', sales_daily:'인테코',
    water_chem_daily:'용수약품', sale_prices_daily:'판매단가', holidays:'공휴일',
    cp_monthly:'CP_TLF', ng_price_monthly:'NG단가', base_charge_monthly:'전기요금설정',
    elec_rate_weekday:'전기요금_평일', elec_rate_saturday:'전기요금_토요일', elec_rate_holiday:'전기요금_공휴일',
    heat_rate_table:'열요금표', water_price_monthly:'수도단가', chem_price_monthly:'약품단가',
    capacity_hourly:'수전량', dispatch_hourly:'급전송전량', heat_constraint_hourly:'열제약송전량',
};
const SHEET_MAP = {};
// 한글 시트명 → 키
Object.entries(KEY_DISPLAY).forEach(([k, v]) => SHEET_MAP[v] = k);
// 영문 키 → 키 (다운로드 Excel 재업로드 호환)
ALL_DATA_KEYS.forEach(k => SHEET_MAP[k] = k);
// 운전최적화 레거시 시트명
Object.assign(SHEET_MAP, {
    '급전':'dispatch_hourly', '열제약':'heat_constraint_hourly',
    '수전량':'capacity_hourly', 'NG사용량':'ng_usage_daily',
    '송전량(급전)':'dispatch_hourly', '송전량(열제약)':'heat_constraint_hourly',
    '급전송전량':'dispatch_hourly', '열제약송전량':'heat_constraint_hourly',
});

async function uploadExcel() {
    const res = await pywebview.api.open_file_dialog(['Excel Files (*.xlsx;*.xls)']);
    if (!res || !res.sheets) return;
    let count = 0;
    for (const [sheetName, csv] of Object.entries(res.sheets)) {
        const key = SHEET_MAP[sheetName];
        if (key && ALL_DATA_KEYS.includes(key)) {
            const parsed = parseCSV(csv);
            if (parsed) {
                DATA[key] = parsed;
                await pywebview.api.save_csv(CURRENT_PROJECT, key, csv);
                count++;
            }
        }
    }
    stripOperationsCols();
    // 업로드된 operations_daily에서 미등록 외부열원 컬럼 자동 감지
    autoDetectExtSources();
    updateDataStatus();
    renderAllPreviews();

    renderHeatBalance();
    if (count) alert(`${count}개 시트를 업로드했습니다.`);
}

// 업로드된 operations_daily에서 미등록 외부열원 컬럼 자동 감지/등록
function autoDetectExtSources() {
    if (!DATA.operations_daily) return;
    const oH = DATA.operations_daily.headers;
    const fixedCols = new Set(['월','일','보유열량','CHP(급전)','CHP(열제약)','CHP 합산','#1 PLB','#2 PLB','#3 PLB','PLB 합산']);
    const knownOps = new Set(EXT_SOURCES.map(s => s.opsName));
    // dual 소스의 판매 컬럼도 알려진 것으로 등록
    EXT_SOURCES.forEach(s => { if (s.type === 'dual') knownOps.add(s.opsName + '(판매)'); });
    let added = false;
    oH.forEach(h => {
        if (!fixedCols.has(h) && !knownOps.has(h)) {
            // 미등록 컬럼 → 자동 등록
            EXT_SOURCES.push({ id: 'ext_' + Date.now() + '_' + Math.random().toString(36).slice(2,6), opsName: h, priceName: h, type: 'supply', enabled: true });
            knownOps.add(h);
            added = true;
        }
    });
    if (added) {
        rebuildTemplateHeaders();
        if (typeof renderExtSourcesUI === 'function') renderExtSourcesUI();
        scheduleSettingsSave();
    }
}

async function downloadExcel() {
    if (!CURRENT_PROJECT) return alert('프로젝트를 먼저 열어주세요');
    rebuildTemplateHeaders(); // 현재 EXT_SOURCES 기준 헤더 갱신
    const data = {};
    ALL_DATA_KEYS.forEach(key => {
        const sheetName = KEY_DISPLAY[key] || key;
        if (DATA[key]) {
            const d = DATA[key];
            const lines = [d.headers.join(',')];
            d.rows.forEach(row => lines.push(row.join(',')));
            data[sheetName] = lines.join('\n');
        } else if (TEMPLATES[key]) {
            // 데이터 없어도 헤더만 빈 시트로 포함
            data[sheetName] = TEMPLATES[key].headers.join(',');
        }
    });
    if (!Object.keys(data).length) return alert('내보낼 데이터가 없습니다');
    const res = await pywebview.api.export_excel(CURRENT_PROJECT, data);
    if (res && res.ok) alert('Excel 파일이 저장되었습니다.');
}

/* ═══════════════════════════════════════════════
   설정 저장/불러오기
   ═══════════════════════════════════════════════ */
const SETTINGS_IDS = [
    'chpMinRunHours', 'chpMinStopHours', 'chpMinLoad', 'plbCount', 'plbAuxCost', 'plbMinRunHours',
    'storageCapacity', 'storageInitial', 'storageChargeRate', 'storageDischargeRate', 'storageMinLevel', 'plbRecoveryTarget',
    'hotTime', 'hotGasRate', 'warmTime', 'warmGasRate', 'coldTime', 'coldGasRate',
    'mt1StartM', 'mt1StartD', 'mt1EndM', 'mt1EndD',
    'mt2StartM', 'mt2StartD', 'mt2EndM', 'mt2EndD',
    'mt3StartM', 'mt3StartD', 'mt3EndM', 'mt3EndD',
    'pricingInc1', 'pricingInc2', 'pricingConst',
    'standbyPowerKW', 'demandGrowthPct', 'demandDayFactor', 'demandNightFactor',
];

function collectSettings() {
    const ui = {};
    SETTINGS_IDS.forEach(id => {
        const el = document.getElementById(id);
        if (el) ui[id] = el.value;
    });

    // loadTable
    ui.loadTable = [];
    document.querySelectorAll('#loadTable tbody tr').forEach(tr => {
        ui.loadTable.push([...tr.querySelectorAll('input')].map(inp => inp.value));
    });

    // plbLoadTable
    ui.plbLoadTable = [];
    document.querySelectorAll('#plbLoadTable tbody tr').forEach(tr => {
        ui.plbLoadTable.push([...tr.querySelectorAll('input')].map(inp => inp.value));
    });

    // 월별 손실율
    ui.heatLossRates = [..._heatLossRates];

    // 단가산정 월별 열량단가
    ui.pricingHeatTable = [];
    document.querySelectorAll('#pricingHeatTable tbody tr').forEach(tr => {
        ui.pricingHeatTable.push([...tr.querySelectorAll('input')].map(inp => inp.value));
    });

    // 외부수열원 레지스트리
    ui.extSources = JSON.parse(JSON.stringify(EXT_SOURCES));

    // 월별 수요 보정계수
    ui.monthlyDemandFactors = [];
    document.querySelectorAll('#monthlyDemandFactors input').forEach(inp => {
        ui.monthlyDemandFactors.push(inp.value);
    });

    // 월별 평균 전력수요 (kW)
    ui.monthlyDemandKW = [];
    document.querySelectorAll('#monthlyDemandKW input').forEach(inp => {
        ui.monthlyDemandKW.push(inp.value);
    });

    // 24시간 수요 패턴
    ui.hourlyDemandPattern = [];
    document.querySelectorAll('#hourlyDemandPattern input').forEach(inp => {
        ui.hourlyDemandPattern.push(inp.value);
    });

    return ui;
}

function applySettings(ui) {
    if (!ui) return;
    SETTINGS_IDS.forEach(id => {
        const el = document.getElementById(id);
        if (el && ui[id] !== undefined) el.value = ui[id];
    });

    if (ui.loadTable) {
        const rows = document.querySelectorAll('#loadTable tbody tr');
        // 열 수가 다르면 구버전 데이터 → 무시 (HTML 기본값 유지)
        const expectedCols = rows[0] ? rows[0].querySelectorAll('input').length : 0;
        if (ui.loadTable[0] && ui.loadTable[0].length === expectedCols) {
            ui.loadTable.forEach((vals, i) => {
                if (!rows[i]) return;
                const inputs = rows[i].querySelectorAll('input');
                vals.forEach((v, j) => { if (inputs[j]) inputs[j].value = v; });
            });
        }
    }

    // plbLoadTable 복원
    if (ui.plbLoadTable) {
        const rows = document.querySelectorAll('#plbLoadTable tbody tr');
        const expectedCols = rows[0] ? rows[0].querySelectorAll('input').length : 0;
        if (ui.plbLoadTable[0] && ui.plbLoadTable[0].length === expectedCols) {
            ui.plbLoadTable.forEach((vals, i) => {
                if (!rows[i]) return;
                const inputs = rows[i].querySelectorAll('input');
                vals.forEach((v, j) => { if (inputs[j]) inputs[j].value = v; });
            });
        }
    }

    // 월별 손실율 복원
    if (ui.heatLossRates && Array.isArray(ui.heatLossRates)) {
        _heatLossRates = ui.heatLossRates.map(v => Number(v) || 0);
        while (_heatLossRates.length < 12) _heatLossRates.push(5);
    }

    // 단가산정 열량단가 복원
    if (ui.pricingHeatTable) {
        const rows = document.querySelectorAll('#pricingHeatTable tbody tr');
        ui.pricingHeatTable.forEach((vals, i) => {
            if (!rows[i]) return;
            const inputs = rows[i].querySelectorAll('input');
            vals.forEach((v, j) => { if (inputs[j]) inputs[j].value = v; });
        });
        updatePricingCalc();
    }

    // 외부수열원 레지스트리 복원
    if (ui.extSources && Array.isArray(ui.extSources)) {
        EXT_SOURCES = ui.extSources;
    } else {
        EXT_SOURCES = JSON.parse(JSON.stringify(DEFAULT_EXT_SOURCES));
    }
    rebuildTemplateHeaders();
    if (typeof renderExtSourcesUI === 'function') renderExtSourcesUI();

    // 월별 수요 보정계수 복원
    if (ui.monthlyDemandFactors && Array.isArray(ui.monthlyDemandFactors)) {
        const inputs = document.querySelectorAll('#monthlyDemandFactors input');
        ui.monthlyDemandFactors.forEach((v, i) => { if (inputs[i]) inputs[i].value = v; });
    }

    // 월별 평균 전력수요 복원
    if (ui.monthlyDemandKW && Array.isArray(ui.monthlyDemandKW)) {
        const inputs = document.querySelectorAll('#monthlyDemandKW input');
        ui.monthlyDemandKW.forEach((v, i) => { if (inputs[i]) inputs[i].value = v; });
    }

    // 24시간 수요 패턴 복원
    if (ui.hourlyDemandPattern && Array.isArray(ui.hourlyDemandPattern)) {
        const inputs = document.querySelectorAll('#hourlyDemandPattern input');
        ui.hourlyDemandPattern.forEach((v, i) => { if (inputs[i]) inputs[i].value = v; });
    }
}

let _saveTimer = null;
function scheduleSettingsSave() {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(saveSettingsNow, 800);
}

async function saveSettingsNow() {
    if (!CURRENT_PROJECT) return;
    const data = { year: CURRENT_YEAR, ui: collectSettings() };
    await pywebview.api.save_csv(CURRENT_PROJECT, '_settings', JSON.stringify(data));
    showPricingSaveStatus('설정 저장됨');
}

/* ═══════════════════════════════════════════════
   외부수열원 관리 UI
   ═══════════════════════════════════════════════ */
function renderExtSourcesUI() {
    const el = document.getElementById('extSourcesTable');
    if (!el) return;
    let html = '<table class="settings-tbl ext-src-tbl"><thead><tr>';
    html += '<th>이름</th><th>단가 컬럼</th><th>유형</th><th>활성</th>';
    html += '</tr></thead><tbody>';
    EXT_SOURCES.forEach((s, i) => {
        const typeLabel = s.type === 'dual' ? '양방향' : '공급';
        const checked = s.enabled ? 'checked' : '';
        html += `<tr>`;
        html += `<td style="font-weight:500;color:#d0d0d0">${s.opsName}</td>`;
        html += `<td style="color:#aaa">${s.priceName}${s.salePriceName ? ' / ' + s.salePriceName : ''}</td>`;
        html += `<td><span class="ext-type-badge ext-type-${s.type}">${typeLabel}</span></td>`;
        html += `<td><label class="ext-toggle"><input type="checkbox" ${checked} onchange="toggleExtSource('${s.id}')"><span class="ext-toggle-slider"></span></label></td>`;
        // 삭제 버튼 제거 — 데이터 탭에서 추가, 설정에서는 활성/비활성만
        html += '</tr>';
    });
    html += '</tbody></table>';
    el.innerHTML = html;
}

function addExtSource() {
    const nameEl = document.getElementById('newExtName');
    const priceEl = document.getElementById('newExtPriceName');
    const typeEl = document.getElementById('newExtType');
    const name = (nameEl.value || '').trim();
    if (!name) { alert('열원 이름을 입력하세요'); return; }
    if (EXT_SOURCES.some(s => s.opsName === name)) { alert('이미 존재하는 열원입니다'); return; }
    const priceName = (priceEl.value || '').trim() || name;
    const type = typeEl.value;
    const id = 'ext_' + Date.now();
    const src = { id, opsName: name, priceName, type, enabled: true };
    if (type === 'dual') src.salePriceName = name + '(판매)';
    EXT_SOURCES.push(src);
    rebuildTemplateHeaders();
    renderExtSourcesUI();
    renderPreview('operations_daily');
    renderPreview('sale_prices_daily');
    scheduleSettingsSave();
    nameEl.value = ''; priceEl.value = '';
}

function renameExtSource(oldName) {
    const src = EXT_SOURCES.find(s => s.opsName === oldName);
    if (!src) return;
    const newName = prompt(`"${oldName}" → 새 이름:`, oldName);
    if (!newName || newName.trim() === '' || newName.trim() === oldName) return;
    const trimmed = newName.trim();
    if (EXT_SOURCES.some(s => s.opsName === trimmed && s.id !== src.id)) { alert('이미 존재하는 이름입니다'); return; }

    const oldOps = src.opsName;
    const oldPrice = src.priceName;
    const oldSale = src.salePriceName;

    // EXT_SOURCES 업데이트
    src.opsName = trimmed;
    // priceName이 oldOps와 같았으면 함께 변경 (커스텀 단가명이면 유지)
    if (src.priceName === oldOps) src.priceName = trimmed;
    if (src.type === 'dual' && src.salePriceName === oldOps + '(판매)') src.salePriceName = trimmed + '(판매)';

    // operations_daily 헤더 변경
    if (DATA.operations_daily) {
        const idx = DATA.operations_daily.headers.indexOf(oldOps);
        if (idx >= 0) DATA.operations_daily.headers[idx] = trimmed;
    }
    // sale_prices_daily 헤더 변경
    if (DATA.sale_prices_daily) {
        const idx = DATA.sale_prices_daily.headers.indexOf(oldPrice);
        if (idx >= 0) DATA.sale_prices_daily.headers[idx] = src.priceName;
        if (oldSale) {
            const idx2 = DATA.sale_prices_daily.headers.indexOf(oldSale);
            if (idx2 >= 0) DATA.sale_prices_daily.headers[idx2] = src.salePriceName;
        }
    }

    rebuildTemplateHeaders();
    renderExtSourcesUI();
    renderPreview('operations_daily');
    renderPreview('sale_prices_daily');
    scheduleSettingsSave();
}

function removeExtSource(id) {
    const src = EXT_SOURCES.find(s => s.id === id);
    if (!src) return;
    if (!confirm(`"${src.opsName}" 수열원을 삭제하시겠습니까?\n관련 데이터 열도 사용되지 않게 됩니다.`)) return;
    EXT_SOURCES = EXT_SOURCES.filter(s => s.id !== id);
    rebuildTemplateHeaders();
    renderExtSourcesUI();
    scheduleSettingsSave();
}

function toggleExtSource(id) {
    const src = EXT_SOURCES.find(s => s.id === id);
    if (src) {
        src.enabled = !src.enabled;
        scheduleSettingsSave();
    }
}

// 차트 Y축 깔끔한 스케일: 천단위→100, 백단위→50 끊기, 정수만
function niceMax(val) {
    if (val <= 0) return 100;
    if (val >= 1000) return Math.ceil(val / 100) * 100;
    if (val >= 100) return Math.ceil(val / 50) * 50;
    return Math.ceil(val / 10) * 10;
}
function niceStep(max) {
    if (max >= 2000) return 200;
    if (max >= 1000) return 100;
    if (max >= 200) return 50;
    return 10;
}

// 쉼표 포함 숫자 파싱 (예: "76,426" → 76426)
function parseNum(v) {
    if (typeof v === 'number') return v;
    return parseFloat(String(v).replace(/,/g, '')) || 0;
}

// 단가산정: NLPC / 무부하비 / 무부하비정산 자동계산
function updatePricingCalc() {
    const con = parseNum(document.getElementById('pricingConst')?.value);
    document.querySelectorAll('#pricingHeatTable tbody tr').forEach(tr => {
        const inp = tr.querySelector('input');
        const nlpcEl = tr.querySelector('td.nlpc');
        const nlcEl = tr.querySelector('td.nlc');
        const nlsEl = tr.querySelector('td.nls');
        if (!inp) return;
        const heatPrice = parseNum(inp.value);
        if (heatPrice === 0) {
            if (nlpcEl) nlpcEl.textContent = '-';
            if (nlcEl) nlcEl.textContent = '-';
            if (nlsEl) nlsEl.textContent = '-';
            return;
        }
        // NLPC = 상수 × 열량단가
        const nlpc = con * heatPrice;
        // 무부하비 = NLPC / 22000 (22MW × 1000kW) — 표시용 별도 반올림
        const noLoadCost = Math.round(nlpc / 22000 * 100) / 100;
        // 무부하비정산 = (상수×열량단가)/22000 × 1/(1+2.6327) — 원본에서 직접 계산
        const noLoadSettle = Math.round(nlpc / 22000 * (1 / (1 + 2.6327)) * 100) / 100;
        if (nlpcEl) nlpcEl.textContent = Math.round(nlpc).toLocaleString();
        if (nlcEl) nlcEl.textContent = noLoadCost.toFixed(2);
        if (nlsEl) nlsEl.textContent = noLoadSettle.toFixed(2);
    });
}

// 저장 피드백 표시
function showPricingSaveStatus(msg) {
    const el = document.getElementById('pricingSaveStatus');
    if (!el) return;
    el.textContent = msg || '저장됨';
    el.style.opacity = '1';
    clearTimeout(el._timer);
    el._timer = setTimeout(() => { el.style.opacity = '0'; }, 2000);
}

/* ═══════════════════════════════════════════════
   데이터 산출: MP(시장가격) 계산
   ═══════════════════════════════════════════════ */
function calculateMP() {
    const statusEl = document.getElementById('calcStatus');
    const show = msg => { if (statusEl) statusEl.textContent = msg; };

    // 1. 필수 데이터 검증
    if (!DATA.smp_hourly) { show('SMP 시간별 데이터가 필요합니다'); return; }
    if (!DATA.cp_monthly) { show('CP/TLF 데이터가 필요합니다'); return; }
    show('계산 중...');

    // 2. 공휴일 Set
    const holidaySet = new Set();
    if (DATA.holidays) {
        DATA.holidays.rows.forEach(row => {
            if (Number(row[3]) === 1) holidaySet.add(`${Number(row[0])}-${Number(row[1])}`);
        });
    }

    // 3. TLF 맵 (월별)
    const findCol = (headers, name) => headers.findIndex(h => h === name);
    const cpH = DATA.cp_monthly.headers;
    const tlfWdIdx = findCol(cpH, 'TLF(평일)');
    const tlfSatIdx = findCol(cpH, 'TLF(토요일)');
    const tlfHolIdx = findCol(cpH, 'TLF(공휴일)');
    const tlfMap = {};
    DATA.cp_monthly.rows.forEach(row => {
        const mm = Number(row[0]);
        if (mm >= 1 && mm <= 12) {
            tlfMap[mm] = {
                weekday: tlfWdIdx >= 0 ? (Number(row[tlfWdIdx]) || 1) : 1,
                saturday: tlfSatIdx >= 0 ? (Number(row[tlfSatIdx]) || 1) : 1,
                holiday: tlfHolIdx >= 0 ? (Number(row[tlfHolIdx]) || 1) : 1,
            };
        }
    });

    // 4. 무부하비정산 월별 (단가산정 설정)
    const noLoadSettle = Array(12).fill(0);
    const pricingConst = parseNum(document.getElementById('pricingConst')?.value);
    document.querySelectorAll('#pricingHeatTable tbody tr').forEach(tr => {
        const mm = parseInt(tr.dataset.m) - 1;
        if (mm >= 0 && mm < 12) {
            const hp = parseNum(tr.querySelector('input')?.value);
            if (hp > 0 && pricingConst > 0) {
                noLoadSettle[mm] = Math.round((pricingConst * hp) / 22000 * (1 / (1 + 2.6327)) * 100) / 100;
            }
        }
    });

    // 5. 365일 MP 생성
    const year = CURRENT_YEAR;
    const isLeap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
    const daysInMonth = [31, isLeap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    const dowNames = ['일','월','화','수','목','금','토'];
    const mpRows = [];

    for (let m = 0; m < 12; m++) {
        for (let d = 1; d <= daysInMonth[m]; d++) {
            const dt = new Date(year, m, d);
            const dow = dt.getDay();
            const key = `${m + 1}-${d}`;
            const isSaturday = dow === 6;
            const isHoliday = holidaySet.has(key);
            const isHolidayOrSunday = dow === 0 || isHoliday;

            // 요일 구분
            let dayType = '평일', dtClass = 'wd';
            if (dow === 0) { dayType = '일요일'; dtClass = 'sun'; }
            else if (isHoliday) { dayType = '공휴일'; dtClass = 'hol'; }
            else if (isSaturday) { dayType = '토요일'; dtClass = 'sat'; }

            // TLF 선택
            const tlfEntry = tlfMap[m + 1] || { weekday: 1, saturday: 1, holiday: 1 };
            const tlf = isHolidayOrSunday ? tlfEntry.holiday
                      : isSaturday ? tlfEntry.saturday
                      : tlfEntry.weekday;
            const nls = noLoadSettle[m];

            // SMP 시간별
            const smpH = Array(24).fill(0);
            const smpRow = DATA.smp_hourly.rows.find(r => Number(r[0]) === (m + 1) && Number(r[1]) === d);
            if (smpRow) {
                for (let h = 0; h < 24; h++) smpH[h] = Number(smpRow[h + 2]) || 0;
            }

            // MP = ROUND(SMP × TLF, 2) + 무부하비정산
            const mpH = smpH.map(v => Math.round(v * tlf * 100) / 100 + nls);
            const mpAvg = Math.round(mpH.reduce((s, v) => s + v, 0) / 24 * 100) / 100;

            mpRows.push({ month: m + 1, day: d, dowName: dowNames[dow], dayType, dtClass, tlf, nls, mpH, mpAvg });
        }
    }

    // 6. 렌더링
    renderMPTable(mpRows);
    show(`${year}년 ${mpRows.length}일 MP 산출 완료`);
}

function renderMPTable(mpRows) {
    const el = document.getElementById('pv_calc_mp');
    if (!el) return;

    let html = '<table class="pv-tbl"><thead><tr>';
    html += '<th style="position:sticky;left:0;z-index:2;background:#f1f5f9">월</th>';
    html += '<th style="position:sticky;left:28px;z-index:2;background:#f1f5f9">일</th>';
    html += '<th>요일</th><th>구분</th><th>TLF</th><th>무부하비정산</th><th>MP평균</th>';
    for (let h = 1; h <= 24; h++) html += `<th>${h}h</th>`;
    html += '</tr></thead><tbody>';

    const dtColors = { wd: '#1e293b', sat: '#d97706', sun: '#dc2626', hol: '#7c3aed' };

    for (const r of mpRows) {
        const c = dtColors[r.dtClass] || '#1e293b';
        html += '<tr>';
        html += `<td style="position:sticky;left:0;background:#fff;font-weight:600">${r.month}</td>`;
        html += `<td style="position:sticky;left:28px;background:#fff">${r.day}</td>`;
        html += `<td style="color:${c}">${r.dowName}</td>`;
        html += `<td style="color:${c};font-weight:600">${r.dayType}</td>`;
        html += `<td>${r.tlf.toFixed(4)}</td>`;
        html += `<td>${r.nls.toFixed(2)}</td>`;
        html += `<td style="font-weight:600;color:#0369a1">${r.mpAvg.toFixed(2)}</td>`;
        for (let h = 0; h < 24; h++) {
            html += `<td>${r.mpH[h].toFixed(2)}</td>`;
        }
        html += '</tr>';
    }

    html += '</tbody></table>';
    el.innerHTML = html;
}

/* ═══════════════════════════════════════════════
   데이터 산출: 용량요금 (일별 CP)
   ═══════════════════════════════════════════════ */
function calculateCapacityCharge() {
    const el = document.getElementById('pv_calc_cp');
    if (!el) return;
    if (!DATA.cp_monthly) { el.innerHTML = '<p class="hint">CP/TLF 데이터가 필요합니다</p>'; return; }

    const cpH = DATA.cp_monthly.headers;
    const cpWdIdx = cpH.indexOf('평일CP');
    const cpHolIdx = cpH.indexOf('휴일CP');
    if (cpWdIdx < 0) { el.innerHTML = '<p class="hint">평일CP 컬럼이 없습니다</p>'; return; }

    const cpMap = {};
    DATA.cp_monthly.rows.forEach(row => {
        const mm = Number(row[0]);
        if (mm >= 1 && mm <= 12) cpMap[mm] = { weekday: Number(row[cpWdIdx]) || 0, holiday: cpHolIdx >= 0 ? (Number(row[cpHolIdx]) || 0) : 0 };
    });

    const holidaySet = new Set();
    if (DATA.holidays) DATA.holidays.rows.forEach(row => { if (Number(row[3]) === 1) holidaySet.add(`${Number(row[0])}-${Number(row[1])}`); });

    const year = CURRENT_YEAR;
    const isLeap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
    const daysInMonth = [31, isLeap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    const dowNames = ['일','월','화','수','목','금','토'];
    const rows = [];

    for (let m = 0; m < 12; m++) {
        for (let d = 1; d <= daysInMonth[m]; d++) {
            const dt = new Date(year, m, d);
            const dow = dt.getDay();
            const isHol = dow === 0 || dow === 6 || holidaySet.has(`${m+1}-${d}`);
            let dayType = '평일', dtClass = 'wd';
            if (dow === 0) { dayType = '일요일'; dtClass = 'sun'; }
            else if (holidaySet.has(`${m+1}-${d}`)) { dayType = '공휴일'; dtClass = 'hol'; }
            else if (dow === 6) { dayType = '토요일'; dtClass = 'sat'; }
            const cp = cpMap[m + 1] || { weekday: 0, holiday: 0 };
            const cpVal = isHol ? cp.holiday : cp.weekday;
            rows.push({ month: m + 1, day: d, dowName: dowNames[dow], dayType, dtClass, cpVal, isHol });
        }
    }

    // 월별 합계
    const monthSum = Array(12).fill(0);
    const monthWd = Array(12).fill(0);
    const monthHol = Array(12).fill(0);
    rows.forEach(r => { monthSum[r.month - 1] += r.cpVal; if (r.isHol) monthHol[r.month - 1]++; else monthWd[r.month - 1]++; });

    const dtColors = { wd: '#1e293b', sat: '#d97706', sun: '#dc2626', hol: '#7c3aed' };
    let html = '<p style="font-size:11px;color:#64748b;margin-bottom:8px">일별 용량요금 = 평일→평일CP, 토·일·공휴일→휴일CP</p>';
    html += '<table class="pv-tbl"><thead><tr>';
    html += '<th style="position:sticky;left:0;z-index:2;background:#f1f5f9">월</th>';
    html += '<th style="position:sticky;left:28px;z-index:2;background:#f1f5f9">일</th>';
    html += '<th>요일</th><th>구분</th><th>CP(원)</th>';
    html += '</tr></thead><tbody>';

    let prevMonth = 0;
    for (const r of rows) {
        if (r.month !== prevMonth && prevMonth > 0) {
            const ms = monthSum[prevMonth - 1];
            html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
            html += `<td>${monthWd[prevMonth - 1]}일</td><td>${monthHol[prevMonth - 1]}일</td>`;
            html += `<td style="color:#1d4ed8">${Math.round(ms).toLocaleString()}</td></tr>`;
        }
        prevMonth = r.month;
        const c = dtColors[r.dtClass];
        html += `<tr><td style="position:sticky;left:0;background:#fff;font-weight:600">${r.month}</td>`;
        html += `<td style="position:sticky;left:28px;background:#fff">${r.day}</td>`;
        html += `<td style="color:${c}">${r.dowName}</td><td style="color:${c};font-weight:600">${r.dayType}</td>`;
        html += `<td style="font-weight:600">${Math.round(r.cpVal).toLocaleString()}</td></tr>`;
    }
    if (prevMonth > 0) {
        const ms = monthSum[prevMonth - 1];
        html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
        html += `<td>${monthWd[prevMonth - 1]}일</td><td>${monthHol[prevMonth - 1]}일</td>`;
        html += `<td style="color:#1d4ed8">${Math.round(ms).toLocaleString()}</td></tr>`;
    }
    const total = monthSum.reduce((a, b) => a + b, 0);
    html += `<tr style="border-top:2px solid #1e293b;font-weight:800"><td colspan="4">연간 합계</td>`;
    html += `<td style="color:#1d4ed8">${Math.round(total).toLocaleString()}</td></tr>`;
    html += '</tbody></table>';
    el.innerHTML = html;
}

/* ═══════════════════════════════════════════════
   데이터 산출: 외부수열비 (일별)
   ═══════════════════════════════════════════════ */
function calculateExternalHeatCost() {
    const el = document.getElementById('pv_calc_ext_heat');
    if (!el) return;
    if (!DATA.operations_daily) { el.innerHTML = '<p class="hint">운영실적 데이터가 필요합니다</p>'; return; }
    if (!DATA.sale_prices_daily) { el.innerHTML = '<p class="hint">판매단가 데이터가 필요합니다</p>'; return; }

    const oH = DATA.operations_daily.headers;
    const pH = DATA.sale_prices_daily.headers;

    // 외부수열 매핑: EXT_SOURCES 레지스트리 기반 (활성 소스만)
    const extMap = EXT_SOURCES.filter(s => s.enabled).map(s => ({ opsName: s.opsName, priceName: s.priceName }));

    // 컬럼 인덱스 찾기
    const sources = extMap.map(e => ({
        name: e.opsName,
        opsCol: oH.indexOf(e.opsName),
        priceCol: pH.indexOf(e.priceName),
    })).filter(s => s.opsCol >= 0);

    if (sources.length === 0) { el.innerHTML = '<p class="hint">외부수열 컬럼을 찾을 수 없습니다</p>'; return; }

    // 판매단가를 일별로 매핑 {m-d: {sourceName: price}}
    const priceMap = {};
    DATA.sale_prices_daily.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        const key = `${m}-${d}`;
        priceMap[key] = {};
        sources.forEach(s => { if (s.priceCol >= 0) priceMap[key][s.name] = parseNum(row[s.priceCol]); });
    });

    // 일별 비용 계산
    const resultRows = [];
    DATA.operations_daily.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        if (m < 1 || m > 12) return;
        const key = `${m}-${d}`;
        const prices = priceMap[key] || {};
        const items = {};
        let dayTotal = 0;
        sources.forEach(s => {
            const qty = parseNum(row[s.opsCol]);
            const price = prices[s.name] || 0;
            const cost = Math.round(qty * price);
            items[s.name] = { qty, price, cost };
            dayTotal += cost;
        });
        resultRows.push({ month: m, day: d, items, total: dayTotal });
    });

    // 월별 합계
    const monthTotals = Array.from({ length: 12 }, () => {
        const obj = { total: 0 };
        sources.forEach(s => obj[s.name] = 0);
        return obj;
    });
    resultRows.forEach(r => {
        const mb = monthTotals[r.month - 1];
        sources.forEach(s => mb[s.name] += r.items[s.name]?.cost || 0);
        mb.total += r.total;
    });

    // 렌더링
    let html = '<p style="font-size:11px;color:#64748b;margin-bottom:8px">외부수열비 = 수열량(Gcal) × 수열단가(원/Gcal)</p>';
    html += '<div style="overflow:auto;max-height:calc(100vh - 260px)">';
    html += '<table class="pv-tbl"><thead><tr>';
    html += '<th>월</th><th>일</th>';
    sources.forEach(s => html += `<th colspan="2" style="background:#ecfdf5">${s.name}</th>`);
    html += '<th style="background:#d1fae5;font-weight:800">합계(원)</th>';
    html += '</tr><tr><th></th><th></th>';
    sources.forEach(() => { html += '<th style="font-size:9px;background:#f0fdf4">Gcal</th><th style="font-size:9px;background:#f0fdf4">비용</th>'; });
    html += '<th></th></tr></thead><tbody>';

    let prevMonth = 0;
    for (const r of resultRows) {
        if (r.month !== prevMonth && prevMonth > 0) {
            const mt = monthTotals[prevMonth - 1];
            html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
            sources.forEach(s => { html += `<td></td><td>${Math.round(mt[s.name]).toLocaleString()}</td>`; });
            html += `<td style="color:#065f46;font-weight:800">${Math.round(mt.total).toLocaleString()}</td></tr>`;
        }
        prevMonth = r.month;
        html += `<tr><td>${r.month}</td><td>${r.day}</td>`;
        sources.forEach(s => {
            const it = r.items[s.name] || { qty: 0, cost: 0 };
            html += `<td style="color:#64748b">${it.qty ? fmtNum(Math.round(it.qty)) : '-'}</td>`;
            html += `<td>${it.cost ? Math.round(it.cost).toLocaleString() : '-'}</td>`;
        });
        html += `<td style="font-weight:600">${Math.round(r.total).toLocaleString()}</td></tr>`;
    }
    if (prevMonth > 0) {
        const mt = monthTotals[prevMonth - 1];
        html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
        sources.forEach(s => { html += `<td></td><td>${Math.round(mt[s.name]).toLocaleString()}</td>`; });
        html += `<td style="color:#065f46;font-weight:800">${Math.round(mt.total).toLocaleString()}</td></tr>`;
    }
    const grandTotal = monthTotals.reduce((s, m) => s + m.total, 0);
    html += `<tr style="border-top:2px solid #1e293b;font-weight:800"><td colspan="2">연간 합계</td>`;
    sources.forEach(s => { const st = monthTotals.reduce((sum, m) => sum + m[s.name], 0); html += `<td></td><td>${Math.round(st).toLocaleString()}</td>`; });
    html += `<td style="color:#065f46">${Math.round(grandTotal).toLocaleString()}</td></tr>`;
    html += '</tbody></table></div>';
    el.innerHTML = html;
}

/* ═══════════════════════════════════════════════
   데이터 산출: 열판매매출 (일별)
   ═══════════════════════════════════════════════ */
function calculateHeatSalesRevenue() {
    const el = document.getElementById('pv_calc_heat_sales');
    if (!el) return;
    if (!DATA.sales_hourly && !DATA.sales_daily) { el.innerHTML = '<p class="hint">판매실적 데이터가 필요합니다</p>'; return; }

    const findCol = (headers, name) => headers.findIndex(h => h === name);

    // 판매 항목 정의 (sales_hourly 기반)
    const heatItemDefs = [
        { name: '주택용', key: '주택용난방', rateKey: '주택용' },
        { name: '업무용난방', key: '업무용난방', rateKey: '업무용' },
        { name: '업무용냉방', key: '업무용냉방', rateKey: '업무용' },
        { name: '공공용난방', key: '공공용난방', rateKey: '공공용' },
        { name: '공공용냉방', key: '공공용냉방', rateKey: '공공용' },
    ];

    // sales_hourly → 일별 항목별 합산
    const dailyHeatQty = {}; // key: "m-d", value: { 주택용난방: qty, ... }
    if (DATA.sales_hourly) {
        const shH = DATA.sales_hourly.headers;
        const colMap = {};
        heatItemDefs.forEach(def => { colMap[def.key] = findCol(shH, def.key); });
        DATA.sales_hourly.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            if (!dailyHeatQty[key]) dailyHeatQty[key] = {};
            heatItemDefs.forEach(def => {
                const ci = colMap[def.key];
                if (ci >= 0) dailyHeatQty[key][def.key] = (dailyHeatQty[key][def.key] || 0) + (Number(row[ci]) || 0);
            });
        });
    }

    // sales_daily → 인테코 일별
    const intecoQtyMap = {};
    if (DATA.sales_daily) {
        const sdH = DATA.sales_daily.headers;
        const intecoCol = findCol(sdH, '인테코');
        if (intecoCol >= 0) {
            DATA.sales_daily.rows.forEach(row => {
                const key = `${Number(row[0])}-${Number(row[1])}`;
                intecoQtyMap[key] = Number(row[intecoCol]) || 0;
            });
        }
    }

    // 열요금표 (월별 원/Gcal) - heat_rate_table
    const heatRateMap = {};
    if (DATA.heat_rate_table) {
        const hrH = DATA.heat_rate_table.headers;
        DATA.heat_rate_table.rows.forEach(row => {
            const m = Number(row[0]);
            const entry = {};
            hrH.forEach((h, i) => { if (i > 0) entry[h] = parseNum(row[i]); });
            heatRateMap[m] = entry;
        });
    }

    // 인테코 판매단가 (일별) - sale_prices_daily
    const intecoPriceMap = {};
    if (DATA.sale_prices_daily) {
        const prH = DATA.sale_prices_daily.headers;
        const intecoPriceCol = prH.indexOf('인테코(판매)');
        if (intecoPriceCol >= 0) {
            DATA.sale_prices_daily.rows.forEach(row => {
                const m = Number(row[0]), d = Number(row[1]);
                intecoPriceMap[`${m}-${d}`] = parseNum(row[intecoPriceCol]);
            });
        }
    }

    // 기본료 (월별)
    const baseFeeMap = {};
    if (DATA.heat_rate_table) {
        const bfIdx = DATA.heat_rate_table.headers.indexOf('기본료');
        if (bfIdx >= 0) DATA.heat_rate_table.rows.forEach(row => { baseFeeMap[Number(row[0])] = parseNum(row[bfIdx]); });
    }

    // 일별 키 수집 (sales_hourly + sales_daily 합집합)
    const allDayKeys = new Set([...Object.keys(dailyHeatQty), ...Object.keys(intecoQtyMap)]);

    // 일별 계산
    const resultRows = [];
    allDayKeys.forEach(key => {
        const [ms, ds] = key.split('-');
        const m = Number(ms), d = Number(ds);
        if (m < 1 || m > 12) return;
        const rates = heatRateMap[m] || {};
        const items = {};
        let dayTotal = 0;
        const hq = dailyHeatQty[key] || {};

        // 열 판매 항목
        heatItemDefs.forEach(def => {
            const qty = hq[def.key] || 0;
            const price = rates[def.rateKey] || 0;
            const rev = Math.round(qty * price);
            items[def.name] = { qty, price, rev };
            dayTotal += rev;
        });

        // 인테코
        const intecoQty = intecoQtyMap[key] || 0;
        const intecoPrice = intecoPriceMap[key] || 0;
        const intecoRev = Math.round(intecoQty * intecoPrice);
        items['인테코'] = { qty: intecoQty, price: intecoPrice, rev: intecoRev };
        dayTotal += intecoRev;

        resultRows.push({ month: m, day: d, items, total: dayTotal });
    });
    resultRows.sort((a, b) => a.month !== b.month ? a.month - b.month : a.day - b.day);

    // 렌더링용 컬럼 목록
    const saleCols = [...heatItemDefs.map(d => ({ name: d.name })), { name: '인테코' }];

    // 월별 합계
    const monthTotals = Array.from({ length: 12 }, () => {
        const obj = { total: 0, baseFee: 0 };
        saleCols.forEach(s => obj[s.name] = 0);
        return obj;
    });
    resultRows.forEach(r => {
        const mb = monthTotals[r.month - 1];
        saleCols.forEach(s => mb[s.name] += r.items[s.name]?.rev || 0);
        mb.total += r.total;
    });
    for (let m = 1; m <= 12; m++) monthTotals[m - 1].baseFee = baseFeeMap[m] || 0;

    // 렌더링
    let html = '<p style="font-size:11px;color:#64748b;margin-bottom:8px">열판매매출 = 판매량(Gcal) × 단가(원/Gcal) + 월별 기본료</p>';
    html += '<div style="overflow:auto;max-height:calc(100vh - 260px)">';
    html += '<table class="pv-tbl"><thead><tr>';
    html += '<th>월</th><th>일</th>';
    saleCols.forEach(s => html += `<th colspan="2" style="background:#fffbeb">${s.name}</th>`);
    html += '<th style="background:#fef3c7;font-weight:800">일합계(원)</th>';
    html += '</tr><tr><th></th><th></th>';
    saleCols.forEach(() => { html += '<th style="font-size:9px;background:#fffdf4">Gcal</th><th style="font-size:9px;background:#fffdf4">매출</th>'; });
    html += '<th></th></tr></thead><tbody>';

    let prevMonth = 0;
    for (const r of resultRows) {
        if (r.month !== prevMonth && prevMonth > 0) {
            const mt = monthTotals[prevMonth - 1];
            html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
            saleCols.forEach(s => { html += `<td></td><td>${Math.round(mt[s.name]).toLocaleString()}</td>`; });
            const mTotal = mt.total + mt.baseFee;
            html += `<td style="color:#92400e;font-weight:800">${Math.round(mTotal).toLocaleString()}</td></tr>`;
            if (mt.baseFee > 0) {
                html += `<tr style="background:#fffbeb"><td colspan="2" style="text-align:right;font-size:10px;color:#92400e">기본료 포함</td>`;
                saleCols.forEach(() => html += '<td colspan="2"></td>');
                html += `<td style="font-size:10px;color:#92400e">+${Math.round(mt.baseFee).toLocaleString()}</td></tr>`;
            }
        }
        prevMonth = r.month;
        html += `<tr><td>${r.month}</td><td>${r.day}</td>`;
        saleCols.forEach(s => {
            const it = r.items[s.name] || { qty: 0, rev: 0 };
            html += `<td style="color:#64748b">${it.qty ? fmtNum(Math.round(it.qty)) : '-'}</td>`;
            html += `<td>${it.rev ? Math.round(it.rev).toLocaleString() : '-'}</td>`;
        });
        html += `<td style="font-weight:600">${Math.round(r.total).toLocaleString()}</td></tr>`;
    }
    if (prevMonth > 0) {
        const mt = monthTotals[prevMonth - 1];
        html += `<tr style="background:#eff6ff;font-weight:700"><td colspan="2">${prevMonth}월 소계</td>`;
        saleCols.forEach(s => { html += `<td></td><td>${Math.round(mt[s.name]).toLocaleString()}</td>`; });
        const mTotal = mt.total + mt.baseFee;
        html += `<td style="color:#92400e;font-weight:800">${Math.round(mTotal).toLocaleString()}</td></tr>`;
        if (mt.baseFee > 0) {
            html += `<tr style="background:#fffbeb"><td colspan="2" style="text-align:right;font-size:10px;color:#92400e">기본료 포함</td>`;
            saleCols.forEach(() => html += '<td colspan="2"></td>');
            html += `<td style="font-size:10px;color:#92400e">+${Math.round(mt.baseFee).toLocaleString()}</td></tr>`;
        }
    }
    const grandTotal = monthTotals.reduce((s, m) => s + m.total + m.baseFee, 0);
    const grandBaseFee = monthTotals.reduce((s, m) => s + m.baseFee, 0);
    html += `<tr style="border-top:2px solid #1e293b;font-weight:800"><td colspan="2">연간 합계</td>`;
    saleCols.forEach(s => { const st = monthTotals.reduce((sum, m) => sum + m[s.name], 0); html += `<td></td><td>${Math.round(st).toLocaleString()}</td>`; });
    html += `<td style="color:#92400e">${Math.round(grandTotal).toLocaleString()}</td></tr>`;
    if (grandBaseFee > 0) {
        html += `<tr><td colspan="2" style="text-align:right;font-size:10px;color:#92400e">기본료 합계</td>`;
        saleCols.forEach(() => html += '<td colspan="2"></td>');
        html += `<td style="font-size:10px;color:#92400e">${Math.round(grandBaseFee).toLocaleString()}</td></tr>`;
    }
    html += '</tbody></table></div>';
    el.innerHTML = html;
}

/* ═══════════════════════════════════════════════
   데이터 산출: 전체 산출
   ═══════════════════════════════════════════════ */
function runAllCalc() {
    if (DATA.smp_hourly) calculateMP();
    calculateCapacityCharge();
    calculateExternalHeatCost();
    calculateHeatSalesRevenue();
    calculateWaterChem();
    window._calcDone = true;
    if (typeof checkPipelineStatus === 'function') checkPipelineStatus();
}

/* ═══════════════════════════════════════════════
   데이터 산출: 용수약품비용
   ═══════════════════════════════════════════════ */
function calculateWaterChem() {
    const statusEl = document.getElementById('calcStatus');
    const show = msg => { if (statusEl) statusEl.textContent = msg; };

    // 1. 용수약품 사용량 검증
    if (!DATA.water_chem_daily) {
        show('용수약품 사용량 데이터가 없습니다');
        return;
    }
    // 2. 수도단가, 약품단가 검증
    if (!DATA.water_price_monthly) {
        show('수도단가 데이터가 없습니다');
        return;
    }
    if (!DATA.chem_price_monthly) {
        show('약품단가 데이터가 없습니다');
        return;
    }

    const wcHeaders = TEMPLATES.water_chem_daily.headers; // 월,일,용수,가성소다,...
    const chemNames = wcHeaders.slice(2); // ['용수사용량(㎥)','가성소다(Kg)',...]

    // 3. 수도단가 → 월별 상수도단가 (300초과 기준 사용)
    const waterPriceByMonth = {};
    if (DATA.water_price_monthly.rows) {
        DATA.water_price_monthly.rows.forEach(r => {
            const m = Number(r[0]);
            // headers: 월, 정액요금, 물이용단가, 상수도단가(300이하), 상수도단가(300초과), 하수도...
            // 상수도(300초과) = index 3
            waterPriceByMonth[m] = parseNum(r[3]) || parseNum(r[2]) || 0;
        });
    }

    // 4. 약품단가 → 월별 {가성소다, 탈산소제, 암모니아수, 염산, 카보하이드라자이드, PH조절제}
    const chemPriceByMonth = {};
    if (DATA.chem_price_monthly.rows) {
        DATA.chem_price_monthly.rows.forEach(r => {
            const m = Number(r[0]);
            chemPriceByMonth[m] = {
                '가성소다': parseNum(r[1]),
                '탈산소제': parseNum(r[2]),
                '암모니아수': parseNum(r[3]),
                '염산': parseNum(r[4]),
                '카보하이드라자이드': parseNum(r[5]),
                'PH조절제': parseNum(r[6]),
            };
        });
    }

    // 5. 일별 비용 계산
    const resultRows = [];
    DATA.water_chem_daily.rows.forEach(row => {
        const m = Number(row[0]);
        const d = Number(row[1]);
        if (m < 1 || m > 12) return;

        const waterQty = parseNum(row[2]);
        const chemQtys = [parseNum(row[3]), parseNum(row[4]), parseNum(row[5]),
                          parseNum(row[6]), parseNum(row[7]), parseNum(row[8])];
        const chemKeys = ['가성소다', '탈산소제', '암모니아수', '염산', '카보하이드라자이드', 'PH조절제'];

        const waterPrice = waterPriceByMonth[m] || 0;
        const waterCost = Math.round(waterQty * waterPrice);

        const cp = chemPriceByMonth[m] || {};
        const chemCosts = chemKeys.map((k, i) => Math.round(chemQtys[i] * (cp[k] || 0)));
        const chemTotal = chemCosts.reduce((s, v) => s + v, 0);

        resultRows.push({
            month: m, day: d,
            waterQty, waterPrice, waterCost,
            chemQtys, chemCosts, chemTotal,
            total: waterCost + chemTotal
        });
    });

    renderWaterChemTable(resultRows);
    show(statusEl.textContent ? statusEl.textContent + ' | 용수약품비용 산출 완료' : '용수약품비용 산출 완료');
}

function renderWaterChemTable(rows) {
    const el = document.getElementById('pv_calc_water_chem');
    if (!el) return;

    const chemLabels = ['가성소다', '탈산소제', '암모니아수', '염산', '카보하이드라자이드', 'PH조절제'];

    let html = '<table class="pv-tbl"><thead>';
    html += '<tr>';
    html += '<th rowspan="2">월</th><th rowspan="2">일</th>';
    html += '<th colspan="3" style="background:#dbeafe;color:#1565c0">용수</th>';
    html += '<th colspan="' + (chemLabels.length * 2) + '" style="background:#fef3c7;color:#e65100">약품</th>';
    html += '<th rowspan="2" style="background:#fce4ec;color:#880e4f;font-weight:800">합계<br>(원)</th>';
    html += '</tr><tr>';
    // 용수 서브헤더
    html += '<th style="background:#eff6ff">사용량<br>(㎥)</th>';
    html += '<th style="background:#eff6ff">단가<br>(원/㎥)</th>';
    html += '<th style="background:#eff6ff">비용<br>(원)</th>';
    // 약품 서브헤더
    for (const name of chemLabels) {
        html += `<th style="background:#fffbeb">${name}<br>(Kg)</th>`;
        html += `<th style="background:#fffbeb">비용<br>(원)</th>`;
    }
    html += '</tr></thead><tbody>';

    // 월별 소계
    const monthTotals = {};

    for (const r of rows) {
        if (!monthTotals[r.month]) monthTotals[r.month] = { waterCost: 0, chemCosts: Array(6).fill(0), total: 0 };
        const mt = monthTotals[r.month];
        mt.waterCost += r.waterCost;
        r.chemCosts.forEach((c, i) => mt.chemCosts[i] += c);
        mt.total += r.total;

        html += '<tr>';
        html += `<td style="text-align:center;color:#64748b">${r.month}</td>`;
        html += `<td style="text-align:center;color:#64748b">${r.day}</td>`;
        html += `<td>${fmtNum(r.waterQty)}</td>`;
        html += `<td>${fmtNum(r.waterPrice)}</td>`;
        html += `<td style="color:#1565c0;font-weight:600">${fmtNum(r.waterCost)}</td>`;
        for (let i = 0; i < 6; i++) {
            html += `<td>${fmtNum(r.chemQtys[i])}</td>`;
            html += `<td style="color:#e65100">${fmtNum(r.chemCosts[i])}</td>`;
        }
        html += `<td style="font-weight:700;color:#c62828">${fmtNum(r.total)}</td>`;
        html += '</tr>';
    }

    // 합계 행
    let grandWater = 0, grandChem = Array(6).fill(0), grandTotal = 0;
    Object.values(monthTotals).forEach(mt => {
        grandWater += mt.waterCost;
        mt.chemCosts.forEach((c, i) => grandChem[i] += c);
        grandTotal += mt.total;
    });

    html += '<tr style="background:#f1f5f9;font-weight:700;border-top:2px solid #94a3b8">';
    html += '<td colspan="2" style="text-align:center">합계</td>';
    html += '<td></td><td></td>';
    html += `<td style="color:#1565c0">${fmtNum(grandWater)}</td>`;
    for (let i = 0; i < 6; i++) {
        html += '<td></td>';
        html += `<td style="color:#e65100">${fmtNum(grandChem[i])}</td>`;
    }
    html += `<td style="color:#c62828">${fmtNum(grandTotal)}</td>`;
    html += '</tr>';

    html += '</tbody></table>';
    el.innerHTML = html;
}

/* ═══════════════════════════════════════════════
   데이터 현황 대시보드
   ═══════════════════════════════════════════════ */


/* ═══════════════════════════════════════════════
   열수급 분석
   ═══════════════════════════════════════════════ */
function computeMonthlyHeatBalance() {
    const salesH = DATA.sales_hourly;
    const salesD = DATA.sales_daily;
    const ops = DATA.operations_daily;
    if ((!salesH && !salesD) || !ops) return null;

    const findCol = (headers, name) => headers.findIndex(h => h === name);

    // 수요 컬럼 (sales_hourly - 시간별)
    const demandKeys = ['주택용난방', '업무용난방', '업무용냉방', '공공용난방', '공공용냉방'];
    let demandDef = [];
    if (salesH) {
        const shH = salesH.headers;
        demandDef = demandKeys.map(k => ({ key: k, col: findCol(shH, k) }));
    }

    // 인테코 (sales_daily)
    const intecoSaleCol = salesD ? findCol(salesD.headers, '인테코') : -1;

    // 외부수열 컬럼 (operations_daily)
    const oH = ops.headers;
    // EXT_SOURCES 레지스트리 기반 (활성 supply만)
    const supplyDef = EXT_SOURCES
        .filter(s => s.enabled && s.type === 'supply')
        .map(s => ({ key: s.opsName, col: findCol(oH, s.opsName) }));
    const dualSrc = EXT_SOURCES.find(s => s.enabled && s.type === 'dual');
    const intecoRecvCol = dualSrc ? findCol(oH, dualSrc.opsName) : -1;

    // 12개월 누적
    const months = Array.from({ length: 12 }, (_, i) => {
        const d = { month: i + 1, 인테코_판매: 0, 인테코_수열: 0 };
        demandKeys.forEach(k => d['d_' + k] = 0);
        supplyDef.forEach(c => d['s_' + c.key] = 0);
        return d;
    });

    // sales_hourly 집계 (시간별 → 월별 합산)
    if (salesH) {
        salesH.rows.forEach(row => {
            const m = Number(row[0]);
            if (m < 1 || m > 12) return;
            const b = months[m - 1];
            demandDef.forEach(c => { if (c.col >= 0) b['d_' + c.key] += Number(row[c.col]) || 0; });
        });
    }

    // sales_daily 인테코 집계
    if (salesD) {
        salesD.rows.forEach(row => {
            const m = Number(row[0]);
            if (m < 1 || m > 12) return;
            if (intecoSaleCol >= 0) months[m - 1].인테코_판매 += Number(row[intecoSaleCol]) || 0;
        });
    }

    // operations_daily 집계
    ops.rows.forEach(row => {
        const m = Number(row[0]);
        if (m < 1 || m > 12) return;
        const b = months[m - 1];
        supplyDef.forEach(c => { if (c.col >= 0) b['s_' + c.key] += Number(row[c.col]) || 0; });
        if (intecoRecvCol >= 0) b.인테코_수열 += Number(row[intecoRecvCol]) || 0;
    });

    // 합계 계산
    const supplyKeys = supplyDef.map(c => c.key);

    months.forEach(b => {
        b.수요합계 = demandKeys.reduce((s, k) => s + b['d_' + k], 0);
        b.외부합계 = supplyKeys.reduce((s, k) => s + b['s_' + k], 0);
        b.인테코_순 = b.인테코_수열 - b.인테코_판매;
        b.필요열량 = b.수요합계 - b.외부합계 - b.인테코_순;
    });

    // 축열조 보유열량 월말 변동
    const storageHCol = findCol(oH, '보유열량');
    if (storageHCol >= 0) {
        const monthLastDay = Array(12).fill(-1);
        const monthLastStorage = Array(12).fill(null);
        ops.rows.forEach(row => {
            const m = Number(row[0]);
            const day = Number(row[1]);
            if (m >= 1 && m <= 12) {
                const val = Number(row[storageHCol]);
                if (!isNaN(val) && day >= monthLastDay[m - 1]) {
                    monthLastDay[m - 1] = day;
                    monthLastStorage[m - 1] = val;
                }
            }
        });
        const initial = parseFloat(document.getElementById('storageInitial')?.value) || 750;
        months.forEach((b, i) => {
            const endVal = monthLastStorage[i];
            const startVal = i === 0 ? initial : monthLastStorage[i - 1];
            b.축열조변동 = (endVal !== null && startVal !== null) ? endVal - startVal : 0;
        });
    } else {
        months.forEach(b => b.축열조변동 = 0);
    }

    return { months, demandKeys, supplyKeys };
}

let _heatLossRates = Array(12).fill(5); // 월별 손실율 (기본 5%)
let _heatBalanceData = null; // 캐시

function renderHeatBalance() {
    const el = document.getElementById('pv_plan_monthly');
    if (!el) return;

    const result = computeMonthlyHeatBalance();
    if (!result) {
        el.innerHTML = '<p class="hint">판매실적과 운영실적 데이터가 필요합니다</p>';
        return;
    }

    _heatBalanceData = result;
    const { months, demandKeys, supplyKeys } = result;
    const fmt = v => fmtNum(Math.round(v));

    let html = '<div style="display:flex;align-items:center;gap:16px;margin-bottom:8px">';
    html += '<h4>월별 열수급 현황 (단위: Gcal)</h4>';
    html += '<label style="font-size:12px;color:#475569;display:flex;align-items:center;gap:6px">';
    html += '배관손실율 일괄(%)';
    html += '<input type="number" id="heatLossBulk" value="" placeholder="일괄" step="0.1" min="0" max="30" ';
    html += 'style="width:60px;padding:3px 6px;border:1px solid #e2e8f0;border-radius:4px;font-size:12px;text-align:right">';
    html += '</label>';
    html += '<span style="font-size:11px;color:#94a3b8">← 값 입력 시 전월 일괄 적용 / 아래에서 월별 개별 조정 가능</span>';
    html += '</div>';

    html += '<div style="overflow:auto; max-height:calc(100vh - 280px)">';
    html += '<table class="pv-tbl"><thead><tr>';
    html += '<th>월</th>';
    demandKeys.forEach(k => html += `<th style="background:#fffbeb">${k}</th>`);
    html += '<th style="background:#fef3c7;color:#92400e;font-weight:800">수요합계</th>';
    supplyKeys.forEach(k => html += `<th style="background:#ecfdf5">${k}</th>`);
    html += '<th style="background:#d1fae5;color:#065f46;font-weight:800">외부합계</th>';
    html += '<th>인테코(수열)</th>';
    html += '<th>인테코(판매)</th>';
    html += '<th style="background:#e0e7ff">인테코(순)</th>';
    html += '<th style="background:#e8eaf6;color:#283593;font-weight:800">축열조변동</th>';
    html += '<th style="background:#fee2e2;color:#991b1b;font-weight:800">필요열량</th>';
    html += '<th style="background:#fce4ec;color:#880e4f;font-weight:800">손실율<br>(%)</th>';
    html += '<th style="background:#fce4ec;color:#880e4f;font-weight:800">필요열량<br>(손실포함)</th>';
    html += '</tr></thead><tbody>';

    months.forEach((b, i) => {
        html += '<tr>';
        html += `<td>${b.month}월</td>`;
        demandKeys.forEach(k => html += `<td>${fmt(b['d_' + k])}</td>`);
        html += `<td style="background:#fffbeb;font-weight:700">${fmt(b.수요합계)}</td>`;
        supplyKeys.forEach(k => html += `<td>${fmt(b['s_' + k])}</td>`);
        html += `<td style="background:#ecfdf5;font-weight:700">${fmt(b.외부합계)}</td>`;
        html += `<td>${fmt(b.인테코_수열)}</td>`;
        html += `<td>${fmt(b.인테코_판매)}</td>`;
        html += `<td style="background:#eef2ff">${fmt(b.인테코_순)}</td>`;
        const stColor = b.축열조변동 > 0 ? 'color:#059669' : (b.축열조변동 < 0 ? 'color:#dc2626' : '');
        html += `<td style="background:#e8eaf6;font-weight:700;${stColor}">${fmt(b.축열조변동)}</td>`;
        html += `<td style="background:#fef2f2;font-weight:800">${fmt(b.필요열량)}</td>`;
        html += `<td style="background:#fff5f7;padding:2px"><input type="number" class="loss-rate-input" data-month="${i}" value="${_heatLossRates[i]}" step="0.1" min="0" max="30" style="width:50px;padding:2px 4px;border:1px solid #f9a8d4;border-radius:3px;font-size:11px;text-align:center;background:#fff"></td>`;
        html += `<td class="loss-cell" style="background:#fce4ec;font-weight:800">0</td>`;
        html += '</tr>';
    });

    // 연간 합계
    html += '<tr style="border-top:2px solid #1e293b;font-weight:800">';
    html += '<td>합계</td>';
    demandKeys.forEach(k => {
        html += `<td>${fmt(months.reduce((s, b) => s + b['d_' + k], 0))}</td>`;
    });
    html += `<td style="background:#fffbeb">${fmt(months.reduce((s, b) => s + b.수요합계, 0))}</td>`;
    supplyKeys.forEach(k => {
        html += `<td>${fmt(months.reduce((s, b) => s + b['s_' + k], 0))}</td>`;
    });
    html += `<td style="background:#ecfdf5">${fmt(months.reduce((s, b) => s + b.외부합계, 0))}</td>`;
    html += `<td>${fmt(months.reduce((s, b) => s + b.인테코_수열, 0))}</td>`;
    html += `<td>${fmt(months.reduce((s, b) => s + b.인테코_판매, 0))}</td>`;
    html += `<td style="background:#eef2ff">${fmt(months.reduce((s, b) => s + b.인테코_순, 0))}</td>`;
    html += `<td style="background:#e8eaf6">${fmt(months.reduce((s, b) => s + b.축열조변동, 0))}</td>`;
    html += `<td style="background:#fef2f2">${fmt(months.reduce((s, b) => s + b.필요열량, 0))}</td>`;
    const avgLoss = (_heatLossRates.reduce((a, b) => a + b, 0) / 12).toFixed(1);
    html += `<td style="background:#fff5f7;font-size:10px;color:#880e4f">평균 ${avgLoss}</td>`;
    html += `<td class="loss-cell" style="background:#fce4ec">0</td>`;
    html += '</tr>';

    html += '</tbody></table></div>';
    el.innerHTML = html;

    // 손실율 셀 초기 계산
    updateLossCells();

    // 월별 손실율 입력 이벤트
    document.querySelectorAll('#pv_plan_monthly .loss-rate-input').forEach(inp => {
        inp.addEventListener('input', (e) => {
            const m = parseInt(e.target.dataset.month);
            _heatLossRates[m] = parseFloat(e.target.value) || 0;
            updateLossCells();
            scheduleSettingsSave();
        });
    });

    // 일괄 적용 이벤트
    document.getElementById('heatLossBulk')?.addEventListener('input', (e) => {
        const val = parseFloat(e.target.value);
        if (isNaN(val)) return;
        _heatLossRates = Array(12).fill(val);
        document.querySelectorAll('#pv_plan_monthly .loss-rate-input').forEach(inp => {
            inp.value = val;
        });
        updateLossCells();
        scheduleSettingsSave();
    });
}

// 손실율 변경 시 해당 셀만 갱신 (DOM 재생성 없음 → 커서 유지)
function updateLossCells() {
    if (!_heatBalanceData) return;
    const { months } = _heatBalanceData;
    const cells = document.querySelectorAll('#pv_plan_monthly .loss-cell');
    const fmt = v => fmtNum(Math.round(v));

    let totalWithLoss = 0;
    months.forEach((b, i) => {
        const lossRate = _heatLossRates[i] / 100;
        const adjExternal = b.외부합계 * (1 - lossRate);
        const adjIntecoRecv = b.인테코_수열 * (1 - lossRate);
        const adjIntecoNet = adjIntecoRecv - b.인테코_판매;
        const withLoss = b.수요합계 - adjExternal - adjIntecoNet;
        if (cells[i]) cells[i].textContent = fmt(Math.round(withLoss));
        totalWithLoss += withLoss;
    });
    // 합계 행 (13번째 셀)
    if (cells[12]) cells[12].textContent = fmt(Math.round(totalWithLoss));
    // 평균 표시 갱신
    const avgCell = document.querySelector('#pv_plan_monthly tr:last-child td[style*="fff5f7"]');
    if (avgCell) avgCell.textContent = '평균 ' + (_heatLossRates.reduce((a, b) => a + b, 0) / 12).toFixed(1);
}

/* ═══════════════════════════════════════════════
   시뮬레이션 엔진
   ═══════════════════════════════════════════════ */
const LOOK_AHEAD_DAYS = 5;
const MJ_PER_GCAL = 4184;

// 정비일 Set 생성
function buildMaintenanceSet() {
    const set = new Set();
    for (let n = 1; n <= 3; n++) {
        const sm = parseInt(document.getElementById(`mt${n}StartM`)?.value);
        const sd = parseInt(document.getElementById(`mt${n}StartD`)?.value);
        const em = parseInt(document.getElementById(`mt${n}EndM`)?.value);
        const ed = parseInt(document.getElementById(`mt${n}EndD`)?.value);
        if (!sm || !sd || !em || !ed) continue;
        const start = new Date(CURRENT_YEAR, sm - 1, sd);
        const end = new Date(CURRENT_YEAR, em - 1, ed);
        for (let dt = new Date(start); dt <= end; dt.setDate(dt.getDate() + 1)) {
            set.add(`${dt.getMonth() + 1}-${dt.getDate()}`);
        }
    }
    return set;
}

// CHP 부하 테이블에서 성능값 읽기 (100% 기본)
function getChpPerf(loadPct) {
    const rows = document.querySelectorAll('#loadTable tbody tr');
    let best = null, bestDiff = Infinity;
    rows.forEach(tr => {
        const inputs = tr.querySelectorAll('input');
        const load = parseInt(tr.querySelector('td')?.textContent);
        if (isNaN(load)) return;
        const diff = Math.abs(load - loadPct);
        if (diff < bestDiff) {
            bestDiff = diff;
            best = {
                load,
                발전MW: parseFloat(inputs[0]?.value) || 0,
                송전MW: parseFloat(inputs[1]?.value) || 0,
                열Gcal: parseFloat(inputs[2]?.value) || 0,
                ngNm3: parseFloat(inputs[3]?.value) || 0, // NG소모량 N㎥/h
            };
        }
    });
    return best || { load: 100, 발전MW: 23.27, 송전MW: 21.89, 열Gcal: 51.39, ngNm3: 8050.1 };
}

// CHP 전체 부하 테이블 배열 반환 (60% 이상만, 부하 내림차순)
function getChpLoadTable() {
    const table = [];
    document.querySelectorAll('#loadTable tbody tr').forEach(tr => {
        const inputs = tr.querySelectorAll('input');
        const load = parseInt(tr.querySelector('td')?.textContent);
        const minLoad = parseFloat(document.getElementById('chpMinLoad')?.value) || 60;
        if (isNaN(load) || load < minLoad) return; // 최소부하 미만 제외
        table.push({
            load,
            발전MW: parseFloat(inputs[0]?.value) || 0,
            송전MW: parseFloat(inputs[1]?.value) || 0,
            열Gcal: parseFloat(inputs[2]?.value) || 0,
            ngNm3: parseFloat(inputs[3]?.value) || 0,
            효율: parseFloat(inputs[4]?.value) || 0,
        });
    });
    table.sort((a, b) => b.load - a.load); // 100% → 60% 내림차순
    return table;
}

// PLB 부하 테이블에서 성능값 읽기
function getPlbPerf(loadPct) {
    const rows = document.querySelectorAll('#plbLoadTable tbody tr');
    let best = null, bestDiff = Infinity;
    rows.forEach(tr => {
        const inputs = tr.querySelectorAll('input');
        const load = parseInt(tr.querySelector('td')?.textContent);
        if (isNaN(load)) return;
        const diff = Math.abs(load - loadPct);
        if (diff < bestDiff) {
            bestDiff = diff;
            best = {
                load,
                열Gcal: parseFloat(inputs[0]?.value) || 0,
                효율: parseFloat(inputs[1]?.value) || 0,
                ngNm3: parseFloat(inputs[2]?.value) || 0,
            };
        }
    });
    return best || { load: 100, 열Gcal: 68.80, 효율: 85.12, ngNm3: 7698 };
}

// PLB 전체 부하 테이블 배열 반환 (효율 최적 부하 선택용)
function getPlbLoadTable() {
    const table = [];
    document.querySelectorAll('#plbLoadTable tbody tr').forEach(tr => {
        const inputs = tr.querySelectorAll('input');
        const load = parseInt(tr.querySelector('td')?.textContent);
        if (isNaN(load)) return;
        table.push({
            load,
            열Gcal: parseFloat(inputs[0]?.value) || 0,
            효율: parseFloat(inputs[1]?.value) || 0,
            ngNm3: parseFloat(inputs[2]?.value) || 0,
        });
    });
    return table.sort((a, b) => b.열Gcal - a.열Gcal); // 열출력 내림차순
}

// 일별 데이터 배열 생성 (365일)
function buildDailyData() {
    const salesH = DATA.sales_hourly;
    const salesD = DATA.sales_daily;
    const ops = DATA.operations_daily;
    if ((!salesH && !salesD) || !ops) return null;

    const findCol = (headers, name) => headers.findIndex(h => h === name);
    const oH = ops.headers;

    // 수요 컬럼 인덱스 (sales_hourly)
    let shDemandCols = [];
    if (salesH) {
        const shH = salesH.headers;
        shDemandCols = ['주택용난방', '업무용난방', '업무용냉방', '공공용난방', '공공용냉방']
            .map(n => findCol(shH, n)).filter(i => i >= 0);
    }
    // 인테코 (sales_daily)
    const intecoSaleCol = salesD ? findCol(salesD.headers, '인테코') : -1;

    // 외부수열 컬럼 인덱스 (EXT_SOURCES 레지스트리 기반)
    const supplyCols = EXT_SOURCES
        .filter(s => s.enabled && s.type === 'supply')
        .map(s => findCol(oH, s.opsName)).filter(i => i >= 0);
    const dualSrc = EXT_SOURCES.find(s => s.enabled && s.type === 'dual');
    const intecoRecvCol = dualSrc ? findCol(oH, dualSrc.opsName) : -1;

    // 공휴일 Set
    const holidaySet = new Set();
    if (DATA.holidays) {
        DATA.holidays.rows.forEach(row => {
            if (Number(row[3]) === 1) holidaySet.add(`${Number(row[0])}-${Number(row[1])}`);
        });
    }

    // SMP 일 평균 맵
    const smpMap = {};
    if (DATA.smp_hourly) {
        DATA.smp_hourly.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            let sum = 0, cnt = 0;
            for (let h = 2; h < row.length; h++) {
                const v = Number(row[h]);
                if (!isNaN(v) && v > 0) { sum += v; cnt++; }
            }
            smpMap[key] = cnt > 0 ? sum / cnt : 0;
        });
    }

    // NG단가 월별 맵 (원/N㎥ 환산)
    const ngCostPerNm3 = Array(12).fill(0); // CHP NG단가 원/N㎥
    const ngPLBperNm3 = Array(12).fill(0);  // PLB NG단가 원/N㎥
    if (DATA.ng_price_monthly) {
        const ngH = DATA.ng_price_monthly.headers;
        const unitHeatCol = findCol(ngH, 'NG단위열량');
        const chpCol = findCol(ngH, 'CHP열량단가(원/MJ)');
        const plbCol = findCol(ngH, 'PLB열량단가(원/MJ)');
        DATA.ng_price_monthly.rows.forEach(row => {
            const m = Number(row[0]) - 1;
            if (m >= 0 && m < 12) {
                const unitHeat = unitHeatCol >= 0 ? (Number(row[unitHeatCol]) || 0) : 0; // MJ/N㎥
                if (chpCol >= 0) ngCostPerNm3[m] = (Number(row[chpCol]) || 0) * unitHeat;
                if (plbCol >= 0) ngPLBperNm3[m] = (Number(row[plbCol]) || 0) * unitHeat;
            }
        });
    }

    // TLF 월별 맵 + CP 월별 맵 (cp_monthly에서 로딩)
    const tlfMap = {};
    const cpMap = {};
    if (DATA.cp_monthly) {
        const cpH = DATA.cp_monthly.headers;
        const cpWdIdx = findCol(cpH, '평일CP');
        const cpHolIdx = findCol(cpH, '휴일CP');
        const tlfWdIdx = findCol(cpH, 'TLF(평일)');
        const tlfSatIdx = findCol(cpH, 'TLF(토요일)');
        const tlfHolIdx = findCol(cpH, 'TLF(공휴일)');
        DATA.cp_monthly.rows.forEach(row => {
            const mm = Number(row[0]);
            if (mm >= 1 && mm <= 12) {
                tlfMap[mm] = {
                    weekday: tlfWdIdx >= 0 ? (Number(row[tlfWdIdx]) || 1) : 1,
                    saturday: tlfSatIdx >= 0 ? (Number(row[tlfSatIdx]) || 1) : 1,
                    holiday: tlfHolIdx >= 0 ? (Number(row[tlfHolIdx]) || 1) : 1,
                };
                cpMap[mm] = {
                    weekdayCP: cpWdIdx >= 0 ? (Number(row[cpWdIdx]) || 0) : 0,
                    holidayCP: cpHolIdx >= 0 ? (Number(row[cpHolIdx]) || 0) : 0,
                };
            }
        });
    }

    // 무부하비정산 월별 (단가산정 설정에서)
    const noLoadSettle = Array(12).fill(0);
    const pricingConst = parseNum(document.getElementById('pricingConst')?.value);
    document.querySelectorAll('#pricingHeatTable tbody tr').forEach(tr => {
        const mm = parseInt(tr.dataset.m) - 1;
        if (mm >= 0 && mm < 12) {
            const hp = parseNum(tr.querySelector('input')?.value);
            if (hp > 0 && pricingConst > 0) {
                noLoadSettle[mm] = Math.round((pricingConst * hp) / 22000 * (1 / (1 + 2.6327)) * 100) / 100;
            }
        }
    });

    // sales_hourly / sales_daily / operations_daily → 일별 맵
    const demandMap = {}, demandHMap = {}, supplyMap = {}, intecoRecvMap = {}, intecoSellMap = {};
    // sales_hourly → 일별 수요 합산 + 시간별 수요 배열
    if (salesH) {
        salesH.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            const h = Number(row[2]) || 0; // 시 (0~23)
            const hourDemand = shDemandCols.reduce((s, c) => s + (Number(row[c]) || 0), 0);
            demandMap[key] = (demandMap[key] || 0) + hourDemand;
            if (!demandHMap[key]) demandHMap[key] = Array(24).fill(0);
            if (h >= 0 && h < 24) demandHMap[key][h] = hourDemand;
        });
    }
    // sales_daily → 인테코만
    if (salesD) {
        salesD.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            intecoSellMap[key] = intecoSaleCol >= 0 ? (Number(row[intecoSaleCol]) || 0) : 0;
        });
    }
    ops.rows.forEach(row => {
        const key = `${Number(row[0])}-${Number(row[1])}`;
        supplyMap[key] = supplyCols.reduce((s, c) => s + (Number(row[c]) || 0), 0);
        intecoRecvMap[key] = intecoRecvCol >= 0 ? (Number(row[intecoRecvCol]) || 0) : 0;
    });

    // 정비일 Set
    const mtSet = buildMaintenanceSet();

    // 365일 배열 생성
    const days = [];
    const isLeap = (CURRENT_YEAR % 4 === 0 && CURRENT_YEAR % 100 !== 0) || CURRENT_YEAR % 400 === 0;
    const daysInMonth = [31, isLeap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

    for (let m = 0; m < 12; m++) {
        for (let d = 1; d <= daysInMonth[m]; d++) {
            const key = `${m + 1}-${d}`;
            const dt = new Date(CURRENT_YEAR, m, d);
            const dow = dt.getDay(); // 0=일, 6=토
            const isWeekend = dow === 0 || dow === 6;
            const isHoliday = holidaySet.has(key);
            const isMaint = mtSet.has(key);
            const loss = _heatLossRates[m] / 100;

            const demand = demandMap[key] || 0;
            // 시간별 수요 배열 (없으면 균등 분배)
            const demandH = demandHMap[key] || Array(24).fill(demand / 24);
            const external = supplyMap[key] || 0;
            const intecoRecv = intecoRecvMap[key] || 0;
            const intecoSell = intecoSellMap[key] || 0;

            const adjExternal = external * (1 - loss);
            const adjIntecoNet = intecoRecv * (1 - loss) - intecoSell;
            const netRequired = demand - adjExternal - adjIntecoNet;

            // SMP 시간별
            const smpH = Array(24).fill(0);
            if (DATA.smp_hourly) {
                const row = DATA.smp_hourly.rows.find(r2 => Number(r2[0]) === (m + 1) && Number(r2[1]) === d);
                if (row) for (let h = 0; h < 24; h++) smpH[h] = Number(row[h + 2]) || 0;
            }

            // MP 시간별: SMP × TLF + 무부하비정산
            const isSaturday = dow === 6;
            const isHolidayOrSunday = dow === 0 || isHoliday;
            const tlfEntry = tlfMap[m + 1] || { weekday: 1, saturday: 1, holiday: 1 };
            const tlf = isHolidayOrSunday ? tlfEntry.holiday
                      : isSaturday ? tlfEntry.saturday
                      : tlfEntry.weekday;
            const nls = noLoadSettle[m];
            const mpH = smpH.map(v => Math.round(v * tlf * 100) / 100 + nls);

            days.push({
                month: m + 1, day: d, dow, isWeekend, isHoliday, isMaint,
                demand, demandH, external, intecoRecv, intecoSell, loss,
                adjExternal, adjIntecoNet, netRequired,
                smpAvg: smpMap[key] || 0, smpH, mpH, tlf, nls,
                ngCostNm3: ngCostPerNm3[m], ngPLBNm3: ngPLBperNm3[m],
                dowName: ['일','월','화','수','목','금','토'][dow],
            });
        }
    }
    return days;
}

// ═══════════════════════════════════════════════
// 글로벌 최적화 시뮬레이션 엔진 (2-Pass)
// ═══════════════════════════════════════════════
function runDailySimulation() {
    const days = buildDailyData();
    if (!days || days.length === 0) return null;

    // ─── 설정값 ───
    const chpLoadTable = getChpLoadTable(); // 최소부하~100% 부하별 성능
    const chp100 = chpLoadTable.find(c => c.load === 100) || chpLoadTable[0];
    const chpLow = chpLoadTable[chpLoadTable.length - 1]; // 최소부하 성능
    const plbCount = parseInt(document.getElementById('plbCount')?.value) || 3;
    const plbTable = getPlbLoadTable();
    const plbMaxCap = plbTable.length > 0 ? plbTable[0].열Gcal : 68.8;
    const storageCap = parseFloat(document.getElementById('storageCapacity')?.value) || 1500;
    const storageInit = parseFloat(document.getElementById('storageInitial')?.value) || 750;
    const storageMin = parseFloat(document.getElementById('storageMinLevel')?.value) || 100;
    const plbRecovery = parseFloat(document.getElementById('plbRecoveryTarget')?.value) || 200;
    const minRunTime = parseFloat(document.getElementById('chpMinRunHours')?.value) || 4;
    const minStopTime = parseFloat(document.getElementById('chpMinStopHours')?.value) || 4;
    const chpMinLoad = parseFloat(document.getElementById('chpMinLoad')?.value) || 60;
    const plbMinRunHours = parseFloat(document.getElementById('plbMinRunHours')?.value) || 2;
    const plb1DayCap = plbMaxCap * 24;

    // CHP 부하별 성능
    const chpPerf = chpLoadTable.map(c => ({
        load: c.load,
        heatRate: c.열Gcal,
        powerMW: c.송전MW,
        ngNm3: c.ngNm3,
    }));
    const perf100 = chpPerf.find(p => p.load === 100) || chpPerf[0];
    const perfLow = chpPerf[chpPerf.length - 1]; // 최소부하 성능

    // ─── 각 일의 CHP 비용/수익 계산 ───
    const dayInfo = days.map((d, i) => {
        const plbRef = plbTable.length > 0 ? plbTable[0] : null;
        const plbCostPerGcal = (plbRef && plbRef.열Gcal > 0)
            ? (d.ngPLBNm3 * plbRef.ngNm3) / plbRef.열Gcal : Infinity;

        // 부하별 CHP 단가 계산 (MP 기반)
        const mpAvg = d.mpH.reduce((s, v) => s + v, 0) / 24;
        const loadCosts = chpPerf.map(p => {
            const ngCostH = d.ngCostNm3 * p.ngNm3;
            const elecRevH = mpAvg * p.powerMW * 1000;
            const netCostH = ngCostH - elecRevH;
            const costPerGcal = p.heatRate > 0 ? netCostH / p.heatRate : Infinity;
            return { load: p.load, ngCostH, elecRevH, netCostH, costPerGcal, heatRate: p.heatRate, powerMW: p.powerMW, ngNm3: p.ngNm3 };
        });

        // 100% 부하 기준 비용
        const cost100 = loadCosts.find(c => c.load === 100) || loadCosts[0];
        // 최적 부하 (비용/Gcal 최소)
        const bestLoad = [...loadCosts].sort((a, b) => a.costPerGcal - b.costPerGcal)[0];

        return {
            idx: i, ...d,
            loadCosts,
            cost100,
            bestLoad,
            plbCostPerGcal,
            smpAvg: d.smpAvg,
            // SMP 이익 지표: 100%부하 기준 시간당 순이익 (높을수록 CHP 우선)
            chpProfitPerH: -cost100.netCostH,  // 양수 = 이익
        };
    });

    // ═══════════════════════════════════════════════
    // PASS 1: SMP 기반 글로벌 우선순위 배정
    // ═══════════════════════════════════════════════

    // 정비일 제외, SMP 수익이 높은 순으로 정렬
    const ranked = dayInfo.filter(d => !d.isMaint)
        .sort((a, b) => b.chpProfitPerH - a.chpProfitPerH);

    // 각 일에 초기 CHP 가동시간과 부하 배정
    const plan = new Array(days.length).fill(null).map(() => ({
        chpHours: 0, chpLoad: 100, chpPerf: perf100,
    }));

    // Step 1: 모든 날에 대해 CHP 가동 필요성 판단
    for (const d of ranked) {
        const cost100 = d.cost100;
        const chpIsProfitable = cost100.costPerGcal < 0;
        const chpIsCheaper = cost100.costPerGcal < d.plbCostPerGcal;

        if (chpIsProfitable) {
            // CHP 이익 → 24시간, 경제성 최적 부하 선택
            plan[d.idx].chpHours = 24;
            plan[d.idx].chpLoad = d.bestLoad.load;
            plan[d.idx].chpPerf = chpPerf.find(p => p.load === d.bestLoad.load) || perf100;
        } else if (chpIsCheaper) {
            // CHP가 PLB보다 쌈 → 열필요 기반 시간, 최적 부하
            const optPerf = chpPerf.find(p => p.load === d.bestLoad.load) || perf100;
            const heatNeeded = Math.max(0, d.netRequired);
            const hoursNeeded = optPerf.heatRate > 0 ? Math.ceil(heatNeeded / optPerf.heatRate) : 0;
            plan[d.idx].chpHours = Math.min(24, hoursNeeded);
            plan[d.idx].chpLoad = d.bestLoad.load;
            plan[d.idx].chpPerf = optPerf;
        } else {
            // PLB가 더 쌈 → CHP 정지
            plan[d.idx].chpHours = 0;
            plan[d.idx].chpLoad = 100;
        }
    }

    // ═══════════════════════════════════════════════
    // PASS 2: 시간순 시뮬레이션 + 축열조 제약 + 부하 조절
    // ═══════════════════════════════════════════════

    let storageLevel = storageInit;
    const results = [];
    let prevChpOn = false; // 전날 CHP 기동 여부 (연속운전 추적)
    let prevChpEndH = -99;  // 전날 CHP 종료시각 (일간 연속성 추적)
    let heurStopDays = 99;  // 정지 경과일수 (기동비 산출용)

    for (let i = 0; i < days.length; i++) {
        const d = dayInfo[i];
        const netReq = d.netRequired;
        const p = plan[i];

        // ─── 정비일 처리 ───
        if (d.isMaint) {
            const hourlyExt_m = d.adjExternal / 24;
            const intecoH_m = d.adjIntecoNet / 24;
            const demandH_m = d.demandH;
            const smpH_m = d.smpH || Array(24).fill(d.smpAvg);

            // PLB 필요 여부: 시간별 시뮬레이션
            let stCheck_m = storageLevel, needPlb_m = false, minSt_m = storageLevel;
            for (let h = 0; h < 24; h++) {
                stCheck_m = stCheck_m + (hourlyExt_m + intecoH_m - demandH_m[h]);
                if (stCheck_m > storageCap) stCheck_m = storageCap;
                if (stCheck_m < minSt_m) minSt_m = stCheck_m;
                if (stCheck_m < storageMin) needPlb_m = true;
            }

            let plbHeat = 0, plbUnits = 0, plbHours = 0, plbNg = 0, plbLoad = 0, plbStartH_m = 0;
            if (needPlb_m) {
                const deficit = storageMin - minSt_m;
                if (deficit > 0) {
                    plbUnits = Math.min(plbCount, Math.ceil(deficit / plb1DayCap));
                    const bestPlb = plbTable[0];
                    plbLoad = bestPlb.load;
                    plbHours = Math.ceil(deficit / (plbUnits * bestPlb.열Gcal));
                    plbHours = Math.max(plbMinRunHours, Math.min(24, plbHours));
                    plbHeat = plbHours * plbUnits * bestPlb.열Gcal;
                    plbNg = plbHours * plbUnits * bestPlb.ngNm3;
                    plbStartH_m = Math.max(0, Math.round(12 - plbHours / 2));
                }
            }

            // 시간별 시뮬레이션
            const hourly_m = [];
            let stH_m = storageLevel;
            for (let h = 0; h < 24; h++) {
                const plb = (plbHours > 0 && h >= plbStartH_m && h < plbStartH_m + plbHours)
                    ? (plbHeat / plbHours) : 0;
                stH_m = stH_m + (hourlyExt_m + intecoH_m + plb - demandH_m[h]);
                if (stH_m > storageCap) stH_m = storageCap;
                if (stH_m < storageMin) stH_m = storageMin;
                hourly_m.push({ hour: h, demand: demandH_m[h], ext: hourlyExt_m + intecoH_m, chp: 0, plb, storage: Math.round(stH_m), smp: smpH_m[h], mp: d.mpH[h] });
            }
            storageLevel = stH_m;

            const plbFuelCost = Math.round(plbNg * d.ngPLBNm3);
            results.push({
                month: d.month, day: d.day, dow: d.dow, dowName: d.dowName,
                isWeekend: d.isWeekend, isHoliday: d.isHoliday, isMaint: true,
                demand: d.demand, adjExternal: d.adjExternal, adjIntecoNet: d.adjIntecoNet,
                netRequired: netReq, smpAvg: d.smpAvg, smpH: smpH_m,
                chpHours: 0, chpHeat: 0, chpPower: 0, chpNg: 0, chpLoad: 0,
                plbUnits, plbHours, plbHeat, plbNg, plbLoad,
                storageLevel, hourly: hourly_m,
                chpFuelCost: 0, chpElecRev: 0, chpNetCost: 0, startupCost: 0,
                plbFuelCost, totalFuelCost: plbFuelCost, totalNetCost: plbFuelCost,
                chpCostPerGcal: 0, plbCostPerGcal: Math.round(d.plbCostPerGcal),
                dayType: '정비',
            });
            prevChpOn = false;
            prevChpEndH = -99;
            heurStopDays++;
            continue;
        }

        // ─── 축열조 기반 가동시간 조정 ───
        let chpHours = p.chpHours;
        let chpLoadPct = p.chpLoad;
        let selectedPerf = chpPerf.find(pp => pp.load === chpLoadPct) || perf100;

        // 축열조 가용 공간 기반 최대 가동시간 (100%)
        // CHP열은 수요 충당 + 축열조 여유공간까지만 (과생산 방지)
        const roomForHeat = Math.max(0, netReq + (storageCap - storageLevel));
        const maxUsefulHours100 = perf100.heatRate > 0 ? Math.floor(roomForHeat / perf100.heatRate) : 0;

        // ─── 핵심: 축열조 가득 차서 24h 못 돌릴 때 부하 낮춤 ───
        if (chpHours > maxUsefulHours100 && chpHours > 0) {
            // 시나리오 A: 100% 부하로 maxUsefulHours만 가동
            const costA_hours = maxUsefulHours100;
            const costA = costA_hours * d.cost100.netCostH;
            // 나머지 열은 PLB or 축열조가 처리

            // 시나리오 B: 부하를 낮춰서 24h 연속운전
            // 필요한 열/h = roomForHeat / 24
            const targetHeatPerH = roomForHeat / 24;
            // 적합한 부하 찾기 (60% 이상, targetHeatPerH에 가장 근접)
            let bestScenarioB = null;
            for (const lc of d.loadCosts) {
                if (lc.load < chpMinLoad) continue;
                if (lc.heatRate <= targetHeatPerH * 1.05) { // 약간 넘는것 허용
                    if (!bestScenarioB || Math.abs(lc.heatRate - targetHeatPerH) < Math.abs(bestScenarioB.heatRate - targetHeatPerH)) {
                        bestScenarioB = lc;
                    }
                }
            }
            // 부하 낮춰도 60%에서도 열이 넘치면 → 시간 줄이기
            if (!bestScenarioB) {
                // 60% 부하로도 넘침 → 60%로 가능한 시간만큼
                const perfLowCost = d.loadCosts[d.loadCosts.length - 1]; // 최소부하
                const maxHours60 = perfLowCost.heatRate > 0 ? Math.floor(roomForHeat / perfLowCost.heatRate) : 0;
                bestScenarioB = { ...perfLowCost, hours: maxHours60 };
            }

            if (bestScenarioB) {
                const hoursB = bestScenarioB.hours || 24;
                const costB = hoursB * bestScenarioB.netCostH;

                // 연속운전 보너스: 전날 켜져있었으면 기동비용 절감
                // 기동비 = Cold 기준 (최악 케이스) 기동시간 × 가스량 × NG단가
                const _coldTime = parseFloat(document.getElementById('coldTime')?.value) || 8;
                const _coldGas = parseFloat(document.getElementById('coldGasRate')?.value) || 400;
                const startupPenalty = prevChpOn ? 0 : (_coldTime * _coldGas * d.ngCostNm3);

                // 비용 비교: 시나리오 A vs B
                if (costB + (prevChpOn ? 0 : 0) < costA + (chpHours > 0 && !prevChpOn ? startupPenalty : 0)) {
                    // 부하 낮춰서 연속운전이 더 경제적
                    chpHours = hoursB;
                    chpLoadPct = bestScenarioB.load;
                    selectedPerf = chpPerf.find(pp => pp.load === chpLoadPct) || perf100;
                } else {
                    chpHours = costA_hours;
                    chpLoadPct = 100;
                    selectedPerf = perf100;
                }
            } else {
                chpHours = maxUsefulHours100;
            }
        }

        // ─── 연속운전 판단: 전날 켜져있고 오늘 끄려면 비용 비교 ───
        if (chpHours === 0 && prevChpOn) {
            // 정지 대신 최소부하로 유지하는 비용 vs 정지+재기동 비용
            const costLow = d.loadCosts[d.loadCosts.length - 1]; // 최소부하

            // 향후 며칠 내 재기동 필요한지 체크
            let daysUntilRestart = 0;
            for (let j = i + 1; j < Math.min(i + 7, days.length); j++) {
                if (dayInfo[j].isMaint) { daysUntilRestart = 999; break; }
                if (plan[j].chpHours > 0) break;
                daysUntilRestart++;
            }

            if (daysUntilRestart < minStopTime) {
                // 정지시간 < 최소정지시간 → 연속운전 강제 (부하 낮춤)
                const room = Math.max(0, netReq + (storageCap - storageLevel));
                const maxHLow = perfLow.heatRate > 0 ? Math.floor(room / perfLow.heatRate) : 0;
                if (maxHLow >= minRunTime) {
                    chpHours = maxHLow;
                    chpLoadPct = chpMinLoad;
                    selectedPerf = perfLow;
                }
            } else if (daysUntilRestart <= 3) {
                // 3일 이내 재기동 예정 → 최소부하 유지 vs 정지+재기동 비교
                const keepCost = 24 * costLow.netCostH;
                const restartCost = 50000; // 기동비 추정
                if (keepCost < restartCost && costLow.costPerGcal < d.plbCostPerGcal) {
                    const room = Math.max(0, netReq + (storageCap - storageLevel));
                    const maxHLow = perfLow.heatRate > 0 ? Math.floor(room / perfLow.heatRate) : 0;
                    if (maxHLow >= minRunTime) {
                        chpHours = Math.min(24, maxHLow);
                        chpLoadPct = chpMinLoad;
                        selectedPerf = perfLow;
                    }
                }
            }
        }

        // ─── 축열조 부족 시 CHP 투입 강제 ───
        if (chpHours === 0 && !d.isMaint) {
            const storageBuffer = storageLevel - storageMin;
            if (netReq > storageBuffer) {
                // 축열조만으로 부족 → CHP 켜야 함
                const deficit = netReq - storageBuffer;
                const hoursNeeded = perf100.heatRate > 0 ? Math.ceil(deficit / perf100.heatRate) : 0;
                if (hoursNeeded >= minRunTime) {
                    chpHours = hoursNeeded;
                    chpLoadPct = 100;
                    selectedPerf = perf100;
                }
            }
        }

        // ─── 최종 제약 적용 ───
        chpHours = Math.min(24, Math.max(0, chpHours));

        // 최소운전시간
        if (chpHours > 0 && chpHours < minRunTime) chpHours = minRunTime;

        // 최소정지시간
        if (chpHours > 0 && (24 - chpHours) > 0 && (24 - chpHours) < minStopTime) {
            const roomMax = Math.max(0, netReq + (storageCap - storageLevel));
            const maxH = selectedPerf.heatRate > 0 ? Math.floor(roomMax / selectedPerf.heatRate) : 0;
            if (maxH >= 24) {
                chpHours = 24;
            } else {
                chpHours = Math.max(0, 24 - minStopTime);
            }
        }

        // 오버플로 최종 방지: CHP열 투입 후 축열조가 storageCap을 넘지 않도록
        // storageLevel + chpHeat - netReq <= storageCap → chpHeat <= netReq + (storageCap - storageLevel)
        const finalRoom = netReq + (storageCap - storageLevel);
        const finalMaxH = (selectedPerf.heatRate > 0 && finalRoom > 0) ? Math.floor(finalRoom / selectedPerf.heatRate) : 0;
        chpHours = Math.min(chpHours, finalMaxH, 24);

        // 최소운전시간 확보 불가 → 정지
        if (chpHours > 0 && chpHours < minRunTime) chpHours = 0;

        // ═══ 2단계: 시간별 스케줄 생성 ═══
        const chpHeatPerH = selectedPerf.heatRate;
        const hourlyExt = d.adjExternal / 24;
        const intecoH = d.adjIntecoNet / 24;
        const demandH = d.demandH;
        const smpH = d.smpH || Array(24).fill(d.smpAvg);

        // CHP 연속 블록: SMP 최대 + overflow 완전 거부 (R2) + 일간 연속성
        let chpStart = 0;
        if (chpHours > 0 && chpHours < 24) {
            let bestScore = -Infinity;
            // overflow 없는 배치를 우선 탐색, 없으면 CHP 시간 축소
            let tryHours = chpHours;
            while (tryHours >= minRunTime) {
                bestScore = -Infinity;
                for (let s = 0; s <= 24 - tryHours; s++) {
                    const gapFromPrev = s + (24 - prevChpEndH);
                    if (prevChpOn && gapFromPrev > 0 && gapFromPrev < minStopTime) continue;

                    let stSim = storageLevel, smpSum = 0, overflow = 0, underflow = 0;
                    for (let h = 0; h < 24; h++) {
                        const chp = (h >= s && h < s + tryHours) ? chpHeatPerH : 0;
                        stSim = stSim + (hourlyExt + intecoH + chp - demandH[h]);
                        if (stSim > storageCap) { overflow += stSim - storageCap; stSim = storageCap; }
                        if (stSim < storageMin) { underflow += storageMin - stSim; stSim = storageMin; }
                        if (h >= s && h < s + tryHours) smpSum += d.mpH[h];
                    }
                    if (overflow > 0) continue; // R2: overflow 배치 완전 거부
                    const score = smpSum - underflow * 200;
                    if (score > bestScore) { bestScore = score; chpStart = s; chpHours = tryHours; }
                }
                if (bestScore > -Infinity) break; // overflow 없는 배치 찾음
                tryHours--; // 모든 배치 overflow → 1시간 줄여서 재시도
            }
            if (bestScore === -Infinity) chpHours = 0; // 어떤 배치도 불가 → 미가동
        }

        // PLB: 축열조가 부족한 시점에 배치
        // 먼저 CHP만으로 시뮬 → 부족 여부 확인
        let stCheck = storageLevel, needPlb = false;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours)) ? chpHeatPerH : 0;
            stCheck = stCheck + (hourlyExt + intecoH + chp - demandH[h]);
            if (stCheck > storageCap) stCheck = storageCap;
            if (stCheck < storageMin) { needPlb = true; break; }
        }

        let plbHeat = 0, plbUnits = 0, plbHours = 0, plbNg = 0, plbLoad = 0;
        let plbStartH = 0;
        if (needPlb) {
            // 부족분 계산 (R1: 축열조 하한 방지)
            let stSim2 = storageLevel, minSt = storageLevel, minStH = 0;
            for (let h = 0; h < 24; h++) {
                const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours)) ? chpHeatPerH : 0;
                stSim2 = stSim2 + (hourlyExt + intecoH + chp - demandH[h]);
                if (stSim2 > storageCap) stSim2 = storageCap;
                if (stSim2 < minSt) { minSt = stSim2; minStH = h; }
            }
            const deficit = storageMin - minSt;
            if (deficit > 0) {
                // PLB 투입량: deficit 기반, 과투입 방지 (deficit의 2배 상한)
                const targetHeat = Math.min(deficit * 2, plb1DayCap);
                const bestPlb = plbTable[0];
                // 대수: 1대로 커버 가능하면 1대
                plbUnits = Math.ceil(deficit / (bestPlb.열Gcal * Math.max(plbMinRunHours, 1)));
                plbUnits = Math.min(plbUnits, plbCount);
                if (plbUnits < 1) plbUnits = 1;
                plbLoad = bestPlb.load;
                plbHours = Math.ceil(deficit / (plbUnits * bestPlb.열Gcal));
                plbHours = Math.max(plbMinRunHours, Math.min(24 - chpHours, plbHours));
                // 과투입 상한: 실제 투입열이 deficit*2를 넘지 않게
                const actualHeat = plbHours * plbUnits * bestPlb.열Gcal;
                if (actualHeat > deficit * 2 && plbHours > plbMinRunHours) {
                    plbHours = Math.max(plbMinRunHours, Math.ceil(deficit * 2 / (plbUnits * bestPlb.열Gcal)));
                }
                plbHeat = plbHours * plbUnits * bestPlb.열Gcal;
                plbNg = plbHours * plbUnits * bestPlb.ngNm3;
                // PLB를 부족 시점 직전에 배치
                plbStartH = Math.max(0, minStH - plbHours);
            }
        }

        // ═══ 3단계: 시간별 시뮬레이션 (확정) ═══
        const hourly = [];
        let stH = storageLevel;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours && chpHours > 0)) ? chpHeatPerH : 0;
            const plb = (plbHours > 0 && h >= plbStartH && h < plbStartH + plbHours)
                ? (plbHeat / plbHours) : 0;
            // R2: overflow 방지 — CHP 열출력을 축열조 상한에 맞게 자동 감소
            let actualChp = chp;
            const projected = stH + hourlyExt + intecoH + actualChp + plb - demandH[h];
            if (projected > storageCap && actualChp > 0) {
                actualChp = Math.max(0, storageCap - stH - hourlyExt - intecoH - plb + demandH[h]);
            }
            stH = stH + (hourlyExt + intecoH + actualChp + plb - demandH[h]);
            if (stH > storageCap) stH = storageCap;
            if (stH < storageMin) stH = storageMin;
            const chpPowerKWh = (actualChp > 0) ? selectedPerf.powerMW * 1000 * (actualChp / (chpHeatPerH || 1)) : 0;
            hourly.push({ hour: h, demand: demandH[h], ext: hourlyExt + intecoH, chp: actualChp, plb, storage: Math.round(stH), smp: smpH[h], mp: d.mpH[h], chpPowerKWh });
        }
        storageLevel = stH;

        // ═══ 4단계: 일별 집계 ═══
        const chpHeat = chpHours * chpHeatPerH;
        const chpPower = chpHours * selectedPerf.powerMW;
        const chpNg = chpHours * selectedPerf.ngNm3;
        const selCost = d.loadCosts.find(c => c.load === chpLoadPct) || d.cost100;
        const chpFuelCost = Math.round(chpHours * selCost.ngCostH);
        // 열제약매출: 시간별 MP × 송전MW × 1000 합산
        let chpElecRev = 0;
        hourly.forEach(hr => { if (hr.chp > 0) chpElecRev += hr.mp * selectedPerf.powerMW * 1000; });
        chpElecRev = Math.round(chpElecRev);
        const chpNetCost = chpFuelCost - chpElecRev;
        const chpCostPerGcal = chpHeat > 0 ? Math.round(chpNetCost / chpHeat) : 0;
        const plbFuelCost = Math.round(plbNg * d.ngPLBNm3);

        // 기동비용: 정지일수 → stopLevel 변환
        let startupCost = 0;
        if (chpHours > 0 && heurStopDays > 0) {
            const _hotTime = parseFloat(document.getElementById('hotTime')?.value) || 4;
            const _hotGas = parseFloat(document.getElementById('hotGasRate')?.value) || 300;
            const _warmTime = parseFloat(document.getElementById('warmTime')?.value) || 6;
            const _warmGas = parseFloat(document.getElementById('warmGasRate')?.value) || 350;
            const _coldTime = parseFloat(document.getElementById('coldTime')?.value) || 8;
            const _coldGas = parseFloat(document.getElementById('coldGasRate')?.value) || 400;
            let suTime, suGas;
            if (heurStopDays <= 1) { suTime = _hotTime; suGas = _hotGas; }
            else if (heurStopDays <= 2) { suTime = _warmTime; suGas = _warmGas; }
            else { suTime = _coldTime; suGas = _coldGas; }
            startupCost = Math.round(suTime * suGas * d.ngCostNm3);
        }
        // 정지일수 갱신
        if (chpHours > 0) { heurStopDays = 0; } else { heurStopDays++; }

        results.push({
            month: d.month, day: d.day, dow: d.dow, dowName: d.dowName,
            isWeekend: d.isWeekend, isHoliday: d.isHoliday, isMaint: d.isMaint,
            demand: d.demand, adjExternal: d.adjExternal, adjIntecoNet: d.adjIntecoNet,
            netRequired: netReq, smpAvg: d.smpAvg, smpH,
            chpHours, chpHeat, chpPower, chpNg, chpLoad: chpLoadPct,
            plbUnits, plbHours, plbHeat, plbNg, plbLoad,
            storageLevel, hourly,
            chpFuelCost, chpElecRev, chpNetCost, startupCost,
            plbFuelCost,
            totalFuelCost: chpFuelCost + plbFuelCost + startupCost,
            totalNetCost: chpNetCost + plbFuelCost + startupCost,
            chpCostPerGcal,
            plbCostPerGcal: Math.round(d.plbCostPerGcal),
            dayType: d.isMaint ? '정비' : (d.isHoliday ? '공휴일' : (d.isWeekend ? '주말' : '평일')),
        });

        prevChpOn = chpHours > 0;
        prevChpEndH = chpHours > 0 ? chpStart + chpHours : -99;
    }

    return results;
}

// ═══════════════════════════════════════════════
// DP (Dynamic Programming) 최적화 솔버
// Bellman Equation 기반 역방향 최적화
// ═══════════════════════════════════════════════
function solveDP() {
    const days = buildDailyData();
    if (!days || days.length === 0) return null;

    // ─── 설정값 ───
    const chpLoadTable = getChpLoadTable();
    const chp100 = chpLoadTable.find(c => c.load === 100) || chpLoadTable[0];
    const plbCount = parseInt(document.getElementById('plbCount')?.value) || 3;
    const plbTable = getPlbLoadTable();
    const plbMaxCap = plbTable.length > 0 ? plbTable[0].열Gcal : 68.8;
    const storageCap = parseFloat(document.getElementById('storageCapacity')?.value) || 1500;
    const storageInit = parseFloat(document.getElementById('storageInitial')?.value) || 750;
    const storageMin = parseFloat(document.getElementById('storageMinLevel')?.value) || 100;
    const plbRecovery = parseFloat(document.getElementById('plbRecoveryTarget')?.value) || 200;
    const minRunTime = parseFloat(document.getElementById('chpMinRunHours')?.value) || 4;
    const minStopTime = parseFloat(document.getElementById('chpMinStopHours')?.value) || 4;
    const chpMinLoad = parseFloat(document.getElementById('chpMinLoad')?.value) || 60;
    const plbMinRunHours = parseFloat(document.getElementById('plbMinRunHours')?.value) || 2;

    // ─── 기동 가스사용량 설정 (Nm³/h) ───
    const startupTimes = [
        0,
        parseFloat(document.getElementById('hotTime')?.value) || 4,
        parseFloat(document.getElementById('warmTime')?.value) || 6,
        parseFloat(document.getElementById('coldTime')?.value) || 8,
    ];
    const startupGasRates = [
        0,
        parseFloat(document.getElementById('hotGasRate')?.value) || 300,
        parseFloat(document.getElementById('warmGasRate')?.value) || 350,
        parseFloat(document.getElementById('coldGasRate')?.value) || 400,
    ];
    // 기동비용 = 기동시간(h) × 가스사용량(Nm³/h) × 월별 NG단가(원/Nm³)
    function calcStartupCost(stopLevel, ngCostNm3) {
        if (stopLevel <= 0) return 0;
        return startupTimes[stopLevel] * startupGasRates[stopLevel] * ngCostNm3;
    }

    const N = days.length;
    const STEP = 10; // 축열조 이산화 단위 (Gcal)
    const numStates = Math.floor((storageCap - storageMin) / STEP) + 1;
    const stateToStorage = j => storageMin + j * STEP;
    const storageToState = s => Math.min(numStates - 1, Math.max(0, Math.round((s - storageMin) / STEP)));

    // ─── 2D 상태공간: 축열조 × 정지레벨 ───
    const NUM_STOP = 4; // 0=ON, 1=Hot(정지1일), 2=Warm(정지2일), 3=Cold(정지3일+)
    const totalStates = numStates * NUM_STOP;
    const dpIdx = (j, k) => j * NUM_STOP + k;

    // CHP 부하별 성능
    const perfOptions = chpLoadTable.filter(c => c.load >= chpMinLoad).map(c => ({
        load: c.load, heatRate: c.열Gcal, powerMW: c.송전MW, ngNm3: c.ngNm3,
    }));

    // PLB 비용/Gcal (100% 기준)
    const plbRef = plbTable.length > 0 ? plbTable[0] : null;
    const plb1DayCap = plbMaxCap * 24 * plbCount;

    // 각 날의 비용 정보 사전 계산 (MP 기반)
    const dayInfo = days.map(d => {
        const mpAvg = d.mpH.reduce((s, v) => s + v, 0) / 24;
        const costs = perfOptions.map(p => {
            const ngCostH = d.ngCostNm3 * p.ngNm3;
            const elecRevH = mpAvg * p.powerMW * 1000;
            return { load: p.load, heatRate: p.heatRate, netCostH: ngCostH - elecRevH, ngCostH, elecRevH, ngNm3: p.ngNm3, powerMW: p.powerMW };
        });
        const plbCostPerGcal = (plbRef && plbRef.열Gcal > 0)
            ? (d.ngPLBNm3 * plbRef.ngNm3) / plbRef.열Gcal : Infinity;
        return { ...d, costs, plbCostPerGcal };
    });

    // DP 테이블: dp[(j,k)] = 현재 day부터 끝까지 최소비용
    let dpCurr = new Float64Array(totalStates).fill(0);
    let dpNext = new Float64Array(totalStates);

    // 결정 기록: [day][j*NUM_STOP+k] = { hours, loadIdx }
    const decisions = Array.from({ length: N }, () =>
        Array.from({ length: totalStates }, () => ({ hours: 0, loadIdx: 0 }))
    );

    // CHP 시간 후보: 0, minRunTime~24
    const hourCandidates = [0];
    for (let h = Math.round(minRunTime); h <= 24; h++) hourCandidates.push(h);

    console.time('DP 솔버');

    // ─── 역방향 DP (Bellman Equation) ───
    for (let i = N - 1; i >= 0; i--) {
        const d = dayInfo[i];
        dpNext.set(dpCurr);
        dpCurr.fill(Infinity);

        for (let j = 0; j < numStates; j++) {
            const s = stateToStorage(j);

            for (let k = 0; k < NUM_STOP; k++) {
                const idx = dpIdx(j, k);
                let bestCost = Infinity;
                let bestDecision = { hours: 0, loadIdx: 0 };

                // 각 부하 × 시간 조합 시도
                for (let li = 0; li < perfOptions.length; li++) {
                    const perf = d.costs[li];

                    for (const h of hourCandidates) {
                        if (h > 0 && d.isMaint) continue;

                        // CHP 열생산 & 비용
                        const chpHeat = h * perf.heatRate;
                        const chpCost = h * perf.netCostH;

                        // 기동비용: CHP 가동 시 정지상태였으면 기동비 추가
                        let suCost = 0;
                        let nextK;
                        if (h > 0) {
                            nextK = 0; // 가동 → ON
                            if (k > 0) suCost = calcStartupCost(k, d.ngCostNm3);
                        } else {
                            // 비가동: 정지레벨 증가
                            nextK = Math.min(k + 1, NUM_STOP - 1);
                        }

                        // 축열조 변화
                        let newS = s + chpHeat - d.netRequired;

                        // PLB 투입 판단
                        let plbHeat = 0, plbCost = 0;
                        if (newS < storageMin) {
                            const deficit = storageMin - newS;
                            const plbMinHeat = plbMinRunHours * plbMaxCap;
                            plbHeat = Math.min(plb1DayCap, Math.max(deficit, plbMinHeat));
                            plbCost = plbHeat * d.plbCostPerGcal;
                            newS += plbHeat;
                            if (newS < storageMin) newS = storageMin;
                        }

                        // 축열조 상한
                        if (newS > storageCap) {
                            if (h === 0) newS = storageCap;
                            else continue;
                        }

                        const jNew = storageToState(newS);
                        const totalCost = chpCost + suCost + plbCost + dpNext[dpIdx(jNew, nextK)];

                        if (totalCost < bestCost) {
                            bestCost = totalCost;
                            bestDecision = { hours: h, loadIdx: li };
                        }
                    }
                }

                dpCurr[idx] = bestCost;
                decisions[i][idx] = bestDecision;
            }
        }
    }

    console.timeEnd('DP 솔버');

    // ─── 순방향 복원: 최적 스케줄 → 시뮬레이션 결과 ───
    let storageLevel = storageInit;
    let sj = storageToState(storageInit);
    let sk = 3; // 초기: Cold (장기간 정지 상태에서 시작)
    const results = [];
    let prevChpOn_dp = false, prevChpEndH_dp = -99;

    for (let i = 0; i < N; i++) {
        const d = dayInfo[i];
        const dec = decisions[i][dpIdx(sj, sk)];
        let chpHours = dec.hours;
        const perf = d.costs[dec.loadIdx];
        const chpLoadPct = perfOptions[dec.loadIdx].load;
        const chpHeatPerH = perf.heatRate;
        const smpH = d.smpH || Array(24).fill(d.smpAvg);

        // 기동비용 산출
        let startupCost = 0;
        if (chpHours > 0 && sk > 0) {
            startupCost = calcStartupCost(sk, d.ngCostNm3);
        }

        // 정지레벨 전이
        if (chpHours > 0) {
            sk = 0;
        } else {
            sk = Math.min(sk + 1, NUM_STOP - 1);
        }

        const hourlyExt = d.adjExternal / 24;
        const intecoH = d.adjIntecoNet / 24;
        const demandH = d.demandH;

        // ═══ 2단계: CHP 연속 블록 배치 (R2: overflow 완전 거부) ═══
        let chpStart = 0;
        if (chpHours > 0 && chpHours < 24) {
            let bestScore = -Infinity;
            let tryHours = chpHours;
            while (tryHours >= minRunTime) {
                bestScore = -Infinity;
                for (let s = 0; s <= 24 - tryHours; s++) {
                    const gapFromPrev = s + (24 - prevChpEndH_dp);
                    if (prevChpOn_dp && gapFromPrev > 0 && gapFromPrev < minStopTime) continue;

                    let stSim = storageLevel, smpSum = 0, overflow = 0, underflow = 0;
                    for (let h = 0; h < 24; h++) {
                        const chp = (h >= s && h < s + tryHours) ? chpHeatPerH : 0;
                        stSim = stSim + (hourlyExt + intecoH + chp - demandH[h]);
                        if (stSim > storageCap) { overflow += stSim - storageCap; stSim = storageCap; }
                        if (stSim < storageMin) { underflow += storageMin - stSim; stSim = storageMin; }
                        if (h >= s && h < s + tryHours) smpSum += d.mpH[h];
                    }
                    if (overflow > 0) continue;
                    const score = smpSum - underflow * 200;
                    if (score > bestScore) { bestScore = score; chpStart = s; chpHours = tryHours; }
                }
                if (bestScore > -Infinity) break;
                tryHours--;
            }
            if (bestScore === -Infinity) chpHours = 0;
        }

        // PLB 필요 여부 시간별 검증 (R1)
        let stCheck = storageLevel, needPlb = false, minSt_dp = storageLevel, minStH_dp = 0;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours && chpHours > 0)) ? chpHeatPerH : 0;
            stCheck = stCheck + (hourlyExt + intecoH + chp - demandH[h]);
            if (stCheck > storageCap) stCheck = storageCap;
            if (stCheck < minSt_dp) { minSt_dp = stCheck; minStH_dp = h; }
            if (stCheck < storageMin) needPlb = true;
        }

        let plbHeat = 0, plbUnits = 0, plbHours = 0, plbNg = 0, plbLoad = 0, plbStartH = 0;
        if (needPlb) {
            const deficit = storageMin - minSt_dp;
            if (deficit > 0) {
                const bestPlb = plbTable[0];
                plbUnits = Math.ceil(deficit / (bestPlb.열Gcal * Math.max(plbMinRunHours, 1)));
                plbUnits = Math.min(plbUnits, plbCount);
                if (plbUnits < 1) plbUnits = 1;
                plbLoad = bestPlb.load;
                plbHours = Math.ceil(deficit / (plbUnits * bestPlb.열Gcal));
                plbHours = Math.max(plbMinRunHours, Math.min(24 - chpHours, plbHours));
                const actualHeat = plbHours * plbUnits * bestPlb.열Gcal;
                if (actualHeat > deficit * 2 && plbHours > plbMinRunHours) {
                    plbHours = Math.max(plbMinRunHours, Math.ceil(deficit * 2 / (plbUnits * bestPlb.열Gcal)));
                }
                plbHeat = plbHours * plbUnits * bestPlb.열Gcal;
                plbNg = plbHours * plbUnits * bestPlb.ngNm3;
                plbStartH = Math.max(0, minStH_dp - plbHours);
            }
        }

        // ═══ 3단계: 시간별 시뮬레이션 ═══
        const hourly = [];
        let stH = storageLevel;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours && chpHours > 0)) ? chpHeatPerH : 0;
            const plb = (plbHours > 0 && h >= plbStartH && h < plbStartH + plbHours)
                ? (plbHeat / plbHours) : 0;
            // R2: overflow 방지
            let actualChp = chp;
            const projected = stH + hourlyExt + intecoH + actualChp + plb - demandH[h];
            if (projected > storageCap && actualChp > 0) {
                actualChp = Math.max(0, storageCap - stH - hourlyExt - intecoH - plb + demandH[h]);
            }
            stH = stH + (hourlyExt + intecoH + actualChp + plb - demandH[h]);
            if (stH > storageCap) stH = storageCap;
            if (stH < storageMin) stH = storageMin;
            const chpPowerKWh = (actualChp > 0) ? perf.powerMW * 1000 * (actualChp / (chpHeatPerH || 1)) : 0;
            hourly.push({ hour: h, demand: demandH[h], ext: hourlyExt + intecoH, chp: actualChp, plb, storage: Math.round(stH), smp: smpH[h], mp: d.mpH[h], chpPowerKWh });
        }
        storageLevel = stH;

        // ═══ 4단계: 일별 집계 ═══
        const chpHeat = chpHours * chpHeatPerH;
        const chpPower = chpHours * perf.powerMW;
        const chpNg = chpHours * perf.ngNm3;
        const chpFuelCost = Math.round(chpHours * perf.ngCostH);
        let chpElecRev = 0;
        hourly.forEach(hr => { if (hr.chp > 0) chpElecRev += hr.mp * perf.powerMW * 1000; });
        chpElecRev = Math.round(chpElecRev);
        const chpNetCost = chpFuelCost - chpElecRev;
        const chpCostPerGcal = chpHeat > 0 ? Math.round(chpNetCost / chpHeat) : 0;
        const plbFuelCost = Math.round(plbNg * d.ngPLBNm3);

        results.push({
            month: d.month, day: d.day, dow: d.dow, dowName: d.dowName,
            isWeekend: d.isWeekend, isHoliday: d.isHoliday, isMaint: d.isMaint,
            demand: d.demand, adjExternal: d.adjExternal, adjIntecoNet: d.adjIntecoNet,
            netRequired: d.netRequired, smpAvg: d.smpAvg, smpH,
            chpHours, chpHeat, chpPower, chpNg, chpLoad: chpLoadPct,
            plbUnits, plbHours, plbHeat, plbNg, plbLoad,
            storageLevel, hourly,
            chpFuelCost, chpElecRev, chpNetCost, startupCost,
            plbFuelCost,
            totalFuelCost: chpFuelCost + plbFuelCost + startupCost,
            totalNetCost: chpNetCost + plbFuelCost + startupCost,
            chpCostPerGcal,
            plbCostPerGcal: Math.round(d.plbCostPerGcal),
            dayType: d.isMaint ? '정비' : (d.isHoliday ? '공휴일' : (d.isWeekend ? '주말' : '평일')),
        });

        sj = storageToState(storageLevel);
        prevChpOn_dp = chpHours > 0;
        prevChpEndH_dp = chpHours > 0 ? chpStart + chpHours : -99;
    }

    return results;
}

// ─── 정밀 최적화 실행 (MPC 방식: 매일 역방향DP + 1일 실행 반복) ───
function runMILPOptimization() {
    const statusEl = document.getElementById('solverStatus');
    if (statusEl) statusEl.textContent = '정밀 최적화 계산 중... (0%)';

    setTimeout(() => {
        try {
            const startTime = performance.now();
            const simResults = solveMPC(30, (pct, day) => {
                if (statusEl) statusEl.textContent = `정밀 최적화 ${pct}% (${day}일차)`;
            });
            const elapsed = Math.round(performance.now() - startTime);

            if (!simResults) {
                alert('데이터가 필요합니다');
                if (statusEl) statusEl.textContent = '';
                return;
            }

            PLAN_RESULTS = simResults;
            if (typeof checkPipelineStatus === 'function') checkPipelineStatus();
            if (statusEl) statusEl.textContent = `정밀 최적화 완료 (${(elapsed/1000).toFixed(1)}초)`;

            const simWrap = document.getElementById('planSimResult');
            if (simWrap) simWrap.style.display = '';

            document.querySelectorAll('#planTabs .inner-tab').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.plan-panel').forEach(p => p.classList.remove('active'));
            document.querySelector('#planTabs .inner-tab[data-ptab="plan_monthly_sim"]')?.classList.add('active');
            document.getElementById('plan_monthly_sim')?.classList.add('active');

            renderPlanMonthly(simResults);
            renderPlanCharts(simResults);
            renderPlanDaily(simResults);
            renderContribution(simResults);
            renderAllCosts(simResults);

            const btnLogic = document.getElementById('btnShowLogicFlow');
            if (btnLogic) btnLogic.style.display = '';
        } catch (e) {
            console.error('정밀 최적화 오류:', e);
            if (statusEl) statusEl.textContent = '오류: ' + e.message;
        }
    }, 50);
}

// ─── MPC 솔버: 매일 DP(horizon일 앞) + 1일 실제실행 반복 ───
function solveMPC(horizon, onProgress) {
    const days = buildDailyData();
    if (!days || days.length === 0) return null;

    // 설정값 (solveDP와 동일)
    const chpLoadTable = getChpLoadTable();
    const plbCount = parseInt(document.getElementById('plbCount')?.value) || 3;
    const plbTable = getPlbLoadTable();
    const plbMaxCap = plbTable.length > 0 ? plbTable[0].열Gcal : 68.8;
    const storageCap = parseFloat(document.getElementById('storageCapacity')?.value) || 1500;
    const storageInit = parseFloat(document.getElementById('storageInitial')?.value) || 750;
    const storageMin = parseFloat(document.getElementById('storageMinLevel')?.value) || 100;
    const minRunTime = parseFloat(document.getElementById('chpMinRunHours')?.value) || 4;
    const minStopTime = parseFloat(document.getElementById('chpMinStopHours')?.value) || 4;
    const chpMinLoad = parseFloat(document.getElementById('chpMinLoad')?.value) || 60;
    const plbMinRunHours = parseFloat(document.getElementById('plbMinRunHours')?.value) || 2;

    const startupTimes = [0,
        parseFloat(document.getElementById('hotTime')?.value) || 4,
        parseFloat(document.getElementById('warmTime')?.value) || 6,
        parseFloat(document.getElementById('coldTime')?.value) || 8];
    const startupGasRates = [0,
        parseFloat(document.getElementById('hotGasRate')?.value) || 300,
        parseFloat(document.getElementById('warmGasRate')?.value) || 350,
        parseFloat(document.getElementById('coldGasRate')?.value) || 400];
    function calcStartupCost(stopLevel, ngCostNm3) {
        if (stopLevel <= 0) return 0;
        return startupTimes[stopLevel] * startupGasRates[stopLevel] * ngCostNm3;
    }

    const STEP = 10;
    const numStates = Math.floor((storageCap - storageMin) / STEP) + 1;
    const stateToStorage = j => storageMin + j * STEP;
    const storageToState = s => Math.min(numStates - 1, Math.max(0, Math.round((s - storageMin) / STEP)));
    const NUM_STOP = 4;
    const totalStates = numStates * NUM_STOP;
    const dpIdx = (j, k) => j * NUM_STOP + k;

    const perfOptions = chpLoadTable.filter(c => c.load >= chpMinLoad).map(c => ({
        load: c.load, heatRate: c.열Gcal, powerMW: c.송전MW, ngNm3: c.ngNm3,
    }));
    const plbRef = plbTable.length > 0 ? plbTable[0] : null;
    const plb1DayCap = plbMaxCap * 24 * plbCount;

    // 일별 비용 사전계산
    const dayInfo = days.map((d, i) => {
        const mpAvg = d.mpH.reduce((s, v) => s + v, 0) / 24;
        const costs = perfOptions.map(p => {
            const ngCostH = d.ngCostNm3 * p.ngNm3;
            const elecRevH = mpAvg * p.powerMW * 1000;
            return { load: p.load, heatRate: p.heatRate, netCostH: ngCostH - elecRevH, ngCostH, elecRevH, ngNm3: p.ngNm3, powerMW: p.powerMW };
        });
        const plbCostPerGcal = plbRef ? (d.ngPLBNm3 * plbRef.ngNm3) / plbRef.열Gcal : 99999;
        return { ...d, idx: i, costs, plbCostPerGcal };
    });

    const N = days.length;
    const results = [];
    let storageLevel = storageInit;
    let sk = 3; // Cold start
    let prevChpOn = false;
    let prevChpEndH = 0;
    let heurStopDays = 99;

    // 시간 후보
    const hourCandidates = [0];
    for (let h = Math.max(1, minRunTime); h <= 24; h++) hourCandidates.push(h);

    // ═══ 사전검증: 각 날/축열조잔량에서 넘치지 않는 최대 CHP 시간 ═══
    // maxSafeHours[dayIdx][storageStateJ] = 넘침 없는 최대 가동시간
    function calcMaxSafeHours(di, startStorage, heatPerH) {
        const d = dayInfo[di];
        const extH = d.adjExternal / 24;
        const intH = d.adjIntecoNet / 24;
        const demH = d.demandH;
        // 24h부터 내려가며 넘침 없는 시간 찾기
        for (let tryH = 24; tryH >= 0; tryH--) {
            if (tryH > 0 && tryH < minRunTime) continue;
            // 최적 시작시간으로 시뮬 (MP 높은 시간대 우선)
            let ok = false;
            if (tryH === 0) { ok = true; }
            else if (tryH >= 24) {
                // 24h: 전 시간 가동, 넘침 확인
                let st = startStorage, overflow = false;
                for (let h = 0; h < 24; h++) {
                    st += extH + intH + heatPerH - demH[h];
                    if (st > storageCap) { overflow = true; break; }
                    if (st < storageMin) st = storageMin;
                }
                ok = !overflow;
            } else {
                // 부분 가동: 어떤 시작시간이든 넘치지 않는지
                for (let s = 0; s <= 24 - tryH; s++) {
                    let st = startStorage, overflow = false;
                    for (let h = 0; h < 24; h++) {
                        const chp = (h >= s && h < s + tryH) ? heatPerH : 0;
                        st += extH + intH + chp - demH[h];
                        if (st > storageCap) { overflow = true; break; }
                        if (st < storageMin) st = storageMin;
                    }
                    if (!overflow) { ok = true; break; }
                }
            }
            if (ok) return tryH;
        }
        return 0;
    }

    // ═══ MPC 루프: 매일 DP + 1일 실행 ═══
    for (let today = 0; today < N; today++) {
        if (onProgress && today % 5 === 0) onProgress(Math.round(today / N * 100), today + 1);

        // ── 미니 DP: today ~ min(today+horizon, N-1) 역방향 탐색 ──
        const endDay = Math.min(today + horizon, N);
        const windowSize = endDay - today;

        let dpCurr = new Float64Array(totalStates).fill(0);
        let dpNext = new Float64Array(totalStates);
        const firstDayDecisions = new Array(totalStates).fill(null).map(() => ({ hours: 0, loadIdx: 0 }));

        for (let wi = windowSize - 1; wi >= 0; wi--) {
            const di = today + wi;
            const d = dayInfo[di];
            dpNext.set(dpCurr);
            dpCurr.fill(Infinity);

            for (let j = 0; j < numStates; j++) {
                const s = stateToStorage(j);
                for (let k = 0; k < NUM_STOP; k++) {
                    const idx = dpIdx(j, k);
                    let bestCost = Infinity, bestDec = { hours: 0, loadIdx: 0 };

                    for (let li = 0; li < perfOptions.length; li++) {
                        const perf = d.costs[li];
                        // 시간별 사전검증: 이 축열조 잔량에서 넘치지 않는 최대 시간
                        const maxH = (wi === 0) ? calcMaxSafeHours(di, s, perf.heatRate) : 24;
                        for (const h of hourCandidates) {
                            if (h > maxH) continue; // 시간별 검증 통과 못하면 스킵
                            if (h > 0 && d.isMaint) continue;

                            const chpHeat = h * perf.heatRate;
                            const chpCost = h * perf.netCostH;

                            let suCost = 0, nextK;
                            if (h > 0) {
                                nextK = 0;
                                if (k > 0) suCost = calcStartupCost(k, d.ngCostNm3);
                            } else {
                                nextK = Math.min(k + 1, NUM_STOP - 1);
                            }

                            let newS = s + chpHeat - d.netRequired;

                            // PLB
                            let plbCost = 0;
                            if (newS < storageMin) {
                                const deficit = storageMin - newS;
                                const plbMinHeat = plbMinRunHours * plbMaxCap;
                                const plbHeat = Math.min(plb1DayCap, Math.max(deficit, Math.min(plbMinHeat, deficit * 2)));
                                plbCost = plbHeat * d.plbCostPerGcal;
                                newS += plbHeat;
                                if (newS < storageMin) newS = storageMin;
                            }

                            // Overflow
                            if (newS > storageCap) {
                                if (h === 0) newS = storageCap;
                                else continue;
                            }

                            const jNew = storageToState(newS);
                            const totalCost = chpCost + suCost + plbCost + dpNext[dpIdx(jNew, nextK)];

                            if (totalCost < bestCost) {
                                bestCost = totalCost;
                                bestDec = { hours: h, loadIdx: li };
                            }
                        }
                    }
                    dpCurr[idx] = bestCost;
                    if (wi === 0) firstDayDecisions[idx] = bestDec;
                }
            }
        }

        // ── 오늘의 최적 결정 추출 ──
        const sj = storageToState(storageLevel);
        const dec = firstDayDecisions[dpIdx(sj, sk)];
        let chpHours = dec.hours;
        const perf = dayInfo[today].costs[dec.loadIdx];
        const chpLoadPct = perfOptions[dec.loadIdx].load;
        const chpHeatPerH = perf.heatRate;
        const d = dayInfo[today];

        // 기동비용
        let startupCost = 0;
        if (chpHours > 0 && sk > 0) {
            startupCost = Math.round(calcStartupCost(sk, d.ngCostNm3));
        }
        if (chpHours > 0) sk = 0; else sk = Math.min(sk + 1, NUM_STOP - 1);

        // ── Stage 2: 블록 배치 (overflow 거부) ──
        const hourlyExt = d.adjExternal / 24;
        const intecoH = d.adjIntecoNet / 24;
        const demandH = d.demandH;
        const smpH = d.smpH || Array(24).fill(d.smpAvg);

        let chpStart = 0;
        if (chpHours > 0 && chpHours < 24) {
            let bestScore = -Infinity;
            let tryHours = chpHours;
            while (tryHours >= minRunTime) {
                bestScore = -Infinity;
                for (let s = 0; s <= 24 - tryHours; s++) {
                    const gapFromPrev = s + (24 - prevChpEndH);
                    if (prevChpOn && gapFromPrev > 0 && gapFromPrev < minStopTime) continue;
                    let stSim = storageLevel, smpSum = 0, overflow = 0, underflow = 0;
                    for (let h = 0; h < 24; h++) {
                        const chp = (h >= s && h < s + tryHours) ? chpHeatPerH : 0;
                        stSim += hourlyExt + intecoH + chp - demandH[h];
                        if (stSim > storageCap) { overflow += stSim - storageCap; stSim = storageCap; }
                        if (stSim < storageMin) { underflow += storageMin - stSim; stSim = storageMin; }
                        if (h >= s && h < s + tryHours) smpSum += d.mpH[h];
                    }
                    if (overflow > 0) continue;
                    const score = smpSum - underflow * 200;
                    if (score > bestScore) { bestScore = score; chpStart = s; chpHours = tryHours; }
                }
                if (bestScore > -Infinity) break;
                tryHours--;
            }
            if (bestScore === -Infinity) chpHours = 0;
        }

        // ── Stage 3: PLB 투입 ──
        let stCheck = storageLevel, needPlb = false, minSt = storageLevel, minStH = 0;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours && chpHours > 0)) ? chpHeatPerH : 0;
            stCheck += hourlyExt + intecoH + chp - demandH[h];
            if (stCheck > storageCap) stCheck = storageCap;
            if (stCheck < minSt) { minSt = stCheck; minStH = h; }
            if (stCheck < storageMin) needPlb = true;
        }

        let plbHeat = 0, plbUnits = 0, plbHours = 0, plbNg = 0, plbLoad = 0, plbStartH = 0;
        if (needPlb) {
            const deficit = storageMin - minSt;
            if (deficit > 0) {
                const bestPlb = plbTable[0];
                plbUnits = Math.ceil(deficit / (bestPlb.열Gcal * Math.max(plbMinRunHours, 1)));
                plbUnits = Math.min(plbUnits, plbCount);
                if (plbUnits < 1) plbUnits = 1;
                plbLoad = bestPlb.load;
                plbHours = Math.ceil(deficit / (plbUnits * bestPlb.열Gcal));
                plbHours = Math.max(plbMinRunHours, Math.min(24 - chpHours, plbHours));
                const actualHeat = plbHours * plbUnits * bestPlb.열Gcal;
                if (actualHeat > deficit * 2 && plbHours > plbMinRunHours) {
                    plbHours = Math.max(plbMinRunHours, Math.ceil(deficit * 2 / (plbUnits * bestPlb.열Gcal)));
                }
                plbHeat = plbHours * plbUnits * bestPlb.열Gcal;
                plbNg = plbHours * plbUnits * bestPlb.ngNm3;
                plbStartH = Math.max(0, minStH - plbHours);
            }
        }

        // ── Stage 4: 시간별 시뮬 ──
        const hourly = [];
        let stH = storageLevel;
        for (let h = 0; h < 24; h++) {
            const chp = (chpHours >= 24 || (h >= chpStart && h < chpStart + chpHours && chpHours > 0)) ? chpHeatPerH : 0;
            const plb = (plbHours > 0 && h >= plbStartH && h < plbStartH + plbHours) ? (plbHeat / plbHours) : 0;
            let actualChp = chp;
            const projected = stH + hourlyExt + intecoH + actualChp + plb - demandH[h];
            if (projected > storageCap && actualChp > 0) {
                actualChp = Math.max(0, storageCap - stH - hourlyExt - intecoH - plb + demandH[h]);
            }
            stH += hourlyExt + intecoH + actualChp + plb - demandH[h];
            if (stH > storageCap) stH = storageCap;
            if (stH < storageMin) stH = storageMin;
            const chpPowerKWh = (actualChp > 0) ? perf.powerMW * 1000 * (actualChp / (chpHeatPerH || 1)) : 0;
            hourly.push({ hour: h, demand: demandH[h], ext: hourlyExt + intecoH, chp: actualChp, plb, storage: Math.round(stH), smp: smpH[h], mp: d.mpH[h], chpPowerKWh });
        }
        storageLevel = stH;

        // 집계
        const chpHeat = hourly.reduce((s, hr) => s + hr.chp, 0);
        const chpPower = hourly.reduce((s, hr) => s + hr.chpPowerKWh, 0) / 1000; // MWh
        const chpNg = chpHours * perf.ngNm3;
        const chpFuelCost = Math.round(chpHours * perf.ngCostH);
        let chpElecRev = 0;
        hourly.forEach(hr => { if (hr.chp > 0) chpElecRev += hr.mp * perf.powerMW * 1000 * (hr.chp / (chpHeatPerH || 1)); });
        chpElecRev = Math.round(chpElecRev);
        const chpNetCost = chpFuelCost - chpElecRev;
        const chpCostPerGcal = chpHeat > 0 ? Math.round(chpNetCost / chpHeat) : 0;
        const plbFuelCost = Math.round(plbNg * d.ngPLBNm3);

        prevChpOn = chpHours > 0;
        prevChpEndH = chpHours > 0 ? chpStart + chpHours : prevChpEndH;
        if (chpHours > 0) heurStopDays = 0; else heurStopDays++;

        results.push({
            month: d.month, day: d.day, dow: d.dow, dowName: d.dowName,
            isWeekend: d.isWeekend, isHoliday: d.isHoliday, isMaint: d.isMaint,
            demand: d.demand, adjExternal: d.adjExternal, adjIntecoNet: d.adjIntecoNet,
            netRequired: d.netRequired, smpAvg: d.smpAvg, smpH,
            chpHours, chpHeat, chpPower, chpNg, chpLoad: chpLoadPct,
            plbUnits, plbHours, plbHeat, plbNg, plbLoad,
            storageLevel: Math.round(stH), hourly,
            chpFuelCost, chpElecRev, chpNetCost, startupCost,
            plbFuelCost, totalFuelCost: chpFuelCost + plbFuelCost + startupCost,
            totalNetCost: chpNetCost + plbFuelCost + startupCost,
            chpCostPerGcal, plbCostPerGcal: plbHeat > 0 ? Math.round(plbFuelCost / plbHeat) : 0,
            dayType: d.isMaint ? '정비' : (d.isHoliday ? '공휴일' : (d.isWeekend ? '주말' : '평일')),
        });
    }

    if (onProgress) onProgress(100, N);
    return results;
}

// ─── 월별 요약 집계 ───
function aggregateMonthly(simResults) {
    const months = Array.from({ length: 12 }, (_, i) => ({
        month: i + 1, days: 0, netRequired: 0,
        chpDays: 0, chpHours: 0, chpHeat: 0, chpPower: 0, chpNg: 0,
        plbDays: 0, plbHours: 0, plbHeat: 0, plbNg: 0, plbMaxUnits: 0,
        storageEnd: 0,
        chpFuelCost: 0, chpElecRev: 0, chpNetCost: 0, startupCost: 0, plbFuelCost: 0, totalNetCost: 0,
        cpCharge: 0, weekdays: 0, holidays: 0,
    }));
    simResults.forEach(r => {
        const b = months[r.month - 1];
        b.days++;
        b.netRequired += r.netRequired;
        if (r.chpHours > 0) { b.chpDays++; b.chpHours += r.chpHours; b.chpHeat += r.chpHeat; b.chpPower += r.chpPower; b.chpNg += r.chpNg; }
        if (r.plbHeat > 0) { b.plbDays++; b.plbHours += r.plbHours; b.plbHeat += r.plbHeat; b.plbNg += r.plbNg; b.plbMaxUnits = Math.max(b.plbMaxUnits, r.plbUnits); }
        b.storageEnd = r.storageLevel;
        b.chpFuelCost += r.chpFuelCost || 0;
        b.chpElecRev += r.chpElecRev || 0;
        b.chpNetCost += r.chpNetCost || 0;
        b.startupCost += r.startupCost || 0;
        b.plbFuelCost += r.plbFuelCost || 0;
        b.totalNetCost += r.totalNetCost || 0;
        // 평일/휴일 카운트 (토+일+공휴일 = 휴일)
        if (r.isWeekend || r.isHoliday) b.holidays++;
        else b.weekdays++;
    });
    // 용량요금 계산: 평일수 × 평일CP + 휴일수 × 휴일CP
    if (DATA.cp_monthly) {
        const cpH = DATA.cp_monthly.headers;
        const cpWdIdx = cpH.indexOf('평일CP');
        const cpHolIdx = cpH.indexOf('휴일CP');
        if (cpWdIdx >= 0) {
            DATA.cp_monthly.rows.forEach(row => {
                const mm = Number(row[0]);
                if (mm >= 1 && mm <= 12) {
                    const b = months[mm - 1];
                    const wdCP = Number(row[cpWdIdx]) || 0;
                    const holCP = cpHolIdx >= 0 ? (Number(row[cpHolIdx]) || 0) : 0;
                    b.cpCharge = b.weekdays * wdCP + b.holidays * holCP;
                }
            });
        }
    }
    return months;
}

// ─── 월별 기동계획 테이블 렌더링 ───
function renderPlanMonthly(simResults) {
    // 열수급현황은 항상 pv_plan_monthly에 렌더링
    renderHeatBalance();

    // 시뮬레이션 결과 → pv_plan_monthly_sim에 렌더링
    const el = document.getElementById('pv_plan_monthly_sim');
    if (!el) return;

    const monthly = aggregateMonthly(simResults);
    const fmt = v => fmtNum(Math.round(v));

    let html = '';
    html += '<h4 style="font-size:14px;font-weight:700;margin-bottom:8px">월별 기동계획 요약</h4>';
    html += '<div style="overflow:auto">';
    html += '<table class="pv-tbl"><thead><tr>';
    html += '<th>월</th><th style="background:#fee2e2">필요열량<br>(손실포함)</th>';
    html += '<th style="background:#dbeafe">CHP<br>가동일</th><th style="background:#dbeafe">CHP<br>가동h</th><th style="background:#dbeafe">CHP<br>생산열</th><th style="background:#dbeafe">CHP<br>발전(MWh)</th>';
    html += '<th style="background:#fef3c7">PLB<br>가동일</th><th style="background:#fef3c7">PLB<br>가동h</th><th style="background:#fef3c7">PLB<br>생산열</th><th style="background:#fef3c7">PLB<br>최대대수</th>';
    html += '<th style="background:#e8eaf6">축열조<br>말잔량</th>';
    html += '<th style="background:#fce4ec">CHP연료비<br>(백만원)</th>';
    html += '<th style="background:#fce4ec">기동비<br>(백만원)</th>';
    html += '<th style="background:#e8f5e9">전력판매<br>(백만원)</th>';
    html += '<th style="background:#fce4ec" class="has-tip" data-tip="CHP단가 = (연료비 + 기동비 − 전력판매) / 생산열량|열 1Gcal을 생산하는 순비용.|이 단가 이상으로 열을 판매해야 마진이 남음.|용량요금은 보일러 보유에 따른 별도 수익으로 제외.">CHP단가<br>(원/Gcal)</th>';
    html += '<th style="background:#fef3c7">PLB연료비<br>(백만원)</th>';
    html += '<th style="background:#fef3c7" class="has-tip" data-tip="PLB단가 = 연료비 / 생산열량|열 원가 그 자체.|전기 판매가 없어 CHP보다 항상 높음.">PLB단가<br>(원/Gcal)</th>';
    html += '<th style="background:#fce4ec">순비용<br>(백만원)</th>';
    html += '</tr></thead><tbody>';

    const fmtM = v => fmtNum(Math.round(v / 1000000)); // 백만원
    const totals = { netReq: 0, chpD: 0, chpH: 0, chpHeat: 0, chpPow: 0, plbD: 0, plbH: 0, plbHeat: 0, plbMax: 0,
                     chpFuelCost: 0, plbFuelCost: 0, startupCost: 0, elecRev: 0, cpCharge: 0, netCost: 0, chpNetCost: 0 };

    monthly.forEach(b => {
        totals.netReq += b.netRequired; totals.chpD += b.chpDays; totals.chpH += b.chpHours;
        totals.chpHeat += b.chpHeat; totals.chpPow += b.chpPower;
        totals.plbD += b.plbDays; totals.plbH += b.plbHours; totals.plbHeat += b.plbHeat;
        totals.plbMax = Math.max(totals.plbMax, b.plbMaxUnits);
        totals.chpFuelCost += b.chpFuelCost || 0;
        totals.plbFuelCost += b.plbFuelCost || 0;
        totals.startupCost += b.startupCost || 0;
        totals.elecRev += b.chpElecRev;
        totals.cpCharge += b.cpCharge;
        totals.netCost += b.totalNetCost;
        totals.chpNetCost += b.chpNetCost;

        const plbColor = b.plbDays > 0 ? 'color:#dc2626;font-weight:700' : '';
        const ncColor = b.totalNetCost < 0 ? 'color:#1565c0' : 'color:#c62828';
        html += '<tr>';
        html += `<td>${b.month}월</td>`;
        html += `<td style="background:#fef2f2;font-weight:700">${fmt(b.netRequired)}</td>`;
        html += `<td style="background:#eff6ff">${b.chpDays}/${b.days}</td>`;
        html += `<td style="background:#eff6ff">${fmt(b.chpHours)}</td>`;
        html += `<td style="background:#eff6ff;font-weight:700">${fmt(b.chpHeat)}</td>`;
        html += `<td style="background:#eff6ff">${fmt(b.chpPower)}</td>`;
        html += `<td style="background:#fffbeb;${plbColor}">${b.plbDays}</td>`;
        html += `<td style="background:#fffbeb;${plbColor}">${fmt(b.plbHours)}</td>`;
        html += `<td style="background:#fffbeb;${plbColor}">${fmt(b.plbHeat)}</td>`;
        html += `<td style="background:#fffbeb;${plbColor}">${b.plbMaxUnits || '-'}</td>`;
        html += `<td style="background:#e8eaf6">${fmt(b.storageEnd)}</td>`;
        html += `<td style="background:#fef0f5">${fmtM(b.chpFuelCost)}</td>`;
        html += `<td style="background:#fef0f5">${(b.startupCost || 0) > 0 ? fmtM(b.startupCost) : '-'}</td>`;
        html += `<td style="background:#f0fdf4;color:#1565c0">${fmtM(b.chpElecRev)}</td>`;
        const chpUnitCost = b.chpHeat > 0 ? Math.round((b.chpFuelCost + (b.startupCost || 0) - b.chpElecRev) / b.chpHeat) : 0;
        const ucColor = chpUnitCost < 0 ? 'color:#1565c0' : 'color:#c62828';
        html += `<td style="background:#fef0f5;font-weight:700;${ucColor}">${b.chpHeat > 0 ? fmtNum(chpUnitCost) : '-'}</td>`;
        html += `<td style="background:#fffbeb">${fmtM(b.plbFuelCost)}</td>`;
        const plbUnitCost = b.plbHeat > 0 ? Math.round(b.plbFuelCost / b.plbHeat) : 0;
        html += `<td style="background:#fffbeb;font-weight:700;color:#e65100">${b.plbHeat > 0 ? fmtNum(plbUnitCost) : '-'}</td>`;
        html += `<td style="background:#fef0f5;font-weight:700;${ncColor}">${fmtM(b.totalNetCost)}</td>`;
        html += '</tr>';
    });

    // 합계
    html += '<tr style="border-top:2px solid #1e293b;font-weight:800">';
    html += '<td>합계</td>';
    html += `<td style="background:#fef2f2">${fmt(totals.netReq)}</td>`;
    html += `<td style="background:#eff6ff">${totals.chpD}</td>`;
    html += `<td style="background:#eff6ff">${fmt(totals.chpH)}</td>`;
    html += `<td style="background:#eff6ff">${fmt(totals.chpHeat)}</td>`;
    html += `<td style="background:#eff6ff">${fmt(totals.chpPow)}</td>`;
    html += `<td style="background:#fffbeb">${totals.plbD}</td>`;
    html += `<td style="background:#fffbeb">${fmt(totals.plbH)}</td>`;
    html += `<td style="background:#fffbeb">${fmt(totals.plbHeat)}</td>`;
    html += `<td style="background:#fffbeb">${totals.plbMax || '-'}</td>`;
    html += `<td style="background:#e8eaf6">${fmt(monthly[11].storageEnd)}</td>`;
    const totNcColor = totals.netCost < 0 ? 'color:#1565c0' : 'color:#c62828';
    html += `<td style="background:#fce4ec">${fmtM(totals.chpFuelCost)}</td>`;
    html += `<td style="background:#fce4ec">${totals.startupCost > 0 ? fmtM(totals.startupCost) : '-'}</td>`;
    html += `<td style="background:#dcfce7;color:#1565c0">${fmtM(totals.elecRev)}</td>`;
    const totChpUnit = totals.chpHeat > 0 ? Math.round((totals.chpFuelCost + totals.startupCost - totals.elecRev) / totals.chpHeat) : 0;
    const totUcColor = totChpUnit < 0 ? 'color:#1565c0' : 'color:#c62828';
    html += `<td style="background:#fce4ec;font-weight:700;${totUcColor}">${totals.chpHeat > 0 ? fmtNum(totChpUnit) : '-'}</td>`;
    html += `<td style="background:#fef3c7">${fmtM(totals.plbFuelCost)}</td>`;
    const totPlbUnit = totals.plbHeat > 0 ? Math.round(totals.plbFuelCost / totals.plbHeat) : 0;
    html += `<td style="background:#fef3c7;font-weight:700;color:#e65100">${totals.plbHeat > 0 ? fmtNum(totPlbUnit) : '-'}</td>`;
    html += `<td style="background:#fce4ec;font-weight:700;${totNcColor}">${fmtM(totals.netCost)}</td>`;
    html += '</tr></tbody></table></div>';

    // 차트 영역 (테이블 아래)
    html += '<div id="monthlyChartWrap" class="chart-row" style="margin-top:16px">';
    html += '<div class="chart-wrap"><div class="chart-title">월별 열생산</div><canvas id="chartMonthlyHeat"></canvas></div>';
    html += '</div>';

    el.innerHTML = html;
}

// ─── 공헌이익 차트: 베지어 라인 ───
function drawCbLineChart(canvasId, data, opts = {}) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    const cssW = Math.max(canvas.parentElement.clientWidth - 40, 200);
    const cssH = opts.height || 280;
    canvas.width = cssW * dpr; canvas.height = cssH * dpr;
    canvas.style.width = cssW + 'px'; canvas.style.height = cssH + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    const pad = { top: 30, right: 20, bottom: 40, left: 56 };
    const cw = cssW - pad.left - pad.right;
    const ch = cssH - pad.top - pad.bottom;
    const color = opts.color || '#10b981';
    const maxVal = Math.max(...data, 1);
    const minVal = Math.min(...data);
    let yMin = 0;
    if (minVal < 0) {
        const negMag = Math.pow(10, Math.floor(Math.log10(Math.abs(minVal) || 1)));
        yMin = -Math.ceil(Math.abs(minVal) / negMag) * negMag;
    } else if (minVal > maxVal * 0.6) {
        const range = maxVal - minVal || 1;
        const mag = Math.pow(10, Math.floor(Math.log10(range)));
        yMin = Math.floor(minVal / mag) * mag;
        if (yMin > minVal * 0.95) yMin = Math.floor(minVal * 0.9 / mag) * mag;
    }
    const yRange = maxVal - yMin || 1;
    const mag = Math.pow(10, Math.floor(Math.log10(yRange)));
    const yMax = yMin + Math.ceil(yRange / mag) * mag;
    const getX = i => pad.left + (cw / 11) * i;
    const getY = v => pad.top + ch - ((v - yMin) / (yMax - yMin)) * ch;
    ctx.save();
    ctx.beginPath(); ctx.roundRect(pad.left, pad.top, cw, ch, 6); ctx.clip();
    ctx.fillStyle = '#1a1a1a'; ctx.fillRect(pad.left, pad.top, cw, ch);
    ctx.setLineDash([3, 3]); ctx.strokeStyle = 'rgba(255,255,255,.06)'; ctx.lineWidth = 0.6;
    for (let i = 1; i <= 4; i++) { const y = pad.top + ch - (ch * i / 5); ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + cw, y); ctx.stroke(); }
    ctx.setLineDash([]);
    if (yMin < 0 && yMax > 0) { const zeroY = getY(0); ctx.strokeStyle = '#666'; ctx.lineWidth = 1; ctx.setLineDash([5, 3]); ctx.beginPath(); ctx.moveTo(pad.left, zeroY); ctx.lineTo(pad.left + cw, zeroY); ctx.stroke(); ctx.setLineDash([]); }
    ctx.restore();
    ctx.font = '10px Pretendard,sans-serif'; ctx.textAlign = 'right'; ctx.fillStyle = '#999';
    for (let i = 0; i <= 5; i++) { const val = Math.round(yMin + (yMax - yMin) * i / 5); ctx.fillText(val.toLocaleString(), pad.left - 8, pad.top + ch - (ch * i / 5) + 3); }
    ctx.textAlign = 'center'; ctx.fillStyle = '#999'; ctx.font = '10px Pretendard,sans-serif';
    for (let m = 1; m <= 12; m++) ctx.fillText(m + '월', getX(m - 1), cssH - pad.bottom + 18);
    const pts = data.map((v, i) => ({ x: getX(i), y: getY(v), v }));
    function bezierPath(isArea) {
        ctx.beginPath();
        if (isArea) { ctx.moveTo(pts[0].x, getY(Math.max(yMin, 0))); ctx.lineTo(pts[0].x, pts[0].y); } else ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) { const cpx = (pts[i - 1].x + pts[i].x) / 2; ctx.bezierCurveTo(cpx, pts[i - 1].y, cpx, pts[i].y, pts[i].x, pts[i].y); }
        if (isArea) { ctx.lineTo(pts[pts.length - 1].x, getY(Math.max(yMin, 0))); ctx.closePath(); }
    }
    const grad = ctx.createLinearGradient(0, pad.top, 0, pad.top + ch);
    grad.addColorStop(0, (opts.gradientTop || color) + '22'); grad.addColorStop(1, (opts.gradientTop || color) + '02');
    bezierPath(true); ctx.fillStyle = grad; ctx.fill();
    bezierPath(false); ctx.strokeStyle = color + '30'; ctx.lineWidth = 8; ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.stroke();
    bezierPath(false); ctx.strokeStyle = color; ctx.lineWidth = 2.5; ctx.stroke();
    pts.forEach(p => { ctx.beginPath(); ctx.arc(p.x, p.y, 9, 0, Math.PI * 2); ctx.fillStyle = color + '15'; ctx.fill(); ctx.beginPath(); ctx.arc(p.x, p.y, 4.5, 0, Math.PI * 2); ctx.fillStyle = '#1a1a1a'; ctx.fill(); ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke(); });
    if (opts.showValues !== false) { ctx.font = 'bold 9px Pretendard,sans-serif'; ctx.textAlign = 'center'; ctx.fillStyle = '#d0d0d0'; pts.forEach(p => { const lbl = opts.fmtVal ? opts.fmtVal(p.v) : Math.round(p.v).toLocaleString(); ctx.fillText(lbl, p.x, p.y - 16); }); }
}

// ─── 공헌이익 차트: 도넛 ───
// 도넛 차트 인스턴스 저장
const _cbDonutCharts = {};
function drawCbDonutChart(canvasId, items, unit, div) {
    if (_cbDonutCharts[canvasId]) { _cbDonutCharts[canvasId].destroy(); }
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const total = items.reduce((s, it) => s + it.value, 0);
    if (total === 0) return;
    const ctx = canvas.getContext('2d');
    _cbDonutCharts[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: items.map(it => it.label),
            datasets: [{
                data: items.map(it => Math.round(it.value / div)),
                backgroundColor: items.map(it => it.color),
                borderColor: '#141414',
                borderWidth: 2,
                hoverOffset: 6
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: true, aspectRatio: 1.8,
            cutout: '62%',
            layout: { padding: { top: 4, bottom: 4, left: 0, right: 0 } },
            plugins: {
                legend: {
                    position: 'right',
                    labels: { color: '#aaa', font: { size: 10, family: "'Pretendard',sans-serif" }, padding: 10, boxWidth: 16, boxHeight: 5, useBorderRadius: true, borderRadius: 3 }
                },
                tooltip: {
                    backgroundColor: 'rgba(20,20,20,.92)', titleColor: '#e0e0e0', bodyColor: '#d0d0d0',
                    padding: 10, cornerRadius: 8, bodyFont: { size: 11 },
                    callbacks: {
                        label: function(ctx) {
                            const v = ctx.parsed;
                            const pct = total > 0 ? (items[ctx.dataIndex].value / total * 100).toFixed(1) : '0';
                            return ` ${ctx.label}: ${v.toLocaleString()} ${unit} (${pct}%)`;
                        }
                    }
                }
            }
        },
        plugins: [{
            id: 'centerText',
            afterDraw(chart) {
                const { ctx: c, chartArea } = chart;
                const cx = (chartArea.left + chartArea.right) / 2;
                const cy = (chartArea.top + chartArea.bottom) / 2;
                c.save();
                c.fillStyle = '#e0e0e0'; c.font = "bold 14px Pretendard,sans-serif"; c.textAlign = 'center'; c.textBaseline = 'middle';
                c.fillText(Math.round(total / div).toLocaleString(), cx, cy - 7);
                c.fillStyle = '#888'; c.font = "9px Pretendard,sans-serif";
                c.fillText(unit, cx, cy + 7);
                c.restore();
            }
        }]
    });
}

// ─── 운영손익 탭: 공헌이익 렌더링 ───
function renderContribution(simResults) {
    const el = document.getElementById('pv_res_contribution');
    if (!el) return;
    if (!simResults || simResults.length === 0) {
        el.innerHTML = '<p class="hint">기동계획 실행 후 손익이 표시됩니다</p>';
        return;
    }

    const monthly = aggregateMonthly(simResults);

    // ── 월별 데이터 수집 ──
    const wcMonthly = calcWaterChemMonthly();
    const extHeatMonthly = calcExternalHeatMonthly();
    const heatSalesMonthly = calcHeatSalesMonthly();
    const elecCostMonthly = simResults
        ? calcElecCostMonthly({ source: 'planned', simResults })
        : calcElecCostMonthly({ source: 'actual' });

    // ── 월별 종합 데이터 ──
    const mData = {};
    let yearRev = 0, yearCost = 0;

    for (let m = 1; m <= 12; m++) {
        const b = monthly[m - 1];
        // 전기판매 = 열제약매출 + 용량요금
        const elecSales = (b.chpElecRev || 0) + (b.cpCharge || 0);
        // 열판매 = 기본료 + 사용요금(주택용+업무용+공공용+인테코판매)
        const heatSales = heatSalesMonthly ? (heatSalesMonthly[m]?.total || 0) : 0;
        const revenue = elecSales + heatSales;

        // 재료비
        const gasCost = (b.chpFuelCost || 0) + (b.plbFuelCost || 0);
        const elecCost = elecCostMonthly ? (elecCostMonthly[m]?.total || 0) : 0;
        const extHeatCost = extHeatMonthly ? (extHeatMonthly[m] || 0) : 0;
        const waterCost = wcMonthly ? (wcMonthly[m]?.waterCost || 0) : 0;
        const chemCost = wcMonthly ? (wcMonthly[m]?.chemTotal || 0) : 0;
        const matTotal = gasCost + elecCost + extHeatCost + waterCost + chemCost;

        const profit = revenue - matTotal;
        const rate = revenue !== 0 ? (profit / revenue * 100) : 0;

        mData[m] = { elecSales, heatSales, revenue, gasCost, elecCost, extHeatCost, waterCost, chemCost, matTotal, profit, rate };
        yearRev += revenue;
        yearCost += matTotal;
    }

    if (yearRev === 0 && yearCost === 0) {
        el.innerHTML = '<p class="hint">매출 또는 비용 데이터가 없습니다</p>';
        return;
    }

    const yearProfit = yearRev - yearCost;
    const yearRate = yearRev !== 0 ? (yearProfit / yearRev * 100) : 0;
    const useThousand = Math.max(yearRev, yearCost) > 1_000_000;
    const unit = useThousand ? '천원' : '원';
    const div = useThousand ? 1000 : 1;
    const fmtCard = v => Math.round(v / div).toLocaleString();

    const ICO = {
        rev: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg>',
        cost: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>',
        profit: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="6"/><path d="M15.477 12.89L17 22l-5-3-5 3 1.523-9.11"/></svg>',
        rate: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/></svg>',
    };

    let yElec = 0, yHeat = 0;
    for (let m = 1; m <= 12; m++) { yElec += mData[m].elecSales; yHeat += mData[m].heatSales; }
    const costRatio = yearRev !== 0 ? (yearCost / yearRev * 100).toFixed(1) : '0.0';

    // ── 요약 카드 ──
    let html = '<div class="cb-cards">';
    html += `<div class="cb-card rev"><div class="cb-icon">${ICO.rev}</div>`;
    html += `<div class="cb-label">연간 매출</div><div class="cb-value">${fmtCard(yearRev)}<small>${unit}</small></div>`;
    html += `<div class="cb-sub">전기 ${fmtCard(yElec)} · 열 ${fmtCard(yHeat)}</div></div>`;
    html += `<div class="cb-card cost"><div class="cb-icon">${ICO.cost}</div>`;
    html += `<div class="cb-label">연간 재료비</div><div class="cb-value">${fmtCard(yearCost)}<small>${unit}</small></div>`;
    html += `<div class="cb-sub">매출 대비 <span class="cb-tag neg">${costRatio}%</span></div></div>`;
    html += `<div class="cb-card profit"><div class="cb-icon">${ICO.profit}</div>`;
    html += `<div class="cb-label">연간 공헌이익</div><div class="cb-value">${fmtCard(yearProfit)}<small>${unit}</small></div>`;
    html += `<div class="cb-sub"><span class="cb-tag ${yearProfit >= 0 ? 'pos' : 'neg'}">${yearProfit >= 0 ? '흑자' : '적자'}</span></div></div>`;
    html += `<div class="cb-card rate"><div class="cb-icon">${ICO.rate}</div>`;
    html += `<div class="cb-label">공헌이익률</div><div class="cb-value">${yearRate.toFixed(1)}<small>%</small></div>`;
    html += `<div class="cb-sub">(매출 − 재료비) / 매출</div></div>`;
    html += '</div>';

    // ── 차트 1행: 공헌이익 추이 + 이익률 ──
    html += '<div class="cb-chart-row">';
    html += `<div class="cb-chart-box main"><div class="cb-chart-head"><div class="cb-dot" style="background:#10b981"></div><h4>월별 공헌이익</h4><span>(${unit})</span></div><canvas id="contribProfitChart" height="360"></canvas></div>`;
    html += `<div class="cb-chart-box sub"><div class="cb-chart-head"><div class="cb-dot" style="background:#f59e0b"></div><h4>공헌이익률</h4><span>(%)</span></div><canvas id="contribRateChart" height="360"></canvas></div>`;
    html += '</div>';

    // ── 차트 2행: 매출·재료비 구성 (도넛) ──
    html += '<div class="cb-chart-row">';
    html += `<div class="cb-chart-box"><div class="cb-chart-head"><div class="cb-dot" style="background:#7BA8F5"></div><h4>연간 매출 구성</h4></div><canvas id="contribRevPieChart" height="220"></canvas></div>`;
    html += `<div class="cb-chart-box"><div class="cb-chart-head"><div class="cb-dot" style="background:#FF6B6B"></div><h4>연간 재료비 구성</h4></div><canvas id="contribCostPieChart" height="220"></canvas></div>`;
    html += '</div>';

    // ── 재료비 항목 정의 ──
    const matItems = [
        { name: '가스비', key: 'gasCost' },
        { name: '전력비', key: 'elecCost' },
        { name: '수열비', key: 'extHeatCost' },
        { name: '용수비', key: 'waterCost' },
        { name: '약품비', key: 'chemCost' },
    ];

    // ── 종합 테이블 ──
    html += '<div class="cb-table-wrap">';
    html += `<table class="cb-tbl"><caption>단위: ${unit}</caption>`;
    html += '<thead><tr>';
    html += '<th rowspan="2" style="text-align:center">월</th>';
    html += '<th colspan="3" class="hd-rev" style="text-align:center;border-bottom:1px solid #bfdbfe">매 출</th>';
    html += `<th colspan="${matItems.length + 1}" class="hd-cost" style="text-align:center;border-bottom:1px solid #fecaca">재 료 비</th>`;
    html += '<th rowspan="2" class="hd-profit">공헌이익</th>';
    html += '<th rowspan="2" class="hd-rate">이익률(%)</th>';
    html += '</tr><tr>';
    html += '<th class="hd-rev">전기판매</th><th class="hd-rev">열판매</th><th class="hd-rev-total">매출합계</th>';
    matItems.forEach(it => html += `<th class="hd-cost">${it.name}</th>`);
    html += '<th class="hd-cost-total">재료비합계</th>';
    html += '</tr></thead><tbody>';

    let sElec = 0, sHeat = 0, sRev = 0, sCost = 0, sProfit = 0;
    const sMat = {}; matItems.forEach(it => sMat[it.name] = 0);

    for (let m = 1; m <= 12; m++) {
        const d = mData[m];
        const eS = d.elecSales / div, hS = d.heatSales / div, rv = d.revenue / div;
        const mc = d.matTotal / div, pr = d.profit / div, rt = d.rate;
        sElec += eS; sHeat += hS; sRev += rv; sCost += mc; sProfit += pr;
        const pCls = pr >= 0 ? 'pos' : 'neg';
        html += `<tr><td>${m}월</td>`;
        html += `<td class="td-rev">${Math.round(eS).toLocaleString()}</td>`;
        html += `<td class="td-rev">${Math.round(hS).toLocaleString()}</td>`;
        html += `<td class="td-rev-total">${Math.round(rv).toLocaleString()}</td>`;
        matItems.forEach(it => {
            const v = d[it.key] / div;
            sMat[it.name] += v;
            html += `<td class="td-cost">${Math.round(v).toLocaleString()}</td>`;
        });
        html += `<td class="td-cost-total">${Math.round(mc).toLocaleString()}</td>`;
        html += `<td class="td-profit ${pCls}">${Math.round(pr).toLocaleString()}</td>`;
        html += `<td class="td-rate">${rt.toFixed(1)}</td>`;
        html += '</tr>';
    }

    // 합계 행
    const totalRate = sRev !== 0 ? (sProfit / sRev * 100) : 0;
    const tpCls = sProfit >= 0 ? 'pos' : 'neg';
    html += `<tr class="cb-total"><td>합계</td>`;
    html += `<td class="td-rev">${Math.round(sElec).toLocaleString()}</td>`;
    html += `<td class="td-rev">${Math.round(sHeat).toLocaleString()}</td>`;
    html += `<td class="td-rev-total">${Math.round(sRev).toLocaleString()}</td>`;
    matItems.forEach(it => html += `<td class="td-cost">${Math.round(sMat[it.name]).toLocaleString()}</td>`);
    html += `<td class="td-cost-total">${Math.round(sCost).toLocaleString()}</td>`;
    html += `<td class="td-profit ${tpCls}">${Math.round(sProfit).toLocaleString()}</td>`;
    html += `<td class="td-rate">${totalRate.toFixed(1)}</td>`;
    html += '</tr></tbody></table></div>';

    el.innerHTML = html;

    // ── 차트 렌더링 (레이아웃 완료 후) ──
    requestAnimationFrame(() => {
        const profitData = [], rateData = [];
        for (let m = 1; m <= 12; m++) { profitData.push(mData[m].profit / div); rateData.push(mData[m].rate); }
        drawCbLineChart('contribProfitChart', profitData, { color: '#10b981', gradientTop: '#10b981', height: 360 });
        drawCbLineChart('contribRateChart', rateData, { color: '#f59e0b', gradientTop: '#f59e0b', fmtVal: v => v.toFixed(1) + '%', height: 360 });

        // 매출 구성 도넛
        drawCbDonutChart('contribRevPieChart', [
            { label: '전기판매', value: yElec, color: '#7BA8F5' },
            { label: '열판매', value: yHeat, color: '#FF6B6B' },
        ], unit, div);

        // 재료비 구성 도넛
        const matColors = { '가스비': '#FFB74D', '전력비': '#3b82f6', '수열비': '#6BCB77', '용수비': '#4DD0E1', '약품비': '#A78BFA' };
        const matPieItems = matItems.map(it => {
            let total = 0;
            for (let m = 1; m <= 12; m++) total += mData[m][it.key] || 0;
            return { label: it.name, value: total, color: matColors[it.name] };
        });
        drawCbDonutChart('contribCostPieChart', matPieItems, unit, div);
    });
}

// ─── 전력수요 프로파일 구축 ───
function buildElecDemandProfile() {
    // 공통 설정값
    const standbyAdj = parseNum(document.getElementById('standbyPowerKW')?.value) || 0;
    const growthPct = parseNum(document.getElementById('demandGrowthPct')?.value) || 0;
    const growthFactor = 1 + growthPct / 100;
    const dayFactor = parseNum(document.getElementById('demandDayFactor')?.value) || 1.0;
    const nightFactor = parseNum(document.getElementById('demandNightFactor')?.value) || 1.0;
    const monthlyFactorInputs = document.querySelectorAll('#monthlyDemandFactors input');
    const monthlyFactor = Array(12).fill(1.0);
    monthlyFactorInputs.forEach((inp, i) => { monthlyFactor[i] = parseNum(inp.value) || 1.0; });

    const cap = DATA.capacity_hourly;
    const hasUpload = cap && cap.rows && cap.rows.length > 0;

    // 업로드 데이터가 없으면 설정 패턴으로 생성
    if (!hasUpload) {
        const mKwInputs = document.querySelectorAll('#monthlyDemandKW input');
        const monthlyKW = Array(12).fill(0);
        mKwInputs.forEach((inp, i) => { monthlyKW[i] = parseNum(inp.value) || 0; });
        const hasSettings = monthlyKW.some(v => v > 0);
        if (!hasSettings) return null;

        const hPatInputs = document.querySelectorAll('#hourlyDemandPattern input');
        const hourPattern = Array(24).fill(1.0);
        hPatInputs.forEach((inp, i) => { hourPattern[i] = parseNum(inp.value) || 1.0; });

        const dailyProfile = {};
        const monthlyTotal = Array(12).fill(0);
        let minHourlyKWh = Infinity;
        const year = CURRENT_YEAR;

        for (let m = 1; m <= 12; m++) {
            const daysInMonth = new Date(year, m, 0).getDate();
            const baseKW = monthlyKW[m - 1];
            for (let d = 1; d <= daysInMonth; d++) {
                const key = `${m}-${d}`;
                const profile = Array(24).fill(0);
                for (let h = 0; h < 24; h++) {
                    const timeFactor = (h >= 7 && h < 22) ? dayFactor : nightFactor;
                    let demand = Math.max(standbyAdj, baseKW * hourPattern[h]) * growthFactor * monthlyFactor[m - 1] * timeFactor;
                    profile[h] = Math.max(0, demand);
                    if (profile[h] < minHourlyKWh) minHourlyKWh = profile[h];
                }
                dailyProfile[key] = profile;
                monthlyTotal[m - 1] += profile.reduce((s, v) => s + v, 0);
            }
        }
        return { dailyProfile, standbyPowerKW: Math.max(0, minHourlyKWh), monthlyTotalKWh: monthlyTotal, source: 'settings' };
    }

    // 업로드 데이터 기반
    const heatConst = DATA.heat_constraint_hourly;
    function buildHourlyMap(data) {
        const map = {};
        if (!data || !data.rows) return map;
        data.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            map[key] = [];
            for (let h = 0; h < 24; h++) map[key][h] = Number(row[h + 2]) || 0;
        });
        return map;
    }
    const hcMap = buildHourlyMap(heatConst);

    const dailyProfile = {};
    const monthlyTotal = Array(12).fill(0);
    let minHourlyKWh = Infinity;

    cap.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        const key = `${m}-${d}`;
        const profile = Array(24).fill(0);
        for (let h = 0; h < 24; h++) {
            const capKWh = Number(row[h + 2]) || 0;
            const hcKWh = hcMap[key]?.[h] || 0;
            let demand = Math.max(standbyAdj, capKWh + hcKWh);
            const timeFactor = (h >= 7 && h < 22) ? dayFactor : nightFactor;
            demand = demand * growthFactor * monthlyFactor[m - 1] * timeFactor;
            profile[h] = Math.max(0, demand);
            if (profile[h] < minHourlyKWh) minHourlyKWh = profile[h];
        }
        dailyProfile[key] = profile;
        monthlyTotal[m - 1] += profile.reduce((s, v) => s + v, 0);
    });

    return { dailyProfile, standbyPowerKW: Math.max(0, minHourlyKWh), monthlyTotalKWh: monthlyTotal, source: 'upload' };
}

// ─── 실적에서 월별 평균 + 시간패턴 추출 ───
function applyDemandFromHistory() {
    const cap = DATA.capacity_hourly;
    const heatConst = DATA.heat_constraint_hourly;
    if (!cap || !cap.rows || cap.rows.length === 0) {
        alert('수전량 데이터를 먼저 업로드하세요');
        return;
    }
    // 월별 합계, 시간별 합계
    const monthlySum = Array(12).fill(0);
    const monthlyHours = Array(12).fill(0);
    const hourlySum = Array(24).fill(0);
    let totalHours = 0;
    const hcMap = {};
    if (heatConst && heatConst.rows) {
        heatConst.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            hcMap[key] = [];
            for (let h = 0; h < 24; h++) hcMap[key][h] = Number(row[h + 2]) || 0;
        });
    }
    cap.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        const key = `${m}-${d}`;
        const daysInMonth = new Date(CURRENT_YEAR, m, 0).getDate();
        for (let h = 0; h < 24; h++) {
            const kwh = (Number(row[h + 2]) || 0) + (hcMap[key]?.[h] || 0);
            monthlySum[m - 1] += kwh;
            hourlySum[h] += kwh;
        }
        monthlyHours[m - 1] += 24;
        totalHours += 24;
    });
    // 월별 평균 kW
    const mKwInputs = document.querySelectorAll('#monthlyDemandKW input');
    mKwInputs.forEach((inp, i) => {
        inp.value = monthlyHours[i] > 0 ? Math.round(monthlySum[i] / monthlyHours[i]) : 0;
    });
    // 시간별 패턴 (전체 평균 대비 비율)
    const avgPerHour = totalHours > 0 ? hourlySum.reduce((s, v) => s + v, 0) / totalHours : 1;
    const hPatInputs = document.querySelectorAll('#hourlyDemandPattern input');
    hPatInputs.forEach((inp, i) => {
        const avg = totalHours > 0 ? hourlySum[i] / (totalHours / 24) : 1;
        inp.value = avgPerHour > 0 ? (avg / avgPerHour).toFixed(2) : '1.0';
    });
    scheduleSettingsSave();
}

// ─── 계획 수전량 산출 ───
function calcPlannedCapacity(simResults) {
    const profile = buildElecDemandProfile();
    if (!profile) return null;

    const daily = {};
    const monthly = {};
    for (let m = 1; m <= 12; m++) monthly[m] = { buyKwh: 0, sellKwh: 0, netKwh: 0 };

    simResults.forEach(r => {
        const key = `${r.month}-${r.day}`;
        const demandH = profile.dailyProfile[key];
        if (!demandH) return;

        const kwhH = Array(24).fill(0);
        let totalBuy = 0, totalSell = 0;
        for (let h = 0; h < 24; h++) {
            const planned = demandH[h] - (r.hourly[h]?.chpPowerKWh || 0);
            kwhH[h] = planned;
            if (planned > 0) totalBuy += planned;
            else totalSell += Math.abs(planned);
        }
        daily[key] = { kwhH, buyKwh: totalBuy, sellKwh: totalSell, netKwh: totalBuy - totalSell };
        monthly[r.month].buyKwh += totalBuy;
        monthly[r.month].sellKwh += totalSell;
        monthly[r.month].netKwh += (totalBuy - totalSell);
    });

    return { daily, monthly, profile };
}

// ─── 전력비 월별 집계 (공헌이익용) ───
// options: { source: 'actual'|'planned', simResults }
function calcElecCostMonthly(options = {}) {
    const { source = 'actual', simResults = null } = options;

    const rateWd = DATA.elec_rate_weekday;
    const bc = DATA.base_charge_monthly;
    if (!rateWd || !rateWd.rows || rateWd.rows.length === 0) return null;
    if (!bc || !bc.rows || bc.rows.length === 0) return null;

    const rateSat = DATA.elec_rate_saturday;
    const rateHol = DATA.elec_rate_holiday;

    function buildRateMap(rd) {
        const rm = {}; for (let m = 1; m <= 12; m++) rm[m] = {};
        if (!rd || !rd.rows) return rm;
        rd.rows.forEach(row => {
            const h = Number(String(row[0]).replace('h', ''));
            if (h < 1 || h > 24) return;
            for (let m = 1; m <= 12; m++) rm[m][h] = Number(row[m]) || 0;
        });
        return rm;
    }
    const rateMapWd = buildRateMap(rateWd);
    const rateMapSat = buildRateMap(rateSat);
    const rateMapHol = buildRateMap(rateHol);
    function getRateMap(dt) {
        if (dt === 'saturday' && rateSat && rateSat.rows && rateSat.rows.length > 0) return rateMapSat;
        if (dt === 'holiday' && rateHol && rateHol.rows && rateHol.rows.length > 0) return rateMapHol;
        return rateMapWd;
    }

    const bcMap = {};
    bc.rows.forEach(row => {
        const m = Number(row[0]); if (m < 1 || m > 12) return;
        bcMap[m] = { base154: Number(row[1])||0, base22: Number(row[2])||0,
                     climate: Number(row[3])||0, fuel: Number(row[4])||0, fundPct: Number(row[5])||0 };
    });

    const holidaySet = new Set();
    if (DATA.holidays) DATA.holidays.rows.forEach(r => {
        if (Number(r[3]) === 1) holidaySet.add(`${Number(r[0])}-${Number(r[1])}`);
    });

    function getDayType(m, d) {
        const dow = new Date(CURRENT_YEAR, m - 1, d).getDay();
        if (dow === 0 || holidaySet.has(`${m}-${d}`)) return 'holiday';
        if (dow === 6) return 'saturday';
        return 'weekday';
    }

    const monthly = {};
    for (let m = 1; m <= 12; m++) monthly[m] = { totalKwh:0, baseFee:0, meterFee:0, climateFee:0, fuelFee:0, fundFee:0, total:0 };

    if (source === 'planned' && simResults) {
        // 계획 수전량 기반
        const planned = calcPlannedCapacity(simResults);
        if (!planned) return null;

        simResults.forEach(r => {
            const m = r.month;
            if (!bcMap[m]) return;
            const key = `${m}-${r.day}`;
            const dayData = planned.daily[key];
            if (!dayData) return;
            const dt = getDayType(m, r.day);
            const rateMap = getRateMap(dt);
            let dailyMeter = 0, dailyKwh = 0;
            for (let h = 0; h < 24; h++) {
                const kwh = Math.max(0, dayData.kwhH[h]); // 잉여(음수)는 전력비에서 제외
                dailyMeter += kwh * (rateMap[m][h + 1] || 0);
                dailyKwh += kwh;
            }
            monthly[m].totalKwh += dailyKwh;
            monthly[m].meterFee += dailyMeter;
            monthly[m].climateFee += dailyKwh * bcMap[m].climate;
            monthly[m].fuelFee += dailyKwh * bcMap[m].fuel;
        });
    } else {
        // 실적 수전량 기반
        const cap = DATA.capacity_hourly;
        if (!cap || !cap.rows || cap.rows.length === 0) return null;
        cap.rows.forEach(row => {
            const m = Number(row[0]), d = Number(row[1]);
            if (m < 1 || m > 12 || !bcMap[m]) return;
            const dt = getDayType(m, d);
            const rateMap = getRateMap(dt);
            let dailyMeter = 0, dailyKwh = 0;
            for (let h = 1; h <= 24; h++) {
                const kwh = Number(row[h + 1]) || 0; // row[2]=1h, row[3]=2h, ...
                dailyMeter += kwh * (rateMap[m][h] || 0);
                dailyKwh += kwh;
            }
            monthly[m].totalKwh += dailyKwh;
            monthly[m].meterFee += dailyMeter;
            monthly[m].climateFee += dailyKwh * bcMap[m].climate;
            monthly[m].fuelFee += dailyKwh * bcMap[m].fuel;
        });
    }

    // 기본료 + 전력기금
    for (let m = 1; m <= 12; m++) {
        if (!bcMap[m]) continue;
        const base = bcMap[m].base154 + bcMap[m].base22;
        monthly[m].baseFee = base;
        const usageSub = base + monthly[m].meterFee + monthly[m].climateFee + monthly[m].fuelFee;
        monthly[m].fundFee = usageSub * (bcMap[m].fundPct / 100);
        monthly[m].total = usageSub + monthly[m].fundFee;
    }
    return monthly;
}

// ─── 용수약품비 월별 집계 (공헌이익용) ───
function calcWaterChemMonthly() {
    if (!DATA.water_chem_daily || !DATA.water_price_monthly || !DATA.chem_price_monthly) return null;
    const waterPriceByMonth = {};
    DATA.water_price_monthly.rows.forEach(r => {
        const m = Number(r[0]);
        waterPriceByMonth[m] = parseNum(r[3]) || parseNum(r[2]) || 0;
    });
    const chemPriceByMonth = {};
    DATA.chem_price_monthly.rows.forEach(r => {
        const m = Number(r[0]);
        chemPriceByMonth[m] = { '가성소다': parseNum(r[1]), '탈산소제': parseNum(r[2]), '암모니아수': parseNum(r[3]), '염산': parseNum(r[4]), '카보하이드라자이드': parseNum(r[5]), 'PH조절제': parseNum(r[6]) };
    });
    const result = {};
    for (let m = 1; m <= 12; m++) result[m] = { waterCost: 0, chemTotal: 0, total: 0 };
    DATA.water_chem_daily.rows.forEach(row => {
        const m = Number(row[0]);
        if (m < 1 || m > 12) return;
        const waterQty = parseNum(row[2]);
        const waterCost = Math.round(waterQty * (waterPriceByMonth[m] || 0));
        const cp = chemPriceByMonth[m] || {};
        const chemKeys = ['가성소다', '탈산소제', '암모니아수', '염산', '카보하이드라자이드', 'PH조절제'];
        let chemTotal = 0;
        chemKeys.forEach((k, i) => { chemTotal += Math.round(parseNum(row[i + 3]) * (cp[k] || 0)); });
        result[m].waterCost += waterCost;
        result[m].chemTotal += chemTotal;
        result[m].total += waterCost + chemTotal;
    });
    return result;
}

// ─── 외부수열비 월별 집계 (공헌이익용) ───
function calcExternalHeatMonthly() {
    if (!DATA.operations_daily || !DATA.sale_prices_daily) return null;
    const oH = DATA.operations_daily.headers;
    const pH = DATA.sale_prices_daily.headers;
    // 외부수열 매핑: EXT_SOURCES 레지스트리 기반 (활성 소스만)
    const extMap = EXT_SOURCES.filter(s => s.enabled).map(s => ({ opsName: s.opsName, priceName: s.priceName }));
    const sources = extMap.map(e => ({ name: e.opsName, opsCol: oH.indexOf(e.opsName), priceCol: pH.indexOf(e.priceName) })).filter(s => s.opsCol >= 0);
    if (sources.length === 0) return null;

    const priceMap = {};
    DATA.sale_prices_daily.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        priceMap[`${m}-${d}`] = {};
        sources.forEach(s => { if (s.priceCol >= 0) priceMap[`${m}-${d}`][s.name] = parseNum(row[s.priceCol]); });
    });

    const result = {};
    for (let m = 1; m <= 12; m++) result[m] = 0;
    DATA.operations_daily.rows.forEach(row => {
        const m = Number(row[0]), d = Number(row[1]);
        if (m < 1 || m > 12) return;
        const prices = priceMap[`${m}-${d}`] || {};
        sources.forEach(s => { result[m] += Math.round(parseNum(row[s.opsCol]) * (prices[s.name] || 0)); });
    });
    return result;
}

// ─── 전력수요 프로파일 시각화 ───
function renderElecDemandProfile(simResults) {
    const el = document.getElementById('pv_calc_elec_demand');
    if (!el) return;
    if (!simResults || simResults.length === 0) {
        el.innerHTML = '<p class="hint">기동계획 실행 후 표시됩니다</p>';
        return;
    }
    const planned = calcPlannedCapacity(simResults);
    if (!planned) {
        el.innerHTML = '<p class="hint">설정 > 전력수요 모델링에서 월별 수요를 입력하세요</p>';
        return;
    }

    // 추정 대기전력 표시
    const estEl = document.getElementById('estimatedStandbyPower');
    if (planned.profile?.standbyPowerKW != null && estEl) {
        estEl.innerHTML = `추정 대기전력: <b style="color:#93c5fd">${planned.profile.standbyPowerKW.toFixed(1)} kW</b> (전 시간대 최소 소비전력)`;
    }

    const unit = 'MWh'; const div = 1000;
    const fmt = v => Math.round(v / div).toLocaleString();

    // 월별 집계
    const mDemand = Array(12).fill(0), mChp = Array(12).fill(0), mBuy = Array(12).fill(0), mSell = Array(12).fill(0);
    simResults.forEach(r => {
        const key = `${r.month}-${r.day}`;
        const demandH = planned.profile?.dailyProfile?.[key];
        if (!demandH) return;
        const dayDemand = demandH.reduce((s, v) => s + v, 0);
        const dayChp = r.hourly.reduce((s, h) => s + (h.chpPowerKWh || 0), 0);
        mDemand[r.month - 1] += dayDemand;
        mChp[r.month - 1] += dayChp;
        mBuy[r.month - 1] += planned.daily[key]?.buyKwh || 0;
        mSell[r.month - 1] += planned.daily[key]?.sellKwh || 0;
    });

    const yDemand = mDemand.reduce((s, v) => s + v, 0);
    const yChp = mChp.reduce((s, v) => s + v, 0);
    const yBuy = mBuy.reduce((s, v) => s + v, 0);

    let html = '<div style="margin-bottom:16px">';
    html += `<div style="font-size:14px;font-weight:700;color:#e0e4ea;margin-bottom:4px">연간 총수요: ${fmt(yDemand)} MWh · CHP자가소비: ${fmt(yDemand - yBuy)} MWh · 계획수전: ${fmt(yBuy)} MWh</div>`;
    html += `<div style="font-size:11px;color:#94a3b8">계획수전 = Σ max(0, 시간별 수요 - CHP발전) — 한전에서 구매해야 하는 전력량</div>`;
    html += '</div>';

    // 월별 테이블
    html += '<table class="pv-tbl" style="max-width:700px"><thead><tr>';
    html += '<th>월</th><th>총수요(MWh)</th><th>CHP발전(MWh)</th><th>계획수전(MWh)</th>';
    html += '</tr></thead><tbody>';
    for (let m = 0; m < 12; m++) {
        html += `<tr><td>${m+1}월</td>`;
        html += `<td style="text-align:right">${fmt(mDemand[m])}</td>`;
        html += `<td style="text-align:right;color:#93c5fd">${fmt(mChp[m])}</td>`;
        html += `<td style="text-align:right;font-weight:600;color:#f59e0b">${fmt(mBuy[m])}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td>`;
    html += `<td style="text-align:right">${fmt(yDemand)}</td>`;
    html += `<td style="text-align:right;color:#93c5fd">${fmt(yChp)}</td>`;
    html += `<td style="text-align:right;color:#f59e0b">${fmt(yBuy)}</td></tr>`;
    html += '</tbody></table>';

    // 차트
    html += '<div style="margin-top:16px"><canvas id="elecDemandChart" height="300"></canvas></div>';
    el.innerHTML = html;

    requestAnimationFrame(() => {
        const ctx = document.getElementById('elecDemandChart')?.getContext('2d');
        if (!ctx) return;
        const labels = Array.from({length:12}, (_, i) => `${i+1}월`);
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: '총수요', data: mDemand.map(v => Math.round(v / div)), backgroundColor: 'rgba(148,163,184,0.25)', borderColor: '#94a3b8', borderWidth: 1, borderRadius: 4 },
                    { label: 'CHP발전', data: mChp.map(v => Math.round(v / div)), backgroundColor: 'rgba(59,130,246,0.3)', borderColor: '#3b82f6', borderWidth: 1, borderRadius: 4 },
                    { label: '계획수전', data: mBuy.map(v => Math.round(v / div)), backgroundColor: 'rgba(251,191,36,0.3)', borderColor: '#f59e0b', borderWidth: 1, borderRadius: 4 },
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { labels: { color: '#94a3b8', boxWidth: 16, boxHeight: 5, borderRadius: 3, useBorderRadius: true } } },
                scales: {
                    y: { beginAtZero: true, title: { display: true, text: unit, color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.04)' }, ticks: { color: '#94a3b8' } },
                    x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
                }
            }
        });
    });
}

// ─── 전력비 비교 시각화 (시뮬 결과 기반) ───
function renderElecCostComparison(simResults) {
    const el = document.getElementById('pv_calc_elec_cost');
    if (!el) return;
    if (!simResults || simResults.length === 0) {
        el.innerHTML = '<p class="hint">기동계획 실행 후 전력비가 표시됩니다</p>';
        return;
    }

    const plannedCost = calcElecCostMonthly({ source: 'planned', simResults });
    if (!plannedCost) {
        el.innerHTML = '<p class="hint">전력수요 설정 + 전기요금 데이터가 필요합니다</p>';
        return;
    }

    const unit = '천원'; const div = 1000;
    const fmt = v => Math.round(v / div).toLocaleString();

    let html = '<div style="margin-bottom:12px;font-size:11px;color:#94a3b8">기동계획 기반 계획 수전량으로 산출한 전력비입니다</div>';

    // 월별 테이블
    html += '<table class="pv-tbl" style="max-width:800px"><thead><tr>';
    html += '<th>월</th><th>수전량(MWh)</th><th>기본료(천원)</th><th>수전요금(천원)</th><th>기후환경(천원)</th><th>연료비(천원)</th><th>전력기금(천원)</th><th>합계(천원)</th>';
    html += '</tr></thead><tbody>';

    let yKwh = 0, yBase = 0, yMeter = 0, yClimate = 0, yFuel = 0, yFund = 0, yTotal = 0;
    for (let m = 1; m <= 12; m++) {
        const c = plannedCost[m];
        yKwh += c.totalKwh; yBase += c.baseFee; yMeter += c.meterFee; yClimate += c.climateFee; yFuel += c.fuelFee; yFund += c.fundFee; yTotal += c.total;
        html += `<tr><td>${m}월</td>`;
        html += `<td style="text-align:right">${(c.totalKwh/1000).toFixed(0)}</td>`;
        html += `<td style="text-align:right">${fmt(c.baseFee)}</td>`;
        html += `<td style="text-align:right">${fmt(c.meterFee)}</td>`;
        html += `<td style="text-align:right">${fmt(c.climateFee)}</td>`;
        html += `<td style="text-align:right">${fmt(c.fuelFee)}</td>`;
        html += `<td style="text-align:right">${fmt(c.fundFee)}</td>`;
        html += `<td style="text-align:right;font-weight:600">${fmt(c.total)}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td>`;
    html += `<td style="text-align:right">${(yKwh/1000).toFixed(0)}</td>`;
    html += `<td style="text-align:right">${fmt(yBase)}</td>`;
    html += `<td style="text-align:right">${fmt(yMeter)}</td>`;
    html += `<td style="text-align:right">${fmt(yClimate)}</td>`;
    html += `<td style="text-align:right">${fmt(yFuel)}</td>`;
    html += `<td style="text-align:right">${fmt(yFund)}</td>`;
    html += `<td style="text-align:right">${fmt(yTotal)}</td></tr>`;
    html += '</tbody></table>';

    // 비교 차트
    html += '<div style="margin-top:16px"><canvas id="elecCostCompChart" height="300"></canvas></div>';

    // 일별 상세 조회
    html += '<div style="margin-top:20px;border-top:1px solid #333;padding-top:16px">';
    html += '<div style="font-size:13px;font-weight:700;color:#e0e4ea;margin-bottom:8px">시간별 상세 조회</div>';
    html += '<div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">';
    html += '<label style="font-size:11px;color:#999">월 <select id="elecDetailMonth" style="background:#1a1a1a;color:#ccc;border:1px solid #333;border-radius:4px;padding:3px 6px;font-size:11px">';
    for (let m = 1; m <= 12; m++) html += `<option value="${m}">${m}월</option>`;
    html += '</select></label>';
    html += '<label style="font-size:11px;color:#999">일 <input type="number" id="elecDetailDay" value="1" min="1" max="31" style="width:50px;background:#1a1a1a;color:#ccc;border:1px solid #333;border-radius:4px;padding:3px 6px;font-size:11px"></label>';
    html += '<button class="btn-sm" onclick="showElecDetailDay()">조회</button>';
    html += '</div>';
    html += '<div id="elecDetailResult"></div>';
    html += '</div>';

    el.innerHTML = html;

    // 비교 차트 그리기
    requestAnimationFrame(() => {
        const ctx = document.getElementById('elecCostCompChart')?.getContext('2d');
        if (!ctx) return;
        const labels = Array.from({length:12}, (_, i) => `${i+1}월`);
        new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: '실적 전력비', data: labels.map((_, i) => Math.round((actualCost?.[i+1]?.total || 0) / div)), backgroundColor: 'rgba(148,163,184,0.3)', borderColor: '#94a3b8', borderWidth: 1, borderRadius: 4 },
                    { label: '계획 전력비', data: labels.map((_, i) => Math.round((plannedCost?.[i+1]?.total || 0) / div)), backgroundColor: 'rgba(59,130,246,0.3)', borderColor: '#3b82f6', borderWidth: 1, borderRadius: 4 },
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: { legend: { labels: { color: '#94a3b8', boxWidth: 16, boxHeight: 5, borderRadius: 3, useBorderRadius: true } } },
                scales: {
                    y: { title: { display: true, text: unit, color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8' } },
                    x: { grid: { display: false }, ticks: { color: '#94a3b8' } }
                }
            }
        });
    });
}

// ─── 시간별 전력 상세 조회 ───
function showElecDetailDay() {
    const m = parseInt(document.getElementById('elecDetailMonth')?.value) || 1;
    const d = parseInt(document.getElementById('elecDetailDay')?.value) || 1;
    const el = document.getElementById('elecDetailResult');
    if (!el) return;

    const profile = buildElecDemandProfile();
    const key = `${m}-${d}`;
    const demandH = profile?.dailyProfile?.[key];

    // 실적 수전량
    const capRow = DATA.capacity_hourly?.rows?.find(r => Number(r[0]) === m && Number(r[1]) === d);
    // CHP 발전 (시뮬 결과)
    const simDay = PLAN_RESULTS?.find(r => r.month === m && r.day === d);

    if (!demandH && !capRow) {
        el.innerHTML = '<p class="hint">해당 일자의 데이터가 없습니다</p>';
        return;
    }

    let html = '<table class="pv-tbl" style="font-size:11px;max-width:100%"><thead><tr>';
    html += '<th>시간</th><th>총수요(kWh)</th><th>실적수전(kWh)</th>';
    if (simDay) html += '<th>CHP발전(kWh)</th><th>계획수전(kWh)</th>';
    html += '<th>요금단가(원/kWh)</th>';
    html += '</tr></thead><tbody>';

    // 요금표
    const rateWd = DATA.elec_rate_weekday;
    const buildRM = (rd) => { const rm = {}; if (!rd?.rows) return rm; rd.rows.forEach(r => { const h = Number(String(r[0]).replace('h','')); if (h >= 1 && h <= 24) rm[h] = Number(r[m]) || 0; }); return rm; };
    const rm = buildRM(rateWd);

    let totalDemand = 0, totalCap = 0, totalChp = 0, totalPlanned = 0;
    for (let h = 0; h < 24; h++) {
        const demand = demandH ? demandH[h] : 0;
        const cap = capRow ? (Number(capRow[h + 2]) || 0) : 0;
        const chpKwh = simDay?.hourly?.[h]?.chpPowerKWh || 0;
        const planned = demand - chpKwh;
        const rate = rm[h + 1] || 0;
        totalDemand += demand; totalCap += cap; totalChp += chpKwh; totalPlanned += planned;
        html += `<tr><td>${h+1}h</td><td style="text-align:right">${demand.toFixed(1)}</td><td style="text-align:right">${cap.toFixed(1)}</td>`;
        if (simDay) {
            html += `<td style="text-align:right;color:#93c5fd">${chpKwh.toFixed(1)}</td>`;
            html += `<td style="text-align:right;color:${planned > 0 ? '#FFB74D' : '#6BCB77'}">${planned.toFixed(1)}</td>`;
        }
        html += `<td style="text-align:right;color:#999">${rate.toFixed(2)}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td><td style="text-align:right">${totalDemand.toFixed(0)}</td><td style="text-align:right">${totalCap.toFixed(0)}</td>`;
    if (simDay) {
        html += `<td style="text-align:right;color:#93c5fd">${totalChp.toFixed(0)}</td>`;
        html += `<td style="text-align:right;color:${totalPlanned > 0 ? '#FFB74D' : '#6BCB77'}">${totalPlanned.toFixed(0)}</td>`;
    }
    html += '<td></td></tr></tbody></table>';
    el.innerHTML = html;
}

// ─── 비용산출: 가스비 렌더링 ───
function renderCostGas(simResults) {
    const el = document.getElementById('pv_cost_gas');
    if (!el) return;
    if (!simResults || simResults.length === 0) {
        el.innerHTML = '<p class="hint">기동계획 실행 후 표시됩니다</p>'; return;
    }
    const monthly = aggregateMonthly(simResults);
    const fmt = v => Math.round(v / 1000).toLocaleString();
    let html = '<div style="font-size:11px;color:#94a3b8;margin-bottom:8px">시뮬레이션 결과 기반 (기동계획 연동)</div>';
    html += '<table class="pv-tbl" style="max-width:700px"><thead><tr>';
    html += '<th>월</th><th>CHP가스비(천원)</th><th>PLB가스비(천원)</th><th>기동비용(천원)</th><th>합계(천원)</th>';
    html += '</tr></thead><tbody>';
    let yChp = 0, yPlb = 0, yStart = 0, yTotal = 0;
    for (let i = 0; i < 12; i++) {
        const b = monthly[i];
        const chp = b.chpFuelCost || 0, plb = b.plbFuelCost || 0, start = b.startupCost || 0;
        const total = chp + plb + start;
        yChp += chp; yPlb += plb; yStart += start; yTotal += total;
        html += `<tr><td>${b.month}월</td><td style="text-align:right">${fmt(chp)}</td><td style="text-align:right">${fmt(plb)}</td>`;
        html += `<td style="text-align:right">${fmt(start)}</td><td style="text-align:right;font-weight:600">${fmt(total)}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td><td style="text-align:right">${fmt(yChp)}</td><td style="text-align:right">${fmt(yPlb)}</td>`;
    html += `<td style="text-align:right">${fmt(yStart)}</td><td style="text-align:right">${fmt(yTotal)}</td></tr>`;
    html += '</tbody></table>';
    el.innerHTML = html;
}

// ─── 비용산출: 용수약품비 렌더링 ───
function renderCostWaterChem() {
    const el = document.getElementById('pv_cost_water_chem');
    if (!el) return;
    const wc = calcWaterChemMonthly();
    if (!wc) { el.innerHTML = '<p class="hint">용수약품 사용량 + 단가 데이터를 입력하세요</p>'; return; }
    const fmt = v => Math.round(v / 1000).toLocaleString();
    let html = '<div style="font-size:11px;color:#f59e0b;margin-bottom:8px">입력 데이터 기반 (월별 계획값)</div>';
    html += '<table class="pv-tbl" style="max-width:500px"><thead><tr>';
    html += '<th>월</th><th>용수비(천원)</th><th>약품비(천원)</th><th>합계(천원)</th>';
    html += '</tr></thead><tbody>';
    let yW = 0, yC = 0, yT = 0;
    for (let m = 1; m <= 12; m++) {
        const w = wc[m]?.waterCost || 0, c = wc[m]?.chemTotal || 0, t = w + c;
        yW += w; yC += c; yT += t;
        html += `<tr><td>${m}월</td><td style="text-align:right">${fmt(w)}</td><td style="text-align:right">${fmt(c)}</td>`;
        html += `<td style="text-align:right;font-weight:600">${fmt(t)}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td><td style="text-align:right">${fmt(yW)}</td>`;
    html += `<td style="text-align:right">${fmt(yC)}</td><td style="text-align:right">${fmt(yT)}</td></tr>`;
    html += '</tbody></table>';
    el.innerHTML = html;
}

// ─── 비용산출: 수열비 렌더링 ───
function renderCostExtHeat() {
    const el = document.getElementById('pv_cost_ext_heat');
    if (!el) return;
    const ext = calcExternalHeatMonthly();
    if (!ext) { el.innerHTML = '<p class="hint">외부거래 + 판매단가 데이터를 입력하세요</p>'; return; }
    const fmt = v => Math.round(v / 1000).toLocaleString();
    let html = '<div style="font-size:11px;color:#f59e0b;margin-bottom:8px">입력 데이터 기반 (월별 계획값)</div>';
    html += '<table class="pv-tbl" style="max-width:400px"><thead><tr>';
    html += '<th>월</th><th>수열비(천원)</th>';
    html += '</tr></thead><tbody>';
    let yTotal = 0;
    for (let m = 1; m <= 12; m++) {
        const v = ext[m] || 0; yTotal += v;
        html += `<tr><td>${m}월</td><td style="text-align:right;font-weight:600">${fmt(v)}</td></tr>`;
    }
    html += `<tr style="font-weight:700;border-top:2px solid #555"><td>합계</td><td style="text-align:right">${fmt(yTotal)}</td></tr>`;
    html += '</tbody></table>';
    el.innerHTML = html;
}

// ─── 비용산출 전체 렌더링 ───
function renderAllCosts(simResults) {
    renderElecDemandProfile(simResults);
    renderElecCostComparison(simResults);
    renderCostGas(simResults);
    renderCostWaterChem();
    renderCostExtHeat();
}

// ─── 열판매매출 월별 집계 (공헌이익용) ───
function calcHeatSalesMonthly() {
    if (!DATA.sales_hourly && !DATA.sales_daily) return null;

    const findCol = (headers, name) => headers.findIndex(h => h === name);
    const heatItemDefs = [
        { name: '주택용', key: '주택용난방', rateKey: '주택용' },
        { name: '업무용난방', key: '업무용난방', rateKey: '업무용' },
        { name: '업무용냉방', key: '업무용냉방', rateKey: '업무용' },
        { name: '공공용난방', key: '공공용난방', rateKey: '공공용' },
        { name: '공공용냉방', key: '공공용냉방', rateKey: '공공용' },
    ];

    // sales_hourly → 일별 항목별 합산
    const dailyHeatQty = {};
    if (DATA.sales_hourly) {
        const shH = DATA.sales_hourly.headers;
        const colMap = {};
        heatItemDefs.forEach(def => { colMap[def.key] = findCol(shH, def.key); });
        DATA.sales_hourly.rows.forEach(row => {
            const key = `${Number(row[0])}-${Number(row[1])}`;
            if (!dailyHeatQty[key]) dailyHeatQty[key] = {};
            heatItemDefs.forEach(def => {
                const ci = colMap[def.key];
                if (ci >= 0) dailyHeatQty[key][def.key] = (dailyHeatQty[key][def.key] || 0) + (Number(row[ci]) || 0);
            });
        });
    }

    // sales_daily → 인테코
    const intecoQtyMap = {};
    if (DATA.sales_daily) {
        const sdH = DATA.sales_daily.headers;
        const intecoCol = findCol(sdH, '인테코');
        if (intecoCol >= 0) DATA.sales_daily.rows.forEach(row => { intecoQtyMap[`${Number(row[0])}-${Number(row[1])}`] = Number(row[intecoCol]) || 0; });
    }

    const heatRateMap = {};
    if (DATA.heat_rate_table) {
        const hrH = DATA.heat_rate_table.headers;
        DATA.heat_rate_table.rows.forEach(row => {
            const m = Number(row[0]);
            const entry = {};
            hrH.forEach((h, i) => { if (i > 0) entry[h] = parseNum(row[i]); });
            heatRateMap[m] = entry;
        });
    }

    const intecoPriceMap = {};
    if (DATA.sale_prices_daily) {
        const prH = DATA.sale_prices_daily.headers;
        const ipCol = prH.indexOf('인테코(판매)');
        if (ipCol >= 0) DATA.sale_prices_daily.rows.forEach(row => { intecoPriceMap[`${Number(row[0])}-${Number(row[1])}`] = parseNum(row[ipCol]); });
    }

    const baseFeeMap = {};
    if (DATA.heat_rate_table) {
        const bfIdx = DATA.heat_rate_table.headers.indexOf('기본료');
        if (bfIdx >= 0) DATA.heat_rate_table.rows.forEach(row => { baseFeeMap[Number(row[0])] = parseNum(row[bfIdx]); });
    }

    const result = {};
    for (let m = 1; m <= 12; m++) result[m] = { usageRev: 0, baseFee: baseFeeMap[m] || 0, total: 0 };

    // 일별 키 수집
    const allDayKeys = new Set([...Object.keys(dailyHeatQty), ...Object.keys(intecoQtyMap)]);
    allDayKeys.forEach(key => {
        const m = Number(key.split('-')[0]);
        if (m < 1 || m > 12) return;
        const rates = heatRateMap[m] || {};
        const hq = dailyHeatQty[key] || {};
        heatItemDefs.forEach(def => {
            const qty = hq[def.key] || 0;
            const price = rates[def.rateKey] || 0;
            result[m].usageRev += Math.round(qty * price);
        });
        const intecoQty = intecoQtyMap[key] || 0;
        const intecoPrice = intecoPriceMap[key] || 0;
        result[m].usageRev += Math.round(intecoQty * intecoPrice);
    });
    for (let m = 1; m <= 12; m++) result[m].total = result[m].usageRev + result[m].baseFee;
    return result;
}

// ─── 일별 상세 테이블 렌더링 ───
function renderPlanDaily(simResults) {
    const el = document.getElementById('pv_plan_daily');
    if (!el) return;

    // 월 선택 버튼 (기본 1월)
    let html = '<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;flex-wrap:wrap">';
    html += '<span style="font-size:13px;font-weight:700;color:#334155;margin-right:4px">월 선택</span>';
    for (let m = 1; m <= 12; m++) {
        const active = m === 1 ? 'background:#3b82f6;color:#fff;border-color:#3b82f6' : '';
        html += `<button class="daily-month-btn" data-month="${m}" style="padding:4px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;background:#fff;color:#475569;transition:all .15s;${active}">${m}월</button>`;
    }
    html += '</div>';

    // 월별 차트
    html += '<div id="dailyChartWrap" class="chart-row" style="margin-bottom:16px">';
    html += '<div class="chart-wrap"><div class="chart-title">일별 열생산 (Gcal)</div><canvas id="chartDailyHeat"></canvas></div>';
    html += '</div>';

    // 테이블
    html += '<div id="dailyTableBody" style="overflow:auto;max-height:calc(100vh - 380px)">';
    html += buildDailyTable(simResults, 1);
    html += '</div>';

    el.innerHTML = html;

    // 1월 차트
    renderDailyChart(simResults, 1);

    // 월 버튼 이벤트
    el.querySelectorAll('.daily-month-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const m = parseInt(btn.dataset.month);
            el.querySelectorAll('.daily-month-btn').forEach(b => {
                b.style.background = '#fff'; b.style.color = '#475569'; b.style.borderColor = '#cbd5e1';
            });
            btn.style.background = '#3b82f6'; btn.style.color = '#fff'; btn.style.borderColor = '#3b82f6';
            document.getElementById('dailyTableBody').innerHTML = buildDailyTable(simResults, m);
            renderDailyChart(simResults, m);
        });
    });

    // 행 클릭 → 모달 상세 보기
    document.getElementById('dailyTableBody')?.addEventListener('click', (e) => {
        const tr = e.target.closest('tr[data-day-idx]');
        if (!tr) return;
        const idx = parseInt(tr.dataset.dayIdx);
        if (!isNaN(idx)) showDayDetailModal(simResults, idx);
    });
}

function buildDailyTable(simResults, filterMonth) {
    const fmt = v => fmtNum(Math.round(v));
    // 원본 인덱스 유지
    const rows = simResults.map((r, i) => ({ ...r, _idx: i }));
    const filtered = filterMonth > 0 ? rows.filter(r => r.month === filterMonth) : rows;

    let html = '<table class="pv-tbl daily-detail"><thead>';
    // 그룹 헤더
    html += '<tr class="group-hdr">';
    html += '<th colspan="4"></th>';
    html += '<th colspan="3" style="background:#e8f5e9;color:#2e7d32">열수급</th>';
    html += '<th style="background:#fee2e2;color:#c62828"></th>';
    html += '<th colspan="3" style="background:#dbeafe;color:#1565c0">CHP</th>';
    html += '<th colspan="3" style="background:#fef3c7;color:#e65100">PLB</th>';
    html += '<th colspan="2" style="background:#e8eaf6;color:#283593">축열조</th>';
    html += '<th colspan="7" style="background:#fce4ec;color:#880e4f">비용</th>';
    html += '</tr>';
    // 컬럼 헤더
    html += '<tr>';
    html += '<th>일</th><th>요일</th><th>유형</th><th>SMP</th>';
    html += '<th style="background:#e8f5e9">수요</th><th style="background:#e8f5e9">외부수열</th><th style="background:#e8f5e9">인테코</th>';
    html += '<th style="background:#fee2e2">순필요열</th>';
    html += '<th style="background:#dbeafe">시간</th><th style="background:#dbeafe">부하</th><th style="background:#dbeafe">생산열</th>';
    html += '<th style="background:#fef3c7">대수</th><th style="background:#fef3c7">시간</th><th style="background:#fef3c7">생산열</th>';
    html += '<th style="background:#e8eaf6">축열조</th>';
    html += '<th style="background:#e8eaf6">축방열</th>';
    html += '<th style="background:#fce4ec">CHP연료비</th>';
    html += '<th style="background:#fce4ec">기동비</th>';
    html += '<th style="background:#e8f5e9">열제약매출</th>';
    html += '<th style="background:#fce4ec">CHP단가<br>원/Gcal</th>';
    html += '<th style="background:#fef3c7">PLB연료비</th>';
    html += '<th style="background:#fef3c7">PLB단가<br>원/Gcal</th>';
    html += '<th style="background:#fce4ec">순비용</th>';
    html += '</tr></thead><tbody>';

    const fmtK = v => fmtNum(Math.round(v / 1000)); // 천원 단위
    const tot = { demand: 0, ext: 0, inteco: 0, net: 0, chpH: 0, chpHeat: 0, plbHeat: 0, stDelta: 0,
                  chpFuelCost: 0, plbFuelCost: 0, startupCost: 0, elecRev: 0, netCost: 0, chpNetCost: 0 };

    filtered.forEach((r, fi) => {
        const typeColor = r.isMaint ? '#ef4444' : (r.isHoliday ? '#8b5cf6' : (r.isWeekend ? '#f59e0b' : '#334155'));
        const typeBg = r.isMaint ? 'background:#fef2f2' : (r.isHoliday ? 'background:#faf5ff' : (r.isWeekend ? 'background:#fffbeb' : ''));
        const plbStyle = r.plbHeat > 0 ? 'color:#dc2626;font-weight:700' : 'color:#94a3b8';
        const stColor = r.storageLevel < 300 ? 'color:#dc2626;font-weight:700' : (r.storageLevel > 1200 ? 'color:#059669' : '');

        // 축방열량: 오늘 잔량 - 전일 잔량 (양수=축열, 음수=방열)
        const prevSt = r._idx > 0 ? simResults[r._idx - 1].storageLevel : (parseFloat(document.getElementById('storageInitial')?.value) || 750);
        const stDelta = r.storageLevel - prevSt;

        tot.demand += r.demand; tot.ext += r.adjExternal; tot.inteco += r.adjIntecoNet;
        tot.net += r.netRequired; tot.chpH += r.chpHours; tot.chpHeat += r.chpHeat; tot.plbHeat += r.plbHeat;
        tot.stDelta += stDelta;
        tot.chpFuelCost += r.chpFuelCost || 0; tot.plbFuelCost += r.plbFuelCost || 0; tot.startupCost += r.startupCost || 0;
        tot.elecRev += r.chpElecRev || 0; tot.netCost += r.totalNetCost || 0; tot.chpNetCost += r.chpNetCost || 0;

        html += `<tr data-day-idx="${r._idx}" style="${typeBg};cursor:pointer" title="클릭하여 상세 보기">`;
        html += `<td style="text-align:center;color:#64748b">${r.day}</td>`;
        html += `<td style="text-align:center">${r.dowName}</td>`;
        html += `<td style="text-align:center;color:${typeColor};font-weight:600">${r.dayType}</td>`;
        html += `<td style="color:#64748b">${r.smpAvg > 0 ? Math.round(r.smpAvg * 10) / 10 : '-'}</td>`;
        // 열수급
        html += `<td style="background:#f1f8e9">${fmt(r.demand)}</td>`;
        html += `<td style="background:#f1f8e9">${fmt(r.adjExternal)}</td>`;
        html += `<td style="background:#f1f8e9">${r.adjIntecoNet !== 0 ? fmt(r.adjIntecoNet) : '-'}</td>`;
        const netColor = r.netRequired < 0 ? 'color:#2e7d32' : 'color:#c62828';
        html += `<td style="background:#fff5f5;font-weight:700;${netColor}">${fmt(r.netRequired)}</td>`;
        // CHP
        html += `<td style="background:#eff6ff">${r.chpHours || '-'}</td>`;
        const loadClr = (r.chpLoad || 100) < 100 ? 'color:#e65100;font-weight:700' : 'color:#64748b';
        html += `<td style="background:#eff6ff;${loadClr}">${r.chpHours > 0 ? (r.chpLoad || 100) + '%' : '-'}</td>`;
        html += `<td style="background:#eff6ff;font-weight:600">${r.chpHeat > 0 ? fmt(r.chpHeat) : '-'}</td>`;
        // PLB
        html += `<td style="background:#fffbeb;${plbStyle}">${r.plbUnits || '-'}</td>`;
        html += `<td style="background:#fffbeb;${plbStyle}">${r.plbHours > 0 ? r.plbHours : '-'}</td>`;
        html += `<td style="background:#fffbeb;${plbStyle}">${r.plbHeat > 0 ? fmt(r.plbHeat) : '-'}</td>`;
        // 축열조 잔량
        html += `<td style="background:#ede7f6;${stColor}">${fmt(r.storageLevel)}</td>`;
        // 축방열
        const deltaColor = stDelta > 0 ? 'color:#1565c0' : (stDelta < 0 ? 'color:#c62828' : 'color:#94a3b8');
        const deltaLabel = stDelta > 0 ? '+' : '';
        html += `<td style="background:#ede7f6;font-weight:600;${deltaColor}">${stDelta !== 0 ? deltaLabel + fmt(stDelta) : '-'}</td>`;
        // 비용 (천원): CHP연료비 → 기동비 → 열제약매출 → CHP단가 → PLB연료비 → PLB단가 → 순비용
        html += `<td style="background:#fef0f5">${r.chpFuelCost ? fmtK(r.chpFuelCost) : '-'}</td>`;
        html += `<td style="background:#fef0f5">${(r.startupCost || 0) > 0 ? fmtK(r.startupCost) : '-'}</td>`;
        html += `<td style="background:#f0fdf4;color:#15803d">${r.chpElecRev ? fmtK(r.chpElecRev) : '-'}</td>`;
        const chpUc = r.chpHeat > 0 ? Math.round((r.chpFuelCost + (r.startupCost || 0) - (r.chpElecRev || 0)) / r.chpHeat) : 0;
        const ucClr = chpUc < 0 ? 'color:#1565c0' : 'color:#c62828';
        html += `<td style="background:#fef0f5;font-weight:600;${ucClr}">${r.chpHeat > 0 ? fmtNum(chpUc) : '-'}</td>`;
        html += `<td style="background:#fffbeb">${r.plbFuelCost ? fmtK(r.plbFuelCost) : '-'}</td>`;
        const plbUc = r.plbHeat > 0 ? Math.round(r.plbFuelCost / r.plbHeat) : 0;
        html += `<td style="background:#fffbeb;font-weight:600;color:#e65100">${r.plbHeat > 0 ? fmtNum(plbUc) : '-'}</td>`;
        const ncColor = (r.totalNetCost || 0) < 0 ? 'color:#1565c0' : 'color:#c62828';
        html += `<td style="background:#fef0f5;font-weight:700;${ncColor}">${r.totalNetCost != null ? fmtK(r.totalNetCost) : '-'}</td>`;
        html += '</tr>';
    });

    // 합계 행
    html += '<tr style="border-top:2px solid #334155;font-weight:800;background:#f8fafc">';
    html += `<td colspan="4" style="text-align:center">합계</td>`;
    html += `<td style="background:#e8f5e9">${fmt(tot.demand)}</td>`;
    html += `<td style="background:#e8f5e9">${fmt(tot.ext)}</td>`;
    html += `<td style="background:#e8f5e9">${tot.inteco !== 0 ? fmt(tot.inteco) : '-'}</td>`;
    html += `<td style="background:#fee2e2">${fmt(tot.net)}</td>`;
    html += `<td style="background:#dbeafe">${fmt(tot.chpH)}</td>`;
    html += `<td style="background:#dbeafe"></td>`;
    html += `<td style="background:#dbeafe">${fmt(tot.chpHeat)}</td>`;
    html += `<td colspan="2" style="background:#fef3c7"></td>`;
    html += `<td style="background:#fef3c7">${fmt(tot.plbHeat)}</td>`;
    html += `<td style="background:#e8eaf6"></td>`;
    const totDeltaColor = tot.stDelta > 0 ? 'color:#1565c0' : (tot.stDelta < 0 ? 'color:#c62828' : '');
    html += `<td style="background:#e8eaf6;${totDeltaColor}">${(tot.stDelta > 0 ? '+' : '') + fmt(tot.stDelta)}</td>`;
    html += `<td style="background:#fce4ec">${fmtK(tot.chpFuelCost)}</td>`;
    html += `<td style="background:#fce4ec">${tot.startupCost > 0 ? fmtK(tot.startupCost) : '-'}</td>`;
    html += `<td style="background:#dcfce7;color:#15803d">${fmtK(tot.elecRev)}</td>`;
    const totDailyUc = tot.chpHeat > 0 ? Math.round((tot.chpFuelCost + tot.startupCost - tot.elecRev) / tot.chpHeat) : 0;
    const totDailyUcClr = totDailyUc < 0 ? 'color:#1565c0' : 'color:#c62828';
    html += `<td style="background:#fce4ec;font-weight:700;${totDailyUcClr}">${tot.chpHeat > 0 ? fmtNum(totDailyUc) : '-'}</td>`;
    html += `<td style="background:#fef3c7">${fmtK(tot.plbFuelCost)}</td>`;
    const totDailyPlbUc = tot.plbHeat > 0 ? Math.round(tot.plbFuelCost / tot.plbHeat) : 0;
    html += `<td style="background:#fef3c7;font-weight:700;color:#e65100">${tot.plbHeat > 0 ? fmtNum(totDailyPlbUc) : '-'}</td>`;
    const totNcColor = tot.netCost < 0 ? 'color:#1565c0' : 'color:#c62828';
    html += `<td style="background:#fce4ec;${totNcColor}">${fmtK(tot.netCost)}</td>`;
    html += '</tr>';

    html += '</tbody></table>';
    return html;
}

// ─── 일 상세 모달 (하루 24시간 차트) ───
let _dayDetailChart = null;

// ── 일별 세부 모달: 데이터 계산 ──
// 차트 데이터: 시뮬레이션 결과의 hourly를 직접 사용
function _calcDayDetailData(simResults, dayIndex) {
    const r = simResults[dayIndex];
    const smpH = r.smpH || Array(24).fill(r.smpAvg || 0);
    const hours = r.hourly || [];
    return { r, smpH, hours };
}

// ── 일별 세부 모달: 내부 콘텐츠 업데이트 (DOM 재사용) ──
let _dayDetailCurrentIndex = 0;
let _dayDetailSimResults = null;

function _updateDayDetailContent(dayIndex) {
    _dayDetailCurrentIndex = dayIndex;
    const simResults = _dayDetailSimResults;
    const { r, smpH, hours } = _calcDayDetailData(simResults, dayIndex);
    const fmt = v => fmtNum(Math.round(v));

    // 헤더 업데이트
    const titleEl = document.getElementById('ddm_title');
    const badgeEl = document.getElementById('ddm_badge');
    const prevBtn = document.getElementById('ddm_prevBtn');
    const nextBtn = document.getElementById('ddm_nextBtn');
    if (titleEl) titleEl.textContent = `${r.month}월 ${r.day}일 (${r.dowName})`;
    if (badgeEl) badgeEl.textContent = r.dayType;
    if (prevBtn) prevBtn.style.visibility = dayIndex > 0 ? 'visible' : 'hidden';
    if (nextBtn) nextBtn.style.visibility = dayIndex < simResults.length - 1 ? 'visible' : 'hidden';

    // 요약 칩 업데이트
    const chipsEl = document.getElementById('ddm_chips');
    if (chipsEl) {
        let chips = '';
        chips += `<div class="day-chip" style="--t:#ef4444">수요 ${fmt(r.demand)}</div>`;
        chips += `<div class="day-chip" style="--t:#6BCB77">외부수열 ${fmt(r.adjExternal)}</div>`;
        if (r.adjIntecoNet !== 0) chips += `<div class="day-chip" style="--t:#6BCB77">인테코 ${fmt(r.adjIntecoNet)}</div>`;
        const loadLabel = (r.chpLoad && r.chpLoad < 100) ? ` (${r.chpLoad}%)` : '';
        chips += `<div class="day-chip" style="--t:#3b82f6">CHP ${r.chpHours}h${loadLabel} / ${fmt(r.chpHeat)}</div>`;
        chips += `<div class="day-chip" style="--t:#f59e0b">PLB ${r.plbHeat > 0 ? r.plbUnits + '\uB300 ' + r.plbHours + 'h / ' + fmt(r.plbHeat) : '-'}</div>`;
        const stColor = r.storageLevel < 300 ? '#FF6B6B' : '#a78bfa';
        chips += `<div class="day-chip" style="--t:${stColor}">축열조 ${fmt(r.storageLevel)}</div>`;
        chips += `<div class="day-chip" style="--t:#E1BEE7">SMP ${r.smpAvg > 0 ? Math.round(r.smpAvg * 10) / 10 + ' \uC6D0' : '-'}</div>`;
        chipsEl.innerHTML = chips;
    }

    // 차트 데이터 업데이트 (destroy 없이)
    if (_dayDetailChart) {
        const demandData = hours.map(h => Math.round((h.demand || 0) * 10) / 10);
        const extData = hours.map(h => Math.round((h.ext || 0) * 10) / 10);
        const chpData = hours.map(h => Math.round(h.chp * 10) / 10);
        const plbData = hours.map(h => Math.round(h.plb * 10) / 10);
        _dayDetailChart.data.datasets[0].data = hours.map(h => h.storage);
        _dayDetailChart.data.datasets[1].data = demandData;
        _dayDetailChart.data.datasets[2].data = extData;
        _dayDetailChart.data.datasets[3].data = chpData;
        _dayDetailChart.data.datasets[4].data = plbData;
        // Y축 max 동적 조정 (정수, 천단위→100, 백단위→50 끊기)
        const maxHeat = Math.max(...demandData, ...chpData, ...plbData, 50);
        const yMax = niceMax(maxHeat);
        _dayDetailChart.options.scales.y.max = yMax;
        _dayDetailChart.options.scales.y.ticks.stepSize = niceStep(yMax);
        // tooltip용 smpH 갱신
        _dayDetailChart._smpH = smpH;
        _dayDetailChart.update('none'); // 애니메이션 없이 즉시 업데이트
    }

    // 축방열 + SMP 배지 업데이트 (차트 x축 정렬)
    if (_dayDetailChart) {
        const ca = _dayDetailChart.chartArea;
        if (ca) {
            const ml = ca.left + 'px';
            const w = (ca.right - ca.left) + 'px';

            // 축방열량
            const stValsEl = document.getElementById('ddm_storageVals');
            if (stValsEl) {
                stValsEl.style.marginLeft = ml;
                stValsEl.style.width = w;
                let stHtml = '';
                for (let h = 0; h < 24; h++) {
                    const prev = h === 0 ? (r.storageLevel || hours[0]?.storage || 0) : hours[h - 1].storage;
                    const cur = hours[h]?.storage || 0;
                    const diff = cur - prev;
                    const isCharge = diff > 0;
                    const bg = diff === 0 ? '#1a1a1a' : (isCharge ? 'rgba(59,130,246,0.15)' : 'rgba(239,68,68,0.15)');
                    const tc = diff === 0 ? '#555' : (isCharge ? '#60a5fa' : '#f87171');
                    const label = diff === 0 ? '·' : (isCharge ? '+' : '') + Math.round(diff);
                    stHtml += `<div style="flex:1;text-align:center;padding:3px 0;border-radius:3px;background:${bg};font-size:10px;font-weight:600;color:${tc}">${label}</div>`;
                }
                stValsEl.innerHTML = stHtml;
            }

            // SMP
            const smpValsEl = document.getElementById('ddm_smpVals');
            if (smpValsEl) {
                smpValsEl.style.marginLeft = ml;
                smpValsEl.style.width = w;
                let smpHtml = '';
                for (let h = 0; h < 24; h++) {
                    const v = smpH[h];
                    const bg = v >= 150 ? '#fee2e2' : (v >= 100 ? '#fef3c7' : '#e8f5e9');
                    const tc = v >= 150 ? '#c62828' : (v >= 100 ? '#92400e' : '#2e7d32');
                    smpHtml += `<div style="flex:1;text-align:center;padding:3px 0;border-radius:3px;background:${bg};font-size:10px;font-weight:600;color:${tc}">${v > 0 ? Math.round(v) : '-'}</div>`;
                }
                smpValsEl.innerHTML = smpHtml;
            }
        }
    }
}

function showDayDetailModal(simResults, dayIndex) {
    _dayDetailSimResults = simResults;
    _dayDetailCurrentIndex = dayIndex;

    // 모달이 이미 있으면 내용만 업데이트 (깜빡임 방지)
    if (document.getElementById('dayDetailModal')) {
        _updateDayDetailContent(dayIndex);
        return;
    }

    // ── 최초 모달 생성 (한 번만) ──
    if (_dayDetailChart) { _dayDetailChart.destroy(); _dayDetailChart = null; }

    const overlay = document.createElement('div');
    overlay.id = 'dayDetailModal';
    overlay.className = 'day-modal-overlay';
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeDayDetail(); });

    let html = '<div class="day-modal-box" style="max-width:1100px;width:95vw">';

    // 헤더: ◀ 날짜 ▶ + 닫기 (고정 구조, 내용은 id로 업데이트)
    html += '<div class="day-modal-header">';
    html += '<div style="display:flex;align-items:center;gap:12px">';
    html += '<button id="ddm_prevBtn" class="day-nav-btn" title="이전일 (←)">\u25C0</button>';
    html += '<span id="ddm_title" class="day-modal-title"></span>';
    html += '<span id="ddm_badge" class="day-modal-badge"></span>';
    html += '<button id="ddm_nextBtn" class="day-nav-btn" title="다음일 (→)">\u25B6</button>';
    html += '</div>';
    html += '<button class="day-close-btn" title="닫기 (ESC)">\u2715</button>';
    html += '</div>';

    // 요약 칩 (컨테이너만, 내용은 업데이트)
    html += '<div id="ddm_chips" style="display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap"></div>';

    // 차트
    html += '<div style="height:480px;margin-bottom:10px"><canvas id="chartDayDetail"></canvas></div>';

    // 축방열량 배지 행
    html += '<div id="ddm_storageWrap" style="position:relative;margin-top:-2px;display:flex;align-items:center">';
    html += '<div style="font-size:10px;font-weight:700;color:#64748b;position:absolute;left:0">축방열</div>';
    html += '<div id="ddm_storageVals" style="display:flex;position:relative"></div>';
    html += '</div>';

    // SMP 배지 행 (차트 x축에 정렬)
    html += '<div id="ddm_smpWrap" style="position:relative;margin-top:2px;display:flex;align-items:center">';
    html += '<div id="ddm_smpLabel" style="font-size:10px;font-weight:700;color:#64748b;position:absolute;left:0">SMP</div>';
    html += '<div id="ddm_smpVals" style="display:flex;position:relative"></div>';
    html += '</div>';

    html += '</div>'; // modal-box
    overlay.innerHTML = html;
    document.body.appendChild(overlay);

    // 네비게이션 이벤트 (한 번만 등록)
    document.getElementById('ddm_prevBtn').addEventListener('click', () => {
        if (_dayDetailCurrentIndex > 0) _updateDayDetailContent(_dayDetailCurrentIndex - 1);
    });
    document.getElementById('ddm_nextBtn').addEventListener('click', () => {
        if (_dayDetailCurrentIndex < _dayDetailSimResults.length - 1) _updateDayDetailContent(_dayDetailCurrentIndex + 1);
    });
    overlay.querySelector('.day-close-btn').addEventListener('click', closeDayDetail);

    // 키보드 (전역, 모달 열려있는 동안)
    document.addEventListener('keydown', _dayDetailKeyHandler);

    // ── Chart.js 생성 (한 번만) ──
    const { smpH, hours } = _calcDayDetailData(simResults, dayIndex);
    const ctx = document.getElementById('chartDayDetail')?.getContext('2d');
    if (!ctx) return;

    const smpAlignPlugin = {
        id: 'smpAlign',
        afterDraw(chart) {
            const ca = chart.chartArea;
            if (!ca) return;
            const ml = ca.left + 'px', w = (ca.right - ca.left) + 'px';
            const smp = document.getElementById('ddm_smpVals');
            const st = document.getElementById('ddm_storageVals');
            if (smp) { smp.style.marginLeft = ml; smp.style.width = w; }
            if (st) { st.style.marginLeft = ml; st.style.width = w; }
        }
    };

    _dayDetailChart = new Chart(ctx, {
        type: 'bar',
        plugins: [smpAlignPlugin],
        data: {
            labels: Array.from({ length: 24 }, (_, i) => `${i}시`),
            datasets: [
                {
                    label: '축열조', type: 'bar',
                    data: hours.map(h => h.storage),
                    backgroundColor: 'rgba(167,139,250,0.15)',
                    borderColor: '#a78bfa', borderWidth: 1,
                    yAxisID: 'y1', order: 6, barPercentage: 0.55,
                },
                {
                    label: '수요', type: 'line',
                    data: hours.map(h => Math.round((h.demand || 0) * 10) / 10),
                    borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,0.08)',
                    borderWidth: 2, pointRadius: 2, pointBackgroundColor: '#ef4444',
                    fill: false, yAxisID: 'y', order: 1,
                },
                {
                    label: '외부수열', type: 'line',
                    data: hours.map(h => Math.round((h.ext || 0) * 10) / 10),
                    borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,0.08)',
                    borderWidth: 2, borderDash: [6, 3], pointRadius: 1, pointBackgroundColor: '#22c55e',
                    fill: false, yAxisID: 'y', order: 4,
                },
                {
                    label: 'CHP', type: 'line',
                    data: hours.map(h => Math.round(h.chp * 10) / 10),
                    borderColor: '#3b82f6', backgroundColor: 'rgba(59,130,246,0.08)',
                    borderWidth: 2, pointRadius: 2, pointBackgroundColor: '#42a5f5',
                    fill: false, yAxisID: 'y', order: 2, stepped: 'before',
                },
                {
                    label: 'PLB', type: 'line',
                    data: hours.map(h => Math.round(h.plb * 10) / 10),
                    borderColor: '#f59e0b', backgroundColor: 'rgba(245,158,11,0.08)',
                    borderWidth: 2, pointRadius: 2, pointBackgroundColor: '#f59e0b',
                    fill: false, yAxisID: 'y', order: 3, stepped: 'before',
                },
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            animation: false, // 네비게이션 시 빠른 응답
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { position: 'top', labels: { color: '#aaa', font: { size: 10, family: "'Pretendard',sans-serif" }, boxWidth: 16, boxHeight: 5, padding: 14, useBorderRadius: true, borderRadius: 3 } },
                tooltip: {
                    callbacks: {
                        afterBody: (items) => {
                            const idx = items[0]?.dataIndex;
                            if (idx == null) return '';
                            const sv = _dayDetailChart._smpH?.[idx] || 0;
                            return `SMP: ${sv > 0 ? Math.round(sv) + ' \uC6D0/kWh' : '-'}`;
                        }
                    }
                }
            },
            scales: (() => {
                const storageCap = parseFloat(document.getElementById('storageCapacity')?.value) || 1500;
                const maxH = Math.max(...hours.map(h => h.demand || 0), ...hours.map(h => h.chp || 0), 50);
                const yMax = niceMax(maxH);
                const sCap = niceMax(storageCap);
                return {
                    y: { beginAtZero: true, max: yMax, position: 'left', title: { display: true, text: 'Gcal/h', font: { size: 10 }, color: '#94a3b8' }, ticks: { font: { size: 9 }, color: '#94a3b8', stepSize: niceStep(yMax) }, grid: { color: 'rgba(255,255,255,0.04)' } },
                    y1: { beginAtZero: true, max: sCap, position: 'right', title: { display: true, text: '축열조 (Gcal)', font: { size: 10 }, color: '#818cf8' }, grid: { drawOnChartArea: false }, ticks: { font: { size: 9 }, color: '#818cf8', stepSize: niceStep(sCap) } },
                    x: { ticks: { font: { size: 9 }, maxRotation: 0, color: '#94a3b8' } }
                };
            })()
        }
    });
    _dayDetailChart._smpH = smpH;

    // 초기 내용 채우기 (차트 레이아웃 완료 후 2프레임 대기)
    requestAnimationFrame(() => requestAnimationFrame(() => _updateDayDetailContent(dayIndex)));
}

function _dayDetailKeyHandler(e) {
    if (!document.getElementById('dayDetailModal')) {
        document.removeEventListener('keydown', _dayDetailKeyHandler);
        return;
    }
    if (e.key === 'Escape') closeDayDetail();
    else if (e.key === 'ArrowLeft' && _dayDetailCurrentIndex > 0) _updateDayDetailContent(_dayDetailCurrentIndex - 1);
    else if (e.key === 'ArrowRight' && _dayDetailCurrentIndex < _dayDetailSimResults.length - 1) _updateDayDetailContent(_dayDetailCurrentIndex + 1);
}

function closeDayDetail() {
    document.getElementById('dayDetailModal')?.remove();
    if (_dayDetailChart) { _dayDetailChart.destroy(); _dayDetailChart = null; }
    document.removeEventListener('keydown', _dayDetailKeyHandler);
    _dayDetailSimResults = null;
}

// ─── 월별 차트 렌더링 ───
let _monthlyChart = null;
let _dailyChart = null;

function renderPlanCharts(simResults) {
    const monthly = aggregateMonthly(simResults);
    const labels = monthly.map(b => `${b.month}월`);

    // 기존 차트 파괴
    if (_monthlyChart) { _monthlyChart.destroy(); _monthlyChart = null; }

    const ctx1 = document.getElementById('chartMonthlyHeat')?.getContext('2d');
    if (ctx1) {
        _monthlyChart = new Chart(ctx1, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: 'CHP 생산열', data: monthly.map(b => Math.round(b.chpHeat)), backgroundColor: '#3b82f6', stack: 'prod' },
                    { label: 'PLB 생산열', data: monthly.map(b => Math.round(b.plbHeat)), backgroundColor: '#f59e0b', stack: 'prod' },
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: {
                    legend: { position: 'top', labels: { color: '#aaa', font: { size: 10, family: "'Pretendard',sans-serif" }, boxWidth: 16, boxHeight: 5, padding: 14, useBorderRadius: true, borderRadius: 3 } },
                    tooltip: {
                        callbacks: {
                            label: item => {
                                const v = item.raw;
                                if (v == null) return '';
                                return `${item.dataset.label}: ${v.toLocaleString()}`;
                            }
                        }
                    }
                },
                scales: (() => {
                    const maxVal = Math.max(...monthly.map(b => b.chpHeat + b.plbHeat), 100);
                    const yMax = niceMax(maxVal);
                    return {
                        y: {
                            beginAtZero: true, max: yMax,
                            title: { display: true, text: 'Gcal', font: { size: 10 }, color: '#94a3b8' },
                            ticks: { font: { size: 9 }, color: '#94a3b8', stepSize: niceStep(yMax) },
                            grid: { color: 'rgba(255,255,255,0.04)' },
                        },
                    };
                })()
            }
        });
    }

}

// ─── 일별 차트 (선택 월) ───
function renderDailyChart(simResults, month) {
    const rows = month > 0 ? simResults.filter(r => r.month === month) : simResults.slice(0, 31);
    const labels = rows.map(r => `${r.day}일`);

    if (_dailyChart) { _dailyChart.destroy(); _dailyChart = null; }

    // 공휴일/주말 배경색 표시용
    const bgColors = rows.map(r => r.isMaint ? 'rgba(239,68,68,0.10)' : (r.isHoliday ? 'rgba(239,68,68,0.08)' : (r.isWeekend ? 'rgba(99,102,241,0.06)' : 'transparent')));
    const holidayPlugin = {
        id: 'holidayBg',
        beforeDraw(chart) {
            const { ctx: c, chartArea: { left, right, top, bottom }, scales: { x } } = chart;
            rows.forEach((r, i) => {
                if (!r.isHoliday && !r.isWeekend && !r.isMaint) return;
                const xPos = x.getPixelForValue(i);
                const halfBar = (right - left) / rows.length / 2;
                c.fillStyle = r.isMaint ? 'rgba(239,68,68,0.12)' : (r.isHoliday ? 'rgba(239,68,68,0.08)' : 'rgba(99,102,241,0.06)');
                c.fillRect(xPos - halfBar, top, halfBar * 2, bottom - top);
            });
        }
    };

    const ctx = document.getElementById('chartDailyHeat')?.getContext('2d');
    if (ctx) {
        _dailyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    { label: '수요량', data: rows.map(r => Math.round(r.demand)), backgroundColor: '#e2e8f0', stack: 'demand', order: 1 },
                    { label: 'CHP열', data: rows.map(r => Math.round(r.chpHeat)), backgroundColor: '#3b82f6', stack: 'prod', order: 0 },
                    { label: 'PLB열', data: rows.map(r => Math.round(r.plbHeat)), backgroundColor: '#f59e0b', stack: 'prod', order: 0 },
                    { label: '축열조', data: rows.map(r => Math.round(r.storageLevel)), borderColor: '#6366f1', borderWidth: 1.5, type: 'line', pointRadius: 2, stack: false, yAxisID: 'y1', order: -1 },
                ]
            },
            plugins: [holidayPlugin],
            options: {
                responsive: true, maintainAspectRatio: false, animation: false,
                plugins: {
                    legend: { position: 'top', labels: { color: '#aaa', font: { size: 10, family: "'Pretendard',sans-serif" }, boxWidth: 16, boxHeight: 5, padding: 14, useBorderRadius: true, borderRadius: 3 } },
                    tooltip: {
                        callbacks: {
                            title(items) {
                                const i = items[0].dataIndex;
                                const r = rows[i];
                                const tag = r.isMaint ? ' [정비]' : (r.isHoliday ? ' [공휴일]' : (r.isWeekend ? ' [주말]' : ''));
                                return `${r.month}/${r.day}일 (${r.dowName})${tag}`;
                            }
                        }
                    }
                },
                scales: (() => {
                    const storageCap = parseFloat(document.getElementById('storageCapacity')?.value) || 1500;
                    const maxDemand = Math.max(...rows.map(r => r.demand), ...rows.map(r => r.chpHeat + r.plbHeat), 100);
                    const yMax = niceMax(maxDemand);
                    return {
                        y: {
                            beginAtZero: true, position: 'left', max: yMax,
                            title: { display: true, text: '열수급 (Gcal)', font: { size: 10 }, color: '#94a3b8' },
                            ticks: { font: { size: 9 }, color: '#94a3b8', stepSize: niceStep(yMax) },
                            grid: { color: 'rgba(255,255,255,0.04)' },
                        },
                        y1: {
                            beginAtZero: true, max: niceMax(storageCap), position: 'right',
                            title: { display: true, text: '축열조 (Gcal)', font: { size: 10 }, color: '#818cf8' },
                            ticks: { font: { size: 9 }, color: '#818cf8', stepSize: niceStep(niceMax(storageCap)) },
                            grid: { drawOnChartArea: false },
                        },
                        x: {
                            ticks: {
                                font: { size: 10 },
                                color: rows.map(r => r.isMaint ? '#ef4444' : (r.isHoliday ? '#ef4444' : (r.isWeekend ? '#6366f1' : '#666')))
                            }
                        }
                    };
                })()
            }
        });
    }

}

// ─── 최적화 실행 ───
function runOptimization() {
    const simResults = runDailySimulation();
    if (!simResults) {
        alert('판매실적과 운영실적 데이터가 필요합니다');
        return;
    }
    PLAN_RESULTS = simResults;
    if (typeof checkPipelineStatus === 'function') checkPipelineStatus();

    // 시뮬 결과 영역 표시
    const simWrap = document.getElementById('planSimResult');
    if (simWrap) simWrap.style.display = '';

    // 월별 탭 활성화
    document.querySelectorAll('#planTabs .inner-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.plan-panel').forEach(p => p.classList.remove('active'));
    document.querySelector('#planTabs .inner-tab[data-ptab="plan_monthly_sim"]')?.classList.add('active');
    document.getElementById('plan_monthly_sim')?.classList.add('active');

    renderPlanMonthly(simResults);
    renderPlanCharts(simResults);
    renderPlanDaily(simResults);
    renderContribution(simResults);
    renderAllCosts(simResults);

    // 로직 흐름 버튼 활성화
    const btnLogic = document.getElementById('btnShowLogicFlow');
    if (btnLogic) btnLogic.style.display = '';
}

// ═══════════════════════════════════════════════
// 로직 플로우차트 모달
// ═══════════════════════════════════════════════

// ─── 흐름도 헬퍼 함수 ───
function _fcNode(type, title, desc, line) {
    const colors = { start: '#a78bfa', end: '#a78bfa', process: '#3b82f6', action: '#22c55e' };
    const bg = { start: 'rgba(167,139,250,0.12)', end: 'rgba(167,139,250,0.12)', process: 'rgba(59,130,246,0.10)', action: 'rgba(34,197,94,0.08)' };
    const border = colors[type] || '#3b82f6';
    const isRound = type === 'start' || type === 'end';
    const radius = isRound ? '24px' : '10px';
    let h = `<div class="fc-node" style="background:${bg[type]};border:1.5px solid ${border};border-radius:${radius};padding:10px 16px;text-align:center;max-width:280px;margin:0 auto">`;
    h += `<div style="font-size:12px;font-weight:700;color:${border}">${title}</div>`;
    if (desc) h += `<div style="font-size:10px;color:#94a3b8;margin-top:3px;line-height:1.5">${desc}</div>`;
    if (line) h += `<div style="font-size:9px;color:#475569;margin-top:2px;font-family:monospace">[${line}]</div>`;
    h += '</div>';
    return h;
}
function _fcArrow() {
    return '<div style="width:2px;height:20px;background:#334155;margin:0 auto"></div>';
}
function _fcDiamond(question, condition, line) {
    let h = '<div style="text-align:center;margin:0 auto;max-width:280px">';
    h += '<div style="width:160px;height:160px;margin:0 auto;transform:rotate(45deg);background:rgba(139,92,246,0.12);border:1.5px solid rgba(139,92,246,0.5);border-radius:12px;display:flex;align-items:center;justify-content:center">';
    h += `<div style="transform:rotate(-45deg);text-align:center;padding:8px">`;
    h += `<div style="font-size:11px;font-weight:700;color:#c4b5fd">${question}</div>`;
    if (condition) h += `<div style="font-size:9px;color:#7c3aed;margin-top:2px">${condition}</div>`;
    if (line) h += `<div style="font-size:8px;color:#475569;margin-top:2px;font-family:monospace">[${line}]</div>`;
    h += '</div></div></div>';
    return h;
}

function showLogicFlowModal() {
    const old = document.getElementById('logicFlowOverlay');
    if (old) old.remove();

    // inject scoped dark-theme styles + animations (once)
    if (!document.getElementById('lfm-styles')) {
        const st = document.createElement('style');
        st.id = 'lfm-styles';
        st.textContent = `
@keyframes lfm-fadeUp { from { opacity:0; transform:translateY(20px); } to { opacity:1; transform:translateY(0); } }
@keyframes lfm-pulse { 0%,100% { box-shadow:0 0 0 0 rgba(139,92,246,0.3); } 50% { box-shadow:0 0 20px 4px rgba(139,92,246,0.15); } }
#logicFlowOverlay * { box-sizing:border-box; }
#logicFlowOverlay .lfm-box { font-family:'Noto Sans KR','SF Pro Display',-apple-system,sans-serif; }
#logicFlowOverlay .lfm-card {
  background:rgba(30,41,59,0.6); backdrop-filter:blur(12px); border:1px solid rgba(51,65,85,0.5);
  border-radius:16px; padding:20px; transition:transform 0.3s cubic-bezier(.34,1.56,.64,1), box-shadow 0.3s ease;
  animation:lfm-fadeUp 0.5s cubic-bezier(.34,1.56,.64,1) both;
}
#logicFlowOverlay .lfm-card:hover { transform:translateY(-4px); box-shadow:0 12px 40px rgba(0,0,0,0.4); }
#logicFlowOverlay .lfm-step-card {
  background:rgba(30,41,59,0.5); backdrop-filter:blur(12px); border:1px solid rgba(51,65,85,0.4);
  border-radius:16px; padding:16px 20px; transition:transform 0.3s cubic-bezier(.34,1.56,.64,1), box-shadow 0.3s ease;
  animation:lfm-fadeUp 0.5s cubic-bezier(.34,1.56,.64,1) both;
}
#logicFlowOverlay .lfm-step-card:hover { transform:translateY(-3px); box-shadow:0 8px 32px rgba(0,0,0,0.35); }
#logicFlowOverlay .lfm-tab {
  padding:8px 18px; border-radius:10px; border:1px solid rgba(51,65,85,0.6); background:transparent;
  color:#94a3b8; font-size:13px; font-weight:600; cursor:pointer; transition:all 0.25s ease;
}
#logicFlowOverlay .lfm-tab:hover { background:rgba(51,65,85,0.5); color:#e2e8f0; }
#logicFlowOverlay .lfm-tab.active { background:rgba(139,92,246,0.25); border-color:rgba(139,92,246,0.5); color:#a78bfa; }
#logicFlowOverlay .lfm-badge {
  display:inline-block; padding:3px 10px; border-radius:8px; font-size:11px; font-weight:600; line-height:1.5;
}
#logicFlowOverlay .lfm-scroll::-webkit-scrollbar { width:6px; }
#logicFlowOverlay .lfm-scroll::-webkit-scrollbar-track { background:rgba(15,23,42,0.5); }
#logicFlowOverlay .lfm-scroll::-webkit-scrollbar-thumb { background:#334155; border-radius:3px; }
`;
        document.head.appendChild(st);
    }

    const overlay = document.createElement('div');
    overlay.id = 'logicFlowOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.7);z-index:9999;display:flex;align-items:center;justify-content:center;animation:lfm-fadeUp 0.3s ease';

    const box = document.createElement('div');
    box.className = 'lfm-box';
    box.style.cssText = 'background:radial-gradient(ellipse at 30% 40%,rgba(15,23,42,1) 0%,rgba(2,6,23,1) 60%);border-radius:20px;padding:28px 32px;max-width:1300px;width:96%;height:85vh;display:flex;flex-direction:column;box-shadow:0 32px 80px rgba(0,0,0,0.6);border:1px solid rgba(51,65,85,0.4)';

    // 헤더 + 탭 (고정)
    let html = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-shrink:0">';
    html += '<div style="display:flex;align-items:center;gap:16px">';
    html += '<h3 style="margin:0;font-size:20px;font-weight:900;color:#f1f5f9;letter-spacing:-0.02em">시뮬레이션 로직 흐름</h3>';
    html += '<div style="display:flex;gap:6px">';
    html += '<button class="lfm-tab active" data-ltab="intro" onclick="_switchLogicTab(\'intro\')">데이터 준비</button>';
    html += '<button class="lfm-tab" data-ltab="basic" onclick="_switchLogicTab(\'basic\')">전체 흐름</button>';
    html += '<button class="lfm-tab" data-ltab="detail" onclick="_switchLogicTab(\'detail\')">최적화 방식</button>';
    html += '</div></div>';
    html += '<button onclick="document.getElementById(\'logicFlowOverlay\').remove()" style="border:none;background:rgba(51,65,85,0.4);width:36px;height:36px;border-radius:10px;font-size:18px;cursor:pointer;color:#94a3b8;display:flex;align-items:center;justify-content:center;transition:all 0.2s" onmouseenter="this.style.background=\'rgba(239,68,68,0.3)\';this.style.color=\'#fca5a5\'" onmouseleave="this.style.background=\'rgba(51,65,85,0.4)\';this.style.color=\'#94a3b8\'">\u2715</button>';
    html += '</div>';

    // 탭 콘텐츠 래퍼 (고정 높이, 내부 스크롤)
    html += '<div class="lfm-scroll" style="flex:1;overflow-y:auto;min-height:0">';


    // ═══ [1] 데이터 준비 탭 ═══
    html += '<div id="logicPanel_intro" class="logic-panel">';
    html += '<p style="font-size:14px;color:#94a3b8;margin:0 0 20px;line-height:1.7">시뮬레이션을 실행하기 전에 필요한 입력 데이터를 준비합니다. 각 데이터는 CSV 파일로 업로드하거나 직접 입력할 수 있습니다.</p>';

    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px">';

    // 열수요
    html += '<div class="lfm-card" style="animation-delay:0s;border-left:3px solid #ef4444">';
    html += '<div style="font-size:15px;font-weight:800;color:#fca5a5;margin-bottom:12px;letter-spacing:-0.01em">열수요 데이터</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:1.9">';
    html += '건물에서 <b style="color:#f1f5f9">시간마다 얼마만큼의 열을 사용</b>하는지를 나타냅니다.<br>';
    html += '<span style="color:#64748b">입력:</span> 판매실적(시간별) \u2014 월, 일, 시간, 용도별 사용량<br>';
    html += '<span style="color:#64748b">용도:</span> 주택용 난방, 업무용 난방/냉방, 공공용 난방/냉방<br>';
    html += '<span style="color:#64748b">처리:</span> 모든 용도의 사용량을 시간별로 합산<br>';
    html += '<div style="margin-top:10px;padding:10px 14px;background:rgba(239,68,68,0.1);border:1px solid rgba(239,68,68,0.2);border-radius:10px;font-size:11px;color:#fca5a5">';
    html += '시간별 데이터가 있으면 <b>실제 수요 패턴</b>(아침 피크, 야간 감소 등)이 반영되고, 없으면 하루 총량을 24시간으로 균등 분배합니다.</div>';
    html += '</div></div>';

    // 외부수열
    html += '<div class="lfm-card" style="animation-delay:0.08s;border-left:3px solid #22c55e">';
    html += '<div style="font-size:15px;font-weight:800;color:#86efac;margin-bottom:12px;letter-spacing:-0.01em">외부수열 데이터</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:1.9">';
    html += 'CHP/보일러 외에 <b style="color:#f1f5f9">다른 곳에서 공급받는 열</b>을 나타냅니다.<br>';
    html += '<span style="color:#64748b">입력:</span> 운영실적 \u2014 월, 일, 외부수열량(하루 총량)<br>';
    html += '<span style="color:#64748b">인테코:</span> 별도 파일로 인테코 수열/송열 관리<br>';
    html += '<span style="color:#64748b">처리:</span> 하루 총량을 24시간으로 나눠 균등 배분<br>';
    html += '<div style="margin-top:10px;padding:10px 14px;background:rgba(34,197,94,0.1);border:1px solid rgba(34,197,94,0.2);border-radius:10px;font-size:11px;color:#86efac">';
    html += '외부수열이 많은 날은 CHP/보일러를 <b>덜 가동</b>해도 되므로, 연료비를 절감할 수 있습니다.</div>';
    html += '</div></div>';

    // 전력가격
    html += '<div class="lfm-card" style="animation-delay:0.16s;border-left:3px solid #3b82f6">';
    html += '<div style="font-size:15px;font-weight:800;color:#93c5fd;margin-bottom:12px;letter-spacing:-0.01em">전력가격 (SMP/MP)</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:1.9">';
    html += 'CHP가 발전한 전기를 <b style="color:#f1f5f9">한전에 팔 때의 가격</b>입니다.<br>';
    html += '<span style="color:#64748b">SMP:</span> 계통한계가격 \u2014 시간별로 변동<br>';
    html += '<span style="color:#64748b">MP:</span> 정산가격 \u2014 SMP에 손실계수와 정산금을 반영<br>';
    html += '<span style="color:#64748b">처리:</span> SMP \u2192 MP로 자동 변환<br>';
    html += '<div style="margin-top:10px;padding:10px 14px;background:rgba(59,130,246,0.1);border:1px solid rgba(59,130,246,0.2);border-radius:10px;font-size:11px;color:#93c5fd">';
    html += 'MP가 높은 시간대에 CHP를 가동하면 <b>전기 판매 수익</b>이 극대화됩니다. 이것이 최적화의 핵심입니다.</div>';
    html += '</div></div>';

    // 기동비용
    html += '<div class="lfm-card" style="animation-delay:0.24s;border-left:3px solid #a78bfa">';
    html += '<div style="font-size:15px;font-weight:800;color:#c4b5fd;margin-bottom:12px;letter-spacing:-0.01em">기동비용</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:1.9">';
    html += 'CHP를 <b style="color:#f1f5f9">껐다가 다시 켤 때</b> 드는 추가 연료비입니다.<br>';
    html += '<span style="color:#64748b">짧은 정지(1일 이내):</span> 예열 시간 짧음 \u2192 비용 적음<br>';
    html += '<span style="color:#64748b">중간 정지(2일):</span> 예열 시간 보통 \u2192 비용 보통<br>';
    html += '<span style="color:#64748b">긴 정지(3일 이상):</span> 완전히 식음 \u2192 비용 큼<br>';
    html += '<div style="margin-top:10px;padding:10px 14px;background:rgba(167,139,250,0.1);border:1px solid rgba(167,139,250,0.2);border-radius:10px;font-size:11px;color:#c4b5fd">';
    html += '기동비용이 크면 "주말에 끄는 것보다 <b>계속 돌리는 게</b> 나을 수 있다"는 판단이 자동으로 반영됩니다.</div>';
    html += '</div></div>';
    html += '</div>';

    // 설정 데이터 요약
    html += '<div class="lfm-card" style="animation-delay:0.32s;border-top:2px solid rgba(100,116,139,0.3)">';
    html += '<div style="font-size:14px;font-weight:800;color:#e2e8f0;margin-bottom:12px">설정 데이터</div>';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;font-size:12px;color:#94a3b8;line-height:1.8">';
    html += '<div><b style="color:#cbd5e1">CHP 성능표:</b> 부하율별 열생산량, 발전량, 가스소비량</div>';
    html += '<div><b style="color:#cbd5e1">보일러 성능표:</b> 부하율별 열생산량, 가스소비량</div>';
    html += '<div><b style="color:#cbd5e1">축열조:</b> 용량, 초기 잔량, 최소 유지량</div>';
    html += '<div><b style="color:#cbd5e1">NG 단가표:</b> 월별 가스 단가 (원/Nm\u00B3)</div>';
    html += '<div><b style="color:#cbd5e1">열요금표:</b> 월별 기본료, 용도별 단가, 수영장비율</div>';
    html += '<div><b style="color:#cbd5e1">운전 조건:</b> 최소 가동시간, 최소 정지시간, 최소 부하율</div>';
    html += '</div></div>';
    html += '</div>';

    // ═══ [2] 전체 흐름 탭 ═══
    html += '<div id="logicPanel_basic" class="logic-panel" style="display:none">';
    html += '<p style="font-size:14px;color:#94a3b8;margin:0 0 20px;line-height:1.7">시뮬레이션은 <b style="color:#e2e8f0">4단계</b>를 거쳐 진행됩니다. 각 단계의 결과가 다음 단계의 입력이 됩니다.</p>';

    const steps = [
        { n:'1', color:'#8b5cf6', glow:'139,92,246', title:'연간 가동계획 수립', desc:'1년 365일 각각에 대해 <b style="color:#f1f5f9">"CHP를 몇 시간, 어느 부하율로 가동할 것인가"</b>를 결정합니다.<br>이때 열수요, 외부수열, 전력가격, 기동비용, 축열조 상태를 <b style="color:#f1f5f9">모두 종합</b>하여 가장 경제적인 조합을 찾습니다.', badge:'예: "1월 15일 \u2192 CHP 18시간, 100% 부하" / "4월 2일 \u2192 CHP 0시간(정비일)"', badgeBg:'rgba(139,92,246,0.15)', badgeColor:'#c4b5fd' },
        { n:'2', color:'#06b6d4', glow:'6,182,212', title:'하루 중 가동 시간대 배치', desc:'1단계에서 "18시간 가동"이 결정되면, <b style="color:#f1f5f9">하루 중 어느 시간에 켜고 끌지</b>를 정합니다.<br>전력가격이 높은 시간대에 가동하면 전기 판매 수익이 커지고, 열수요가 많은 시간대에 맞추면 축열조가 안정됩니다.<br>모든 가능한 시작 시간을 시뮬레이션해서 <b style="color:#f1f5f9">가장 수익이 높고 축열조가 안전한</b> 배치를 선택합니다.', badge:'보일러(PLB)도 축열조가 부족해지는 시점을 예측하여 자동 배치됩니다', badgeBg:'rgba(6,182,212,0.15)', badgeColor:'#67e8f9' },
        { n:'3', color:'#10b981', glow:'16,185,129', title:'시간별 시뮬레이션 실행', desc:'2단계에서 정한 배치대로 <b style="color:#f1f5f9">0시부터 23시까지 한 시간씩</b> 시뮬레이션합니다.<br>매 시간 축열조에 들어오는 열(외부수열, CHP, 보일러)과 나가는 열(수요)을 계산하여 <b style="color:#f1f5f9">축열조 잔량</b>을 추적합니다.<br>전력가격과 CHP 발전량을 곱해 <b style="color:#f1f5f9">시간별 전기 판매 수익</b>도 계산합니다.', badge:'각 날의 마지막 축열조 잔량이 다음 날의 시작값이 됩니다', badgeBg:'rgba(16,185,129,0.15)', badgeColor:'#6ee7b7' },
        { n:'4', color:'#f59e0b', glow:'245,158,11', title:'결과 집계 및 손익 계산', desc:'시간별 결과를 <b style="color:#f1f5f9">일별 \u2192 월별</b>로 모아서 최종 보고서를 만듭니다.<br><b style="color:#cbd5e1">비용:</b> CHP 연료비 + 보일러 연료비 + 기동비용<br><b style="color:#cbd5e1">수입:</b> 전력 판매(열제약매출) + 용량요금(한전 고정 수입)<br><b style="color:#cbd5e1">운영손익:</b> 수입 - 비용 = 공헌이익', badge:'월별 차트, 일별 상세 보기, 열판매 수익 등을 확인할 수 있습니다', badgeBg:'rgba(245,158,11,0.15)', badgeColor:'#fcd34d' },
    ];
    steps.forEach((s, i) => {
        html += `<div style="display:flex;gap:16px;margin-bottom:${i<3?'16':'0'}px">`;
        html += `<div style="flex-shrink:0;width:52px;height:52px;background:${s.color};border-radius:14px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:22px;font-weight:900;box-shadow:0 4px 20px rgba(${s.glow},0.35);animation:lfm-fadeUp 0.5s cubic-bezier(.34,1.56,.64,1) ${i*0.1}s both">${s.n}</div>`;
        html += `<div class="lfm-step-card" style="flex:1;animation-delay:${i*0.1}s;border-left:2px solid ${s.color}40">`;
        html += `<div style="font-size:16px;font-weight:800;color:${s.color};margin-bottom:8px;letter-spacing:-0.02em">${s.title}</div>`;
        html += `<div style="font-size:12px;color:#cbd5e1;line-height:1.9">${s.desc}<br>`;
        html += `<span class="lfm-badge" style="margin-top:8px;background:${s.badgeBg};color:${s.badgeColor}">${s.badge}</span>`;
        html += '</div></div></div>';
    });
    html += '</div>';

    // ═══ [3] 최적화 방식 탭 ═══
    html += '<div id="logicPanel_detail" class="logic-panel" style="display:none">';
    html += '<p style="font-size:14px;color:#94a3b8;margin:0 0 20px;line-height:1.7">1단계(연간 가동계획)에서 사용할 수 있는 <b style="color:#e2e8f0">두 가지 최적화 방식</b>을 비교합니다.</p>';

    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px">';

    // 정밀 최적화
    html += '<div class="lfm-card" style="animation-delay:0s;border-top:3px solid #8b5cf6">';
    html += '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">';
    html += '<div style="width:36px;height:36px;background:linear-gradient(135deg,#8b5cf6,#6d28d9);border-radius:10px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:14px;font-weight:900;box-shadow:0 4px 16px rgba(139,92,246,0.3)">DP</div>';
    html += '<div style="font-size:17px;font-weight:900;color:#c4b5fd;letter-spacing:-0.02em">정밀 최적화</div>';
    html += '</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:2.0">';
    html += '<div style="margin-bottom:10px;font-weight:700;color:#e2e8f0;font-size:13px">작동 원리</div>';
    html += '1년 전체를 <b style="color:#f1f5f9">한꺼번에</b> 보고, 가장 비용이 적은 경로를 찾습니다.<br>';
    html += '마지막 날(12/31)부터 첫날(1/1)로 <b style="color:#f1f5f9">거꾸로</b> 계산하면서,<br>';
    html += '"오늘 CHP를 켜면 내일 이후 비용이 어떻게 달라지나?"를 모든 경우에 대해 비교합니다.<br><br>';
    html += '<div style="margin-bottom:10px;font-weight:700;color:#e2e8f0;font-size:13px">고려하는 요소</div>';
    html += '<span style="color:#a78bfa">\u2022</span> 현재 축열조에 열이 얼마나 남아 있는지<br>';
    html += '<span style="color:#a78bfa">\u2022</span> CHP가 꺼진 지 며칠 됐는지 (기동비용 결정)<br>';
    html += '<span style="color:#a78bfa">\u2022</span> 각 부하율(60%~100%)에서의 연료비와 전기 수익<br>';
    html += '<span style="color:#a78bfa">\u2022</span> 내일부터 연말까지의 총 비용까지 모두 고려<br><br>';
    html += '<div style="padding:12px 16px;background:rgba(139,92,246,0.1);border:1px solid rgba(139,92,246,0.2);border-radius:12px;font-size:11px;color:#c4b5fd;line-height:1.7">';
    html += '<b>장점:</b> 전체적으로 가장 저렴한 답을 보장<br>';
    html += '<b>예시:</b> "주말에 끄면 월요일 기동비가 비싸니까, 주말에도 최소 부하로 돌리는 게 낫다"는 판단을 자동으로 내림';
    html += '</div>';
    html += '</div></div>';

    // 빠른 최적화
    html += '<div class="lfm-card" style="animation-delay:0.1s;border-top:3px solid #22c55e">';
    html += '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">';
    html += '<div style="width:36px;height:36px;background:linear-gradient(135deg,#22c55e,#16a34a);border-radius:10px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:900;box-shadow:0 4px 16px rgba(34,197,94,0.3)">Fast</div>';
    html += '<div style="font-size:17px;font-weight:900;color:#86efac;letter-spacing:-0.02em">빠른 최적화</div>';
    html += '</div>';
    html += '<div style="font-size:12px;color:#cbd5e1;line-height:2.0">';
    html += '<div style="margin-bottom:10px;font-weight:700;color:#e2e8f0;font-size:13px">작동 원리</div>';
    html += '두 번에 걸쳐 판단합니다:<br><br>';
    html += '<b style="color:#86efac">1차 판단 \u2014 우선순위 정하기</b><br>';
    html += '365일을 전력가격(SMP) 순으로 정렬합니다.<br>';
    html += 'SMP가 높은 날(전기 비싸게 팔 수 있는 날)부터 CHP 가동을 배정합니다.<br><br>';
    html += '<b style="color:#86efac">2차 판단 \u2014 날짜 순서대로 시뮬레이션</b><br>';
    html += '1월 1일부터 시간 순서대로 하루씩 시뮬레이션합니다.<br>';
    html += '축열조가 넘칠 것 같으면 CHP 가동시간을 줄이고,<br>';
    html += '전날 CHP를 켰는데 오늘 끄면 기동비가 발생하므로 계속 돌릴지 비교합니다.<br><br>';
    html += '<div style="padding:12px 16px;background:rgba(34,197,94,0.1);border:1px solid rgba(34,197,94,0.2);border-radius:12px;font-size:11px;color:#86efac;line-height:1.7">';
    html += '<b>장점:</b> 즉시 실행, 직관적인 결과<br>';
    html += '<b>한계:</b> "하루 단위"로 판단하므로, 전체적으로는 정밀 최적화보다 비용이 다소 높을 수 있음';
    html += '</div>';
    html += '</div></div>';
    html += '</div>';

    // 핵심 개념 설명
    html += '<div class="lfm-card" style="animation-delay:0.2s;border-top:2px solid rgba(100,116,139,0.3)">';
    html += '<div style="font-size:15px;font-weight:800;color:#e2e8f0;margin-bottom:14px;letter-spacing:-0.01em">핵심 개념 설명</div>';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;font-size:12px;color:#94a3b8;line-height:1.9">';

    html += '<div style="padding:14px;background:rgba(15,23,42,0.6);border-radius:12px;border:1px solid rgba(51,65,85,0.4)">';
    html += '<div style="font-weight:800;color:#e2e8f0;margin-bottom:8px;font-size:13px">축열조란?</div>';
    html += '열을 임시로 저장하는 큰 물탱크입니다. CHP/보일러가 만든 열을 저장했다가, 수요가 있을 때 공급합니다. 항상 최소 잔량 이상을 유지해야 합니다.</div>';

    html += '<div style="padding:14px;background:rgba(15,23,42,0.6);border-radius:12px;border:1px solid rgba(51,65,85,0.4)">';
    html += '<div style="font-weight:800;color:#e2e8f0;margin-bottom:8px;font-size:13px">CHP vs 보일러(PLB)</div>';
    html += 'CHP는 열과 전기를 동시에 만듭니다. 전기를 팔아 수익을 얻지만 효율이 보일러보다 낮습니다. 보일러는 열만 만들지만 더 효율적입니다. 보일러는 CHP만으로 부족할 때 보조로 투입됩니다.</div>';

    html += '<div style="padding:14px;background:rgba(15,23,42,0.6);border-radius:12px;border:1px solid rgba(51,65,85,0.4)">';
    html += '<div style="font-weight:800;color:#e2e8f0;margin-bottom:8px;font-size:13px">최적화의 목표</div>';
    html += '"열을 안정적으로 공급하면서, 총 비용(연료비 + 기동비 - 전기판매)을 최소화하는 것"입니다. CHP를 많이 돌리면 전기는 많이 팔지만 연료비도 늘어나므로, 균형점을 찾는 것이 핵심입니다.</div>';

    html += '</div></div>';
    html += '</div>';

    html += '</div>'; // 스크롤 래퍼 end

    box.innerHTML = html;
    overlay.appendChild(box);
    document.body.appendChild(overlay);

    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
    document.addEventListener('keydown', function _lfKey(e) {
        if (e.key === 'Escape') { overlay.remove(); document.removeEventListener('keydown', _lfKey); }
    });
}

function _switchLogicTab(tab) {
    document.querySelectorAll('#logicFlowOverlay .lfm-tab').forEach(b => b.classList.remove('active'));
    document.querySelector(`#logicFlowOverlay .lfm-tab[data-ltab="${tab}"]`)?.classList.add('active');
    document.querySelectorAll('#logicFlowOverlay .logic-panel').forEach(p => p.style.display = 'none');
    const panel = document.getElementById(`logicPanel_${tab}`);
    if (panel) {
        panel.style.display = '';
        // replay stagger animations on tab switch
        panel.querySelectorAll('.lfm-card,.lfm-step-card').forEach(el => {
            el.style.animation = 'none';
            el.offsetHeight; // trigger reflow
            el.style.animation = '';
        });
    }
}

/* ═══════════════════════════════════════════════
   탭 전환 이벤트
   ═══════════════════════════════════════════════ */

// ═══ 파이프라인 상태 관리 ═══
const PIPE_STEPS = document.querySelectorAll('.pipe-step');
const PIPE_ARROWS = document.querySelectorAll('.pipe-arrow');

// 필수 데이터 키 (이것만 있으면 시뮬 가능)
const REQUIRED_DATA = ['smp_hourly', 'operations_daily', 'ng_price_monthly', 'cp_monthly'];
// 필수 설정 (값이 0이 아닌지)
const REQUIRED_SETTINGS = ['chpCapacity', 'storageCapacity'];

function checkPipelineStatus() {
    // Step 0: 데이터 — 필수 데이터 입력 여부
    const dataLoaded = REQUIRED_DATA.filter(k => DATA[k] && DATA[k].rows && DATA[k].rows.length > 0).length;
    const dataTotal = REQUIRED_DATA.length;
    const dataDone = dataLoaded >= dataTotal;

    // Step 1: 설정 — 주요 설정값 존재 여부
    const settingsOk = REQUIRED_SETTINGS.every(id => {
        const el = document.getElementById(id);
        return el && parseFloat(el.value) > 0;
    });

    // Step 2: 데이터 산출 — 데이터+설정 충족 시 ready, CALC_DONE 시 done
    const calcReady = dataDone && settingsOk;
    const calcDone = !!window._calcDone;

    // Step 3: 기동계획 — 데이터 산출 완료 시 ready, PLAN_RESULTS 있으면 done
    const planReady = calcDone;
    const planDone = !!PLAN_RESULTS;

    // Step 4: 비용산출 — 기동계획 완료 시 ready/done
    const costReady = planDone;
    const costDone = planDone;

    // Step 5: 운영손익 — 비용산출과 동일
    const resultReady = planDone;
    const resultDone = planDone;

    const states = [
        { done: dataDone, ready: true, status: dataDone ? '완료' : `${dataLoaded}/${dataTotal}` },
        { done: settingsOk, ready: true, status: settingsOk ? '완료' : '미설정' },
        { done: calcDone, ready: calcReady, status: calcDone ? '완료' : calcReady ? '준비됨' : '대기' },
        { done: planDone, ready: planReady, status: planDone ? '완료' : planReady ? '준비됨' : '대기' },
        { done: costDone, ready: costReady, status: costDone ? '완료' : '대기' },
        { done: resultDone, ready: resultReady, status: resultDone ? '완료' : '대기' },
    ];

    PIPE_STEPS.forEach((step, i) => {
        const s = states[i];
        step.classList.remove('locked', 'ready', 'done');
        if (s.done) step.classList.add('done');
        else if (s.ready) step.classList.add('ready');
        else step.classList.add('locked');

        const statusEl = step.querySelector('.pipe-status');
        if (statusEl) {
            statusEl.textContent = s.status;
            statusEl.style.display = (s.status && s.status !== '대기') ? 'inline-block' : 'none';
        }
    });

    // 화살표 상태
    PIPE_ARROWS.forEach((arrow, i) => {
        arrow.classList.remove('lit', 'active');
        if (states[i] && states[i].done) arrow.classList.add('lit');
        else if (states[i] && states[i].ready) arrow.classList.add('active');
    });
}

// 메인 탭 전환
document.querySelectorAll('.pipe-step').forEach(btn => {
    btn.addEventListener('click', () => {
        if (btn.classList.contains('locked')) return; // 잠긴 탭 클릭 불가
        document.querySelectorAll('.pipe-step').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.tab).classList.add('active');
        // 데이터 산출 탭 진입 시 자동 산출
        if (btn.dataset.tab === 'tab-calc') {
            runAllCalc();
        }
        // 운영손익 탭 진입 시 차트 리렌더
        if (btn.dataset.tab === 'tab-result' && PLAN_RESULTS) {
            requestAnimationFrame(() => renderContribution(PLAN_RESULTS));
        }
        // 비용산출 탭 진입 시 렌더
        if (btn.dataset.tab === 'tab-cost') {
            requestAnimationFrame(() => renderAllCosts(PLAN_RESULTS));
        }
    });
});

// 데이터 주파수 탭
document.querySelectorAll('.freq-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.freq-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.freq-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('fp_' + btn.dataset.freq).classList.add('active');
    });
});

// 이너 탭 (데이터 - 입력)
document.querySelectorAll('#inputTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#inputTabs .inner-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('#fp_input .data-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('dp_' + btn.dataset.key).classList.add('active');
    });
});

// 이너 탭 (데이터 - 단가)
document.querySelectorAll('#pricingTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#pricingTabs .inner-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('#fp_pricing .data-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('dp_' + btn.dataset.key).classList.add('active');
    });
});

// 설정 스텝 탭
document.querySelectorAll('.step-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.step-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.step).classList.add('active');
    });
});

// ─── 설정값 변경 감지 & 적용 버튼 ───
(function() {
    const settingsContainer = document.getElementById('tab-settings');
    const applyBar = document.getElementById('settingsApplyBar');
    const applyBtn = document.getElementById('settingsApplyBtn');
    const resetBtn = document.getElementById('settingsResetBtn');
    if (!settingsContainer || !applyBar) return;

    // 고유 키 생성: id가 있으면 id, 없으면 테이블명+행+셀 인덱스
    function inputKey(inp) {
        if (inp.id) return inp.id;
        const tr = inp.closest('tr');
        const td = inp.closest('td');
        const tbl = inp.closest('table');
        const tblId = tbl ? (tbl.id || tbl.closest('.step-panel')?.id || '') : '';
        return tblId + '_r' + (tr ? tr.rowIndex : '') + '_c' + (td ? td.cellIndex : '');
    }

    // 스냅샷
    let snapshot = {};
    function takeSnapshot() {
        snapshot = {};
        settingsContainer.querySelectorAll('input[type="text"]').forEach(inp => {
            snapshot[inputKey(inp)] = inp.value;
        });
    }
    takeSnapshot();

    // 변경 감지
    function checkChanged() {
        let changed = false;
        settingsContainer.querySelectorAll('input[type="text"]').forEach(inp => {
            const k = inputKey(inp);
            if (snapshot[k] !== undefined && snapshot[k] !== inp.value) {
                changed = true;
                inp.style.outline = '2px solid #FFB74D';
            } else {
                inp.style.outline = '';
            }
        });
        applyBar.style.display = changed ? 'flex' : 'none';
    }

    settingsContainer.addEventListener('input', checkChanged);

    // 적용
    applyBtn.addEventListener('click', () => {
        takeSnapshot();
        applyBar.style.display = 'none';
        settingsContainer.querySelectorAll('input[type="text"]').forEach(inp => inp.style.outline = '');
        if (typeof updatePriceCalc === 'function') updatePriceCalc();
        if (typeof runAllCalc === 'function') runAllCalc();
        applyBtn.textContent = '적용 완료 ✓';
        applyBtn.style.background = '#22C55E';
        setTimeout(() => { applyBtn.textContent = '값 적용'; applyBtn.style.background = '#4B83F0'; }, 1500);
    });

    // 되돌리기
    resetBtn.addEventListener('click', () => {
        settingsContainer.querySelectorAll('input[type="text"]').forEach(inp => {
            const k = inputKey(inp);
            if (snapshot[k] !== undefined) inp.value = snapshot[k];
            inp.style.outline = '';
        });
        applyBar.style.display = 'none';
    });

    // 설정 탭 진입 시 스냅샷 갱신
    document.querySelectorAll('.pipe-step').forEach(btn => {
        btn.addEventListener('click', () => {
            if (btn.dataset.tab === 'tab-settings') setTimeout(takeSnapshot, 50);
        });
    });
})();

// 기동계획 탭
document.querySelectorAll('#planTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#planTabs .inner-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.plan-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.ptab).classList.add('active');
    });
});

// 데이터 산출 탭
document.querySelectorAll('#calcTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#calcTabs .inner-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.calc-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.ctab).classList.add('active');
    });
});

// ─── 항목 설명 tooltip ───
(function() {
    const box = document.createElement('div');
    box.id = 'tipBox';
    document.body.appendChild(box);
    let currentTip = null;

    setInterval(() => {
        if (!currentTip) { box.style.display = 'none'; return; }
        if (!currentTip.matches(':hover')) {
            box.style.display = 'none';
            currentTip = null;
        }
    }, 100);

    document.body.addEventListener('mouseover', function(e) {
        const el = e.target.closest('.has-tip');
        if (!el || el === currentTip) return;
        currentTip = el;
        const tip = el.getAttribute('data-tip');
        if (!tip) return;
        box.innerHTML = tip.split('|').map((line, i) =>
            i === 0 ? '<b style="color:#93c5fd">' + line + '</b>' : line
        ).join('<br>');
        const r = el.getBoundingClientRect();
        box.style.left = Math.min(r.left, window.innerWidth - 340) + 'px';
        box.style.top = (r.bottom + 8) + 'px';
        box.style.display = 'block';
    });
})();

// 비용산출 탭 전환
document.querySelectorAll('#costTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#costTabs .inner-tab').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.cost-panel').forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.costtab)?.classList.add('active');
    });
});

// 전력수급 서브탭 전환
document.querySelectorAll('.elec-data-subtabs .elec-sub').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.elec-data-subtabs .elec-sub').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.elec-data-pane').forEach(p => p.classList.remove('active'));
        const target = document.getElementById('edp_' + btn.dataset.edkey);
        if (target) target.classList.add('active');
    });
});

// 전기요금표 서브탭 전환
document.querySelectorAll('#elecRateTabs .inner-tab').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('#elecRateTabs .inner-tab').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.elec-rate-panel').forEach(p => p.classList.remove('active'));
        const target = document.getElementById('erp_' + btn.dataset.ertab);
        if (target) target.classList.add('active');
    });
});

// 설정 자동저장 + 단가산정 자동계산
document.getElementById('tab-settings').addEventListener('input', () => {
    scheduleSettingsSave();
    updatePricingCalc();
});

// 단가산정 테이블 직접 리스너 (즉각적인 계산 반영)
document.getElementById('pricingHeatTable')?.addEventListener('input', () => {
    updatePricingCalc();
    scheduleSettingsSave();
});
// 계수 변경 시에도 전 월 무부하비정산 재계산
['pricingInc1', 'pricingInc2', 'pricingConst'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', updatePricingCalc);
});

// 월별 편집 가능 테이블 이벤트 위임 (input → 자동저장, paste → 엑셀 붙여넣기)
MONTHLY_KEYS.forEach(key => {
    const body = document.getElementById('pv_' + key);
    if (!body) return;
    body.addEventListener('input', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('.editable-tbl')) {
            scheduleMonthlyAutoSave(key);
        }
    });
    body.addEventListener('paste', (e) => {
        if (e.target.tagName === 'INPUT' && e.target.closest('.editable-tbl')) {
            handleMonthlyPaste(e, key);
        }
    });
});

/* ═══════════════════════════════════════════════
   초기화
   ═══════════════════════════════════════════════ */
window.addEventListener('pywebviewready', () => loadMainScreen());
