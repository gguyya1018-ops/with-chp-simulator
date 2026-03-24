"""
MPC 테스트 앱 — pywebview 기반, 자동 실행 후 종료
실제 앱과 동일한 브라우저 엔진(V8)으로 실행하여 동일한 결과 보장

Usage:
    python test/test_app.py --project 24위드
    python test/test_app.py --project 24위드 --output test/results/result.json
    python test/test_app.py --project 24위드 --maxdays 31
"""
import webview
import json
import os
import sys
import glob
import argparse
import calendar
import time
import threading

# 프로젝트 루트 (test/ 상위)
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

parser = argparse.ArgumentParser()
parser.add_argument('--project', default='24위드')
parser.add_argument('--output', default=os.path.join(ROOT_DIR, 'test', 'results', 'test_result.json'))
parser.add_argument('--maxdays', type=int, default=366)
parser.add_argument('--horizon', type=int, default=30)
args = parser.parse_args()


class TestApi:
    """실제 앱의 Api와 동일한 인터페이스 (테스트에 필요한 것만)"""

    def list_projects(self):
        data_dir = os.path.join(ROOT_DIR, 'data')
        if not os.path.isdir(data_dir):
            return []
        return sorted([d for d in os.listdir(data_dir)
                       if os.path.isdir(os.path.join(data_dir, d))])

    def create_project(self, name):
        return {'ok': True}

    def load_project(self, project_name):
        proj_dir = os.path.join(ROOT_DIR, 'data', project_name)
        if not os.path.isdir(proj_dir):
            return {'error': f'프로젝트 폴더 없음: {proj_dir}'}
        result = {}
        for csv_path in glob.glob(os.path.join(proj_dir, '*.csv')):
            key = os.path.splitext(os.path.basename(csv_path))[0]
            try:
                with open(csv_path, 'r', encoding='utf-8-sig') as f:
                    result[key] = f.read()
            except Exception as e:
                result[key] = f'ERROR:{e}'
        return result

    def save_csv(self, project_name, file_key, content):
        return {'ok': True}

    def get_year_days(self, year):
        weekdays = ['월', '화', '수', '목', '금', '토', '일']
        days = []
        for m in range(1, 13):
            max_d = calendar.monthrange(year, m)[1]
            for d in range(1, max_d + 1):
                wd_idx = calendar.weekday(year, m, d)
                days.append({'m': m, 'd': d, 'wd': weekdays[wd_idx], 'wdi': wd_idx})
        return days

    def save_test_result(self, json_str):
        """테스트 결과 JSON을 파일로 저장"""
        os.makedirs(os.path.dirname(args.output), exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(json_str)
        print(f'[TestApp] 결과 저장: {args.output}')
        return {'ok': True}

    def get_test_config(self):
        """테스트 설정 반환"""
        return {
            'project': args.project,
            'horizon': args.horizon,
            'maxdays': args.maxdays,
        }

    def test_done(self):
        """테스트 완료 → 앱 종료"""
        print('[TestApp] 완료, 종료합니다.')
        # 약간 딜레이 후 종료 (JS 응답 보내기 위해)
        threading.Timer(0.5, lambda: window.destroy()).start()
        return {'ok': True}

    # 더미 (앱 코드에서 호출될 수 있는 것들)
    def win_minimize(self): pass
    def win_toggle_max(self): pass
    def win_close(self): window.destroy()
    def open_file_dialog(self, *a): return None
    def save_file_dialog(self, *a): return None
    def export_excel(self, *a): return None
    def delete_csv(self, *a): return {'ok': True}
    def delete_project(self, *a): return {'ok': True}


api = TestApi()
window = webview.create_window(
    'MPC Test Runner',
    os.path.join(ROOT_DIR, 'index.html'),
    js_api=api, width=800, height=600,
    background_color='#0a0a0a',
)


def on_loaded():
    """윈도우 로드 후 자동 테스트 실행"""
    time.sleep(1)  # DOM 안정화 대기
    print(f'[TestApp] 프로젝트: {args.project}, horizon: {args.horizon}, maxdays: {args.maxdays}')

    # 전체를 하나의 async로 실행 (openProject 포함)
    js_code = f"""
    (async function() {{
        // 진행률 오버레이 생성
        const overlay = document.createElement('div');
        overlay.id = 'testOverlay';
        overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(10,10,10,0.95);z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:monospace;color:#e2e8f0';
        overlay.innerHTML = '<div style="font-size:24px;font-weight:bold;margin-bottom:20px">MPC 테스트 실행 중</div>'
            + '<div id="testStatus" style="font-size:16px;color:#94a3b8;margin-bottom:30px">프로젝트 로딩...</div>'
            + '<div style="width:400px;height:8px;background:#1e293b;border-radius:4px;overflow:hidden"><div id="testBar" style="width:0%;height:100%;background:#3b82f6;transition:width 0.3s"></div></div>'
            + '<div id="testPct" style="font-size:14px;color:#64748b;margin-top:10px">0%</div>'
            + '<div id="testLog" style="margin-top:30px;font-size:12px;color:#475569;max-height:200px;overflow:auto;width:500px"></div>';
        document.body.appendChild(overlay);

        function setStatus(msg) {{ document.getElementById('testStatus').textContent = msg; }}
        function setProgress(pct) {{ document.getElementById('testBar').style.width = pct+'%'; document.getElementById('testPct').textContent = pct+'%'; }}
        function addLog(msg) {{ const el = document.getElementById('testLog'); el.innerHTML += msg + '<br>'; el.scrollTop = el.scrollHeight; }}

        try {{
            // 1. 프로젝트 열기
            setStatus('프로젝트 로딩: {args.project}');
            await openProject('{args.project}');
            addLog('프로젝트 로드 완료');

            // 2. solveMPC 실행
            setStatus('MPC 최적화 실행 중...');
            const startT = Date.now();
            const results = await solveMPCAsync({args.horizon}, (pct, day) => {{
                setProgress(pct);
                setStatus('MPC 최적화: ' + pct + '% (' + day + '/' + 366 + '일)');
            }});

            if (!results || results.length === 0) {{
                setStatus('결과 없음 - 종료');
                await pywebview.api.test_done();
                return;
            }}

            const maxdays = {args.maxdays};
            const limited = maxdays < results.length ? results.slice(0, maxdays) : results;
            const elapsed = ((Date.now() - startT) / 1000).toFixed(1);
            setStatus('완료! ' + limited.length + '일, ' + elapsed + '초');
            setProgress(100);
            addLog('PLB 가동일: ' + limited.filter(r => r.plbHeat > 0).length + '일');
            addLog('PLB 총열량: ' + Math.round(limited.reduce((s,r) => s + (r.plbHeat||0), 0)) + ' Gcal');
            addLog('저장 중...');

            // 월별 집계
            let monthly = null;
            try {{ monthly = aggregateMonthly(results); }} catch(e) {{}}

            // JSON 구성
            const warnings = results._warnings || [];
            const output = {{
                project: '{args.project}',
                horizon: {args.horizon},
                elapsed: parseFloat(elapsed),
                summary: {{
                    totalDays: limited.length,
                    plbDays: limited.filter(r => r.plbHeat > 0).length,
                    totalPlbHeat: Math.round(limited.reduce((s,r) => s + (r.plbHeat||0), 0)),
                    totalPlbFuelCost: Math.round(limited.reduce((s,r) => s + (r.plbFuelCost||0), 0)),
                    totalChpFuelCost: Math.round(limited.reduce((s,r) => s + (r.chpFuelCost||0), 0)),
                    totalStartupCost: Math.round(limited.reduce((s,r) => s + (r.startupCost||0), 0)),
                    totalNetCost: Math.round(limited.reduce((s,r) => s + (r.totalNetCost||0), 0)),
                    underflows: warnings.filter(w => w.type === 'underflow').length,
                    overflows: warnings.filter(w => w.type === 'overflow').length,
                }},
                daily: limited.map(r => ({{
                    month: r.month, day: r.day, dayType: r.dayType,
                    storageStart: r.storageStart, storageEnd: r.storageLevel,
                    chpHours: r.chpHours, chpLoad: r.chpLoad,
                    chpHeat: Math.round(r.chpHeat||0), chpStart: r.chpStart,
                    plbUnits: r.plbUnits||0, plbHours: r.plbHours||0,
                    plbStartH: r.plbStartH||0, plbHeat: Math.round(r.plbHeat||0),
                    chpFuelCost: Math.round(r.chpFuelCost||0),
                    plbFuelCost: Math.round(r.plbFuelCost||0),
                    startupCost: Math.round(r.startupCost||0),
                    totalNetCost: Math.round(r.totalNetCost||0),
                    hourly: r.hourly ? r.hourly.map(h => ({{
                        storage: Math.round(h.storage),
                        chp: Math.round((h.chp||0)*100)/100,
                        plb: Math.round((h.plb||0)*100)/100,
                        demand: Math.round((h.demand||0)*100)/100,
                    }})) : null,
                }})),
                monthly: monthly,
                warnings: warnings.map(w => ({{ type: w.type, day: w.day, msg: w.msg }})),
            }};

            // 저장
            await pywebview.api.save_test_result(JSON.stringify(output, null, 2));
            await pywebview.api.test_done();
        }} catch(e) {{
            console.error('[Test] 오류:', e.message);
            // 에러를 결과 파일에 저장
            await pywebview.api.save_test_result(JSON.stringify({{error: e.message, stack: e.stack}}, null, 2));
            await pywebview.api.test_done();
        }}
    }})();
    """
    window.evaluate_js(js_code)


print(f'[TestApp] 시작...')
webview.start(func=on_loaded, debug=False)
print(f'[TestApp] 종료')
