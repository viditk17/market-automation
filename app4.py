import os
import uuid
import threading
import sys
import traceback
import calendar
import requests
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# Load environment variables from .env file (if exists)
load_dotenv()

# Get base directory for templates and static
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

app = Flask(__name__, template_folder=TEMPLATE_DIR)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.jinja_env.auto_reload = True
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['OUTPUT_FOLDER'] = os.path.join(BASE_DIR, 'outputs')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

jobs = {}
fetch_jobs = {}  # for async broker_master / no-date fetches

# ── OM Insights credentials (from environment variables or .env file) ──
OM_EMAIL = os.getenv('OM_EMAIL', 'vidit.kalra@olsc.in')
OM_PASSWORD = os.getenv('OM_PASSWORD', 'v#1234K%')
OM_BASE_URL = "https://ominsights.omlogistics.co.in"

# ── Fetchable report configs (from work_assign_summary_automated.py + screenshots) ──
FETCHABLE_REPORTS = {
    "work_summary": {
        "name": "0016 - WORK ASSIGN SUMMARY",
        "card_id": "1293",
        "params": [
            {"name": "P_EMP_CODE", "type": "category", "value": "*"},
            {"name": "P_FROM_DT", "type": "date/single", "role": "from"},
            {"name": "P_TO_DT", "type": "date/single", "role": "to"}
        ],
        "format": "xlsx",
        "needs_dates": True
    },
    "staff_detail": {
        "name": "0003 - VEHICLE HIRING INCENTIVE",
        "card_id": "1272",
        "params": [
            {"name": "FROM_DT", "type": "date/single", "role": "from"},
            {"name": "TO_DT", "type": "date/single", "role": "to"}
        ],
        "format": "csv",
        "needs_dates": True
    },
    "broker_master": {
        "name": "0005 - BROKER MASTER REPORT",
        "card_id": "1274",
        "params": [],
        "format": "csv",
        "needs_dates": False
    }
}

# ── Required files for processing ──
REQUIRED_FILES = [
    {'key': 'branch_master',    'label': 'BRANCH MASTER XL',             'icon': '📋', 'desc': 'Zone lookup mapping',          'fetchable': True,  'fetch_type': 'no_dates'},
    {'key': 'work_summary',     'label': 'Work Assign Summary (0016)',    'icon': '📊', 'desc': 'Main data source',             'fetchable': True,  'fetch_type': 'date_range'},
    {'key': 'cancel_report',    'label': 'Cancel Remark Report',          'icon': '❌', 'desc': 'Cancellation analysis',        'fetchable': False, 'fetch_type': None},
    {'key': 'challenge_report', 'label': 'Challenge Price Report',        'icon': '💰', 'desc': 'Bid & savings data',           'fetchable': False, 'fetch_type': None},
    {'key': 'staff_detail',     'label': 'Vehicle Hiring Incentive (0003)','icon': '🚛', 'desc': 'Staff detail report',          'fetchable': True,  'fetch_type': 'date_range'},
    {'key': 'broker_master',    'label': 'Broker Master (0005)',          'icon': '🤝', 'desc': 'New vendor registration',      'fetchable': True,  'fetch_type': 'no_dates'},
    {'key': 'ho_file',          'label': 'HO & Branch Segregation',       'icon': '🏢', 'desc': 'Employee code segregation',    'fetchable': False, 'fetch_type': None},
]


@app.route('/')
def index():
    try:
        print(f"[DEBUG] Template folder: {TEMPLATE_DIR}")
        print(f"[DEBUG] Templates exist: {os.path.exists(TEMPLATE_DIR)}")
        index_path = os.path.join(TEMPLATE_DIR, 'index.html')
        print(f"[DEBUG] index.html exists: {os.path.exists(index_path)}")
        return render_template('index.html', required_files=REQUIRED_FILES)
    except Exception as e:
        return f"Error: {str(e)}<br>Template folder: {TEMPLATE_DIR}<br>Files: {os.listdir(TEMPLATE_DIR) if os.path.exists(TEMPLATE_DIR) else 'FOLDER NOT FOUND'}", 500


# ═══════════════════════════════════════════════
# AUTO-FETCH ENDPOINT — launches background thread, returns job_id immediately
# ═══════════════════════════════════════════════
@app.route('/api/fetch-report', methods=['POST'])
def fetch_report():
    """Starts a background fetch job and returns job_id immediately."""
    try:
        data = request.get_json(force=True, silent=True) or {}
        report_key = data.get('report_key')

        if not report_key or report_key not in FETCHABLE_REPORTS:
            return jsonify({'error': f'Unknown report: {report_key}'}), 400

        report = FETCHABLE_REPORTS[report_key]
        from_date = data.get('from_date')
        to_date = data.get('to_date')

        if report['needs_dates'] and (not from_date or not to_date):
            return jsonify({'error': 'from_date and to_date required'}), 400

        # Start background thread and return job_id immediately
        fj_id = str(uuid.uuid4())[:8]
        fetch_jobs[fj_id] = {'status': 'running', 'filename': None, 'filepath': None,
                             'size_kb': None, 'report_name': report['name'], 'error': None}

        t = threading.Thread(
            target=_run_fetch_bg,
            args=(fj_id, report, report_key, from_date, to_date)
        )
        t.daemon = True
        t.start()

        return jsonify({'job_id': fj_id})

    except Exception as e:
        return jsonify({'error': f'{type(e).__name__}: {str(e)}'}), 500


def _run_fetch_bg(fj_id, report, report_key, from_date, to_date):
    """Runs in background thread: login, download, save file."""
    try:
        print(f"[FETCH-BG] Starting fetch for: {report_key}")
        session = requests.Session()

        # Login
        login_resp = session.post(
            f"{OM_BASE_URL}/api/session",
            json={"username": OM_EMAIL, "password": OM_PASSWORD},
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        print(f"[FETCH-BG] Login status: {login_resp.status_code}")
        if login_resp.status_code not in [200, 202]:
            fetch_jobs[fj_id]['status'] = 'failed'
            fetch_jobs[fj_id]['error'] = f'Login failed: {login_resp.status_code} — {login_resp.text[:200]}'
            return

        # Build parameters
        parameters = []
        for param in report['params']:
            if param['type'] == 'category':
                parameters.append({
                    "type": param['type'],
                    "target": ["variable", ["template-tag", param['name']]],
                    "value": param['value']
                })
            elif param['type'] == 'date/single':
                date_value = from_date if param.get('role') == 'from' else to_date
                parameters.append({
                    "type": param['type'],
                    "target": ["variable", ["template-tag", param['name']]],
                    "value": date_value
                })

        payload = {"parameters": parameters} if parameters else {}

        # Download
        download_url = f"{OM_BASE_URL}/api/card/{report['card_id']}/query/{report['format']}"
        print(f"[FETCH-BG] Downloading: {download_url}")
        file_resp = session.post(
            download_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=300
        )
        print(f"[FETCH-BG] Download status: {file_resp.status_code}")

        if file_resp.status_code != 200:
            fetch_jobs[fj_id]['status'] = 'failed'
            fetch_jobs[fj_id]['error'] = f'Download failed: {file_resp.status_code} — {file_resp.text[:300]}'
            return

        # Save file
        fetch_dir = os.path.join(BASE_DIR, 'uploads', 'fetched')
        os.makedirs(fetch_dir, exist_ok=True)

        filename = (f"{report_key}_{from_date}_to_{to_date}.{report['format']}"
                    if report['needs_dates'] else f"{report_key}.{report['format']}")
        filepath = os.path.join(fetch_dir, filename)
        with open(filepath, 'wb') as f:
            f.write(file_resp.content)

        size_kb = round(os.path.getsize(filepath) / 1024, 1)
        print(f"[FETCH-BG] Saved: {filename} ({size_kb} KB)")

        fetch_jobs[fj_id].update({'status': 'completed', 'filename': filename,
                                   'filepath': filepath, 'size_kb': size_kb})

    except Exception as e:
        tb = traceback.format_exc()
        print(f"[FETCH-BG ERROR] {tb}")
        fetch_jobs[fj_id]['status'] = 'failed'
        fetch_jobs[fj_id]['error'] = f'{type(e).__name__}: {str(e)}'


@app.route('/api/fetch-status/<fj_id>')
def fetch_status(fj_id):
    """Poll this endpoint to check background fetch job result."""
    if fj_id not in fetch_jobs:
        return jsonify({'error': 'Job not found'}), 404
    job = fetch_jobs[fj_id]
    if job['status'] == 'completed':
        return jsonify({
            'status': 'completed',
            'filename': job['filename'],
            'filepath': job['filepath'],
            'size_kb': job['size_kb'],
            'report_name': job['report_name']
        })
    elif job['status'] == 'failed':
        return jsonify({'status': 'failed', 'error': job['error']})
    else:
        return jsonify({'status': 'running'})


@app.route('/api/fetched-file/<path:filename>')
def serve_fetched_file(filename):
    fetch_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'fetched')
    return send_file(os.path.join(fetch_dir, filename))


# ═══════════════════════════════════════════════
# PROCESS ENDPOINT — runs market.py
# ═══════════════════════════════════════════════
@app.route('/api/process', methods=['POST'])
def process_files():
    job_id = str(uuid.uuid4())[:8]
    job_dir = os.path.join(app.config['UPLOAD_FOLDER'], job_id)
    os.makedirs(job_dir, exist_ok=True)

    saved_files = {}
    for finfo in REQUIRED_FILES:
        key = finfo['key']

        # Check if file is in form upload
        if key in request.files and request.files[key].filename != '':
            f = request.files[key]
            fname = secure_filename(f.filename)
            fpath = os.path.join(job_dir, fname)
            f.save(fpath)
            saved_files[key] = fpath
        # Check if filepath was provided (auto-fetched file)
        elif request.form.get(f'{key}_filepath'):
            fetched_path = request.form.get(f'{key}_filepath')
            if os.path.exists(fetched_path):
                saved_files[key] = fetched_path
            else:
                return jsonify({'error': f"Fetched file not found: {finfo['label']}"}), 400
        else:
            return jsonify({'error': f"Missing file: {finfo['label']}"}), 400

    jobs[job_id] = {
        'status': 'queued', 'progress': 0, 'logs': [],
        'output_file': None, 'error': None,
    }

    thread = threading.Thread(target=run_market_py, args=(job_id, saved_files))
    thread.daemon = True
    thread.start()

    return jsonify({'job_id': job_id})


@app.route('/api/status/<job_id>')
def job_status(job_id):
    if job_id not in jobs:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify(jobs[job_id])


@app.route('/api/download/<job_id>')
def download_file(job_id):
    if job_id not in jobs:
        return jsonify({'error': 'Job not found'}), 404
    job = jobs[job_id]
    if job['status'] != 'completed' or not job['output_file']:
        return jsonify({'error': 'File not ready'}), 400
    return send_file(job['output_file'], as_attachment=True,
                     download_name='WORK_ASSIGN_SUMMARY_PROCESSED.xlsx')


# ═══════════════════════════════════════════════
# MARKET.PY EXECUTOR
# ═══════════════════════════════════════════════
def run_market_py(job_id, saved_files):
    try:
        jobs[job_id]['status'] = 'running'
        add_log(job_id, "🚀 Starting market.py processing...", 2)

        market_py_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'market.py')
        if not os.path.exists(market_py_path):
            raise FileNotFoundError("market.py not found! Place it next to app.py")

        with open(market_py_path, 'r', encoding='utf-8') as f:
            source_code = f.read()

        add_log(job_id, "📄 market.py loaded", 5)

        file_queue = [
            saved_files['branch_master'],
            saved_files['work_summary'],
            saved_files['cancel_report'],
            saved_files['challenge_report'],
            saved_files['staff_detail'],
            saved_files['broker_master'],
            saved_files['ho_file'],
        ]

        modified = source_code
        modified = modified.replace("import tkinter as tk", "# [WEB] import tkinter as tk")
        modified = modified.replace("from tkinter import filedialog", "# [WEB] from tkinter import filedialog")
        modified = modified.replace("root = tk.Tk()", "# [WEB] root = tk.Tk()")
        modified = modified.replace("root.withdraw()", "# [WEB] root.withdraw()")

        old_func_start = "def select_file(title, filetypes):"
        old_func_end = "    return filepath"
        start_idx = modified.find(old_func_start)
        if start_idx != -1:
            end_idx = modified.find(old_func_end, start_idx)
            if end_idx != -1:
                end_idx += len(old_func_end)
                new_func = '''def select_file(title, filetypes):
    """[WEB] Returns pre-uploaded file paths"""
    _queue = ''' + repr(file_queue) + '''
    if not hasattr(select_file, '_call_count'):
        select_file._call_count = 0
    idx = select_file._call_count
    select_file._call_count += 1
    if idx < len(_queue):
        return _queue[idx]
    raise Exception(f"No more files for: {title}")'''
                modified = modified[:start_idx] + new_func + modified[end_idx:]

        output_file = os.path.join(app.config['OUTPUT_FOLDER'],
                                   f'{job_id}_WORK_ASSIGN_SUMMARY_PROCESSED.xlsx')
        output_file_safe = output_file.replace('\\', '/')
        modified = modified.replace(
            "output_filename = 'WORK_ASSIGN_SUMMARY_PROCESSED.xlsx'",
            f"output_filename = '{output_file_safe}'"
        )

        add_log(job_id, "🔧 Configured with files:", 8)
        for finfo in REQUIRED_FILES:
            fname = os.path.basename(saved_files[finfo['key']])
            add_log(job_id, f"  {finfo['icon']} {finfo['label']}: {fname}", None)
        add_log(job_id, "\n⚡ Running market.py...\n", 12)

        class LogCapture:
            def __init__(self, jid):
                self.jid = jid
                self.line_buf = ""
            def write(self, text):
                self.line_buf += text
                while '\n' in self.line_buf:
                    line, self.line_buf = self.line_buf.split('\n', 1)
                    if line.strip():
                        add_log(self.jid, line.rstrip(), self._pct(line))
            def flush(self):
                if self.line_buf.strip():
                    add_log(self.jid, self.line_buf.rstrip(), self._pct(self.line_buf))
                    self.line_buf = ""
            def _pct(self, text):
                steps = {
                    'Step 1': 10, 'Step 2.1': 18, 'Step 2.2': 20, 'Step 2.3': 22,
                    'Step 2': 15, 'Step 3': 28, 'Step 4': 32, 'Step 5': 36,
                    'Step 6': 40, 'Step 7': 42, 'Step 8': 44, 'Step 9': 46,
                    'Step 10': 48, 'Step 11': 50, 'Step 12': 52,
                    'Step 13.1': 55, 'Step 13': 54,
                    'Step 14.1': 62, 'Step 14': 60,
                    'Step 15': 65, 'Step 16': 67, 'Step 17': 69,
                    'Step 18': 71, 'Step 19': 73, 'Step 20': 75,
                    'Step 21': 77, 'Step 22': 78, 'Step 23.5': 82, 'Step 23': 80,
                    'Step 24.5': 86, 'Step 24': 84,
                    'Step 25.5': 90, 'Step 25': 88,
                    'Step 26': 92, 'Step 27': 94, 'Step 28': 97,
                    'PROCESS COMPLETED': 100,
                }
                for s in sorted(steps.keys(), key=len, reverse=True):
                    if s in text:
                        return steps[s]
                return None

        old_stdout = sys.stdout
        capture = LogCapture(job_id)
        sys.stdout = capture

        try:
            exec(compile(modified, 'market.py', 'exec'), {'__name__': '__market_exec__'})
        finally:
            capture.flush()
            sys.stdout = old_stdout

        if os.path.exists(output_file):
            jobs[job_id]['status'] = 'completed'
            jobs[job_id]['output_file'] = output_file
            jobs[job_id]['progress'] = 100
            add_log(job_id, "\n🎉 COMPLETE! File ready for download.", 100)
        else:
            raise Exception("Output file not created — check logs above.")

    except Exception as e:
        jobs[job_id]['status'] = 'failed'
        jobs[job_id]['error'] = str(e)
        add_log(job_id, f"\n❌ FAILED: {str(e)}", None)
        for line in traceback.format_exc().split('\n'):
            if line.strip():
                add_log(job_id, line, None)
        sys.stdout = sys.__stdout__


def add_log(job_id, msg, progress=None):
    if job_id in jobs:
        jobs[job_id]['logs'].append(msg)
        if progress is not None:
            jobs[job_id]['progress'] = progress


if __name__ == '__main__':
    mp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'market.py')
    print("=" * 50)
    print("🚀 Market Automation Web Server")
    print("=" * 50)
    print(f"📄 market.py: {'✅ FOUND' if os.path.exists(mp) else '❌ NOT FOUND!'}")
    print(f"📂 Uploads: {app.config['UPLOAD_FOLDER']}")
    print(f"📂 Outputs: {app.config['OUTPUT_FOLDER']}")
    port = int(os.getenv('PORT', 8888))
    print(f"\n🌐 Open http://localhost:{port}\n")
    app.run(host='0.0.0.0', port=port, debug=True)
