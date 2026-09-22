"""
Rider Payout Converter — web app
Upload a raw attendance file, pick company + store, click Generate,
download the finished payout workbook.

Run locally:
    pip install -r requirements.txt
    python app.py
Then open http://localhost:5000 in a browser.

To let others on your network use it, opt in explicitly (the default is
localhost-only, because the app has no login):
    HOST=0.0.0.0 python app.py
"""
import os
import time
import uuid
import zipfile
import traceback
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from payout_engine import COMPANY_CONFIG, build, preview_summary, build_rate_card_rows, \
    build_store_mg_flags, store_has_mg, MissingColumnError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024  # 60MB upload cap (shipment files can be large)

# Only .xlsx is supported: the engine reads workbooks with openpyxl, which can't open the legacy .xls format.
ALLOWED_EXTENSIONS = ('.xlsx',)
XLSX_ONLY_HINT = 'Please upload an Excel .xlsx file (for an older .xls file, open it in Excel and use Save As \u2192 .xlsx).'
RETENTION_SECONDS = 24 * 60 * 60  # uploads/outputs older than this are deleted at startup


def is_xlsx(filename):
    return bool(filename) and filename.lower().endswith(ALLOWED_EXTENSIONS)


def is_configured(company, store):
    """True only for a vendor/store pair that exists in COMPANY_CONFIG (same case-insensitive matching the engine uses)."""
    return (store or '').strip().upper() in COMPANY_CONFIG.get((company or '').strip().upper(), {})


def purge_old_files(max_age_seconds=RETENTION_SECONDS):
    """Delete scratch files (uploaded raw files and generated workbooks) older than max_age_seconds.
    Payout files can contain personal data, so they shouldn't pile up on disk indefinitely."""
    cutoff = time.time() - max_age_seconds
    removed = 0
    for folder in (UPLOAD_DIR, OUTPUT_DIR):
        for entry in os.scandir(folder):
            try:
                if entry.is_file() and entry.stat().st_mtime < cutoff:
                    os.remove(entry.path)
                    removed += 1
            except OSError:
                pass
    return removed


def error_response(e, fallback_message):
    """Structured error payload. For a missing-column problem, surfaces exactly
    which columns were missing and what was actually in the file, so the UI
    can show something actionable instead of a generic failure message."""
    if isinstance(e, MissingColumnError):
        return jsonify({
            'error': f"Missing required column(s): {', '.join(e.missing)}.",
            'missing_columns': e.missing,
            'found_columns': e.found,
        }), 400
    return jsonify({'error': f'{fallback_message}: {e}'}), 500


@app.route('/')
def index():
    # company -> sorted list of store names, for the dropdowns
    companies = {c: sorted(stores.keys()) for c, stores in COMPANY_CONFIG.items()}
    return render_template('index.html', companies=companies, rate_card=build_rate_card_rows(),
                           store_mg_flags=build_store_mg_flags())


def _save_upload(file):
    job_id = uuid.uuid4().hex[:10]
    raw_path = os.path.join(UPLOAD_DIR, f'{job_id}_{secure_filename(file.filename) or "upload.xlsx"}')
    file.save(raw_path)
    return raw_path, job_id


@app.route('/rate-card')
def rate_card_page():
    return render_template('rate_card.html', rate_card=build_rate_card_rows())


@app.route('/preview', methods=['POST'])
def preview():
    try:
        company = request.form.get('company', '').strip()
        store = request.form.get('store', '').strip()
        file = request.files.get('raw_file')

        if not company or not store:
            return jsonify({'error': 'Please select both a vendor and a store.'}), 400
        if not is_configured(company, store):
            return jsonify({'error': 'Unknown vendor/store combination.'}), 400
        if not file or file.filename == '':
            return jsonify({'error': 'Please choose a raw file to upload.'}), 400
        if not is_xlsx(file.filename):
            return jsonify({'error': XLSX_ONLY_HINT}), 400

        raw_path, _ = _save_upload(file)
        summary = preview_summary(raw_path, company=company, store=store)
        return jsonify(summary)

    except Exception as e:
        traceback.print_exc()
        return error_response(e, 'Could not read the raw file')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        company = request.form.get('company', '').strip()
        store = request.form.get('store', '').strip()
        file = request.files.get('raw_file')
        shipment_file = request.files.get('shipment_file')

        if not company or not store:
            return jsonify({'error': 'Please select both a vendor and a store.'}), 400
        if not is_configured(company, store):
            return jsonify({'error': 'Unknown vendor/store combination.'}), 400
        if not file or file.filename == '':
            return jsonify({'error': 'Please choose a raw file to upload.'}), 400
        if not is_xlsx(file.filename):
            return jsonify({'error': XLSX_ONLY_HINT}), 400
        if shipment_file and shipment_file.filename and not is_xlsx(shipment_file.filename):
            return jsonify({'error': 'Shipment file: ' + XLSX_ONLY_HINT}), 400

        raw_path, job_id = _save_upload(file)
        shipment_path = None
        # Shipment cross-check doesn't apply to MG-based stores — payout isn't
        # driven purely by reported delivered orders there. Ignore it even if
        # one slipped through (the UI hides the upload for these stores).
        if shipment_file and shipment_file.filename and not store_has_mg(company, store):
            shipment_path, _ = _save_upload(shipment_file)

        safe_store = store.replace(' ', '_')
        out_name = f'{company}_{safe_store}_FINAL_PAYOUT_{job_id}.xlsx'
        out_path = os.path.join(OUTPUT_DIR, out_name)

        build(raw_path, out_path, company=company, store=store, shipment_path=shipment_path)

        download_name = f'{company}_{safe_store}_FINAL_PAYOUT.xlsx'
        return send_file(out_path, as_attachment=True, download_name=download_name)

    except Exception as e:
        traceback.print_exc()
        return error_response(e, 'Could not generate the payout file')


@app.route('/generate-bulk', methods=['POST'])
def generate_bulk():
    try:
        row_count = int(request.form.get('row_count', '0') or '0')
        if row_count <= 0:
            return jsonify({'error': 'No files to process.'}), 400

        results = []
        # One shared shipment file for the whole batch (same file regardless of
        # vendor/store) — saved once, then reused per row. Skipped automatically
        # for MG-based stores, same as single-file mode.
        shared_shipment_file = request.files.get('shipment_file')
        shared_shipment_path = None
        if shared_shipment_file and shared_shipment_file.filename:
            if not is_xlsx(shared_shipment_file.filename):
                return jsonify({'error': 'Shipment file: ' + XLSX_ONLY_HINT}), 400
            shared_shipment_path, _ = _save_upload(shared_shipment_file)

        for i in range(row_count):
            company = request.form.get(f'company_{i}', '').strip()
            store = request.form.get(f'store_{i}', '').strip()
            file = request.files.get(f'raw_file_{i}')
            label = f'{company} / {store}' if company and store else f'File {i + 1}'

            if not company or not store:
                results.append({'label': label, 'ok': False, 'error': 'Vendor and store are required.'})
                continue
            if not is_configured(company, store):
                results.append({'label': label, 'ok': False, 'error': 'Unknown vendor/store combination.'})
                continue
            if not file or file.filename == '':
                results.append({'label': label, 'ok': False, 'error': 'No file uploaded.'})
                continue
            if not is_xlsx(file.filename):
                results.append({'label': label, 'ok': False, 'error': 'File must be an .xlsx workbook.'})
                continue

            try:
                raw_path, job_id = _save_upload(file)
                row_shipment_path = shared_shipment_path if not store_has_mg(company, store) else None
                safe_store = store.replace(' ', '_')
                out_name = f'{company}_{safe_store}_FINAL_PAYOUT_{job_id}.xlsx'
                out_path = os.path.join(OUTPUT_DIR, out_name)
                build(raw_path, out_path, company=company, store=store, shipment_path=row_shipment_path)
                results.append({
                    'label': label, 'ok': True, 'out_path': out_path,
                    'download_name': f'{company}_{safe_store}_FINAL_PAYOUT.xlsx',
                })
            except MissingColumnError as e:
                results.append({'label': label, 'ok': False,
                                 'error': f"Missing required column(s): {', '.join(e.missing)}."})
            except Exception as e:
                results.append({'label': label, 'ok': False, 'error': str(e)})

        succeeded = [r for r in results if r['ok']]
        failed = [r for r in results if not r['ok']]

        if not succeeded:
            return jsonify({
                'error': 'All files failed to process.',
                'results': [{'label': r['label'], 'error': r['error']} for r in failed],
            }), 400

        zip_name = f'bulk_payouts_{uuid.uuid4().hex[:8]}.zip'
        zip_path = os.path.join(OUTPUT_DIR, zip_name)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            used_names = set()
            for r in succeeded:
                name = r['download_name']
                n, suffix = name, 1
                while n in used_names:
                    base, ext = os.path.splitext(name)
                    n = f'{base}_{suffix}{ext}'
                    suffix += 1
                used_names.add(n)
                zf.write(r['out_path'], arcname=n)
            if failed:
                manifest = 'The following files failed to process and are NOT included above:\n\n' + \
                    '\n'.join(f"- {r['label']}: {r['error']}" for r in failed)
                zf.writestr('FAILURES.txt', manifest)

        summary = f"{len(succeeded)} succeeded" + (f", {len(failed)} failed" if failed else "")
        resp = send_file(zip_path, as_attachment=True, download_name='bulk_payouts.zip')
        resp.headers['X-Bulk-Summary'] = summary
        return resp

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Could not process bulk upload: {e}'}), 500


if __name__ == '__main__':
    host = os.environ.get('HOST', '127.0.0.1')
    port = int(os.environ.get('PORT', '5000'))
    removed = purge_old_files()
    if removed:
        print(f'Removed {removed} old file(s) from uploads/ and outputs/')
    if host not in ('127.0.0.1', 'localhost'):
        print(f'WARNING: listening on {host} - anyone on your network can upload files to this app (no login).')
    app.run(host=host, port=port, debug=False)
