import os
import sys
import io
import time
import secrets
import threading
import traceback
import webbrowser
from collections import OrderedDict

from flask import Flask, render_template, request, jsonify, send_file

from core_engine import (
    IS_CLOUD,
    parse_and_process_file,
    parse_and_process_multiple_files,
    process_in_memory,
    get_latest_download_file,
    get_latest_download_files,
    get_base_dirs,
)

is_frozen = getattr(sys, 'frozen', False)
template_folder = os.path.join(sys._MEIPASS, 'templates') if is_frozen else 'templates'
static_folder = os.path.join(sys._MEIPASS, 'static') if is_frozen else 'static'
BASE_DIR = os.path.dirname(sys.executable) if is_frozen else os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)

app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024  # 64 MB - siparis dosyasi yukleme siniri

# ---------------------------------------------------------------------------
# Bulutta indirilebilir dosyalar icin bellek deposu.
# Sunucuda kalici disk yok; Excel byte'lari kisa sureligine tutulur.
# ---------------------------------------------------------------------------
DOWNLOAD_STORE = OrderedDict()
DOWNLOAD_TTL_SECONDS = 30 * 60
DOWNLOAD_STORE_MAX = 20
_store_lock = threading.Lock()


def _prune_store(now):
    # Degerler uc elemanlidir: (veri, dosya_adi, zaman)
    for key in [k for k, (_, _, ts) in DOWNLOAD_STORE.items() if now - ts > DOWNLOAD_TTL_SECONDS]:
        DOWNLOAD_STORE.pop(key, None)
    while len(DOWNLOAD_STORE) > DOWNLOAD_STORE_MAX:
        DOWNLOAD_STORE.popitem(last=False)


def store_download(data, filename):
    if not data:
        return None
    now = time.time()
    with _store_lock:
        _prune_store(now)
        token = secrets.token_urlsafe(24)
        DOWNLOAD_STORE[token] = (data, filename, now)
        return token


def pop_download(token):
    """Tek seferlik indirme: token kullanilir ve bellekten silinir."""
    if not token:
        return None
    with _store_lock:
        _prune_store(time.time())
        entry = DOWNLOAD_STORE.pop(token, None)
    return entry


@app.route('/')
def index():
    return render_template('index.html', is_cloud=IS_CLOUD)


@app.route('/api/status', methods=['GET'])
def get_status():
    if IS_CLOUD:
        # Bulutta kullanici indirme klasoru okunamaz; yalnizca arayuz bilgisi doner.
        return jsonify({
            'status': 'online',
            'is_cloud': True,
            'latest_download': None,
            'recent_downloads': [],
        })

    latest_file = get_latest_download_file()
    latest_info = None
    if latest_file and os.path.exists(latest_file):
        latest_info = {
            'path': latest_file,
            'name': os.path.basename(latest_file),
            'size': f"{os.path.getsize(latest_file) / 1024:.1f} KB",
            'mtime': os.path.getmtime(latest_file),
        }

    recent_files = get_latest_download_files(max_files=5)
    recent_info_list = []
    for rf in recent_files:
        if os.path.exists(rf):
            recent_info_list.append({
                'path': rf,
                'name': os.path.basename(rf),
                'size': f"{os.path.getsize(rf) / 1024:.1f} KB",
                'mtime': os.path.getmtime(rf),
            })

    return jsonify({
        'status': 'online',
        'is_cloud': False,
        'latest_download': latest_info,
        'recent_downloads': recent_info_list,
        'app_dir': BASE_DIR,
    })


def _attach_downloads(result):
    """Bulutta uretilen Excel byte'larini indirme deposuna koyar, JSON'dan cikarir."""
    main_bytes = result.pop('_main_bytes', None)
    beko_bytes = result.pop('_beko_bytes', None)
    result['main_download_id'] = store_download(main_bytes, 'Urun_Toplama_Listesi_A4_Cikti.xlsx')
    result['beko_download_id'] = store_download(beko_bytes, 'Beko_Urun_Toplama_Listesi_A4.xlsx')
    return result


@app.route('/api/process-latest', methods=['POST'])
def process_latest():
    if IS_CLOUD:
        return jsonify({
            'success': False,
            'error': 'Bulut surumunde sunucu yerel indirme klasurunu goremez. Excel dosyanizi yukleyin.'
        }), 501

    latest_file = get_latest_download_file()
    if not latest_file:
        return jsonify({'success': False, 'error': 'İndirilenler klasöründe uygun sipariş Excel dosyası bulunamadı.'}), 404

    data = request.get_json(silent=True) or {}
    filter_beko = data.get('filter_beko', False)

    try:
        result = parse_and_process_file(latest_file, filter_beko=filter_beko)
        return jsonify({'success': True, 'result': result})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/process-all-recent', methods=['POST'])
def process_all_recent():
    if IS_CLOUD:
        return jsonify({
            'success': False,
            'error': 'Bulut surumunde sunucu yerel indirme klasorunu goremez. Excel dosyalarinizi yukleyin.'
        }), 501

    recent_files = get_latest_download_files(max_files=10)
    if not recent_files:
        return jsonify({'success': False, 'error': 'İndirilenler klasöründe uygun sipariş Excel dosyası bulunamadı.'}), 404

    data = request.get_json(silent=True) or {}
    filter_beko = data.get('filter_beko', False)

    try:
        result = parse_and_process_multiple_files(recent_files, filter_beko=filter_beko)
        return jsonify({'success': True, 'result': result})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/upload', methods=['POST'])
def upload_file():
    uploaded_files = request.files.getlist('files') or request.files.getlist('file')
    valid_files = [f for f in uploaded_files if f and f.filename.strip() != '']
    if not valid_files:
        return jsonify({'success': False, 'error': 'Dosya seçilmedi veya geçerli dosya bulunamadı.'}), 400

    filter_beko = (request.form.get('filter_beko', 'false').lower() == 'true')

    try:
        if IS_CLOUD:
            result = process_in_memory(valid_files, filter_beko=filter_beko)
            result = _attach_downloads(result)
        else:
            result = parse_and_process_multiple_files(valid_files, filter_beko=filter_beko)
        return jsonify({'success': True, 'result': result})
    except Exception as e:
        app.logger.exception('Yukleme hatasi')
        print(f"[HATA] /api/upload: {e}", file=sys.stderr)
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    if IS_CLOUD:
        return jsonify({
            'success': False,
            'error': 'Bulut surumunde sunucuda klasor acilamaz. Indirdiginiz Excel dosyasini bilgisayarinizda acin.'
        }), 501

    data = request.get_json(silent=True) or {}
    folder_path = data.get('folder_path')
    if not folder_path or not os.path.exists(folder_path):
        folder_path = BASE_DIR
    try:
        os.startfile(folder_path)
        return jsonify({'success': True, 'message': 'Klasör açıldı.'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/open-file', methods=['POST'])
def open_file():
    if IS_CLOUD:
        return jsonify({
            'success': False,
            'error': 'Bulut surumunde dosya sunucuda acilamaz. Once indirin, sonra bilgisayarinizda acin.'
        }), 501

    data = request.get_json(silent=True) or {}
    file_path = data.get('file_path')
    if not file_path or not os.path.exists(file_path) or not _is_allowed_local_path(file_path):
        return jsonify({'success': False, 'error': 'Dosya bulunamadı.'}), 404
    try:
        os.startfile(file_path)
        return jsonify({'success': True, 'message': 'Excel dosyası açıldı.'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


def _is_allowed_local_path(path):
    """Yerel dosya erisimi yalnizca proje klasorlerine sinirlidir."""
    try:
        real = os.path.realpath(path)
    except Exception:
        return False
    allowed = [os.path.realpath(b) for b in get_base_dirs()]
    return any(real.startswith(a + os.sep) or real == a for a in allowed)


@app.route('/api/download-file', methods=['GET'])
def download_file():
    if IS_CLOUD:
        entry = pop_download(request.args.get('id'))
        if not entry:
            return jsonify({'success': False, 'error': 'Dosya bulunamadi veya suresi doldu.'}), 404
        data, filename, _ = entry
        return send_file(io.BytesIO(data), as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    path = request.args.get('path')
    if not path or not os.path.exists(path) or not _is_allowed_local_path(path):
        return "Dosya bulunamadı", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


def open_browser():
    webbrowser.open_new('http://127.0.0.1:5000')


if __name__ == '__main__':
    threading.Timer(0.8, open_browser).start()
    print("=" * 66)
    print("  DEPO PANELİ | SİPARİŞ & ÜRÜN TOPLAMA SİSTEMİ")
    print(f"  Mod    : {'BULUT' if IS_CLOUD else 'MASAÜSTÜ'}")
    print("  Adres  : http://127.0.0.1:5000")
    print("=" * 66)
    app.run(host='127.0.0.1', port=5000, debug=False)
