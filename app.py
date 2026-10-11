import os
import sys
import io
import time
import secrets
import threading
import traceback
import tempfile
import webbrowser
from collections import OrderedDict
from werkzeug.utils import secure_filename

from flask import Flask, render_template, request, jsonify, send_file, g

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
    custom_note = data.get('custom_note', '').strip()

    try:
        result = parse_and_process_file(latest_file, filter_beko=filter_beko, custom_note=custom_note)
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
    custom_note = data.get('custom_note', '').strip()

    try:
        result = parse_and_process_multiple_files(recent_files, filter_beko=filter_beko, custom_note=custom_note)
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
    custom_note = request.form.get('custom_note', '').strip()

    try:
        if IS_CLOUD:
            result = process_in_memory(valid_files, filter_beko=filter_beko, custom_note=custom_note)
            result = _attach_downloads(result)
        else:
            result = parse_and_process_multiple_files(valid_files, filter_beko=filter_beko, custom_note=custom_note)
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
    # Guvenlik: yol yalnizca proje klasorleri / Indirilenler icinde olabilir.
    if not folder_path or not os.path.isdir(folder_path) or not _is_allowed_local_path(folder_path):
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



# ---------------------------------------------------------------------------
# TOPLU BARKOD & 10x10cm TERMAL KARGO ETIKET MODULU (PROJE 3 ENTEGRASYONU)
# ---------------------------------------------------------------------------
import glob
from core_barcode import (
    read_excel_or_csv as barcode_read_excel_or_csv,
    extract_items as barcode_extract_items,
    generate_labels_pdf as barcode_generate_labels_pdf,
    create_sample_excel as barcode_create_sample_excel,
    init_fonts as barcode_init_fonts,
)

# Termal fontlari baslat
barcode_init_fonts()

if IS_CLOUD:
    BARCODE_UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'barcode_uploads')
else:
    BARCODE_UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp_barcode_uploads')
try:
    os.makedirs(BARCODE_UPLOAD_FOLDER, exist_ok=True)
except Exception:
    pass

# Kargo dosyasi (basliklar/satirlar/eslesme) tarayici oturumuna gore ayrilir.
# Once tek bir modul seviyesi dict kullaniliyordu; bulutta eszamanli iki kullanicinin
# musteri/adres/telefon verilerini birbirine gormesine yol aciyordu (H-10).
BARCODE_COOKIE_NAME = 'barcode_sid'
BARCODE_CACHE_TTL_SECONDS = 60 * 60
BARCODE_CACHE_MAX = 20
BARCODE_SESSION_CACHE = {}
_barcode_cache_lock = threading.Lock()


def _empty_barcode_cache():
    return {'headers': [], 'data_rows': [], 'mapping': {}, 'items': []}


def _valid_barcode_sid(sid):
    return (
        isinstance(sid, str)
        and 16 <= len(sid) <= 64
        and all(ch.isalnum() or ch in '-_' for ch in sid)
    )


def _barcode_sid():
    """Istek/oturum bazli benzersiz anahtar. Gerekirse yeni bir tane uretir."""
    sid = request.cookies.get(BARCODE_COOKIE_NAME)
    if _valid_barcode_sid(sid):
        g.barcode_sid = sid
        g.barcode_sid_is_new = False
    else:
        g.barcode_sid = secrets.token_urlsafe(16)
        g.barcode_sid_is_new = True
    return g.barcode_sid


@app.after_request
def _persist_barcode_sid(response):
    if getattr(g, 'barcode_sid_is_new', False) and getattr(g, 'barcode_sid', None):
        response.set_cookie(
            BARCODE_COOKIE_NAME,
            g.barcode_sid,
            max_age=BARCODE_CACHE_TTL_SECONDS,
            httponly=True,
            samesite='Lax',
            secure=IS_CLOUD,
            path='/',
        )
    return response


def _barcode_cache_get():
    sid = _barcode_sid()
    with _barcode_cache_lock:
        entry = BARCODE_SESSION_CACHE.get(sid)
        if not entry:
            return _empty_barcode_cache()
        if time.time() - entry['ts'] > BARCODE_CACHE_TTL_SECONDS:
            BARCODE_SESSION_CACHE.pop(sid, None)
            return _empty_barcode_cache()
        return entry['data']


def _barcode_cache_set(headers, data_rows, mapping, items):
    sid = _barcode_sid()
    with _barcode_cache_lock:
        now = time.time()
        for key in [k for k, v in BARCODE_SESSION_CACHE.items() if now - v['ts'] > BARCODE_CACHE_TTL_SECONDS]:
            BARCODE_SESSION_CACHE.pop(key, None)
        while len(BARCODE_SESSION_CACHE) >= BARCODE_CACHE_MAX:
            oldest = min(BARCODE_SESSION_CACHE, key=lambda k: BARCODE_SESSION_CACHE[k]['ts'])
            BARCODE_SESSION_CACHE.pop(oldest, None)
        BARCODE_SESSION_CACHE[sid] = {
            'ts': now,
            'data': {
                'headers': headers,
                'data_rows': data_rows,
                'mapping': mapping,
                'items': items,
            },
        }


@app.route('/barkod')
def barkod_view():
    return render_template('index.html', is_cloud=IS_CLOUD, initial_tab='barkod')


@app.route('/api/barcode/upload', methods=['POST'])
def barcode_upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'Lütfen bir dosya seçin.'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Dosya seçilmedi.'}), 400

    filename = secure_filename(file.filename)
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ['.xlsx', '.xls', '.csv']:
        return jsonify({'success': False, 'message': 'Yalnızca .xlsx, .xls veya .csv dosyaları desteklenir.'}), 400

    file_path = os.path.join(BARCODE_UPLOAD_FOLDER, f"upload_{secrets.token_hex(4)}_{filename}")
    file.save(file_path)

    try:
        headers, data_rows, mapping = barcode_read_excel_or_csv(file_path)
        if not headers or not data_rows:
            return jsonify({'success': False, 'message': 'Dosyada geçerli veri veya başlık satırı bulunamadı.'}), 400

        items = barcode_extract_items(headers, data_rows, mapping)

        _barcode_cache_set(headers, data_rows, mapping, items)

        total_labels = sum(item.get('quantity', 1) for item in items)

        return jsonify({
            'success': True,
            'filename': filename,
            'headers': headers,
            'mapping': mapping,
            'items': items,
            'total_items': len(items),
            'total_labels': total_labels
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'Dosya işlenirken hata oluştu: {str(e)}'}), 500
    finally:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass


@app.route('/api/barcode/remap', methods=['POST'])
def barcode_remap_columns():
    data = request.get_json(silent=True) or {}
    cache = _barcode_cache_get()
    headers = cache['headers']
    data_rows = cache['data_rows']

    if not headers or not data_rows:
        return jsonify({'success': False, 'message': 'Aktif kargo dosyası bulunamadı.'}), 400

    def parse_idx(key):
        val = data.get(key)
        return int(val) if val is not None and str(val).isdigit() else None

    mapping = {
        'cust_idx': parse_idx('cust_idx'),
        'addr_idx': parse_idx('addr_idx'),
        'dist_idx': parse_idx('dist_idx'),
        'city_idx': parse_idx('city_idx'),
        'phone_idx': parse_idx('phone_idx'),
        'barcode_idx': parse_idx('barcode_idx'),
        'name_idx': parse_idx('name_idx'),
        'qty_idx': parse_idx('qty_idx'),
    }

    items = barcode_extract_items(headers, data_rows, mapping)
    _barcode_cache_set(headers, data_rows, mapping, items)

    total_labels = sum(item.get('quantity', 1) for item in items)
    return jsonify({
        'success': True,
        'mapping': mapping,
        'items': items,
        'total_items': len(items),
        'total_labels': total_labels
    })


@app.route('/api/barcode/load-sample', methods=['GET'])
def barcode_load_sample_data():
    sample_bytes = barcode_create_sample_excel()
    sample_path = os.path.join(BARCODE_UPLOAD_FOLDER, f'sample_{secrets.token_hex(4)}.xlsx')
    with open(sample_path, 'wb') as f:
        f.write(sample_bytes)

    try:
        headers, data_rows, mapping = barcode_read_excel_or_csv(sample_path)
        items = barcode_extract_items(headers, data_rows, mapping)

        _barcode_cache_set(headers, data_rows, mapping, items)

        total_labels = sum(item.get('quantity', 1) for item in items)
        return jsonify({
            'success': True,
            'filename': 'Ornek_Kargo_Listesi.xlsx',
            'headers': headers,
            'mapping': mapping,
            'items': items,
            'total_items': len(items),
            'total_labels': total_labels
        })
    finally:
        if os.path.exists(sample_path):
            try:
                os.remove(sample_path)
            except Exception:
                pass


@app.route('/api/barcode/download-sample-excel', methods=['GET'])
def barcode_download_sample_excel():
    buf_bytes = barcode_create_sample_excel()
    return send_file(
        io.BytesIO(buf_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Ornek_Kargo_Listesi.xlsx'
    )


@app.route('/api/barcode/load-downloads', methods=['GET'])
def barcode_load_recent_downloads():
    if IS_CLOUD:
        return jsonify({'success': False, 'message': 'Bulut sürümünde yerel İndirilenler klasörü taranamıyor. Lütfen dosya yükleyin.'}), 501

    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    if not os.path.exists(downloads_path):
        return jsonify({'success': False, 'message': 'İndirilenler klasörü bulunamadı.'}), 404

    patterns = ['*.xlsx', '*.xls', '*.csv']
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(os.path.join(downloads_path, pat)))

    candidates = [f for f in candidates if not os.path.basename(f).startswith('~$')]
    if not candidates:
        return jsonify({'success': False, 'message': 'İndirilenler klasöründe uygun Excel/CSV dosyası bulunamadı.'}), 404

    latest_file = max(candidates, key=os.path.getmtime)
    filename = os.path.basename(latest_file)

    try:
        headers, data_rows, mapping = barcode_read_excel_or_csv(latest_file)
        if not headers or not data_rows:
            return jsonify({'success': False, 'message': f'{filename} dosyasında geçerli veri bulunamadı.'}), 400

        items = barcode_extract_items(headers, data_rows, mapping)
        _barcode_cache_set(headers, data_rows, mapping, items)

        total_labels = sum(item.get('quantity', 1) for item in items)
        return jsonify({
            'success': True,
            'filename': filename,
            'headers': headers,
            'mapping': mapping,
            'items': items,
            'total_items': len(items),
            'total_labels': total_labels
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'Dosya okunurken hata: {str(e)}'}), 500


@app.route('/api/barcode/generate-pdf', methods=['POST'])
def barcode_generate_pdf_endpoint():
    data = request.get_json(silent=True) or {}
    items = data.get('items', [])
    options = data.get('options', {})

    if not items:
        # Yalnizca ayni tarayici oturumunun kargo dosyasindan tamamlanir.
        items = _barcode_cache_get()['items']

    if not items:
        return jsonify({'success': False, 'message': 'PDF üretilecek sipariş listesi bulunamadı.'}), 400

    try:
        pdf_bytes = barcode_generate_labels_pdf(items, options)
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name='Kargo_Barkod_10x10cm.pdf'
        )
    except Exception as e:
        return jsonify({'success': False, 'message': f'PDF oluşturulamadı: {str(e)}'}), 500



# ---------------------------------------------------------------------------
# OLIZ KAMPANYA & INDIRIM ANALIZ MODULU (PROJE 3 ENTEGRASYONU)
# ---------------------------------------------------------------------------
from core_campaign import CampaignEngine

if IS_CLOUD:
    OLIZ_DATA_DIR = os.path.join(BASE_DIR, "data")
    OLIZ_UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "oliz_uploads")
else:
    OLIZ_DATA_DIR = os.path.join(BASE_DIR, "data")
    OLIZ_UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")
    try:
        os.makedirs(OLIZ_DATA_DIR, exist_ok=True)
    except Exception:
        pass

try:
    os.makedirs(OLIZ_UPLOAD_DIR, exist_ok=True)
except Exception:
    pass

default_campaign_file = None
if os.path.exists(OLIZ_DATA_DIR):
    candidate = os.path.join(OLIZ_DATA_DIR, "1-15_EKIM.xlsx")
    if os.path.exists(candidate):
        default_campaign_file = candidate
    else:
        try:
            c_files = [os.path.join(OLIZ_DATA_DIR, f) for f in os.listdir(OLIZ_DATA_DIR) if f.endswith(".xlsx")]
            default_campaign_file = c_files[0] if c_files else None
        except Exception:
            default_campaign_file = None

try:
    campaign_engine = CampaignEngine(default_campaign_file)
except Exception as e:
    print(f"Oliz CampaignEngine init fallback: {e}")
    campaign_engine = CampaignEngine(None)

# ---------------------------------------------------------------------------
# BARKOD & SERI NO ARSIVLEME MODULU (BARKOD_ARSIV / PROJE 4 ENTEGRASYONU)
# CANLI BULUT: https://barkod-eight.vercel.app/
# ---------------------------------------------------------------------------
from core_arsiv_db import (
    init_db as init_arsiv_db,
    add_barcode_record as add_arsiv_record,
    check_duplicate as check_arsiv_duplicate,
    search_records as search_arsiv_records,
    get_stats as get_arsiv_stats,
    get_record_by_id as get_arsiv_record_by_id,
    export_records_to_excel as export_arsiv_excel,
    save_active_orders_file as save_arsiv_active_orders,
    get_active_orders_file as get_arsiv_active_orders,
    clear_active_orders_file as clear_arsiv_active_orders,
    delete_record_by_id as delete_arsiv_record_by_id,
    delete_record_by_barcode as delete_arsiv_record_by_barcode,
    update_record_by_id as update_arsiv_record_by_id,
    update_record_by_barcode as update_arsiv_record_by_barcode,
)
from core_arsiv_parser import parse_order_file as parse_arsiv_order_file
from core_arsiv_cloud import (
    save_orders_to_cloud as save_arsiv_cloud,
    get_orders_from_cloud as get_arsiv_cloud,
    clear_orders_from_cloud as clear_arsiv_cloud,
    update_order_progress_cloud as update_arsiv_cloud,
    remove_serial_from_cloud as remove_arsiv_serial_cloud,
    update_serial_in_cloud as update_arsiv_serial_cloud,
    sync_db_with_cloud_archive as sync_arsiv_cloud_db,
    restore_scanned_status_from_archive as restore_arsiv_scanned_status,
)

try:
    init_arsiv_db()
    # Arka planda kalıcı bulut arşivini SQLite ile çift yönlü eşitle
    threading.Thread(target=sync_arsiv_cloud_db, daemon=True).start()
except Exception as e:
    print(f"[UYARI] Arsiv DB baslatma / bulut senkron hatasi: {e}", file=sys.stderr)

@app.context_processor
def inject_global_template_context():
    try:
        st = campaign_engine.get_stats()
    except Exception:
        st = {}
    try:
        ast = get_arsiv_stats()
    except Exception:
        ast = {}
    return dict(stats=st, arsiv_stats=ast)


@app.route('/oliz')
def oliz_view():
    return render_template('index.html', is_cloud=IS_CLOUD, initial_tab='oliz')


@app.route('/api/oliz/stats', methods=['GET'])
@app.route('/api/stats', methods=['GET'])
def oliz_stats_endpoint():
    return jsonify({
        'success': True,
        'stats': campaign_engine.get_stats()
    })


@app.route('/api/oliz/autocomplete', methods=['GET'])
@app.route('/api/autocomplete', methods=['GET'])
def oliz_autocomplete_endpoint():
    q = request.args.get('q', '').strip()
    try:
        limit = int(request.args.get('limit', 15))
    except (TypeError, ValueError):
        limit = 15
    limit = max(1, min(limit, 100))
    results = campaign_engine.search_products(q, limit=limit)
    return jsonify({
        'success': True,
        'results': results
    })


@app.route('/api/oliz/analyze', methods=['POST'])
@app.route('/api/analyze', methods=['POST'])
def oliz_analyze_endpoint():
    data = request.get_json(silent=True) or {}
    products = data.get('products', [])
    if not isinstance(products, list):
        products = [products]
    clean_products = [str(p).strip() for p in products if p and str(p).strip()][:4]
    analysis = campaign_engine.analyze_bundle(clean_products)
    return jsonify({
        'success': True,
        'data': analysis
    })


@app.route('/api/oliz/single-search', methods=['GET'])
@app.route('/api/single-search', methods=['GET'])
def oliz_single_search_endpoint():
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify({'success': False, 'message': 'Arama sorgusu boş.'})
    res = campaign_engine.analyze_single_product(q)
    if not res:
        return jsonify({'success': False, 'message': 'Ürün bulunamadı veya kampanya listesinde yer almıyor.'})
    return jsonify({
        'success': True,
        'data': res
    })


@app.route('/api/oliz/catalog', methods=['GET'])
@app.route('/api/catalog', methods=['GET'])
def oliz_catalog_endpoint():
    cat_type = request.args.get('type', 'tekil').lower()
    if cat_type == 'tekil':
        items = campaign_engine.get_all_tekil()
    elif cat_type == 'toptan':
        items = campaign_engine.get_all_toptan()
    elif cat_type == 'paket':
        items = campaign_engine.get_all_paket()
    else:
        items = campaign_engine.get_all_tekil()
    return jsonify({
        'success': True,
        'type': cat_type,
        'count': len(items),
        'items': items
    })


@app.route('/api/oliz/upload', methods=['POST'])
def oliz_upload_endpoint():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'Dosya yüklenmedi.'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'Seçilen dosya boş.'}), 400
    if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        return jsonify({'success': False, 'message': 'Lütfen geçerli bir Excel (.xlsx / .xls) dosyası yükleyin.'}), 400

    filename = secure_filename(file.filename) or 'kampanya.xlsx'
    # Ayni adi tasiyan iki yukleme birbirini ezmesin diye benzersiz onek (H-02).
    dest_path = os.path.join(OLIZ_UPLOAD_DIR, f"kampanya_{secrets.token_hex(4)}_{filename}")
    file.save(dest_path)
    try:
        campaign_engine.load_from_excel(dest_path)
        return jsonify({
            'success': True,
            'message': f"'{filename}' başarıyla yüklendi ve işlendi!",
            'stats': campaign_engine.get_stats()
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Dosya işlenirken hata oluştu: {str(e)}'
        }), 500


# ---------------------------------------------------------------------------
# BARKOD & SERİ NO ARŞİVLEME SİSTEMİ ROTALARI (BARKOD_ARSIV / BULUT ENTEGRASYON)
# CANLI YAYIN: https://barkod-eight.vercel.app/
# ---------------------------------------------------------------------------

@app.route('/arsiv')
@app.route('/barkod_arsiv')
def arsiv_view():
    return render_template('index.html', is_cloud=IS_CLOUD, initial_tab='arsiv')


def get_arsiv_temp_upload_dir():
    if IS_CLOUD:
        temp_dir = os.path.join(tempfile.gettempdir(), 'temp_uploads')
    else:
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'temp_uploads')
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir


@app.route('/api/arsiv/load-recent-download', methods=['POST'])
@app.route('/api/load-recent-download', methods=['POST'])
def arsiv_load_recent_download():
    """Yerelde İndirilenler klasöründeki en güncel pazaryeri veya e-ticaret web sipariş dosyasını otomatik yükler."""
    downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    if not os.path.exists(downloads_dir):
        return jsonify({'status': 'error', 'message': 'İndirilenler klasörü bulunamadı'}), 404

    recent_candidates = []
    for f in os.listdir(downloads_dir):
        if f.startswith('~') or f.startswith('.'):
            continue
        fl = f.lower()
        if not fl.endswith(('.xlsx', '.xls', '.csv')):
            continue
        if any(fl.startswith(prefix) for prefix in ['barkod_serino_kayitlari_', 'urun_toplama_listesi_', 'beko_urun_toplama_']):
            continue
        fp = os.path.join(downloads_dir, f)
        try:
            mtime = os.path.getmtime(fp)
            recent_candidates.append((mtime, fp, f))
        except Exception:
            pass

    if not recent_candidates:
        return jsonify({'status': 'error', 'message': 'İndirilenler klasöründe sipariş dosyası bulunamadı'}), 404

    recent_candidates.sort(key=lambda x: x[0], reverse=True)
    selected_name = None
    orders = []

    for _, fp, fname in recent_candidates[:10]:
        try:
            parsed = parse_arsiv_order_file(fp)
            if parsed and len(parsed) > 0:
                selected_name = fname
                orders = parsed
                break
        except Exception:
            continue

    if not selected_name or not orders:
        return jsonify({'status': 'error', 'message': 'İndirilenler klasöründe geçerli sipariş satırı içeren dosya bulunamadı'}), 404

    try:
        orders = restore_arsiv_scanned_status(orders)
        save_arsiv_active_orders(selected_name, orders)
        save_arsiv_cloud(selected_name, orders)
    except Exception as e:
        print(f"[ArsivRestore/CloudSync] Hata: {e}")

    return jsonify({
        'status': 'success',
        'filename': selected_name,
        'order_count': len(orders),
        'orders': orders
    })


@app.route('/api/arsiv/upload-orders', methods=['POST'])
@app.route('/api/upload-orders', methods=['POST'])
def arsiv_upload_orders():
    raw_files = request.files.getlist('files')
    if not raw_files:
        raw_files = request.files.getlist('file')
    valid_files = [f for f in raw_files if f and getattr(f, 'filename', '').strip() != '']

    if not valid_files:
        return jsonify({'status': 'error', 'message': 'Hiçbir dosya seçilmedi veya dosya adı geçersiz'}), 400

    upload_dir = get_arsiv_temp_upload_dir()
    saved_paths = []
    file_names = []

    try:
        all_packages_map = {}
        for f in valid_files:
            fname = f.filename
            file_names.append(fname)
            unique_name = f"{secrets.token_hex(4)}_{fname}"
            temp_path = os.path.join(upload_dir, unique_name)
            f.save(temp_path)
            saved_paths.append(temp_path)

            file_orders = parse_arsiv_order_file(temp_path)
            for pkg in file_orders:
                key = pkg.get('kargo_barkodu') or pkg.get('siparis_no') or pkg.get('paket_no')
                if not key:
                    key = f"pkg_{len(all_packages_map)+1}"

                if key not in all_packages_map:
                    all_packages_map[key] = pkg
                else:
                    existing = all_packages_map[key]
                    existing_items = existing.setdefault('urunler', [])
                    for item in pkg.get('urunler', []):
                        item_copy = dict(item)
                        item_copy['item_id'] = f"item_{len(existing_items) + 1}"
                        existing_items.append(item_copy)

        merged_orders = list(all_packages_map.values())
        if not merged_orders:
            return jsonify({'status': 'error', 'message': 'Seçilen dosyalarda uygun sipariş satırı tespit edilemedi'}), 400

        if len(file_names) == 1:
            summary_name = file_names[0]
        else:
            names_summary = ", ".join(file_names[:2])
            if len(file_names) > 2:
                names_summary += f" (+{len(file_names)-2} dosya)"
            summary_name = f"{len(file_names)} Dosya ({names_summary})"

        try:
            merged_orders = restore_arsiv_scanned_status(merged_orders)
            save_arsiv_active_orders(summary_name, merged_orders)
            save_arsiv_cloud(summary_name, merged_orders)
        except Exception as e:
            print(f"[CloudSync] Hata: {e}")

        return jsonify({
            'status': 'success',
            'filename': summary_name,
            'file_count': len(file_names),
            'order_count': len(merged_orders),
            'orders': merged_orders
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Dosyalar işlenirken hata: {str(e)}'}), 500
    finally:
        for p in saved_paths:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


@app.route('/api/arsiv/record', methods=['POST'])
@app.route('/api/record', methods=['POST'])
def arsiv_record():
    data = request.get_json() or {}
    barkod_seri_no = data.get('barkod_seri_no', '').strip()
    if not barkod_seri_no:
        return jsonify({'status': 'error', 'message': 'Barkod / Seri No boş olamaz'}), 400

    kargo_barkodu = data.get('kargo_barkodu', '').strip()
    siparis_no = data.get('siparis_no', '').strip()
    musteri_adi = data.get('musteri_adi', '').strip()
    urun_adi = data.get('urun_adi', '').strip()
    pazaryeri = data.get('pazaryeri', 'Bilinmiyor').strip()
    okutma_yontemi = data.get('okutma_yontemi', 'El Terminali')
    notlar = data.get('notlar', '').strip()

    try:
        update_arsiv_cloud(kargo_barkodu, barkod_seri_no)
    except Exception:
        pass

    record_data, is_dup, dup_info = add_arsiv_record(
        kargo_barkodu=kargo_barkodu,
        siparis_no=siparis_no,
        musteri_adi=musteri_adi,
        urun_adi=urun_adi,
        barkod_seri_no=barkod_seri_no,
        pazaryeri=pazaryeri,
        okutma_yontemi=okutma_yontemi,
        notlar=notlar
    )

    stats = get_arsiv_stats()
    return jsonify({
        'status': 'success',
        'record': record_data,
        'is_duplicate': is_dup,
        'duplicate_info': dup_info,
        'stats': {
            'toplam_kayit': stats['toplam_kayit'],
            'bugun_kayit': stats['bugun_kayit']
        }
    })


@app.route('/api/arsiv/check-duplicate', methods=['GET'])
@app.route('/api/check-duplicate', methods=['GET'])
def arsiv_check_duplicate():
    barcode = request.args.get('barcode', '').strip()
    is_dup, dup_info = check_arsiv_duplicate(barcode)
    return jsonify({
        'is_duplicate': is_dup,
        'duplicate_info': dup_info
    })


@app.route('/api/arsiv/search', methods=['GET'])
@app.route('/api/search', methods=['GET'])
def arsiv_search():
    query = request.args.get('q', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    records = search_arsiv_records(query=query, start_date=start_date, end_date=end_date, limit=200)
    return jsonify({
        'status': 'success',
        'count': len(records),
        'records': records
    })


@app.route('/api/arsiv/stats', methods=['GET'])
def arsiv_stats_endpoint():
    return jsonify(get_arsiv_stats())


@app.route('/api/arsiv/export-excel', methods=['GET'])
@app.route('/api/export-excel', methods=['GET'])
def arsiv_export_excel():
    query = request.args.get('q', '').strip()
    start_date = request.args.get('start_date', '').strip()
    end_date = request.args.get('end_date', '').strip()
    records = search_arsiv_records(query=query, start_date=start_date, end_date=end_date, limit=5000)
    wb = export_arsiv_excel(records)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    filename = f"Barkod_SeriNo_Kayitlari_{get_arsiv_stats()['bugun_tarih']}.xlsx"
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/arsiv/proof/<int:record_id>', methods=['GET'])
@app.route('/api/proof/<int:record_id>', methods=['GET'])
def arsiv_get_proof(record_id):
    record = get_arsiv_record_by_id(record_id)
    if not record:
        return jsonify({'status': 'error', 'message': 'Kayıt bulunamadı'}), 404
    return jsonify({'status': 'success', 'record': record})


@app.route('/api/arsiv/delete-record/<int:record_id>', methods=['POST', 'DELETE'])
@app.route('/api/delete-record/<int:record_id>', methods=['POST', 'DELETE'])
def arsiv_delete_record(record_id):
    success, msg = delete_arsiv_record_by_id(record_id)
    if not success:
        return jsonify({'status': 'error', 'message': msg}), 400
    stats = get_arsiv_stats()
    return jsonify({
        'status': 'success',
        'message': 'Kayıt başarıyla silindi',
        'stats': {
            'toplam_kayit': stats['toplam_kayit'],
            'bugun_kayit': stats['bugun_kayit']
        }
    })


@app.route('/api/arsiv/delete-record-by-barcode', methods=['POST'])
@app.route('/api/delete-record-by-barcode', methods=['POST'])
@app.route('/api/arsiv/remove-barcode', methods=['POST'])
@app.route('/api/remove-barcode', methods=['POST'])
def arsiv_remove_barcode():
    data = request.get_json() or {}
    barkod_seri_no = data.get('barkod_seri_no', '').strip()
    kargo_barkodu = data.get('kargo_barkodu', '').strip()
    siparis_no = data.get('siparis_no', '').strip()

    if not barkod_seri_no:
        return jsonify({'status': 'error', 'message': 'Barkod / Seri No boş olamaz'}), 400

    try:
        remove_arsiv_serial_cloud(kargo_barkodu, barkod_seri_no)
    except Exception:
        pass

    success, msg = delete_arsiv_record_by_barcode(barkod_seri_no, kargo_barkodu=kargo_barkodu, siparis_no=siparis_no)
    stats = get_arsiv_stats()
    return jsonify({
        'status': 'success',
        'message': msg,
        'stats': {
            'toplam_kayit': stats['toplam_kayit'],
            'bugun_kayit': stats['bugun_kayit']
        }
    })


@app.route('/api/arsiv/update-record/<int:record_id>', methods=['POST', 'PUT'])
@app.route('/api/update-record/<int:record_id>', methods=['POST', 'PUT'])
def arsiv_update_record(record_id):
    data = request.get_json() or {}
    success, msg = update_arsiv_record_by_id(record_id, data)
    if not success:
        return jsonify({'status': 'error', 'message': msg}), 400
    rec = get_arsiv_record_by_id(record_id)
    return jsonify({
        'status': 'success',
        'message': 'Kayıt güncellendi',
        'record': rec
    })


@app.route('/api/arsiv/update-barcode', methods=['POST'])
@app.route('/api/update-barcode', methods=['POST'])
def arsiv_update_barcode():
    data = request.get_json() or {}
    old_sn = data.get('old_barkod_seri_no', '').strip()
    new_sn = data.get('new_barkod_seri_no', '').strip()
    kargo_barkodu = data.get('kargo_barkodu', '').strip()
    siparis_no = data.get('siparis_no', '').strip()

    if not old_sn or not new_sn:
        return jsonify({'status': 'error', 'message': 'Eski ve yeni barkod belirtilmelidir'}), 400

    try:
        update_arsiv_serial_cloud(kargo_barkodu, old_sn, new_sn)
    except Exception:
        pass

    success, msg = update_arsiv_record_by_barcode(old_sn, new_sn, kargo_barkodu=kargo_barkodu, siparis_no=siparis_no)
    return jsonify({'status': 'success', 'message': msg})


@app.route('/api/arsiv/get-current-orders', methods=['GET'])
@app.route('/api/get-current-orders', methods=['GET'])
def arsiv_get_current_orders():
    data = get_arsiv_cloud() or get_arsiv_active_orders()
    if data and data.get('orders'):
        return jsonify({
            'status': 'success',
            'filename': data.get('filename', 'Aktif Siparisler'),
            'updated_at': data.get('updated_at', ''),
            'order_count': len(data.get('orders', [])),
            'orders': data.get('orders', [])
        })
    return jsonify({'status': 'empty', 'message': 'Kayitli aktif siparis bulunmuyor', 'orders': []})


@app.route('/api/arsiv/clear-current-orders', methods=['POST'])
@app.route('/api/clear-current-orders', methods=['POST'])
def arsiv_clear_current_orders():
    clear_arsiv_active_orders()
    try:
        clear_arsiv_cloud()
    except Exception:
        pass
    return jsonify({'status': 'success', 'message': 'Aktif siparis listesi temizlendi'})


@app.route('/api/arsiv/sync-cloud', methods=['POST'])
@app.route('/api/sync-cloud', methods=['POST'])
def arsiv_sync_cloud():
    """Tüm cihazlarda arşiv kayıtlarını ve aktif siparişleri anlık olarak GitHub bulutla eşitler."""
    try:
        res = sync_arsiv_cloud_db()
        cloud_orders = get_arsiv_cloud(force_refresh=True)
        stats = get_arsiv_stats()
        return jsonify({
            'status': 'success',
            'message': 'Bulut arşiv senkronizasyonu tamamlandı',
            'sync_result': res,
            'active_orders': cloud_orders,
            'stats': stats
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Bulut senkronizasyon hatası: {str(e)}'}), 500


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
