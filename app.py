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
    return render_template('index.html', is_cloud=IS_CLOUD, stats=campaign_engine.get_stats())


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



# ---------------------------------------------------------------------------
# TOPLU BARKOD & 10x10cm TERMAL KARGO ETIKET MODULU (PROJE 3 ENTEGRASYONU)
# ---------------------------------------------------------------------------
import glob
from werkzeug.utils import secure_filename
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
    BARCODE_UPLOAD_FOLDER = '/tmp'
else:
    BARCODE_UPLOAD_FOLDER = os.path.join(BASE_DIR, 'temp_barcode_uploads')
os.makedirs(BARCODE_UPLOAD_FOLDER, exist_ok=True)

BARCODE_SESSION_CACHE = {
    'headers': [],
    'data_rows': [],
    'mapping': {},
    'items': []
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

        BARCODE_SESSION_CACHE['headers'] = headers
        BARCODE_SESSION_CACHE['data_rows'] = data_rows
        BARCODE_SESSION_CACHE['mapping'] = mapping
        BARCODE_SESSION_CACHE['items'] = items

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
    headers = BARCODE_SESSION_CACHE.get('headers', [])
    data_rows = BARCODE_SESSION_CACHE.get('data_rows', [])

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
    BARCODE_SESSION_CACHE['mapping'] = mapping
    BARCODE_SESSION_CACHE['items'] = items

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

        BARCODE_SESSION_CACHE['headers'] = headers
        BARCODE_SESSION_CACHE['data_rows'] = data_rows
        BARCODE_SESSION_CACHE['mapping'] = mapping
        BARCODE_SESSION_CACHE['items'] = items

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
        BARCODE_SESSION_CACHE['headers'] = headers
        BARCODE_SESSION_CACHE['data_rows'] = data_rows
        BARCODE_SESSION_CACHE['mapping'] = mapping
        BARCODE_SESSION_CACHE['items'] = items

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
        items = BARCODE_SESSION_CACHE.get('items', [])

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

OLIZ_DATA_DIR = os.path.join(BASE_DIR, "data")
OLIZ_UPLOAD_DIR = os.path.join(BASE_DIR, "temp_uploads")
os.makedirs(OLIZ_DATA_DIR, exist_ok=True)
os.makedirs(OLIZ_UPLOAD_DIR, exist_ok=True)

default_campaign_file = os.path.join(OLIZ_DATA_DIR, "1-15_EKIM.xlsx")
if not os.path.exists(default_campaign_file):
    c_files = [os.path.join(OLIZ_DATA_DIR, f) for f in os.listdir(OLIZ_DATA_DIR) if f.endswith(".xlsx")]
    default_campaign_file = c_files[0] if c_files else None

campaign_engine = CampaignEngine(default_campaign_file)

@app.context_processor
def inject_global_template_context():
    try:
        st = campaign_engine.get_stats()
    except Exception:
        st = {}
    return dict(stats=st)



@app.route('/oliz')
def oliz_view():
    stats = campaign_engine.get_stats()
    return render_template('index.html', is_cloud=IS_CLOUD, initial_tab='oliz', stats=stats)


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
    limit = int(request.args.get('limit', 15))
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
    dest_path = os.path.join(OLIZ_DATA_DIR, filename)
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
