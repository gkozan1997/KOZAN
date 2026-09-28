import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break
from collections import defaultdict
import datetime
import io
import os
import glob
import re
import sys
import tempfile
import unicodedata

# Bulut (Vercel) ortamı: kalıcı disk ve Windows Gezgini yok.
# Yerel masaüstü sürümü bu bayrağı False görür ve eski davranışını sürdürür.
IS_CLOUD = bool(os.environ.get('VERCEL'))

# Geçici dosya dizini. Vercel'de TMPDIR=/tmp verir; yerelde işletim sisteminin
# geçici dizinine düşer (sabit /tmp yolu Windows'ta C:\tmp yaratırdı).
CLOUD_TMP_DIR = os.environ.get('TMPDIR') or tempfile.gettempdir()

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass


def safe_filename(name, fallback='uploaded_orders.xlsx'):
    """
    Yüklenen dosya adını güvenli hale getirir.
    Path traversal (..\\..\\etc\\passwd) ve ayraç temizleme.
    """
    base = unicodedata.normalize('NFC', str(name or '')).strip()
    base = base.replace('\\', '/').split('/')[-1]          # yol bileşenlerini at
    base = re.sub(r'[\x00-\x1f<>:"|?*]', '', base)        # yasaklı karakterler
    base = base.strip(' .')
    if not base:
        return fallback
    if not base.lower().endswith(('.xls', '.xlsx')):
        base = fallback
    return base[:120]

KNOWN_BRANDS = [
    'PHILIPS', 'TEFAL', 'BABYLISS', 'BRAUN', 'WMF', 'KENWOOD', 
    'LENOVO', 'BEKO', 'GRUNDIG', 'BISSELL', 'ARIETE', 'LAURASTAR', 
    'TEKA', 'IPHONE', 'APPLE', 'FAKIR', 'ARZUM', 'KARACA', 
    'KORKMAZ', 'DYSON', 'SAMSUNG', 'XIAOMI', 'BOSCH', 'SIEMENS',
    'HOTPOINT', 'ROWENTA'
]

BEKO_MODELS = [
    'KMX', '7053', 'CM', 'CMX', 'B 600', 'B 710', 'BKK', 'BK ', 'BFC', 'BDE', 
    'TKM', 'BEU', 'KMB', '9704', '9705', '31825', '74826', 'FR 8374', 'FRA', 
    'RHB', 'CFM', 'FK 81'
]

def normalize_text(s):
    if not s:
        return ""
    s = unicodedata.normalize('NFC', str(s).strip())
    return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8').lower()

def canonical_key(s):
    """
    Marka/anahtar metnini kanonik biçime getirir (ASCII + büyük harf).

    Neden gerekli: Python'un upper() metodu Türkçe noktalı İ (U+0130) ve
    î/û/ö/ç/ş/ğ harflerini ASCII karşılıklarına katmaz. Bu yüzden
    "BABYLİSS" ile "BABYLISS" farklı anahtarlar olarak görülüyor ve aynı
    marka iki ayrı sekmede/listede bölünüyordu. normalize_text (NFKD + ASCII)
    bu katlamayı zaten doğru yapıyor.
    """
    return normalize_text(s).upper()

def normalize_tr(s):
    if not s:
        return ""
    tr_map = str.maketrans({
        'ı': 'i', 'İ': 'i', 'I': 'i', 'i': 'i',
        'ş': 's', 'Ş': 's',
        'ğ': 'g', 'Ğ': 'g',
        'ü': 'u', 'Ü': 'u',
        'ö': 'o', 'Ö': 'o',
        'ç': 'c', 'Ç': 'c'
    })
    return str(s).translate(tr_map).lower()

EXCLUDE_PRODUCT_KEYWORDS = [
    'camasir kurutma makinesi',
    'camasir makinesi',
    'bulasik makinesi',
    'buzdolabi',
    'ankastre firin',
    'mini firin',
    'derin dondurucu',
    'davlumbaz',
    'klima',
    'ek garanti',
    'garanti uzatma',
    'garanti paketi',
    'yazarkasa',
    'yazar kasa',
    'pos cihazi',
    'odeme kaydedici',
]

def is_conditional_appliance(name):
    norm = normalize_tr(name)
    return ('aspirator' in norm or 'ocak' in norm)

def is_major_appliance_or_warranty(name):
    norm = normalize_tr(name)
    if not norm:
        return False

    # 1. KESİNLİKLE KORUNACAKLAR (İstisnalar)
    if 'lenovo' in norm and 'garanti' in norm:
        return False
    if 'sac kurutma' in norm:
        return False
    if 'mikrodalga' in norm:
        return False

    # 2. TV / Televizyon kontrolü
    if 'televizyon' in norm:
        return True
    cleaned_words = norm.replace('"', ' ').replace("'", ' ').replace('-', ' ').replace(',', ' ').replace('.', ' ').split()
    if 'tv' in cleaned_words or 'qled' in cleaned_words or 'oled' in cleaned_words:
        return True

    # 3. Kurutma makinesi (çamaşır / büyük)
    if 'kurutma makinesi' in norm:
        return True

    # 4. Fırınlar (Mikrodalga yukarıda istisna yapıldığı için diğer ankastre/mini fırınlar elenir)
    if 'ankastre firin' in norm or 'mini firin' in norm or 'buhar destekli firin' in norm or 'solo firin' in norm or ' firin' in norm or norm.endswith('firin'):
        return True

    # 5. Yazarkasa / POS Cihazları (Beko X30 TR Yazarkasa POS vb.)
    if 'yazarkasa' in norm or 'yazar kasa' in norm or 'pos cihazi' in norm or 'odeme kaydedici' in norm or 'okc' in cleaned_words:
        return True
    if 'pos' in cleaned_words and ('yazar' in norm or 'kasa' in norm or 'x30' in norm or '300 tr' in norm or '220' in norm or '400' in norm or 'mobil' in norm or 'cihaz' in norm or 'eft' in norm):
        return True

    # 6. Diğer beyaz eşya & garanti listesi
    for kw in EXCLUDE_PRODUCT_KEYWORDS:
        if kw in norm:
            return True

    return False

def is_excluded_product(name, brand=''):
    if brand and brand != 'BEKO' and 'BEKO' not in canonical_key(name):
        return False
    return is_major_appliance_or_warranty(name)

def find_column_index(headers, possible_names):
    norm_headers = [normalize_text(h) for h in headers]
    for name in possible_names:
        norm_name = normalize_text(name)
        if norm_name in norm_headers:
            return norm_headers.index(norm_name)
            
    for name in possible_names:
        norm_name = normalize_text(name)
        for idx, h in enumerate(norm_headers):
            if norm_name and (norm_name == h or norm_name in h):
                return idx
    return -1

def detect_brand(full_name, store='', brand_col=''):
    if brand_col:
        m = canonical_key(brand_col)
        if m:
            for kb in KNOWN_BRANDS:
                if kb in m:
                    return kb
            return m

    u_upper = canonical_key(full_name)
    words = full_name.split()
    first_word = canonical_key(words[0]) if words else ''

    # 1. First word or start of title matches known brand
    for kb in KNOWN_BRANDS:
        if kb in first_word or u_upper.startswith(kb):
            return kb

    # 2. Known Beko model codes (like 7053MB, KMX 1002, etc.)
    for bmp in BEKO_MODELS:
        if bmp in u_upper:
            return 'BEKO'

    # 3. Store name has Beko or title has Beko
    if 'BEKO' in canonical_key(store) or 'BEKO' in u_upper:
        return 'BEKO'

    # 4. Known brand anywhere in title
    for kb in KNOWN_BRANDS:
        if f" {kb} " in f" {u_upper} " or f"({kb})" in u_upper or f"[{kb}]" in u_upper:
            return kb

    return first_word or 'DİĞER'

def get_base_dirs():
    """Returns dynamic base dirs that work on ANY computer."""
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
        
    user_profile = os.environ.get('USERPROFILE', os.path.expanduser('~'))
    downloads_dir = os.path.join(user_profile, 'Downloads')
    
    dirs = [app_dir]
    if os.path.exists(downloads_dir) and downloads_dir not in dirs:
        dirs.append(downloads_dir)
    return dirs

def get_latest_download_file():
    files = get_latest_download_files(max_files=1)
    return files[0] if files else None

def get_latest_download_files(max_files=10):
    user_profile = os.environ.get('USERPROFILE', os.path.expanduser('~'))
    downloads_dir = os.path.join(user_profile, 'Downloads')
    if not os.path.exists(downloads_dir):
        return []
        
    files = glob.glob(os.path.join(downloads_dir, '*.xls*'))
    valid_files = []
    for f in files:
        base = unicodedata.normalize('NFC', os.path.basename(f))
        if base.startswith('~$'):
            continue
        if base.startswith('Urun_Toplama_Listesi') or base.startswith('Beko_Urun_Toplama'):
            continue
        valid_files.append(f)

    if not valid_files:
        return []
    valid_files.sort(key=os.path.getmtime, reverse=True)
    return valid_files[:max_files]

def save_workbook_safely(wb, target_path):
    """
    Saves workbook safely. If target_path is open/locked in Excel (PermissionError),
    automatically creates an alternative file (_1.xlsx, _2.xlsx, etc.) so it NEVER crashes.
    """
    try:
        wb.save(target_path)
        return target_path, False
    except PermissionError:
        dir_name = os.path.dirname(target_path)
        base_name, ext = os.path.splitext(os.path.basename(target_path))
        for counter in range(1, 100):
            alt_path = os.path.join(dir_name, f"{base_name}_{counter}{ext}")
            try:
                wb.save(alt_path)
                return alt_path, True
            except PermissionError:
                continue
        raise PermissionError(f"Dosya Excel'de açık ve kaydedilemedi: {target_path}")

def parse_single_file_rows(file_item, custom_filename=None, temp_dir=None):
    """
    Parses a single file (filepath, bytes, or Flask FileStorage) and returns (raw_rows, filename).
    """
    if isinstance(file_item, str):
        file_path = file_item
        filename = unicodedata.normalize('NFC', os.path.basename(file_path))
        is_xls = file_path.lower().endswith('.xls')
    else:
        if hasattr(file_item, 'filename') and file_item.filename:
            filename = safe_filename(file_item.filename)
        else:
            filename = safe_filename(custom_filename)

        temp_dir = temp_dir or os.getcwd()
        os.makedirs(temp_dir, exist_ok=True)
        file_path = os.path.join(temp_dir, filename)

        # Save securely
        if hasattr(file_item, 'save'):
            file_item.save(file_path)
        elif hasattr(file_item, 'read'):
            with open(file_path, 'wb') as f:
                f.write(file_item.read())
        else:
            with open(file_path, 'wb') as f:
                f.write(file_item)

        is_xls = filename.lower().endswith('.xls')

    raw_rows = []

    if is_xls:
        wb_in = xlrd.open_workbook(file_path)
        sh = wb_in.sheet_by_index(0)
        
        # Smart header detection: check first 5 rows
        header_row_idx = 0
        headers = []
        for r in range(min(5, sh.nrows)):
            row_vals = [str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)]
            row_str = ' '.join(row_vals).lower()
            if any(k in row_str for k in ['ürün', 'urun', 'stok', 'barkod', 'adet', 'miktar', 'sipariş', 'siparis']):
                header_row_idx = r
                headers = row_vals
                break
        if not headers and sh.nrows > 0:
            headers = [str(sh.cell_value(0, c)).strip() for c in range(sh.ncols)]
        
        idx_name = find_column_index(headers, ['ürün adı', 'ürün', 'urun adi', 'urun', 'ürün ismi', 'product name'])
        idx_variant = find_column_index(headers, ['varyant', 'variant', 'seçenek', 'özellik'])
        idx_code = find_column_index(headers, ['stok kodu', 'ürün kodu', 'urun kodu', 'stok', 'sku', 'model kodu'])
        idx_qty = find_column_index(headers, ['adet', 'miktar', 'miktar (adet)', 'sipariş adedi', 'quantity', 'qty'])
        idx_ord = find_column_index(headers, ['sipariş no', 'sipariş numarası', 'siparis no', 'paket no', 'order no'])
        idx_bar = find_column_index(headers, ['gtin (barkod)', 'barkod', 'gtin', 'barcode', 'ean'])
        idx_cust = find_column_index(headers, ['üye adı soyadı', 'fatura - müşteri', 'müşteri', 'musteri', 'alıcı', 'alici', 'ad soyad', 'pazaryeri kullanıcı'])
        idx_magaza = find_column_index(headers, ['mağaza', 'magaza', 'satıcı', 'store'])
        idx_marka = find_column_index(headers, ['marka', 'brand'])
        idx_pazar = find_column_index(headers, ['pazaryeri', 'pazar yeri', 'platform'])

        for r in range(header_row_idx + 1, sh.nrows):
            name_val = str(sh.cell_value(r, idx_name)).strip() if idx_name != -1 else ''
            if not name_val:
                continue
            
            var_val = str(sh.cell_value(r, idx_variant)).strip() if idx_variant != -1 else ''
            if var_val and var_val.lower() not in name_val.lower():
                full_name = f"{name_val} ({var_val})"
            else:
                full_name = name_val
            
            stok_val = str(sh.cell_value(r, idx_code)).strip() if idx_code != -1 else ''
            if stok_val.endswith('.0'): stok_val = stok_val[:-2]
            
            bar_val = str(sh.cell_value(r, idx_bar)).strip() if idx_bar != -1 else ''
            if bar_val.endswith('.0'): bar_val = bar_val[:-2]
            
            qty_val = 1
            if idx_qty != -1:
                q = sh.cell_value(r, idx_qty)
                try: qty_val = int(float(q))
                except Exception: qty_val = 1
                
            ord_val = str(sh.cell_value(r, idx_ord)).strip() if idx_ord != -1 else ''
            if ord_val.endswith('.0'): ord_val = ord_val[:-2]

            cust_val = str(sh.cell_value(r, idx_cust)).strip() if idx_cust != -1 else ''
            magaza_val = str(sh.cell_value(r, idx_magaza)).strip() if idx_magaza != -1 else ''
            marka_val = str(sh.cell_value(r, idx_marka)).strip() if idx_marka != -1 else ''
            pazar_val = str(sh.cell_value(r, idx_pazar)).strip() if idx_pazar != -1 else ''
            
            brand = detect_brand(full_name, magaza_val, marka_val)

            raw_rows.append({
                'name': full_name,
                'stok': stok_val,
                'barkod': bar_val,
                'qty': qty_val,
                'order_no': ord_val,
                'customer': cust_val,
                'brand': brand,
                'magaza': magaza_val,
                'pazar': pazar_val,
                'source_file': filename
            })
    else:
        wb_in = openpyxl.load_workbook(file_path, data_only=True)
        sh = wb_in.active
        
        # Smart header detection: check first 5 rows
        header_row_idx = 1
        headers = []
        for r in range(1, min(6, sh.max_row + 1)):
            row_vals = [str(sh.cell(r, c).value or '').strip() for c in range(1, sh.max_column + 1)]
            row_str = ' '.join(row_vals).lower()
            if any(k in row_str for k in ['ürün', 'urun', 'stok', 'barkod', 'adet', 'miktar', 'sipariş', 'siparis']):
                header_row_idx = r
                headers = row_vals
                break
        if not headers and sh.max_row > 0:
            headers = [str(sh.cell(1, c).value or '').strip() for c in range(1, sh.max_column + 1)]
        
        idx_name = find_column_index(headers, ['ürün adı', 'ürün', 'urun adi', 'urun', 'ürün ismi', 'product name'])
        idx_variant = find_column_index(headers, ['varyant', 'variant', 'seçenek', 'özellik'])
        idx_code = find_column_index(headers, ['stok kodu', 'ürün kodu', 'urun kodu', 'stok', 'sku', 'model kodu'])
        idx_qty = find_column_index(headers, ['adet', 'miktar', 'miktar (adet)', 'sipariş adedi', 'quantity', 'qty'])
        idx_ord = find_column_index(headers, ['sipariş no', 'sipariş numarası', 'siparis no', 'paket no', 'order no'])
        idx_bar = find_column_index(headers, ['gtin (barkod)', 'barkod', 'gtin', 'barcode', 'ean'])
        idx_cust = find_column_index(headers, ['fatura - müşteri', 'üye adı soyadı', 'müşteri', 'musteri', 'alıcı', 'alici', 'ad soyad', 'pazaryeri kullanıcı'])
        idx_magaza = find_column_index(headers, ['mağaza', 'magaza', 'satıcı', 'store'])
        idx_marka = find_column_index(headers, ['marka', 'brand'])
        idx_pazar = find_column_index(headers, ['pazaryeri', 'pazar yeri', 'platform'])

        for r in range(header_row_idx + 1, sh.max_row + 1):
            name_val = str(sh.cell(r, idx_name + 1).value or '').strip() if idx_name != -1 else ''
            if not name_val:
                continue
            
            var_val = str(sh.cell(r, idx_variant + 1).value or '').strip() if idx_variant != -1 else ''
            if var_val and var_val.lower() not in name_val.lower():
                full_name = f"{name_val} ({var_val})"
            else:
                full_name = name_val
            
            stok_val = str(sh.cell(r, idx_code + 1).value or '').strip() if idx_code != -1 else ''
            if stok_val.endswith('.0'): stok_val = stok_val[:-2]
            
            bar_val = str(sh.cell(r, idx_bar + 1).value or '').strip() if idx_bar != -1 else ''
            if bar_val.endswith('.0'): bar_val = bar_val[:-2]
            
            qty_val = 1
            if idx_qty != -1:
                q = sh.cell(r, idx_qty + 1).value
                try: qty_val = int(float(q))
                except Exception: qty_val = 1
                
            ord_val = str(sh.cell(r, idx_ord + 1).value or '').strip() if idx_ord != -1 else ''
            if ord_val.endswith('.0'): ord_val = ord_val[:-2]

            cust_val = str(sh.cell(r, idx_cust + 1).value or '').strip() if idx_cust != -1 else ''
            magaza_val = str(sh.cell(r, idx_magaza + 1).value or '').strip() if idx_magaza != -1 else ''
            marka_val = str(sh.cell(r, idx_marka + 1).value or '').strip() if idx_marka != -1 else ''
            pazar_val = str(sh.cell(r, idx_pazar + 1).value or '').strip() if idx_pazar != -1 else ''
            
            brand = detect_brand(full_name, magaza_val, marka_val)

            raw_rows.append({
                'name': full_name,
                'stok': stok_val,
                'barkod': bar_val,
                'qty': qty_val,
                'order_no': ord_val,
                'customer': cust_val,
                'brand': brand,
                'magaza': magaza_val,
                'pazar': pazar_val,
                'source_file': filename
            })
        wb_in.close()

    return raw_rows, filename

def read_rows_from_files(files_or_paths, temp_dir=None):
    """
    Bir veya birden fazla dosyayı ayrıştırır ve ham satırları döner.
    Konsolidasyon yapmaz, diske çıktı yazmaz.
    """
    if not files_or_paths:
        raise ValueError("İşlenecek dosya bulunamadı.")

    all_raw_rows = []
    source_filenames = []

    for item in files_or_paths:
        try:
            rows, fname = parse_single_file_rows(item, temp_dir=temp_dir)
            all_raw_rows.extend(rows)
            if fname not in source_filenames:
                source_filenames.append(fname)
        except Exception as e:
            print(f"Uyarı: {getattr(item, 'filename', item)} ayrıştırılırken hata oluştu: {e}")

    if not all_raw_rows:
        raise ValueError("Yüklenen dosyalarda uygun sipariş satırı bulunamadı.")

    return all_raw_rows, source_filenames


def clean_product_name(s):
    """
    Ürün adındaki pazaryeri varyant/özellik kalıntılarını ve köşeli parantezli ekleri temizler.
    Örn: "[SSD Kapasitesi:512 GB, Ram (Sistem Belleği):8 GB]", "[Renk:Beyaz]", ", one size" vb.
    """
    if not s:
        return ""
    # 1. Parantez içi pazaryeri eklerini temizle: [Renk:Beyaz], [SSD Kapasitesi:...], [Beden:...] vb.
    s = re.sub(r'\s*\[[^\]]+\]', '', s)
    # 2. ", one size" veya ", standart" gibi pazaryeri kalıntılarını temizle
    s = re.sub(r',\s*(one size|standart|tek ebat)\b', '', s, flags=re.IGNORECASE)
    # 3. Fazla boşlukları temizle
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def pick_best_name(names, brand=None):
    """
    Aynı ürün için gelen farklı ad varyasyonları arasından en temiz ve açıklayıcı olanı seçer.
    Marka adını içeren ve gereksiz sistem kodları taşımayan başlığa öncelik verir.
    """
    cleaned_names = [clean_product_name(n) for n in names if n]
    if not cleaned_names:
        return ""
    b_upper = (brand or "").upper()
    def score(n):
        has_brand = 1 if b_upper and b_upper in n.upper() else 0
        has_garbage = -1 if re.search(r'TYC[A-Z0-9]{15,}', n) else 0
        return (has_brand, has_garbage, len(n))
    cleaned_names.sort(key=score, reverse=True)
    return cleaned_names[0]


def consolidate_and_build(all_raw_rows, source_filenames, filter_beko=False):
    """
    Mükerrer ürünleri konsolide eder ve A4 formatlı Excel dosyalarını BELLEKTE üretir.
    Bu fonksiyon hiçbir disk yazma işlemi yapmaz; hem masaüstü hem bulut yolu bunu kullanır.
    """
    # Group by brand & CONSOLIDATE duplicate products across ALL files
    brand_consolidated = defaultdict(dict)
    beko_consolidated = dict()

    # Büyük beyaz eşya içeren sipariş/müşteri kümesi (SADECE BEKO İÇİN)
    # Kural: Bu filtreleme sadece Beko markası için uygulanır. Geriye kalan tüm markalar (Tefal, Philips, Teka vb.) doğrudan listelenir.
    orders_with_beko_major = set()
    for item in all_raw_rows:
        is_beko = (item['brand'] == 'BEKO' or 'BEKO' in canonical_key(item['name']))
        if is_beko and is_major_appliance_or_warranty(item['name']):
            if item.get('order_no'):
                orders_with_beko_major.add(item['order_no'])
            if item.get('customer'):
                orders_with_beko_major.add(canonical_key(item['customer']))

    def is_item_excluded(item):
        is_beko = (item['brand'] == 'BEKO' or 'BEKO' in canonical_key(item['name']))
        # Sadece BEKO markasına uygulanır, diğer markalar asla elenmez
        if not is_beko:
            return False

        name = item['name']
        if is_major_appliance_or_warranty(name):
            return True
        if is_conditional_appliance(name):
            ord_no = item.get('order_no', '')
            cust = canonical_key(item.get('customer', ''))
            return bool((ord_no and ord_no in orders_with_beko_major) or (cust and cust in orders_with_beko_major))
        return False

    for item in all_raw_rows:
        if is_item_excluded(item):
            continue

        brand = item['brand']
        stok = (item.get('stok') or '').strip().upper()
        if stok.endswith('.0'):
            stok = stok[:-2]

        is_beko = (brand == 'BEKO' or 'BEKO' in canonical_key(item['name']))
        if is_beko:
            brand = 'BEKO'

        clean_name = clean_product_name(item['name'])
        is_valid_stok = bool(stok and len(stok) >= 3 and stok.lower() not in ('-', 'yok', '0', 'none', 'null'))

        if is_valid_stok:
            merge_key = ('STOK', brand, stok)
        else:
            merge_key = ('NAME', brand, canonical_key(clean_name))

        target_dict = beko_consolidated if (filter_beko and (is_beko or brand == 'GRUNDIG')) else brand_consolidated[brand]

        if merge_key not in target_dict:
            target_dict[merge_key] = {
                'names': [],
                'stok': stok,
                'qty': 0,
                'brand': brand
            }
        target_dict[merge_key]['names'].append(item['name'])
        if not target_dict[merge_key]['stok'] and stok:
            target_dict[merge_key]['stok'] = stok
        target_dict[merge_key]['qty'] += item['qty']

    brand_orders = {}
    for brand, items_dict in brand_consolidated.items():
        processed_list = []
        for merge_key, data in items_dict.items():
            best_name = pick_best_name(data['names'], brand=brand)
            processed_list.append({
                'name': best_name,
                'stok': data['stok'],
                'qty': data['qty'],
                'brand': data['brand']
            })
        brand_orders[brand] = sorted(processed_list, key=lambda x: x['qty'], reverse=True)

    beko_list = list(brand_orders.get('BEKO', []))
    grundig_list = list(brand_orders.get('GRUNDIG', []))
    beko_grundig_orders = sorted(beko_list + grundig_list, key=lambda x: x['qty'], reverse=True)

    if filter_beko:
        processed_beko = []
        for merge_key, data in beko_consolidated.items():
            best_name = pick_best_name(data['names'], brand=data['brand'])
            processed_beko.append({
                'name': best_name,
                'stok': data['stok'],
                'qty': data['qty'],
                'brand': data['brand']
            })
        beko_orders = sorted(processed_beko, key=lambda x: x['qty'], reverse=True)
        beko_grundig_orders = beko_orders
    else:
        beko_orders = beko_grundig_orders

    other_brand_keys = [b for b in brand_orders.keys() if b not in ('BEKO', 'GRUNDIG')]
    sorted_other_keys = sorted(other_brand_keys, key=lambda b: sum(x['qty'] for x in brand_orders[b]), reverse=True)

    grouped_items = []
    for brand in sorted_other_keys:
        grouped_items.extend(brand_orders[brand])
    other_items_count = len(grouped_items)

    if not filter_beko:
        grouped_items.extend(beko_grundig_orders)

    font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    font_regular = Font(name='Segoe UI', size=10)
    font_qty = Font(name='Segoe UI', size=11, bold=True)
    font_code = Font(name='Segoe UI', size=10, bold=True, color='1E293B')

    fill_header = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    fill_zebra = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')

    border_thin = Side(border_style='thin', color='94A3B8')
    cell_border = Border(top=border_thin, left=border_thin, right=border_thin, bottom=border_thin)

    def setup_a4_sheet(ws, title, rows_data, page_break_after_idx=None):
        ws.title = title
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1

        has_page_break = (page_break_after_idx is not None and 0 < page_break_after_idx < len(rows_data))
        if has_page_break:
            # 2 sayfa A4: 1. sayfa diger markalar, 2. sayfa Beko & Grundig
            ws.page_setup.fitToHeight = 2
            ws.print_title_rows = '1:1'
        else:
            # Tek sayfa A4
            ws.page_setup.fitToHeight = 1

        # Sayfa basina dusen maksimum satir sayisina gore dinamik optimizasyon
        if has_page_break:
            max_page_rows = max(page_break_after_idx, len(rows_data) - page_break_after_idx)
        else:
            max_page_rows = len(rows_data)

        is_dense = (max_page_rows > 35)

        ws.page_margins.left = 0.35 if is_dense else 0.4
        ws.page_margins.right = 0.35 if is_dense else 0.4
        ws.page_margins.top = 0.4 if is_dense else 0.5
        ws.page_margins.bottom = 0.4 if is_dense else 0.5
        ws.page_margins.header = 0.2 if is_dense else 0.25
        ws.page_margins.footer = 0.2 if is_dense else 0.25

        ws.print_options.horizontalCentered = True
        ws.print_options.gridLines = True

        f_header = Font(name='Segoe UI', size=10 if is_dense else 11, bold=True, color='FFFFFF')
        f_regular = Font(name='Segoe UI', size=9.5 if is_dense else 10)
        f_qty = Font(name='Segoe UI', size=10.5 if is_dense else 11, bold=True)
        f_code = Font(name='Segoe UI', size=9.5 if is_dense else 10, bold=True, color='1E293B')

        headers = ['Miktar', 'Ürün Adı', 'Stok Kodu']
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.font = f_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = cell_border
        ws.row_dimensions[1].height = 24 if is_dense else 28

        row_h = 20 if is_dense else 24

        for idx, itm in enumerate(rows_data, 1):
            row_num = idx + 1

            c_q = ws.cell(row=row_num, column=1, value=itm['qty'])
            c_q.alignment = Alignment(horizontal='center', vertical='center')
            c_q.font = f_qty

            c_n = ws.cell(row=row_num, column=2, value=itm['name'])
            c_n.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            c_n.font = f_regular

            c_s = ws.cell(row=row_num, column=3, value=itm['stok'])
            c_s.alignment = Alignment(horizontal='center', vertical='center')
            c_s.font = f_code

            ws.row_dimensions[row_num].height = row_h

            for c in range(1, 4):
                cell = ws.cell(row=row_num, column=c)
                cell.border = cell_border
                if idx % 2 == 0:
                    cell.fill = fill_zebra

        if has_page_break:
            ws.row_breaks.append(Break(id=page_break_after_idx + 1))

        # Tablo sütun genişliklerini içerik uzunluğuna göre dinamik ayarla
        max_q_len = max([len('Miktar')] + [len(str(itm.get('qty', ''))) for itm in rows_data]) if rows_data else 6
        max_s_len = max([len('Stok Kodu')] + [len(str(itm.get('stok', ''))) for itm in rows_data]) if rows_data else 9

        col_a_w = max(8, min(max_q_len + 4, 12))
        col_c_w = max(14, min(max_s_len + 4, 22))

        # A4 dikey toplam genişliği (~84-86 pt) içinde kalan genişliği Ürün Adı'na ver
        target_total_w = 84 if is_dense else 86
        col_b_w = max(50, target_total_w - col_a_w - col_c_w)

        ws.column_dimensions['A'].width = col_a_w
        ws.column_dimensions['B'].width = col_b_w
        ws.column_dimensions['C'].width = col_c_w

    # 1. Main Workbook (Non-Beko or All)
    wb_main = openpyxl.Workbook()
    ws1 = wb_main.active
    # Sheet 1: Genel Toplama Listesi (Sayfa 1: Diger Markalar, Sayfa 2: Beko & Grundig)
    page_break_idx = other_items_count if (other_items_count > 0 and len(beko_grundig_orders) > 0 and not filter_beko) else None
    setup_a4_sheet(ws1, "Ürün Toplama Listesi", grouped_items, page_break_after_idx=page_break_idx)

    # Sheet 2: BEKO & GRUNDIG Sekmesi (Eger Beko & Grundig siparisi varsa)
    if beko_grundig_orders and not filter_beko:
        ws_bg = wb_main.create_sheet(title="BEKO & GRUNDIG")
        setup_a4_sheet(ws_bg, "BEKO & GRUNDIG", beko_grundig_orders)

    # Sheet 3+: Diger Marka Sekmeleri (Tefal, Babyliss, Philips vb.)
    for brand in sorted_other_keys:
        sheet_title = brand[:30]
        ws_b = wb_main.create_sheet(title=sheet_title)
        setup_a4_sheet(ws_b, sheet_title, brand_orders[brand])

    # 2. Beko & Grundig Workbook (Ayri dosya)
    wb_beko = None
    if beko_orders:
        wb_beko = openpyxl.Workbook()
        ws_beko = wb_beko.active
        setup_a4_sheet(ws_beko, "Beko & Grundig Toplama Listesi", beko_orders)

    brand_stats = []
    if beko_grundig_orders and not filter_beko:
        brand_stats.append({
            'brand': 'BEKO & GRUNDIG',
            'count': len(beko_grundig_orders),
            'qty': sum(x['qty'] for x in beko_grundig_orders)
        })
    for b in sorted_other_keys:
        brand_stats.append({
            'brand': b,
            'count': len(brand_orders[b]),
            'qty': sum(x['qty'] for x in brand_orders[b])
        })

    excluded_rows = [item for item in all_raw_rows if is_item_excluded(item)]

    summary_filename = source_filenames[0] if len(source_filenames) == 1 else f"{len(source_filenames)} Dosya Birleştirildi"

    return {
        'source_filename': summary_filename,
        'source_files': source_filenames,
        'total_files': len(source_filenames),
        'total_orders': len(all_raw_rows),
        'excluded_rows': len(excluded_rows),
        'excluded_qty': sum(x['qty'] for x in excluded_rows),
        'non_beko_count': len(grouped_items),
        'non_beko_qty': sum(x['qty'] for x in grouped_items),
        'other_items_count': other_items_count,
        'beko_count': len(beko_orders),
        'beko_qty': sum(x['qty'] for x in beko_orders),
        'filter_beko': filter_beko,
        'brands': brand_stats,
        'grouped_items': grouped_items,
        'beko_items': beko_orders,
        'is_cloud': IS_CLOUD,
        '_wb_main': wb_main,
        '_wb_beko': wb_beko,
    }


def workbooks_to_bytes(wb_main, wb_beko=None):
    """Üretilmiş workbook'ları indirilebilir byte dizisine çevirir (disk kullanmaz)."""
    main_buf = io.BytesIO()
    wb_main.save(main_buf)
    beko_buf = None
    if wb_beko is not None:
        beko_buf = io.BytesIO()
        wb_beko.save(beko_buf)
    return main_buf.getvalue(), (beko_buf.getvalue() if beko_buf else None)


def parse_and_process_multiple_files(files_or_paths, filter_beko=False, target_date=None):
    """
    MASAÜSTÜ YOLU: dosyaları işler, konsolide eder ve A4 Excel çıktılarını
    tarih klasörlerine diske yazar. Davranış eskisiyle birebir aynıdır.
    """
    if target_date is None:
        target_date = datetime.date.today().strftime('%Y-%m-%d')

    base_dirs = get_base_dirs()

    date_folders = []
    for b_dir in base_dirs:
        d_folder = os.path.join(b_dir, target_date)
        os.makedirs(d_folder, exist_ok=True)
        date_folders.append(d_folder)

    all_raw_rows, source_filenames = read_rows_from_files(files_or_paths, temp_dir=date_folders[0])

    result = consolidate_and_build(all_raw_rows, source_filenames, filter_beko=filter_beko)
    wb_main = result.pop('_wb_main')
    wb_beko = result.pop('_wb_beko')

    # Save safely
    saved_files = []
    main_filename = "Urun_Toplama_Listesi_A4_Cikti.xlsx"
    beko_filename = "Beko_Urun_Toplama_Listesi_A4.xlsx"
    primary_main_path = None
    primary_beko_path = None
    was_renamed = False

    for idx, d_folder in enumerate(date_folders):
        path_main = os.path.join(d_folder, main_filename)
        actual_main_path, renamed_m = save_workbook_safely(wb_main, path_main)
        saved_files.append(actual_main_path)
        if idx == 0:
            primary_main_path = actual_main_path
            if renamed_m: was_renamed = True

        if wb_beko:
            path_beko = os.path.join(d_folder, beko_filename)
            actual_beko_path, renamed_b = save_workbook_safely(wb_beko, path_beko)
            saved_files.append(actual_beko_path)
            if idx == 0:
                primary_beko_path = actual_beko_path
                if renamed_b: was_renamed = True

    result.update({
        'date': target_date,
        'saved_files': saved_files,
        'main_excel_path': primary_main_path,
        'beko_excel_path': primary_beko_path,
        'folder_path': date_folders[0],
        'was_renamed': was_renamed,
    })
    return result


def process_in_memory(files_or_paths, filter_beko=False):
    """
    BULUT YOLU: hiçbir disk yazma işlemi yapmaz. Excel çıktıları bellekte üretilir
    ve byte olarak döner. Vercel gibi kalıcı diski olmayan ortamlar içindir.
    """
    temp_dir = CLOUD_TMP_DIR if IS_CLOUD else None
    all_raw_rows, source_filenames = read_rows_from_files(files_or_paths, temp_dir=temp_dir)

    result = consolidate_and_build(all_raw_rows, source_filenames, filter_beko=filter_beko)
    wb_main = result.pop('_wb_main')
    wb_beko = result.pop('_wb_beko')

    main_bytes, beko_bytes = workbooks_to_bytes(wb_main, wb_beko)

    result.update({
        'date': datetime.date.today().strftime('%Y-%m-%d'),
        'main_excel_path': None,
        'beko_excel_path': None,
        'folder_path': None,
        'was_renamed': False,
        '_main_bytes': main_bytes,
        '_beko_bytes': beko_bytes,
    })
    return result

def parse_and_process_file(file_path_or_bytes, custom_filename=None, filter_beko=False, target_date=None):
    """
    Backward-compatible wrapper for single file or list of files.
    """
    if isinstance(file_path_or_bytes, (list, tuple)):
        return parse_and_process_multiple_files(file_path_or_bytes, filter_beko=filter_beko, target_date=target_date)
    return parse_and_process_multiple_files([file_path_or_bytes], filter_beko=filter_beko, target_date=target_date)
