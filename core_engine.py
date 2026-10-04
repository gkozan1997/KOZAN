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
import secrets
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
    'HOTPOINT', 'ROWENTA',
    'DELONGHI', 'THOR', 'NESPRESSO', 'KRUPS', 'MELITTA', 'SAGE', 
    'JURA', 'ELECTROLUX', 'KARCHER', 'MOULINEX', 'SIMFER', 'KUMTEL', 
    'LUXELL', 'SINBO', 'KING', 'GOLDMASTER', 'REMINGTON', 'CONTESSE',
    'PANASONIC', 'SONY', 'LG', 'HUAWEI', 'HONOR', 'OPPO', 'VIVO', 
    'REALME', 'ASUS', 'ACER', 'HP', 'DELL', 'MSI', 'CASPER', 
    'MONSTER', 'TOSHIBA', 'ANKER', 'JBL', 'MARSHALL', 'HARMAN KARDON', 
    'LOGITECH', 'RAZER', 'STEELSERIES', 'CORSAIR', 'STANLEY', 'COSORI', 
    'NINJA', 'INSTANT POT', 'SCHAFER', 'EMSAN', 'NEVA', 'TAC', 'TACH', 
    'BERGHOFF', 'ZWILLING', 'PASABAHCE', 'LAV', 'BORCAM', 'SINFONIA', 
    'SMEG', 'BERETTA', 'FERROLI', 'DEMIRDOKUM', 'ECA', 'BAYMAK', 
    'VAILLANT', 'VIESSMANN', 'BUDERUS', 'DAIKIN', 'MITSUBISHI',
    'HITACHI', 'SHARP', 'PIONEER', 'HOMEND', 'COSMED', 'AWOX'
]

BRAND_ALIASES = {
    "DE'LONGHI": "DELONGHI",
    "DELONGHI": "DELONGHI",
    "DE LONGHI": "DELONGHI",
    "THOR KITCHEN": "THOR",
    "THOR": "THOR",
    "BABYLSS": "BABYLISS",
    "BABBYLISS": "BABYLISS",
    "IPHONE": "IPHONE",
    "APPLE": "IPHONE",
}

BEKO_START_MODELS = [
    'KMX', '7053', 'CM ', 'CMX', 'B 600', 'B 710', 'B600', 'B710',
    'BKK', 'BK RHC', 'BK ', 'BFC', 'BDE', 'TKM', 'BEU', 'KMB', 
    '9704', '9705', '31825', '74826', 'FR 8374', 'FRA ', 'FRA', 
    'RHB', 'CFM', 'FK 81', 'FK81', 'ADP', 'ADE', 'BOCD'
]

BEKO_INLINE_REGEX = re.compile(
    r'\b(7053MB|BKK\s*\d+|TKM\s*\d+|BEU\s*\d+|BFC\s*\d+|BDE\s*\d+[A-Z0-9]*|ADP\s*\d+[A-Z0-9]*|ADE\s*\d+[A-Z0-9]*|BOCD\s*[A-Z0-9]+|KMX\s*\d+|CMX\s*\d+|FRA\s*\d+|RHB\s*\d+|CFM\s*\d+|CM\s*\d{3,4}|BK\s*(?:RHC|\d{3,4}))\b',
    re.IGNORECASE
)

BEKO_MODELS = BEKO_START_MODELS


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
    'klima',
    'ek garanti',
    'garanti uzatma',
    'garanti paketi',
    'yazarkasa',
    'yazar kasa',
    'pos cihazi',
    'odeme kaydedici',
    'termosifon',
    'sofben',
    'kombi',
    'boyler',
    'ani su isitici',
    'ani su isiticisi',
]

def is_conditional_appliance(name):
    norm = normalize_tr(name)
    if 'aspirator' in norm or 'ocak' in norm or 'davlumbaz' in norm:
        return True
    if re.search(r'\b(adp\s*\d+[a-z0-9]*|bde\s*\d+[a-z0-9]*|ade\s*\d+[a-z0-9]*|hde\s*\d+[a-z0-9]*|cde\s*\d+[a-z0-9]*|bocd\s*[a-z0-9]+|p\s*38|p\s*27)\b', norm):
        return True
    return False

def is_major_appliance_or_warranty(name):
    norm = normalize_tr(name)
    if not norm:
        return False

    # 1. KESİNLİKLE KORUNACAKLAR (İstisnalar)
    if ('lenovo' in norm or 'ideapad' in norm) and 'garanti' in norm:
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

    # 6. Termosifon, şofben, kombi, boyler ve ani su ısıtıcıları (Büyük su ısıtma ve ısıtma cihazları)
    if 'termosifon' in norm or 'sofben' in norm or 'boyler' in norm or 'ani su isitici' in norm or 'ani su isiticisi' in norm:
        return True
    if re.search(r'\bkombi\b', norm):
        return True

    # 7. Diğer beyaz eşya & garanti listesi
    for kw in EXCLUDE_PRODUCT_KEYWORDS:
        if kw in norm:
            return True

    return False

def is_excluded_product(name, brand=''):
    if brand and brand != 'BEKO' and 'BEKO' not in canonical_key(name):
        return False
    u_name = canonical_key(name)
    if any(kb in u_name for kb in KNOWN_BRANDS if kb not in ('BEKO', 'GRUNDIG', 'LENOVO')):
        return False
    return is_major_appliance_or_warranty(name)

def safe_sheet_title(title, used_titles=None):
    """
    Excel sekme adını geçerli ve benzersiz hale getirir.
    Excel '[]:*?/\\' karakterlerini sessizce reddeder, 31 karakter sınırı koyar.
    Aksi halde openpyxl ValueError fırlatır ve tüm yükleme HTTP 500 döner (H-03).
    """
    used_titles = used_titles if used_titles is not None else set()
    clean = re.sub(r'[\[\]:\*\?/\\]', '', str(title or '')).strip()
    clean = re.sub(r'\s+', ' ', clean)
    clean = clean.strip("'")
    if not clean:
        clean = 'Diger'
    base = clean[:31]
    candidate = base
    suffix = 2
    while candidate.lower() in used_titles:
        tail = f'_{suffix}'
        candidate = base[:31 - len(tail)] + tail
        suffix += 1
    used_titles.add(candidate.lower())
    return candidate


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
            m_norm = BRAND_ALIASES.get(m, m)
            if 'IDEAPAD' in m_norm or 'LENOVO' in m_norm:
                return 'LENOVO'
            for kb in sorted(KNOWN_BRANDS, key=len, reverse=True):
                if kb == m_norm or f" {kb} " in f" {m_norm} ":
                    return BRAND_ALIASES.get(kb, kb)
            return m_norm

    u_upper = canonical_key(full_name)
    words = full_name.split()
    first_word = canonical_key(words[0]) if words else ''
    first_word_norm = BRAND_ALIASES.get(first_word, first_word)

    # 0. Lenovo & IdeaPad kontrolü (IdeaPad içeren veya başlayan tüm ürünler Lenovo olarak atanır ve 2. sayfaya gider)
    if 'IDEAPAD' in u_upper or 'LENOVO' in u_upper or 'IDEAPAD' in first_word or 'LENOVO' in first_word:
        return 'LENOVO'

    # 1. Başlangıç veya ilk kelime bilinen marka mı? (Delonghi, Thor, Philips, Tefal, vb.)
    for kb in sorted(KNOWN_BRANDS, key=len, reverse=True):
        if kb == first_word_norm or first_word.startswith(kb) or u_upper.startswith(kb):
            return BRAND_ALIASES.get(kb, kb)

    # 2. Ürün adının herhangi bir yerinde tam kelime / ayraç içi olarak bilinen marka var mı?
    for kb in sorted(KNOWN_BRANDS, key=len, reverse=True):
        pattern = r'(?:^|[\s\(\[\-\/\,\.\"])' + re.escape(kb) + r'(?:$|[\s\)\]\-\/\,\.\"])'
        if re.search(pattern, u_upper):
            return BRAND_ALIASES.get(kb, kb)

    # 3. Ürün adında açıkça BEKO geçiyor mu?
    if 'BEKO' in u_upper:
        return 'BEKO'

    # 4. Ürün adı doğrudan bir Beko model koduyla başlıyor mu? (KMX, 7053MB, CM 5964, B 710 vb.)
    for bmp in BEKO_START_MODELS:
        if u_upper.startswith(bmp) or first_word.startswith(bmp):
            return 'BEKO'

    # 5. Başlıkta net Beko model kodu kalıbı var mı? (örn: CM 5964, BKK 2300, BK RHC 1800)
    if BEKO_INLINE_REGEX.search(u_upper):
        return 'BEKO'

    # 6. Mağaza adında Beko geçiyor mu?
    if 'BEKO' in canonical_key(store):
        return 'BEKO'

    # 7. Yedek marka: ilk kelime gerçekten bir marka adına benziyor mu?
    # Model/seri numarası içeren ilk kelimeler (örn. "6715de", "X1450", "CM5964")
    # marka sanılırsa sahte bir Excel sekmesi ve sahte web sekmesi üretirdi (H-05).
    if (first_word_norm
            and len(first_word_norm) >= 3
            and not re.search(r'\d', first_word_norm)):
        return first_word_norm
    return 'DİĞER'

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

HEADER_HINTS = ['ürün', 'urun', 'stok', 'barkod', 'adet', 'miktar', 'sipariş', 'siparis']


def detect_header_row(row_values_list):
    """
    Verilen satır listesinden gerçek başlık satırını bulur (yasal uyarı olsa dahil).
    Sonda 'ürün/stok/adet/sipariş' ipuçlarından biri geçen ilk satır başlıktır.
    """
    for idx, row_vals in enumerate(row_values_list):
        row_str = ' '.join(str(v or '').strip() for v in row_vals).lower()
        if any(k in row_str for k in HEADER_HINTS):
            return idx, [str(v or '').strip() for v in row_vals]
    if row_values_list:
        return 0, [str(v or '').strip() for v in row_values_list[0]]
    return 0, []


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
        # Aynı adlı iki dosya (örn. Hepsiburada/Siparis.xlsx + Trendyol/Siparis.xlsx)
        # aynı geçici yola yazılırsa biri diğerini sessizce ezerdi (H-02).
        # Bu yüzden geçici dosya adına benzersiz bir ön ek eklenir.
        unique_name = f"{secrets.token_hex(4)}_{filename}"
        file_path = os.path.join(temp_dir, unique_name)

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

        # Smart header detection: ilk 5 satir taranir
        preview = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(min(5, sh.nrows))]
        header_row_idx, headers = detect_header_row(preview)
        
        idx_name = find_column_index(headers, ['ürün adı', 'ürün', 'urun adi', 'urun', 'ürün ismi', 'product name'])
        idx_variant = find_column_index(headers, ['varyant', 'variant', 'seçenek', 'özellik'])
        idx_code = find_column_index(headers, ['stok kodu', 'ürün kodu', 'urun kodu', 'stok', 'sku', 'model kodu'])
        idx_qty = find_column_index(headers, ['adet', 'miktar', 'miktar (adet)', 'sipariş adedi', 'quantity', 'qty'])
        idx_ord = find_column_index(headers, ['sipariş no', 'sipariş numarası', 'siparis no', 'paket no', 'order no'])
        idx_bar = find_column_index(headers, ['gtin (barkod)', 'barkod', 'gtin', 'barcode', 'ean'])
        idx_cust = find_column_index(headers, ['üye adı soyadı', 'fatura - müşteri', 'müşteri', 'musteri', 'alıcı', 'alici', 'ad soyad', 'pazaryeri kullanıcı'])
        idx_magaza = find_column_index(headers, ['mağaza', 'magaza', 'satıcı', 'store'])
        idx_marka = find_column_index(headers, ['marka', 'brand'])
        idx_pazar = find_column_index(headers, ['pazaryeri', 'pazar yeri', 'platform', 'marketplace'])

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

        # Smart header detection: ilk 5 satir taranir (1. satır yasal uyarı olabilir)
        preview = [
            [sh.cell(r, c).value for c in range(1, sh.max_column + 1)]
            for r in range(1, min(6, sh.max_row + 1))
        ]
        header_row_idx, headers = detect_header_row(preview)
        header_row_idx += 1  # openpyxl satirlari 1 tabanlidir
        
        idx_name = find_column_index(headers, ['ürün adı', 'ürün', 'urun adi', 'urun', 'ürün ismi', 'product name'])
        idx_variant = find_column_index(headers, ['varyant', 'variant', 'seçenek', 'özellik'])
        idx_code = find_column_index(headers, ['stok kodu', 'ürün kodu', 'urun kodu', 'stok', 'sku', 'model kodu'])
        idx_qty = find_column_index(headers, ['adet', 'miktar', 'miktar (adet)', 'sipariş adedi', 'quantity', 'qty'])
        idx_ord = find_column_index(headers, ['sipariş no', 'sipariş numarası', 'siparis no', 'paket no', 'order no'])
        idx_bar = find_column_index(headers, ['gtin (barkod)', 'barkod', 'gtin', 'barcode', 'ean'])
        idx_cust = find_column_index(headers, ['üye adı soyadı', 'fatura - müşteri', 'müşteri', 'musteri', 'alıcı', 'alici', 'ad soyad', 'pazaryeri kullanıcı'])
        idx_magaza = find_column_index(headers, ['mağaza', 'magaza', 'satıcı', 'store'])
        idx_marka = find_column_index(headers, ['marka', 'brand'])
        idx_pazar = find_column_index(headers, ['pazaryeri', 'pazar yeri', 'platform', 'marketplace'])

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
        n_up = n.upper()
        has_brand = 2 if (b_upper and b_upper in n_up) else (1 if (b_upper == 'LENOVO' and 'IDEAPAD' in n_up) else 0)
        has_garbage = -1 if re.search(r'TYC[A-Z0-9]{15,}', n) else 0
        return (has_brand, has_garbage, len(n))
    cleaned_names.sort(key=score, reverse=True)
    return cleaned_names[0]


COLOR_WORDS = [
    'beyaz', 'siyah', 'kirmizi', 'mavi', 'gri', 'inox', 'antrasit',
    'pembe', 'mor', 'bordo', 'yesil', 'sari', 'bej', 'gold', 'silver',
    'gumus', 'bakir', 'turuncu', 'krem', 'rose', 'rosegold',
    'white', 'black', 'red', 'blue', 'gray', 'grey', 'green', 'yellow', 'pink', 'purple', 'orange'
]

UNITS_OF_MEASURE = {
    'W', 'V', 'KG', 'GR', 'CM', 'MM', 'GB', 'TB', 
    'BTU', 'HZ', 'BAR', 'KPA', 'RPM', 'DEVIR', 'LT'
}

MODEL_CODE_RE = re.compile(
    r'\b([A-Z]{1,5}(?:\s+[A-Z]{1,2})?\s*\d{2,5})\s*([A-Z]{1,3})\b',
    re.IGNORECASE
)

def extract_product_sort_keys(name):
    """
    Ürün adından model gövdesini (base) ve renk/varyant kodunu (color_key) ayıklar.
    Böylece aynı modelin farklı renkleri (ör: CM 5964 R ve CM 5964 B) aynı base anahtarına sahip olur.
    """
    norm = normalize_tr(name)
    found_color = ""
    for c in COLOR_WORDS:
        if re.search(r'\b' + re.escape(c) + r'\b', norm):
            found_color = c
            break

    m = MODEL_CODE_RE.search(name)
    if m and m.group(2).upper() not in UNITS_OF_MEASURE:
        model_num = m.group(1).upper()
        suffix = m.group(2).upper()
        clean_base = name[:m.start()] + model_num + name[m.end():]
        if found_color:
            clean_base = re.sub(r'\b' + re.escape(found_color) + r'\b', '', clean_base, flags=re.IGNORECASE)
        clean_base = re.sub(r'\s+', ' ', clean_base).strip()
        color_key = suffix + ("_" + found_color if found_color else "")
        return canonical_key(clean_base), color_key
    else:
        clean_base = name
        if found_color:
            clean_base = re.sub(r'\b' + re.escape(found_color) + r'\b', '', clean_base, flags=re.IGNORECASE)
        clean_base = re.sub(r'\s+', ' ', clean_base).strip()
        color_key = found_color or "ZZZ"
        return canonical_key(clean_base), color_key

def sort_items_by_family_and_color(items):
    """
    Aynı ürün/model ailesine ait farklı renk ve varyantları bir araya toplar (alt alta)
    ve kendi içinde renklerine göre sıralar.
    
    Örnek:
    - Beko CM 5964 B Floral Çay Makinesi
    - Beko CM 5964 R Floral Çay Makinesi
    - Beko TKM 2341 Keyf-i Bol Beyaz
    - Beko TKM 2341 Keyf-i Bol Siyah
    """
    if not items:
        return []

    # 1. Ürün ailesine göre grupla
    groups = defaultdict(list)
    for itm in items:
        base_key, color_key = extract_product_sort_keys(itm['name'])
        groups[base_key].append((color_key, itm))

    # 2. Her aileyi kendi içinde renge ve miktara göre sırala
    sorted_groups = []
    for base_key, member_list in groups.items():
        member_list.sort(key=lambda x: (x[0], -x[1]['qty']))
        sorted_members = [m[1] for m in member_list]
        max_qty = max(m['qty'] for m in sorted_members)
        sorted_groups.append((max_qty, base_key, sorted_members))

    # 3. Grupları sırala: önce grup içindeki en yüksek miktar (büyükten küçüğe),
    #    miktarlar eşitse model/ürün adına göre alfabetik
    sorted_groups.sort(key=lambda g: (-g[0], g[1]))

    # 4. Düz liste haline getir
    result = []
    for _, _, members in sorted_groups:
        result.extend(members)
    return result


def consolidate_and_build(all_raw_rows, source_filenames, filter_beko=False, custom_note=""):
    """
    Mükerrer ürünleri konsolide eder ve A4 formatlı Excel dosyalarını BELLEKTE üretir.
    Bu fonksiyon hiçbir disk yazma işlemi yapmaz; hem masaüstü hem bulut yolu bunu kullanır.
    """
    # Group by brand & CONSOLIDATE duplicate products across ALL files
    brand_consolidated = defaultdict(dict)
    beko_consolidated = dict()

    # Büyük beyaz eşya içeren sipariş/müşteri kümesi (SADECE BEKO İÇİN)
    # Kural: Bu filtreleme sadece Beko markası için uygulanır. Geriye kalan tüm markalar (Delonghi, Thor, Tefal, Philips, Teka vb.) doğrudan listelenir.
    orders_with_beko_major = set()
    for item in all_raw_rows:
        i_brand = item.get('brand', '')
        if i_brand in KNOWN_BRANDS and i_brand not in ('BEKO', 'GRUNDIG', 'LENOVO'):
            is_beko = False
        else:
            is_beko = (i_brand == 'BEKO' or ('BEKO' in canonical_key(item['name']) and not any(kb in canonical_key(item['name']) for kb in KNOWN_BRANDS if kb not in ('BEKO', 'GRUNDIG', 'LENOVO'))))
        if is_beko and is_major_appliance_or_warranty(item['name']):
            if item.get('order_no'):
                orders_with_beko_major.add(item['order_no'])
            if item.get('customer'):
                orders_with_beko_major.add(canonical_key(item['customer']))

    def is_item_excluded(item):
        i_brand = item.get('brand', '')
        if i_brand in KNOWN_BRANDS and i_brand not in ('BEKO', 'GRUNDIG', 'LENOVO'):
            return False
        is_beko = (i_brand == 'BEKO' or ('BEKO' in canonical_key(item['name']) and not any(kb in canonical_key(item['name']) for kb in KNOWN_BRANDS if kb not in ('BEKO', 'GRUNDIG', 'LENOVO'))))
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

        if brand in KNOWN_BRANDS and brand not in ('BEKO', 'GRUNDIG', 'LENOVO'):
            is_beko = False
        else:
            is_beko = (brand == 'BEKO' or ('BEKO' in canonical_key(item['name']) and not any(kb in canonical_key(item['name']) for kb in KNOWN_BRANDS if kb not in ('BEKO', 'GRUNDIG', 'LENOVO'))))
        if is_beko:
            brand = 'BEKO'

        is_lenovo = (brand in ('LENOVO', 'IDEAPAD') or 'LENOVO' in canonical_key(item['name']) or 'IDEAPAD' in canonical_key(item['name']))
        if is_lenovo:
            brand = 'LENOVO'

        clean_name = clean_product_name(item['name'])
        is_valid_stok = bool(stok and len(stok) >= 3 and stok.lower() not in ('-', 'yok', '0', 'none', 'null'))

        if is_valid_stok:
            merge_key = ('STOK', brand, stok)
        else:
            merge_key = ('NAME', brand, canonical_key(clean_name))

        is_special_brand = (is_beko or brand in ('GRUNDIG', 'LENOVO'))
        target_dict = beko_consolidated if (filter_beko and is_special_brand) else brand_consolidated[brand]

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
        brand_orders[brand] = sort_items_by_family_and_color(processed_list)

    beko_list = list(brand_orders.get('BEKO', []))
    grundig_list = list(brand_orders.get('GRUNDIG', []))
    lenovo_list = list(brand_orders.get('LENOVO', []))
    beko_grundig_orders = beko_list + grundig_list + lenovo_list

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
        b_beko = [x for x in processed_beko if x['brand'] == 'BEKO']
        b_grundig = [x for x in processed_beko if x['brand'] == 'GRUNDIG']
        b_lenovo = [x for x in processed_beko if x['brand'] == 'LENOVO']
        b_other = [x for x in processed_beko if x['brand'] not in ('BEKO', 'GRUNDIG', 'LENOVO')]
        beko_orders = (
            sort_items_by_family_and_color(b_beko) +
            sort_items_by_family_and_color(b_grundig) +
            sort_items_by_family_and_color(b_lenovo) +
            sort_items_by_family_and_color(b_other)
        )
        beko_grundig_orders = beko_orders
    else:
        beko_orders = beko_grundig_orders

    other_brand_keys = [b for b in brand_orders.keys() if b not in ('BEKO', 'GRUNDIG', 'LENOVO')]
    sorted_other_keys = sorted(other_brand_keys, key=lambda b: sum(x['qty'] for x in brand_orders[b]), reverse=True)

    grouped_items = []
    for brand in sorted_other_keys:
        grouped_items.extend(brand_orders[brand])
    other_items_count = len(grouped_items)

    if not filter_beko:
        grouped_items.extend(beko_grundig_orders)

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
        # Hücrelere zaten kenarlık çizildiği için ızgara çizgileri basılırsa
        # tablo dışı boş alanlara da çizgi basılır (H-11).
        ws.print_options.gridLines = False

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
            is_note_row = (str(itm.get('qty', '')).upper() == 'NOT')

            c_q = ws.cell(row=row_num, column=1, value=itm['qty'])
            c_q.alignment = Alignment(horizontal='center', vertical='center')
            c_q.font = Font(name='Segoe UI', size=10, bold=True, color='B45309') if is_note_row else f_qty

            c_n = ws.cell(row=row_num, column=2, value=itm['name'])
            c_n.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            c_n.font = Font(name='Segoe UI', size=9.5, bold=True, color='000000') if is_note_row else f_regular

            c_s = ws.cell(row=row_num, column=3, value=itm.get('stok', ''))
            c_s.alignment = Alignment(horizontal='center', vertical='center')
            c_s.font = f_code

            ws.row_dimensions[row_num].height = row_h

            for c in range(1, 4):
                cell = ws.cell(row=row_num, column=c)
                cell.border = cell_border
                if is_note_row:
                    cell.fill = PatternFill(start_color='FFFBEB', end_color='FFFBEB', fill_type='solid')
                elif idx % 2 == 0:
                    cell.fill = fill_zebra

            if is_note_row:
                ws.merge_cells(start_row=row_num, start_column=2, end_row=row_num, end_column=3)

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
    sheet1_items = list(grouped_items)
    page_break_idx = other_items_count if (other_items_count > 0 and len(beko_grundig_orders) > 0 and not filter_beko) else None

    clean_note = (custom_note or '').strip()
    if clean_note:
        note_row = {'qty': 'NOT', 'name': f'Not: {clean_note}', 'stok': '', 'brand': 'NOTE'}
        if other_items_count > 0:
            sheet1_items.insert(other_items_count, note_row)
            if page_break_idx is not None:
                page_break_idx += 1
        else:
            sheet1_items.append(note_row)

    setup_a4_sheet(ws1, "Ürün Toplama Listesi", sheet1_items, page_break_after_idx=page_break_idx)

    # Sheet 2: BEKO, GRUNDIG & LENOVO Sekmesi (Eger Beko, Grundig veya Lenovo siparisi varsa)
    if beko_grundig_orders and not filter_beko:
        ws_bg = wb_main.create_sheet(title="BEKO, GRUNDIG & LENOVO")
        setup_a4_sheet(ws_bg, "BEKO, GRUNDIG & LENOVO", beko_grundig_orders)

    # Sheet 3+: Diger Marka Sekmeleri (Tefal, Babyliss, Philips vb.)
    used_sheet_titles = set()
    for brand in sorted_other_keys:
        sheet_title = safe_sheet_title(brand, used_sheet_titles)
        ws_b = wb_main.create_sheet(title=sheet_title)
        setup_a4_sheet(ws_b, sheet_title, brand_orders[brand])

    # 2. Beko, Grundig & Lenovo Workbook (Ayri dosya)
    wb_beko = None
    if beko_orders:
        wb_beko = openpyxl.Workbook()
        ws_beko = wb_beko.active
        setup_a4_sheet(ws_beko, "Beko Grundig Lenovo Listesi", beko_orders)

    brand_stats = []
    if beko_grundig_orders and not filter_beko:
        brand_stats.append({
            'brand': 'BEKO, GRUNDIG & LENOVO',
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

    # "Diğer Markalar" KPI'ları SADECE 1. sayfadaki kalemlerden hesaplanır.
    # grouped_items, filter_beko=False iken Beko/Grundig/Lenovo kalemlerini de
    # içerdiği için tamamının toplamı kullanılırsa Beko adedi iki kez sayılırdı (H-04).
    other_page_items = grouped_items[:other_items_count]

    summary_filename = source_filenames[0] if len(source_filenames) == 1 else f"{len(source_filenames)} Dosya Birleştirildi"

    return {
        'source_filename': summary_filename,
        'source_files': source_filenames,
        'total_files': len(source_filenames),
        'total_orders': len(all_raw_rows),
        'excluded_rows': len(excluded_rows),
        'excluded_qty': sum(x['qty'] for x in excluded_rows),
        'non_beko_count': len(other_page_items),
        'non_beko_qty': sum(x['qty'] for x in other_page_items),
        'other_items_count': other_items_count,
        'beko_count': len(beko_orders),
        'beko_qty': sum(x['qty'] for x in beko_orders),
        'filter_beko': filter_beko,
        'brands': brand_stats,
        'grouped_items': grouped_items,
        'beko_items': beko_orders,
        'custom_note': (custom_note or '').strip(),
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


def parse_and_process_multiple_files(files_or_paths, filter_beko=False, target_date=None, custom_note=""):
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

    result = consolidate_and_build(all_raw_rows, source_filenames, filter_beko=filter_beko, custom_note=custom_note)
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


def process_in_memory(files_or_paths, filter_beko=False, custom_note=""):
    """
    BULUT YOLU: hiçbir disk yazma işlemi yapmaz. Excel çıktıları bellekte üretilir
    ve byte olarak döner. Vercel gibi kalıcı diski olmayan ortamlar içindir.
    """
    temp_dir = CLOUD_TMP_DIR if IS_CLOUD else None
    all_raw_rows, source_filenames = read_rows_from_files(files_or_paths, temp_dir=temp_dir)

    result = consolidate_and_build(all_raw_rows, source_filenames, filter_beko=filter_beko, custom_note=custom_note)
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

def parse_and_process_file(file_path_or_bytes, custom_filename=None, filter_beko=False, target_date=None, custom_note=""):
    """
    Backward-compatible wrapper for single file or list of files.
    """
    if isinstance(file_path_or_bytes, (list, tuple)):
        return parse_and_process_multiple_files(file_path_or_bytes, filter_beko=filter_beko, target_date=target_date, custom_note=custom_note)
    return parse_and_process_multiple_files([file_path_or_bytes], filter_beko=filter_beko, target_date=target_date, custom_note=custom_note)
