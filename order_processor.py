import sys
import os
import unicodedata
from core_engine import parse_and_process_file, get_base_dirs, get_latest_download_file

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

def process_excel_orders(source_file_path, base_output_dirs=None):
    """
    Sipariş Excel dosyasını (.xls veya .xlsx) okur, pazaryeri ve Beko ayrımını yapar,
    A4 formatında yazdırılabilir toplama listelerini oluşturur.
    """
    return parse_and_process_file(source_file_path, filter_beko=False)

if __name__ == '__main__':
    source = get_latest_download_file()
    if not source:
        source = r'C:\Users\ugurk\Downloads\Siparişler (1).xlsx'
    print(f"İşleniyor: {source}")
    res = process_excel_orders(source)
    print("\nİşlem Başarıyla Tamamlandı:")
    print(f"  Toplam Sipariş Satırı: {res['total_orders']}")
    print(f"  Ana Liste: {res['non_beko_count']} çeşit ({res['non_beko_qty']} adet)")
    print(f"  Beko Listesi: {res['beko_count']} çeşit ({res['beko_qty']} adet)")
    print(f"  Klasör: {res['folder_path']}")
