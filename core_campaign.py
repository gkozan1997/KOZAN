# -*- coding: utf-8 -*-
"""
Oliz Kampanya & İndirim Analiz Motoru
Beko & Arçelik Oliz Kampanya Excel Dosyalarını Çözümleme ve 1, 2, 3, 4 Ürünlü Analiz Sistemi
"""

import os
import re
import openpyxl
from datetime import datetime

class CampaignEngine:
    def __init__(self, excel_path=None):
        self.excel_path = excel_path
        self.tekil_kampanyalar = {}       # sku -> item dict
        self.tekil_by_name = {}           # normalized_name -> item dict
        self.toptan_kuponlar = {}         # sku -> item dict
        self.toptan_by_name = {}          # normalized_name -> item dict
        self.paket_kampanyalar = []       # list of package rules
        self.package_members = {}         # sheet_name -> { role_name: [sku_list] }
        self.all_products = {}            # sku -> basic info {sku, name, brand, group}
        self.file_info = {}
        
        if excel_path and os.path.exists(excel_path):
            self.load_from_excel(excel_path)

    @staticmethod
    def normalize_key(val):
        if val is None:
            return ""
        s = str(val).strip().upper()
        # Clean multiple spaces
        s = re.sub(r'\s+', ' ', s)
        return s

    @staticmethod
    def clean_sku(val):
        if val is None:
            return ""
        s = str(val).strip()
        # If float like 7295520276.0 -> 7295520276
        if '.' in s:
            try:
                s = str(int(float(s)))
            except:
                pass
        return s

    @staticmethod
    def to_float(val):
        if val is None or val == "":
            return 0.0
        try:
            return float(val)
        except:
            # try removing dots or comma
            s = str(val).replace('.', '').replace(',', '.')
            try:
                return float(s)
            except:
                return 0.0

    def load_from_excel(self, file_path):
        self.excel_path = file_path
        self.tekil_kampanyalar.clear()
        self.tekil_by_name.clear()
        self.toptan_kuponlar.clear()
        self.toptan_by_name.clear()
        self.paket_kampanyalar.clear()
        self.package_members.clear()
        self.all_products.clear()

        wb = openpyxl.load_workbook(file_path, data_only=True)
        self.file_info = {
            "filename": os.path.basename(file_path),
            "sheets": wb.sheetnames,
            "load_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        # 1. TEKİL KAMPANYALAR
        if "TEKİL KAMPANYALAR" in wb.sheetnames or "TEKL KAMPANYALAR" in wb.sheetnames:
            sheet_name = "TEKİL KAMPANYALAR" if "TEKİL KAMPANYALAR" in wb.sheetnames else "TEKL KAMPANYALAR"
            ws = wb[sheet_name]
            # headers in row 1
            for r in range(2, ws.max_row + 1):
                brand = str(ws.cell(r, 1).value or '').strip()
                group = str(ws.cell(r, 2).value or '').strip()
                raw_sku = ws.cell(r, 3).value
                raw_desc = ws.cell(r, 4).value
                if not raw_sku and not raw_desc:
                    continue
                
                sku = self.clean_sku(raw_sku)
                name = str(raw_desc or '').strip()
                discount = self.to_float(ws.cell(r, 5).value)
                arcelik_support = self.to_float(ws.cell(r, 6).value)
                bayi_support = max(0.0, discount - arcelik_support)
                start_date = ws.cell(r, 7).value
                end_date = ws.cell(r, 8).value
                camp_code = str(ws.cell(r, 11).value or '').strip()
                camp_desc = str(ws.cell(r, 12).value or '').strip()
                notes = str(ws.cell(r, 13).value or '').strip() if ws.max_column >= 13 else ''

                item = {
                    "sku": sku,
                    "name": name,
                    "brand": brand,
                    "group": group,
                    "discount": discount,
                    "arcelik_support": arcelik_support,
                    "bayi_support": bayi_support,
                    "camp_code": camp_code,
                    "camp_desc": camp_desc,
                    "start_date": str(start_date)[:10] if start_date else "",
                    "end_date": str(end_date)[:10] if end_date else "",
                    "notes": notes
                }
                if sku:
                    self.tekil_kampanyalar[sku] = item
                    self.all_products[sku] = {"sku": sku, "name": name, "brand": brand, "group": group}
                if name:
                    norm_n = self.normalize_key(name)
                    self.tekil_by_name[norm_n] = item

        # 2. TOPTAN KUPON KAMPANYALARI
        if "TOPTAN KUPON KAMPANYALARI" in wb.sheetnames:
            ws = wb["TOPTAN KUPON KAMPANYALARI"]
            for r in range(2, ws.max_row + 1):
                raw_sku = ws.cell(r, 3).value
                raw_desc = ws.cell(r, 4).value
                if not raw_sku and not raw_desc:
                    continue
                sku = self.clean_sku(raw_sku)
                name = str(raw_desc or '').strip()
                brand = str(ws.cell(r, 1).value or '').strip()
                group = str(ws.cell(r, 2).value or '').strip()
                discount = self.to_float(ws.cell(r, 5).value)
                coupon_val = self.to_float(ws.cell(r, 6).value)
                target_group = str(ws.cell(r, 7).value or '').strip()
                camp_code = str(ws.cell(r, 12).value or '').strip()
                camp_desc = str(ws.cell(r, 13).value or '').strip()

                item = {
                    "sku": sku,
                    "name": name,
                    "brand": brand,
                    "group": group,
                    "discount": discount,
                    "coupon_val": coupon_val,
                    "target_group": target_group,
                    "camp_code": camp_code,
                    "camp_desc": camp_desc
                }
                if sku:
                    self.toptan_kuponlar[sku] = item
                    if sku not in self.all_products:
                        self.all_products[sku] = {"sku": sku, "name": name, "brand": brand, "group": group}
                if name:
                    self.toptan_by_name[self.normalize_key(name)] = item

        # 3. PAKET KAMPANYALAR TABLOSU
        if "PAKET KAMPANYALAR" in wb.sheetnames:
            ws = wb["PAKET KAMPANYALAR"]
            for r in range(2, ws.max_row + 1):
                desc = ws.cell(r, 1).value
                if not desc or "bilgilendirme" in str(desc).lower():
                    continue
                camp_desc = str(desc).strip()
                start_date = ws.cell(r, 2).value
                end_date = ws.cell(r, 3).value
                bayi_share = self.to_float(ws.cell(r, 4).value)
                arcelik_share = self.to_float(ws.cell(r, 5).value)
                camp_code = str(ws.cell(r, 6).value or '').strip()
                camp_title = str(ws.cell(r, 7).value or '').strip()
                notes = str(ws.cell(r, 8).value or '').strip()

                total_disc = bayi_share + arcelik_share

                # Extract discount from text if total_disc is 0
                if total_disc == 0:
                    disc_match = re.search(r'([\d\.]+)\s*TL', camp_desc)
                    if disc_match:
                        total_disc = self.to_float(disc_match.group(1))

                pkg_item = {
                    "id": r,
                    "title": camp_title or camp_desc[:60],
                    "desc": camp_desc,
                    "camp_code": camp_code,
                    "total_discount": total_disc,
                    "bayi_share": bayi_share,
                    "arcelik_share": arcelik_share,
                    "start_date": str(start_date)[:10] if start_date else "",
                    "end_date": str(end_date)[:10] if end_date else "",
                    "notes": notes,
                    "sheet_target": None
                }
                self.paket_kampanyalar.append(pkg_item)

        # 4. PACKAGE DETAIL SHEETS (Sheets 4+)
        # Mapping rules from sheet names to package campaigns
        sheet_mapping = {
            "BE+KURUTUCU": 1161143,
            "ANK SET+KURUTUCU": 1161156,
            "TV+TEKL TKM": 1157227,
            "TV+TEKLİ TKM": 1157227,
            "65TV+32TV": 1157266,
            "65TV+KEA": 1157265,
            "85TV+43TV": 1153680,
            "TV+MD": 1154701,
            "TV+HAVADAR": 1154702,
            "PORTATF KLM+HAVADAR": 1157426,
            "PORTATİF KLM+HAVADAR": 1157426,
            "SALON TP KLM+HAVADAR10.699": 1157231,
            "SALON TİPİ KLM+HAVADAR10.699": 1157231,
            "SALON TP KLM+HAVADAR16.199": 1157425,
            "SALON TİPİ KLM+HAVADAR16.199": 1157425,
            "BUZDOLABI+SPRGE": 1161129,
            "BUZDOLABI+SÜPÜRGE": 1161129,
            "AMAIR+T": 1161140,
            "ÇAMAŞIR+ÜTÜ": 1161140,
            "ESPRESSO+BUZ MAK.": 1159844
        }

        for s_name in wb.sheetnames[3:]:
            ws = wb[s_name]
            roles = {}
            for c in range(1, ws.max_column + 1, 5):
                role_title = str(ws.cell(1, c).value or '').strip()
                if not role_title:
                    continue
                role_items = []
                for r in range(3, ws.max_row + 1):
                    raw_sku = ws.cell(r, c).value
                    if not raw_sku:
                        continue
                    sku = self.clean_sku(raw_sku)
                    name = str(ws.cell(r, c+1).value or '').strip() if c+1 <= ws.max_column else ''
                    brand = str(ws.cell(r, c+2).value or '').strip() if c+2 <= ws.max_column else ''
                    group = str(ws.cell(r, c+3).value or '').strip() if c+3 <= ws.max_column else ''

                    if sku and sku != 'None':
                        item_info = {"sku": sku, "name": name, "brand": brand, "group": group}
                        role_items.append(item_info)
                        if sku not in self.all_products:
                            self.all_products[sku] = item_info

                roles[role_title] = role_items

            self.package_members[s_name] = roles

            # Link package sheet with package campaign row
            camp_code_target = sheet_mapping.get(s_name)
            for pkg in self.paket_kampanyalar:
                if (camp_code_target and str(pkg["camp_code"]) == str(camp_code_target)) or (s_name.replace('', '').replace(' ', '') in pkg["title"].replace(' ', '')):
                    pkg["sheet_target"] = s_name
                    pkg["roles"] = list(roles.keys())
                    break

        print(f"[OK] Kampanya verisi yüklendi: {len(self.tekil_kampanyalar)} Tekil, {len(self.toptan_kuponlar)} Toptan, {len(self.paket_kampanyalar)} Paket.")
        return True

    def find_product(self, query):
        """SKU veya ürün adı ile arama yapar."""
        if not query:
            return None
        q = str(query).strip()
        q_sku = self.clean_sku(q)
        q_norm = self.normalize_key(q)

        # 1. Exact SKU match
        if q_sku in self.all_products:
            p = dict(self.all_products[q_sku])
            return p

        # 2. Match in Tekil / Toptan by name
        if q_norm in self.tekil_by_name:
            t = self.tekil_by_name[q_norm]
            return {"sku": t["sku"], "name": t["name"], "brand": t["brand"], "group": t["group"]}

        # 3. Substring match in all products
        for sku, p in self.all_products.items():
            if q_sku and (q_sku == sku or q_sku in sku):
                return dict(p)
            if q_norm and (q_norm == self.normalize_key(p["name"]) or q_norm in self.normalize_key(p["name"])):
                return dict(p)

        return None

    def search_products(self, query, limit=15):
        """Autocomplete için ürün listesi arar."""
        if not query:
            return []
        q = str(query).strip().lower()
        q_sku = self.clean_sku(q)
        results = []

        for sku, p in self.all_products.items():
            name = (p.get("name") or "").lower()
            group = (p.get("group") or "").lower()
            p_sku = str(sku).lower()

            if q_sku and q_sku in p_sku:
                results.append(p)
            elif q in name or q in group:
                results.append(p)

            if len(results) >= limit:
                break
        return results

    def analyze_single_product(self, sku_or_name):
        """Tek bir ürün için tekil kampanya, toptan kupon ve girebileceği paketleri bulur."""
        product = self.find_product(sku_or_name)
        if not product:
            return None

        sku = product["sku"]
        name = product["name"]
        norm_name = self.normalize_key(name)

        # Tekil kampanya kontrolü
        tekil = self.tekil_kampanyalar.get(sku)
        if not tekil and norm_name:
            tekil = self.tekil_by_name.get(norm_name)

        # Toptan kupon kontrolü
        toptan = self.toptan_kuponlar.get(sku)
        if not toptan and norm_name:
            toptan = self.toptan_by_name.get(norm_name)

        # Girebileceği paket kampanyalarını tara
        eligible_packages = []
        for sheet_name, roles in self.package_members.items():
            for role_name, member_list in roles.items():
                is_member = any(m["sku"] == sku or (norm_name and self.normalize_key(m["name"]) == norm_name) for m in member_list)
                if is_member:
                    # Find package campaign info
                    pkg_info = next((p for p in self.paket_kampanyalar if p.get("sheet_target") == sheet_name), None)
                    eligible_packages.append({
                        "sheet_name": sheet_name,
                        "role_name": role_name,
                        "pkg_title": pkg_info["title"] if pkg_info else sheet_name,
                        "pkg_code": pkg_info["camp_code"] if pkg_info else "",
                        "pkg_discount": pkg_info["total_discount"] if pkg_info else 0.0,
                        "bayi_share": pkg_info["bayi_share"] if pkg_info else 0.0,
                        "arcelik_share": pkg_info["arcelik_share"] if pkg_info else 0.0,
                        "desc": pkg_info["desc"] if pkg_info else ""
                    })

        return {
            "product": product,
            "tekil": tekil,
            "toptan": toptan,
            "eligible_packages": eligible_packages,
            "has_discount": bool(tekil or toptan or eligible_packages)
        }

    def analyze_bundle(self, products_list):
        """
        1, 2, 3, 4 ürün verildiğinde:
        - Her ürünün tekil analizini yapar (1., 2., 3., 4. ayrı hesap)
        - Ürünlerin birlikte oluşturduğu 2'li, 3'lü, 4'lü paket kampanyalarını tespit eder
        - En avantajlı kombinasyonu hesaplar
        """
        parsed_items = []
        for i, code in enumerate(products_list):
            if code and str(code).strip():
                analysis = self.analyze_single_product(code)
                if analysis:
                    analysis["slot_index"] = i + 1
                    parsed_items.append(analysis)
                else:
                    parsed_items.append({
                        "slot_index": i + 1,
                        "product": {"sku": str(code).strip(), "name": "Bilinmeyen Ürün / Kampanya Dışı", "brand": "-", "group": "-"},
                        "tekil": None,
                        "toptan": None,
                        "eligible_packages": [],
                        "has_discount": False
                    })

        # 1, 2, 3, 4 AYRI HESAPLAR
        individual_results = []
        total_single_discount = 0.0
        total_arcelik_single = 0.0
        total_bayi_single = 0.0
        total_coupon_val = 0.0

        for item in parsed_items:
            slot = item["slot_index"]
            prod = item["product"]
            tekil = item.get("tekil")
            toptan = item.get("toptan")

            disc = tekil["discount"] if tekil else 0.0
            arc = tekil["arcelik_support"] if tekil else 0.0
            bayi = tekil["bayi_support"] if tekil else 0.0
            coup = toptan["coupon_val"] if toptan else 0.0

            total_single_discount += disc
            total_arcelik_single += arc
            total_bayi_single += bayi
            total_coupon_val += coup

            individual_results.append({
                "slot": slot,
                "sku": prod["sku"],
                "name": prod["name"],
                "brand": prod.get("brand", ""),
                "group": prod.get("group", ""),
                "has_tekil": bool(tekil),
                "tekil_discount": disc,
                "arcelik_support": arc,
                "bayi_support": bayi,
                "camp_code": tekil["camp_code"] if tekil else "",
                "camp_desc": tekil["camp_desc"] if tekil else "",
                "has_toptan": bool(toptan),
                "coupon_val": coup,
                "target_group": toptan.get("target_group", "") if toptan else "",
                "eligible_packages_count": len(item.get("eligible_packages", []))
            })

        # PAKET KAMPANYASI TESPİTİ
        matched_packages = []

        # Ankastre 4'lü paket kontrolü (Fırın + Ocak + Davlumbaz + Kurutucu)
        for sheet_name, roles in self.package_members.items():
            role_names = list(roles.keys())
            matched_roles = {}
            for r_name, members in roles.items():
                member_skus = {m["sku"] for m in members}
                for item in parsed_items:
                    it_sku = item["product"]["sku"]
                    if it_sku in member_skus:
                        matched_roles[r_name] = item["product"]
                        break

            # If all roles in package sheet are satisfied!
            if len(matched_roles) == len(role_names) and len(role_names) > 0:
                pkg_info = next((p for p in self.paket_kampanyalar if p.get("sheet_target") == sheet_name), None)
                matched_packages.append({
                    "sheet_name": sheet_name,
                    "title": pkg_info["title"] if pkg_info else sheet_name,
                    "camp_code": pkg_info["camp_code"] if pkg_info else "",
                    "total_discount": pkg_info["total_discount"] if pkg_info else 0.0,
                    "bayi_share": pkg_info["bayi_share"] if pkg_info else 0.0,
                    "arcelik_share": pkg_info["arcelik_share"] if pkg_info else 0.0,
                    "desc": pkg_info["desc"] if pkg_info else "",
                    "matched_roles": matched_roles,
                    "required_items_count": len(role_names)
                })

        # EN AVANTAJLI SEÇENEK
        best_package = max(matched_packages, key=lambda x: x["total_discount"]) if matched_packages else None
        best_pkg_discount = best_package["total_discount"] if best_package else 0.0

        if best_pkg_discount > total_single_discount:
            recommendation = {
                "type": "PAKET",
                "message": f"Bu sepet için en avantajlı seçenek '{best_package['title']}' paket kampanyasıdır.",
                "total_discount": best_pkg_discount,
                "arcelik_share": best_package["arcelik_share"],
                "bayi_share": best_package["bayi_share"],
                "advantage_diff": best_pkg_discount - total_single_discount
            }
        elif total_single_discount > 0:
            recommendation = {
                "type": "TEKİL",
                "message": "Ürünlerin tekil kampanyalarından yararlanmak daha avantajlıdır.",
                "total_discount": total_single_discount,
                "arcelik_share": total_arcelik_single,
                "bayi_share": total_bayi_single,
                "advantage_diff": total_single_discount - best_pkg_discount
            }
        else:
            recommendation = {
                "type": "YOK",
                "message": "Girilen ürünler için aktif bir indirim kampanyası tespit edilemedi.",
                "total_discount": 0.0,
                "arcelik_share": 0.0,
                "bayi_share": 0.0,
                "advantage_diff": 0.0
            }

        return {
            "input_count": len(parsed_items),
            "individual_results": individual_results,
            "total_single_discount": total_single_discount,
            "total_arcelik_single": total_arcelik_single,
            "total_bayi_single": total_bayi_single,
            "total_coupon_val": total_coupon_val,
            "matched_packages": matched_packages,
            "best_package": best_package,
            "recommendation": recommendation
        }

    def get_stats(self):
        return {
            "tekil_count": len(self.tekil_kampanyalar),
            "toptan_count": len(self.toptan_kuponlar),
            "paket_count": len(self.paket_kampanyalar),
            "total_products": len(self.all_products),
            "file_info": self.file_info
        }

    def get_all_tekil(self):
        return list(self.tekil_kampanyalar.values())

    def get_all_toptan(self):
        return list(self.toptan_kuponlar.values())

    def get_all_paket(self):
        return self.paket_kampanyalar