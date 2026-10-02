/**
 * Toplu Barkod & 10x10cm Termal Kargo Etiket Sistemi
 * Dinamik JsBarcode SVG Render, Çoklu Yazdırma ve PDF Yönetimi
 */

window.barcodeAppState = {
    items: [],
    headers: [],
    mapping: {},
    filename: '',
    printMode: 'quantity',
    searchQuery: '',
    options: {
        barcode_height: 24,
        font_size: 12,
        show_border: false
    }
};

document.addEventListener('DOMContentLoaded', () => {
    loadSettingsFromStorage();
    initEventListeners();
});

function loadSettingsFromStorage() {
    try {
        const saved = localStorage.getItem('toplu_kargo_barkod_settings');
        if (saved) {
            window.barcodeAppState.options = { ...window.barcodeAppState.options, ...JSON.parse(saved) };
            applySettingsToUI();
        }
    } catch (e) {
        console.error('Ayarlar yuklenemedi', e);
    }
}

function saveSettingsToStorage() {
    try {
        localStorage.setItem('toplu_kargo_barkod_settings', JSON.stringify(window.barcodeAppState.options));
    } catch (e) {
        console.error('Ayarlar kaydedilemedi', e);
    }
}

function applySettingsToUI() {
    document.getElementById('setting-barcode-height').value = window.barcodeAppState.options.barcode_height || 24;
    document.getElementById('setting-font-size').value = window.barcodeAppState.options.font_size || 12;
    document.getElementById('setting-show-border').checked = !!window.barcodeAppState.options.show_border;
}

function initEventListeners() {
    const dropArea = document.getElementById('drop-area');
    const fileInput = document.getElementById('barcode-file-input');

    dropArea.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
            uploadFile(e.target.files[0]);
        }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropArea.classList.add('drag-over');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropArea.classList.remove('drag-over');
        });
    });

    dropArea.addEventListener('drop', (e) => {
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            uploadFile(e.dataTransfer.files[0]);
        }
    });

    document.getElementById('btn-sample-load').addEventListener('click', loadSampleData);
    document.getElementById('btn-load-downloads').addEventListener('click', loadRecentDownloads);
    document.getElementById('btn-apply-remap').addEventListener('click', applyRemapping);

    const btnToggleMapping = document.getElementById('btn-toggle-mapping');
    if (btnToggleMapping) {
        btnToggleMapping.addEventListener('click', toggleMappingBar);
    }

    const btnCloseMapping = document.getElementById('btn-close-mapping');
    if (btnCloseMapping) {
        btnCloseMapping.addEventListener('click', () => setMappingBarVisibility(false));
    }

    document.getElementById('check-select-all').addEventListener('change', (e) => {
        const isChecked = e.target.checked;
        window.barcodeAppState.items.forEach(item => item.selected = isChecked);
        renderLabels();
        updateCounts();
    });

    document.getElementById('input-search').addEventListener('input', (e) => {
        window.barcodeAppState.searchQuery = e.target.value.toLowerCase().trim();
        renderLabels();
    });

    document.getElementById('btn-mode-qty').addEventListener('click', () => {
        window.barcodeAppState.printMode = 'quantity';
        document.getElementById('btn-mode-qty').classList.add('active');
        document.getElementById('btn-mode-single').classList.remove('active');
        renderLabels();
        updateCounts();
    });

    document.getElementById('btn-mode-single').addEventListener('click', () => {
        window.barcodeAppState.printMode = 'single';
        document.getElementById('btn-mode-single').classList.add('active');
        document.getElementById('btn-mode-qty').classList.remove('active');
        renderLabels();
        updateCounts();
    });

    document.getElementById('btn-print-bulk').addEventListener('click', triggerBulkPrint);
    document.getElementById('btn-download-pdf').addEventListener('click', downloadBulkPDF);

    // Ayarlar Modalı
    const modalSettings = document.getElementById('modal-settings');
    document.getElementById('btn-open-settings').addEventListener('click', () => {
        applySettingsToUI();
        modalSettings.style.display = 'flex';
    });
    document.getElementById('btn-close-settings').addEventListener('click', () => modalSettings.style.display = 'none');
    document.getElementById('btn-cancel-settings').addEventListener('click', () => modalSettings.style.display = 'none');
    document.getElementById('btn-save-settings').addEventListener('click', () => {
        window.barcodeAppState.options.barcode_height = parseInt(document.getElementById('setting-barcode-height').value) || 24;
        window.barcodeAppState.options.font_size = parseInt(document.getElementById('setting-font-size').value) || 12;
        window.barcodeAppState.options.show_border = document.getElementById('setting-show-border').checked;
        saveSettingsToStorage();
        modalSettings.style.display = 'none';
        renderLabels();
    });

    // Manuel Etiket Modalı
    const modalManual = document.getElementById('modal-manual');
    document.getElementById('btn-manual-add').addEventListener('click', () => {
        document.getElementById('manual-cust').value = '';
        document.getElementById('manual-addr').value = '';
        document.getElementById('manual-dist').value = '';
        document.getElementById('manual-city').value = '';
        document.getElementById('manual-phone').value = '';
        document.getElementById('manual-barcode').value = '';
        document.getElementById('manual-name').value = '';
        document.getElementById('manual-qty').value = '1';
        modalManual.style.display = 'flex';
    });
    document.getElementById('btn-close-manual').addEventListener('click', () => modalManual.style.display = 'none');
    document.getElementById('btn-cancel-manual').addEventListener('click', () => modalManual.style.display = 'none');
    document.getElementById('btn-save-manual').addEventListener('click', () => {
        const cust = document.getElementById('manual-cust').value.trim();
        const addr = document.getElementById('manual-addr').value.trim();
        const dist = document.getElementById('manual-dist').value.trim();
        const city = document.getElementById('manual-city').value.trim();
        const phone = document.getElementById('manual-phone').value.trim();
        const barcode = document.getElementById('manual-barcode').value.trim();
        const name = document.getElementById('manual-name').value.trim();
        const qty = parseInt(document.getElementById('manual-qty').value) || 1;

        if (!barcode && !name) {
            alert('Lütfen en azından bir Barkod veya Ürün Adı girin.');
            return;
        }

        let city_line = '';
        if (dist && city) city_line = `(${dist}) ${city}`;
        else city_line = city || dist;

        const newItem = {
            id: Date.now(),
            customer: cust,
            address: addr,
            district: dist,
            city: city,
            city_line: city_line,
            phone: phone,
            barcode: barcode,
            name: name || 'İsimsiz Ürün',
            quantity: qty,
            selected: true
        };

        window.barcodeAppState.items.unshift(newItem);
        modalManual.style.display = 'none';
        
        document.getElementById('toolbar').style.display = 'flex';
        renderLabels();
        updateCounts();
    });
}

function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    showLoading('Dosya yükleniyor ve kargo sütunları taranıyor...');

    fetch('/api/barcode/upload', {
        method: 'POST',
        body: formData
    })
    .then(r => r.json())
    .then(res => {
        hideLoading();
        if (!res.success) {
            alert('Hata: ' + res.message);
            return;
        }
        handleDataLoaded(res);
    })
    .catch(err => {
        hideLoading();
        alert('Sunucu hatası: ' + err.message);
    });
}

function loadSampleData() {
    showLoading('Örnek veriler yükleniyor...');
    fetch('/api/barcode/load-sample')
    .then(r => r.json())
    .then(res => {
        hideLoading();
        if (res.success) {
            handleDataLoaded(res);
        } else {
            alert('Hata: ' + res.message);
        }
    })
    .catch(err => {
        hideLoading();
        alert('Hata: ' + err.message);
    });
}

function loadRecentDownloads() {
    showLoading('İndirilenler klasörü taranıyor...');
    fetch('/api/barcode/load-downloads')
    .then(r => r.json())
    .then(res => {
        hideLoading();
        if (res.success) {
            handleDataLoaded(res);
        } else {
            alert('Bilgi: ' + res.message);
        }
    })
    .catch(err => {
        hideLoading();
        alert('Hata: ' + err.message);
    });
}

function handleDataLoaded(res) {
    window.barcodeAppState.items = res.items || [];
    window.barcodeAppState.headers = res.headers || [];
    window.barcodeAppState.mapping = res.mapping || {};
    window.barcodeAppState.filename = res.filename || 'Dosya';

    document.getElementById('active-filename').textContent = window.barcodeAppState.filename;
    populateColumnSelects();

    // Sütun eşleşmesi varsayılan olarak gizli kalır
    setMappingBarVisibility(false);
    document.getElementById('toolbar').style.display = 'flex';

    renderLabels();
    updateCounts();
}

function populateColumnSelects() {
    const headers = window.barcodeAppState.headers;
    const mapping = window.barcodeAppState.mapping;

    const fillSelect = (selectId, activeIdx) => {
        const sel = document.getElementById(selectId);
        sel.innerHTML = '<option value="">(Seçilmedi)</option>';
        headers.forEach((h, idx) => {
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = `${idx + 1}. ${h}`;
            if (activeIdx !== null && activeIdx !== undefined && activeIdx === idx) {
                opt.selected = true;
            }
            sel.appendChild(opt);
        });
    };

    fillSelect('select-cust-col', mapping.cust_idx);
    fillSelect('select-addr-col', mapping.addr_idx);
    fillSelect('select-dist-col', mapping.dist_idx);
    fillSelect('select-city-col', mapping.city_idx);
    fillSelect('select-phone-col', mapping.phone_idx);
    fillSelect('select-barcode-col', mapping.barcode_idx);
    fillSelect('select-name-col', mapping.name_idx);
    fillSelect('select-qty-col', mapping.qty_idx);
}

function applyRemapping() {
    const cVal = document.getElementById('select-cust-col').value;
    const aVal = document.getElementById('select-addr-col').value;
    const dVal = document.getElementById('select-dist-col').value;
    const ctVal = document.getElementById('select-city-col').value;
    const pVal = document.getElementById('select-phone-col').value;
    const bVal = document.getElementById('select-barcode-col').value;
    const nVal = document.getElementById('select-name-col').value;
    const qVal = document.getElementById('select-qty-col').value;

    showLoading('Sütunlar yeniden eşleştiriliyor...');

    fetch('/api/barcode/remap', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            cust_idx: cVal !== '' ? parseInt(cVal) : null,
            addr_idx: aVal !== '' ? parseInt(aVal) : null,
            dist_idx: dVal !== '' ? parseInt(dVal) : null,
            city_idx: ctVal !== '' ? parseInt(ctVal) : null,
            phone_idx: pVal !== '' ? parseInt(pVal) : null,
            barcode_idx: bVal !== '' ? parseInt(bVal) : null,
            name_idx: nVal !== '' ? parseInt(nVal) : null,
            qty_idx: qVal !== '' ? parseInt(qVal) : null
        })
    })
    .then(r => r.json())
    .then(res => {
        hideLoading();
        if (res.success) {
            window.barcodeAppState.items = res.items || [];
            window.barcodeAppState.mapping = res.mapping || {};
            renderLabels();
            updateCounts();
            setMappingBarVisibility(false);
        } else {
            alert('Hata: ' + res.message);
        }
    })
    .catch(err => {
        hideLoading();
        alert('Hata: ' + err.message);
    });
}

function updateCounts() {
    const selectedItems = window.barcodeAppState.items.filter(i => i.selected);
    const totalSelected = selectedItems.length;
    const totalItems = window.barcodeAppState.items.length;

    document.getElementById('count-selected').textContent = totalSelected;
    document.getElementById('count-total-items').textContent = totalItems;

    let totalPrintLabels = 0;
    if (window.barcodeAppState.printMode === 'quantity') {
        totalPrintLabels = selectedItems.reduce((acc, i) => acc + (i.quantity || 1), 0);
    } else {
        totalPrintLabels = totalSelected;
    }

    document.getElementById('badge-total-labels').textContent = `${totalPrintLabels} Etiket`;
}

// 10cm x 10cm Kargo Etiketlerini Ekrana Çiz (Kullanıcı Şablonu)
function renderLabels() {
    const container = document.getElementById('labels-container');
    container.innerHTML = '';

    const query = window.barcodeAppState.searchQuery;
    const filtered = window.barcodeAppState.items.filter(item => {
        if (!query) return true;
        const c = (item.customer || '').toLowerCase();
        const b = (item.barcode || '').toLowerCase();
        const n = (item.name || '').toLowerCase();
        const ct = (item.city || '').toLowerCase();
        return c.includes(query) || b.includes(query) || n.includes(query) || ct.includes(query);
    });

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">🔍</div>
                <h3>Eşleşen sipariş etiketi bulunamadı</h3>
                <p>Arama filtrenizi temizleyin veya farklı bir arama yapın.</p>
            </div>
        `;
        return;
    }

    // Sadece aynı sipariş numaralı etiketlerde ardışık oran hesapla (1/3, 2/3, 3/3 ...)
    const orderTotalsScreen = {};
    filtered.forEach(i => {
        const b = String(i.barcode || '').trim();
        if (b) {
            const mult = (window.barcodeAppState.printMode === 'quantity' ? (i.quantity || 1) : 1);
            orderTotalsScreen[b] = (orderTotalsScreen[b] || 0) + mult;
        }
    });

    const orderCountersScreen = {};
    filtered.forEach(item => {
        const b = String(item.barcode || '').trim();
        const tot = orderTotalsScreen[b] || 0;
        if (tot > 1) {
            const q = (window.barcodeAppState.printMode === 'quantity' ? (item.quantity || 1) : 1);
            const startCnt = (orderCountersScreen[b] || 0) + 1;
            const endCnt = startCnt + q - 1;
            orderCountersScreen[b] = endCnt;

            if (startCnt === endCnt) {
                item.package_ratio = `${startCnt}/${tot}`;
            } else {
                item.package_ratio = `${startCnt}-${endCnt}/${tot}`;
            }
        } else {
            item.package_ratio = '';
        }
        const cardWrapper = document.createElement('div');
        cardWrapper.className = 'label-card-wrapper';

        // Üst Seçim ve Yazdırma Araçları
        const actionsBar = document.createElement('div');
        actionsBar.className = 'label-card-actions';

        const checkLabel = document.createElement('label');
        checkLabel.className = 'checkbox-label';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = !!item.selected;
        checkbox.addEventListener('change', (e) => {
            item.selected = e.target.checked;
            labelCard.classList.toggle('deselected', !item.selected);
            updateCounts();
        });
        checkLabel.appendChild(checkbox);
        checkLabel.appendChild(document.createTextNode(item.quantity > 1 ? `${item.quantity} Adet` : '1 Adet'));

        const singlePrintBtn = document.createElement('button');
        singlePrintBtn.className = 'btn btn-secondary btn-sm';
        singlePrintBtn.innerHTML = '🖨️';
        singlePrintBtn.title = 'Sadece bu etiketi yazdır';
        singlePrintBtn.addEventListener('click', () => printSingleLabel(item));

        actionsBar.appendChild(checkLabel);
        actionsBar.appendChild(singlePrintBtn);

        // 10cm x 10cm Kare Kart (Kullanıcının Görseliyle Birebir Şablon)
        const labelCard = document.createElement('div');
        labelCard.className = `label-card ${item.selected ? '' : 'deselected'}`;
        if (window.barcodeAppState.options.show_border) {
            labelCard.style.border = '1px dashed #94A3B8';
        }

        // 1. Üst Kısım: Müşteri Bilgileri
        const headerInfo = document.createElement('div');
        headerInfo.className = 'label-header-info';

        if (item.customer) {
            const custEl = document.createElement('div');
            custEl.className = 'label-customer-name';
            custEl.textContent = item.customer;
            headerInfo.appendChild(custEl);
        }

        if (item.address) {
            const addrEl = document.createElement('div');
            addrEl.className = 'label-customer-addr';
            addrEl.textContent = item.address;
            headerInfo.appendChild(addrEl);
        }

        if (item.city_line) {
            const cityEl = document.createElement('div');
            cityEl.className = 'label-customer-city';
            cityEl.textContent = item.city_line;
            headerInfo.appendChild(cityEl);
        }

        if (item.phone) {
            const phoneEl = document.createElement('div');
            phoneEl.className = 'label-customer-phone';
            phoneEl.textContent = item.phone;
            headerInfo.appendChild(phoneEl);
        }

        labelCard.appendChild(headerInfo);

        // 2. Orta Kısım: Barkod (Code128) ve Sağ Tarafta 1/2, 2/2
        const bcBox = document.createElement('div');
        bcBox.className = 'label-barcode-container';

        const bcWrapper = document.createElement('div');
        bcWrapper.className = 'barcode-svg-wrapper';

        const svgEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        bcWrapper.appendChild(svgEl);

        if (item.package_ratio) {
            const ratioEl = document.createElement('div');
            ratioEl.className = 'label-package-ratio';
            ratioEl.textContent = item.package_ratio;
            bcWrapper.appendChild(ratioEl);
        }

        bcBox.appendChild(bcWrapper);
        labelCard.appendChild(bcBox);

        try {
            if (item.barcode) {
                JsBarcode(svgEl, item.barcode, {
                    format: "CODE128",
                    lineColor: "#000",
                    width: 2.2,
                    height: window.barcodeAppState.options.barcode_height || 24,
                    displayValue: true,
                    fontSize: 14,
                    textMargin: 3,
                    font: "sans-serif"
                });
            } else {
                svgEl.outerHTML = '<div style="font-size: 11px; font-weight: bold; color: #DC2626;">[ BARKOD YOK ]</div>';
            }
        } catch (e) {
            svgEl.outerHTML = `<div style="font-size: 11px; font-weight: bold;">Barkod: ${item.barcode}</div>`;
        }

        // 3. Alt Kısım (Kırmızıyla İşaretlenen Yer): Ürün Adı
        const prodNameEl = document.createElement('div');
        prodNameEl.className = 'label-product-name-bottom';
        prodNameEl.style.fontSize = `${window.barcodeAppState.options.font_size || 12}px`;
        prodNameEl.textContent = item.name;
        labelCard.appendChild(prodNameEl);

        cardWrapper.appendChild(actionsBar);
        cardWrapper.appendChild(labelCard);
        container.appendChild(cardWrapper);
    });
}

function printSingleLabel(item) {
    prepareAndPrintDOM([{ ...item, selected: true }]);
}

function triggerBulkPrint() {
    const selectedItems = window.barcodeAppState.items.filter(i => i.selected);
    if (selectedItems.length === 0) {
        alert('Lütfen yazdırılacak en az bir sipariş seçin.');
        return;
    }
    prepareAndPrintDOM(selectedItems);
}

function prepareAndPrintDOM(itemsToPrint) {
    const printContainer = document.getElementById('print-sheet-container');
    printContainer.innerHTML = '';

    const printMode = window.barcodeAppState.printMode;
    const barcodeHeight = window.barcodeAppState.options.barcode_height || 24;
    const fontSize = window.barcodeAppState.options.font_size || 12;
    const showBorder = window.barcodeAppState.options.show_border;

    // Önce adetlere göre etiketleri aç
    const expandedPrint = [];
    itemsToPrint.forEach(item => {
        const count = printMode === 'quantity' ? (item.quantity || 1) : 1;
        for (let i = 1; i <= count; i++) {
            expandedPrint.push({ ...item });
        }
    });

    // Sadece aynı sipariş numaralı etiket sayısı > 1 ise hesapla
    const orderTotals = {};
    expandedPrint.forEach(p => {
        const b = String(p.barcode || '').trim();
        if (b) orderTotals[b] = (orderTotals[b] || 0) + 1;
    });

    const orderCounters = {};
    expandedPrint.forEach(p => {
        const b = String(p.barcode || '').trim();
        const tot = orderTotals[b] || 0;
        if (tot > 1) {
            const cnt = (orderCounters[b] || 0) + 1;
            orderCounters[b] = cnt;
            p.package_ratio = `${cnt}/${tot}`;
        } else {
            p.package_ratio = '';
        }
    });

    expandedPrint.forEach(item => {
        const pageEl = document.createElement('div');
            pageEl.className = 'thermal-label-page';
            if (showBorder) {
                pageEl.style.border = '1px dashed #000';
            }

            // 1. Müşteri Bilgileri
            const headerInfo = document.createElement('div');
            headerInfo.className = 'thermal-header-info';

            if (item.customer) {
                const cEl = document.createElement('div');
                cEl.className = 'thermal-customer-name';
                cEl.textContent = item.customer;
                headerInfo.appendChild(cEl);
            }

            if (item.address) {
                const aEl = document.createElement('div');
                aEl.className = 'thermal-customer-addr';
                aEl.textContent = item.address;
                headerInfo.appendChild(aEl);
            }

            if (item.city_line) {
                const ctEl = document.createElement('div');
                ctEl.className = 'thermal-customer-city';
                ctEl.textContent = item.city_line;
                headerInfo.appendChild(ctEl);
            }

            if (item.phone) {
                const pEl = document.createElement('div');
                pEl.className = 'thermal-customer-phone';
                pEl.textContent = item.phone;
                headerInfo.appendChild(pEl);
            }

            pageEl.appendChild(headerInfo);

            // 2. Barkod SVG ve Sağ Tarafta 1/2, 2/2
            const bcBox = document.createElement('div');
            bcBox.className = 'thermal-barcode-box';

            const bcWrapper = document.createElement('div');
            bcWrapper.className = 'thermal-barcode-wrapper';

            const svgEl = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
            bcWrapper.appendChild(svgEl);

            if (item.package_ratio) {
                const ratioEl = document.createElement('div');
                ratioEl.className = 'thermal-package-ratio';
                ratioEl.textContent = item.package_ratio;
                bcWrapper.appendChild(ratioEl);
            }

            bcBox.appendChild(bcWrapper);
            pageEl.appendChild(bcBox);

            try {
                if (item.barcode) {
                    JsBarcode(svgEl, item.barcode, {
                        format: "CODE128",
                        lineColor: "#000",
                        width: 2.2,
                        height: barcodeHeight,
                        displayValue: true,
                        fontSize: 13,
                        textMargin: 2,
                        font: "sans-serif"
                    });
                }
            } catch (e) {
                console.error(e);
            }

            // 3. Ürün Adı (Kırmızı işaretli yer)
            const prodEl = document.createElement('div');
            prodEl.className = 'thermal-prod-name-bottom';
            prodEl.style.fontSize = `${fontSize}pt`;
            prodEl.textContent = item.name;
            pageEl.appendChild(prodEl);

            printContainer.appendChild(pageEl);
    });

    setTimeout(() => {
        
    document.body.classList.add('active-tab-barkod');
    document.body.classList.remove('active-tab-siparis');
    setTimeout(() => {
        window.print();
    }, 150);

        setTimeout(() => {
            if (printContainer) printContainer.innerHTML = '';
        }, 1500);
    }, 100);
}

// Yazdırma penceresi kapandığında veya iptal edildiğinde ekran temizlensin
window.addEventListener('afterprint', () => {
    const printContainer = document.getElementById('print-sheet-container');
    if (printContainer) printContainer.innerHTML = '';
});

function downloadBulkPDF() {
    const selectedItems = window.barcodeAppState.items.filter(i => i.selected);
    if (selectedItems.length === 0) {
        alert('Lütfen indirilecek en az bir sipariş seçin.');
        return;
    }

    showLoading('10cm x 10cm formatında PDF oluşturuluyor...');

    const payload = {
        items: selectedItems,
        options: {
            ...window.barcodeAppState.options,
            print_count_mode: window.barcodeAppState.printMode
        }
    };

    fetch('/api/barcode/generate-pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    })
    .then(response => {
        if (!response.ok) throw new Error('PDF üretimi başarısız oldu.');
        return response.blob();
    })
    .then(blob => {
        hideLoading();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = `Kargo_Barkod_10x10_${new Date().toISOString().slice(0, 10)}.pdf`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
    })
    .catch(err => {
        hideLoading();
        alert('Hata: ' + err.message);
    });
}

function showLoading(msg) {
    let loader = document.getElementById('global-loader');
    if (!loader) {
        loader = document.createElement('div');
        loader.id = 'global-loader';
        loader.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.7);z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#FFF;font-weight:600;font-size:16px;';
        loader.innerHTML = '<div style="font-size:40px;margin-bottom:12px;animation:spin 1s linear infinite;">⏳</div><div id="loader-msg">Yükleniyor...</div>';
        document.body.appendChild(loader);
    }
    document.getElementById('loader-msg').textContent = msg || 'Lütfen bekleyin...';
    loader.style.display = 'flex';
}

function hideLoading() {
    const loader = document.getElementById('global-loader');
    if (loader) loader.style.display = 'none';
}

// Sütun Eşleşmesi Gizle / Aç Yönetimi
function setMappingBarVisibility(show) {
    const bar = document.getElementById('mapping-bar');
    const btn = document.getElementById('btn-toggle-mapping');
    const arrow = document.getElementById('mapping-btn-arrow');
    const label = document.getElementById('mapping-btn-label');

    if (!bar) return;

    if (show) {
        bar.style.display = 'block';
        if (btn) btn.classList.add('active');
        if (arrow) arrow.textContent = '▴';
        if (label) label.textContent = 'Sütunları Gizle';
    } else {
        bar.style.display = 'none';
        if (btn) btn.classList.remove('active');
        if (arrow) arrow.textContent = '▾';
        if (label) label.textContent = 'Sütun Eşleşmesi';
    }
}

function toggleMappingBar() {
    const bar = document.getElementById('mapping-bar');
    const isCurrentlyOpen = bar && bar.style.display !== 'none';
    setMappingBarVisibility(!isCurrentlyOpen);
}
