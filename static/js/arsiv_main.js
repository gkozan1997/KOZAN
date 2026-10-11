window.handleFileInputChange = function(files) {
    if (files && files.length > 0) {
        handleFileUpload(files);
    }
};
let isUploadingExcel = false;
let lastUploadTimestamp = 0;

// --------------------------------------------------------------
// CANLI OKUTMA SAYACI & TEYİT MOTORU
// --------------------------------------------------------------
function getTotalScannedUnitsCount() {
    let totalDone = 0;
    if (Array.isArray(loadedOrders)) {
        loadedOrders.forEach(order => {
            if (order.urunler) {
                order.urunler.forEach(item => {
                    if (item.okunan_adet !== undefined && item.okunan_adet !== null && item.okunan_adet > 0) {
                        totalDone += item.okunan_adet;
                    } else if (item.tamamlandi) {
                        totalDone += (item.adet || 1);
                    }
                });
            }
        });
    }
    return totalDone;
}

function getTotalExcelUnitsCount() {
    let total = 0;
    if (Array.isArray(loadedOrders)) {
        loadedOrders.forEach(order => {
            if (order.urunler) {
                order.urunler.forEach(item => {
                    total += (item.adet || 1);
                });
            }
        });
    }
    return total;
}

function updateScanCounters(lastCode = '', lastName = '') {
    const totalDone = getTotalScannedUnitsCount();
    const totalAll = getTotalExcelUnitsCount();

    // 1. Canlı giriş kutusu yanındaki sayaç rozeti
    const counterVal = document.getElementById('scanCounterVal');
    if (counterVal) {
        counterVal.textContent = totalDone;
    }
    const badge = document.getElementById('scanCounterBadge');
    if (badge && lastCode) {
        badge.classList.remove('pulse-hit');
        void badge.offsetWidth;
        badge.classList.add('pulse-hit');
    }

    // 2. Header sayacı
    const headerUnits = document.getElementById('headerUnitsCount');
    if (headerUnits) {
        headerUnits.textContent = totalAll > 0 ? `${totalDone} / ${totalAll}` : totalDone;
    }

    // 3. Kamera modali içi sayaç
    const camVal = document.getElementById('camCounterVal');
    if (camVal) {
        camVal.textContent = totalDone;
    }

    // 4. Canlı Teyit Şeridi
    const confirmBar = document.getElementById('scanConfirmationBar');
    if (confirmBar) {
        if (lastCode) {
            confirmBar.style.display = 'flex';
            const codeEl = document.getElementById('confirmBarcodeCode');
            if (codeEl) codeEl.textContent = lastCode;
            const nameEl = document.getElementById('confirmProductName');
            if (nameEl) nameEl.textContent = lastName ? `• ${lastName}` : '';
            const tagEl = document.getElementById('confirmCountTag');
            if (tagEl) tagEl.textContent = `🎯 Okutulan: ${totalDone} Adet`;
        } else if (totalDone === 0) {
            confirmBar.style.display = 'none';
        }
    }
}

window.resetScanCounter = function() {
    const totalCurrent = getTotalScannedUnitsCount();
    if (totalCurrent === 0) {
        const alertBox = document.getElementById('orderScanAlert');
        showAlert(alertBox, 'info', 'ℹ️ Sayaç zaten 0 adette.');
        return;
    }

    if (!confirm(`Okutulan ürün sayacını (${totalCurrent} Adet) sıfırlamak istediğinize emin misiniz?`)) {
        return;
    }

    // 1. Tüm yüklü siparişlerdeki okutulma durumlarını sıfırla
    if (Array.isArray(loadedOrders)) {
        loadedOrders.forEach(order => {
            order.tamamlandi = false;
            if (order.urunler) {
                order.urunler.forEach(item => {
                    item.tamamlandi = false;
                    item.okunan_adet = 0;
                    item.seri_nolar = [];
                });
            }
        });
        saveOrdersToStorage(loadedFileName, loadedOrders);
    }

    // 2. Canlı sayaçları ve UI bileşenlerini güncelle
    updateScanCounters();
    renderOrderCards(document.getElementById('orderFilterInput')?.value.trim().toLowerCase() || '');
    if (activeOrderIndex !== -1) {
        renderActiveOrderItems();
        updateOrderProgress();
    }

    // 3. Teyit şeridini gizle
    const confirmBar = document.getElementById('scanConfirmationBar');
    if (confirmBar) confirmBar.style.display = 'none';

    // 4. Kullanıcıya net bildirim ver
    const alertBox = document.getElementById('orderScanAlert');
    showAlert(alertBox, 'success', '🎯 Okutma sayacı başarıyla sıfırlandı! (0 Adet)');
    playSuccessBeep();
};

// ==============================================================
// BARKOD & SERİ NO ARŞİVLEME SİSTEMİ - SAĞLAM VE HIZLI JS MOTORU
// ==============================================================

let html5QrCode = null;
let activeCamTarget = null;
let loadedOrders = [];
let loadedFileName = '';
let activeOrderIndex = -1;
let scanStep = 1; // 1: Kargo Barkodu, 2: Ürün Seri No
let packageStatusFilter = 'all'; // 'all', 'pending', 'completed'
let activePendingItemIndex = -1;

const ORDERS_STORAGE_KEY = 'barkod_arsiv_persisted_orders_v1';
const ORDERS_META_KEY = 'barkod_arsiv_persisted_meta_v1';

// --------------------------------------------------------------
// PERSISTENCE & KALICI SİPARİŞ DEPOLAMA (YENİ EXCEL'E KADAR GEÇERLİ)
// --------------------------------------------------------------
function saveOrdersToStorage(filename, orders) {
    try {
        if (!orders || orders.length === 0) return;
        localStorage.setItem(ORDERS_STORAGE_KEY, JSON.stringify(orders));
        localStorage.setItem(ORDERS_META_KEY, JSON.stringify({
            filename: filename || loadedFileName || 'Yüklü Sipariş Listesi',
            count: orders.length,
            savedAt: new Date().toISOString()
        }));
    } catch (e) {
        console.warn('localStorage kaydetme hatası:', e);
    }
}

function restorePersistedOrders() {
    try {
        const savedMeta = localStorage.getItem(ORDERS_META_KEY);
        const savedOrders = localStorage.getItem(ORDERS_STORAGE_KEY);
        if (savedMeta && savedOrders) {
            const meta = JSON.parse(savedMeta);
            const orders = JSON.parse(savedOrders);
            if (Array.isArray(orders) && orders.length > 0) {
                loadedOrders = orders;
                loadedFileName = meta.filename || 'Yüklü Sipariş Listesi';
                updateOrderUIElements(loadedFileName, loadedOrders.length);
                renderOrderCards();
                console.log(`[Barkod Arşiv] ${loadedOrders.length} paket yerel depolamadan başarıyla geri yüklendi.`);
                return true;
            }
        }
    } catch (e) {
        console.warn('restorePersistedOrders hatası:', e);
    }
    return false;
}

async function fetchServerPersistedOrders(silent = false) {
    // Eğer şu an bir yükleme yapılıyorsa veya son 15 saniyede yükleme yapıldıysa buluttan çekip ezme!
    if (isUploadingExcel || (Date.now() - lastUploadTimestamp < 15000)) {
        return false;
    }
    try {
        const res = await fetch('/api/arsiv/get-current-orders?t=' + Date.now());
        const data = await res.json();
        if (data && data.status === 'success' && Array.isArray(data.orders) && data.orders.length > 0) {
            loadedOrders = data.orders;
            loadedFileName = data.filename || 'Aktif Siparişler';
            saveOrdersToStorage(loadedFileName, loadedOrders);
            updateOrderUIElements(loadedFileName, loadedOrders.length);
            renderOrderCards(document.getElementById('orderFilterInput')?.value.trim().toLowerCase() || '');
            console.log(`[Barkod Arşiv] Buluttan ${loadedOrders.length} paket tüm cihazlarla senkronize edildi.`);
            if (!silent) {
                const badge = document.getElementById('cloudSyncStatusText');
                if (badge) badge.textContent = '☁️ Bulutla Senkronize';
            }
            return true;
        } else if (data && data.status === 'empty') {
            // SADECE yerelde de hiçbir sipariş yoksa boş bırak; kullanıcının mevcut siparişlerini ASLA ezme!
            if (loadedOrders.length === 0) {
                loadedOrders = [];
                loadedFileName = '';
                updateOrderUIElements('', 0);
                renderOrderCards();
            }
        }
    } catch (e) {
        console.warn('Sunucudan sipariş çekme hatası:', e);
    }
    return false;
}

window.triggerManualCloudSync = async function() {
    const btn = document.getElementById('btnSyncCloud');
    const icon = document.getElementById('syncCloudIcon');
    const txt = document.getElementById('syncCloudText');
    if (btn) btn.disabled = true;
    if (icon) icon.textContent = '⏳';
    if (txt) txt.textContent = 'Eşitleniyor...';

    try {
        const syncUrl = window.location.pathname.startsWith('/arsiv') || window.location.pathname.startsWith('/barkod_arsiv')
            ? '/api/arsiv/sync-cloud'
            : (window.location.port === '5004' ? '/api/sync-cloud' : '/api/arsiv/sync-cloud');

        const res = await fetch(syncUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        });
        const data = await res.json();

        // Aktif siparişleri de buluttan tazele
        await fetchServerPersistedOrders(false);

        // Güncel istatistikleri ve arama tablosunu tazele
        if (typeof loadStats === 'function') loadStats();
        if (typeof performArchiveSearch === 'function') performArchiveSearch();

        if (icon) icon.textContent = '✅';
        if (txt) txt.textContent = 'Senkronize';

        const alertBox = document.getElementById('orderScanAlert');
        const count = data?.sync_result?.total_records || data?.stats?.toplam_kayit || loadedOrders.length;
        showAlert(alertBox, 'success', `☁️ Bulut arşiv ve cihazlar eşitlendi (${count} arşiv kaydı, ${loadedOrders.length} paket).`);
        playSuccessBeep();

        setTimeout(() => {
            if (icon) icon.textContent = '☁️';
            if (txt) txt.textContent = 'Bulutla Eşitle';
            if (btn) btn.disabled = false;
        }, 2000);
    } catch (e) {
        console.error('Senkronizasyon hatası:', e);
        if (icon) icon.textContent = '⚠️';
        if (txt) txt.textContent = 'Hata';
        if (btn) btn.disabled = false;
        setTimeout(() => {
            if (icon) icon.textContent = '☁️';
            if (txt) txt.textContent = 'Bulutla Eşitle';
        }, 3000);
    }
};
window.syncOrdersNow = window.triggerManualCloudSync;

window.clearLoadedOrders = async function() {
    if (!confirm('Yüklü sipariş listesini tamamen temizlemek istediğinize emin misiniz?')) return;
    loadedOrders = [];
    loadedFileName = '';
    activeOrderIndex = -1;
    localStorage.removeItem(ORDERS_STORAGE_KEY);
    localStorage.removeItem(ORDERS_META_KEY);
    try {
        await fetch('/api/arsiv/clear-current-orders', { method: 'POST' });
    } catch (e) {}
    updateOrderUIElements('', 0);
    renderOrderCards();
    resetToWaitingState();
    const alertBox = document.getElementById('orderScanAlert');
    showAlert(alertBox, 'info', 'Sipariş listesi temizlendi. Yeni bir Excel yükleyebilirsiniz.');
};

window.setPackageStatusFilter = function(filter) {
    packageStatusFilter = filter;
    document.querySelectorAll('.filter-pill').forEach(btn => btn.classList.remove('active'));
    if (filter === 'all') document.getElementById('btnFilterAll')?.classList.add('active');
    else if (filter === 'pending') document.getElementById('btnFilterPending')?.classList.add('active');
    else if (filter === 'completed') document.getElementById('btnFilterCompleted')?.classList.add('active');
    renderOrderCards(document.getElementById('orderFilterInput')?.value.trim().toLowerCase() || '');
};

function updateOrderUIElements(filename, count) {
    const navOrderCount = document.getElementById('navOrderCount');
    if (navOrderCount) navOrderCount.textContent = count;

    const loadedCount = document.getElementById('loadedOrderCount');
    if (loadedCount) loadedCount.textContent = count;

    const headerPkgCount = document.getElementById('headerPkgCount');
    if (headerPkgCount) headerPkgCount.textContent = count;

    const btnPkgCount = document.getElementById('btnPkgCount');
    if (btnPkgCount) btnPkgCount.textContent = count;

    const btnGoToPackages = document.getElementById('btnGoToPackages');
    if (btnGoToPackages) btnGoToPackages.style.display = count > 0 ? 'inline-flex' : 'none';

    const btnClearOrders = document.getElementById('btnClearOrders');
    if (btnClearOrders) btnClearOrders.style.display = count > 0 ? 'inline-flex' : 'none';

    const dropTitle = document.getElementById('dropzoneMainTitle');
    const dropSub = document.getElementById('dropzoneSubtitle');
    if (count > 0 && filename) {
        if (dropTitle) dropTitle.innerHTML = `📁 Yüklü Liste: <span style="color:#60A5FA;">${filename}</span> (${count} Paket)`;
        if (dropSub) dropSub.textContent = 'Yeni Excel yüklenene kadar geçerli kalır • Yukarıdaki "Paketler" sekmesinden görebilirsiniz.';
    } else {
        if (dropTitle) dropTitle.textContent = 'Pazaryeri Excel Sipariş Listesini Sürükleyin veya Seçin';
        if (dropSub) dropSub.textContent = '(Trendyol, Hepsiburada, Amazon, Çiçeksepeti, Pazarama vb. • Yeni Excel yüklenene kadar geçerli kalır)';
    }

    const activeFileNameText = document.getElementById('activeFileNameText');
    if (activeFileNameText) {
        activeFileNameText.textContent = count > 0 ? `${filename} (${count} Paket)` : 'Henüz Excel Yüklenmedi';
    updateScanCounters();
    }
}

// Web Audio API (Lazy, güvenli başlatma)
let audioCtx = null;
function getAudioContext() {
    try {
        if (!audioCtx) {
            const AudioClass = window.AudioContext || window.webkitAudioContext;
            if (AudioClass) audioCtx = new AudioClass();
        }
        if (audioCtx && audioCtx.state === 'suspended') {
            audioCtx.resume();
        }
    } catch (e) {
        console.warn('AudioContext uyarısı:', e);
    }
    return audioCtx;
}

function playSuccessBeep() {
    try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.frequency.setValueAtTime(880, now);
        osc.frequency.setValueAtTime(1320, now + 0.08);
        gain.gain.setValueAtTime(0.2, now);
        gain.gain.exponentialRampToValueAtTime(0.001, now + 0.25);
        osc.start(now);
        osc.stop(now + 0.25);
    } catch (e) {}
}

function playWarningBuzzer() {
    try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.frequency.setValueAtTime(320, now);
        osc.frequency.setValueAtTime(220, now + 0.1);
        gain.gain.setValueAtTime(0.3, now);
        gain.gain.exponentialRampToValueAtTime(0.001, now + 0.35);
        osc.start(now);
        osc.stop(now + 0.35);
    } catch (e) {}
}

// --------------------------------------------------------------
// UYGULAMA BAŞLATICI (HER KOŞULDA ÇALIŞIR)
// --------------------------------------------------------------
function initApp() {
    console.log('[Barkod Arşiv] Başlatılıyor...');
    try { initTabs(); } catch (e) { console.error('Tabs error:', e); }
        try { 
        // 1. Cihazdaki yerel depolamayı hemen göster (0ms beklemesiz açılış)
        restorePersistedOrders();
        // 2. Ardından tüm cihazlar (telefon, tablet, başka PC) ile tam eşitlenmek için buluttan çek
        fetchServerPersistedOrders(true);
        // Sayfa tekrar öne geldiğinde (focus) otomatik bulut kontrolü yap
        window.addEventListener('focus', () => { fetchServerPersistedOrders(true); });
    } catch (e) { console.error('Persistence error:', e); }
    try { initExcelUpload(); } catch (e) { console.error('Excel upload error:', e); }
    try { initOrderWorkflow(); } catch (e) { console.error('Order workflow error:', e); }
    try { initFreeScan(); } catch (e) { console.error('Free scan error:', e); }
    try { initSearchAndArchive(); } catch (e) { console.error('Search error:', e); }
    try { initCameraScanner(); } catch (e) { console.error('Camera error:', e); }
    console.log('[Barkod Arşiv] Hazır!');
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}

// --------------------------------------------------------------
// 1. TAB GEÇİŞLERİ
// --------------------------------------------------------------
function initTabs() {
    const tabs = document.querySelectorAll('#view-arsiv .nav-tab');
    const panes = document.querySelectorAll('#view-arsiv .tab-pane');

    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            e.preventDefault();
            const targetId = tab.dataset.tab;
            if (!targetId) return;

            tabs.forEach(t => t.classList.remove('active'));
            panes.forEach(p => p.classList.remove('active'));

            tab.classList.add('active');
            const targetPane = document.getElementById(targetId);
            if (targetPane) {
                targetPane.classList.add('active');
            }

            if (targetId === 'tab-siparis') {
                const scanInput = document.getElementById('activeScanInput');
                if (scanInput) scanInput.focus();
            } else if (targetId === 'tab-paketler') {
                const f = document.getElementById('orderFilterInput');
                if (f) f.focus();
                renderOrderCards(f ? f.value.trim().toLowerCase() : '');
            } else if (targetId === 'tab-serbest') {
                const freeInput = document.getElementById('freeScanInput');
                if (freeInput) freeInput.focus();
            } else if (targetId === 'tab-arama') {
                const searchInput = document.getElementById('archiveSearchInput');
                if (searchInput) searchInput.focus();
            }
        });
    });
}

// --------------------------------------------------------------
// 2. EXCEL YÜKLEME & İNDİRİLENLER'DEN GETİR
// --------------------------------------------------------------
function initExcelUpload() {
    const dropzone = document.getElementById('excelDropzone');
    const fileInput = document.getElementById('excelFileInput');
    const browseBtn = document.getElementById('browseExcelBtn');
    const recentBtn = document.getElementById('btnLoadRecentExcel');

    if (browseBtn && fileInput) {
        browseBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileInput.click();
        });
    }

    if (dropzone && fileInput) {
        dropzone.addEventListener('click', (e) => {
            if (e.target.closest('button') || e.target.closest('input')) return;
            fileInput.click();
        });

        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.style.borderColor = '#3B82F6';
        });

        dropzone.addEventListener('dragleave', () => {
            dropzone.style.borderColor = 'rgba(59, 130, 246, 0.35)';
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.style.borderColor = 'rgba(59, 130, 246, 0.35)';
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleFileUpload(e.dataTransfer.files);
            }
        });
    }

    if (fileInput) {
        fileInput.addEventListener('change', () => {
            if (fileInput.files && fileInput.files.length > 0) {
                handleFileUpload(fileInput.files);
            }
        });
    }

    if (recentBtn) {
        recentBtn.addEventListener('click', async (e) => {
            e.preventDefault();
            e.stopPropagation();
            const orig = recentBtn.textContent;
            recentBtn.textContent = 'Yükleniyor...';
            try {
                const response = await fetch('/api/arsiv/load-recent-download', { method: 'POST' });
                const res = await response.json();
                if (res.status === 'success') {
                    loadedOrders = res.orders || [];
                    loadedFileName = res.filename || 'İndirilenler Sipariş Listesi';
                    saveOrdersToStorage(loadedFileName, loadedOrders);
                    updateOrderUIElements(loadedFileName, loadedOrders.length);
                    renderOrderCards();
                    resetToWaitingState();
                    const alertBox = document.getElementById('orderScanAlert');
                    const restoredCount = loadedOrders.filter(o => o.tamamlandi).length;
                    let successMsg = `📦 ${loadedOrders.length} Paket yüklendi ve kaydedildi!`;
                    if (restoredCount > 0) {
                        successMsg += ` (${restoredCount} paket geçmiş arşivden tamamlandı olarak geri yüklendi)`;
                    }
                    showAlert(alertBox, 'success', `${successMsg} Kolideki KARGO BARKODUNU okutun.`);
                    playSuccessBeep();
                } else {
                    alert('Hata: ' + (res.message || 'Dosya okunamadı. Lütfen "Dosya Seç" ile yükleyin.'));
                    playWarningBuzzer();
                }
            } catch (err) {
                alert('İndirilenler klasörüne doğrudan erişim yerel sunucuda (127.0.0.1:5004) çalışır. Bulutta dosyanızı "Dosya Seç" butonu ile yükleyebilirsiniz.');
                playWarningBuzzer();
            } finally {
                recentBtn.textContent = orig;
            }
        });
    }
}

async function handleFileUpload(filesInput) {
    if (isUploadingExcel) {
        console.warn('Yükleme zaten devam ediyor.');
        return;
    }
    if (!filesInput) return;

    const files = (filesInput instanceof FileList || Array.isArray(filesInput)) 
        ? Array.from(filesInput) 
        : [filesInput];

    const validFiles = files.filter(f => f && f.name && /\.(xlsx|xls|csv)$/i.test(f.name));
    if (validFiles.length === 0) {
        alert('Lütfen geçerli bir Excel (.xlsx, .xls) veya CSV dosyası seçin.');
        return;
    }

    isUploadingExcel = true;
    lastUploadTimestamp = Date.now();

    const formData = new FormData();
    validFiles.forEach(f => {
        formData.append('files', f);
    });

    const dropText = document.querySelector('.dropzone-text strong');
    const origText = dropText ? dropText.textContent : '';
    if (dropText) {
        if (validFiles.length === 1) {
            dropText.textContent = `"${validFiles[0].name}" yükleniyor ve okunuyor...`;
        } else {
            dropText.textContent = `${validFiles.length} adet Excel dosyası yükleniyor ve birleştiriliyor...`;
        }
    }

    try {
        const response = await fetch('/api/arsiv/upload-orders', {
            method: 'POST',
            body: formData
        });
        const res = await response.json();

        if (res.status === 'success' && Array.isArray(res.orders) && res.orders.length > 0) {
            loadedOrders = res.orders;
            loadedFileName = res.filename || (validFiles.length > 1 ? `${validFiles.length} Dosya` : validFiles[0].name);
            lastUploadTimestamp = Date.now();
            saveOrdersToStorage(loadedFileName, loadedOrders);
            updateOrderUIElements(loadedFileName, loadedOrders.length);

            // Filtreleri sıfırla ki yeni yüklenen tüm paketler anında görünsün
            packageStatusFilter = 'all';
            document.querySelectorAll('.filter-pill').forEach(b => b.classList.remove('active'));
            document.getElementById('btnFilterAll')?.classList.add('active');

            const fInp = document.getElementById('orderFilterInput');
            if (fInp) fInp.value = '';

            renderOrderCards('');
            resetToWaitingState();

            const alertBox = document.getElementById('orderScanAlert');
            const summaryBadge = validFiles.length > 1 ? `${validFiles.length} Dosya Birleştirildi` : loadedFileName;
            const restoredCount = loadedOrders.filter(o => o.tamamlandi).length;
            let successMsg = `📦 ${loadedOrders.length} Paket yüklendi (${summaryBadge})!`;
            if (restoredCount > 0) {
                successMsg += ` (${restoredCount} paket geçmiş arşivden tamamlandı olarak geri yüklendi)`;
            }
            showAlert(alertBox, 'success', `${successMsg} Kolideki kargo barkodunu veya sipariş numarasını okutun.`);
            playSuccessBeep();
        } else {
            alert('Hata: ' + (res.message || 'Dosyalar okunamadı veya sipariş bulunamadı'));
            if (dropText) dropText.textContent = origText;
            playWarningBuzzer();
        }
    } catch (err) {
        alert('Sunucu hatası: ' + err.message);
        if (dropText) dropText.textContent = origText;
        playWarningBuzzer();
    } finally {
        isUploadingExcel = false;
        const fi = document.getElementById('excelFileInput');
        if (fi) fi.value = '';
    }
}

// --------------------------------------------------------------
// 3. 2 ADIMLI AKILLI DEPO TARAMA AKIŞI
// --------------------------------------------------------------
function setScanStep(step) {
    scanStep = step;
    const chip1 = document.getElementById('stepChip1');
    const chip2 = document.getElementById('stepChip2');
    const title = document.getElementById('scanLabelTitle');
    const hint = document.getElementById('scanLabelHint');
    const input = document.getElementById('activeScanInput');

    if (!input) return;

    if (step === 1) {
        if (chip1) chip1.className = 'step-chip active';
        if (chip2) chip2.className = 'step-chip';
        if (title) {
            title.innerHTML = '📦 1. Adım: Kargo Barkodu / Sipariş No Okutun';
            title.style.color = '#38BDF8';
        }
        if (hint) hint.textContent = 'Kolideki kargo etiketini veya sipariş numarasını (Örn: 1099201) okutun, paket açılsın';
        input.className = 'scan-large-input step1-mode';
        input.placeholder = '1. ADIM: Kargo Barkodu veya Sipariş No Okutun...';
    } else {
        if (chip1) chip1.className = 'step-chip';
        if (chip2) chip2.className = 'step-chip active-step2';
        const order = (activeOrderIndex !== -1 && loadedOrders[activeOrderIndex]) ? loadedOrders[activeOrderIndex] : null;
        const musteriStr = order ? `(${order.musteri_adi})` : '';
        const pendingItem = order ? (
            (activePendingItemIndex !== -1 && order.urunler[activePendingItemIndex] && !order.urunler[activePendingItemIndex].tamamlandi)
                ? order.urunler[activePendingItemIndex]
                : order.urunler.find(i => !i.tamamlandi)
        ) : null;

        let itemInfo = '';
        if (pendingItem) {
            const nextCount = (pendingItem.okunan_adet || 0) + 1;
            const totalCount = pendingItem.adet || 1;
            itemInfo = totalCount > 1 
                ? ` — Beklenen: ${pendingItem.urun_adi.substring(0, 26)}... [${nextCount}/${totalCount}. Adet]`
                : ` — Beklenen: ${pendingItem.urun_adi.substring(0, 26)}...`;
        }
        if (title) {
            title.innerHTML = `🎯 2. Adım: Ürün Seri No / Barkodunu Okutun ${musteriStr}${itemInfo}`;
            title.style.color = '#34D399';
        }
        if (hint) {
            if (pendingItem && (pendingItem.adet || 1) > 1) {
                const nextCount = (pendingItem.okunan_adet || 0) + 1;
                hint.textContent = `Bu üründen ${pendingItem.adet} adet var. Şu an ${nextCount}/${pendingItem.adet}. adedin seri numarası/barkodu bekleniyor.`;
            } else {
                hint.textContent = 'Pakete koyduğunuz cihazın seri numarasını (SN/IMEI) veya barkodunu okutun';
            }
        }
        input.className = 'scan-large-input step2-mode';
        input.placeholder = pendingItem && (pendingItem.adet || 1) > 1
            ? `2. ADIM: ${(pendingItem.okunan_adet || 0) + 1}/${pendingItem.adet}. Adet Seri No / Barkod Okutun...`
            : '2. ADIM: Ürün Seri Numarası / Barkod Okutun...';
    }
    input.disabled = false;
    input.value = '';
    input.focus();
}


// --------------------------------------------------------------
// BEKLEME DURUMUNA SIFIRLA (KOLİ BARKODU BEKLENİYOR MODU)
// --------------------------------------------------------------
function resetToWaitingState() {
    activeOrderIndex = -1;
    activePendingItemIndex = -1;
    setScanStep(1);

    const mBadge = document.getElementById('activeMarketplace');
    if (mBadge) mBadge.textContent = 'Paket Bekleniyor';

    const oTitle = document.getElementById('activeOrderTitle');
    if (oTitle) oTitle.textContent = '📦 Kargo Barkodunu Okutun';

    const cName = document.getElementById('activeCustomerName');
    if (cName) cName.textContent = 'Kolideki kargo etiketini okutun, müşteri ve ürünler otomatik seçilsin.';

    const progressEl = document.getElementById('orderProgressPercent');
    if (progressEl) progressEl.textContent = '%0';

    const banner = document.getElementById('orderCompleteBanner');
    if (banner) banner.style.display = 'none';

    const tbody = document.getElementById('activeOrderItemsBody');
    if (tbody) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted" style="padding: 30px;">📦 Kolideki kargo barkodunu okuttuğunuzda müşteri ve ürünler buraya gelecektir.</td></tr>';
    }

    renderOrderCards(document.getElementById('orderFilterInput')?.value.trim().toLowerCase() || '');

    const scanInput = document.getElementById('activeScanInput');
    if (scanInput) {
        scanInput.value = '';
        scanInput.focus();
    }
}

function initOrderWorkflow() {
    const filterInput = document.getElementById('orderFilterInput');
    if (filterInput) {
        filterInput.addEventListener('input', () => {
            renderOrderCards(filterInput.value.trim().toLowerCase());
        });
    }

    const scanInput = document.getElementById('activeScanInput');
    if (scanInput) {
        scanInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const code = scanInput.value.trim();
                if (code) {
                    handleSmartScan(code);
                }
            }
        });
    }

    const resetCounterBtn = document.getElementById('btnResetScanCounter');
    if (resetCounterBtn) {
        resetCounterBtn.addEventListener('click', (e) => {
            e.preventDefault();
            if (typeof window.resetScanCounter === 'function') {
                window.resetScanCounter();
            }
        });
    }

    const resetBtn = document.getElementById('btnResetToStep1');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            resetToWaitingState();
            const alertBox = document.getElementById('orderScanAlert');
            showAlert(alertBox, 'info', 'Başa dönüldü: Lütfen koli kargo barkodunu okutun.');
        });
    }

    const nextBtn = document.getElementById('btnNextOrder');
    if (nextBtn) {
        nextBtn.addEventListener('click', () => {
            resetToWaitingState();
            const alertBox = document.getElementById('orderScanAlert');
            showAlert(alertBox, 'info', 'Sıradaki koli için kargo barkodunu okutun.');
        });
    }

    setScanStep(1);
}

async function handleSmartScan(code) {
    const cleanCode = code.trim();
    const alertBox = document.getElementById('orderScanAlert');
    const scanInput = document.getElementById('activeScanInput');

    if (!cleanCode) return;

    if (loadedOrders.length === 0) {
        playWarningBuzzer();
        showAlert(alertBox, 'error', 'Henüz sipariş listesi yüklenmedi! Lütfen yukarıdan "İndirilenler\'den Getir" veya "Dosya Seç" ile listenizi yükleyin.');
        scanInput.value = '';
        return;
    }

    // 1. Kargo Barkodu Eşleşme Kontrolü
    const cleanAlnum = cleanCode.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

    let matchedIdx = loadedOrders.findIndex(o => 
        (o.kargo_barkodu && o.kargo_barkodu.trim().toLowerCase() === cleanCode.toLowerCase()) ||
        (o.siparis_no && o.siparis_no.trim().toLowerCase() === cleanCode.toLowerCase()) ||
        (o.paket_no && o.paket_no.trim().toLowerCase() === cleanCode.toLowerCase())
    );

    if (matchedIdx === -1 && cleanAlnum) {
        matchedIdx = loadedOrders.findIndex(o => {
            const kbAlnum = (o.kargo_barkodu || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            const spAlnum = (o.siparis_no || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            const pkAlnum = (o.paket_no || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            return (kbAlnum && kbAlnum === cleanAlnum) || 
                   (spAlnum && spAlnum === cleanAlnum) || 
                   (pkAlnum && pkAlnum === cleanAlnum);
        });
    }

    if (matchedIdx === -1 && cleanCode.length >= 6) {
        matchedIdx = loadedOrders.findIndex(o => {
            const kb = (o.kargo_barkodu || '').trim().toLowerCase();
            const sp = (o.siparis_no || '').trim().toLowerCase();
            return (kb && (kb.endsWith(cleanCode.toLowerCase()) || cleanCode.toLowerCase().endsWith(kb))) ||
                   (sp && (sp.endsWith(cleanCode.toLowerCase()) || cleanCode.toLowerCase().endsWith(sp)));
        });
    }

    if (scanStep === 1 || matchedIdx !== -1) {
        if (matchedIdx !== -1) {
            selectOrder(matchedIdx);
            playSuccessBeep();
            const order = loadedOrders[matchedIdx];
            showAlert(alertBox, 'success', `📦 MÜŞTERİ SEÇİLDİ: ${order.musteri_adi} | Kargo: ${order.kargo_barkodu || order.siparis_no} (${order.pazaryeri}). Şimdi ürün seri numarasını okutun!`);
            return;
        } else {
            playWarningBuzzer();
            showAlert(alertBox, 'error', `⚠️ Bu kargo barkodu listede bulunamadı: "${cleanCode}". Kolideki doğru kargo barkodunu okutun.`);
            scanInput.value = '';
            scanInput.focus();
            return;
        }
    }

    if (scanStep === 2) {
        if (activeOrderIndex === -1) {
            setScanStep(1);
            return;
        }
        await processProductSerialScan(cleanCode);
    }
}

async function processProductSerialScan(barcode) {
    if (activeOrderIndex === -1) return;
    const order = loadedOrders[activeOrderIndex];
    const alertBox = document.getElementById('orderScanAlert');
    const scanInput = document.getElementById('activeScanInput');

    let pendingItem = null;
    if (activePendingItemIndex !== -1 && order.urunler[activePendingItemIndex] && !order.urunler[activePendingItemIndex].tamamlandi) {
        pendingItem = order.urunler[activePendingItemIndex];
    }
    if (!pendingItem) {
        pendingItem = order.urunler.find(i => !i.tamamlandi && (
            (i.gtin && i.gtin.trim().toLowerCase() === barcode.trim().toLowerCase()) ||
            (i.stok_kod && i.stok_kod.trim().toLowerCase() === barcode.trim().toLowerCase())
        ));
    }
    if (!pendingItem) {
        pendingItem = order.urunler.find(i => !i.tamamlandi);
    }

    if (!pendingItem) {
        showAlert(alertBox, 'duplicate', 'Bu paketteki tüm ürünler zaten tamamlanmış!');
        scanInput.value = '';
        setScanStep(1);
        return;
    }

    try {
        const response = await fetch('/api/arsiv/record', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                kargo_barkodu: order.kargo_barkodu || '',
                siparis_no: order.siparis_no || '',
                musteri_adi: order.musteri_adi,
                urun_adi: pendingItem.urun_adi,
                barkod_seri_no: barcode,
                pazaryeri: order.pazaryeri,
                okutma_yontemi: 'El Terminali / Kargo Eşleme'
            })
        });
        const res = await response.json();

        if (res.status === 'success') {
            if (!pendingItem.seri_nolar) pendingItem.seri_nolar = [];
            pendingItem.seri_nolar.push(barcode);
            pendingItem.okunan_adet = (pendingItem.okunan_adet || 0) + 1;

            const isItemDone = pendingItem.okunan_adet >= (pendingItem.adet || 1);
            pendingItem.tamamlandi = isItemDone;

            if (isItemDone && activePendingItemIndex !== -1) {
                activePendingItemIndex = -1;
            }

            renderActiveOrderItems();
            updateOrderProgress();
            renderOrderCards();
            updateScanCounters(barcode, pendingItem.urun_adi);
            updateHeaderStats(res.stats);
            saveOrdersToStorage(loadedFileName, loadedOrders);

            const allDone = order.urunler.every(i => i.tamamlandi);

            if (res.is_duplicate) {
                playWarningBuzzer();
                if (isItemDone) {
                    showAlert(alertBox, 'duplicate', `⚠️ MÜKERRER UYARISI: Bu Seri No (${barcode}) daha önce kaydedilmiş! Yine de arşive mühürlendi. [${pendingItem.okunan_adet}/${pendingItem.adet} Tamamlandı]`);
                } else {
                    showAlert(alertBox, 'duplicate', `⚠️ MÜKERRER UYARISI: Bu Seri No (${barcode}) daha önce kaydedilmiş! [${pendingItem.okunan_adet}/${pendingItem.adet} Okundu]. Lütfen bu ürünün ${pendingItem.okunan_adet + 1}. seri numarasını/barkodunu okutun.`);
                }
            } else {
                playSuccessBeep();
                if (isItemDone) {
                    showAlert(alertBox, 'success', `✅ ÜRÜN MÜHÜRLENDİ & TAMAMLANDI (${pendingItem.okunan_adet}/${pendingItem.adet}): ${pendingItem.urun_adi.substring(0, 32)}... (Seri No: ${barcode})`);
                } else {
                    showAlert(alertBox, 'success', `✅ ADET MÜHÜRLENDİ (${pendingItem.okunan_adet}/${pendingItem.adet}): ${pendingItem.urun_adi.substring(0, 28)}... (Seri No: ${barcode}) — Lütfen bu ürünün ${pendingItem.okunan_adet + 1}. seri numarasını/barkodunu okutun!`);
                }
            }

            if (allDone) {
                playSuccessBeep();
                setTimeout(() => {
                    showAlert(alertBox, 'success', `🎉 TEBRİKLER! Bu paketin (${order.musteri_adi}) tüm ürünleri tamamlandı. Sıradaki koli için KARGO BARKODU okutun.`);
                    resetToWaitingState();
                }, 1200);
            } else {
                setScanStep(2);
                scanInput.value = '';
                scanInput.focus();
            }
        } else {
            playWarningBuzzer();
            showAlert(alertBox, 'error', 'Hata: ' + (res.message || 'Kaydedilemedi'));
        }
    } catch (err) {
        playWarningBuzzer();
        showAlert(alertBox, 'error', 'Bağlantı hatası: ' + err.message);
    }
}

function renderOrderCards(filterText = '') {
    const container = document.getElementById('orderCardsContainer');
    if (!container) return;

    container.innerHTML = '';
    if (loadedOrders.length === 0) {
        container.innerHTML = '<div class="empty-state">Henüz sipariş yüklenmedi.<br>Yukarıdaki \'Sipariş & Barkod Okutma\' sekmesinden Excel dosyanızı seçin.</div>';
        return;
    }

    const countAll = loadedOrders.length;
    const countDone = loadedOrders.filter(o => o.tamamlandi || (o.urunler && o.urunler.length > 0 && o.urunler.every(i => i.tamamlandi))).length;
    const countPending = countAll - countDone;

    const cfAll = document.getElementById('countFilterAll');
    if (cfAll) cfAll.textContent = countAll;
    const cfPending = document.getElementById('countFilterPending');
    if (cfPending) cfPending.textContent = countPending;
    const cfCompleted = document.getElementById('countFilterCompleted');
    if (cfCompleted) cfCompleted.textContent = countDone;

    let matchedCount = 0;

    loadedOrders.forEach((order, idx) => {
        const isCompleted = order.tamamlandi || (order.urunler && order.urunler.length > 0 && order.urunler.every(i => i.tamamlandi));

        // Durum Filtresi
        if (packageStatusFilter === 'pending' && isCompleted) return;
        if (packageStatusFilter === 'completed' && !isCompleted) return;

        // Metin Filtresi
        if (filterText) {
            const matchKargo = (order.kargo_barkodu || '').toLowerCase().includes(filterText);
            const matchOrder = (order.siparis_no || '').toLowerCase().includes(filterText);
            const matchCustomer = (order.musteri_adi || '').toLowerCase().includes(filterText);
            if (!matchKargo && !matchOrder && !matchCustomer) return;
        }

        matchedCount++;
        const isSelected = (idx === activeOrderIndex);
        const card = document.createElement('div');
        card.className = `order-card-item ${isSelected ? 'selected' : ''} ${isCompleted ? 'completed' : ''}`;
        
        const isWebOrder = order.is_web_order || (order.kargo_barkodu === order.siparis_no);
        const kargoGosterim = isWebOrder ? `Sipariş: ${order.siparis_no}` : (order.kargo_barkodu || order.siparis_no || 'Barkod Yok');
        const kargoBilgi = order.kargo_firmasi ? ` • ${order.kargo_firmasi}` : '';
        const sehirText = (order.sehir || order.ilce) ? ` • ${[order.ilce, order.sehir].filter(Boolean).join('/')}` : '';

        const statusBadge = isCompleted 
            ? '<span class="status-badge-done">Tamamlandı ✅</span>'
            : '<span class="status-badge-pending">Bekliyor ⏳</span>';

        const itemCount = order.urunler ? order.urunler.length : 1;
        const totalUnits = order.urunler ? order.urunler.reduce((sum, i) => sum + (parseInt(i.adet, 10) || 1), 0) : 1;
        const readUnits = order.urunler ? order.urunler.reduce((sum, i) => sum + (parseInt(i.okunan_adet, 10) || 0), 0) : 0;

        card.innerHTML = `
            <div class="card-item-top">
                <span class="card-order-no" style="font-size: 14.5px; color: #38BDF8;">📦 ${kargoGosterim}</span>
                <span class="card-market-tag">${order.pazaryeri || 'Pazar'}${kargoBilgi}</span>
            </div>
            <div class="card-customer" style="font-size: 15px; font-weight: 700; color: #F1F5F9;">${order.musteri_adi || 'Müşteri'}${sehirText}</div>
            <div class="card-item-summary" style="display: flex; justify-content: space-between; align-items: center; margin-top: 6px;">
                <span style="font-size: 12px; color: var(--text-muted);">${itemCount} Kalem (${readUnits}/${totalUnits} Adet) • Sipariş: ${order.siparis_no || '-'}</span>
                ${statusBadge}
            </div>
        `;
        card.addEventListener('click', () => selectOrder(idx));
        container.appendChild(card);
    });

    if (matchedCount === 0 && loadedOrders.length > 0) {
        container.innerHTML = '<div class="empty-state">Filtreye uygun paket bulunamadı.</div>';
    }
}

function selectOrder(index) {
    if (index < 0 || index >= loadedOrders.length) return;
    activeOrderIndex = index;
    activePendingItemIndex = -1;
    const order = loadedOrders[index];

    // Otomatik olarak Sipariş & Tarama sekmesine geç
    if (typeof switchTab === 'function') {
        switchTab('tab-siparis');
    }

    renderOrderCards(document.getElementById('orderFilterInput')?.value.trim().toLowerCase() || '');

    const isWebOrder = order.is_web_order || (order.kargo_barkodu === order.siparis_no);
    const kargoBilgisi = order.kargo_firmasi ? ` (${order.kargo_firmasi})` : '';
    const mBadge = document.getElementById('activeMarketplace');
    if (mBadge) mBadge.textContent = (order.pazaryeri || 'Pazaryeri') + kargoBilgisi;

    const oTitle = document.getElementById('activeOrderTitle');
    const displayCode = isWebOrder ? `Sipariş No: ${order.siparis_no}` : `Kargo Barkodu: ${order.kargo_barkodu || order.siparis_no}`;
    if (oTitle) oTitle.textContent = `📦 ${displayCode}`;

    const sehirBilgi = (order.sehir || order.ilce) ? ` (${[order.ilce, order.sehir].filter(Boolean).join(' / ')})` : '';
    const telBilgi = order.telefon ? ` • Tel: ${order.telefon}` : '';
    const cName = document.getElementById('activeCustomerName');
    if (cName) cName.textContent = `Alıcı: ${order.musteri_adi || 'Bilinmiyor'}${sehirBilgi}${telBilgi} • Sipariş No: ${order.siparis_no || '-'}`;

    renderActiveOrderItems();
    updateOrderProgress();

    setScanStep(2);

    const alertBox = document.getElementById('orderScanAlert');
    showAlert(alertBox, 'info', `🎯 Seçilen Paket: ${order.musteri_adi} (${displayCode}). Lütfen ürün seri numarasını okutun.`);
}

function renderActiveOrderItems() {
    const tbody = document.getElementById('activeOrderItemsBody');
    if (!tbody || activeOrderIndex === -1) return;

    const order = loadedOrders[activeOrderIndex];
    tbody.innerHTML = '';

    const currentPendingItem = (activePendingItemIndex !== -1 && order.urunler[activePendingItemIndex] && !order.urunler[activePendingItemIndex].tamamlandi)
        ? order.urunler[activePendingItemIndex]
        : order.urunler.find(i => !i.tamamlandi);

    order.urunler.forEach((item, itemIdx) => {
        const isDone = item.tamamlandi;
        const isPendingTarget = (item === currentPendingItem && !isDone);
        const inProgress = (!isDone && (item.okunan_adet || 0) > 0);

        const tr = document.createElement('tr');
        if (isDone) {
            tr.className = 'completed-row';
        } else if (isPendingTarget) {
            tr.className = 'active-pending-row';
        } else if (inProgress) {
            tr.className = 'in-progress-row';
        }

        tr.style.cursor = isDone ? 'default' : 'pointer';
        tr.title = isDone ? 'Tamamlandı' : 'Bu ürünü okutmak için tıklayın';
        tr.addEventListener('click', () => {
            if (!isDone) {
                activePendingItemIndex = itemIdx;
                renderActiveOrderItems();
                setScanStep(2);
            }
        });

        const serialsHtml = (item.seri_nolar && item.seri_nolar.length > 0)
            ? item.seri_nolar.map(sn => `
                <span class="code-pill serial-item-pill" title="Seri No: ${sn}">
                    <span class="sn-text">${sn}</span>
                    <button type="button" class="sn-action-btn sn-edit" title="Seri Numarasını Düzelt" onclick="event.stopPropagation(); editScannedSerial(${itemIdx}, '${sn}')">✏️</button>
                    <button type="button" class="sn-action-btn sn-delete" title="Yanlış Okunan Seri Numarasını Sil" onclick="event.stopPropagation(); deleteScannedSerial(${itemIdx}, '${sn}')">🗑️</button>
                </span>
            `).join(' ')
            : '<span class="text-muted">Okutulmadı</span>';

        let statusIcon = '⏳';
        if (isDone) statusIcon = '✅';
        else if (inProgress) statusIcon = '🔄';

        tr.innerHTML = `
            <td><span class="item-check-icon">${statusIcon}</span></td>
            <td>
                <strong>${item.urun_adi}</strong>
                ${isPendingTarget ? '<span class="next-target-tag">Sıradaki</span>' : ''}
                ${item.stok_kod ? `<div style="font-size: 11px; color: var(--text-muted); font-family: monospace;">Kod: ${item.stok_kod}</div>` : ''}
            </td>
            <td style="text-align: center;"><strong>${item.okunan_adet || 0} / ${item.adet}</strong></td>
            <td>${serialsHtml}</td>
        `;
        tbody.appendChild(tr);
    });
}

function updateOrderProgress() {
    if (activeOrderIndex === -1) return;
    const order = loadedOrders[activeOrderIndex];

    const totalQty = order.urunler.reduce((sum, i) => sum + (parseInt(i.adet, 10) || 1), 0);
    const readQty = order.urunler.reduce((sum, i) => sum + (parseInt(i.okunan_adet, 10) || 0), 0);

    const allItemsCompleted = order.urunler.length > 0 && order.urunler.every(i => i.tamamlandi);
    let percent = totalQty > 0 ? Math.min(100, Math.round((readQty / totalQty) * 100)) : 0;
    if (allItemsCompleted) percent = 100;

    const progressEl = document.getElementById('orderProgressPercent');
    if (progressEl) progressEl.textContent = `%${percent}`;

    const banner = document.getElementById('orderCompleteBanner');
    if (percent === 100 && allItemsCompleted) {
        order.tamamlandi = true;
        if (banner) banner.style.display = 'flex';
    } else {
        order.tamamlandi = false;
        if (banner) banner.style.display = 'none';
    }
}

// --------------------------------------------------------------
// 4. HIZLI SERBEST KAYIT
// --------------------------------------------------------------
function initFreeScan() {
    const scanInput = document.getElementById('freeScanInput');
    const saveBtn = document.getElementById('btnSaveFreeScan');
    const refreshBtn = document.getElementById('btnRefreshStats');

    if (scanInput) {
        scanInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                executeFreeScan();
            }
        });
    }

    if (saveBtn) {
        saveBtn.addEventListener('click', executeFreeScan);
    }

    if (refreshBtn) {
        refreshBtn.addEventListener('click', refreshDashboardStats);
    }
}

async function executeFreeScan() {
    const scanInput = document.getElementById('freeScanInput');
    const barcode = scanInput ? scanInput.value.trim() : '';
    const alertBox = document.getElementById('freeScanAlert');

    if (!barcode) {
        showAlert(alertBox, 'error', 'Lütfen geçerli bir barkod / seri no okutun!');
        return;
    }

    const kargoBarcode = document.getElementById('freeKargoBarcode')?.value.trim() || '';
    const orderNo = document.getElementById('freeOrderNo')?.value.trim() || '';
    const customer = document.getElementById('freeCustomerName')?.value.trim() || '';
    const product = document.getElementById('freeProductName')?.value.trim() || '';
    const marketplace = document.getElementById('freeMarketplace')?.value || 'Diğer';

    try {
        const response = await fetch('/api/arsiv/record', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                kargo_barkodu: kargoBarcode,
                siparis_no: orderNo,
                musteri_adi: customer,
                urun_adi: product,
                barkod_seri_no: barcode,
                pazaryeri: marketplace,
                okutma_yontemi: 'Hızlı Serbest Kayıt'
            })
        });
        const res = await response.json();

        if (res.status === 'success') {
            updateHeaderStats(res.stats);
            addRecentScanRow(res.record);

            if (res.is_duplicate) {
                playWarningBuzzer();
                showAlert(alertBox, 'duplicate', `⚠️ MÜKERRER BARKOD: Bu Seri No (${barcode}) daha önce kaydedilmiş! Kayıt mühürlendi.`);
            } else {
                playSuccessBeep();
                showAlert(alertBox, 'success', `✅ Seri No Başarıyla Diske Yazıldı: ${barcode}`);
            }

            scanInput.value = '';
            scanInput.focus();
        } else {
            playWarningBuzzer();
            showAlert(alertBox, 'error', 'Hata: ' + (res.message || 'Kayıt başarısız'));
        }
    } catch (err) {
        playWarningBuzzer();
        showAlert(alertBox, 'error', 'Sunucu hatası: ' + err.message);
    }
}

function addRecentScanRow(record) {
    const tbody = document.getElementById('recentScansBody');
    if (!tbody) return;

    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${record.kayit_tarihi}</td>
        <td><span class="code-pill">${record.barkod_seri_no}</span></td>
        <td><strong>${record.kargo_barkodu || record.siparis_no || '-'}</strong></td>
        <td>${record.musteri_adi || '-'}</td>
        <td><span class="market-tag">${record.pazaryeri || 'Diğer'}</span></td>
    `;
    tbody.insertBefore(tr, tbody.firstChild);
}

// --------------------------------------------------------------
// 5. ARAMA & KANIT ARŞİVİ
// --------------------------------------------------------------
function initSearchAndArchive() {
    const searchBtn = document.getElementById('btnExecuteSearch');
    const searchInput = document.getElementById('archiveSearchInput');
    const excelBtn = document.getElementById('btnDownloadExcelArchive');
    const fullBackupBtn = document.getElementById('btnFullExcelExport');

    if (searchBtn) searchBtn.addEventListener('click', executeArchiveSearch);
    if (searchInput) {
        searchInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                executeArchiveSearch();
            }
        });
    }

    if (excelBtn) {
        excelBtn.addEventListener('click', () => {
            const q = document.getElementById('archiveSearchInput')?.value.trim() || '';
            const start = document.getElementById('archiveStartDate')?.value || '';
            const end = document.getElementById('archiveEndDate')?.value || '';
            window.location.href = `/api/arsiv/export-excel?q=${encodeURIComponent(q)}&start_date=${start}&end_date=${end}`;
        });
    }

    if (fullBackupBtn) {
        fullBackupBtn.addEventListener('click', () => {
            window.location.href = '/api/arsiv/export-excel';
        });
    }

    document.getElementById('btnCloseProofModal')?.addEventListener('click', () => {
        const m = document.getElementById('proofModal');
        if (m) m.style.display = 'none';
    });
    document.getElementById('btnDismissProof')?.addEventListener('click', () => {
        const m = document.getElementById('proofModal');
        if (m) m.style.display = 'none';
    });
    document.getElementById('btnPrintProof')?.addEventListener('click', () => {
        window.print();
    });
}

async function executeArchiveSearch() {
    const q = document.getElementById('archiveSearchInput')?.value.trim() || '';
    const start = document.getElementById('archiveStartDate')?.value || '';
    const end = document.getElementById('archiveEndDate')?.value || '';
    const tbody = document.getElementById('searchResultsBody');
    const countSpan = document.getElementById('searchResultCount');

    if (tbody) tbody.innerHTML = '<tr><td colspan="8" class="text-center">Aranıyor...</td></tr>';

    try {
        const response = await fetch(`/api/arsiv/search?q=${encodeURIComponent(q)}&start_date=${start}&end_date=${end}`);
        const res = await response.json();

        if (res.status === 'success') {
            if (countSpan) countSpan.textContent = res.count;
            if (!tbody) return;

            if (res.records.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted">Kayıt bulunamadı.</td></tr>';
                return;
            }

            tbody.innerHTML = '';
            res.records.forEach(r => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>#${r.id}</td>
                    <td>${r.kayit_tarihi}</td>
                    <td><strong style="color: #38BDF8;">🏷️ ${r.kargo_barkodu || '-'}</strong></td>
                    <td><span class="code-pill">${r.barkod_seri_no}</span></td>
                    <td>${r.siparis_no || '-'}</td>
                    <td>${r.musteri_adi || '-'}</td>
                    <td>${r.urun_adi || '-'}</td>
                    <td><span class="market-tag">${r.pazaryeri}</span></td>
                    <td style="white-space: nowrap;">
                        <button class="btn-secondary btn-sm" onclick="openProofModal(${r.id})" title="Kanıt Belgesi">Kanıt 📜</button>
                        <button class="btn-secondary btn-sm" onclick="openEditRecordModal(${r.id})" style="margin-left: 4px; color: #38BDF8; border-color: rgba(56, 189, 248, 0.4);" title="Kaydı Düzenle">✏️</button>
                        <button class="btn-secondary btn-sm" onclick="deleteArchiveRecord(${r.id}, '${r.barkod_seri_no}')" style="margin-left: 4px; color: #EF4444; border-color: rgba(239, 68, 68, 0.4);" title="Kaydı Sil">🗑️</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }
    } catch (err) {
        if (tbody) tbody.innerHTML = `<tr><td colspan="8" class="text-center text-danger">Arama hatası: ${err.message}</td></tr>`;
    }
}

window.openProofModal = async function(recordId) {
    try {
        const response = await fetch(`/api/arsiv/proof/${recordId}`);
        const res = await response.json();

        if (res.status === 'success') {
            const r = res.record;
            document.getElementById('proofId').textContent = r.id;
            document.getElementById('proofCertDate').textContent = 'Tarih: ' + r.kayit_tarihi;
            const proofKargoEl = document.getElementById('proofKargoBarcode');
            if (proofKargoEl) proofKargoEl.textContent = r.kargo_barkodu || r.siparis_no || '-';
            document.getElementById('proofOrderNo').textContent = r.siparis_no || 'Belirtilmedi';
            document.getElementById('proofCustomer').textContent = r.musteri_adi || 'Belirtilmedi';
            document.getElementById('proofProduct').textContent = r.urun_adi || 'Genel Ürün';
            document.getElementById('proofMarketplace').textContent = r.pazaryeri || 'Doğrudan';
            document.getElementById('proofBarcode').textContent = r.barkod_seri_no;
            document.getElementById('proofTime').textContent = r.kayit_tarihi + ' (İstanbul UTC+3)';

            const m = document.getElementById('proofModal');
            if (m) m.style.display = 'flex';
        }
    } catch (err) {
        alert('Kanıt belgesi yüklenemedi: ' + err.message);
    }
};

// --------------------------------------------------------------
// 6. KAMERA İLE TARAMA
// --------------------------------------------------------------
// ==============================================================
// 6. KAMERA İLE AKILLI 2 ADIMLI CANLI TARAMA (ELLER SERBEST)
// ==============================================================
let isCameraProcessing = false;
let cameraCooldownUntil = 0;
let activeCamStep = 1;

function initCameraScanner() {
    const btnOrder = document.getElementById('btnToggleCamOrder');
    const btnFree = document.getElementById('btnToggleCamFree');
    const closeBtn = document.getElementById('btnCloseCamModal');
    const footerCloseBtn = document.getElementById('btnDismissCamFooter');
    const camSwitchStepBtn = document.getElementById('btnCamSwitchStep');

    if (btnOrder) btnOrder.addEventListener('click', () => startCameraScanner('order'));
    if (btnFree) btnFree.addEventListener('click', () => startCameraScanner('free'));
    if (closeBtn) closeBtn.addEventListener('click', stopCameraScanner);
    if (footerCloseBtn) footerCloseBtn.addEventListener('click', stopCameraScanner);
    if (camSwitchStepBtn) {
        camSwitchStepBtn.addEventListener('click', () => {
            resetToWaitingState();
            updateCameraUIForStep(1);
        });
    }
}

function updateCameraUIForStep(step) {
    activeCamStep = step;
    const pill = document.getElementById('camStepPill');
    const title = document.getElementById('camModalTitle');
    const banner = document.getElementById('camOrderBanner');
    const statusText = document.getElementById('camStatusText');
    const switchBtn = document.getElementById('btnCamSwitchStep');

    if (step === 1) {
        if (pill) {
            pill.className = 'cam-step-pill step1-pill';
            pill.textContent = '📦 1. Adım: Kargo Barkodu';
        }
        if (title) title.textContent = 'Kolideki Kargo Barkodunu Okutun';
        if (banner) banner.style.display = 'none';
        if (switchBtn) switchBtn.style.display = 'none';
        if (statusText) {
            statusText.className = 'camera-status status-step1';
            statusText.textContent = '📦 Kolideki kargo takip etiketini kameranın çerçevesine hizalayın...';
        }
    } else {
        const order = (activeOrderIndex !== -1 && loadedOrders[activeOrderIndex]) ? loadedOrders[activeOrderIndex] : null;
        if (pill) {
            pill.className = 'cam-step-pill step2-pill';
            pill.textContent = '🎯 2. Adım: Ürün Seri No';
        }
        if (title) {
            title.textContent = order ? `Ürün Okutma (${order.musteri_adi})` : 'Ürün Seri No / Barkod Okutun';
        }
        if (banner && order) {
            banner.style.display = 'block';
            banner.style.background = 'rgba(17, 24, 39, 0.85)';
            banner.style.borderColor = 'rgba(16, 185, 129, 0.35)';

            const cCust = document.getElementById('camBannerCustomer');
            const cKargo = document.getElementById('camBannerKargo');
            const cProd = document.getElementById('camBannerProduct');

            const pending = (activePendingItemIndex !== -1 && order.urunler[activePendingItemIndex] && !order.urunler[activePendingItemIndex].tamamlandi)
                ? order.urunler[activePendingItemIndex]
                : (order.urunler.find(i => !i.tamamlandi) || order.urunler[0]);
            const completedCount = order.urunler.filter(i => i.tamamlandi).length;

            if (cCust) cCust.textContent = `👤 ${order.musteri_adi}`;
            if (cKargo) cKargo.textContent = `🏷️ ${order.kargo_barkodu || order.siparis_no}`;
            if (cProd && pending) {
                const curScan = (pending.okunan_adet || 0) + 1;
                const qtyText = pending.adet > 1 ? ` [${curScan}/${pending.adet}. Adet]` : '';
                cProd.innerHTML = `<strong>Beklenen:</strong> ${pending.urun_adi.substring(0, 45)}${qtyText} (${completedCount}/${order.urunler.length} Kalem)`;
            }
        }
        if (switchBtn) switchBtn.style.display = 'inline-block';
        if (statusText) {
            statusText.className = 'camera-status status-step2';
            statusText.textContent = '🎯 Pakete koyduğunuz ürünün Seri No / Barkodunu kameraya tutun...';
        }
    }
}

function startCameraScanner(target = 'order') {
    activeCamTarget = target;
    const modal = document.getElementById('camModal');
    if (modal) modal.style.display = 'flex';

    if (target === 'order') {
        if (activeOrderIndex !== -1 && loadedOrders[activeOrderIndex]) {
            updateCameraUIForStep(2);
        } else {
            updateCameraUIForStep(1);
        }
    } else {
        const pill = document.getElementById('camStepPill');
        if (pill) {
            pill.className = 'cam-step-pill step1-pill';
            pill.textContent = '⚡ Hızlı Serbest Kayıt';
        }
        const title = document.getElementById('camModalTitle');
        if (title) title.textContent = 'Barkod / Seri No Okutun';
        const banner = document.getElementById('camOrderBanner');
        if (banner) banner.style.display = 'none';
        const statusText = document.getElementById('camStatusText');
        if (statusText) statusText.textContent = 'Barkod veya seri numarasını kamera çerçevesine tutun...';
    }

    launchHtml5QrCode();
}

function launchHtml5QrCode() {
    const statusText = document.getElementById('camStatusText');

    if (typeof Html5Qrcode === 'undefined') {
        if (statusText) statusText.textContent = 'Kamera kütüphanesi yüklenemedi. Lütfen sayfayı yenileyin.';
        return;
    }

    if (html5QrCode) {
        html5QrCode.stop().catch(() => {}).finally(() => {
            html5QrCode = null;
            execLaunchCamera();
        });
    } else {
        execLaunchCamera();
    }
}

function execLaunchCamera() {
    const statusText = document.getElementById('camStatusText');
    try {
        html5QrCode = new Html5Qrcode("cameraPreview");
        const config = {
            fps: 15,
            qrbox: { width: 280, height: 180 },
            aspectRatio: 1.333334
        };

        html5QrCode.start(
            { facingMode: "environment" },
            config,
            (decodedText) => {
                onCameraBarcodeDecoded(decodedText);
            },
            () => {}
        ).then(() => {
            if (statusText && !isCameraProcessing) {
                if (activeCamTarget === 'order') {
                    if (activeCamStep === 1) {
                        statusText.textContent = '📦 Kamera aktif: Kolideki kargo takip etiketini gösterin...';
                    } else {
                        statusText.textContent = '🎯 Kamera aktif: Ürünün seri numarasını gösterin...';
                    }
                } else {
                    statusText.textContent = 'Kamera aktif: Barkodu kutu içine hizalayın...';
                }
            }
        }).catch(err => {
            if (statusText) statusText.textContent = 'Kamera erişim hatası: ' + err;
        });
    } catch (e) {
        if (statusText) statusText.textContent = 'Kamera başlatılamadı: ' + e.message;
    }
}

async function onCameraBarcodeDecoded(decodedText) {
    const now = Date.now();
    if (isCameraProcessing || now < cameraCooldownUntil) {
        return;
    }
    const cleanCode = decodedText.trim();
    if (!cleanCode) return;

    isCameraProcessing = true;
    const statusText = document.getElementById('camStatusText');

    try {
        if (activeCamTarget === 'free') {
            stopCameraScanner();
            const freeInput = document.getElementById('freeScanInput');
            if (freeInput) {
                freeInput.value = cleanCode;
                executeFreeScan();
            }
            return;
        }

        // ActiveCamTarget === 'order'
        if (activeCamStep === 1 || activeOrderIndex === -1) {
            // STEP 1: Scan cargo barcode to find customer
            if (loadedOrders.length === 0) {
                playWarningBuzzer();
                if (statusText) statusText.textContent = '⚠️ Henüz sipariş listesi yüklenmedi! Lütfen yukarıdan Excel seçin.';
                cameraCooldownUntil = Date.now() + 2000;
                return;
            }

            const cleanAlnum = cleanCode.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            let matchedIdx = loadedOrders.findIndex(o => 
                (o.kargo_barkodu && o.kargo_barkodu.trim().toLowerCase() === cleanCode.toLowerCase()) ||
                (o.siparis_no && o.siparis_no.trim().toLowerCase() === cleanCode.toLowerCase()) ||
                (o.paket_no && o.paket_no.trim().toLowerCase() === cleanCode.toLowerCase())
            );

            if (matchedIdx === -1 && cleanAlnum) {
                matchedIdx = loadedOrders.findIndex(o => {
                    const kbAlnum = (o.kargo_barkodu || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                    const spAlnum = (o.siparis_no || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                    const pkAlnum = (o.paket_no || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                    return (kbAlnum && kbAlnum === cleanAlnum) || (spAlnum && spAlnum === cleanAlnum) || (pkAlnum && pkAlnum === cleanAlnum);
                });
            }

            if (matchedIdx === -1 && cleanCode.length >= 6) {
                matchedIdx = loadedOrders.findIndex(o => {
                    const kb = (o.kargo_barkodu || '').trim().toLowerCase();
                    const sp = (o.siparis_no || '').trim().toLowerCase();
                    return (kb && (kb.endsWith(cleanCode.toLowerCase()) || cleanCode.toLowerCase().endsWith(kb))) ||
                           (sp && (sp.endsWith(cleanCode.toLowerCase()) || cleanCode.toLowerCase().endsWith(sp)));
                });
            }

            if (matchedIdx !== -1) {
                // MATCHED! Customer found!
                playSuccessBeep();
                selectOrder(matchedIdx);
                const order = loadedOrders[matchedIdx];

                updateCameraUIForStep(2);
                if (statusText) {
                    statusText.innerHTML = `✅ <strong>MÜŞTERİ BULUNDU:</strong> ${order.musteri_adi}<br>🎯 Şimdi ürünün seri numarasını kameraya tutun!`;
                }

                // Give 1.5 seconds cooldown so the cargo barcode isn't instantly re-read
                cameraCooldownUntil = Date.now() + 1500;
            } else {
                playWarningBuzzer();
                if (statusText) {
                    statusText.textContent = `⚠️ Kargo barkodu listede bulunamadı: "${cleanCode}". Kolideki doğru barkodu tutun.`;
                }
                cameraCooldownUntil = Date.now() + 1500;
            }
        } else if (activeCamStep === 2) {
            // STEP 2: Scan product barcode/serial
            const order = loadedOrders[activeOrderIndex];

            // Safety check: Did user accidentally scan the cargo barcode again?
            const isSameKargo = (order.kargo_barkodu && order.kargo_barkodu.trim().toLowerCase() === cleanCode.toLowerCase()) ||
                                (order.siparis_no && order.siparis_no.trim().toLowerCase() === cleanCode.toLowerCase());
            if (isSameKargo) {
                playWarningBuzzer();
                if (statusText) statusText.textContent = `⚠️ Bu koli kargo barkodu! Lütfen kutudaki ÜRÜNÜN seri numarasını gösterin.`;
                cameraCooldownUntil = Date.now() + 1200;
                return;
            }

            const pendingItem = (activePendingItemIndex !== -1 && order.urunler[activePendingItemIndex] && !order.urunler[activePendingItemIndex].tamamlandi)
                ? order.urunler[activePendingItemIndex]
                : order.urunler.find(i => !i.tamamlandi);
            if (!pendingItem) {
                playSuccessBeep();
                if (statusText) statusText.textContent = `🎉 Bu paketin tüm ürünleri tamamlanmış! Sıradaki koliye geçiliyor...`;
                setTimeout(() => {
                    resetToWaitingState();
                    updateCameraUIForStep(1);
                }, 1200);
                cameraCooldownUntil = Date.now() + 1500;
                return;
            }

            if (statusText) statusText.textContent = `Ürün kaydediliyor: ${cleanCode}...`;

            const response = await fetch('/api/arsiv/record', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    kargo_barkodu: order.kargo_barkodu || '',
                    siparis_no: order.siparis_no || '',
                    musteri_adi: order.musteri_adi,
                    urun_adi: pendingItem.urun_adi,
                    barkod_seri_no: cleanCode,
                    pazaryeri: order.pazaryeri,
                    okutma_yontemi: 'Kamera / Canlı Tarama'
                })
            });
            const res = await response.json();

            if (res.status === 'success') {
                if (!pendingItem.seri_nolar) pendingItem.seri_nolar = [];
                pendingItem.seri_nolar.push(cleanCode);
                pendingItem.okunan_adet = (pendingItem.okunan_adet || 0) + 1;

                const isItemDone = pendingItem.okunan_adet >= (pendingItem.adet || 1);
                pendingItem.tamamlandi = isItemDone;

                if (isItemDone && activePendingItemIndex !== -1) {
                    activePendingItemIndex = -1;
                }

                renderActiveOrderItems();
                updateOrderProgress();
                renderOrderCards();
                updateScanCounters(cleanCode, pendingItem.urun_adi);
                updateHeaderStats(res.stats);
                saveOrdersToStorage(loadedFileName, loadedOrders);

                const allDone = order.urunler.every(i => i.tamamlandi);

                if (res.is_duplicate) {
                    playWarningBuzzer();
                    if (statusText) statusText.innerHTML = isItemDone
                        ? `⚠️ <strong>MÜKERRER:</strong> ${cleanCode} daha önce kaydedilmiş! Mühürlendi. [${pendingItem.okunan_adet}/${pendingItem.adet} Tamamlandı]`
                        : `⚠️ <strong>MÜKERRER:</strong> ${cleanCode} daha önce kaydedilmiş! [${pendingItem.okunan_adet}/${pendingItem.adet} Okundu - Sıradaki adedi tutun]`;
                } else {
                    playSuccessBeep();
                    if (statusText) statusText.innerHTML = isItemDone
                        ? `✅ <strong>TAMAMLANDI:</strong> ${pendingItem.urun_adi.substring(0, 25)}... (${cleanCode}) [${pendingItem.okunan_adet}/${pendingItem.adet}]`
                        : `✅ <strong>ADET OKUNDU:</strong> ${pendingItem.urun_adi.substring(0, 25)}... (${cleanCode}) [${pendingItem.okunan_adet}/${pendingItem.adet} - Sıradaki adedi tutun]`;
                }

                if (allDone) {
                    playSuccessBeep();
                    const banner = document.getElementById('camOrderBanner');
                    if (banner) {
                        banner.style.background = 'rgba(16, 185, 129, 0.35)';
                        banner.style.borderColor = '#10B981';
                    }
                    if (statusText) {
                        statusText.innerHTML = `🎉 <strong>TEBRİKLER!</strong> ${order.musteri_adi} paketi tamamlandı! 2 saniye sonra sıradaki koliye geçiliyor...`;
                    }
                    setTimeout(() => {
                        resetToWaitingState();
                        updateCameraUIForStep(1);
                    }, 2000);
                    cameraCooldownUntil = Date.now() + 2500;
                } else {
                    updateCameraUIForStep(2);
                    cameraCooldownUntil = Date.now() + 1400;
                }
            } else {
                playWarningBuzzer();
                if (statusText) statusText.textContent = 'Hata: ' + (res.message || 'Kaydedilemedi');
                cameraCooldownUntil = Date.now() + 1500;
            }
        }
    } catch (err) {
        playWarningBuzzer();
        if (statusText) statusText.textContent = 'Bağlantı hatası: ' + err.message;
        cameraCooldownUntil = Date.now() + 1500;
    } finally {
        isCameraProcessing = false;
    }
}

function stopCameraScanner() {
    const modal = document.getElementById('camModal');
    if (modal) modal.style.display = 'none';

    if (html5QrCode) {
        html5QrCode.stop().then(() => {
            try { html5QrCode.clear(); } catch (e) {}
            html5QrCode = null;
        }).catch(() => {
            html5QrCode = null;
        });
    }

    setTimeout(() => {
        const activeInput = document.getElementById('activeScanInput');
        if (activeInput) activeInput.focus();
    }, 100);
}

// --------------------------------------------------------------
// YARDIMCI FONKSİYONLAR
// --------------------------------------------------------------
function showAlert(el, type, message) {
    if (!el) return;
    el.className = `scan-feedback-alert ${type}`;
    el.textContent = message;
    el.style.display = 'block';

    setTimeout(() => {
        if (el && el.className.includes(type)) {
            el.style.display = 'none';
        }
    }, 6000);
}

function updateHeaderStats(stats) {
    if (!stats) return;
    if (stats.bugun_kayit !== undefined) {
        const el = document.getElementById('headerTodayCount');
        if (el) el.textContent = stats.bugun_kayit;
        const el2 = document.getElementById('statTodayScans');
        if (el2) el2.textContent = stats.bugun_kayit;
    }
    if (stats.toplam_kayit !== undefined) {
        const el = document.getElementById('statTotalScans');
        if (el) el.textContent = stats.toplam_kayit;
    }
}

async function refreshDashboardStats() {
    try {
        const response = await fetch('/api/arsiv/stats');
        const data = await response.json();
        updateHeaderStats(data);
    } catch (e) {
        console.error('Stats refresh error:', e);
    }
}


// ==============================================================
// GLOBAL WINDOW ERİŞİMLERİ (DOĞRUDAN HTML VE BUTON TIKLAMALARI İÇİN)
// ==============================================================
window.switchArsivTab = function(targetId) {
    if (!targetId) return;
    document.querySelectorAll('#view-arsiv .nav-tab').forEach(t => {
        if (t.dataset.tab === targetId) t.classList.add('active');
        else t.classList.remove('active');
    });
    document.querySelectorAll('#view-arsiv .tab-pane').forEach(p => {
        if (p.id === targetId) p.classList.add('active');
        else p.classList.remove('active');
    });
    if (targetId === 'tab-siparis') {
        const inp = document.getElementById('activeScanInput');
        if (inp) inp.focus();
    } else if (targetId === 'tab-paketler') {
        const f = document.getElementById('orderFilterInput');
        if (f) f.focus();
        if (typeof renderOrderCards === 'function') {
            renderOrderCards(f ? f.value.trim().toLowerCase() : '');
        }
    } else if (targetId === 'tab-serbest') {
        const inp = document.getElementById('freeScanInput');
        if (inp) inp.focus();
    } else if (targetId === 'tab-arama') {
        const inp = document.getElementById('archiveSearchInput');
        if (inp) inp.focus();
    }
};
window.switchTab = window.switchArsivTab;


window.handleFileInputChange = function(files) {
    if (files && files.length > 0) {
        handleFileUpload(files);
    }
};

window.loadRecentExcel = async function() {
    const recentBtn = document.getElementById('btnLoadRecentExcel');
    const orig = recentBtn ? recentBtn.textContent : '⚡ İndirilenler\'den Getir';
    if (recentBtn) recentBtn.textContent = 'Yükleniyor...';
    try {
        const response = await fetch('/api/arsiv/load-recent-download', { method: 'POST' });
        const res = await response.json();
        if (res.status === 'success') {
            loadedOrders = res.orders || [];
            loadedFileName = res.filename || 'İndirilenler Sipariş Listesi';
            saveOrdersToStorage(loadedFileName, loadedOrders);
            updateOrderUIElements(loadedFileName, loadedOrders.length);
            renderOrderCards();
            resetToWaitingState();
            const alertBox = document.getElementById('orderScanAlert');
            showAlert(alertBox, 'success', `📦 ${loadedOrders.length} Paket yüklendi ve sabitlendi! Kolideki KARGO BARKODUNU okutun.`);
            playSuccessBeep();
        } else {
            alert('Hata: ' + (res.message || 'Dosya okunamadı. Lütfen "Dosya Seç" ile yükleyin.'));
            playWarningBuzzer();
        }
    } catch (err) {
        alert('İndirilenler klasörüne doğrudan erişim yerel sunucuda (127.0.0.1:5004) çalışır. Bulutta dosyanızı "Dosya Seç" butonu ile yükleyebilirsiniz.');
        playWarningBuzzer();
    } finally {
        if (recentBtn) recentBtn.textContent = orig;
    }
};

window.selectOrder = selectOrder;
window.handleSmartScan = handleSmartScan;
window.processProductSerialScan = processProductSerialScan;
window.executeFreeScan = executeFreeScan;
window.executeArchiveSearch = executeArchiveSearch;
window.startCameraScanner = startCameraScanner;
window.stopCameraScanner = stopCameraScanner;
window.setScanStep = setScanStep;
window.refreshDashboardStats = refreshDashboardStats;

window.resetToWaitingState = resetToWaitingState;

window.updateCameraUIForStep = updateCameraUIForStep;


window.initArsivModule = function() {
    if (typeof initApp === 'function') initApp();
};


// --------------------------------------------------------------
// YANLIŞ OKUTULAN BARKODU SİLME VE DÜZENLEME FONKSİYONLARI
// --------------------------------------------------------------
async function deleteScannedSerial(itemIdx, serial) {
    if (activeOrderIndex === -1) return;
    const order = loadedOrders[activeOrderIndex];
    if (!order || !order.urunler || !order.urunler[itemIdx]) return;
    const item = order.urunler[itemIdx];

    const confirmed = confirm(`"${serial}" seri numarasını/barkodunu bu üründen silmek istiyor musunuz?\n\nSilindiğinde bu ürün için adedi yeniden okutmanız gerekecektir.`);
    if (!confirmed) return;

    try {
        const response = await fetch('/api/arsiv/remove-barcode', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                kargo_barkodu: order.kargo_barkodu || '',
                siparis_no: order.siparis_no || '',
                barkod_seri_no: serial
            })
        });
        const res = await response.json();

        // Yerel durumdan seriyi çıkar
        const snIndex = (item.seri_nolar || []).indexOf(serial);
        if (snIndex !== -1) {
            item.seri_nolar.splice(snIndex, 1);
        }
        item.okunan_adet = Math.max(0, (item.okunan_adet || 1) - 1);
        item.tamamlandi = false;
        order.tamamlandi = false;

        // Hedef kalemi bu eksilen kaleme sabitle
        activePendingItemIndex = itemIdx;

        renderActiveOrderItems();
        updateOrderProgress();
        renderOrderCards();
        if (res.stats) updateHeaderStats(res.stats);
        saveOrdersToStorage(loadedFileName, loadedOrders);

        setScanStep(2);

        const alertBox = document.getElementById('orderScanAlert');
        showAlert(alertBox, 'success', `🗑️ Seri No ("${serial}") başarıyla silindi! Sıradaki doğru seri numarasını okutabilirsiniz.`);
        playSuccessBeep();
    } catch (err) {
        alert('Silme sırasında hata oluştu: ' + err.message);
    }
}

async function editScannedSerial(itemIdx, serial) {
    if (activeOrderIndex === -1) return;
    const order = loadedOrders[activeOrderIndex];
    if (!order || !order.urunler || !order.urunler[itemIdx]) return;
    const item = order.urunler[itemIdx];

    const newSerial = prompt(`"${item.urun_adi}" kalemi için yanlış okunan seri numarası:\n${serial}\n\nLütfen yeni / doğru seri numarasını girin veya okutun:`, serial);
    if (!newSerial || newSerial.trim() === '' || newSerial.trim() === serial) return;

    const cleanNew = newSerial.trim();

    try {
        const response = await fetch('/api/arsiv/update-barcode', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                kargo_barkodu: order.kargo_barkodu || '',
                siparis_no: order.siparis_no || '',
                old_barkod_seri_no: serial,
                new_barkod_seri_no: cleanNew
            })
        });
        const res = await response.json();

        if (res.status === 'success') {
            const snIndex = (item.seri_nolar || []).indexOf(serial);
            if (snIndex !== -1) {
                item.seri_nolar[snIndex] = cleanNew;
            } else {
                item.seri_nolar.push(cleanNew);
            }

            renderActiveOrderItems();
            updateOrderProgress();
            renderOrderCards();
            saveOrdersToStorage(loadedFileName, loadedOrders);

            const alertBox = document.getElementById('orderScanAlert');
            showAlert(alertBox, 'success', `✏️ Seri No güncellendi: "${cleanNew}"`);
            playSuccessBeep();
        } else {
            alert('Hata: ' + (res.message || 'Güncellenemedi'));
        }
    } catch (err) {
        alert('Güncelleme hatası: ' + err.message);
    }
}

async function deleteArchiveRecord(recordId, barcode) {
    const confirmed = confirm(`#${recordId} numaralı kaydı ("${barcode}") arşivden tamamen silmek istiyor musunuz?`);
    if (!confirmed) return;

    try {
        const response = await fetch(`/api/arsiv/delete-record/${recordId}`, { method: 'POST' });
        const res = await response.json();
        if (res.status === 'success') {
            alert('Kayıt başarıyla silindi.');
            if (res.stats) updateHeaderStats(res.stats);
            if (typeof executeArchiveSearch === 'function') executeArchiveSearch();
        } else {
            alert('Hata: ' + (res.message || 'Silinemedi'));
        }
    } catch (err) {
        alert('Bağlantı hatası: ' + err.message);
    }
}

async function openEditRecordModal(recordId) {
    try {
        const response = await fetch(`/api/arsiv/proof/${recordId}`);
        const res = await response.json();
        if (res.status === 'success' && res.record) {
            const r = res.record;
            const elId = document.getElementById('editRecordId');
            const elBc = document.getElementById('editBarcodeVal');
            const elKb = document.getElementById('editKargoBarcodeVal');
            const elOn = document.getElementById('editOrderNoVal');
            const elCu = document.getElementById('editCustomerVal');
            const elPr = document.getElementById('editProductVal');

            if (elId) elId.value = r.id;
            if (elBc) elBc.value = r.barkod_seri_no || '';
            if (elKb) elKb.value = r.kargo_barkodu || '';
            if (elOn) elOn.value = r.siparis_no || '';
            if (elCu) elCu.value = r.musteri_adi || '';
            if (elPr) elPr.value = r.urun_adi || '';

            const modal = document.getElementById('editRecordModal');
            if (modal) modal.style.display = 'flex';
            if (elBc) elBc.focus();
        } else {
            alert('Kayıt detayları alınamadı.');
        }
    } catch (err) {
        alert('Hata: ' + err.message);
    }
}

function closeEditRecordModal() {
    const modal = document.getElementById('editRecordModal');
    if (modal) modal.style.display = 'none';
}

async function saveRecordEdit() {
    const recordId = document.getElementById('editRecordId')?.value;
    const barcode = document.getElementById('editBarcodeVal')?.value.trim();
    if (!recordId || !barcode) {
        alert('Ürün Seri No / Barkod boş olamaz!');
        return;
    }

    const payload = {
        barkod_seri_no: barcode,
        kargo_barkodu: document.getElementById('editKargoBarcodeVal')?.value.trim() || '',
        siparis_no: document.getElementById('editOrderNoVal')?.value.trim() || '',
        musteri_adi: document.getElementById('editCustomerVal')?.value.trim() || '',
        urun_adi: document.getElementById('editProductVal')?.value.trim() || ''
    };

    try {
        const response = await fetch(`/api/arsiv/update-record/${recordId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        const res = await response.json();
        if (res.status === 'success') {
            closeEditRecordModal();
            alert('Kayıt başarıyla güncellendi!');
            if (typeof executeArchiveSearch === 'function') executeArchiveSearch();
        } else {
            alert('Hata: ' + (res.message || 'Güncellenemedi'));
        }
    } catch (err) {
        alert('Bağlantı hatası: ' + err.message);
    }
}

window.deleteScannedSerial = deleteScannedSerial;
window.editScannedSerial = editScannedSerial;
window.deleteArchiveRecord = deleteArchiveRecord;
window.openEditRecordModal = openEditRecordModal;
window.closeEditRecordModal = closeEditRecordModal;
window.saveRecordEdit = saveRecordEdit;
