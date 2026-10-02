// ==========================================================================
// OLİZ KAMPANYA & İNDİRİM ANALİZ SİSTEMİ - İNTERAKTİF JS MOTORU
// ==========================================================================

let activeSlots = ["", "", "", ""];
let catalogData = [];
let currentCatalogType = "tekil";

document.addEventListener("DOMContentLoaded", () => {
    initSlots();
    initChips();
    initCatalogTabs();
    initModal();
    loadCatalog("tekil");
});

// Format currency
function formatTL(num) {
    if (num === null || num === undefined || isNaN(num)) return "0 TL";
    return Number(num).toLocaleString("tr-TR", { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + " TL";
}

// ==========================================================================
// 4 SLOTS & AUTOCOMPLETE
// ==========================================================================
function initSlots() {
    for (let i = 1; i <= 4; i++) {
        const input = document.getElementById(`slot-input-${i}`);
        const dropdown = document.getElementById(`autocomplete-${i}`);
        const clearBtn = document.getElementById(`clear-slot-${i}`);

        if (!input) continue;

        let debounceTimer = null;

        input.addEventListener("input", (e) => {
            const val = e.target.value.trim();
            activeSlots[i - 1] = val;

            clearTimeout(debounceTimer);
            if (val.length < 2) {
                dropdown.classList.remove("active");
                dropdown.innerHTML = "";
                return;
            }

            debounceTimer = setTimeout(() => {
                fetchAutocomplete(val, dropdown, i);
            }, 250);
        });

        input.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
                e.preventDefault();
                dropdown.classList.remove("active");
                runAnalysis();
            }
        });

        clearBtn.addEventListener("click", () => {
            clearSlot(i);
        });
    }

    // Close dropdowns on outside click
    document.addEventListener("click", (e) => {
        if (!e.target.closest(".input-wrapper")) {
            document.querySelectorAll(".autocomplete-dropdown").forEach(d => d.classList.remove("active"));
        }
    });

    // Analyze Button
    const analyzeBtn = document.getElementById("analyze-btn");
    if (analyzeBtn) {
        analyzeBtn.addEventListener("click", runAnalysis);
    }

    // Clear All Button
    const clearAllBtn = document.getElementById("clear-all-btn");
    if (clearAllBtn) {
        clearAllBtn.addEventListener("click", clearAllSlots);
    }

    // Print Button
    const printBtn = document.getElementById("print-btn");
    if (printBtn) {
        printBtn.addEventListener("click", () => window.print());
    }
}

async function fetchAutocomplete(query, dropdownEl, slotIndex) {
    try {
        const res = await fetch(`/api/autocomplete?q=${encodeURIComponent(query)}&limit=12`);
        const data = await res.json();
        if (!data.success || !data.results || data.results.length === 0) {
            dropdownEl.classList.remove("active");
            return;
        }

        dropdownEl.innerHTML = data.results.map(item => `
            <div class="autocomplete-item" data-sku="${item.sku}" data-name="${item.name}">
                <div class="item-title">${item.name}</div>
                <div class="item-meta">
                    <span class="item-sku">SKU: ${item.sku}</span>
                    <span class="item-brand">${item.brand} • ${item.group}</span>
                </div>
            </div>
        `).join("");

        dropdownEl.classList.add("active");

        dropdownEl.querySelectorAll(".autocomplete-item").forEach(item => {
            item.addEventListener("click", () => {
                const sku = item.getAttribute("data-sku");
                const name = item.getAttribute("data-name");
                selectProductForSlot(slotIndex, sku, name);
                dropdownEl.classList.remove("active");
            });
        });
    } catch (err) {
        console.error("Autocomplete error:", err);
    }
}

function selectProductForSlot(slotIndex, sku, name) {
    const input = document.getElementById(`slot-input-${slotIndex}`);
    const preview = document.getElementById(`preview-${slotIndex}`);
    const previewName = document.getElementById(`preview-name-${slotIndex}`);

    if (input) input.value = sku;
    activeSlots[slotIndex - 1] = sku;

    if (preview && previewName) {
        previewName.textContent = name;
        preview.classList.add("active");
    }

    runAnalysis();
}

function clearSlot(slotIndex) {
    const input = document.getElementById(`slot-input-${slotIndex}`);
    const preview = document.getElementById(`preview-${slotIndex}`);
    const dropdown = document.getElementById(`autocomplete-${slotIndex}`);

    if (input) input.value = "";
    if (preview) preview.classList.remove("active");
    if (dropdown) dropdown.classList.remove("active");
    activeSlots[slotIndex - 1] = "";

    runAnalysis();
}

function clearAllSlots() {
    for (let i = 1; i <= 4; i++) {
        clearSlot(i);
    }
    const container = document.getElementById("results-container");
    if (container) container.classList.remove("active");
}

// ==========================================================================
// RUN BUNDLE & INDIVIDUAL ANALYSIS (1, 2, 3, 4)
// ==========================================================================
async function runAnalysis() {
    const filledProducts = activeSlots.filter(s => s && s.trim().length > 0);
    const container = document.getElementById("results-container");

    if (filledProducts.length === 0) {
        if (container) container.classList.remove("active");
        return;
    }

    try {
        const res = await fetch("/api/oliz/analyze", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ products: activeSlots })
        });
        const json = await res.json();
        if (json.success) {
            renderResults(json.data);
            if (container) {
                container.classList.add("active");
                container.scrollIntoView({ behavior: "smooth", block: "start" });
            }
        }
    } catch (err) {
        console.error("Analysis error:", err);
    }
}

function renderResults(data) {
    renderIndividualCards(data.individual_results);
    renderMatchedPackages(data.matched_packages);
    renderSummaryComparison(data);
}

// Render 1, 2, 3, 4 Individual Cards
function renderIndividualCards(items) {
    const grid = document.getElementById("individual-results-grid");
    if (!grid) return;

    grid.innerHTML = items.map(item => {
        const slot = item.slot;
        const hasTekil = item.has_tekil;
        const discVal = item.tekil_discount;
        const coupVal = item.coupon_val;

        return `
            <div class="ind-card slot-${slot}">
                <div>
                    <div class="ind-card-header">
                        <span class="slot-badge">Ürün ${slot}</span>
                        <span style="font-family: monospace; font-size: 11px; color: #38BDF8;">SKU: ${item.sku}</span>
                    </div>
                    <div class="ind-card-title">${item.name}</div>
                    <div class="ind-card-meta">${item.brand} • ${item.group}</div>

                    ${hasTekil ? `
                        <div class="ind-discount-badge">
                            <div class="ind-discount-val">${formatTL(discVal)}</div>
                            <div class="ind-discount-label">TEKİL KAMPANYA İNDİRİMİ</div>
                        </div>
                        <div class="ind-breakdown">
                            <div class="breakdown-row">
                                <span>Arçelik Katkısı:</span>
                                <span class="val" style="color: #60A5FA;">${formatTL(item.arcelik_support)}</span>
                            </div>
                            <div class="breakdown-row">
                                <span>Bayi Katkı Payı:</span>
                                <span class="val" style="color: #F87171;">${formatTL(item.bayi_support)}</span>
                            </div>
                            ${item.camp_code ? `
                            <div class="breakdown-row">
                                <span>Kampanya Kodu:</span>
                                <span class="val">${item.camp_code}</span>
                            </div>` : ''}
                        </div>
                    ` : `
                        <div style="background: rgba(255,255,255,0.03); border: 1px dashed rgba(255,255,255,0.1); border-radius: 8px; padding: 14px; text-align: center; margin-bottom: 12px; color: var(--text-dim); font-size: 12px;">
                            Tekil İndirim Tanımlı Değil
                        </div>
                    `}
                </div>

                <div>
                    ${item.has_toptan ? `
                        <div class="ind-coupon-badge">
                            <span>🎁 Toptan Kupon:</span>
                            <strong>${formatTL(coupVal)}</strong>
                        </div>
                    ` : ''}
                    
                    <div style="margin-top: 10px; font-size: 11.5px; color: var(--text-dim);">
                        <span>Dahil Olabileceği Paket Fırsatları: </span>
                        <strong style="color: #C4B5FD;">${item.eligible_packages_count} Adet</strong>
                    </div>
                </div>
            </div>
        `;
    }).join("");
}

// Render Matched Package Banner
function renderMatchedPackages(matched) {
    const container = document.getElementById("matched-packages-container");
    if (!container) return;

    if (!matched || matched.length === 0) {
        container.innerHTML = `
            <div style="background: rgba(255,255,255,0.02); border: 1px dashed var(--border-subtle); border-radius: var(--radius-md); padding: 18px 24px; text-align: center; margin-bottom: 24px;">
                <span style="color: var(--text-muted); font-size: 13.5px;">
                    ℹ️ Girilen ürün kombinasyonu henüz aktif bir çoklu paket kampanyasını (2'li, 3'lü veya 4'lü) tamamlamadı. İlgili diğer ürünleri ekleyerek paket fırsatlarını yakalayabilirsiniz.
                </span>
            </div>
        `;
        return;
    }

    container.innerHTML = matched.map(pkg => `
        <div class="package-match-card">
            <div class="pkg-banner-header">
                <div class="pkg-tag">
                    <span>🎉</span> ${pkg.required_items_count}'LÜ PAKET KAMPANYASI YAKALANDI
                </div>
                ${pkg.camp_code ? `<span class="badge badge-info">Kod: ${pkg.camp_code}</span>` : ''}
            </div>

            <div class="pkg-title">${pkg.title}</div>
            <div class="pkg-desc">${pkg.desc}</div>

            <div class="pkg-financials-grid">
                <div class="fin-item">
                    <span class="label">Toplam Paket İndirimi</span>
                    <span class="amount green">${formatTL(pkg.total_discount)}</span>
                </div>
                <div class="fin-item">
                    <span class="label">Arçelik Katılım Payı</span>
                    <span class="amount blue">${formatTL(pkg.arcelik_share)}</span>
                </div>
                <div class="fin-item">
                    <span class="label">Bayi Maliyet Payı</span>
                    <span class="amount purple">${formatTL(pkg.bayi_share)}</span>
                </div>
            </div>
        </div>
    `).join("");
}

// Render Comparison Table & Recommendation
function renderSummaryComparison(data) {
    const compTable = document.getElementById("comparison-table-body");
    const recBox = document.getElementById("recommendation-box-content");

    if (compTable) {
        compTable.innerHTML = `
            <tr>
                <td><strong>1. Tekil İndirimler Toplamı</strong></td>
                <td>${data.input_count} Ürün Ayrı Ayrı</td>
                <td><strong style="color: #34D399;">${formatTL(data.total_single_discount)}</strong></td>
                <td>${formatTL(data.total_arcelik_single)}</td>
                <td>${formatTL(data.total_bayi_single)}</td>
            </tr>
            <tr>
                <td><strong>2. Paket Kampanyası İndirimi</strong></td>
                <td>${data.best_package ? data.best_package.title : 'Eşleşen Paket Yok'}</td>
                <td><strong style="color: #60A5FA;">${formatTL(data.best_package ? data.best_package.total_discount : 0)}</strong></td>
                <td>${formatTL(data.best_package ? data.best_package.arcelik_share : 0)}</td>
                <td>${formatTL(data.best_package ? data.best_package.bayi_share : 0)}</td>
            </tr>
            <tr>
                <td><strong>3. Toptan Kupon Yüklemesi</strong></td>
                <td>Hedef Grubu Yükleme</td>
                <td><strong style="color: #FBBF24;">${formatTL(data.total_coupon_val)}</strong></td>
                <td>-</td>
                <td>-</td>
            </tr>
        `;
    }

    if (recBox) {
        const rec = data.recommendation;
        recBox.innerHTML = `
            <div class="rec-title">
                <span>💡</span> EN AVANTAJLI STRATEJİ
            </div>
            <div class="rec-msg">
                ${rec.message}
            </div>
            <div class="rec-total-badge">
                <div class="rec-total-val">${formatTL(rec.total_discount)}</div>
                <div class="rec-total-label">MAKSİMUM KAMPANYA İNDİRİMİ</div>
            </div>
        `;
    }
}

// ==========================================================================
// QUICK CHIPS (DEMO SENARYOLARI)
// ==========================================================================
function initChips() {
    const chips = [
        {
            id: "chip-ankastre-4",
            items: ["7768220221", "7751620235", "7703330203", "7182570230"] // Fırın + Ocak + Davlumbaz + Kurutucu
        },
        {
            id: "chip-tekil-buzdolabi",
            items: ["7295520276", "", "", ""] // 25.000 TL Tekil Soğutucu
        },
        {
            id: "chip-camasir-utu",
            items: ["7178520100", "8838321200", "", ""] // Çamaşır + Ütü
        },
        {
            id: "chip-tv-kea",
            items: ["B65 O 990 B", "TKM 8961", "", ""] // TV + KEA
        }
    ];

    chips.forEach(chip => {
        const btn = document.getElementById(chip.id);
        if (btn) {
            btn.addEventListener("click", () => {
                clearAllSlots();
                chip.items.forEach((code, idx) => {
                    if (code) {
                        const input = document.getElementById(`slot-input-${idx + 1}`);
                        if (input) input.value = code;
                        activeSlots[idx] = code;
                    }
                });
                runAnalysis();
            });
        }
    });
}

// ==========================================================================
// CATALOG EXPLORER (TEKİL, PAKET, TOPTAN)
// ==========================================================================
function initCatalogTabs() {
    const tabs = document.querySelectorAll(".nav-tab");
    tabs.forEach(tab => {
        tab.addEventListener("click", () => {
            tabs.forEach(t => t.classList.remove("active"));
            tab.classList.add("active");
            currentCatalogType = tab.getAttribute("data-type");
            loadCatalog(currentCatalogType);
        });
    });

    const searchInput = document.getElementById("catalog-search-input");
    if (searchInput) {
        searchInput.addEventListener("input", (e) => {
            filterCatalog(e.target.value.trim().toLowerCase());
        });
    }
}

async function loadCatalog(type) {
    try {
        const res = await fetch(`/api/catalog?type=${type}`);
        const data = await res.json();
        if (data.success) {
            catalogData = data.items;
            renderCatalogTable(catalogData, type);
        }
    } catch (err) {
        console.error("Catalog load error:", err);
    }
}

function renderCatalogTable(items, type) {
    const thead = document.getElementById("catalog-thead");
    const tbody = document.getElementById("catalog-tbody");
    if (!thead || !tbody) return;

    if (type === "tekil") {
        thead.innerHTML = `
            <tr>
                <th>SKU</th>
                <th>Ürün Tanımı / Model</th>
                <th>Marka & Grup</th>
                <th>İndirim Tutarı</th>
                <th>Arçelik Katkısı</th>
                <th>Bayi Payı</th>
                <th>Kampanya Tanımı</th>
                <th>İşlem</th>
            </tr>
        `;
        tbody.innerHTML = items.map(it => `
            <tr>
                <td><strong style="color: #38BDF8; font-family: monospace;">${it.sku}</strong></td>
                <td><strong>${it.name}</strong></td>
                <td>${it.brand} • ${it.group}</td>
                <td><span style="color: #34D399; font-weight: 700;">${formatTL(it.discount)}</span></td>
                <td>${formatTL(it.arcelik_support)}</td>
                <td>${formatTL(it.bayi_support)}</td>
                <td style="font-size: 11.5px; color: var(--text-muted);">${it.camp_desc}</td>
                <td>
                    <button class="btn btn-secondary" style="padding: 4px 10px; font-size: 11.5px;" onclick="addFirstEmptySlot('${it.sku}')">
                        + Ekle
                    </button>
                </td>
            </tr>
        `).join("");
    } else if (type === "toptan") {
        thead.innerHTML = `
            <tr>
                <th>SKU</th>
                <th>Ürün Tanımı</th>
                <th>Grup</th>
                <th>Kupon Yüklenecek Tutar</th>
                <th>Hedef Grubu</th>
                <th>Kampanya</th>
                <th>İşlem</th>
            </tr>
        `;
        tbody.innerHTML = items.map(it => `
            <tr>
                <td><strong style="color: #38BDF8; font-family: monospace;">${it.sku}</strong></td>
                <td><strong>${it.name}</strong></td>
                <td>${it.group}</td>
                <td><span style="color: #FBBF24; font-weight: 700;">${formatTL(it.coupon_val)}</span></td>
                <td><span class="badge badge-info">${it.target_group}</span></td>
                <td style="font-size: 11.5px; color: var(--text-muted);">${it.camp_desc}</td>
                <td>
                    <button class="btn btn-secondary" style="padding: 4px 10px; font-size: 11.5px;" onclick="addFirstEmptySlot('${it.sku}')">
                        + Ekle
                    </button>
                </td>
            </tr>
        `).join("");
    } else if (type === "paket") {
        thead.innerHTML = `
            <tr>
                <th>Kampanya Başlığı</th>
                <th>Toplam İndirim</th>
                <th>Bayi Payı</th>
                <th>Arçelik Payı</th>
                <th>Kampanya Kodu</th>
                <th>Açıklama / Şartlar</th>
            </tr>
        `;
        tbody.innerHTML = items.map(it => `
            <tr>
                <td><strong>${it.title}</strong></td>
                <td><span style="color: #34D399; font-weight: 700;">${formatTL(it.total_discount)}</span></td>
                <td>${formatTL(it.bayi_share)}</td>
                <td>${formatTL(it.arcelik_share)}</td>
                <td><span class="badge badge-info">${it.camp_code || '-'}</span></td>
                <td style="font-size: 12px; color: var(--text-muted); max-width: 320px;">${it.desc}</td>
            </tr>
        `).join("");
    }
}

function filterCatalog(term) {
    if (!term) {
        renderCatalogTable(catalogData, currentCatalogType);
        return;
    }
    const filtered = catalogData.filter(it => {
        const str = JSON.stringify(it).toLowerCase();
        return str.includes(term);
    });
    renderCatalogTable(filtered, currentCatalogType);
}

function addFirstEmptySlot(sku) {
    for (let i = 0; i < 4; i++) {
        if (!activeSlots[i]) {
            selectProductForSlot(i + 1, sku, sku);
            window.scrollTo({ top: 120, behavior: "smooth" });
            return;
        }
    }
    // If all full, replace slot 1
    selectProductForSlot(1, sku, sku);
    window.scrollTo({ top: 120, behavior: "smooth" });
}

// ==========================================================================
// EXCEL UPLOAD MODAL
// ==========================================================================
function initModal() {
    const modal = document.getElementById("upload-modal");
    const openBtn = document.getElementById("open-upload-modal");
    const closeBtn = document.getElementById("close-modal-btn");
    const dropzone = document.getElementById("file-dropzone");
    const fileInput = document.getElementById("oliz-file-input");

    if (openBtn && modal) {
        openBtn.addEventListener("click", () => modal.classList.add("active"));
    }
    if (closeBtn && modal) {
        closeBtn.addEventListener("click", () => modal.classList.remove("active"));
    }

    if (dropzone && fileInput) {
        dropzone.addEventListener("click", () => fileInput.click());

        dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
            dropzone.classList.add("dragover");
        });

        dropzone.addEventListener("dragleave", () => {
            dropzone.classList.remove("dragover");
        });

        dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
            dropzone.classList.remove("dragover");
            if (e.dataTransfer.files.length > 0) {
                uploadExcel(e.dataTransfer.files[0]);
            }
        });

        fileInput.addEventListener("change", () => {
            if (fileInput.files.length > 0) {
                uploadExcel(fileInput.files[0]);
            }
        });
    }
}

async function uploadExcel(file) {
    const formData = new FormData();
    formData.append("file", file);

    const statusEl = document.getElementById("upload-status");
    if (statusEl) {
        statusEl.innerHTML = `<span style="color: #38BDF8;">⏳ '${file.name}' yükleniyor ve kampanya analiz motoruna aktarılıyor...</span>`;
    }

    try {
        const res = await fetch("/api/oliz/upload", {
            method: "POST",
            body: formData
        });
        const json = await res.json();
        if (json.success) {
            statusEl.innerHTML = `<span style="color: #34D399;">✅ ${json.message}</span>`;
            setTimeout(() => {
                window.location.reload();
            }, 1200);
        } else {
            statusEl.innerHTML = `<span style="color: #EF4444;">❌ ${json.message}</span>`;
        }
    } catch (err) {
        if (statusEl) statusEl.innerHTML = `<span style="color: #EF4444;">❌ Yükleme hatası: ${err.message}</span>`;
    }
}