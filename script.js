document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const syncBtn = document.getElementById('syncData');
    const downloadExcelBtn = document.getElementById('downloadExcel');
    const applyFiltersBtn = document.getElementById('applyFilters');
    const productFilter = document.getElementById('productFilter');
    const totalRecordsEl = document.getElementById('totalRecords');
    const lumpsComplianceEl = document.getElementById('lumpsCompliance');
    const leakageComplianceEl = document.getElementById('leakageCompliance');
    const sealComplianceEl = document.getElementById('sealCompliance');
    const materialComplianceEl = document.getElementById('materialCompliance');
    const sizeComplianceEl = document.getElementById('sizeCompliance');
    const oxygenComplianceEl = document.getElementById('oxygenCompliance');

    const lumpsCountEl = document.getElementById('lumpsCount');
    const leakageCountEl = document.getElementById('leakageCount');
    const sealCountEl = document.getElementById('sealCount');
    const materialCountEl = document.getElementById('materialCount');
    const sizeCountEl = document.getElementById('sizeCount');
    const oxygenCountEl = document.getElementById('oxygenCount');

    const lastUpdatedEl = document.getElementById('lastUpdated');
    const navItems = document.querySelectorAll('.nav-item');
    const tabContents = document.querySelectorAll('.tab-content');
    const dashboardTableBody = document.getElementById('dashboardTableBody');
    const nitrogenFilter = document.getElementById('nitrogenFilter');
    const datePreset = document.getElementById('datePreset');
    const shiftFilter = document.getElementById('shiftFilter');
    const checkedByFilter = document.getElementById('checkedByFilter');

    let allData = [];
    let filteredData = [];
    let complianceChart;

    // --- Configuration ---
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzNVAaFAR1aMuJlk9vUWljryMv9HgRvGendH_mmqweCYp6dV0q48l-49_myQRudgwBo/exec';

    // --- Product Oxygen % Target Ranges ---
    const productRanges = {
        "Popular California Almonds Farmley Standee Pouch 1 kg": [8, 10],
        "Premium California Almonds Farmley Standee Pouch 1 Kg": [8, 10],
        "Premium W320 Cashew Farmley Standee Pouch 1 Kg": [8, 10],
        "Premium Raisin Long Farmley Standee Pouch 1 kg": [10, 12],
        "Chia Seeds Farmley Standee Pouch 120 g": [8, 10],
        "Premium Chia Seeds Farmley Standee Pouch 200 g": [8, 10],
        "Chia seeds Farmley Standee Pouch 300 g": [8, 10],
        "Chia Seed Farmley Standee Pouch 500g": [8, 10],
        "Quinoa seeds Farmley Standee Pouch 500g": [8, 10],
        "Seed mix Farmley 160g Standee Pouches": [8, 10],
        "Seed Mix Farmley Standee Pouch 200 g": [8, 10],
        "Basil Seeds Farmley Standee pouch 300g": [8, 10],
        "Premium Raisin Long Farmley Standee Pouch 500 g": [5, 8],
        "Premium Raisin Long Farmley Standee Pouch 200 g": [5, 8],
        "Roasted Flax Seeds Farmley Standee Pouch 150g": [5, 8],
        "Premium Flax Seeds Farmley Standee Pouch 200 g": [5, 8],
        "Premium Jumbo Pumpkin Seeds Farmley Standee Pouch 200 g": [5, 8],
        "Pumpkin seeds Farmley Standee Pouch 300 g": [5, 8],
        "Premium Sunflower Seeds Farmley Standee Pouch 200 g": [5, 8],
        "Sunflower seeds Farmley Standee Pouch 300 g": [5, 8],
        "Popular W400 Cashew Farmley Generic Pouch 500 g": [5, 8],
        "Watermelon Seeds Farmley Standee Pouch 500g": [5, 8],
        "Premium Broken Walnut Kernels Farmley Standee Pouch 250 g": [3, 5],
        "Premium Broken Walnut Kernels Farmley Standee Pouch 200 g": [3, 5],
        "Premium Extra Light Halves Walnut Kernels Farmley Standee Pouch 200 g": [3, 5],
        "Popular California Almonds Farmley Standee Pouch 500 g": [3, 5],
        "Popular California Almonds Farmley Standee Pouch 250 g": [3, 5],
        "Popular Almond Farmley Standee Pouch 60g": [3, 5],
        "Premium California Almonds Farmley Standee Pouch 500 g": [3, 5],
        "Premium California Almonds Farmley Standee Pouch 250 g": [3, 5],
        "Roasted and Salted Cashew - Farmley Standee Pouch 200 g": [3, 5],
        "Roasted & Salted Cashew - Farmley Standee Pouch 150 g": [3, 5],
        "Premium W320 Cashew Farmley Standee Pouch 500 g": [3, 5],
        "Premium Omani Fard Dates Farmley standee pouch 400 g": [3, 5],
        "Premium Omani Fard Dates Farmley standee pouch 500 g": [3, 5],
        "Barkat dates Farmley Standee Pouch 250 g": [3, 5],
        "Barkat Dates Farmley Standee Pouch 500g": [3, 5],
        "Kalmi Dates Farmley Standee Pouch 200 gms": [3, 5],
        "Trail Mix Farmley Standee Pouch 160 g": [3, 5],
        "Trail Mix Farmley Standee Pouch 200 g": [3, 5],
        "Berry Mix Farmley 160g Standee Pouch": [3, 5],
        "Berry Mix Farmley Standee Pouch 200 g": [3, 5],
        "Premium Anjeer Farmley Standee Pouch 250 g": [3, 5],
        "Premium Anjeer Farmley Standee Pouch 200 g": [3, 5],
        "Jumbo Iranian Roasted & Salted Pistachios Farmley Standee Pouch 200 g": [3, 5],
        "Premium California Roasted & Salted Pistachios Farmley Standee Pouch 35 g": [3, 5],
        "Premium California Roasted & Salted Pistachios Farmley Standee Pouch 200 g": [3, 5],
        "Premium Panchmeva Farmley Standee Pouch 160g": [3, 5],
        "Farmley Premium California Roasted & Salted Pistachios standee pouch 250 gm": [3, 5],
        "Premium Turkish Dried Apricot Farmley Standee Pouch 200 g": [3, 5],
        "Premium California Whole Dried Cranberries Farmley Standee Pouch 200 g": [3, 5],
        "Premium Omani Fard Dates Farmley Standee Pouch 35 g": [3, 5],
        "Smokey Almond Farmley 50g": [1, 3],
        "Salted Cashews Farmley 50g": [1, 3],
        "Premium Panchmeva Farmley Pillow Pouch 16g- Phool": [1, 3],
        "Panchmeva Farmley Pillow Pouch 18g": [1, 3],
        "Premium Panchmeva Farmley Pillow Pouch 21g": [1, 3],
        "Premium Panchmeva Farmley Pillow Pouch 30g": [1, 3],
        "Nut Mix Farmley 50g": [1, 3],
        "Sweet and Salty Mix Farmley Pillow Pouch 23g": [1, 3],
        "Party Mix Farmley Pillow Pouch 16g-Phool": [1, 3],
        "Mexican Peri Peri Snack Mix Farmley Pillow Pouch 21g": [1, 3],
        "Party Mix Farmley Pillow Pouch 35g": [1, 3],
        "Smokey BBQ Party Mix Pillow Pouch Farmley 28 g": [1, 3],
        "Fruit N Nut Mix Farmley Pillow Pouch 28g": [1, 3]
    };

    // --- Products where Material Uniformity (Mixing) is always NA ---
    const MATERIAL_NA_PRODUCTS = new Set([
        "Premium Anjeer Farmley Standee Pouch 200 g",
        "Watermelon Seeds Farmley Standee Pouch 500g",
        "Premium Raisin Long Farmley Standee Pouch 1 kg",
        "Premium California Roasted & Salted Pistachios Farmley Standee Pouch 35 g",
        "Premium Omani Fard Dates Farmley standee pouch 400 g",
        "Premium California Whole Dried Cranberries Farmley Standee Pouch 200 g",
        "Roasted & Salted Cashew - Farmley Standee Pouch 36 g - Retail",
        "Roasted & Salted Almonds Farmley Standee Pouch 40 g",
        "Premium Raisin Long Farmley Standee Pouch 200 g",
        "Date Powder Farmley Standee pouch 200 g",
        "Premium Turkish Dried Apricot Farmley Standee Pouch 200 g",
        "Premium Raisin Long Farmley Standee Pouch 500 g",
        "Premium Omani Fard Dates Farmley Standee Pouch 35 g",
        "Quinoa seeds Farmley Jar 1kg",
        "Premium Broken Walnut Kernels Farmley Standee Pouch 200 g",
        "Chia Seed Farmley Standee Pouch 500g",
        "Barkat Dates Farmley Standee Pouch 500g",
        "Farmley Seeds Combo - Chia Seeds, Flax Seeds, Pumpkin Seeds, Sunflower Seeds Farmley Standee Pouch 800 g (200 g Each)",
        "Premium Chia Seeds Farmley Standee Pouch 200 g",
        "Kalmi Dates Farmley Standee Pouch 200 gms",
        "Quinoa seeds Farmley Standee Pouch 500g",
        "Farmley Seed combo-Sunflower Seeds, Pumpkin Seeds Farmley Standee Pouch 400g (Pack of 2, 200g each)",
        "Barkat dates Farmley Standee Pouch 250 g",
        "Popular California Almonds Farmley Standee Pouch 250 g",
        "Farmley Dry Fruits Combo - Popular California Almonds, W400 Cashews (500 g Each)",
        "Popular W400 Cashew Farmley Generic Pouch 500 g",
        "Premium Munakka Farmley Standee Pouch 200 g",
        "Popular California Almonds Farmley Standee Pouch 500 g",
        "Barkat Dates Farmley Standee Pouch 1 kg -Pack of 2 (500 g each)",
        "Premium Chia Seeds Farmley Jar 1 Kg",
        "Pumpkin Seeds Farmley Standee Pouch 70 g",
        "Roasted Flax Seeds Farmley Standee Pouch 150g",
        "Premium Jumbo Pumpkin Seeds Farmley Standee Pouch 200 g",
        "Premium Extra Light Halves Walnut Kernels Farmley Standee Pouch 200 g",
        "Basil Seeds Farmley Standee pouch 300g",
        "Premium California Almonds Farmley Standee Pouch 500 g",
        "Premium California Almonds Farmley Standee Pouch 250 g",
        "Chia seeds Farmley Standee Pouch 300 g",
        "Premium California Almonds Farmley Standee Pouch 1 Kg",
        "Premium Flax Seeds Farmley Standee Pouch 200 g",
        "Popular Almond Farmley Standee Pouch 60g",
        "Premium W320 Cashew Farmley Vacuum Standee Pouch 250 g",
        "Popular California Almonds Farmley Standee Pouch 1 kg",
        "Premium Dried Mango Farmley Standee Pouch 200 g",
        "Date Melts Farmley Standee pouch 100g",
        "Premium Chia Seeds Farmley Standee pouch 400 g - Pack of 2 (200g each)",
        "Premium Omani Fard Dates Farmley standee pouch 500 g",
        "Premium Anjeer Farmley Standee Pouch 250 g",
        "Farmley Premium Almonds 250g, Premium W320 Cashews 250g, Premium Raisins 200g Large Potli 700g",
        "Premium Broken Walnut Kernels Farmley Standee Pouch 250 g",
        "Premium Chia Seeds Farmley Standee Pouch - Pack of 5 (200g each)",
        "Premium Omani Fard Dates Farmley Standee Pouch 800 g - Pack of 2 (400g each)",
        "Premium Figs Farmley Standee Pouch 400 g - Pack of 2 (200 g each)",
        "Chia Seeds Farmley Standee Pouch 120 g",
        "Premium Dried Cranberries Farmley Standee Pouch - Pack of 2 (200g each)",
        "Premium Dried Blueberries Farmley Standee Pouch 200 g",
        "Premium W320 Cashew Farmley Standee Pouch 500 g",
        "Cranberry Farmley Pillow Pouch 18g- Zepto",
        "Seed combo 800g",
        "Seed combo 400g",
        "Premium Dried Papaya Farmley Standee Pouch 200 g",
        "Premium Fard Dates Vendor Corrugated Box 10 Kg",
        "Prasadam Makhana Center Seal Pouch 100 g- NEW",
        "Premium chia seed jar 1kg",
        "Quinoa  seed jar 1kg",
        "Pumpkin seed 200g",
        "Chia  seed 200g",
        "Sunflower seed 200g",
        "Salted Pumpkin Seeds Pillow Pouch Farmley 24g-Ladi"
    ]);

    function normalizeData(data) {
        return data.map(row => {
            const normalizedRow = {};
            for (let key in row) {
                const normKey = key.trim().replace(/\s+/g, ' ');
                normalizedRow[normKey] = row[key];
            }
            const productName = String(normalizedRow['Product Name'] || '').trim();
            if (MATERIAL_NA_PRODUCTS.has(productName)) {
                normalizedRow['Material Uniformity (Mixing)'] = 'NA';
            }
            for (let key in normalizedRow) {
                if (key !== '_parsedDate') {
                    normalizedRow[key] = translateValue(normalizedRow[key]);
                }
            }
            if (!normalizedRow['Machine'] || normalizedRow['Machine'] === 'N/A') {
                normalizedRow['Machine'] = normalizedRow['Table No'] || 'N/A';
            }
            return normalizedRow;
        });
    }

    function translateValue(val) {
        if (val === null || val === undefined) return val;
        const s = String(val).trim().toLowerCase();
        if (s === 'okay' || s === 'ok') return 'Yes';
        if (s === 'not okay' || s === 'not ok') return 'No';
        if (s === 'n/a' || s === 'na') return 'N/A';
        return val;
    }

    const refreshIcon = document.getElementById('refreshIcon');
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');

    init();

    async function init() {
        const now = new Date();
        const todayStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
        if (startDateInput) startDateInput.value = todayStr;
        if (endDateInput) endDateInput.value = todayStr;
        if (datePreset) datePreset.value = 'today';

        showSkeleton();
        setupEventListeners();

        const cachedBody = localStorage.getItem('farmley_ipqc_data');
        if (cachedBody) {
            try {
                const parsed = JSON.parse(cachedBody);
                if (parsed && parsed.length > 0) {
                    allData = normalizeData(parsed);
                    allData.forEach(parseRowDate);
                    populateProductFilter(allData);
                    populateStaffFilter(allData);
                    populateShiftFilter(allData);
                    filteredData = filterData(allData);
                    processAndRender(filteredData);
                    if (lastUpdatedEl) lastUpdatedEl.textContent = 'Showing cached data — syncing live...';
                }
            } catch (e) { console.warn("Cache parse failed", e); }
        }
        fetchData();
    }

    function setupEventListeners() {
        if (applyFiltersBtn) {
            applyFiltersBtn.addEventListener('click', () => {
                if (allData.length === 0) return; 
                filteredData = filterData(allData);
                processAndRender(filteredData);
            });
        }
        [productFilter, nitrogenFilter, shiftFilter, checkedByFilter].forEach(filter => {
            if (filter) {
                filter.addEventListener('change', () => {
                    if (allData.length === 0) return;
                    filteredData = filterData(allData);
                    processAndRender(filteredData);
                });
            }
        });
        if (syncBtn) syncBtn.addEventListener('click', () => fetchData());
        if (downloadExcelBtn) downloadExcelBtn.addEventListener('click', () => exportToExcel());
        if (datePreset) {
            datePreset.addEventListener('change', () => {
                const val = datePreset.value;
                if (val === 'custom') return;
                const now = new Date();
                let start = new Date();
                let end = new Date();
                if (val === 'today') {} 
                else if (val === 'yesterday') { start.setDate(now.getDate() - 1); end.setDate(now.getDate() - 1); }
                else if (val === 'last7days') { start.setDate(now.getDate() - 6); }
                startDateInput.value = start.getFullYear() + '-' + String(start.getMonth() + 1).padStart(2, '0') + '-' + String(start.getDate()).padStart(2, '0');
                endDateInput.value = end.getFullYear() + '-' + String(end.getMonth() + 1).padStart(2, '0') + '-' + String(end.getDate()).padStart(2, '0');
                if (allData.length > 0) { filteredData = filterData(allData); processAndRender(filteredData); }
            });
        }
    }

    function parseRowDate(row) {
        if (row._parsedDate !== undefined) return;
        const raw = row['Date'] || row['Timestamp'] || row['date'] || row['timestamp'] || row['id'] || row['ID'];
        if (!raw) { row._parsedDate = null; return; }
        if (typeof raw === 'string') {
            const INformat = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
            if (INformat) { row._parsedDate = new Date(parseInt(INformat[3], 10), parseInt(INformat[2], 10) - 1, parseInt(INformat[1], 10)); return; }
        }
        const d = new Date(raw);
        if (!isNaN(d.getTime())) { row._parsedDate = new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
        else { row._parsedDate = null; }
    }

    async function fetchData() {
        try {
            if (refreshIcon) refreshIcon.classList.add('rotating');
            if (lastUpdatedEl && allData.length > 0) lastUpdatedEl.textContent = 'Syncing latest data...';

            // --- Tiered Fetching Strategy ---
            // Stage 1: Fast fetch for Today & Yesterday
            const now = new Date();
            const yesterday = new Date(now);
            yesterday.setDate(now.getDate() - 1);
            
            const formatDateParam = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
            const startDateParam = formatDateParam(yesterday);
            const endDateParam = formatDateParam(now);

            if (lastUpdatedEl) lastUpdatedEl.textContent = 'Fetching Today/Yesterday...';
            
            // Short timeout for Stage 1 (15s)
            const partialData = await fetchWithTimeout(`${APPS_SCRIPT_URL}?startDate=${startDateParam}&endDate=${endDateParam}&t=${Date.now()}`, 15000);
            
            if (partialData && partialData.length > 0) {
                updateAllData(partialData);
                if (lastUpdatedEl) lastUpdatedEl.textContent = 'Recent data loaded. Fetching history...';
            }

            // Stage 2: Background fetch for full history (last 30 days)
            // Longer timeout for Stage 2 (60s)
            const fullData = await fetchWithTimeout(`${APPS_SCRIPT_URL}?t=${Date.now()}`, 60000);
            
            if (fullData && fullData.length > 0) {
                localStorage.setItem('farmley_ipqc_data', JSON.stringify(fullData));
                updateAllData(fullData);
                if (lastUpdatedEl) lastUpdatedEl.textContent = `Last updated: ${new Date().toLocaleTimeString()}`;
            }

        } catch (error) {
            console.error('Fetch error:', error);
            if (lastUpdatedEl) lastUpdatedEl.textContent = 'Sync failed — showing last cached data';
            if (allData.length === 0) {
                allData = normalizeData(mockData || []);
                allData.forEach(parseRowDate);
                filteredData = filterData(allData);
                processAndRender(filteredData);
            }
        } finally {
            if (refreshIcon) setTimeout(() => refreshIcon.classList.remove('rotating'), 500);
        }
    }

    async function fetchWithTimeout(url, timeout) {
        const controller = new AbortController();
        const id = setTimeout(() => controller.abort(), timeout);
        try {
            const response = await fetch(url, { signal: controller.signal });
            return await response.json();
        } finally {
            clearTimeout(id);
        }
    }

    function updateAllData(newData) {
        const processed = normalizeData(newData);
        processed.forEach(parseRowDate);
        
        // Merge with existing allData or replace if it's a large set
        if (newData.length < 500) { 
            const map = new Map();
            // Use a combination of Date, Time and Product as a fallback unique key if ID is missing
            const getKey = (r) => r.ID || `${r.Date}-${r.Time}-${r['Product Name']}`;
            
            allData.forEach(r => map.set(getKey(r), r));
            processed.forEach(r => map.set(getKey(r), r));
            allData = Array.from(map.values());
        } else {
            allData = processed;
        }

        populateProductFilter(allData);
        populateStaffFilter(allData);
        populateShiftFilter(allData);
        filteredData = filterData(allData);
        
        if (filteredData.length === 0 && allData.length > 0) {
            filteredData = showLatestAvailableDate(allData);
        }

        requestAnimationFrame(() => {
            processAndRender(filteredData);
        });
    }

    function processAndRender(data) {
        const sorted = [...data].sort((a, b) => (b._parsedDate ? b._parsedDate.getTime() : 0) - (a._parsedDate ? a._parsedDate.getTime() : 0));
        renderTable(sorted, dashboardTableBody);
        renderCharts(sorted);
        updateStats(sorted);
    }

    function showSkeleton() {
        const statValues = ['totalRecords', 'lumpsCompliance', 'leakageCompliance', 'sealCompliance', 'materialCompliance', 'sizeCompliance', 'oxygenCompliance'];
        statValues.forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = '<div class="skeleton skeleton-value"></div>'; });
        if (dashboardTableBody) {
            let skelHtml = '';
            for (let i = 0; i < 6; i++) {
                skelHtml += '<tr class="table-skeleton-row">';
                for (let c = 0; c < 15; c++) skelHtml += '<td><div class="skeleton skeleton-cell"></div></td>';
                skelHtml += '</tr>';
            }
            dashboardTableBody.innerHTML = skelHtml;
        }
    }

    function getOxygenStatus(productName, value) {
        const val = parseFloat(value);
        if (isNaN(val)) return 'neutral';
        const range = productRanges[productName];
        if (!range) return 'compliant';
        return (val <= range[1]) ? 'compliant' : 'non-compliant';
    }

    function updateStats(data) {
        const total = data.length;
        if (total === 0) {
            ['totalRecords', 'lumpsCompliance', 'leakageCompliance', 'sealCompliance', 'materialCompliance', 'sizeCompliance', 'oxygenCompliance'].forEach(id => {
               const el = document.getElementById(id); if (el) el.textContent = '0%';
            });
            totalRecordsEl.textContent = '0';
            return;
        }
        const calcComp = (key, passVal) => {
            const relevant = data.filter(r => r[key] && String(r[key]).toLowerCase() !== 'n/a');
            if (relevant.length === 0) return { percent: 'N/A', count: 0, total: 0 };
            const passed = relevant.filter(r => String(r[key]).toLowerCase() === passVal.toLowerCase()).length;
            return { percent: ((passed / relevant.length) * 100).toFixed(1), count: passed, total: relevant.length };
        };
        const calcMaterial = () => {
            const relevant = data.filter(r => {
                const v = String(r['Material Uniformity (Mixing)'] || r['rial Uniformity (Mack & Seal)'] || '').toLowerCase().trim();
                return v === 'yes' || v === 'no' || v === 'ok' || v === 'not ok';
            });
            if (relevant.length === 0) return { percent: 'N/A', count: 0, total: 0 };
            const passed = relevant.filter(r => {
                const v = String(r['Material Uniformity (Mixing)'] || r['rial Uniformity (Mack & Seal)'] || '').toLowerCase().trim();
                return v === 'yes' || v === 'ok';
            }).length;
            return { percent: ((passed / relevant.length) * 100).toFixed(1), count: passed, total: relevant.length };
        };
        const lumps = calcComp('Lumps', 'No');
        const leakage = calcComp('Leakage Test', 'Yes');
        const seal = calcComp('Pack & Seal Integrity', 'Yes');
        const material = calcMaterial();
        const size = calcComp('Size Uniformity (Slice of Mixes)', 'Yes');
        const oxygenChecked = data.filter(r => !isNaN(parseFloat(r['Oxygen % Check'])));
        const compliantOxygen = oxygenChecked.filter(r => getOxygenStatus(r['Product Name'], r['Oxygen % Check']) === 'compliant').length;
        const oxygenPercent = oxygenChecked.length > 0 ? ((compliantOxygen / oxygenChecked.length) * 100).toFixed(1) : 'N/A';
        totalRecordsEl.textContent = total;
        const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val === 'N/A' ? 'N/A' : `${val}%`; };
        setVal('lumpsCompliance', lumps.percent); setVal('leakageCompliance', leakage.percent); setVal('sealCompliance', seal.percent); setVal('materialCompliance', material.percent); setVal('sizeCompliance', size.percent); setVal('oxygenCompliance', oxygenPercent);
    }

    function renderCharts(data) {
        if (complianceChart) complianceChart.destroy();
        const labels = ['Lumps', 'Leakage', 'Pack & Seal', 'Material Uniformity', 'Size Uniformity', 'Oxygen %'];
        const passData = [], failData = [];
        const getStatsBreakdown = (key, passVal) => {
            const rel = data.filter(r => r[key] && String(r[key]).toLowerCase() !== 'n/a');
            if (rel.length === 0) return { pass: 0, fail: 0 };
            const pass = rel.filter(r => String(r[key]).toLowerCase() === passVal.toLowerCase()).length;
            return { pass, fail: rel.length - pass };
        };
        const l = getStatsBreakdown('Lumps', 'No'); passData.push(l.pass); failData.push(l.fail);
        const lk = getStatsBreakdown('Leakage Test', 'Yes'); passData.push(lk.pass); failData.push(lk.fail);
        const s = getStatsBreakdown('Pack & Seal Integrity', 'Yes'); passData.push(s.pass); failData.push(s.fail);
        const mPass = data.filter(r => { const v = String(r['Material Uniformity (Mixing)'] || r['rial Uniformity (Mack & Seal)'] || '').toLowerCase(); return v === 'yes' || v === 'ok'; }).length;
        const mRelCount = data.filter(r => { const v = String(r['Material Uniformity (Mixing)'] || r['rial Uniformity (Mack & Seal)'] || '').toLowerCase(); return v === 'yes' || v === 'no' || v === 'ok' || v === 'not ok'; }).length;
        passData.push(mPass); failData.push(mRelCount - mPass);
        const sz = getStatsBreakdown('Size Uniformity (Slice of Mixes)', 'Yes'); passData.push(sz.pass); failData.push(sz.fail);
        const oxygenChecked = data.filter(r => !isNaN(parseFloat(r['Oxygen % Check'])));
        const oxyPass = oxygenChecked.filter(r => getOxygenStatus(r['Product Name'], r['Oxygen % Check']) === 'compliant').length;
        passData.push(oxyPass); failData.push(oxygenChecked.length - oxyPass);
        complianceChart = new Chart(document.getElementById('complianceChart'), {
            type: 'bar', data: { labels: labels, datasets: [{ label: 'Pass', data: passData, backgroundColor: '#10b981', borderRadius: 4 }, { label: 'Fail', data: failData, backgroundColor: '#ef4444', borderRadius: 4 }] },
            options: { responsive: true, maintainAspectRatio: false, scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } }, plugins: { legend: { position: 'top', align: 'end' } } }
        });
    }

    function filterData(data) {
        const product = productFilter.value;
        const nitrogen = nitrogenFilter ? nitrogenFilter.value : 'all';
        const shift = shiftFilter ? shiftFilter.value : 'all';
        const checkedBy = checkedByFilter ? checkedByFilter.value : 'all';
        const startVal = startDateInput.value, endVal = endDateInput.value;
        let startDate = startVal ? new Date(startVal.split('-')[0], startVal.split('-')[1]-1, startVal.split('-')[2]) : null;
        let endDate = endVal ? new Date(endVal.split('-')[0], endVal.split('-')[1]-1, endVal.split('-')[2]) : null;
        return data.filter(row => {
            if (product !== 'all' && row['Product Name'] !== product) return false;
            if (startDate || endDate) { const rd = row._parsedDate; if (!rd) return false; if (startDate && rd < startDate) return false; if (endDate && rd > endDate) return false; }
            if (nitrogen !== 'all' && String(row['Nitrogen Flush'] || '').toLowerCase() !== nitrogen.toLowerCase()) return false;
            if (shift !== 'all' && String(row['Shift'] || '').toLowerCase() !== shift.toLowerCase()) return false;
            if (checkedBy !== 'all' && row['Checked By'] !== checkedBy) return false;
            return true;
        });
    }

    function populateStaffFilter(data) {
        if (!checkedByFilter) return;
        const staff = [...new Set(data.map(item => item['Checked By']).filter(Boolean))].sort();
        const frag = document.createDocumentFragment();
        const allOpt = document.createElement('option'); allOpt.value = 'all'; allOpt.textContent = 'All Staff'; frag.appendChild(allOpt);
        staff.forEach(name => { const opt = document.createElement('option'); opt.value = name; opt.textContent = name; frag.appendChild(opt); });
        checkedByFilter.innerHTML = ''; checkedByFilter.appendChild(frag);
    }

    function populateShiftFilter(data) {
        if (!shiftFilter) return;
        const shifts = [...new Set(data.map(item => item['Shift']).filter(Boolean))].sort();
        const frag = document.createDocumentFragment();
        const allOpt = document.createElement('option'); allOpt.value = 'all'; allOpt.textContent = 'All Shifts'; frag.appendChild(allOpt);
        shifts.forEach(s => { const opt = document.createElement('option'); opt.value = s; opt.textContent = s; frag.appendChild(opt); });
        shiftFilter.innerHTML = ''; shiftFilter.appendChild(frag);
    }

    function checkRowFailure(row) {
        const qualityFields = [
            'Lumps', 'Leakage Test', 'Pack & Seal Integrity', 'Material Uniformity (Mixing)', 
            'Size Uniformity (Slice of Mixes)', 'Print Window', 'Burn Marks', 'Wrinkles in Seal',
            'RM Quality (Freshness/Aroma)', 'Seal Offset - Top', 'Seal Offset - Centre', 'Seal Offset - Bottom',
            'Moisture Condition', 'Aroma Type', 'Texture', 'Taste Profile 1', 'Taste Profile 2', 'Cooking Status'
        ];
        for (let field of qualityFields) if (String(row[field] || '').toLowerCase() === 'no') return true;
        if (getOxygenStatus(row['Product Name'], row['Oxygen % Check']) === 'non-compliant') return true;
        if (productRanges[row['Product Name']] && String(row['Nitrogen Flush'] || '').toLowerCase() === 'no') return true;
        return false;
    }

    function renderTable(data, targetBody) {
        if (!targetBody) return;
        if (data.length === 0) { targetBody.innerHTML = `<tr><td colspan="55" class="loading">No matching records found.</td></tr>`; return; }
        const rows = data.map(row => {
            const oxyStatus = getOxygenStatus(row['Product Name'], row['Oxygen % Check']);
            const failClass = checkRowFailure(row) ? ' class="row-failure"' : '';
            const renderCell = (val) => {
                const s = String(val || 'N/A');
                const highlight = (s === 'No') ? ' class="breach-highlight"' : '';
                return `<td${highlight}>${s}</td>`;
            };
            return `<tr${failClass}>
                <td>${row['ID'] || 'N/A'}</td>
                <td>${row['Batch Code'] || 'N/A'}</td>
                <td>${row['Type'] || 'N/A'}</td>
                <td>${formatDate(row['Date'] || '')}</td>
                <td>${row['Time'] || 'N/A'}</td>
                <td>${row['Product Name'] || 'N/A'}</td>
                <td>${row['MRP'] || 'N/A'}</td>
                <td>${row['USP'] || 'N/A'}</td>
                <td>${row['Shift'] || 'N/A'}</td>
                <td>${row['Machine'] || 'N/A'}</td>
                <td>${row['Table No'] || 'N/A'}</td>
                <td>${row['Net Weight'] || 'N/A'}</td>
                <td>${row['Use By Date'] || 'N/A'}</td>
                <td>${row['Serving Per Pack'] || 'N/A'}</td>
                <td>${row['Production Tables'] || 'N/A'}</td>
                <td>${row['Manufacturing Date'] || 'N/A'}</td>
                <td>${row['Remarks'] || 'N/A'}</td>
                <td>${row['MOD'] || 'N/A'}</td>
                <td>${row['Checked By'] || 'N/A'}</td>
                <td>${row['Verified By'] || 'N/A'}</td>
                <td>${row['Average Weight'] || 'N/A'}</td>
                <td>${row['Empty Pouch Weight'] || 'N/A'}</td>
                <td>${row['Weights'] || 'N/A'}</td>
                <td>${row['Nitrogen Flush'] || 'N/A'}</td>
                ${renderCell(row['Leakage Test'])}
                <td>${row['Mono Carton'] || 'N/A'}</td>
                ${renderCell(row['Material Uniformity (Mixing)'])}
                ${renderCell(row['Pack & Seal Integrity'])}
                ${renderCell(row['Size Uniformity (Slice of Mixes)'])}
                ${renderCell(row['Print Window'])}
                ${renderCell(row['Lumps'])}
                <td>${row['EAN Scan Check'] || 'N/A'}</td>
                <td>${row['Damaged/Broken'] || 'N/A'}</td>
                <td>${row['Carton Label & Size'] || 'N/A'}</td>
                ${renderCell(row['Burn Marks'])}
                ${renderCell(row['Wrinkles in Seal'])}
                ${renderCell(row['RM Quality (Freshness/Aroma)'])}
                ${renderCell(row['Seal Offset - Top'])}
                ${renderCell(row['Seal Offset - Centre'])}
                ${renderCell(row['Seal Offset - Bottom'])}
                <td>${row['RTV Use & Mixing Qty/Ratio'] || 'N/A'}</td>
                <td>${row['Mixing Ratio'] || 'N/A'}</td>
                <td>${row['RTV Remarks'] || 'N/A'}</td>
                <td class="${oxyStatus === 'non-compliant' ? 'breach-highlight' : ''}">${row['Oxygen % Check'] || 'N/A'}${row['Oxygen % Check'] && row['Oxygen % Check'] !== 'N/A' ? '%' : ''}</td>
                <td>${row['Pouch Height(mm)'] || 'N/A'}</td>
                <td>${row['Carton Weight'] || 'N/A'}</td>
                ${renderCell(row['Crunch/Sogginess'])}
                <td>${row['Empty Carton Weight'] || 'N/A'}</td>
                <td>${row['RM Quality Remarks'] || 'N/A'}</td>
                ${renderCell(row['Moisture Condition'])}
                <td>${row['Aroma Type'] || 'N/A'}</td>
                <td>${row['Texture'] || 'N/A'}</td>
                <td>${row['Taste Profile 1'] || 'N/A'}</td>
                <td>${row['Taste Profile 2'] || 'N/A'}</td>
                <td>${row['Cooking Status'] || 'N/A'}</td>
            </tr>`;
        });
        targetBody.innerHTML = rows.join('');
    }

    function exportToExcel() {
        if (filteredData.length === 0) { alert('No data to export'); return; }
        const headers = [
            'ID', 'Batch Code', 'Type', 'Date', 'Time', 'Product Name', 'MRP', 'USP', 'Shift', 'Machine', 'Table No', 'Net Weight', 'Use By Date', 'Serving Per Pack', 'Production Tables', 'Manufacturing Date', 'Remarks', 'MOD', 'Checked By', 'Verified By', 'Average Weight', 'Empty Pouch Weight', 'Weights', 'Nitrogen Flush', 'Leakage Test', 'Mono Carton', 'Material Uniformity (Mixing)', 'Pack & Seal Integrity', 'Size Uniformity (Slice of Mixes)', 'Print Window', 'Lumps', 'EAN Scan Check', 'Damaged/Broken', 'Carton Label & Size', 'Burn Marks', 'Wrinkles in Seal', 'RM Quality (Freshness/Aroma)', 'Seal Offset - Top', 'Seal Offset - Centre', 'Seal Offset - Bottom', 'RTV Use & Mixing Qty/Ratio', 'Mixing Ratio', 'RTV Remarks', 'Oxygen % Check', 'Pouch Height(mm)', 'Carton Weight', 'Crunch/Sogginess', 'Empty Carton Weight', 'RM Quality Remarks', 'Moisture Condition', 'Aroma Type', 'Texture', 'Taste Profile 1', 'Taste Profile 2', 'Cooking Status'
        ];
        let csvContent = "data:text/csv;charset=utf-8," + headers.join(",") + "\n";
        filteredData.forEach(row => {
            const rowData = headers.map(header => {
                let cell = row[header] || 'N/A';
                if (typeof cell === 'string' && cell.includes(',')) cell = `"${cell}"`;
                return cell;
            });
            csvContent += rowData.join(",") + "\n";
        });
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `IPQC_Indore_Export_${new Date().toISOString().split('T')[0]}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function populateProductFilter(data) {
        if (!productFilter) return;
        const products = [...new Set(data.map(r => r['Product Name']))].sort();
        const frag = document.createDocumentFragment();
        const allOpt = document.createElement('option'); allOpt.value = 'all'; allOpt.textContent = 'All Products'; frag.appendChild(allOpt);
        products.forEach(p => { const opt = document.createElement('option'); opt.value = p; opt.textContent = p; frag.appendChild(opt); });
        productFilter.innerHTML = ''; productFilter.appendChild(frag);
    }

    function formatDate(dateStr) {
        if (!dateStr) return 'N/A';
        let date = new Date(dateStr);
        if (typeof dateStr === 'string') {
            const INformat = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
            if (INformat) date = new Date(parseInt(INformat[3], 10), parseInt(INformat[2], 10) - 1, parseInt(INformat[1], 10));
        }
        if (isNaN(date.getTime())) return dateStr;
        return date.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
    }
});
