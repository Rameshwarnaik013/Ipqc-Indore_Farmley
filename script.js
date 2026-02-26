document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    // const tableBody = document.getElementById('tableBody');
    const syncBtn = document.getElementById('syncData');
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
    const statusFilter = document.getElementById('statusFilter');
    const datePreset = document.getElementById('datePreset');
    const shiftFilter = document.getElementById('shiftFilter');
    const checkedByFilter = document.getElementById('checkedByFilter');

    let allData = [];
    let complianceChart;

    // --- Configuration ---
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyL9c9pk5Z8gNCIcWXurkGb33zitvQmhivi9eK-M-kzNye6gZi_UESv2ADQe2oShkDb/exec';

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

    // Normalise: force Material Uniformity to 'NA' for products in the NA list
    function normalizeData(data) {
        return data.map(row => {
            const product = String(row['Product Name'] || '').trim();
            if (MATERIAL_NA_PRODUCTS.has(product)) {
                return { ...row, 'Material Uniformity (Mixing)': 'NA' };
            }
            return row;
        });
    }

    // --- Mock Data ---
    const mockData = [
        { "Date": "2026-02-04", "Product Name": "Date Powder Farmley Standee pouch 200 g", "Lumps": "No", "Leakage Test": "Yes", "Pack & Seal Integrity": "Yes", "Material Uniformity (Mixing)": "Yes", "Size Uniformity (Slice of Mixes)": "Yes", "Nitrogen Flush": "Yes", "Oxygen % Check": "2.5", "Shift": "Day", "Checked By": "Deepak" },
        { "Date": "2026-02-04", "Product Name": "Popular California Almonds Farmley Standee Pouch 1 kg", "Lumps": "No", "Leakage Test": "Yes", "Pack & Seal Integrity": "Yes", "Material Uniformity (Mixing)": "Yes", "Size Uniformity (Slice of Mixes)": "Yes", "Nitrogen Flush": "Yes", "Oxygen % Check": "9.2", "Shift": "Day", "Checked By": "Saloni" },
        { "Date": "2026-02-04", "Product Name": "Smokey Almond Farmley 50g", "Lumps": "Yes", "Leakage Test": "No", "Pack & Seal Integrity": "No", "Material Uniformity (Mixing)": "No", "Size Uniformity (Slice of Mixes)": "No", "Nitrogen Flush": "No", "Oxygen % Check": "4.5", "Shift": "Night", "Checked By": "John" }
    ];

    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');

    init();

    async function init() {
        // Set default dates to Today
        const todayStr = new Date().toISOString().split('T')[0];
        if (startDateInput) startDateInput.value = todayStr;
        if (endDateInput) endDateInput.value = todayStr;

        await fetchData();
        populateProductFilter(allData);
        populateStaffFilter(allData);
        setupEventListeners();

        // Initial filter for Today's data
        filteredData = filterData(allData);
        processAndRender(filteredData);
    }

    function setupEventListeners() {

        if (applyFiltersBtn) {
            applyFiltersBtn.addEventListener('click', () => {
                filteredData = filterData(allData);
                processAndRender(filteredData);
            });
        }

        // Auto-apply for all filters
        [productFilter, statusFilter, shiftFilter, checkedByFilter].forEach(filter => {
            if (filter) {
                filter.addEventListener('change', () => {
                    filteredData = filterData(allData);
                    processAndRender(filteredData);
                });
            }
        });

        if (syncBtn) {
            syncBtn.addEventListener('click', () => {
                fetchData();
            });
        }


        if (datePreset) {
            datePreset.addEventListener('change', () => {
                const val = datePreset.value;
                if (val === 'custom') return;

                const now = new Date();
                let start = new Date();
                let end = new Date();

                if (val === 'today') {
                    // today already set
                } else if (val === 'yesterday') {
                    start.setDate(now.getDate() - 1);
                    end.setDate(now.getDate() - 1);
                } else if (val === 'last7days') {
                    start.setDate(now.getDate() - 6);
                }

                startDateInput.value = start.toISOString().split('T')[0];
                endDateInput.value = end.toISOString().split('T')[0];

                // Auto apply
                filteredData = filterData(allData);
                processAndRender(filteredData);
            });
        }

        [startDateInput, endDateInput].forEach(inp => {
            if (inp) {
                inp.addEventListener('input', () => {
                    if (datePreset) datePreset.value = 'custom';
                });
            }
        });
    }

    async function fetchData() {
        try {
            renderTable([], dashboardTableBody, true); // Loading state

            const response = await fetch(`${APPS_SCRIPT_URL}?t=${Date.now()}`);
            const data = await response.json();

            console.log(`Fetched ${data.length} records from GSheet`);

            if (data.error) {
                console.error('Server error:', data.error);
                dashboardTableBody.innerHTML = `<tr><td colspan="11" class="loading" style="color:red">Server Error: ${data.error}</td></tr>`;
                return;
            }

            allData = normalizeData(data);
            populateStaffFilter(allData);
            populateShiftFilter(allData);
            filteredData = [...allData];
            processAndRender(allData);
            lastUpdatedEl.textContent = `Last updated: ${new Date().toLocaleTimeString()}`;
        } catch (error) {
            console.error('Error fetching data:', error);
            allData = normalizeData(mockData);
            processAndRender(allData);
        }
    }

    function processAndRender(data) {
        renderTable(data, dashboardTableBody);
        renderCharts(data);
        updateStats(data);
    }

    function getOxygenStatus(productName, value) {
        const range = productRanges[productName];
        if (!range) return 'neutral';
        const val = parseFloat(value);
        if (isNaN(val)) return 'neutral';
        return (val >= range[0] && val <= range[1]) ? 'compliant' : 'non-compliant';
    }


    function updateStats(data) {
        const total = data.length;
        if (total === 0) {
            const zeros = ['totalRecords', 'lumpsCompliance', 'leakageCompliance', 'sealCompliance', 'materialCompliance', 'sizeCompliance', 'oxygenCompliance'];
            zeros.forEach(id => document.getElementById(id).textContent = '0%');
            totalRecordsEl.textContent = '0';
            return;
        }

        // Helper to calculate compliance % excluding N/A
        const calcComp = (key, passVal) => {
            const relevant = data.filter(r => r[key] && String(r[key]).toLowerCase() !== 'n/a');
            if (relevant.length === 0) return { percent: 'N/A', count: 0, total: 0 };
            const passed = relevant.filter(r => String(r[key]).toLowerCase() === passVal.toLowerCase()).length;
            return {
                percent: ((passed / relevant.length) * 100).toFixed(1),
                count: passed,
                total: relevant.length
            };
        };

        // Material Uniformity: treat as NA by default — only count explicit Yes/No values
        const calcMaterial = () => {
            const relevant = data.filter(r => {
                const v = String(r['Material Uniformity (Mixing)'] || '').toLowerCase().trim();
                return v === 'yes' || v === 'no';
            });
            if (relevant.length === 0) return { percent: 'N/A', count: 0, total: 0 };
            const passed = relevant.filter(r => String(r['Material Uniformity (Mixing)']).toLowerCase().trim() === 'yes').length;
            return {
                percent: ((passed / relevant.length) * 100).toFixed(1),
                count: passed,
                total: relevant.length
            };
        };

        const lumps = calcComp('Lumps', 'No');
        const leakage = calcComp('Leakage Test', 'Yes');
        const seal = calcComp('Pack & Seal Integrity', 'Yes');
        const material = calcMaterial();
        const size = calcComp('Size Uniformity (Slice of Mixes)', 'Yes');

        const oxygenChecked = data.filter(r => productRanges[r['Product Name']] && String(r['Oxygen % Check']).toLowerCase() !== 'n/a');
        const compliantOxygen = oxygenChecked.filter(r => getOxygenStatus(r['Product Name'], r['Oxygen % Check']) === 'compliant').length;
        const oxygenPercent = oxygenChecked.length > 0 ? ((compliantOxygen / oxygenChecked.length) * 100).toFixed(1) : 'N/A';

        const updateCardUI = (cardId, percent) => {
            const cardEl = document.getElementById(cardId);
            if (!cardEl) return;
            const card = cardEl.closest('.stat-card');
            if (!card) return;
            card.classList.remove('high-compliance', 'mid-compliance', 'low-compliance');
            if (percent === 'N/A') return;
            const p = parseFloat(percent);
            if (p >= 95) card.classList.add('high-compliance');
            else if (p >= 80) card.classList.add('mid-compliance');
            else card.classList.add('low-compliance');
        };

        totalRecordsEl.textContent = total;

        const setVal = (el, val) => el.textContent = val === 'N/A' ? 'N/A' : `${val}%`;
        setVal(lumpsComplianceEl, lumps.percent);
        setVal(leakageComplianceEl, leakage.percent);
        setVal(sealComplianceEl, seal.percent);
        setVal(materialComplianceEl, material.percent);
        setVal(sizeComplianceEl, size.percent);
        setVal(oxygenComplianceEl, oxygenPercent);

        lumpsCountEl.textContent = `${lumps.count} / ${lumps.total} compliant`;
        leakageCountEl.textContent = `${leakage.count} / ${leakage.total} compliant`;
        sealCountEl.textContent = `${seal.count} / ${seal.total} compliant`;
        materialCountEl.textContent = `${material.count} / ${material.total} compliant`;
        sizeCountEl.textContent = `${size.count} / ${size.total} compliant`;
        oxygenCountEl.textContent = oxygenPercent === 'N/A' ? 'N/A' : `${compliantOxygen} / ${oxygenChecked.length} compliant`;

        updateCardUI('lumpsCompliance', lumps.percent);
        updateCardUI('leakageCompliance', leakage.percent);
        updateCardUI('sealCompliance', seal.percent);
        updateCardUI('materialCompliance', material.percent);
        updateCardUI('sizeCompliance', size.percent);
        updateCardUI('oxygenCompliance', oxygenPercent);
    }

    function renderCharts(data) {
        if (complianceChart) complianceChart.destroy();

        const labels = ['Lumps', 'Leakage', 'Pack & Seal', 'Material Uniformity (Mixing)', 'Size Uniformity (Slice of Mixes)', 'Oxygen %'];
        const passData = [];
        const failData = [];

        const getStatsBreakdown = (key, passVal) => {
            const rel = data.filter(r => r[key] && String(r[key]).toLowerCase() !== 'n/a');
            if (rel.length === 0) return { pass: 0, fail: 0 };
            const pass = rel.filter(r => String(r[key]).toLowerCase() === passVal.toLowerCase()).length;
            return { pass, fail: rel.length - pass };
        };

        // Lumps (No = Pass)
        const l = getStatsBreakdown('Lumps', 'No'); passData.push(l.pass); failData.push(l.fail);
        // Leakage (Yes = Pass)
        const lk = getStatsBreakdown('Leakage Test', 'Yes'); passData.push(lk.pass); failData.push(lk.fail);
        // Seal (Yes = Pass)
        const s = getStatsBreakdown('Pack & Seal Integrity', 'Yes'); passData.push(s.pass); failData.push(s.fail);
        // Material Uniformity: only count explicit Yes/No (NA by default)
        const mRel = data.filter(r => { const v = String(r['Material Uniformity (Mixing)'] || '').toLowerCase().trim(); return v === 'yes' || v === 'no'; });
        const mPass = mRel.filter(r => String(r['Material Uniformity (Mixing)']).toLowerCase().trim() === 'yes').length;
        passData.push(mPass); failData.push(mRel.length - mPass);
        // Size (Yes = Pass)
        const sz = getStatsBreakdown('Size Uniformity (Slice of Mixes)', 'Yes'); passData.push(sz.pass); failData.push(sz.fail);
        // Oxygen
        const oxygenChecked = data.filter(r => productRanges[r['Product Name']] && String(r['Oxygen % Check']).toLowerCase() !== 'n/a');
        const oxyPass = oxygenChecked.filter(r => getOxygenStatus(r['Product Name'], r['Oxygen % Check']) === 'compliant').length;
        passData.push(oxyPass); failData.push(oxygenChecked.length - oxyPass);

        complianceChart = new Chart(document.getElementById('complianceChart'), {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Pass',
                        data: passData,
                        backgroundColor: '#10b981',
                        borderRadius: 6,
                        categoryPercentage: 0.6,
                        barPercentage: 0.8
                    },
                    {
                        label: 'Fail',
                        data: failData,
                        backgroundColor: '#ef4444',
                        borderRadius: 6,
                        categoryPercentage: 0.6,
                        barPercentage: 0.8
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'x',
                scales: {
                    x: {
                        stacked: true,
                        grid: { display: false }
                    },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        grid: { color: '#f3f4f6' },
                        title: { display: true, text: 'Count of Checks' }
                    }
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        align: 'end',
                        labels: {
                            usePointStyle: true,
                            boxWidth: 8
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(31, 41, 55, 0.9)',
                        padding: 12,
                        titleFont: { size: 14 },
                        bodyFont: { size: 13 },
                        callbacks: {
                            label: function (context) {
                                let label = context.dataset.label || '';
                                let value = context.parsed.y;
                                let idx = context.dataIndex;
                                let total = passData[idx] + failData[idx];
                                let percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                                return ` ${label}: ${value} (${percentage}%)`;
                            }
                        }
                    }
                }
            }
        });
    }

    function filterData(data) {
        const product = productFilter.value;
        const status = statusFilter.value;
        const shift = shiftFilter ? shiftFilter.value : 'all';
        const checkedBy = checkedByFilter ? checkedByFilter.value : 'all';
        const startVal = startDateInput.value;
        const endVal = endDateInput.value;

        const startDate = startVal ? new Date(startVal) : null;
        const endDate = endVal ? new Date(endVal) : null;

        if (startDate) startDate.setHours(0, 0, 0, 0);
        if (endDate) endDate.setHours(23, 59, 59, 999);

        return data.filter(row => {
            const productMatch = product === 'all' || row['Product Name'] === product;

            const rowDateRaw = row['Date'] || row['Timestamp'] || row['date'] || row['timestamp'];
            const rowDate = rowDateRaw ? new Date(rowDateRaw) : null;

            let dateMatch = true;
            if (rowDate && !isNaN(rowDate.getTime())) {
                rowDate.setHours(0, 0, 0, 0);
                if (startDate && rowDate < startDate) dateMatch = false;
                if (endDate && rowDate > endDate) dateMatch = false;
            } else {
                // If no valid date found, only show if no date filter is active or if we can't determine
                if (startDate || endDate) dateMatch = false;
            }

            // Compliance status filter
            let statusMatch = true;
            if (status !== 'all') {
                const hasFailure = checkRowFailure(row);
                statusMatch = (status === 'fail') ? hasFailure : !hasFailure;
            }

            const shiftMatch = shift === 'all' || String(row['Shift'] || '').toLowerCase() === shift.toLowerCase();
            const staffMatch = checkedBy === 'all' || row['Checked By'] === checkedBy;

            return productMatch && dateMatch && statusMatch && shiftMatch && staffMatch;
        });
    }

    function populateStaffFilter(data) {
        if (!checkedByFilter) return;
        const staff = [...new Set(data.map(item => item['Checked By']).filter(Boolean))].sort();
        const currentVal = checkedByFilter.value;
        checkedByFilter.innerHTML = '<option value="all">All Staff</option>';
        staff.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            checkedByFilter.appendChild(option);
        });
        if (staff.includes(currentVal)) checkedByFilter.value = currentVal;
    }

    function populateShiftFilter(data) {
        if (!shiftFilter) return;
        const shifts = [...new Set(data.map(item => item['Shift']).filter(Boolean))].sort();
        const currentVal = shiftFilter.value;
        shiftFilter.innerHTML = '<option value="all">All Shifts</option>';
        shifts.forEach(s => {
            const option = document.createElement('option');
            option.value = s;
            option.textContent = s;
            shiftFilter.appendChild(option);
        });
        if (shifts.includes(currentVal)) shiftFilter.value = currentVal;
    }

    function checkRowFailure(row) {
        const lpFail = String(row['Lumps'] || '').toLowerCase() === 'yes';
        const lkFail = String(row['Leakage Test'] || '').toLowerCase() === 'no';
        const sFail = String(row['Pack & Seal Integrity'] || '').toLowerCase() === 'no';
        const mFail = String(row['Material Uniformity (Mixing)'] || '').toLowerCase() === 'no';
        const szFail = String(row['Size Uniformity (Slice of Mixes)'] || '').toLowerCase() === 'no';

        const oxyStatus = getOxygenStatus(row['Product Name'], row['Oxygen % Check']);
        const oFail = oxyStatus === 'non-compliant';

        const isGasProd = productRanges[row['Product Name']];
        const nFail = (isGasProd && String(row['Nitrogen Flush'] || '').toLowerCase() === 'no');

        return lpFail || lkFail || sFail || mFail || szFail || oFail || nFail;
    }

    function renderTable(data, targetBody, isLoading = false) {
        if (!targetBody) return;
        targetBody.innerHTML = '';

        if (isLoading) {
            targetBody.innerHTML = `<tr><td colspan="11" class="loading">Loading data...</td></tr>`;
            return;
        }

        if (data.length === 0) {
            targetBody.innerHTML = `<tr><td colspan="11" class="loading">No matching records found.</td></tr>`;
            return;
        }

        data.forEach(row => {
            const tr = document.createElement('tr');

            const lumpsS = String(row['Lumps'] || '').toLowerCase() === 'yes' ? 'status-yes' : 'status-no';
            const leakageS = String(row['Leakage Test'] || '').toLowerCase() === 'no' ? 'status-yes' : 'status-no';
            const sealS = String(row['Pack & Seal Integrity'] || '').toLowerCase() === 'no' ? 'status-yes' : 'status-no';
            const materialS = String(row['Material Uniformity (Mixing)'] || '').toLowerCase() === 'no' ? 'status-yes' : 'status-no';
            const sizeS = String(row['Size Uniformity (Slice of Mixes)'] || '').toLowerCase() === 'no' ? 'status-yes' : 'status-no';

            const oxyStatus = getOxygenStatus(row['Product Name'], row['Oxygen % Check']);
            const oxyClass = oxyStatus === 'non-compliant' ? 'breach-cell' : (oxyStatus === 'compliant' ? 'status-no' : '');

            const isGasProd = productRanges[row['Product Name']];
            const nWarning = (isGasProd && String(row['Nitrogen Flush'] || '').toLowerCase() === 'no') ? 'status-yes' : '';

            if (checkRowFailure(row)) tr.classList.add('row-failure');

            const dateVal = row['Date'] || row['Timestamp'] || row['date'] || row['timestamp'];
            const oxygenStatus = getOxygenStatus(row['Product Name'], row['Oxygen % Check']);

            tr.innerHTML = `
                <td>${formatDate(dateVal)}</td>
                <td title="${row['Product Name']}">${row['Product Name']}</td>
                <td class="${String(row['Lumps'] || '').toLowerCase() === 'yes' ? 'breach-highlight' : ''}">${row['Lumps'] || 'N/A'}</td>
                <td class="${String(row['Leakage Test'] || '').toLowerCase() === 'no' ? 'breach-highlight' : ''}">${row['Leakage Test'] || 'N/A'}</td>
                <td class="${String(row['Pack & Seal Integrity'] || '').toLowerCase() === 'no' ? 'breach-highlight' : ''}">${row['Pack & Seal Integrity'] || 'N/A'}</td>
                <td class="${String(row['Material Uniformity (Mixing)'] || '').toLowerCase() === 'no' ? 'breach-highlight' : ''}">${row['Material Uniformity (Mixing)'] || 'N/A'}</td>
                <td class="${String(row['Size Uniformity (Slice of Mixes)'] || '').toLowerCase() === 'no' ? 'breach-highlight' : ''}">${row['Size Uniformity (Slice of Mixes)'] || 'N/A'}</td>
                <td>${row['Nitrogen % Check'] || 'N/A'}</td>
                <td class="${oxygenStatus === 'breach' || oxygenStatus === 'non-compliant' ? 'breach-highlight' : ''}">${row['Oxygen % Check'] || 'N/A'}${oxyStatus !== 'neutral' ? '%' : ''}</td>
                <td>${row['Shift'] || 'N/A'}</td>
                <td title="${row['Checked By']}">${row['Checked By'] || 'N/A'}</td>
            `;
            targetBody.appendChild(tr);
        });
    }
    function populateProductFilter(data) {
        if (!productFilter) return;
        const currentVal = productFilter.value;
        productFilter.innerHTML = '<option value="all">All Products</option>';
        const products = [...new Set(data.map(r => r['Product Name']))].sort();
        products.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p;
            opt.textContent = p;
            if (p === currentVal) opt.selected = true;
            productFilter.appendChild(opt);
        });
    }

    function formatDate(dateStr) {
        if (!dateStr) return 'N/A';
        const date = new Date(dateStr);
        if (isNaN(date.getTime())) return dateStr;

        // Return DD/MM/YYYY or similar human readable format
        return date.toLocaleDateString('en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
    }
});
