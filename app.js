let ehrFiles = [];
let irisFiles = [];
let combinedFiles = [];

let ehrData = [];
let irisData = [];

// Store datatable instances to destroy them before re-rendering
const dtInstances = {};

// LocalStorage Cache for Working List
const CACHE_KEY = 'vaccine_script_working_list_cache';

function getResolvedCache() {
    const cache = localStorage.getItem(CACHE_KEY);
    return cache ? JSON.parse(cache) : {};
}

function setResolvedStatus(key, isResolved) {
    const cache = getResolvedCache();
    if (isResolved) {
        cache[key] = true;
    } else {
        delete cache[key];
    }
    localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
}

function clearResolvedCache() {
    localStorage.removeItem(CACHE_KEY);
}

function initDataTable(tableId) {
    if (dtInstances[tableId]) {
        dtInstances[tableId].destroy();
    }
    dtInstances[tableId] = new simpleDatatables.DataTable(`#${tableId}`, {
        searchable: true,
        fixedHeight: false,
        perPage: 10,
        perPageSelect: [10, 25, 50, 100],
    });
}

document.addEventListener('DOMContentLoaded', () => {
    setupDropzone('ehr-dropzone', 'ehr-upload', 'ehr-file-list', ehrFiles, () => checkReady());
    setupDropzone('iris-dropzone', 'iris-upload', 'iris-file-list', irisFiles, () => checkReady());
    setupDropzone('combined-dropzone', 'combined-upload', 'combined-file-list', combinedFiles, () => checkReady());

    document.getElementById('btn-export').addEventListener('click', () => {
        exportExcel();
    });
    document.getElementById('btn-query').addEventListener('click', runQuery);

    // Global delegated listener for resolved checkboxes (survives DataTables DOM rewrites)
    document.addEventListener('change', (e) => {
        if (e.target && e.target.classList.contains('resolved-checkbox')) {
            const key = e.target.getAttribute('data-key');
            setResolvedStatus(key, e.target.checked);
            
            // Visual feedback
            if (e.target.checked) {
                e.target.closest('tr').style.opacity = '0.5';
            } else {
                e.target.closest('tr').style.opacity = '1';
            }
        }
    });

    // Clear Progress Button
    document.getElementById('btn-clear-progress').addEventListener('click', () => {
        if (confirm('Are you sure you want to clear your progress? This will uncheck all boxes in your Working List and cannot be undone.')) {
            clearResolvedCache();
            // Re-render the working list to uncheck boxes if data exists
            if (window.currentWorkingListData) {
                renderWorkingList('table-working-list', window.currentWorkingListData);
            }
        }
    });

    // Tab switching
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            e.target.classList.add('active');
            document.getElementById(e.target.getAttribute('data-target')).classList.add('active');
        });
    });
});

function setupDropzone(zoneId, inputId, listId, fileArray, onUpdate) {
    const dropzone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    const list = document.getElementById(listId);

    // Removed dropzone click listener because the input naturally covers the area

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files, fileArray, list, onUpdate);
    });

    input.addEventListener('change', (e) => {
        handleFiles(e.target.files, fileArray, list, onUpdate);
        input.value = ''; // Reset
    });
}

function handleFiles(files, fileArray, listElement, onUpdate) {
    if (!files) return;
    Array.from(files).forEach(file => {
        const name = file.name.toLowerCase();
        if (name.endsWith('.csv') || name.endsWith('.xlsx') || name.endsWith('.xls')) {
            fileArray.push(file);
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `<span>${file.name}</span> <span style="cursor:pointer" onclick="event.stopPropagation(); this.parentElement.remove(); window.removeFile('${file.name}', '${fileArray === ehrFiles ? 'ehr' : (fileArray === irisFiles ? 'iris' : 'combined')}')">❌</span>`;
            listElement.appendChild(item);
        }
    });
    onUpdate();
}

window.removeFile = function(fileName, type) {
    if (type === 'ehr') {
        ehrFiles = ehrFiles.filter(f => f.name !== fileName);
    } else if (type === 'iris') {
        irisFiles = irisFiles.filter(f => f.name !== fileName);
    } else {
        combinedFiles = combinedFiles.filter(f => f.name !== fileName);
    }
    checkReady();
}

function checkReady() {
    const btnExport = document.getElementById('btn-export');
    const btnQuery = document.getElementById('btn-query');
    const status = document.getElementById('data-status');

    if ((ehrFiles.length > 0 && irisFiles.length > 0) || combinedFiles.length > 0) {
        btnExport.disabled = combinedFiles.length > 0; // Disable export if using a combined file
        btnQuery.disabled = false;
        if (combinedFiles.length > 0) {
            status.innerHTML = `✅ Ready (Combined Excel Loaded)`;
        } else {
            status.innerHTML = `✅ Ready (${ehrFiles.length} EHR, ${irisFiles.length} IRIS files)`;
        }
        status.style.color = 'var(--success-color)';
        
        // Auto-load data to show preview tables
        if (!window.isParsing) {
            window.isParsing = true;
            loadAllData().then(() => {
                window.isParsing = false;
            });
        }
    } else {
        btnExport.disabled = true;
        btnQuery.disabled = true;
        
        if (ehrFiles.length > 0 || irisFiles.length > 0) {
            status.innerHTML = `Waiting for both types... (${ehrFiles.length} EHR, ${irisFiles.length} IRIS)`;
            status.style.color = 'var(--warning-color)';
        } else {
            status.innerHTML = 'Awaiting files...';
            status.style.color = 'inherit';
        }
    }
}

async function parseCSV(file) {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: false, // Read as 2D array first to find the true header
            skipEmptyLines: true,
            complete: (results) => {
                resolve(extractDataFrom2DArray(results.data));
            },
            error: (err) => reject(err)
        });
    });
}

async function parseExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                resolve(extractDataFrom2DArray(rows));
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsArrayBuffer(file);
    });
}

function extractDataFrom2DArray(rows) {
    if (rows.length < 2) return [];
    
    // As per user request, row 1 is junk, row 2 (index 1) is the header
    const headerRowIndex = 1;
    
    const headers = rows[headerRowIndex].map(h => String(h).trim());
    const data = [];
    
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        const obj = {};
        // Only add rows that have at least some data
        let hasData = false;
        headers.forEach((header, index) => {
            obj[header] = row[index];
            if (row[index]) hasData = true;
        });
        if (hasData) {
            // Concatenate EHR names to match IRIS "Last, First" format
            if (obj['patient lastname'] !== undefined && obj['patient firstname'] !== undefined) {
                const last = String(obj['patient lastname'] || '').trim();
                const first = String(obj['patient firstname'] || '').trim();
                if (last && first) {
                    obj['Patient Name'] = `${last}, ${first}`;
                } else if (last || first) {
                    obj['Patient Name'] = last || first;
                }
            }
            data.push(obj);
        }
    }
    return data;
}

async function parseCombinedExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const ehrSheet = workbook.Sheets["EHR Data"];
                const irisSheet = workbook.Sheets["IRIS Data"];
                
                const ehr = ehrSheet ? XLSX.utils.sheet_to_json(ehrSheet, { defval: "" }) : [];
                const iris = irisSheet ? XLSX.utils.sheet_to_json(irisSheet, { defval: "" }) : [];
                resolve({ ehr, iris });
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = (err) => reject(err);
        reader.readAsArrayBuffer(file);
    });
}

async function loadAllData() {
    ehrData = [];
    irisData = [];
    
    if (combinedFiles.length > 0) {
        document.getElementById('data-status').innerText = 'Parsing combined Excel...';
        const data = await parseCombinedExcel(combinedFiles[0]);
        ehrData = data.ehr;
        irisData = data.iris;
    } else {
        document.getElementById('data-status').innerText = 'Parsing files...';

        for (let file of ehrFiles) {
            const data = file.name.toLowerCase().endsWith('.csv') ? await parseCSV(file) : await parseExcel(file);
            ehrData.push(...data);
        }
        
        for (let file of irisFiles) {
            const data = file.name.toLowerCase().endsWith('.csv') ? await parseCSV(file) : await parseExcel(file);
            irisData.push(...data);
        }
    }

    document.getElementById('data-status').innerHTML = `✅ Data Loaded!`;
    
    // Filter IRIS data for preview to only show "Immunizations Given"
    const irisPreviewData = irisData.filter(row => {
        const reason = String(row['Transaction Reason'] || '').trim().toLowerCase();
        return reason === 'immunizations given';
    });
    
    // Render the raw data tables for visual human review
    renderRawTable('raw-ehr-table', ehrData, ['Patient Name', 'vaccine lot #', 'vaccine admin date', 'vaccine administered by']);
    renderRawTable('raw-iris-table', irisPreviewData, ['Patient Name', 'Lot Number', 'Transaction Reason', 'Quantity', 'Transaction Date']);
}

function renderRawTable(tableId, data, columns) {
    const tableEl = document.getElementById(tableId);
    if (!tableEl) return;
    
    // Destroy existing table instance before rebuilding DOM
    if (dtInstances[tableId]) {
        dtInstances[tableId].destroy();
        delete dtInstances[tableId];
    }
    
    const thead = tableEl.querySelector('thead');
    const tbody = tableEl.querySelector('tbody');
    
    thead.innerHTML = '';
    tbody.innerHTML = '';
    
    // Create Headers
    const trHead = document.createElement('tr');
    columns.forEach(col => {
        const th = document.createElement('th');
        th.innerText = col;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);
    
    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${columns.length}" class="empty-state">No data available.</td></tr>`;
    } else {
        // Create Rows
        data.forEach(row => {
            const tr = document.createElement('tr');
            columns.forEach(col => {
                const td = document.createElement('td');
                td.innerText = row[col] || '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }
    
    document.getElementById('raw-data-section').classList.remove('hidden');
    
    if (data.length > 0) {
        initDataTable(tableId);
    }
}

async function exportExcel() {
    await loadAllData();
    
    const wb = XLSX.utils.book_new();
    
    const ehrSheet = XLSX.utils.json_to_sheet(ehrData);
    XLSX.utils.book_append_sheet(wb, ehrSheet, "EHR Data");
    
    const irisSheet = XLSX.utils.json_to_sheet(irisData);
    XLSX.utils.book_append_sheet(wb, irisSheet, "IRIS Data");
    
    XLSX.writeFile(wb, "Combined_Vaccine_Data.xlsx");
}

function cleanString(str) {
    if (!str) return "";
    return String(str).trim().toLowerCase();
}

function normalizeName(str) {
    if (!str) return "";
    return String(str).toLowerCase().replace(/\s*,\s*/g, ', ').replace(/\s+/g, ' ').trim();
}

async function runQuery() {
    await loadAllData();
    
    document.getElementById('data-status').innerText = 'Running query...';

    const records = {}; // Key: "name|lot", Value: { name, lot, ehrCount, irisCount }
    const infoRecordsByReason = {}; // Store informational IRIS rows by reason
    const lotTracker = {}; // Key: "lot", Value: { received: 0, irisGiven: 0, ehrGiven: 0 }

    // Process EHR Data
    ehrData.forEach(row => {
        const name = normalizeName(row['Patient Name']);
        const lot = cleanString(row['vaccine lot #']);
        
        if (lot) {
            if (!lotTracker[lot]) lotTracker[lot] = { received: 0, irisGiven: 0, ehrGiven: 0 };
            lotTracker[lot].ehrGiven += 1;
        }

        if (!name && !lot) return; // Skip empty rows

        const key = `${name}|${lot}`;
        if (!records[key]) {
            records[key] = { name: name, lot: lot, ehrCount: 0, irisCount: 0 };
        }
        records[key].ehrCount += 1;
    });

    // Process IRIS Data
    irisData.forEach(row => {
        const name = normalizeName(row['Patient Name']);
        const lot = cleanString(row['Lot Number']);
        const reason = cleanString(row['Transaction Reason']);
        const qty = parseFloat(row['Quantity']) || 0;

        // Tracking for Lot KPI
        if (lot) {
            if (!lotTracker[lot]) lotTracker[lot] = { received: 0, irisGiven: 0, ehrGiven: 0 };
            
            if (reason === 'immunizations given') {
                lotTracker[lot].irisGiven += Math.abs(qty);
            } else if (reason === 'receipt of inventory') {
                lotTracker[lot].received += qty;
            }
        }

        // User requested ONLY 'immunizations given' run against EHR data
        if (reason === 'immunizations given') {
            if (!name) return; // Skip empty rows for actual patient records
            
            const key = `${name}|${lot}`;
            if (!records[key]) {
                records[key] = { name: name, lot: lot, ehrCount: 0, irisCount: 0 };
            }
            
            // Quantity is negative for "Given"
            // We want net given, so -1 * qty
            records[key].irisCount += (-1 * qty);
        } else {
            // It's informational, group by reason
            const rawReason = row['Transaction Reason'] || 'Unknown Reason';
            if (!infoRecordsByReason[rawReason]) {
                infoRecordsByReason[rawReason] = [];
            }
            infoRecordsByReason[rawReason].push(row);
        }
    });

    // Categorize
    const missingIris = [];
    const missingEhr = [];
    const mismatch = [];
    const allDiscrepancies = [];

    Object.keys(records).forEach(key => {
        const rec = records[key];
        // Round iris count to avoid float precision issues
        rec.irisCount = Math.round(rec.irisCount);
        
        let type = '';
        if (rec.ehrCount > 0 && rec.irisCount === 0) {
            missingIris.push(rec);
            type = 'Missing in IRIS';
        } else if (rec.ehrCount === 0 && rec.irisCount > 0) {
            missingEhr.push(rec);
            type = 'Missing in EHR';
        } else if (rec.ehrCount !== rec.irisCount && rec.ehrCount > 0 && rec.irisCount > 0) {
            mismatch.push(rec);
            type = 'Mismatched Counts';
        }
        
        if (type) {
            const uniqueKey = `${rec.name}|${rec.lot}|${type}`;
            allDiscrepancies.push({
                ...rec,
                type: type,
                uniqueKey: uniqueKey
            });
        }
    });
    
    // Store globally so the clear button can re-render it
    window.currentWorkingListData = allDiscrepancies;

    // Update UI
    document.getElementById('count-missing-iris').innerText = missingIris.length;
    document.getElementById('count-missing-ehr').innerText = missingEhr.length;
    document.getElementById('count-mismatch').innerText = mismatch.length;

    renderTable('table-missing-iris', missingIris);
    renderTable('table-missing-ehr', missingEhr);
    renderTable('table-mismatch', mismatch);
    
    // Render Lot KPI Table
    renderLotKpiTable('table-lot-kpi', lotTracker);
    
    // Render Working List Table
    renderWorkingList('table-working-list', allDiscrepancies);
    
    // Render the informational IRIS records dynamically
    renderInfoTabs(infoRecordsByReason);

    document.getElementById('results-section').classList.remove('hidden');
    document.getElementById('working-list-section').classList.remove('hidden');
    
    if (Object.keys(infoRecordsByReason).length > 0) {
        document.getElementById('info-section').classList.remove('hidden');
    }
    
    document.getElementById('data-status').innerHTML = `✅ Query Complete`;
}

function renderTable(tableId, data) {
    if (dtInstances[tableId]) {
        dtInstances[tableId].destroy();
        delete dtInstances[tableId];
    }
    
    const tbody = document.getElementById(tableId).querySelector('tbody');
    tbody.innerHTML = '';
    
    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="empty-state">No discrepancies found in this category.</td></tr>`;
        return;
    }

    data.forEach(rec => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="text-transform: capitalize">${rec.name}</td>
            <td style="text-transform: uppercase">${rec.lot}</td>
            <td>${rec.ehrCount}</td>
            <td>${rec.irisCount}</td>
        `;
        tbody.appendChild(tr);
    });
    
    initDataTable(tableId);
}

function renderLotKpiTable(tableId, lotTracker) {
    if (dtInstances[tableId]) {
        dtInstances[tableId].destroy();
        delete dtInstances[tableId];
    }
    
    const tbody = document.getElementById(tableId).querySelector('tbody');
    tbody.innerHTML = '';
    
    const lots = Object.keys(lotTracker).sort();
    
    if (lots.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="empty-state">No lot numbers found.</td></tr>`;
        return;
    }

    lots.forEach(lot => {
        const data = lotTracker[lot];
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="text-transform: uppercase; font-weight: bold;">${lot}</td>
            <td>${data.received}</td>
            <td>${data.irisGiven}</td>
            <td>${data.ehrGiven}</td>
        `;
        tbody.appendChild(tr);
    });
    
    initDataTable(tableId);
}

function renderWorkingList(tableId, data) {
    if (dtInstances[tableId]) {
        dtInstances[tableId].destroy();
        delete dtInstances[tableId];
    }
    
    const tableEl = document.getElementById(tableId);
    const tbody = tableEl.querySelector('tbody');
    tbody.innerHTML = '';
    
    if (data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-state">No discrepancies found! You're all caught up.</td></tr>`;
        return;
    }

    const cache = getResolvedCache();

    data.forEach(rec => {
        const tr = document.createElement('tr');
        const isChecked = cache[rec.uniqueKey] ? 'checked' : '';
        
        tr.innerHTML = `
            <td style="text-align: center;">
                <input type="checkbox" class="resolved-checkbox" data-key="${rec.uniqueKey}" ${isChecked} style="width: 20px; height: 20px; cursor: pointer;">
            </td>
            <td style="text-transform: capitalize">${rec.name}</td>
            <td style="text-transform: uppercase">${rec.lot}</td>
            <td><span class="badge ${rec.type === 'Missing in EHR' ? 'error' : (rec.type === 'Missing in IRIS' ? 'warning' : 'caution')}">${rec.type}</span></td>
            <td>${rec.ehrCount}</td>
            <td>${rec.irisCount}</td>
        `;
        tbody.appendChild(tr);
    });
    
    // Set initial opacity before datatable init just to be safe
    tbody.querySelectorAll('.resolved-checkbox:checked').forEach(cb => {
        cb.closest('tr').style.opacity = '0.5';
    });
    
    initDataTable(tableId);
    
    // DataTables aggressively redraws the DOM on pagination/sorting, which erases checkbox states.
    // Use a MutationObserver to constantly force the DOM checkboxes to match our LocalStorage cache.
    const dtWrapper = tableEl.closest('.dataTable-wrapper');
    if (dtWrapper) {
        const observer = new MutationObserver(() => {
            const cache = getResolvedCache();
            dtWrapper.querySelectorAll('.resolved-checkbox').forEach(cb => {
                const key = cb.getAttribute('data-key');
                cb.checked = !!cache[key];
                if (cb.checked) {
                    cb.closest('tr').style.opacity = '0.5';
                } else {
                    cb.closest('tr').style.opacity = '1';
                }
            });
        });
        
        const newTbody = dtWrapper.querySelector('tbody');
        if (newTbody) {
            observer.observe(newTbody, { childList: true, subtree: true });
            
            // Trigger initially
            const cache = getResolvedCache();
            newTbody.querySelectorAll('.resolved-checkbox').forEach(cb => {
                const key = cb.getAttribute('data-key');
                cb.checked = !!cache[key];
                if (cb.checked) {
                    cb.closest('tr').style.opacity = '0.5';
                }
            });
        }
    }
}

function renderInfoTabs(recordsByReason) {
    const tabsContainer = document.getElementById('info-tabs');
    const contentsContainer = document.getElementById('info-tab-contents');
    
    tabsContainer.innerHTML = '';
    contentsContainer.innerHTML = '';
    
    const reasons = Object.keys(recordsByReason).sort();
    if (reasons.length === 0) return;
    
    reasons.forEach((reason, index) => {
        const idFriendly = 'info-tab-' + reason.toLowerCase().replace(/[^a-z0-9]/g, '-');
        
        // Create tab button
        const btn = document.createElement('button');
        btn.className = `tab ${index === 0 ? 'active' : ''}`;
        btn.setAttribute('data-target', idFriendly);
        btn.innerText = `${reason} (${recordsByReason[reason].length})`;
        tabsContainer.appendChild(btn);
        
        // Create tab content container
        const contentDiv = document.createElement('div');
        contentDiv.className = `tab-content ${index === 0 ? 'active' : ''}`;
        contentDiv.id = idFriendly;
        
        // Create table
        const tableId = `table-${idFriendly}`;
        contentDiv.innerHTML = `
            <table id="${tableId}">
                <thead>
                    <tr>
                        <th>Patient Name</th>
                        <th>Lot Number</th>
                        <th>Quantity</th>
                        <th>Transaction Date</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        `;
        contentsContainer.appendChild(contentDiv);
        
        const tbody = contentDiv.querySelector('tbody');
        recordsByReason[reason].forEach(row => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td style="text-transform: capitalize">${row['Patient Name'] || ''}</td>
                <td style="text-transform: uppercase">${row['Lot Number'] || ''}</td>
                <td>${row['Quantity'] || ''}</td>
                <td>${row['Transaction Date'] || ''}</td>
            `;
            tbody.appendChild(tr);
        });
        
        // Initialize datatable
        initDataTable(tableId);
    });
    
    // Add event listeners to new tabs
    tabsContainer.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabsContainer.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            contentsContainer.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            e.target.classList.add('active');
            document.getElementById(e.target.getAttribute('data-target')).classList.add('active');
        });
    });
}
