let groupedData = {};
let selectedCategories = new Set(["Frame", "Lens", "Contactlens", "Service", "Accessories", "น้ำยา"]);

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const filterSection = document.getElementById('filterSection');
const groupList = document.getElementById('groupList');
const previewSection = document.getElementById('previewSection');
const previewTableBody = document.querySelector('#previewTable tbody');
const downloadBtn = document.getElementById('downloadBtn');
const errorBanner = document.getElementById('errorBanner');
const previewTabBar = document.getElementById('previewTabBar');

// Admin Elements
const adminBtn = document.getElementById('adminBtn');
const adminModal = document.getElementById('adminModal');
const closeModal = adminModal.querySelector('.close');
const adminPasswordInput = document.getElementById('adminPassword');
const loginBtn = document.getElementById('loginBtn');
const passwordSection = document.getElementById('passwordSection');
const settingsSection = document.getElementById('settingsSection');
const mappingList = document.getElementById('mappingList');
const addMappingBtn = document.getElementById('addMappingBtn');
const saveMappingBtn = document.getElementById('saveMappingBtn');
const supabaseUrlInput = document.getElementById('supabaseUrl');
const supabaseKeyInput = document.getElementById('supabaseKey');
const importMappingBtn = document.getElementById('importMappingBtn');
const mappingFileInput = document.getElementById('mappingFileInput');

let supabaseClient = null;
let brandMappings = {};
let activePreviewTab = 'All';

// --- SUPABASE CONFIG (HARDCODED) ---
const DEFAULT_SUPABASE_URL = 'https://qvuviyueajtchprbafbk.supabase.co';
const DEFAULT_SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InF2dXZpeXVlYWp0Y2hwcmJhZmJrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc4NDYwMjgsImV4cCI6MjA5MzQyMjAyOH0.JapT1lildN2H_9TeHF_iNiL3ABbT7NJewaf_wqeT0Cg';

async function initSupabase() {
    // Users will use hardcoded values, Admin can override via UI
    const url = DEFAULT_SUPABASE_URL;
    const key = DEFAULT_SUPABASE_KEY;

    if (url && key && typeof supabase !== 'undefined') {
        supabaseClient = supabase.createClient(url, key);
        await syncMappings();
    } else {
        console.warn("Supabase SDK not loaded or credentials missing.");
    }
}

// Initial Sync for all users (Background)
initSupabase();

async function syncMappings() {
    if (!supabaseClient) return;
    try {
        const { data, error } = await supabaseClient.from('brand_mappings').select('*');
        if (error) throw error;

        const newMappings = {};
        data.forEach(item => {
            newMappings[item.original_name] = item.replacement_name;
        });
        brandMappings = newMappings;
        console.log("Mappings synced from Supabase:", Object.keys(brandMappings).length);
    } catch (err) {
        console.error("Supabase sync error:", err);
        showToast(`Sync ล้มเหลว: ${err.message || 'โปรดตรวจสอบชื่อตาราง brand_mappings'}`, 'warning');
    }
}

const CATEGORY_MAP = {
    'Frame': 'Frame',
    'Lens': 'Lens',
    'Contactlens': 'Contactlens',
    'Service': 'Service',
    'Accessories': 'Accessories',
    '': 'น้ำยา'
};

const TARGET_CATEGORIES = ["Frame", "Lens", "Contactlens", "Service", "Accessories", "น้ำยา"];

// UI Events
dropZone.onclick = () => fileInput.click();
fileInput.onchange = (e) => handleFile(e.target.files[0]);

dropZone.ondragover = (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
};

dropZone.ondragleave = () => dropZone.classList.remove('dragover');

dropZone.ondrop = (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
};

// Search Logic removed

// Admin Logic
adminBtn.onclick = () => {
    adminModal.style.display = 'block';
    passwordSection.style.display = 'flex';
    settingsSection.style.display = 'none';
    adminPasswordInput.value = '';
    loginError.style.display = 'none';
    setTimeout(() => adminPasswordInput.focus(), 100);
};

const adminCancelBtn = document.getElementById('adminCancelBtn');
const togglePasswordBtn = document.getElementById('togglePasswordBtn');
const loginError = document.getElementById('loginError');

closeModal.onclick = () => adminModal.style.display = 'none';
adminCancelBtn.onclick = () => adminModal.style.display = 'none';
window.onclick = (e) => { if (e.target == adminModal) adminModal.style.display = 'none'; };

togglePasswordBtn.onclick = () => {
    const isPassword = adminPasswordInput.type === 'password';
    adminPasswordInput.type = isPassword ? 'text' : 'password';
    togglePasswordBtn.innerHTML = isPassword
        ? `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"></path><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`
        : `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
};

// Enter key on password field
adminPasswordInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') loginBtn.click();
});

// Helper to hash password (SHA-256)
async function hashPassword(password) {
    const msgUint8 = new TextEncoder().encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgUint8);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

loginBtn.onclick = async () => {
    const inputHash = await hashPassword(adminPasswordInput.value);

    if (inputHash === '27a307d3e70ee464d6cbdd13812e501e3010ada7318ddac7c5d3696d9613df0c') {
        loginError.style.display = 'none';
        passwordSection.style.display = 'none';
        settingsSection.style.display = 'block';

        // Reveal credentials only to Admin
        supabaseUrlInput.value = DEFAULT_SUPABASE_URL;
        supabaseKeyInput.value = DEFAULT_SUPABASE_KEY;

        await initSupabase();
        await syncMappings(); // Force re-sync to be sure
        renderMappings();
        showToast('เข้าสู่ระบบ Admin สำเร็จ', 'success');
    } else {
        loginError.style.display = 'block';
        adminPasswordInput.value = '';
        adminPasswordInput.focus();
    }
};

function renderMappings() {
    mappingList.innerHTML = '';
    Object.entries(brandMappings).forEach(([key, val], index) => {
        addMappingRow(key, val);
    });
}

function addMappingRow(key = '', val = '') {
    const div = document.createElement('div');
    div.className = 'mapping-row';
    div.innerHTML = `
        <input type="text" class="map-key settings-input" placeholder="Original (เช่น FMT หรือ 8851234567)" value="${key}">
        <span style="text-align:center; color: var(--text-muted);">&#8594;</span>
        <input type="text" class="map-val settings-input" placeholder="Replace with (เช่น Mykita)" value="${val}">
        <button class="remove-mapping" title="ลบ">&times;</button>
    `;
    div.querySelector('.remove-mapping').onclick = () => div.remove();
    mappingList.appendChild(div);
}

addMappingBtn.onclick = () => addMappingRow();

// Import Mappings from Excel
importMappingBtn.onclick = () => mappingFileInput.click();
mappingFileInput.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        json.forEach((row, i) => {
            if (i === 0 || !row[0]) return; // Skip header or empty
            addMappingRow(String(row[0]).trim(), String(row[1] || '').trim());
        });
        mappingFileInput.value = ''; // Reset
    };
    reader.readAsArrayBuffer(file);
};

saveMappingBtn.onclick = async () => {
    if (!supabaseClient) {
        // Fallback to init if user just pasted credentials
        const url = supabaseUrlInput.value.trim();
        const key = supabaseKeyInput.value.trim();
        if (!url || !key) return alert('กรุณาระบุ Supabase URL และ Key ก่อนบันทึก!');
        supabaseClient = supabase.createClient(url, key);
        localStorage.setItem('supabaseUrl', url);
        localStorage.setItem('supabaseKey', key);
    }

    const rowsToSave = [];
    document.querySelectorAll('.mapping-row').forEach(row => {
        const key = row.querySelector('.map-key').value.trim();
        const val = row.querySelector('.map-val').value.trim();
        if (key) rowsToSave.push({ original_name: key, replacement_name: val });
    });

    try {
        saveMappingBtn.textContent = "⌛ กำลังบันทึก...";
        saveMappingBtn.disabled = true;

        // Simple sync strategy: Clear and re-insert
        await supabaseClient.from('brand_mappings').delete().neq('id', 0);
        const { error } = await supabaseClient.from('brand_mappings').insert(rowsToSave);

        if (error) throw error;

        await syncMappings();
        alert('บันทึกข้อมูลลง Supabase เรียบร้อยแล้ว!');
        adminModal.style.display = 'none';
        if (Object.keys(groupedData).length > 0) updatePreview();
    } catch (err) {
        console.error("Save error:", err);
        alert('เกิดข้อผิดพลาดในการบันทึก: ' + err.message);
    } finally {
        saveMappingBtn.textContent = "บันทึกลง Database (Supabase)";
        saveMappingBtn.disabled = false;
    }
};

function getMappedName(category, rawDept, matCode) {
    if (!rawDept) rawDept = 'General';
    
    // Clean rawDept (remove "Dept Name:" prefix and trim)
    let deptName = String(rawDept).replace(/Dept\s*Name:\s*/gi, '').trim();
    const cat = String(category || '').trim().toLowerCase();
    
    if (cat === 'frame') {
        // If Group name = frame, lookup using Deptname from database (brandMappings)
        if (brandMappings[deptName]) {
            return brandMappings[deptName];
        }
        return deptName;
    } else if (cat === 'contactlens') {
        // If Group name = Contactlens, use first 10 characters of Matcode to lookup
        const code = String(matCode || '').trim();
        const first10 = code.substring(0, 10);
        if (brandMappings[first10]) {
            return brandMappings[first10];
        }
        return deptName; // Fallback to raw/cleaned Dept name if not found
    } else {
        // For other groups, lookup using Deptname from database as fallback
        if (brandMappings[deptName]) {
            return brandMappings[deptName];
        }
        return deptName;
    }
}

async function handleFile(file) {
    if (!file) return;

    // Visual Loading State
    const dropZoneText = dropZone.querySelector('p');
    const originalText = dropZoneText.textContent;
    dropZoneText.textContent = "⌛ กำลังประมวลผลไฟล์... (Processing)";
    dropZone.style.opacity = "0.6";
    dropZone.style.pointerEvents = "none";

    resetState();

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log("File loaded, rows:", json.length);
            const result = detectPattern(json);

            if (result && result.type) {
                console.log("Pattern detected:", result.type, result.mapping);
                processRawData(json, result);
            } else {
                console.warn("No pattern detected.", result);
                showFormatError(result ? result.missingColumns : []);
            }
        } catch (err) {
            console.error("Error processing file:", err);
            showError();
        } finally {
            // Restore UI
            dropZoneText.textContent = originalText;
            dropZone.style.opacity = "1";
            dropZone.style.pointerEvents = "all";
        }
    };
    reader.onerror = () => {
        showError();
        dropZoneText.textContent = originalText;
        dropZone.style.opacity = "1";
        dropZone.style.pointerEvents = "all";
    };
    reader.readAsArrayBuffer(file);
}

function detectPattern(rows) {
    // 1. Check for Legacy Dept Marker (Row-based)
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
        const firstCell = String(rows[i][0] || '').trim();
        if (firstCell.startsWith('Dept Name:')) return { type: 'LEGACY_DEPT_MARKER' };
    }

    // 2. Smart Header Detection (Search for column keywords)
    const headerKeywords = {
        code: ['mat code', 'รหัสสินค้า', 'product code', 'item code'],
        name: ['mat name', 'ชื่อสินค้า', 'product name', 'item name', 'description'],
        balance: ['current balance', 'คงเหลือ', 'ยอดคงเหลือ', 'qty', 'balance', 'on hand'],
        dept: ['dept name', 'ชื่อแผนก', 'department'],
        groupCode: ['group code', 'รหัสกลุ่ม'],
        groupName: ['group name', 'ชื่อกลุ่ม']
    };

    // Track best candidate row for diagnostics
    let bestMatch = { matchCount: 0, foundKeys: [], row: -1 };

    for (let i = 0; i < Math.min(rows.length, 100); i++) {
        const row = rows[i];
        if (!row || row.length < 3) continue;

        const mapping = {};
        let matchCount = 0;
        const foundKeys = [];

        row.forEach((cell, index) => {
            if (cell === undefined || cell === null) return;
            const cellText = String(cell).toLowerCase().trim();
            if (!cellText) return;

            for (const [key, aliases] of Object.entries(headerKeywords)) {
                const isMatch = aliases.some(alias => cellText.includes(alias));
                if (isMatch && mapping[key] === undefined) {
                    mapping[key] = index;
                    foundKeys.push(key);
                    matchCount++; // Increment for any recognized column
                }
            }
        });

        // Found the core columns? (Strict: Need code, name, balance, and groupName)
        const hasRequired = mapping.code !== undefined &&
            mapping.name !== undefined &&
            mapping.balance !== undefined &&
            mapping.groupName !== undefined;

        if (hasRequired) {
            return {
                type: 'FLAT_TABLE',
                headerRowIndex: i,
                mapping: mapping
            };
        }

        // Keep track of the best partial match for diagnostics
        if (matchCount > bestMatch.matchCount) {
            bestMatch = { matchCount, foundKeys, row: i };
        }
    }

    // Build diagnostic: which required columns were missing
    const requiredLabels = {
        code: 'Mat Code / รหัสสินค้า',
        name: 'Mat Name / ชื่อสินค้า',
        balance: 'Balance / ยอดคงเหลือ',
        groupName: 'Group Name / ชื่อกลุ่ม'
    };
    const missing = Object.keys(requiredLabels).filter(k => !bestMatch.foundKeys.includes(k));
    const missingColumns = missing.map(k => requiredLabels[k]);

    return { type: null, missingColumns };
}

function showError() {
    errorBanner.style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function showFormatError(missingColumns = []) {
    showError();

    // Build a descriptive toast message with solutions
    let title = 'รูปแบบไฟล์ไม่ถูกต้อง';
    let body = 'ระบบไม่สามารถตรวจพบข้อมูลที่จำเป็นได้ครบถ้วน:\n';

    if (missingColumns.length > 0) {
        missingColumns.forEach(col => body += `\n  ❌ ไม่พบ: ${col}`);
    } else {
        body += '\n  ❌ ไม่พบคอลัมน์ Group Name (ชื่อกลุ่ม)';
    }

    body += '\n\n----------------------------\n';
    body += '💡 แนวทางการแก้ไข:\n';
    body += '1. ตรวจสอบว่าชื่อหัวคอลัมน์ในไฟล์ตรงตามมาตรฐาน (เช่น Dept Name, Group Name)\n';
    body += '2. ตรวจสอบว่าคุณอัปโหลดไฟล์ "รายงานคงเหลือตามคลังสินค้า" จากโปรแกรม Inventory\n';
    body += '3. ตรวจสอบว่าข้อมูลใน Excel เริ่มต้นที่ Sheet แรกเสมอ\n';
    body += '4. หากยังพบปัญหา โปรดดู "วิธีโหลดไฟล์" ที่ปุ่มด้านล่างขวาของหน้าจอ';

    showToast(`${title}\n${body}`, 'error');
}

function showToast(message, type = 'info') {
    // Remove existing toast
    const existing = document.getElementById('toast-notification');
    if (existing) existing.remove();

    const colors = {
        error: { bg: '#dc2626', icon: '❌' },
        warning: { bg: '#d97706', icon: '⚠️' },
        success: { bg: '#16a34a', icon: '✅' },
        info: { bg: '#2563eb', icon: 'ℹ️' }
    };
    const { bg, icon } = colors[type] || colors.info;

    const toast = document.createElement('div');
    toast.id = 'toast-notification';
    toast.style.cssText = `
        position: fixed;
        bottom: 2rem;
        right: 2rem;
        background: ${bg};
        color: white;
        padding: 1.25rem 1.5rem;
        border-radius: 16px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.35);
        z-index: 9999;
        max-width: 380px;
        font-family: 'Prompt', sans-serif;
        font-size: 0.9rem;
        line-height: 1.6;
        white-space: pre-line;
        animation: slideUp 0.3s ease-out;
        cursor: pointer;
    `;
    const lines = message.split('\n');
    toast.innerHTML = `<strong style="font-size:1rem; display:block; margin-bottom:0.4rem;">${icon} ${lines[0]}</strong>${lines.slice(1).join('<br>')}`;
    toast.onclick = () => toast.remove();
    document.body.appendChild(toast);

    // Auto remove after 10 seconds
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 10000);
}

function resetState() {
    groupedData = {};
    activePreviewTab = 'All';
    previewTabBar.innerHTML = '';
    previewTabBar.style.display = 'none';
    errorBanner.style.display = 'none';
    filterSection.style.display = 'none';
    previewSection.style.display = 'none';
    previewTableBody.innerHTML = '';
}

function processRawData(rows, result) {
    groupedData = {};
    const pattern = result.type;

    if (pattern === 'LEGACY_DEPT_MARKER') {
        let currentDeptRaw = null;
        rows.forEach(row => {
            if (!row || row.length === 0) return;
            const firstCell = String(row[0] || '').trim();

            if (firstCell.startsWith('Dept Name:')) {
                currentDeptRaw = firstCell.replace(/Dept\s*Name:\s*/gi, '').trim();
            }
            else if (currentDeptRaw && row[3] && row[0]) {
                let rawCat = String(row[0] || '').trim();
                let cat = CATEGORY_MAP[rawCat] || rawCat;
                let balance = parseFloat(row[15]) || 0;

                if (balance === 0) return;

                let code = String(row[3]).trim();
                let finalDeptName = getMappedName(cat, currentDeptRaw, code);

                if (!groupedData[finalDeptName]) {
                    groupedData[finalDeptName] = [];
                }

                groupedData[finalDeptName].push({
                    category: cat,
                    type: row[1],
                    dept: finalDeptName,
                    code: code,
                    description: row[4],
                    balance: balance
                });
            }
        });
    } else if (pattern === 'FLAT_TABLE') {
        const { headerRowIndex, mapping } = result;

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || !row[mapping.code]) continue;

            let rawCat = mapping.groupName !== undefined ? String(row[mapping.groupName] || '').trim() : '';
            let cat = CATEGORY_MAP[rawCat] || rawCat;
            let balance = parseFloat(row[mapping.balance]) || 0;

            if (balance === 0) continue; // Fix: use continue instead of return to skip 0 balance items

            const rawDept = mapping.dept !== undefined ? String(row[mapping.dept] || '').trim() : 'General';
            const code = String(row[mapping.code]).trim();
            const finalDeptName = getMappedName(cat, rawDept, code);

            if (!groupedData[finalDeptName]) {
                groupedData[finalDeptName] = [];
            }

            groupedData[finalDeptName].push({
                category: cat,
                type: mapping.groupCode !== undefined ? String(row[mapping.groupCode] || '').trim() : '',
                dept: finalDeptName,
                code: code,
                description: String(row[mapping.name] || '').trim(),
                balance: balance
            });
        }
    }

    renderFilters();
    updatePreview();
}

function renderFilters() {
    groupList.innerHTML = '';
    TARGET_CATEGORIES.forEach(cat => {
        const div = document.createElement('div');
        div.className = 'group-item';
        div.innerHTML = `
            <input type="checkbox" id="chk-${cat}" ${selectedCategories.has(cat) ? 'checked' : ''}>
            <label for="chk-${cat}">${cat}</label>
        `;
        div.querySelector('input').onchange = (e) => {
            if (e.target.checked) selectedCategories.add(cat);
            else selectedCategories.delete(cat);
            updatePreview();
        };
        groupList.appendChild(div);
    });
    filterSection.style.display = 'block';
}

document.getElementById('selectAll').onclick = () => {
    document.querySelectorAll('.group-item input').forEach(i => {
        i.checked = true;
        selectedCategories.add(i.id.replace('chk-', ''));
    });
    updatePreview();
};

document.getElementById('deselectAll').onclick = () => {
    document.querySelectorAll('.group-item input').forEach(i => {
        i.checked = false;
        selectedCategories.clear();
    });
    updatePreview();
};

function updatePreview() {
    previewTableBody.innerHTML = '';
    let rowCount = 0;

    // Find all categories that actually have data
    const categoriesWithData = new Set();
    Object.values(groupedData).forEach(items => {
        items.forEach(item => {
            if (selectedCategories.has(item.category)) {
                categoriesWithData.add(item.category);
            }
        });
    });

    const activeCategories = Array.from(categoriesWithData);

    // Keep active preview tab in check
    if (activePreviewTab !== 'All' && !categoriesWithData.has(activePreviewTab)) {
        activePreviewTab = 'All';
    }

    // Render Tab Bar if there are multiple active categories
    if (activeCategories.length > 1) {
        previewTabBar.innerHTML = '';
        previewTabBar.style.display = 'flex';

        // 1. All tab
        const totalItemsCount = Object.values(groupedData).reduce((sum, items) => {
            return sum + items.filter(item => selectedCategories.has(item.category)).length;
        }, 0);

        const allTab = document.createElement('div');
        allTab.className = `preview-tab ${activePreviewTab === 'All' ? 'active' : ''}`;
        allTab.innerHTML = `ทั้งหมด <span class="count-badge">${totalItemsCount}</span>`;
        allTab.onclick = () => {
            activePreviewTab = 'All';
            updatePreview();
        };
        previewTabBar.appendChild(allTab);

        // 2. Individual product tabs
        activeCategories.forEach(cat => {
            const catItemsCount = Object.values(groupedData).reduce((sum, items) => {
                return sum + items.filter(item => item.category === cat).length;
            }, 0);

            const tab = document.createElement('div');
            tab.className = `preview-tab ${activePreviewTab === cat ? 'active' : ''}`;
            tab.innerHTML = `${cat} <span class="count-badge">${catItemsCount}</span>`;
            tab.onclick = () => {
                activePreviewTab = cat;
                updatePreview();
            };
            previewTabBar.appendChild(tab);
        });
    } else {
        previewTabBar.style.display = 'none';
        activePreviewTab = 'All';
    }

    Object.keys(groupedData).forEach(dept => {
        let filteredItems = groupedData[dept].filter(item => {
            if (activePreviewTab === 'All') {
                return selectedCategories.has(item.category);
            } else {
                return item.category === activePreviewTab;
            }
        });

        if (filteredItems.length === 0) return;

        const subtotal = filteredItems.reduce((sum, item) => sum + item.balance, 0);

        const headerTr = document.createElement('tr');
        headerTr.className = 'dept-row';
        headerTr.innerHTML = `
            <td colspan="5">Dept Name: ${dept}</td>
            <td>${subtotal.toLocaleString()}</td>
            <td></td>
            <td></td>
            <td></td>
        `;
        previewTableBody.appendChild(headerTr);

        filteredItems.forEach(item => {
            rowCount++;
            if (rowCount > 100) return; // Limit preview for performance

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.category}</td>
                <td>${item.type}</td>
                <td>${item.dept}</td>
                <td>${item.code}</td>
                <td>${item.description}</td>
                <td>${item.balance.toLocaleString()}</td>
                <td></td>
                <td></td>
                <td></td>
            `;
            previewTableBody.appendChild(tr);
        });
    });

    previewSection.style.display = rowCount > 0 ? 'block' : 'none';
}

async function exportToExcel(branchCode) {
    const workbook = new ExcelJS.Workbook();

    // 1. Instruction Sheet (First Tab)
    const insSheet = workbook.addWorksheet('คู่มือการใช้งาน', { properties: { tabColor: { argb: 'FFFF0000' } } });
    insSheet.columns = [{ width: 5 }, { width: 80 }];

    const insTitle = insSheet.addRow(["", "คู่มือการใช้งานไฟล์ Audit Stock"]);
    insTitle.font = { size: 18, bold: true, color: { argb: 'FF4F46E5' } };
    insSheet.addRow([]); // Blank

    const instructions = [
        "1. ตรวจสอบข้อมูลแผนกและหมวดหมู่สินค้าในแต่ละแท็บ (เช่น Frame, Lens, Contactlens)",
        "2. กรอกจำนวนสินค้าที่นับได้จริงในคอลัมน์ 'Actual Count' (ช่องสีขาว)",
        "3. ระบบจะคำนวณผลต่าง (Variance) ให้โดยอัตโนมัติในคอลัมน์ 'Variance'",
        "4. การจัดการในรูปแบบ Group: สามารถใช้เครื่องหมาย (+) และ (-) ทางด้านซ้ายมือเพื่อย่อหรือขยายรายละเอียดในแต่ละ Group ได้",
        "5. ยอดรวมตาม Group: บรรทัดสีเทาเข้มจะแสดงผลรวมของสินค้าใน Group นั้นๆ ซึ่งจะขยับตามจำนวนที่คุณกรอกจริง"
    ];

    instructions.forEach((text, i) => {
        const row = insSheet.addRow(["", text]);
        row.font = { size: 12 };
        if (i === 1) row.getCell(2).font = { size: 12, bold: true, color: { argb: 'FFFF0000' } };
        insSheet.addRow([]); // Space between points
    });

    // 2. Audit Sheets (Separate Tab for each Category)
    const activeCategories = Array.from(selectedCategories).filter(cat => {
        return Object.values(groupedData).some(items => items.some(item => item.category === cat));
    });

    activeCategories.forEach(cat => {
        // Slick, harmonious brand-specific tab colors
        let tabColor = 'FF4F46E5'; // Default: Indigo
        if (cat === 'Frame') tabColor = 'FF4F46E5';
        else if (cat === 'Lens') tabColor = 'FF10B981'; // Emerald
        else if (cat === 'Contactlens') tabColor = 'FFF59E0B'; // Amber
        else if (cat === 'Service') tabColor = 'FF8B5CF6'; // Purple
        else if (cat === 'Accessories') tabColor = 'FFEC4899'; // Pink
        else if (cat === 'น้ำยา') tabColor = 'FF06B6D4'; // Cyan

        const worksheet = workbook.addWorksheet(cat, {
            views: [{ state: 'frozen', ySplit: 2 }],
            properties: {
                tabColor: { argb: tabColor },
                outlineLevelCol: 0,
                outlineLevelRow: 1,
                outlineProperties: { summaryBelow: false }
            }
        });

        // Main Header with Branch Code and Category
        const titleText = branchCode ? `Audit Stock Report (${cat}) สาขา ${branchCode}` : `Audit Stock Report (${cat})`;
        const titleRow = worksheet.addRow([titleText]);
        worksheet.mergeCells('A1:I1');
        titleRow.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.alignment = { vertical: 'middle', horizontal: 'center' };
        titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tabColor } };

        // Table Headers
        const headerRow = worksheet.addRow(["Category", "Type", "Dept", "Code", "Description", "System Stock", "Actual Count", "Variance", "Remark"]);
        headerRow.font = { bold: true };
        headerRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
            cell.border = { bottom: { style: 'thin' } };
        });

        Object.keys(groupedData).forEach(dept => {
            const filteredItems = groupedData[dept].filter(item => item.category === cat);
            if (filteredItems.length === 0) return;

            // Add Dept Name Header Row (Styled with #404040 background and white text)
            const deptRow = worksheet.addRow([`Dept Name: ${dept}`, "", "", "", "", 0, 0, 0, ""]);
            deptRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // White Text
            deptRow.eachCell(cell => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } }; // Dark Gray
            });

            const startRow = worksheet.rowCount + 1;
            filteredItems.forEach(item => {
                const row = worksheet.addRow([
                    item.category,
                    item.type,
                    item.dept,
                    item.code,
                    item.description,
                    item.balance,
                    null,
                    null,
                    null
                ]);
                row.outlineLevel = 1;
                const rowIndex = row.number;
                row.getCell(8).value = { formula: `G${rowIndex}-F${rowIndex}` };
            });

            // Add 2 blank rows for manual entry per Dept
            for (let j = 0; j < 2; j++) {
                const blankRow = worksheet.addRow(["", "", "", "", "-", 0, null, null, null]);
                blankRow.outlineLevel = 1;
                const rowIndex = blankRow.number;
                blankRow.getCell(8).value = { formula: `G${rowIndex}-F${rowIndex}` };

                // Add borders to blank row cells
                blankRow.eachCell({ includeEmpty: true }, (cell) => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            }

            const endRow = worksheet.rowCount;

            // Set dynamic SUM formulas for System and Actual
            deptRow.getCell(6).value = { formula: `SUM(F${startRow}:F${endRow})` };
            deptRow.getCell(7).value = { formula: `SUM(G${startRow}:G${endRow})` };

            // Variance formula for Dept row: Actual (G) - System (F)
            const deptRowIndex = deptRow.number;
            deptRow.getCell(8).value = { formula: `G${deptRowIndex}-F${deptRowIndex}` };
        });

        // Column Widths
        worksheet.columns = [
            { width: 15 }, { width: 10 }, { width: 10 }, { width: 25 }, { width: 45 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 25 }
        ];
    });

    // Filename logic: include current date (YYYYMMDD) in filename
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const dateStr = `${yyyy}${mm}${dd}`;

    const cleanBranch = branchCode ? String(branchCode).trim() : '';
    const filename = cleanBranch 
        ? `Audit_Stock_${cleanBranch}_${dateStr}.xlsx` 
        : `Audit_Stock_${dateStr}.xlsx`;

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), filename);
}

downloadBtn.onclick = () => {
    const branchCode = prompt("กรุณากรอกรหัสสาขาเพื่อระบุในรายงานและชื่อไฟล์:");
    if (branchCode === null) {
        // Cancel download if user clicks Cancel
        return;
    }
    exportToExcel(branchCode);
};
