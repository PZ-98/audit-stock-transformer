let groupedData = {};
let selectedCategories = new Set(["Frame", "Lens", "Contactlens", "Accessories", "น้ำยา", "ไม่พบข้อมูลสินค้า"]);

// Raw Stock State Variables (for re-processing upon scan file updates)
let rawStockRows = null;
let rawStockResult = null;
let rawStockFileName = '';
let rawStockFileSize = 0;

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
    } catch (err) {
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

function formatFileBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const kilobytes = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const index = Math.floor(Math.log(bytes) / Math.log(kilobytes));
    const formattedSize = parseFloat((bytes / Math.pow(kilobytes, index)).toFixed(2));
    return `${formattedSize} ${sizes[index]}`;
}

function translatePatternType(type) {
    if (type === 'LEGACY_DEPT_MARKER') {
        return 'ตารางแยกแผนก (Legacy)';
    }
    if (type === 'FLAT_TABLE') {
        return 'ตารางทั่วไป (Flat Table)';
    }
    return 'ไม่ระบุรูปแบบ';
}

function calculateCategoryStatistics(dataGroup) {
    const stats = {};
    const categoriesToCount = [...TARGET_CATEGORIES, 'ไม่พบข้อมูลสินค้า'];
    categoriesToCount.forEach(category => {
        stats[category] = 0;
    });

    Object.values(dataGroup).forEach(items => {
        items.forEach(item => {
            if (stats[item.category] !== undefined) {
                stats[item.category] += item.balance;
            }
        });
    });
    return stats;
}

function renderStatCard(category, count, container) {
    if (count === 0) return;
    const categoryColors = {
        'Frame': '#4f46e5',
        'Lens': '#10b981',
        'Contactlens': '#f59e0b',
        'Service': '#8b5cf6',
        'Accessories': '#ec4899',
        'น้ำยา': '#06b6d4',
        'ไม่พบข้อมูลสินค้า': '#64748b'
    };
    const color = categoryColors[category] || '#64748b';
    const card = document.createElement('div');
    card.className = 'popup-stat-card';
    card.innerHTML = `
        <div class="stat-category-info">
            <span class="stat-color-dot" style="background-color: ${color}"></span>
            <span class="stat-cat-name">${category}</span>
        </div>
        <span class="stat-value">${count.toLocaleString()} ชิ้น</span>
    `;
    container.appendChild(card);
}

function showUploadSuccessPopup(fileName, fileSize, patternType, dataGroup) {
    const modal = document.getElementById('uploadSuccessModal');
    const nameSpan = document.getElementById('popupFileName');
    const sizeSpan = document.getElementById('popupFileSize');
    const patternSpan = document.getElementById('popupPattern');
    const deptSpan = document.getElementById('popupDeptCount');
    const statsList = document.getElementById('popupStatsList');
    const closeBtn = document.getElementById('closePopupBtn');

    nameSpan.textContent = fileName;
    sizeSpan.textContent = formatFileBytes(fileSize);
    patternSpan.textContent = translatePatternType(patternType);
    deptSpan.textContent = Object.keys(dataGroup).length.toString();

    statsList.innerHTML = '';
    const stats = calculateCategoryStatistics(dataGroup);
    Object.entries(stats).forEach(([category, count]) => {
        renderStatCard(category, count, statsList);
    });

    modal.style.display = 'flex';
    closeBtn.onclick = () => {
        modal.style.display = 'none';
        const filterElement = document.getElementById('filterSection');
        if (filterElement) {
            filterElement.scrollIntoView({ behavior: 'smooth' });
        }
    };
}

function renderUploadZoneActive(fileName) {
    dropZone.classList.add('active-file');
    dropZone.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none"
            stroke="#10b981" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"
            style="margin-bottom: 1rem;">
            <path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path>
            <polyline points="14 2 14 8 20 8"></polyline>
        </svg>
        <p style="font-weight: 600; color: #10b981; margin-bottom: 0.5rem;">กำลังใช้งานไฟล์:</p>
        <p id="activeFileName" style="font-weight: 700; color: var(--text-main); font-size: 1.1rem; margin-bottom: 1rem;" class="text-truncate">${fileName}</p>
        <button type="button" id="changeFileBtn" class="btn-secondary" style="display: inline-block; width: auto; padding: 0.5rem 1.5rem;">เปลี่ยนไฟล์</button>
    `;
    const btn = document.getElementById('changeFileBtn');
    btn.onclick = (e) => {
        e.stopPropagation();
        fileInput.click();
    };
}

function renderUploadZoneInactive() {
    dropZone.classList.remove('active-file');
    dropZone.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none"
            stroke="var(--primary)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"
            style="margin-bottom: 1rem;">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="17 8 12 3 7 8"></polyline>
            <line x1="12" y1="3" x2="12" y2="15"></line>
        </svg>
        <p style="font-weight: 600;">ลากไฟล์ .xls หรือ .xlsx มาวางที่นี่</p>
        <p style="color: var(--text-muted); font-size: 0.85rem; margin-top: 0.5rem;">
            หรือคลิกเพื่อเลือกไฟล์จากคอมพิวเตอร์ของคุณ</p>
    `;
}

function renderUploadZoneLoading() {
    dropZone.style.opacity = "0.6";
    dropZone.style.pointerEvents = "none";
    dropZone.innerHTML = `
        <p style="font-weight: 600;">⌛ กำลังประมวลผลไฟล์... (Processing)</p>
    `;
}

function restoreUploadZoneEvents() {
    dropZone.style.opacity = "1";
    dropZone.style.pointerEvents = "all";
}

function parseAndProcessExcel(file, arrayBuffer) {
    try {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const result = detectPattern(json);

        if (result && result.type) {
            processRawData(json, result, file.name, file.size);
            renderUploadZoneActive(file.name);
            return;
        }
        showFormatError(result ? result.missingColumns : []);
        renderUploadZoneInactive();
    } catch (err) {
        showError();
        renderUploadZoneInactive();
    }
}

async function handleFile(file) {
    if (!file) return;

    resetState();
    renderUploadZoneLoading();

    const reader = new FileReader();
    reader.onload = (e) => {
        parseAndProcessExcel(file, e.target.result);
        restoreUploadZoneEvents();
    };
    reader.onerror = () => {
        showError();
        restoreUploadZoneEvents();
        renderUploadZoneInactive();
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
    rawStockRows = null;
    rawStockResult = null;
    rawStockFileName = '';
    rawStockFileSize = 0;
    activePreviewTab = 'All';
    previewTabBar.innerHTML = '';
    previewTabBar.style.display = 'none';
    errorBanner.style.display = 'none';
    filterSection.style.display = 'none';
    previewSection.style.display = 'none';
    previewTableBody.innerHTML = '';
    renderUploadZoneInactive();
}

function processLegacyRow(row, currentDeptState) {
    if (!row || row.length === 0) return;
    const firstCell = String(row[0] || '').trim();

    if (firstCell.startsWith('Dept Name:')) {
        currentDeptState.name = firstCell.replace(/Dept\s*Name:\s*/gi, '').trim();
        return;
    }

    if (!currentDeptState.name || !row[3] || !row[0]) return;

    const rawCat = String(row[0] || '').trim();
    const cat = CATEGORY_MAP[rawCat] || rawCat;
    const balance = parseFloat(row[15]) || 0;
    if (balance === 0) return;

    const code = String(row[3]).trim();
    const finalDeptName = getMappedName(cat, currentDeptState.name, code);

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

function processLegacyData(rows) {
    const currentDeptState = { name: null };
    rows.forEach(row => {
        processLegacyRow(row, currentDeptState);
    });
}

function processFlatTableRow(row, mapping) {
    if (!row || !row[mapping.code]) return;

    const rawCat = mapping.groupName !== undefined ? String(row[mapping.groupName] || '').trim() : '';
    const cat = CATEGORY_MAP[rawCat] || rawCat;
    const balance = parseFloat(row[mapping.balance]) || 0;
    if (balance === 0) return;

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

function processFlatTableData(rows, result) {
    const { headerRowIndex, mapping } = result;
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        processFlatTableRow(rows[i], mapping);
    }
}

function processRawData(rows, result, fileName, fileSize) {
    rawStockRows = rows;
    rawStockResult = result;
    rawStockFileName = fileName;
    rawStockFileSize = fileSize;
    rebuildGroupedData();
    showUploadSuccessPopup(fileName, fileSize, rawStockResult.type, groupedData);
}

function rebuildGroupedData() {
    if (!rawStockRows || !rawStockResult) return;
    groupedData = {};
    const pattern = rawStockResult.type;
    if (pattern === 'LEGACY_DEPT_MARKER') {
        processLegacyData(rawStockRows);
    } else if (pattern === 'FLAT_TABLE') {
        processFlatTableData(rawStockRows, rawStockResult);
    }
    renderFilters();
    updatePreview();
}

function renderFilters() {
    groupList.innerHTML = '';
    const activeFilterCategories = [...TARGET_CATEGORIES];
    const hasUnmatched = Object.values(groupedData).some(items =>
        items.some(item => item.category === 'ไม่พบข้อมูลสินค้า')
    );
    if (hasUnmatched) activeFilterCategories.push('ไม่พบข้อมูลสินค้า');

    activeFilterCategories.forEach(cat => {
        if (cat === 'Service') return; // Hide Service option
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
    const categoriesWithData = getCategoriesWithData();
    const activeCategories = Array.from(categoriesWithData);

    if (activePreviewTab !== 'All' && !categoriesWithData.has(activePreviewTab)) {
        activePreviewTab = 'All';
    }

    renderPreviewTabBar(activeCategories);
    const rowCount = renderPreviewRows();
    previewSection.style.display = rowCount > 0 ? 'block' : 'none';
}

function getCategoriesWithData() {
    const categories = new Set();
    Object.values(groupedData).forEach(items => {
        items.forEach(item => {
            if (selectedCategories.has(item.category)) {
                categories.add(item.category);
            }
        });
    });
    return categories;
}

function renderPreviewTabBar(activeCategories) {
    if (activeCategories.length <= 1) {
        previewTabBar.style.display = 'none';
        activePreviewTab = 'All';
        return;
    }
    previewTabBar.innerHTML = '';
    previewTabBar.style.display = 'flex';
    renderAllTabButton();
    activeCategories.forEach(renderCategoryTabButton);
}

function renderAllTabButton() {
    const totalItemsCount = Object.values(groupedData).reduce((sum, items) => {
        return sum + items
            .filter(item => selectedCategories.has(item.category))
            .reduce((itemSum, item) => itemSum + item.balance, 0);
    }, 0);
    const allTab = document.createElement('div');
    allTab.className = `preview-tab ${activePreviewTab === 'All' ? 'active' : ''}`;
    allTab.innerHTML = `ทั้งหมด <span class="count-badge">${totalItemsCount}</span>`;
    allTab.onclick = () => {
        activePreviewTab = 'All';
        updatePreview();
    };
    previewTabBar.appendChild(allTab);
}

function renderCategoryTabButton(cat) {
    const catItemsCount = Object.values(groupedData).reduce((sum, items) => {
        return sum + items
            .filter(item => item.category === cat)
            .reduce((itemSum, item) => itemSum + item.balance, 0);
    }, 0);
    const tab = document.createElement('div');
    tab.className = `preview-tab ${activePreviewTab === cat ? 'active' : ''}`;
    tab.innerHTML = `${cat} <span class="count-badge">${catItemsCount}</span>`;
    tab.onclick = () => {
        activePreviewTab = cat;
        updatePreview();
    };
    previewTabBar.appendChild(tab);
}

function renderPreviewRows() {
    let rowCount = 0;
    Object.keys(groupedData).forEach(dept => {
        const filteredItems = groupedData[dept].filter(isItemInActiveTab);
        if (filteredItems.length === 0) return;
        renderDeptHeaderRow(dept, filteredItems);
        filteredItems.forEach(item => {
            rowCount++;
            if (rowCount > 100) return;
            renderItemRow(item);
        });
    });
    return rowCount;
}

function isItemInActiveTab(item) {
    if (activePreviewTab === 'All') {
        return selectedCategories.has(item.category);
    }
    return item.category === activePreviewTab;
}

function renderDeptHeaderRow(dept, filteredItems) {
    const systemSubtotal = filteredItems.reduce((sum, item) => sum + item.balance, 0);
    const actualSubtotal = filteredItems.reduce((sum, item) => sum + (item.actualCount || 0), 0);
    const hasScanData = filteredItems.some(item => item.actualCount !== undefined && item.actualCount !== null);
    const actualText = hasScanData ? actualSubtotal.toLocaleString() : '';
    const varianceText = hasScanData ? (actualSubtotal - systemSubtotal).toLocaleString() : '';
    const headerTr = document.createElement('tr');
    headerTr.className = 'dept-row';
    headerTr.innerHTML = `
        <td colspan="5">Dept Name: ${dept}</td>
        <td>${systemSubtotal.toLocaleString()}</td>
        <td>${actualText}</td>
        <td>${varianceText}</td>
        <td></td>
    `;
    previewTableBody.appendChild(headerTr);
}

function renderItemRow(item) {
    const hasActual = item.actualCount !== undefined && item.actualCount !== null;
    const actualText = hasActual ? item.actualCount.toLocaleString() : '';
    const varianceText = hasActual ? (item.actualCount - item.balance).toLocaleString() : '';
    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${item.category}</td>
        <td>${item.type}</td>
        <td>${item.dept}</td>
        <td>${item.code}</td>
        <td>${item.description}</td>
        <td>${item.balance.toLocaleString()}</td>
        <td>${actualText}</td>
        <td>${varianceText}</td>
        <td></td>
    `;
    previewTableBody.appendChild(tr);
}

async function exportToExcel(branchCode) {
    const workbook = new ExcelJS.Workbook();
    createInstructionSheet(workbook);
    createScanSheet(workbook);

    const activeCategories = Array.from(selectedCategories).filter(cat => {
        return Object.values(groupedData).some(items => items.some(item => item.category === cat));
    });

    activeCategories.forEach(cat => {
        createCategorySheet(workbook, cat, branchCode);
    });

    const filename = getExportFilename(branchCode);
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), filename);
}

function createScanSheet(workbook) {
    const scanSheet = workbook.addWorksheet('สแกน', { properties: { tabColor: { argb: '#5c5c5c' } } });
    scanSheet.columns = [
        { header: 'รายการสแกน (Barcode)', key: 'barcode', width: 25 },
        { header: 'สถานะ / รายละเอียดสินค้า', key: 'status', width: 45 }
    ];
    const headerRow = scanSheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
    headerRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
    headerRow.getCell(1).border = { bottom: { style: 'thin' } };
    headerRow.getCell(2).border = { bottom: { style: 'thin' } };

    for (let r = 2; r <= 5000; r++) {
        const row = scanSheet.getRow(r);
        row.getCell(2).value = { formula: `IF(A${r}="","",_xlfn.XLOOKUP(A${r},'Frame'!D:D,'Frame'!E:E,"ไม่พบข้อมูล"))` };
    }
}

function createInstructionSheet(workbook) {
    const insSheet = workbook.addWorksheet('คู่มือการใช้งาน', { properties: { tabColor: { argb: 'FFFF0000' } } });
    insSheet.columns = [{ width: 5 }, { width: 80 }];
    const insTitle = insSheet.addRow(["", "คู่มือการใช้งานไฟล์ Audit Stock"]);
    insTitle.font = { size: 18, bold: true, color: { argb: 'FF4F46E5' } };
    insSheet.addRow([]);

    const instructions = [
        "1. ตรวจสอบข้อมูลแผนกและหมวดหมู่สินค้าในแต่ละแท็บ (เช่น Frame, Lens, Contactlens)",
        "2. สำหรับสินค้าหมวดแว่น (Frame): ให้สลับไปที่แท็บ 'สแกน' แล้วใช้เครื่องสแกนยิงบาร์โค้ดลงในคอลัมน์ A (รายการสแกน) ได้ทันที",
        "3. ในแท็บ 'สแกน' คอลัมน์ B จะแสดงชื่อสินค้าโดยอัตโนมัติ หากบาร์โค้ดนั้นไม่มีอยู่ในสต็อกจะขึ้นแจ้งเตือนว่า 'ไม่พบข้อมูล'",
        "4. จำนวนที่ยิงสแกนได้ในแท็บ 'สแกน' จะถูกนำไปคำนวณและป้อนในคอลัมน์ 'Actual Count' ของแท็บ 'Frame' ให้โดยอัตโนมัติ",
        "5. สำหรับสินค้าหมวดอื่นๆ (Lens, Contactlens, Accessories, น้ำยา): ให้ใช้วิธีนับมือแล้วกรอกลงในคอลัมน์ 'Actual Count' (ช่องสีขาว)",
        "6. ระบบจะคำนวณผลต่าง (Variance) และยอดรวมแยกตามแผนก/กลุ่มสินค้าให้โดยอัตโนมัติ",
        "7. การจัดการรูปแบบ Group: สามารถใช้เครื่องหมาย (+) และ (-) ทางด้านซ้ายมือเพื่อย่อหรือขยายรายละเอียดในแต่ละแผนกได้"
    ];
    instructions.forEach((text, i) => {
        const row = insSheet.addRow(["", text]);
        row.font = { size: 12 };
        if (i === 1 || i === 2 || i === 3) {
            row.getCell(2).font = { size: 12, bold: true, color: { argb: 'FF4F46E5' } };
        }
        insSheet.addRow([]);
    });
}

function getCategoryTabColor(cat) {
    const tabColors = {
        'Frame': 'FF4F46E5',
        'Lens': 'FF10B981',
        'Contactlens': 'FFF59E0B',
        'Service': 'FF8B5CF6',
        'Accessories': 'FFEC4899',
        'น้ำยา': 'FF06B6D4',
        'ไม่พบข้อมูลสินค้า': 'FF64748B'
    };
    return tabColors[cat] || 'FF4F46E5';
}

function createCategorySheet(workbook, cat, branchCode) {
    const tabColor = getCategoryTabColor(cat);
    const worksheet = workbook.addWorksheet(cat, {
        views: [{ state: 'frozen', ySplit: 2 }],
        properties: {
            tabColor: { argb: tabColor },
            outlineLevelCol: 0,
            outlineLevelRow: 1,
            outlineProperties: { summaryBelow: false }
        }
    });
    setupSheetHeaders(worksheet, cat, branchCode, tabColor);
    populateSheetData(worksheet, cat);
    worksheet.columns = [
        { width: 15 }, { width: 10 }, { width: 10 }, { width: 25 }, { width: 45 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 25 }
    ];
}

function setupSheetHeaders(worksheet, cat, branchCode, tabColor) {
    const titleText = branchCode ? `Audit Stock Report (${cat}) สาขา ${branchCode}` : `Audit Stock Report (${cat})`;
    const titleRow = worksheet.addRow([titleText]);
    worksheet.mergeCells('A1:I1');
    titleRow.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' };
    titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tabColor } };

    const headerRow = worksheet.addRow(["Category", "Type", "Dept", "Code", "Description", "System Stock", "Actual Count", "Variance", "Remark"]);
    headerRow.font = { bold: true };
    headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
        cell.border = { bottom: { style: 'thin' } };
    });
}

function populateSheetData(worksheet, cat) {
    Object.keys(groupedData).forEach(dept => {
        const filteredItems = groupedData[dept].filter(item => item.category === cat);
        if (filteredItems.length === 0) return;

        const deptRow = worksheet.addRow([`Dept Name: ${dept}`, "", "", "", "", 0, 0, 0, ""]);
        deptRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        deptRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF404040' } };
        });

        const startRow = worksheet.rowCount + 1;
        filteredItems.forEach(item => {
            const row = worksheet.addRow([
                item.category, item.type, item.dept, item.code, item.description,
                item.balance, null, null, null
            ]);
            row.outlineLevel = 1;
            if (cat === 'Frame') {
                row.getCell(7).value = { formula: `COUNTIF('สแกน'!A:A, D${row.number})` };
            } else {
                row.getCell(7).value = item.actualCount !== undefined && item.actualCount !== null ? item.actualCount : null;
            }
            row.getCell(8).value = { formula: `G${row.number}-F${row.number}` };
        });

        addBlankRowsForDept(worksheet);
        const endRow = worksheet.rowCount;
        deptRow.getCell(6).value = { formula: `SUM(F${startRow}:F${endRow})` };
        deptRow.getCell(7).value = { formula: `SUM(G${startRow}:G${endRow})` };
        deptRow.getCell(8).value = { formula: `G${deptRow.number}-F${deptRow.number}` };
    });
}

function addBlankRowsForDept(worksheet) {
    for (let j = 0; j < 2; j++) {
        const blankRow = worksheet.addRow(["", "", "", "", "-", 0, null, null, null]);
        blankRow.outlineLevel = 1;
        blankRow.getCell(8).value = { formula: `G${blankRow.number}-F${blankRow.number}` };

        blankRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
        });
    }
}

function getExportFilename(branchCode) {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const dateStr = `${yyyy}${mm}${dd}`;
    const cleanBranch = branchCode ? String(branchCode).trim() : '';
    return cleanBranch
        ? `Audit_Stock_${cleanBranch}_${dateStr}.xlsx`
        : `Audit_Stock_${dateStr}.xlsx`;
}

const branchCodeModal = document.getElementById('branchCodeModal');
const branchInput = document.getElementById('branchInput');
const branchInputError = document.getElementById('branchInputError');
const cancelBranchBtn = document.getElementById('cancelBranchBtn');
const submitBranchBtn = document.getElementById('submitBranchBtn');

downloadBtn.onclick = () => {
    if (branchCodeModal && branchInput) {
        branchInput.value = '';
        branchInputError.style.display = 'none';
        submitBranchBtn.disabled = true;
        branchCodeModal.style.display = 'flex';
        setTimeout(() => branchInput.focus(), 100);
    }
};

if (branchCodeModal) {
    cancelBranchBtn.onclick = () => {
        branchCodeModal.style.display = 'none';
    };

    branchInput.addEventListener('input', (e) => {
        let value = e.target.value;
        const cleaned = value.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

        if (value !== cleaned) {
            e.target.value = cleaned;
            value = cleaned;
            branchInputError.textContent = "⚠️ กรุณากรอกเฉพาะภาษาอังกฤษและตัวเลขเท่านั้น";
            branchInputError.style.display = 'block';
            return;
        }

        const pattern = /^[BMPQ]\d{4}$/;
        if (value.length === 0) {
            branchInputError.style.display = 'none';
            submitBranchBtn.disabled = true;
        } else if (!/^[BMPQ]/i.test(value)) {
            branchInputError.textContent = "❌ ต้องขึ้นต้นด้วยตัวอักษร B, M, P หรือ Q เท่านั้น";
            branchInputError.style.display = 'block';
            submitBranchBtn.disabled = true;
        } else if (value.length < 5) {
            branchInputError.textContent = "❌ ต้องระบุตัวเลขตามหลังอีก 4 หลัก (เช่น B0001)";
            branchInputError.style.display = 'block';
            submitBranchBtn.disabled = true;
        } else if (!pattern.test(value)) {
            branchInputError.textContent = "❌ รูปแบบไม่ถูกต้อง ต้องเป็น Bxxxx, Mxxxx, Pxxxx หรือ Qxxxx (เช่น B0001, P5001)";
            branchInputError.style.display = 'block';
            submitBranchBtn.disabled = true;
        } else {
            branchInputError.style.display = 'none';
            submitBranchBtn.disabled = false;
        }
    });

    branchInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !submitBranchBtn.disabled) {
            submitBranchBtn.click();
        }
    });

    submitBranchBtn.onclick = () => {
        const branchCode = branchInput.value.trim().toUpperCase();
        branchCodeModal.style.display = 'none';
        exportToExcel(branchCode);
    };
}

// Scroll to Bottom Logic
const scrollToBottomBtn = document.getElementById('scrollToBottomBtn');
if (scrollToBottomBtn) {
    window.addEventListener('scroll', () => {
        const threshold = 100; // Pixels from bottom
        const totalHeight = document.documentElement.scrollHeight;
        const viewportHeight = window.innerHeight;
        const currentScroll = window.scrollY || window.pageYOffset;

        // Show if page is scrollable and we are not near the bottom
        if (totalHeight - viewportHeight - currentScroll > threshold) {
            scrollToBottomBtn.classList.add('visible');
        } else {
            scrollToBottomBtn.classList.remove('visible');
        }
    });

    scrollToBottomBtn.onclick = () => {
        window.scrollTo({
            top: document.body.scrollHeight,
            behavior: 'smooth'
        });
    };
}
