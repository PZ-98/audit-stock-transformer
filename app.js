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

async function handleFile(file) {
    if (!file) return;
    resetState();
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const pattern = detectPattern(json);
            if (pattern) {
                processRawData(json, pattern);
            } else {
                showError();
            }
        } catch (err) {
            console.error(err);
            showError();
        }
    };
    reader.readAsArrayBuffer(file);
}

function detectPattern(rows) {
    // Pattern 1: Legacy Dept Marker
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
        const firstCell = String(rows[i][0] || '').trim();
        if (firstCell.startsWith('Dept Name:')) return 'LEGACY_DEPT_MARKER';
    }
    
    // Pattern 2: Flat Table (Ex2.xls)
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const row = rows[i];
        if (row && row[0] === 'Group Code' && row[1] === 'Group Name' && row[4] === 'Mat Code') {
            return 'FLAT_TABLE';
        }
    }
    return null;
}

function showError() {
    errorBanner.style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function resetState() {
    groupedData = {};
    errorBanner.style.display = 'none';
    filterSection.style.display = 'none';
    previewSection.style.display = 'none';
    previewTableBody.innerHTML = '';
}

function processRawData(rows, pattern) {
    groupedData = {};

    if (pattern === 'LEGACY_DEPT_MARKER') {
        let currentDept = null;
        rows.forEach(row => {
            if (!row || row.length === 0) return;
            const firstCell = String(row[0] || '').trim();
            
            if (firstCell.startsWith('Dept Name:')) {
                currentDept = firstCell.replace('Dept Name:', '').trim();
                if (!groupedData[currentDept]) groupedData[currentDept] = [];
            } 
            else if (currentDept && row[3] && row[0]) {
                let rawCat = String(row[0] || '').trim();
                let cat = CATEGORY_MAP[rawCat] || rawCat;
                let balance = parseFloat(row[15]) || 0;
                
                // Skip items with 0 balance
                if (balance === 0) return;

                groupedData[currentDept].push({
                    category: cat,
                    type: row[1],
                    dept: row[2],
                    code: row[3],
                    description: row[4],
                    balance: balance
                });
            }
        });
    } else if (pattern === 'FLAT_TABLE') {
        let dataStarted = false;
        rows.forEach(row => {
            if (!row || row.length < 15) return;
            
            // Detect header row
            if (row[0] === 'Group Code' && row[4] === 'Mat Code') {
                dataStarted = true;
                return;
            }
            
            if (!dataStarted) return;
            if (!row[4]) return; // Skip if no Mat Code

            const deptName = String(row[2] || 'General').trim();
            if (!groupedData[deptName]) groupedData[deptName] = [];

            let rawCat = String(row[1] || '').trim();
            let cat = CATEGORY_MAP[rawCat] || rawCat;
            let balance = parseFloat(row[14]) || 0;
            
            // Skip items with 0 balance
            if (balance === 0) return;

            groupedData[deptName].push({
                category: cat,
                type: row[0], // Group Code
                dept: row[2], // Dept Name
                code: row[4], // Mat Code
                description: row[5], // Mat Name
                balance: balance
            });
        });
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
    
    Object.keys(groupedData).forEach(dept => {
        const filteredItems = groupedData[dept].filter(item => selectedCategories.has(item.category));
        if (filteredItems.length === 0) return;

        const subtotal = filteredItems.reduce((sum, item) => sum + item.balance, 0);

        const headerTr = document.createElement('tr');
        headerTr.className = 'dept-row';
        headerTr.innerHTML = `
            <td colspan="5">Dept Name: ${dept}</td>
            <td>${subtotal.toLocaleString()}</td>
            <td></td>
            <td></td>
        `;
        previewTableBody.appendChild(headerTr);

        filteredItems.forEach(item => {
            if (rowCount > 100) return;
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
            `;
            previewTableBody.appendChild(tr);
            rowCount++;
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
        "1. ตรวจสอบข้อมูลแผนกและหมวดหมู่สินค้าในหน้า 'Audit Stock'",
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

    // 2. Audit Sheet (Second Tab)
    const worksheet = workbook.addWorksheet('Audit Stock', {
        views: [{ state: 'frozen', ySplit: 2 }],
        properties: { outlineLevelCol: 0, outlineLevelRow: 1 }
    });

    // Main Header with Branch Code
    const titleText = branchCode ? `Audit Stock Report สาขา ${branchCode}` : 'Audit Stock Report';
    const titleRow = worksheet.addRow([titleText]);
    worksheet.mergeCells('A1:H1');
    titleRow.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' };
    titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; 

    // Table Headers
    const headerRow = worksheet.addRow(["Category", "Type", "Dept", "Code", "Description", "System Stock", "Actual Count", "Variance"]);
    headerRow.font = { bold: true };
    headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
        cell.border = { bottom: { style: 'thin' } };
    });

    Object.keys(groupedData).forEach(dept => {
        const filteredItems = groupedData[dept].filter(item => selectedCategories.has(item.category));
        if (filteredItems.length === 0) return;

        // Add Dept Name Header Row (Styled with #404040 background and white text)
        const deptRow = worksheet.addRow([`Dept Name: ${dept}`, "", "", "", "", 0, 0, 0]);
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
                null 
            ]);
            row.outlineLevel = 1;
            const rowIndex = row.number;
            row.getCell(8).value = { formula: `G${rowIndex}-F${rowIndex}` };
        });
        const endRow = worksheet.rowCount;

        // Set dynamic SUM formulas
        deptRow.getCell(6).value = { formula: `SUM(F${startRow}:F${endRow})` };
        deptRow.getCell(7).value = { formula: `SUM(G${startRow}:G${endRow})` };
        deptRow.getCell(8).value = { formula: `SUM(H${startRow}:H${endRow})` };
    });

    // Column Widths
    worksheet.columns = [
        { width: 15 }, { width: 10 }, { width: 10 }, { width: 25 }, { width: 45 }, { width: 15 }, { width: 15 }, { width: 15 }
    ];

    // Filename logic
    const filename = branchCode ? `Audit_Stock_${branchCode}.xlsx` : 'Audit_Stock_Advanced.xlsx';

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), filename);
}

downloadBtn.onclick = () => {
    const branchCode = prompt("กรุณากรอกรหัสสาขาเพื่อระบุในรายงานและชื่อไฟล์:");
    exportToExcel(branchCode);
};
