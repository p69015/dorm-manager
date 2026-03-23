// Initial state for 5 dormitories
const INITIAL_DORMS = [
    { id: 1, name: '女生第一宿舍 (晨曦樓)', rooms: {
        '2': { count: 20, price: 3000 },
        '4': { count: 50, price: 2000 },
        '6': { count: 30, price: 1500 }
    }},
    { id: 2, name: '女生第二宿舍 (星辰樓)', rooms: {
        '2': { count: 10, price: 3200 },
        '4': { count: 60, price: 2100 },
        '6': { count: 40, price: 1600 }
    }},
    { id: 3, name: '德惠宿舍 (翠林樓)', rooms: {
        '2': { count: 15, price: 3100 },
        '4': { count: 45, price: 2000 },
        '6': { count: 20, price: 1500 }
    }},
    { id: 4, name: '新德惠宿舍 (雅風樓)', rooms: {
        '2': { count: 5, price: 3500 },
        '4': { count: 80, price: 2200 },
        '6': { count: 10, price: 1800 }
    }},
    { id: 5, name: '大直宿舍 (弘道樓)', rooms: {
        '2': { count: 30, price: 2800 },
        '4': { count: 40, price: 1900 },
        '6': { count: 50, price: 1400 }
    }}
];

// App State
let dormsData = [];
// State for residents: { [dormIndex]: { [bedId]: { studentData } } }
let residentsData = {};
// State for beds from Excel: [[bed1, bed2...], [bed1...], ...] index mapped to dorm
let allBedsData = [[], [], [], [], []];

let activeModalDormIndex = null;
let activeModalBedId = null;

// DOM Elements
const dormGrid = document.getElementById('dorm-grid');
const totalCapacityEl = document.getElementById('total-capacity');
const totalRevenueEl = document.getElementById('total-revenue');
const totalOccupantsEl = document.getElementById('total-occupants');
const totalEmptyEl = document.getElementById('total-empty');
const totalPaidEl = document.getElementById('total-paid');
const totalActualRevenueEl = document.getElementById('total-actual-revenue');
const statsTableBody = document.getElementById('stats-table-body');
const modalEl = document.getElementById('resident-modal');

// Initialize App
async function init() {
    // 1. Load Data from localStorage
    const storedDorms = localStorage.getItem('dormData');
    if (storedDorms) dormsData = JSON.parse(storedDorms);
    else dormsData = JSON.parse(JSON.stringify(INITIAL_DORMS));
    
    const storedResidents = localStorage.getItem('residentsData');
    if (storedResidents) residentsData = JSON.parse(storedResidents);

    // 2. Load Bed Numbers from data.js
    try {
        if (typeof RAW_BED_DATA !== 'undefined') {
            const sheet = RAW_BED_DATA['工作表1'];
            if (sheet && sheet.length > 1) {
                const numCols = sheet[0].length;
                for (let row = 1; row < sheet.length; row++) {
                    for (let col = 0; col < numCols; col++) {
                        const bedVal = sheet[row][col];
                        if (bedVal && typeof bedVal === 'string' && bedVal.trim() !== '') {
                            if (col < 5) {
                                allBedsData[col].push(bedVal.trim());
                            }
                        }
                    }
                }
            }
        } else {
            console.error("RAW_BED_DATA is not defined. Ensure data.js is loaded.");
        }
    } catch (e) {
        console.error("Failed to load bed data from RAW_BED_DATA:", e);
    }

    // 3. Process Imported Residents if not already processed
    if (typeof IMPORTED_RESIDENTS !== 'undefined' && !localStorage.getItem('importedResidentDataProcessed_v2')) {
        IMPORTED_RESIDENTS.forEach(item => {
            const row = item.value || item;
            const id = row[0];
            const name = row[1];
            
            // Skip headers or empty rows
            if (!id || id === '學號') return;

            let paymentStr = row[2] || "";
            let payment = "未繳費";
            if (paymentStr.includes("銷帳")) payment = "已繳費";
            else if (paymentStr.includes("貸") || (row[3] && row[3].includes("貸"))) payment = "分期/就貸";
            
            let nationality = row[5] || "台灣";
            let dormName = row[8] || "";
            let bedId = row[9] || "";
            let className = row[10] || "";
            let phone = row[14] || "";
            
            let targetIndex = -1;
            if (dormName.includes("一")) targetIndex = 0;
            else if (dormName.includes("二")) targetIndex = 1;
            else if (dormName.includes("新德惠")) targetIndex = 3;
            else if (dormName.includes("德惠")) targetIndex = 2;
            else if (dormName.includes("大直")) targetIndex = 4;
            
            if (targetIndex !== -1 && bedId) {
                if (!residentsData[targetIndex]) residentsData[targetIndex] = {};
                
                // Extract year and dept from class (e.g., 日工2甲 -> dept '工', year '2')
                let dept = className.length >= 3 ? className.substring(1, 2) + "系" : "未分類";
                let yearMatch = className.match(/\d/);
                let year = yearMatch ? yearMatch[0] + "年級" : "一年級";

                residentsData[targetIndex][bedId] = {
                    id: id, 
                    name: name, 
                    phone: phone, 
                    dept: dept, 
                    nationality: nationality, 
                    year: year, 
                    className: className, 
                    checkIn: new Date().toISOString().split('T')[0], 
                    payment: payment
                };
            }
        });
        localStorage.setItem('residentsData', JSON.stringify(residentsData));
        localStorage.setItem('importedResidentDataProcessed_v2', 'true');
    }

    render();
}

// Stats Calculators
function getDormCapacity(dorm) {
    return Object.entries(dorm.rooms).reduce((total, [cap, data]) => {
        return total + (parseInt(cap) * data.count);
    }, 0);
}

function getDormRevenue(dorm) {
    return Object.entries(dorm.rooms).reduce((total, [cap, data]) => {
        return total + (parseInt(cap) * data.count * data.price);
    }, 0);
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('zh-TW', { style: 'currency', currency: 'TWD', maximumFractionDigits: 0 }).format(amount);
}

// ----- Edit Mode Logic -----
function toggleEditMode(index, isEdit) {
    const card = document.querySelector(`.dorm-card[data-index="${index}"]`);
    const viewMode = card.querySelector('.view-mode');
    const editMode = card.querySelector('.edit-mode');
    const editBtn = card.querySelector('.edit-btn');
    
    if (isEdit) {
        viewMode.classList.add('hidden');
        editMode.classList.remove('hidden');
        editBtn.classList.add('hidden');
    } else {
        viewMode.classList.remove('hidden');
        editMode.classList.add('hidden');
        editBtn.classList.remove('hidden');
    }
}

function syncDraftState(index) {
    const card = document.querySelector(`.dorm-card[data-index="${index}"]`);
    if (!card) return;
    
    const newName = card.querySelector('.input-name').value;
    const roomRows = card.querySelectorAll('.room-type-row');
    
    const newRooms = {};
    roomRows.forEach(row => {
        const cap = row.dataset.capacity;
        const count = parseInt(row.querySelector('.input-count').value) || 0;
        const price = parseInt(row.querySelector('.input-price').value) || 0;
        newRooms[cap] = { count, price };
    });
    
    dormsData[index].name = newName;
    dormsData[index].rooms = newRooms;
}

function saveDorm(index) {
    syncDraftState(index);
    localStorage.setItem('dormData', JSON.stringify(dormsData));
    render();
}

function addRoomType(index) {
    const capacity = prompt('請輸入新房型的容納人數 (例如輸入 3 代表三人房):');
    if (!capacity || isNaN(capacity) || parseInt(capacity) <= 0) return;
    
    const cap = parseInt(capacity).toString();
    syncDraftState(index);
    
    if (!dormsData[index].rooms[cap]) {
        dormsData[index].rooms[cap] = { count: 0, price: 0 };
    } else {
        alert('該房型已經存在！');
    }
    
    render();
    toggleEditMode(index, true);
}

function removeRoomType(index, capacity) {
    if (confirm(`確定要刪除 ${capacity}人房 的設定嗎？`)) {
        syncDraftState(index);
        delete dormsData[index].rooms[capacity];
        render();
        toggleEditMode(index, true);
    }
}

// ----- Resident Modal Logic -----
function openResidentModal(index) {
    const dorm = dormsData[index];
    if (!residentsData[index]) residentsData[index] = {};

    activeModalDormIndex = index;
    activeModalBedId = null;

    document.getElementById('modal-dorm-name').textContent = `${dorm.name} - 床位與住宿生管理`;
    modalEl.classList.remove('hidden');
    
    // Default Empty Panel
    const panel = document.getElementById('bed-detail-panel');
    panel.innerHTML = `
        <div class="empty-state">
            <div>
                <i class="fa-solid fa-bed" style="font-size: 3rem; color: #E5E7EB; margin-bottom: 1rem;"></i>
                <p>請從左側列表選擇床位以檢視或新增住宿生資料</p>
            </div>
        </div>
    `;

    renderBedsSidebar();
}

// Expose modal close to window for onclick outside script block
window.closeResidentModal = function() {
    modalEl.classList.add('hidden');
    activeModalDormIndex = null;
    activeModalBedId = null;
};

// Also close on click outside the modal content
modalEl.addEventListener('click', (e) => {
    if (e.target === modalEl) closeResidentModal();
});

const infoModal = document.getElementById('info-modal');
infoModal.addEventListener('click', (e) => {
    if (e.target === infoModal) infoModal.classList.add('hidden');
});

const listModal = document.getElementById('list-modal');
if (listModal) {
    listModal.addEventListener('click', (e) => {
        if (e.target === listModal) listModal.classList.add('hidden');
    });
}

function renderBedsSidebar() {
    const listEl = document.getElementById('beds-list');
    listEl.innerHTML = '';
    
    const index = activeModalDormIndex;
    if (index === null) return;

    // Use fetched beds or fallback mock beds if fetch failed
    let beds = allBedsData[index] || [];
    if (beds.length === 0) {
        beds = ["無床位資料 (無法讀取 Excel)"];
    }

    beds.forEach(bedId => {
        const isOccupied = residentsData[index][bedId] !== undefined;
        const statusClass = isOccupied ? 'occupied' : 'empty';
        const statusText = isOccupied ? '已入住' : '空床';
        const activeClass = activeModalBedId === bedId ? 'active' : '';

        const item = document.createElement('div');
        item.className = `bed-item ${statusClass} ${activeClass}`;
        item.onclick = () => selectBed(index, bedId);
        
        let label = bedId;
        if (isOccupied && residentsData[index][bedId].name) {
            label = `${bedId} (${residentsData[index][bedId].name})`;
        }
        
        item.innerHTML = `
            <span class="bed-id" style="font-family: monospace; font-weight: 500;"><i class="fa-solid fa-bed text-muted"></i> ${label}</span>
            <span class="status">${statusText}</span>
        `;
        listEl.appendChild(item);
    });
}

function selectBed(index, bedId) {
    if (bedId.includes("無床位資料")) return;
    activeModalBedId = bedId;
    renderBedsSidebar(); // Update active highlights
    
    const panel = document.getElementById('bed-detail-panel');
    const student = residentsData[index][bedId];
    const today = new Date().toISOString().split('T')[0];

    if (student) {
        // Edit Existing Resident
        panel.innerHTML = `
            <div class="resident-form-title">
                <span><i class="fa-solid fa-bed text-primary"></i> 床位 ${bedId} 
                    <span class="badge" style="background:#EEF2FF; color:var(--primary); font-size:0.875rem; padding:0.2rem 0.6rem; border-radius:1rem; margin-left:0.5rem; vertical-align:middle;">已被申請</span>
                </span>
                <button class="btn btn-secondary" style="color:var(--danger); border:1px solid var(--danger); font-size:0.875rem; padding:0.4rem 0.8rem;" onclick="removeResident(${index}, '${bedId}')">
                    <i class="fa-solid fa-user-minus"></i> 申請退宿
                </button>
            </div>
            
            <div class="resident-form">
                <div class="form-group">
                    <label>學號 Student ID</label>
                    <input type="text" class="form-control" id="res-id" value="${student.id || ''}">
                </div>
                <div class="form-group">
                    <label>姓名 Name</label>
                    <input type="text" class="form-control" id="res-name" value="${student.name || ''}">
                </div>
                <div class="form-group">
                    <label>手機號碼 Phone</label>
                    <input type="text" class="form-control" id="res-phone" value="${student.phone || ''}">
                </div>
                <div class="form-group">
                    <label>科系 Department</label>
                    <input type="text" class="form-control" id="res-dept" value="${student.dept || ''}">
                </div>
                <div class="form-group">
                    <label>國籍 Nationality</label>
                    <input type="text" class="form-control" id="res-nat" value="${student.nationality || ''}">
                </div>
                <div class="form-group">
                    <label>年級 Year</label>
                    <input type="text" class="form-control" id="res-year" value="${student.year || ''}">
                </div>
                <div class="form-group">
                    <label>班級 Class</label>
                    <input type="text" class="form-control" id="res-class" value="${student.className || ''}">
                </div>
                <div class="form-group">
                    <label>入住日期 Check-in Date</label>
                    <input type="date" class="form-control" id="res-date" value="${student.checkIn || today}">
                </div>
                <div class="form-group">
                    <label>繳費情形 Payment Status</label>
                    <select class="form-control" id="res-payment">
                        <option value="已繳費" ${student.payment === '已繳費' ? 'selected' : ''}>已繳費 (Paid)</option>
                        <option value="未繳費" ${student.payment === '未繳費' ? 'selected' : ''}>未繳費 (Unpaid)</option>
                        <option value="分期/就貸" ${student.payment === '分期/就貸' ? 'selected' : ''}>分期/就學貸款 (Installment/Loan)</option>
                    </select>
                </div>
                <div class="form-group full-width" style="display:flex; justify-content:flex-end; gap:1rem; margin-top:2rem;">
                    <button class="btn btn-primary" onclick="saveResident(${index}, '${bedId}')"><i class="fa-solid fa-save"></i> 儲存變更 (Save Changes)</button>
                </div>
            </div>
        `;
    } else {
        // Add New Resident
        panel.innerHTML = `
            <div class="resident-form-title">
                <span><i class="fa-solid fa-bed text-muted"></i> 床位 ${bedId} 
                    <span class="badge" style="background:#F3F4F6; color:var(--text-muted); font-size:0.875rem; padding:0.2rem 0.6rem; border-radius:1rem; margin-left:0.5rem; vertical-align:middle;">空床</span>
                </span>
            </div>
            
            <div class="resident-form">
                <div class="form-group">
                    <label>學號 Student ID</label>
                    <input type="text" class="form-control" id="res-id" placeholder="例如：B1140001">
                </div>
                <div class="form-group">
                    <label>姓名 Name</label>
                    <input type="text" class="form-control" id="res-name" placeholder="例如：王小明">
                </div>
                <div class="form-group">
                    <label>手機號碼 Phone</label>
                    <input type="text" class="form-control" id="res-phone" placeholder="例如：0912-345-678">
                </div>
                <div class="form-group">
                    <label>科系 Department</label>
                    <input type="text" class="form-control" id="res-dept" placeholder="例如：資訊工程學系">
                </div>
                <div class="form-group">
                    <label>國籍 Nationality</label>
                    <input type="text" class="form-control" id="res-nat" value="台灣 (Taiwan)">
                </div>
                <div class="form-group">
                    <label>年級 Year</label>
                    <input type="text" class="form-control" id="res-year" placeholder="例如：大一">
                </div>
                <div class="form-group">
                    <label>班級 Class</label>
                    <input type="text" class="form-control" id="res-class" placeholder="例如：甲班">
                </div>
                <div class="form-group">
                    <label>入住日期 Check-in Date</label>
                    <input type="date" class="form-control" id="res-date" value="${today}">
                </div>
                <div class="form-group">
                    <label>繳費情形 Payment Status</label>
                    <select class="form-control" id="res-payment">
                        <option value="未繳費">未繳費 (Unpaid)</option>
                        <option value="已繳費">已繳費 (Paid)</option>
                        <option value="分期/就貸">分期/就學貸款 (Installment/Loan)</option>
                    </select>
                </div>
                <div class="form-group full-width" style="display:flex; justify-content:flex-end; gap:1rem; margin-top:2rem;">
                    <button class="btn btn-primary" onclick="saveResident(${index}, '${bedId}')"><i class="fa-solid fa-user-plus"></i> 新增住宿生 (Add Resident)</button>
                </div>
            </div>
        `;
    }
}

window.saveResident = function(index, bedId) {
    if (!residentsData[index]) residentsData[index] = {};
    
    const id = document.getElementById('res-id').value.trim();
    const name = document.getElementById('res-name').value.trim();
    
    if (!id || !name) {
        alert("學號與姓名為必填欄位！");
        return;
    }
    
    residentsData[index][bedId] = {
        id: id,
        name: name,
        phone: document.getElementById('res-phone').value.trim(),
        dept: document.getElementById('res-dept').value.trim(),
        nationality: document.getElementById('res-nat').value.trim(),
        year: document.getElementById('res-year').value.trim(),
        className: document.getElementById('res-class').value.trim(),
        checkIn: document.getElementById('res-date').value,
        payment: document.getElementById('res-payment').value
    };
    
    localStorage.setItem('residentsData', JSON.stringify(residentsData));
    
    // UI Updates
    selectBed(index, bedId);    // Refresh detail panel
    renderBedsSidebar();        // Refresh list decorators (colors/badges)
}

window.removeResident = function(index, bedId) {
    if (confirm(`確定要為床位 ${bedId} 辦理退宿嗎？所有住宿生資料將從系統中刪除。`)) {
        delete residentsData[index][bedId];
        localStorage.setItem('residentsData', JSON.stringify(residentsData));
        selectBed(index, bedId);
    }
}

// ----- Main Rendering -----
function render() {
    dormGrid.innerHTML = '';
    
    let totalGlobalCapacity = 0;
    let totalGlobalRevenue = 0;
    let totalGlobalActualRevenue = 0;
    let totalGlobalOccupants = 0;
    let totalGlobalPaid = 0;
    
    let statsTableHtml = '';
    
    dormsData.forEach((dorm, index) => {
        const capacity = getDormCapacity(dorm);
        const revenue = getDormRevenue(dorm);
        let actualRevenue = 0;
        
        totalGlobalCapacity += capacity;
        totalGlobalRevenue += revenue;
        
        // Count actual occupants and paid status
        let occupants = 0;
        let paid = 0;
        if (residentsData[index]) {
            const students = Object.values(residentsData[index]);
            occupants = students.length;
            students.forEach(s => {
                if (s.payment === '已繳費') paid++;
            });
        }
        
        // Calculate Actual Revenue based on sequential bed mapping
        let bedPointers = 0;
        Object.entries(dorm.rooms)
            .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
            .forEach(([capStr, data]) => {
                const totalBeds = parseInt(capStr) * data.count;
                for(let k = 0; k < totalBeds; k++) {
                    const bedId = allBedsData[index] ? allBedsData[index][bedPointers] : null;
                    if(bedId) {
                        const student = residentsData[index] ? residentsData[index][bedId] : null;
                        if (student && student.payment === '已繳費') {
                            actualRevenue += data.price;
                        }
                    }
                    bedPointers++;
                }
            });
        
        totalGlobalOccupants += occupants;
        totalGlobalPaid += paid;

        // View Mode Rooms
        let viewRoomsHtml = '';
        Object.entries(dorm.rooms)
            .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
            .forEach(([cap, data]) => {
                viewRoomsHtml += `
                    <div class="room-item">
                        <div class="room-type">
                            <i class="fa-solid fa-bed"></i> ${cap}人房
                        </div>
                        <div class="room-details">
                            <span class="room-count">${data.count} 間</span>
                            <span class="room-price">${formatCurrency(data.price)} /學期</span>
                        </div>
                    </div>
                `;
            });

        // Edit Mode Rooms
        let editRoomsHtml = '';
        Object.entries(dorm.rooms)
            .sort((a, b) => parseInt(a[0]) - parseInt(b[0]))
            .forEach(([cap, data]) => {
                editRoomsHtml += `
                    <div class="form-row room-type-row" data-capacity="${cap}">
                        <div class="form-row-title">
                            <span><i class="fa-solid fa-bed"></i> ${cap}人房設定</span>
                            <button type="button" class="btn btn-outline text-danger" style="width: auto; flex-shrink: 0; margin: 0; padding: 0.25rem 0.6rem; font-size: 0.85rem; border-color: var(--danger); display: flex; align-items: center; gap: 0.35rem; background: transparent;" onclick="removeRoomType(${index}, '${cap}')" title="刪除此房型">
                                <i class="fa-solid fa-trash-can"></i> 刪除
                            </button>
                        </div>
                        <div class="form-group">
                            <label>間數</label>
                            <input type="number" min="0" class="form-control input-count" value="${data.count}">
                        </div>
                        <div class="form-group">
                            <label>價格 (每學期)</label>
                            <input type="number" min="0" step="100" class="form-control input-price" value="${data.price}">
                        </div>
                    </div>
                `;
            });
            
        totalGlobalActualRevenue += actualRevenue;
        
        // Append row to stats table
        statsTableHtml += `
            <tr style="border-bottom: 1px solid var(--border-color); animation: fadeIn 0.3s ease;">
                <td style="padding: 0.75rem; font-weight:600;">${dorm.name}</td>
                <td style="padding: 0.75rem;">${capacity} 人</td>
                <td style="padding: 0.75rem; color:var(--primary); font-weight:bold;">${occupants} 人</td>
                <td style="padding: 0.75rem; color:var(--text-muted);">${capacity - occupants} 床</td>
                <td style="padding: 0.75rem; color:var(--secondary); font-weight:bold;">${paid} 人</td>
                <td style="padding: 0.75rem; color:#059669; font-weight:bold;">${formatCurrency(actualRevenue)}</td>
            </tr>
        `;

        // Card HTML
        const card = document.createElement('div');
        card.className = 'dorm-card';
        card.setAttribute('data-index', index);
        
        card.innerHTML = `
            <div class="dorm-header">
                <h3><i class="fa-solid fa-hotel text-primary"></i> ${dorm.name}</h3>
                <button class="btn-icon edit-btn" onclick="toggleEditMode(${index}, true)" title="編輯宿舍設定">
                    <i class="fa-solid fa-pen"></i>
                </button>
            </div>
            
            <div class="dorm-content">
                <!-- VIEW MODE -->
                <div class="view-mode">
                    <div class="room-list">
                        ${viewRoomsHtml}
                    </div>
                    
                    <div class="dorm-summary" style="margin-bottom: 1.5rem; display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem;">
                        <div class="summary-item">
                            <span class="label">可容納人數</span>
                            <span class="value" style="font-size:1rem;">${capacity} 人</span>
                        </div>
                        <div class="summary-item" style="text-align:right;">
                            <span class="label">當前空床</span>
                            <span class="value" style="font-size:1rem; color:var(--text-muted);">${capacity - occupants} 床</span>
                        </div>
                        <div class="summary-item">
                            <span class="label">當前入住</span>
                            <span class="value" style="font-size:1rem; color:var(--primary);">${occupants} 人</span>
                        </div>
                        <div class="summary-item" style="text-align:right;">
                            <span class="label">已繳費</span>
                            <span class="value" style="font-size:1rem; color:var(--secondary);">${paid} 人</span>
                        </div>
                    </div>
                    
                    <div style="display:flex; flex-direction:column; gap:0.5rem; width:100%;">
                        <div style="display:flex; gap:0.5rem; width:100%;">
                            <button class="btn btn-outline" onclick="openResidentModal(${index})" style="flex:1; border: 1px solid var(--border-color); color: var(--text-main); background: #fdfdfd; display: flex; justify-content: center; align-items: center; gap: 0.5rem;">
                                <i class="fa-solid fa-users"></i> 床位設定
                            </button>
                            <button class="btn btn-outline" onclick="queryEmptyBeds(${index})" style="flex:1; border: 1px solid var(--border-color); color: var(--text-main); background: #fdfdfd; display: flex; justify-content: center; align-items: center; gap: 0.5rem;">
                                <i class="fa-solid fa-bed-pulse"></i> 查詢空床
                            </button>
                        </div>
                        <div style="display:flex; gap:0.5rem; width:100%;">
                            <button class="btn btn-outline" onclick="showDormStudents(${index})" style="flex:1; border: 1px solid var(--primary); color: var(--primary); background: #fdfdfd; display: flex; justify-content: center; align-items: center; gap: 0.5rem;">
                                <i class="fa-solid fa-list-ul"></i> 住宿名單
                            </button>
                            <button class="btn btn-outline" onclick="exportDormToExcel(${index})" style="flex:1; border: 1px solid #107c41; color: #107c41; background: #fcfcfc; display: flex; justify-content: center; align-items: center; gap: 0.5rem; transition: all 0.2s ease;">
                                <i class="fa-solid fa-file-excel"></i> 匯出名單
                            </button>
                        </div>
                    </div>
                </div>
                
                <!-- EDIT MODE -->
                <div class="edit-mode hidden">
                    <div class="edit-form">
                        <div class="form-group">
                            <label>宿舍名稱</label>
                            <input type="text" class="form-control input-name" value="${dorm.name}">
                        </div>
                        
                        ${editRoomsHtml}
                        
                        <button class="btn btn-outline" style="border-style: dashed; color: var(--primary); border-color: var(--primary);" onclick="addRoomType(${index})">
                            <i class="fa-solid fa-plus"></i> 新增房型屬性
                        </button>
                        
                        <div class="form-actions" style="margin-top: 2rem;">
                            <button class="btn btn-secondary" onclick="toggleEditMode(${index}, false)">取消</button>
                            <button class="btn btn-primary" onclick="saveDorm(${index})"><i class="fa-solid fa-check"></i> 儲存設定</button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        dormGrid.appendChild(card);
    });
    
    // Update global dashboard stats
    if(totalCapacityEl) totalCapacityEl.textContent = totalGlobalCapacity.toLocaleString() + ' 人';
    if(totalRevenueEl) totalRevenueEl.textContent = formatCurrency(totalGlobalRevenue);
    if(totalActualRevenueEl) totalActualRevenueEl.textContent = formatCurrency(totalGlobalActualRevenue);
    if(totalOccupantsEl) totalOccupantsEl.textContent = totalGlobalOccupants.toLocaleString() + ' 人';
    if(totalEmptyEl) totalEmptyEl.textContent = (totalGlobalCapacity - totalGlobalOccupants).toLocaleString() + ' 床';
    if(totalPaidEl) totalPaidEl.textContent = totalGlobalPaid.toLocaleString() + ' 人';
    
    // Update Stats Table
    if(statsTableBody) statsTableBody.innerHTML = statsTableHtml;
}

// ----- Search and Query Logic -----
function showInfoModal(title, htmlContent) {
    document.getElementById('info-modal-title').innerHTML = title;
    document.getElementById('info-modal-body').innerHTML = htmlContent;
    document.getElementById('info-modal').classList.remove('hidden');
}

function showListModal(title, htmlContent) {
    document.getElementById('list-modal-title').innerHTML = title;
    document.getElementById('list-modal-body').innerHTML = htmlContent;
    document.getElementById('list-modal').classList.remove('hidden');
}

window.addResidentFromEmptyBed = function(index, bedId) {
    document.getElementById('info-modal').classList.add('hidden');
    openResidentModal(index);
    // Add small delay to allow modal render to complete
    setTimeout(() => {
        selectBed(index, bedId);
        const activeItem = document.querySelector('.bed-item.active');
        if (activeItem) {
            activeItem.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }, 50);
}

window.searchStudent = function() {
    const query = document.getElementById('student-search-input').value.trim();
    if (!query) {
        alert("請輸入學號！");
        return;
    }
    
    const results = [];
    
    dormsData.forEach((dorm, index) => {
        if (!residentsData[index]) return;
        
        for (const [bedId, student] of Object.entries(residentsData[index])) {
            if (student.id && student.id.toLowerCase() === query.toLowerCase()) {
                results.push({ dorm: dorm.name, bedId: bedId, student: student });
            }
        }
    });
    
    if (results.length === 0) {
        showInfoModal('<i class="fa-solid fa-circle-xmark text-danger"></i> 查無資料', `<p>找不到學號為 <b>${query}</b> 的學生住宿紀錄。</p>`);
    } else {
        let html = '';
        results.forEach(res => {
            const s = res.student;
            html += `
                <div style="background:var(--body-bg); padding:1rem; border-radius:8px; margin-bottom:1rem; border:1px solid var(--border-color);">
                    <h4 style="margin:0 0 0.5rem 0; color:var(--primary);"><i class="fa-solid fa-hotel"></i> ${res.dorm} - 床位 ${res.bedId}</h4>
                    <table style="width:100%; border-collapse:collapse; font-size:0.9rem;">
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted); width:35%;">學號</td><td style="font-weight:500;">${s.id}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">姓名</td><td style="font-weight:500;">${s.name}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">科系/班級</td><td style="font-weight:500;">${s.dept} ${s.year} ${s.className}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">國籍</td><td style="font-weight:500;">${s.nationality}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">手機</td><td style="font-weight:500;">${s.phone || '未填寫'}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">入住日期</td><td style="font-weight:500;">${s.checkIn || '未填寫'}</td></tr>
                        <tr><td style="padding:0.25rem 0; color:var(--text-muted);">繳費狀態</td><td style="font-weight:500; color:${s.payment === '已繳費' ? 'var(--secondary)' : 'var(--danger)'};">${s.payment}</td></tr>
                    </table>
                </div>
            `;
        });
        showInfoModal('<i class="fa-solid fa-magnifying-glass text-primary"></i> 學生查詢結果', html);
    }
};

window.queryEmptyBeds = function(index) {
    const dorm = dormsData[index];
    const beds = allBedsData[index] || [];
    
    // Calculate total actual beds mapped dynamically based on room setup
    let requiredBeds = 0;
    Object.values(dorm.rooms).forEach(room => {
        requiredBeds += (room.count * Object.keys(dorm.rooms).find(k => dorm.rooms[k] === room)); 
        // this calculation gets messy easily since Object.values doesn't know capacity directly.
    });
    
    // Let's accurately re-simulate the bed distribution
    const sortedRooms = Object.entries(dorm.rooms).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    
    let totalAssignedBeds = 0;
    sortedRooms.forEach(([capStr, data]) => {
        totalAssignedBeds += parseInt(capStr) * data.count;
    });
    
    const relevantBeds = beds.slice(0, totalAssignedBeds);
    
    const emptyBeds = relevantBeds.filter(bedId => !residentsData[index] || !residentsData[index][bedId]);
    
    if (emptyBeds.length === 0) {
        showInfoModal(`<i class="fa-solid fa-building-circle-check text-primary"></i> ${dorm.name} 空床查詢`, `<p>目前 <b>${dorm.name}</b> 沒有任何空床，已完全滿租！</p>`);
    } else {
        const bedBadges = emptyBeds.map(b => `
            <button onclick="addResidentFromEmptyBed(${index}, '${b}')" class="empty-bed-btn" title="點選為此床位新增住宿生">
                <i class="fa-solid fa-plus" style="font-size:0.75rem; opacity:0.6;"></i> ${b}
            </button>
        `).join('');
        
        showInfoModal(`<i class="fa-solid fa-bed-pulse text-primary"></i> ${dorm.name} 空床清單 (共 ${emptyBeds.length} 床)`, `
            <p>以下為目前無住宿生配置的床號：</p>
            <div style="background:var(--body-bg); padding:1rem; border-radius:8px; border:1px solid var(--border-color);">
                ${bedBadges}
            </div>
        `);
    }
};

window.showDormStudents = function(index) {
    const dorm = dormsData[index];
    const residents = residentsData[index];

    if (!residents || Object.keys(residents).length === 0) {
        showListModal(`<i class="fa-solid fa-list-ul text-primary"></i> ${dorm.name} 住宿名單`, `<div style="text-align:center; padding: 3rem; color: var(--text-muted); font-size: 1.25rem;">目前 <b>${dorm.name}</b> 沒有任何住宿生資料。</div>`);
        return;
    }

    const beds = allBedsData[index] || [];
    const sortedRooms = Object.entries(dorm.rooms).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    let totalAssignedBeds = 0;
    sortedRooms.forEach(([capStr, data]) => {
        totalAssignedBeds += parseInt(capStr) * data.count;
    });
    const relevantBeds = beds.slice(0, totalAssignedBeds);

    let html = `<div style="overflow-x: auto; height: 100%; border: 1px solid var(--border-color); border-radius: var(--radius-md);"><table style="width:100%; border-collapse:collapse; font-size:1.1rem; text-align: left; min-width: 800px; background: white;">
        <thead style="position: sticky; top: 0; background: #F9FAFB; z-index: 10; font-weight: 600;">
            <tr style="border-bottom: 2px solid var(--border-color); color:var(--text-muted);">
                <th style="padding: 1rem; background: #F9FAFB;">床位</th>
                <th style="padding: 1rem; background: #F9FAFB;">學號</th>
                <th style="padding: 1rem; background: #F9FAFB;">姓名</th>
                <th style="padding: 1rem; background: #F9FAFB;">科系班級</th>
                <th style="padding: 1rem; background: #F9FAFB;">繳費狀態</th>
            </tr>
        </thead>
        <tbody>`;
    
    let studentCount = 0;

    relevantBeds.forEach(bedId => {
        const s = residents[bedId];
        if (s) {
            studentCount++;
            html += `
                <tr style="border-bottom: 1px solid var(--border-color); transition: background-color 0.2s;">
                    <td style="padding: 1rem; font-family: monospace; font-size: 1.15rem;">${bedId}</td>
                    <td style="padding: 1rem;">${s.id || ''}</td>
                    <td style="padding: 1rem; font-weight: 600; color: var(--text-main);">${s.name || ''}</td>
                    <td style="padding: 1rem; color:var(--text-muted);">${(s.dept || '')} ${(s.className || '')}</td>
                    <td style="padding: 1rem; font-weight:600; color:${s.payment === '已繳費' ? 'var(--secondary)' : 'var(--danger)'};">
                        <span style="background: ${s.payment === '已繳費' ? '#ECFDF5' : '#FEF2F2'}; padding: 0.25rem 0.75rem; border-radius: 999px;">${s.payment || '未繳費'}</span>
                    </td>
                </tr>
            `;
        }
    });

    html += `</tbody></table></div>`;
    
    if (studentCount === 0) {
        html = `<div style="text-align:center; padding: 3rem; color: var(--text-muted); font-size: 1.25rem;">目前 <b>${dorm.name}</b> 的床位上沒有任何住宿生資料。</div>`;
    }

    showListModal(`<i class="fa-solid fa-list-ul text-primary"></i> ${dorm.name} 住宿名單 (共 ${studentCount} 人)`, html);
};

window.showEmptyBeds = function() {
    let html = '';
    let totalEmptyGlobal = 0;
    
    dormsData.forEach((dorm, index) => {
        const beds = allBedsData[index] || [];
        const sortedRooms = Object.entries(dorm.rooms).sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
        let totalAssignedBeds = 0;
        sortedRooms.forEach(([capStr, data]) => {
            totalAssignedBeds += parseInt(capStr) * data.count;
        });
        const relevantBeds = beds.slice(0, totalAssignedBeds);
        const emptyBeds = relevantBeds.filter(bedId => !residentsData[index] || !residentsData[index][bedId]);
        
        totalEmptyGlobal += emptyBeds.length;
        
        if (emptyBeds.length > 0) {
            const bedBadges = emptyBeds.map(b => `
                <button onclick="addResidentFromEmptyBed(${index}, '${b}')" class="empty-bed-btn" title="點選為此床位新增住宿生">
                    <i class="fa-solid fa-plus" style="font-size:0.75rem; opacity:0.6;"></i> ${b}
                </button>
            `).join('');
            html += `<h4 style="margin:1rem 0 0.5rem 0; color:var(--primary);"><i class="fa-solid fa-hotel"></i> ${dorm.name} (${emptyBeds.length} 床)</h4>`;
            html += `<div style="background:var(--body-bg); padding:0.5rem; border-radius:8px; border:1px solid var(--border-color);">${bedBadges}</div>`;
        }
    });
    
    if (totalEmptyGlobal === 0) {
        showInfoModal(`<i class="fa-solid fa-building-circle-check text-primary"></i> 全校空床查詢`, `<p>目前全校所有宿舍皆無空床，已完全滿租！</p>`);
    } else {
        showInfoModal(`<i class="fa-solid fa-bed-pulse text-primary"></i> 全校空床清單 (共 ${totalEmptyGlobal} 床)`, html);
    }
};

// Handle Enter key for search
document.addEventListener('DOMContentLoaded', () => {
    init();
    
    const searchInput = document.getElementById('student-search-input');
    if (searchInput) {
        searchInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                searchStudent();
            }
        });
    }
});

// Excel Export Functions
window.exportDormToExcel = function(index) {
    if (typeof XLSX === 'undefined') {
        alert('載入 Excel 匯出模組失敗，請檢查網路連線。');
        return;
    }
    
    const dorm = dormsData[index];
    const residents = residentsData[index];
    if (!residents || Object.keys(residents).length === 0) {
        alert(`${dorm.name} 目前沒有任何住宿生資料可供匯出。`);
        return;
    }

    const data = [];
    
    for (const [bedId, student] of Object.entries(residents)) {
        data.push({
            '宿舍名稱': dorm.name,
            '床位號碼': bedId,
            '學號': student.id || '',
            '姓名': student.name || '',
            '科系': student.dept || '',
            '班級': student.className || '',
            '年級': student.year || '',
            '國籍': student.nationality || '',
            '手機號碼': student.phone || '',
            '入住日期': student.checkIn || '',
            '繳費狀態': student.payment || ''
        });
    }
    
    // Sort by bedId simply
    data.sort((a, b) => a['床位號碼'].localeCompare(b['床位號碼']));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '住宿名單');
    
    const filename = `${dorm.name}_住宿學生名單_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
};

window.exportAllDormsToExcel = function() {
    if (typeof XLSX === 'undefined') {
        alert('載入 Excel 匯出模組失敗，請檢查網路連線。');
        return;
    }
    
    const wb = XLSX.utils.book_new();
    let hasData = false;

    dormsData.forEach((dorm, index) => {
        const residents = residentsData[index];
        if (!residents || Object.keys(residents).length === 0) return;
        
        hasData = true;
        const data = [];
        
        for (const [bedId, student] of Object.entries(residents)) {
            data.push({
                '宿舍名稱': dorm.name,
                '床位號碼': bedId,
                '學號': student.id || '',
                '姓名': student.name || '',
                '科系': student.dept || '',
                '班級': student.className || '',
                '年級': student.year || '',
                '國籍': student.nationality || '',
                '手機號碼': student.phone || '',
                '入住日期': student.checkIn || '',
                '繳費狀態': student.payment || ''
            });
        }
        
        data.sort((a, b) => a['床位號碼'].localeCompare(b['床位號碼']));
        const ws = XLSX.utils.json_to_sheet(data);
        
        // Excel worksheet names have a max limit of 31 chars and cannot contain certain chars
        // Safe sheet name, replace invalid characters like : \ / ? * [ ]
        let sheetName = dorm.name.replace(/[:\\/?*\[\]]/g, '').substring(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    
    if (!hasData) {
        alert('目前所有宿舍皆無住宿生資料可供匯出。');
        return;
    }
    
    const filename = `全校宿舍_住宿學生名單_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(wb, filename);
};

// Download Import Template
window.downloadImportTemplate = function() {
    if (typeof XLSX === 'undefined') {
        alert('載入 Excel 模組失敗，請檢查網路連線。');
        return;
    }
    
    // Create sample data matching the required format
    const sampleData = [
        {
            '宿舍名稱': '女生第一宿舍 (晨曦樓)',
            '床位號碼': 'W1-105-1',
            '學號': 'B1140001',
            '姓名': '王小明',
            '科系': '資訊工程學系',
            '班級': '甲班',
            '年級': '大一',
            '國籍': '台灣',
            '手機號碼': '0912-345-678',
            '入住日期': new Date().toISOString().split('T')[0],
            '繳費狀態': '已繳費'
        },
        {
            '宿舍名稱': '大直宿舍 (弘道樓)',
            '床位號碼': 'D1-201-1',
            '學號': 'B1140002',
            '姓名': '李大華',
            '科系': '電機工程學系',
            '班級': '乙班',
            '年級': '大二',
            '國籍': '台灣',
            '手機號碼': '0987-654-321',
            '入住日期': new Date().toISOString().split('T')[0],
            '繳費狀態': '未繳費'
        }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    
    // Add some helpful column widths
    const wscols = [
        {wch: 25}, // 宿舍名稱
        {wch: 15}, // 床位號碼
        {wch: 12}, // 學號
        {wch: 10}, // 姓名
        {wch: 15}, // 科系
        {wch: 10}, // 班級
        {wch: 10}, // 年級
        {wch: 10}, // 國籍
        {wch: 15}, // 手機號碼
        {wch: 15}, // 入住日期
        {wch: 15}  // 繳費狀態
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '住宿匯入範例');
    
    XLSX.writeFile(wb, '批次匯入範例格式.xlsx');
};

// Handle Excel Import
window.handleExcelImport = function(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (typeof XLSX === 'undefined') {
        alert('載入 Excel 解析模組失敗，請檢查網路連線。');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            
            // Assume data is in the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            if (json.length === 0) {
                alert("未在檔案中找到資料，請確認格式是否正確。");
                event.target.value = "";
                return;
            }

            let successCount = 0;
            let errorCount = 0;

            json.forEach((row) => {
                const dormName = row['宿舍名稱'] ? row['宿舍名稱'].toString().trim() : '';
                const bedId = row['床位號碼'] ? row['床位號碼'].toString().trim() : '';
                const studentId = row['學號'] ? row['學號'].toString().trim() : '';
                const studentName = row['姓名'] ? row['姓名'].toString().trim() : '';
                
                if (!dormName || !bedId || !studentId || !studentName) {
                    errorCount++;
                    return; // Skip invalid rows
                }

                let targetIndex = -1;
                if (dormName.includes("一") || dormName.includes("晨曦")) targetIndex = 0;
                else if (dormName.includes("二") || dormName.includes("星辰")) targetIndex = 1;
                else if (dormName.includes("新德惠") || dormName.includes("雅風")) targetIndex = 3;
                else if (dormName.includes("德惠") || dormName.includes("翠林")) targetIndex = 2;
                else if (dormName.includes("大直") || dormName.includes("弘道")) targetIndex = 4;
                
                if (targetIndex !== -1) {
                    if (!residentsData[targetIndex]) residentsData[targetIndex] = {};
                    
                    let payment = row['繳費狀態'] ? row['繳費狀態'].toString().trim() : '未繳費';
                    if (!['已繳費', '未繳費', '分期/就貸'].includes(payment)) {
                        if (payment.includes('已')) payment = '已繳費';
                        else if (payment.includes('貸') || payment.includes('分期')) payment = '分期/就貸';
                        else payment = '未繳費';
                    }

                    residentsData[targetIndex][bedId] = {
                        id: studentId,
                        name: studentName,
                        dept: row['科系'] ? row['科系'].toString().trim() : '',
                        className: row['班級'] ? row['班級'].toString().trim() : '',
                        year: row['年級'] ? row['年級'].toString().trim() : '',
                        nationality: row['國籍'] ? row['國籍'].toString().trim() : '台灣',
                        phone: row['手機號碼'] ? row['手機號碼'].toString().trim() : '',
                        checkIn: row['入住日期'] ? row['入住日期'].toString().trim() : new Date().toISOString().split('T')[0],
                        payment: payment
                    };
                    successCount++;
                } else {
                    errorCount++;
                }
            });

            if (successCount > 0) {
                localStorage.setItem('residentsData', JSON.stringify(residentsData));
                alert(`匯入成功！成功匯入 ${successCount} 筆學生資料。` + (errorCount > 0 ? `\n有 ${errorCount} 筆資料因必填欄位不完整或無法對應宿舍而被略過。` : ''));
                render(); // refresh UI
            } else {
                alert(`匯入失敗，未能成功辨識出有效的學生資料 (共 ${errorCount} 筆錯誤)。請確認您的宿舍名稱、床位號碼、學號與姓名皆正確填寫。`);
            }
        } catch (error) {
            console.error(error);
            alert("讀取檔案發生錯誤，請確保您上傳的是正確的 Excel 檔案。");
        } finally {
            // Reset input so the exact same file can be uploaded again if needed
            event.target.value = "";
        }
    };
    reader.readAsArrayBuffer(file);
};
