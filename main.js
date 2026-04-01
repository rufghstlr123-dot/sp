const LS_KEYS = {
    SCHEDULE: 'sd_scheduleData',
    CUSTOM_HOLIDAYS: 'sd_customHolidaysData',
    CHECKED_STATUS: 'sd_checkedData',
    MEMO: 'sd_memo',
    EMPLOYEES: 'sd_employeesData_all',
    ROW_ORDER: 'sd_rowOrder_2',
    EXCLUDED_BRANDS: 'sd_excluded_brands',
    INTEREST_FREE: 'sd_interest_free'
};

// --- Firebase Initialization ---
const firebaseConfig = {
    databaseURL: "https://spcalander-default-rtdb.asia-southeast1.firebasedatabase.app/"
};
firebase.initializeApp(firebaseConfig);
const db = firebase.database();

let allMonthsEmployeesData = {};

function saveLocalState(key, data) {
    try {
        // Save to Firebase (Remote)
        db.ref(key).set(data);
        // Backup to LocalStorage
        localStorage.setItem(key, JSON.stringify(data));
    } catch (e) {
        console.error("Firebase Sync Error:", e);
    }
}

function loadLocalState(key, defaultVal) {
    try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : defaultVal;
    } catch {
        return defaultVal;
    }
}

function getRowOrder(calendarType) {
    const orders = loadLocalState(LS_KEYS.ROW_ORDER, {});
    return orders[calendarType] || [];
}

function saveRowOrder(calendarType, orderArray) {
    const orders = loadLocalState(LS_KEYS.ROW_ORDER, {});
    orders[calendarType] = orderArray;
    saveLocalState(LS_KEYS.ROW_ORDER, orders);
}

document.documentElement.lang = "ko";

// --- Application State ---
let currentUserRole = null; // 'admin' or 'guest'
let currentDate = new Date();
let scheduleData = {};  // Format: { "YYYY-MM-DD_eventId": "Memo/Status" }
let employeesData = {}; // Format: { "eventId": { name: "Sale", details: "...", startDate: "...", endDate: "..." } }
let customHolidaysData = {}; // Custom holiday name overrides
let checkedData = {}; // Format: { "cellKey": true/false } для 연차/대휴 체크상태

// Selections
let sidebarSelectedShift = null;
let selectedKeys = new Set();
let pivotKey = null;
let activeKey = null;

function parseLocalDate(dateStr) {
    if (!dateStr) return new Date();
    const [y, m, d] = dateStr.split('-');
    return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
}
let allCellCoords = {};
let coordsToKey = {};
let isDragging = false;
let draggedEmpId = null;
let clipboardBuffer = null; // { baseRow, baseCol, shifts: { "rowOffset,colOffset": "Shift" } }

// Undo Stack
let undoStack = [];
const MAX_UNDO_STACK = 50;

const DAYS_KR = ['일', '월', '화', '수', '목', '금', '토'];

// --- DOM References ---
const loginContainer = document.getElementById('login-container');
const mainApp = document.getElementById('main-app');
const loginIdInput = document.getElementById('login-id');
const loginPwInput = document.getElementById('login-pw');
const loginBtn = document.getElementById('login-btn');
const loginError = document.getElementById('login-error');
const logoutBtn = document.getElementById('logout-btn');
const userRoleDisplay = document.getElementById('user-role-display');

const currentMonthDisplay = document.getElementById('current-month-display');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const rosterGrid = document.getElementById('roster-grid');

const eventTypeInput = document.getElementById('event-type');
const eventTypeSelect = document.getElementById('event-type-select');
const empNameInput = document.getElementById('emp-name');
const eventDetailsInput = document.getElementById('event-details');
const eventBrandInput = document.getElementById('event-brand');
const eventStartInput = document.getElementById('event-start');
const eventEndInput = document.getElementById('event-end');
const eventBudgetInput = document.getElementById('event-budget');
// ... other DOM refs
const eventModal = document.getElementById('event-modal');
const modalCloseBtn = document.getElementById('modal-close-btn');
const modalEditBtn = document.getElementById('modal-edit-btn');
const modalName = document.getElementById('modal-event-name');
const modalType = document.getElementById('modal-event-type');
const modalPeriod = document.getElementById('modal-event-period');
const modalBudget = document.getElementById('modal-event-budget');
const modalDetails = document.getElementById('modal-event-details');
const modalDetailsLabel = document.getElementById('modal-details-label');
const modalMemo = document.getElementById('modal-event-memo');
const modalMemoGroup = document.getElementById('modal-memo-group');
const modalBrand = document.getElementById('modal-event-brand');
const sidebarBrandGroup = document.getElementById('sidebar-brand-group');
const sidebarBudgetGroup = document.getElementById('sidebar-budget-group');
const modalBudgetGroup = document.getElementById('modal-budget-group');
const modalBrandGroup = document.getElementById('modal-brand-group');
const eventDetailsLabel = document.getElementById('event-details-label');
const sidebarMemoSection = document.getElementById('sidebar-memo-section');
const memoArea = document.getElementById('memo-area');
const memoStatus = document.getElementById('memo-status');
const empErrorMsg = document.getElementById('emp-error-msg');
const saveEmpBtn = document.getElementById('save-emp-btn');
const editBtnGroup = document.getElementById('edit-btn-group');
const updateEmpBtn = document.getElementById('update-emp-btn');
const duplicateEmpBtn = document.getElementById('duplicate-emp-btn');
const deleteEmpBtn = document.getElementById('delete-emp-btn');
const cancelEditBtn = document.getElementById('cancel-edit-btn');
const calTabs = document.querySelectorAll('.cal-tab');
const sidebarEventTypeLabel = document.getElementById('sidebar-event-type-label');
const sidebarEventNameGroup = document.getElementById('sidebar-event-name-group');
const sidebarEventDetailsGroup = document.getElementById('sidebar-event-details-group');
const sidebarBrandLabel = document.getElementById('sidebar-brand-label');
const sidebarVenueFloorGroup = document.getElementById('sidebar-venue-floor-group');
const sidebarVenueFloorInput = document.getElementById('event-venue-floor');
const sidebarVenueNameGroup = document.getElementById('sidebar-venue-name-group');
const sidebarVenueNameInput = document.getElementById('event-venue-name');
const sidebarVenueDetailGroup = document.getElementById('sidebar-venue-detail-group');
const sidebarVenueDetailInput = document.getElementById('event-venue-detail');
const sidebarTeamGroup = document.getElementById('sidebar-team-group');
const sidebarTeamInput = document.getElementById('event-team');
const modalVenueFloorGroup = document.getElementById('modal-venue-floor-group');
const modalVenueFloorText = document.getElementById('modal-event-venue-floor');
const modalVenueNameGroup = document.getElementById('modal-venue-name-group');
const modalVenueNameText = document.getElementById('modal-event-venue-name');
const modalVenueDetailGroup = document.getElementById('modal-venue-detail-group');
const modalVenueDetailText = document.getElementById('modal-event-venue-detail');
const modalTeamGroup = document.getElementById('modal-team-group');
const modalTeamText = document.getElementById('modal-event-team');
const modalBrandLabel = document.getElementById('modal-brand-label');
const modalEventDetailsGroup = document.getElementById('modal-event-details-group');
const excludedBrandsBtn = document.getElementById('excluded-brands-btn');
const excludedBrandsModal = document.getElementById('excluded-brands-modal');
const excludedModalCloseBtn = document.getElementById('excluded-modal-close-btn');
const excludedBrandsList = document.getElementById('excluded-brands-list');
const excludedViewContainer = document.getElementById('excluded-view-container');
const excludedEditContainer = document.getElementById('excluded-edit-container');
const excludedBrandsInput = document.getElementById('excluded-brands-input');
const excludedEditBtn = document.getElementById('excluded-edit-btn');
const excludedSaveBtn = document.getElementById('excluded-save-btn');
const excludedCancelBtn = document.getElementById('excluded-cancel-btn');
const interestFreeBtn = document.getElementById('interest-free-btn');
const interestFreeModal = document.getElementById('interest-free-modal');
const interestFreeModalCloseBtn = document.getElementById('interest-free-modal-close-btn');
const interestFreeList = document.getElementById('interest-free-list');
const interestFreeViewContainer = document.getElementById('interest-free-view-container');
const interestFreeEditContainer = document.getElementById('interest-free-edit-container');
const interestFreeInput = document.getElementById('interest-free-input');
const interestFreeEditBtn = document.getElementById('interest-free-edit-btn');
const interestFreeSaveBtn = document.getElementById('interest-free-save-btn');
const interestFreeCancelBtn = document.getElementById('interest-free-cancel-btn');

const searchBtn = document.getElementById('search-btn');
const searchModal = document.getElementById('search-modal');
const searchModalCloseBtn = document.getElementById('search-modal-close-btn');
const searchInput = document.getElementById('search-input');
const searchResultsList = document.getElementById('search-results-list');
const searchNoResults = document.getElementById('search-no-results');

let currentCalendarType = '사은행사';
let editingEventId = null;
// --- Initialization ---
function init() {
    // Check session
    const savedRole = localStorage.getItem('sd_roster_role');
    if (savedRole) {
        showApp(savedRole);
    }

    setupLoginListeners();
    // loadAllLocalData(); // Replaced by Real-time Sync

    // --- Start Real-time Sync from Firebase ---
    Object.keys(LS_KEYS).forEach(k => {
        const key = LS_KEYS[k];
        db.ref(key).on('value', (snapshot) => {
            const data = snapshot.val();
            
            // Backup to localStorage so that getExcludedBrands etc. read the latest data
            if (data !== null) {
                localStorage.setItem(key, JSON.stringify(data));
            } else {
                localStorage.removeItem(key);
            }

            const fallbackData = data || {};

            // Map Firebase data to local state variables
            if (key === LS_KEYS.SCHEDULE) scheduleData = fallbackData;
            if (key === LS_KEYS.CUSTOM_HOLIDAYS) customHolidaysData = fallbackData;
            if (key === LS_KEYS.CHECKED_STATUS) checkedData = fallbackData;
            if (key === LS_KEYS.EMPLOYEES) {
                let isNested = false;
                if (fallbackData) {
                    for (let fKey in fallbackData) {
                        if (fKey.match(/^\d{4}-\d{1,2}$/)) {
                            isNested = true;
                            break;
                        }
                    }
                }

                if (isNested) {
                    let flattened = {};
                    for (let mKey in fallbackData) {
                        if (typeof fallbackData[mKey] === 'object') {
                            Object.assign(flattened, fallbackData[mKey]);
                        }
                    }
                    employeesData = flattened;
                    db.ref(LS_KEYS.EMPLOYEES).set(flattened);
                } else {
                    employeesData = fallbackData;
                }
            }

            if (key === LS_KEYS.EXCLUDED_BRANDS) {
                const excludedModal = document.getElementById('excluded-brands-modal');
                if (excludedModal && !excludedModal.classList.contains('hidden')) {
                    renderExcludedBrands();
                }
            }

            if (typeof renderRoster === 'function') {
                renderRoster();
            }
        });
    });
    
    listenToCurrentMonthInterestFree();

    setupEventListeners();
    setupGlobalKeyboard();

    // Ensure sidebar UI matches initial currentCalendarType
    const initialTab = Array.from(calTabs).find(t => t.dataset.type === currentCalendarType);
    if (initialTab) {
        initialTab.click();
    }

    // Global Mouse Up to stop dragging
    window.addEventListener('mouseup', () => {
        isDragging = false;
    });
}

function setupLoginListeners() {
    loginBtn.onclick = handleLogin;
    loginPwInput.onkeydown = (e) => { if (e.key === 'Enter') handleLogin(); };
    logoutBtn.onclick = () => {
        localStorage.removeItem('sd_roster_role');
        location.reload();
    };
}

function handleLogin() {
    const id = loginIdInput.value.trim();
    const pw = loginPwInput.value.trim();
    loginError.textContent = "";

    if (id === "admin" && pw === "0626") {
        showApp('admin');
    } else if (id === "ar" && pw === "2222") {
        showApp('guest');
    } else {
        loginError.textContent = "아이디 또는 비밀번호가 잘못되었습니다.";
    }
}

function showApp(role) {
    currentUserRole = role;
    localStorage.setItem('sd_roster_role', role);
    loginContainer.style.display = 'none';
    mainApp.style.display = 'flex';
    userRoleDisplay.textContent = role === 'admin' ? "관리자 모드" : "게스트 모드";

    // Hide or show general administrative UI
    if (role === 'guest') {
        document.querySelectorAll('.admin-only').forEach(el => el.style.display = 'none');
    } else {
        document.querySelectorAll('.admin-only').forEach(el => el.style.display = '');
    }

    // Ensure correct sidebar visibility on app show
    const activeTab = Array.from(calTabs).find(t => t.classList.contains('active'));
    if (activeTab) {
        activeTab.click();
    }

    renderRoster();
}

function hasPermission(action, targetEmpId = null) {
    if (currentUserRole === 'admin') return true;

    // For read-only mode, only admin has permissions
    return false;
}

let currentInterestFreeRef = null;
function listenToCurrentMonthInterestFree() {
    if (currentInterestFreeRef) {
        currentInterestFreeRef.off();
    }
    const monthKey = `${LS_KEYS.INTEREST_FREE}_${getMonthKey()}`;
    currentInterestFreeRef = db.ref(monthKey);
    currentInterestFreeRef.on('value', (snapshot) => {
        const data = snapshot.val();
        if (data !== null) {
            localStorage.setItem(monthKey, JSON.stringify(data));
        } else {
            localStorage.removeItem(monthKey);
        }
        
        const modal = document.getElementById('interest-free-modal');
        if (modal && !modal.classList.contains('hidden')) {
            renderInterestFree();
        }
    });
}


// --- Local Storage Sync ---
function getMonthKey() {
    return `${currentDate.getFullYear()}-${currentDate.getMonth() + 1}`;
}

function loadAllLocalData() {
    scheduleData = loadLocalState(LS_KEYS.SCHEDULE, {});
    customHolidaysData = loadLocalState(LS_KEYS.CUSTOM_HOLIDAYS, {});
    checkedData = loadLocalState(LS_KEYS.CHECKED_STATUS, {});
    allMonthsEmployeesData = loadLocalState(LS_KEYS.EMPLOYEES, {});

    loadEmployeesForCurrentMonth();
    renderRoster();
}

function loadEmployeesForCurrentMonth() {
    // No-op: Data is now global and unpartitioned, so we don't need to load per month
}

function saveCurrentMonthEmployees() {
    saveLocalState(LS_KEYS.EMPLOYEES, employeesData);
}



function duplicateEvent() {
    if (!editingEventId || !employeesData[editingEventId]) return;

    // Use the period currently in the input fields for the duplicate
    const newStart = eventStartInput.value;
    const newEnd = eventEndInput.value;

    if (!newStart || !newEnd) {
        alert("복사할 행사의 기간(시작일과 종료일)을 먼저 입력해주세요.");
        return;
    }

    const original = employeesData[editingEventId];
    const newId = 'evt_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);

    employeesData[newId] = {
        ...original,
        name: original.name,
        startDate: newStart,
        endDate: newEnd,
        sortOrder: Date.now()
    };

    saveCurrentMonthEmployees();
    renderRoster();
    cancelEditEvent();
    alert(`상담/행사가 새로운 기간(${newStart} ~ ${newEnd})으로 복사되었습니다.`);
}

function saveEmployee(eventType, name, details, startDate, endDate, budget, brand, memo) {
    if (!hasPermission('add_employee')) return;
    const newId = 'evt_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
    employeesData[newId] = { category: currentCalendarType, type: eventType || currentCalendarType, name, details, startDate, endDate, budget, brand, memo, sortOrder: Date.now() };
    saveCurrentMonthEmployees();
    renderRoster();
}

function deleteEmployee(empId) {
    if (!hasPermission('delete_employee')) return;
    if (confirm("정말 이 행사를 삭제하시겠습니까?")) {
        delete employeesData[empId];
        saveCurrentMonthEmployees();

        // Remove shifts for this employee
        Object.keys(scheduleData).forEach(key => {
            if (key.endsWith('_' + empId)) {
                delete scheduleData[key];
            }
        });
        saveLocalState(LS_KEYS.SCHEDULE, scheduleData);

        renderRoster();
    }
}

window.showEventModal = function (empId) {
    const e = employeesData[empId];
    if (!e) return;

    modalPeriod.textContent = `${e.startDate} ~ ${e.endDate}`;

    if (e.category === '행사장') {
        if (modalName) {
            modalName.style.display = 'block';
            modalName.textContent = e.brand || '브랜드 정보 없음';
        }
        if (modalType) modalType.style.display = 'none';
        if (modalVenueFloorGroup) modalVenueFloorGroup.style.display = 'block';
        if (modalVenueFloorText) {
            modalVenueFloorText.textContent = e.floor || '';
            modalVenueFloorText.style.fontWeight = '400'; // Thin/Normal
        }
        if (modalVenueNameGroup) modalVenueNameGroup.style.display = 'block';
        if (modalVenueNameText) {
            modalVenueNameText.textContent = e.type || '';
            modalVenueNameText.style.fontWeight = '400'; // Thin/Normal
        }
        if (modalVenueDetailGroup) modalVenueDetailGroup.style.display = 'block';
        if (modalVenueDetailText) {
            modalVenueDetailText.textContent = e.venueDetail || '상세위치 없음';
            modalVenueDetailText.style.fontWeight = '400';
        }
        if (modalTeamGroup) {
            if (e.team) {
                modalTeamGroup.style.display = 'block';
                modalTeamText.textContent = e.team;
            } else {
                modalTeamGroup.style.display = 'none';
            }
        }
        if (modalEventDetailsGroup) modalEventDetailsGroup.style.display = 'none';
        if (modalBudgetGroup) modalBudgetGroup.style.display = 'none';
        if (modalBrandGroup) modalBrandGroup.style.display = 'none';
    } else {
        if (modalName) {
            modalName.style.display = 'block';
            modalName.textContent = e.name || '행사 정보';
        }
        if (modalType) {
            modalType.style.display = 'block';
            modalType.textContent = e.type || e.category || '자사';
        }
        if (modalVenueFloorGroup) modalVenueFloorGroup.style.display = 'none';
        if (modalVenueNameGroup) modalVenueNameGroup.style.display = 'none';
        if (modalTeamGroup) modalTeamGroup.style.display = 'none';
        if (modalBrandLabel) modalBrandLabel.textContent = '대상 브랜드';
        if (modalBrandGroup) modalBrandGroup.style.display = 'block';
        if (modalEventDetailsGroup) {
            modalEventDetailsGroup.style.display = 'block';
            if (modalDetailsLabel) modalDetailsLabel.textContent = e.category === '이벤트' ? '장소' : '행사 내용';
            if (modalDetails) modalDetails.textContent = e.details || '상세내용 없음';
        }
        if (modalBudgetGroup) {
            if (e.budget && e.category !== '이벤트') {
                modalBudget.textContent = `${e.budget} 만원`;
                modalBudgetGroup.style.display = 'block';
            } else {
                modalBudgetGroup.style.display = 'none';
            }
        }
    }

    if (e.memo) {
        modalMemo.textContent = e.memo;
        if (modalMemoGroup) modalMemoGroup.style.display = 'block';
    } else {
        if (modalMemoGroup) modalMemoGroup.style.display = 'none';
    }

    if (e.brand && e.category !== '이벤트' && e.category !== '행사장') {
        modalBrand.textContent = e.brand;
        if (modalBrandGroup) modalBrandGroup.style.display = 'block';
    } else {
        if (modalBrandGroup) modalBrandGroup.style.display = 'none';
    }

    if (hasPermission('add_employee')) {
        const modalEditGroup = document.getElementById('modal-edit-btn-group');
        if (modalEditGroup) modalEditGroup.style.display = 'flex';
        modalEditBtn.onclick = () => {
            closeModal();
            editEvent(empId);
        };
    } else {
        const modalEditGroup = document.getElementById('modal-edit-btn-group');
        if (modalEditGroup) modalEditGroup.style.display = 'none';
    }

    eventModal.classList.remove('hidden');
};

function closeModal() {
    eventModal.classList.add('hidden');
}

window.editEvent = function (empId) {
    if (!hasPermission('add_employee')) return;
    const e = employeesData[empId];
    if (!e) return;

    editingEventId = empId;

    // Switch to category tab
    const tabToClick = Array.from(calTabs).find(t => t.dataset.type === (e.category || '사은행사'));
    if (tabToClick && currentCalendarType !== tabToClick.dataset.type) {
        tabToClick.click();
    }

    setTimeout(() => {
        if (currentCalendarType === '사은행사') {
            const predefinedOpts = ['자사', '타사', '전관', '부문(패션)', '부문(라이프스타일)', '사은품'];
            if (predefinedOpts.includes(e.type || '자사')) {
                eventTypeSelect.value = e.type || '자사';
                eventTypeInput.style.display = 'none';
            } else {
                eventTypeSelect.value = '직접입력';
                eventTypeInput.style.display = 'block';
                eventTypeInput.placeholder = '행사 종류를 직접 입력하세요';
                eventTypeInput.readOnly = false;
                eventTypeInput.style.backgroundColor = 'white';
                eventTypeInput.style.color = 'black';
                eventTypeInput.value = e.type || '';
            }
        } else if (currentCalendarType === '행사장') {
            sidebarVenueFloorInput.value = e.floor || '';
            sidebarVenueNameInput.value = e.type || '';
            if (sidebarVenueDetailInput) sidebarVenueDetailInput.value = e.venueDetail || '';
            if (sidebarTeamInput) sidebarTeamInput.value = e.team || '';
        } else {
            eventTypeInput.value = e.type || currentCalendarType;
        }

        empNameInput.value = e.name || '';
        eventDetailsInput.value = e.details || '';
        eventBrandInput.value = e.brand || '';
        eventBudgetInput.value = e.budget || '';
        memoArea.value = e.memo || '';
        eventStartInput.value = e.startDate || '';
        eventEndInput.value = e.endDate || '';

        saveEmpBtn.style.display = 'none';
        editBtnGroup.style.display = 'flex';
    }, 50);
};

window.cancelEditEvent = function () {
    editingEventId = null;
    empNameInput.value = '';
    eventDetailsInput.value = '';
    eventBrandInput.value = '';
    eventBudgetInput.value = '';
    memoArea.value = '';
    eventStartInput.value = '';
    eventEndInput.value = '';
    if (sidebarVenueFloorInput) sidebarVenueFloorInput.value = '';
    if (sidebarVenueNameInput) sidebarVenueNameInput.value = '';
    if (sidebarVenueDetailInput) sidebarVenueDetailInput.value = '';
    if (sidebarTeamInput) sidebarTeamInput.value = '';

    if (currentCalendarType === '행사장') {
        eventTypeInput.value = '';
    } else if (currentCalendarType === '사은행사') {
        eventTypeSelect.value = '자사';
        eventTypeInput.value = '';
        eventTypeInput.style.display = 'none';
    } else {
        eventTypeInput.value = '이벤트';
    }

    saveEmpBtn.style.display = 'block';
    editBtnGroup.style.display = 'none';
};

// handleDrop removed as dragging event rows is no longer applicable

function applyShiftChanges(updates, pushToUndo = true) {
    if (Object.keys(updates).length === 0) return;

    if (pushToUndo) {
        const previousState = {};
        for (const key in updates) {
            previousState[key] = scheduleData[key] || "";
        }
        undoStack.push(previousState);
        if (undoStack.length > MAX_UNDO_STACK) undoStack.shift();
    }

    let changed = false;
    for (const key in updates) {
        const firstUnderscore = key.indexOf('_');
        if (firstUnderscore !== -1) {
            const empId = key.substring(firstUnderscore + 1);
            if (hasPermission('edit_shift', empId)) {
                let val = updates[key];
                const value = (typeof val === 'string') ? val.toUpperCase() : val;
                if (value === "") {
                    delete scheduleData[key];
                } else {
                    scheduleData[key] = value;
                }
                changed = true;
            }
        }
    }

    if (changed) {
        saveLocalState(LS_KEYS.SCHEDULE, scheduleData);
        renderRoster();
    }
}

function handleUndo() {
    if (undoStack.length === 0) return;
    const prevState = undoStack.pop();
    applyShiftChanges(prevState, false);
}

function saveCustomHoliday(dateStr, holidayName) {
    if (!hasPermission('edit_holiday')) return;
    if (holidayName === "") {
        delete customHolidaysData[dateStr];
    } else {
        customHolidaysData[dateStr] = holidayName;
    }
    saveLocalState(LS_KEYS.CUSTOM_HOLIDAYS, customHolidaysData);
    renderRoster();
}

// --- Holiday Logic ---
const PUBLIC_HOLIDAYS_FIXED = {
    "01-01": "신정", "03-01": "삼일절", "05-05": "어린이날", "06-06": "현충일",
    "08-15": "광복절", "10-03": "개천절", "10-09": "한글날", "12-25": "성탄절"
};

const HOLIDAYS_VAR_2026 = {
    "2026-02-16": "설날 연휴", "2026-02-17": "설날", "2026-02-18": "설날 연휴",
    "2026-03-02": "대체공휴일", "2026-05-24": "부처님오신날", "2026-05-25": "대체공휴일",
    "2026-08-17": "대체공휴일", "2026-09-24": "추석 연휴", "2026-09-25": "추석",
    "2026-09-26": "추석 연휴", "2026-10-05": "대체공휴일",
};

function getHolidayName(year, month, day) {
    const mmdd = String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    const yyyymmdd = year + '-' + mmdd;

    // Check custom overrides first0
    if (customHolidaysData[yyyymmdd]) return customHolidaysData[yyyymmdd];

    if (year === 2026 && HOLIDAYS_VAR_2026[yyyymmdd]) return HOLIDAYS_VAR_2026[yyyymmdd];
    if (PUBLIC_HOLIDAYS_FIXED[mmdd]) return PUBLIC_HOLIDAYS_FIXED[mmdd];
    return null;
}

// --- Selection Helpers ---
function performRangeSelection(startKey, endKey) {
    const start = allCellCoords[startKey];
    const end = allCellCoords[endKey];
    if (!start || !end) return;

    const r1 = Math.min(start.row, end.row);
    const r2 = Math.max(start.row, end.row);
    const c1 = Math.min(start.col, end.col);
    const c2 = Math.max(start.col, end.col);

    selectedKeys.clear();
    Object.keys(allCellCoords).forEach(k => {
        const coord = allCellCoords[k];
        if (coord.row >= r1 && coord.row <= r2 && coord.col >= c1 && coord.col <= c2) {
            selectedKeys.add(k);
        }
    });
}

function handleMouseDown(e, key) {
    // If Painter mode is active, don't drag select, just apply
    if (sidebarSelectedShift !== null) {
        applyShiftToTarget(key, sidebarSelectedShift);
        return;
    }

    if (e.ctrlKey || e.metaKey) {
        if (selectedKeys.has(key)) selectedKeys.delete(key);
        else selectedKeys.add(key);
        pivotKey = key;
        activeKey = key;
    } else if (e.shiftKey && pivotKey) {
        activeKey = key;
        performRangeSelection(pivotKey, activeKey);
    } else {
        selectedKeys.clear();
        selectedKeys.add(key);
        pivotKey = key;
        activeKey = key;
        isDragging = true;
    }
    renderRoster();
}

function handleMouseOver(e, key) {
    if (isDragging && pivotKey) {
        performRangeSelection(pivotKey, key);
        renderRoster();
    }
}

function applyShiftToTarget(key, shift) {
    const info = allCellCoords[key];
    if (!info) return;

    if (selectedKeys.has(key)) {
        // Toggle logic for bulk: if first cell in selection already has this shift, we clear all.
        // Otherwise, we set all. (Standard on/off toggle behavior)
        const currentShift = scheduleData[key];
        const targetShift = (currentShift === shift) ? "" : shift;

        const updates = {};
        selectedKeys.forEach(k => {
            const ki = allCellCoords[k];
            if (ki) updates[`${ki.date}_${ki.eId}`] = targetShift;
        });
        applyShiftChanges(updates);
    } else {
        const currentShift = scheduleData[key];
        const targetShift = (currentShift === shift) ? "" : shift;
        applyShiftChanges({ [`${info.date}_${info.eId}`]: targetShift });
    }
}

// --- Display Logic ---
function renderRoster() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();

    currentMonthDisplay.textContent = `${year}년 ${month + 1}월`;
    rosterGrid.innerHTML = '';
    rosterGrid.classList.remove('cal-type-행사장', 'cal-type-이벤트', 'cal-type-사은행사');
    rosterGrid.classList.add(`cal-type-${currentCalendarType}`);

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    if (currentCalendarType === '행사장') {
        rosterGrid.style.gridTemplateColumns = `60px 120px repeat(${daysInMonth}, minmax(80px, 1fr))`;
    } else {
        rosterGrid.style.gridTemplateColumns = `minmax(180px, auto) repeat(${daysInMonth}, minmax(80px, 1fr))`;
    }

    let currentRowIdx = 0;
    allCellCoords = {};
    coordsToKey = {};

    let unOrderedTypesToRender = [];
    if (currentCalendarType === '행사장') {
        const eventsInCat = Object.values(employeesData).filter(e => {
            const cat = e.category || (['사은행사', '이벤트', '행사장'].includes(e.type) ? e.type : '사은행사');
            return cat === '행사장';
        });
        const firstDayOfMonth = new Date(year, month, 1);
        const lastDayOfMonth = new Date(year, month + 1, 0);

        // Use a composite key including venueDetail to separate rows
        unOrderedTypesToRender = [...new Set(eventsInCat.map(e => (e.floor || ' - ') + '|||' + e.type + '|||' + (e.venueDetail || '')))].sort().filter(compKey => {
            const [f, t, vd] = compKey.split('|||');
            return eventsInCat.some(e => {
                if (e.floor !== (f === ' - ' ? undefined : f) && e.floor !== f) return false;
                if (e.type !== t) return false;
                if ((e.venueDetail || '') !== vd) return false;
                if (!e.startDate || !e.endDate) return false;
                const start = parseLocalDate(e.startDate);
                const end = parseLocalDate(e.endDate);
                return start <= lastDayOfMonth && end >= firstDayOfMonth;
            });
        });
    } else if (currentCalendarType === '사은행사') {
        const defaultTypes = ['자사', '타사', '전관', '부문(패션)', '부문(라이프스타일)', '사은품'];
        const eventsInCat = Object.values(employeesData).filter(e => {
            const cat = e.category || (['사은행사', '이벤트', '행사장'].includes(e.type) ? e.type : '사은행사');
            return cat === '사은행사';
        });
        const additionalTypes = [...new Set(eventsInCat.map(e => e.type))].filter(t => !defaultTypes.includes(t));

        const firstDayOfMonth = new Date(year, month, 1);
        const lastDayOfMonth = new Date(year, month + 1, 0);

        unOrderedTypesToRender = [...defaultTypes, ...additionalTypes].filter(type => {
            return eventsInCat.some(e => {
                if (e.type !== type) return false;
                if (!e.startDate || !e.endDate) return false;
                const start = parseLocalDate(e.startDate);
                const end = parseLocalDate(e.endDate);
                return start <= lastDayOfMonth && end >= firstDayOfMonth;
            });
        });
    } else if (currentCalendarType === '이벤트') {
        const eventsInCat = Object.values(employeesData).filter(e => {
            const cat = e.category || (['사은행사', '이벤트', '행사장'].includes(e.type) ? e.type : '사은행사');
            return cat === '이벤트';
        });
        const firstDayOfMonth = new Date(year, month, 1);
        const lastDayOfMonth = new Date(year, month + 1, 0);

        unOrderedTypesToRender = [...new Set(eventsInCat.map(e => e.type))].sort().filter(type => {
            return eventsInCat.some(e => {
                if (e.type !== type) return false;
                if (!e.startDate || !e.endDate) return false;
                const start = parseLocalDate(e.startDate);
                const end = parseLocalDate(e.endDate);
                return start <= lastDayOfMonth && end >= firstDayOfMonth;
            });
        });
    }

    const savedOrder = getRowOrder(currentCalendarType);
    let typesToRender = [...unOrderedTypesToRender].sort((a, b) => {
        const indexA = savedOrder.indexOf(a);
        const indexB = savedOrder.indexOf(b);
        if (indexA === -1 && indexB === -1) return 0;
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
    });

    if (currentCalendarType !== '행사장' && currentCalendarType !== '사은행사' && currentCalendarType !== '이벤트') {
        typesToRender = [currentCalendarType];
    }

    // --- 2. Date Header ---
    if (currentCalendarType === '행사장') {
        const floorHeader = document.createElement('div');
        floorHeader.className = 'r-cell r-corner';
        floorHeader.textContent = '위치';
        rosterGrid.appendChild(floorHeader);

        const nameHeader = document.createElement('div');
        nameHeader.className = 'r-cell r-corner';
        nameHeader.textContent = '행사장명';
        rosterGrid.appendChild(nameHeader);
    } else {
        const cornerCell = document.createElement('div');
        cornerCell.className = 'r-cell r-corner';
        cornerCell.textContent = '일자';
        rosterGrid.appendChild(cornerCell);
    }

    for (let day = 1; day <= daysInMonth; day++) {
        const dayOfWeek = new Date(year, month, day).getDay();
        const headerCell = document.createElement('div');
        headerCell.className = `r-cell r-header`;
        if (dayOfWeek === 0 || dayOfWeek === 6 || getHolidayName(year, month + 1, day)) {
            headerCell.classList.add('is-holiday');
            headerCell.classList.add('is-holiday-bg');
        }
        if (dayOfWeek === 0) {
            headerCell.classList.add('sun-border');
        }

        const numSpan = document.createElement('span');
        numSpan.className = 'date-num';
        numSpan.textContent = day;

        const dowSpan = document.createElement('span');
        dowSpan.textContent = DAYS_KR[dayOfWeek];

        headerCell.appendChild(numSpan);
        headerCell.appendChild(dowSpan);
        rosterGrid.appendChild(headerCell);
    }

    // --- 3. Holiday Row ---
    const holidayRowLabel = document.createElement('div');
    holidayRowLabel.className = 'r-cell r-col-header holiday-row-label';
    if (currentCalendarType === '행사장') {
        holidayRowLabel.style.gridColumn = 'span 2';
    }
    holidayRowLabel.textContent = '공휴일';
    holidayRowLabel.style.borderBottom = '1px solid black';
    rosterGrid.appendChild(holidayRowLabel);

    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = new Date(year, month, day).getDay();
        const holidayName = getHolidayName(year, month + 1, day);
        const holiCell = document.createElement('div');
        holiCell.className = 'r-cell holiday-cell';
        holiCell.style.borderBottom = '1px solid black';
        if (dayOfWeek === 0) holiCell.classList.add('sun-border');
        if (holidayName) holiCell.textContent = holidayName;

        holiCell.contentEditable = hasPermission('edit_holiday') ? "true" : "false";
        holiCell.onblur = () => {
            if (!hasPermission('edit_holiday')) return;
            const newValue = holiCell.innerText.trim();
            saveCustomHoliday(dateStr, newValue);
        };
        holiCell.onkeydown = (e) => {
            if (e.key === 'Enter') {
                if (e.altKey) {
                    e.preventDefault();
                    document.execCommand('insertLineBreak');
                    return;
                }
                e.preventDefault();
                holiCell.blur();
            }
        };

        rosterGrid.appendChild(holiCell);
    }

    // --- 4. Event Type Row ---
    typesToRender.forEach((compType, typeIndex) => {
        let type = compType;
        let floor = '';
        let vDetail = '';

        if (currentCalendarType === '행사장') {
            const parts = compType.split('|||');
            floor = parts[0];
            type = parts[1];
            vDetail = parts[2] || '';
            if (floor === ' - ') floor = '';
        }

        const nextCompType = typesToRender[typeIndex + 1];
        const isLastTypeInFloor = currentCalendarType === '행사장' &&
            (!nextCompType || nextCompType.split('|||')[0] !== compType.split('|||')[0]);
        const bottomBorderStyle = isLastTypeInFloor ? '2.5px solid #222' : '1px solid black';

        try {
            const firstDayOfMonth = new Date(year, month, 1);
            const lastDayOfMonth = new Date(year, month + 1, 0);

            const eventsForType = Object.entries(employeesData)
                .filter(([id, e]) => {
                    const evtCategory = e.category || (['사은행사', '이벤트', '행사장'].includes(e.type) ? e.type : '사은행사');
                    if (evtCategory !== currentCalendarType) return false;
                    if (currentCalendarType === '행사장') {
                        if (e.type !== type) return false;
                        if ((e.floor || '') !== floor) return false;
                        if ((e.venueDetail || '') !== vDetail) return false; // Filter by venueDetail
                    } else {
                        if (e.type !== type) return false;
                    }
                    if (!e.startDate || !e.endDate) return false;

                    const start = parseLocalDate(e.startDate);
                    const end = parseLocalDate(e.endDate);
                    // Check if event overlaps with current month
                    return start <= lastDayOfMonth && end >= firstDayOfMonth;
                })
                .sort((a, b) => parseLocalDate(a[1].startDate) - parseLocalDate(b[1].startDate) || a[0].localeCompare(b[0]));

            if (eventsForType.length === 0) return;

            const lanes = [];
            eventsForType.forEach(eventData => {
                const [id, e] = eventData;
                const start = parseLocalDate(e.startDate).getTime();
                const end = parseLocalDate(e.endDate).getTime();
                let placed = false;
                for (let i = 0; i < lanes.length; i++) {
                    let overlaps = false;
                    for (const laneEvt of lanes[i]) {
                        const lStart = parseLocalDate(laneEvt[1].startDate).getTime();
                        const lEnd = parseLocalDate(laneEvt[1].endDate).getTime();
                        if (start <= lEnd && end >= lStart) {
                            overlaps = true; break;
                        }
                    }
                    if (!overlaps) {
                        lanes[i].push(eventData);
                        placed = true;
                        break;
                    }
                }
                if (!placed) lanes.push([eventData]);
            });

            if (lanes.length === 0) return; // Skip rendering if no lanes (no events)

            // Base color coding for floors in Venue calendar
            const getFloorColors = (f) => {
                const floorMap = {
                    '1': '#dcfce7', // Light Green
                    '2': '#dbeafe', // Light Blue
                    '3': '#fee2e2', // Light Red/Pink
                    '4': '#fef9c3', // Light Yellow
                    '5': '#f3e8ff', // Light Purple
                    '6': '#ffedd5', // Light Orange
                    '7': '#ccfbf1', // Light Teal
                    '8': '#fce7f3', // Light Pink
                    '9': '#cffafe', // Light Cyan
                    'B1': '#f1f5f9', // Light Slate
                    'B2': '#e2e8f0'  // Light Gray
                };
                const nText = f.replace(/[^0-9]/g, '');
                const isB = f.toUpperCase().includes('B');
                const key = isB ? 'B' + nText : nText;
                const floorColor = floorMap[key] || '#f8fafc';
                return { f: floorColor, n: floorColor, g: '#ffffff', e: floorColor };
            };
            const floorColors = getFloorColors(floor);

            if (currentCalendarType === '행사장') {
                const floorCell = document.createElement('div');
                floorCell.className = 'r-cell r-col-header';
                floorCell.style.setProperty('border-bottom', bottomBorderStyle, 'important');
                if (lanes.length > 1) floorCell.style.gridRow = `span ${lanes.length}`;
                floorCell.style.justifyContent = 'center';
                floorCell.style.textAlign = 'center';
                floorCell.style.setProperty('background-color', floorColors.f, 'important');
                floorCell.style.borderRight = '1px solid #ddd';

                floorCell.innerHTML = `<span style="font-size: 0.85rem; font-weight: bold; color: var(--text-main); width: 100%; text-align: center;">${floor}</span>`;
                rosterGrid.appendChild(floorCell);
            }

            const nameCell = document.createElement('div');
            nameCell.className = 'r-cell r-col-header employee-name-cell';
            nameCell.style.setProperty('border-bottom', bottomBorderStyle, 'important');
            nameCell.style.justifyContent = 'flex-start';
            nameCell.style.paddingLeft = '12px';
            if (lanes.length > 1) {
                nameCell.style.gridRow = `span ${lanes.length}`;
            }

            // Apply name column color if in Venue calendar
            if (currentCalendarType === '행사장') {
                nameCell.style.setProperty('background-color', floorColors.n, 'important');
            }

            // Use the vDetail from the composite key
            const venueDetailToDisplay = vDetail;

            nameCell.innerHTML = `
                    <div class="emp-display-name-container" style="text-align: left; width: 100%; display: flex; flex-direction: column; gap: 2px;">
                        <span class="emp-display-name" style="font-size: 0.95rem; font-weight: bold; color: black; line-height: 1.1;">${type}</span>
                        ${venueDetailToDisplay ? `<span style="font-size: 0.65rem; font-weight: 600; color: #666; font-style: normal; line-height: 1.1;">${venueDetailToDisplay}</span>` : ''}
                    </div>
                `;

            if (hasPermission('edit_holiday')) {
                nameCell.draggable = true;
                nameCell.dataset.eventType = compType; // Use composite key for drag/drop logic
                nameCell.addEventListener('dragstart', handleRowDragStart);
                nameCell.addEventListener('dragover', handleRowDragOver);
                nameCell.addEventListener('drop', handleRowDrop);
                nameCell.addEventListener('dragenter', handleRowDragEnter);
                nameCell.addEventListener('dragleave', handleRowDragLeave);
            }

            rosterGrid.appendChild(nameCell);

            lanes.forEach((laneEvents, laneIdx) => {
                for (let day = 1; day <= daysInMonth; day++) {
                    const currentCellDate = new Date(year, month, day);
                    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

                    const overlappingEvents = laneEvents.filter(([id, e]) => {
                        const start = parseLocalDate(e.startDate);
                        const end = parseLocalDate(e.endDate);
                        return currentCellDate >= start && currentCellDate <= end;
                    });

                    let spanLength = 1;
                    let eventData = null;

                    if (overlappingEvents.length > 0) {
                        eventData = overlappingEvents[0];
                        const endD = parseLocalDate(eventData[1].endDate);
                        let evtLastDay = endD.getDate();
                        if (endD.getFullYear() > year || (endD.getFullYear() === year && endD.getMonth() > month)) {
                            evtLastDay = daysInMonth;
                        } else if (endD.getFullYear() < year || (endD.getFullYear() === year && endD.getMonth() < month)) {
                            evtLastDay = 0; // should be caught by filter but just in case
                        }
                        spanLength = evtLastDay - day + 1;
                        if (spanLength < 1) spanLength = 1;
                    }

                    const eventIdStr = eventData ? `_E${eventData[0]}` : '';
                    const cellKey = `${dateStr}_${compType}_L${laneIdx}${eventIdStr}`;

                    // Set coordinates for all spanned days
                    for (let s = 0; s < spanLength; s++) {
                        const sDay = day + s;
                        const sDateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(sDay).padStart(2, '0')}`;
                        coordsToKey[`${currentRowIdx},${sDay}`] = cellKey;
                    }
                    allCellCoords[cellKey] = { row: currentRowIdx, col: day, date: dateStr, eId: eventData ? eventData[0] : type };

                    const cell = document.createElement('div');
                    cell.className = 'r-cell r-entry';
                    cell.dataset.key = cellKey;
                    cell.setAttribute('lang', 'ko');
                    cell.setAttribute('inputmode', 'text');
                    cell.setAttribute('spellcheck', 'false');
                    cell.tabIndex = 0;

                    if (spanLength > 1) {
                        cell.style.gridColumn = `span ${spanLength}`;
                    }

                    if (laneIdx < lanes.length - 1) {
                        cell.style.setProperty('border-bottom', '2px solid #888', 'important');
                    } else {
                        cell.style.setProperty('border-bottom', bottomBorderStyle, 'important');
                    }

                    const endDayOfSpan = day + spanLength - 1;
                    const dayOfWeek = new Date(year, month, endDayOfSpan).getDay();
                    if (dayOfWeek === 0) {
                        cell.classList.add('sun-border');
                    }

                    let eventNamesHTML = "";
                    let cellBackgroundColor = "";
                    if (eventData) {
                        const [id, e] = eventData;

                        // Define color palette for different event types with better distinction
                        const typeColors = {
                            '자사': '#dcfce7',           // Light Green
                            '타사': '#ffedd5',           // Light Orange
                            '전관': '#bae6fd',           // Sky Blue
                            '부문(패션)': '#e0e7ff',      // Light Indigo
                            '부문(라이프스타일)': '#fee2e2', // Light Red/Pink
                            '사은품': '#ccfbf1',         // Light Teal
                            '이벤트': '#fef9c3',         // Light Yellow
                            '행사장': '#e2e8f0'          // Light Gray
                        };

                        // In Venue calendar, use floor color for events to satisfy "shading" request while grid is white
                        if (currentCalendarType === '행사장') {
                            cellBackgroundColor = floorColors.e;
                        } else {
                            cellBackgroundColor = typeColors[e.type] || "#f0fdf4";
                        }

                        const startD = parseLocalDate(e.startDate);
                        const isStart = currentCellDate.getTime() === startD.getTime() || day === 1;

                        if (hasPermission('add_employee')) {
                            cell.style.cursor = 'pointer';
                        }
                        if (isStart) {
                            const gridDisplayName = e.category === '행사장' ? (e.brand || '') : (e.name || '');
                            const gridDisplayDetail = e.category === '행사장' ? '' : (e.details || '');

                            eventNamesHTML = `<div title="${gridDisplayDetail.replace(/"/g, '&quot;')} (클릭하여 수정)" 
                                   style="font-size: 0.85rem; color: var(--text-main); text-align: left; position: absolute; left: 8px; right: 8px; top: 50%; transform: translateY(-50%); white-space: normal; word-break: break-all; z-index: 20; pointer-events: none; display: flex; flex-direction: column; gap: 5px;">
                                <div style="display: flex; align-items: baseline; gap: 6px; flex-wrap: wrap;">
                                    <strong style="font-weight:700; line-height: 1.1;">${gridDisplayName}</strong>
                                    ${(e.category !== '행사장' && e.budget) ? `<span style="font-size: 0.72rem; font-weight: 700; color: #ef4444; opacity: 1;">[${e.budget}만]</span>` : ''}
                                </div>
                                ${gridDisplayDetail ? `<span style="font-size: 0.72rem; font-weight: 700; opacity: 1; line-height: 1.1; padding-left: 4px;">${gridDisplayDetail}</span>` : ''}
                            </div>`;
                        }
                    }



                    if (selectedKeys.has(cellKey)) cell.classList.add('selected-cell');
                    if (activeKey === cellKey) {
                        cell.classList.add('active-cell');
                        cell.classList.add('is-navigating');
                    }

                    if (cellBackgroundColor) {
                        cell.style.setProperty('background-color', cellBackgroundColor, 'important');
                    }

                    cell.innerHTML = `
                        <div style="display: flex; flex-direction: column; width: 100%; align-items: center; justify-content: center; height: 100%; padding: 0;">
                            ${eventNamesHTML}
                        </div>
                    `;

                    cell.onmousedown = (e) => {
                        if (eventData) {
                            showEventModal(eventData[0]);
                        } else {
                            cancelEditEvent();
                        }
                        if (document.activeElement === cell) return;
                        if (selectedKeys.size === 1 && selectedKeys.has(cellKey)) return;
                        e.preventDefault();
                        handleMouseDown(e, cellKey);
                    };
                    cell.onmouseover = (e) => handleMouseOver(e, cellKey);

                    cell.onkeydown = (e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            e.preventDefault(); moveSelection(1, 0);
                        }
                        if (e.key === 'Tab') { e.preventDefault(); moveSelection(0, e.shiftKey ? -1 : 1); }
                        if (e.key === 'Escape') { renderRoster(); cell.blur(); }
                    };
                    rosterGrid.appendChild(cell);

                    day += spanLength - 1;
                }
                currentRowIdx++; // Increment logic per lane
            });
        } catch (err) {
            console.error("Error rendering type row:", err);
        }
    });
}

// --- Excel Navigation Logic ---
function moveSelection(dRow, dCol, shift = false, ctrl = false) {
    if (!activeKey || !allCellCoords[activeKey]) return;
    const current = allCellCoords[activeKey];

    let targetRow = current.row;
    let targetCol = current.col;

    if (ctrl) {
        // Find grid boundaries
        const rows = [...new Set(Object.values(allCellCoords).map(c => c.row))];
        const cols = [...new Set(Object.values(allCellCoords).map(c => c.col))];
        const minR = Math.min(...rows), maxR = Math.max(...rows);
        const minC = Math.min(...cols), maxC = Math.max(...cols);

        if (dRow < 0) targetRow = minR;
        if (dRow > 0) targetRow = maxR;
        if (dCol < 0) targetCol = minC;
        if (dCol > 0) targetCol = maxC;
    } else {
        if (dCol !== 0) {
            let step = Math.sign(dCol);
            while (true) {
                targetCol += step;
                const nextKey = coordsToKey[`${targetRow},${targetCol}`];
                if (!nextKey || nextKey !== activeKey) {
                    break;
                }
            }
        } else {
            targetRow += dRow;
        }
    }

    const nextKey = coordsToKey[`${targetRow},${targetCol}`];
    if (nextKey) {
        activeKey = nextKey;
        if (shift) {
            performRangeSelection(pivotKey, activeKey);
        } else {
            selectedKeys.clear();
            selectedKeys.add(nextKey);
            pivotKey = nextKey;
        }
        renderRoster();

        // Pre-emptive focus for IME support
        setTimeout(() => {
            const activeEl = document.querySelector('.active-cell');
            if (activeEl) {
                activeEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                activeEl.focus();
            }
        }, 0);
    }
}

function setupGlobalKeyboard() {
    window.onkeydown = (e) => {
        const isCtrl = e.ctrlKey || e.metaKey;
        const isEditing = document.activeElement.contentEditable === "true";

        // If user is typing in an input or focused in a modal/sidebar control, don't trigger global nav
        const isControlInsideModal = e.target.closest('.modal') !== null;
        const isControlInsideSidebar = e.target.closest('.employee-sidebar') !== null;
        const isInput = e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA';

        if (isInput || isControlInsideModal || isControlInsideSidebar) {
            return; // Let native behavior handle it
        }

        // Allow shortcuts and navigation keys even when editing a cell
        const allowedKeys = ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Enter', 'Tab', 'Delete', 'Backspace'];
        if (isEditing && !isCtrl && !allowedKeys.includes(e.key)) {
            return;
        }

        if ((selectedKeys.size > 0 && pivotKey) || activeKey) {
            // Arrow Navigation
            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
                e.preventDefault();
                const dRow = e.key === 'ArrowUp' ? -1 : (e.key === 'ArrowDown' ? 1 : 0);
                const dCol = e.key === 'ArrowLeft' ? -1 : (e.key === 'ArrowRight' ? 1 : 0);
                moveSelection(dRow, dCol, e.shiftKey, e.ctrlKey);
                return;
            }

            // Delete / Backspace: Clear entire cell content immediately
            if (e.key === 'Delete' || e.key === 'Backspace') {
                e.preventDefault();
                const updates = {};
                selectedKeys.forEach(k => {
                    const info = allCellCoords[k];
                    if (info) updates[`${info.date}_${info.eId}`] = "";

                    // Instant UI feedback
                    const cellEl = document.querySelector(`.r-entry[data-key="${k}"]`);
                    if (cellEl) {
                        cellEl.innerText = "";
                        cellEl.classList.remove('is-navigating');
                        // Clean shift classes
                        const classes = Array.from(cellEl.classList).filter(c => c.startsWith('shift-'));
                        classes.forEach(c => cellEl.classList.remove(c));
                    }
                });
                applyShiftChanges(updates);
                return;
            }

            // Instant Typing: Handled locally in cell.onkeydown for focused cells,
            // but for safety if window still catches it:
            if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !document.activeElement.classList.contains('r-entry')) {
                const targetCell = document.querySelector(`.r-entry.active-cell`);
                if (targetCell) {
                    targetCell.focus();
                    targetCell.innerText = "";
                    targetCell.classList.remove('is-navigating');
                }
            }

            // Copy/Paste/Undo
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
                if (isControlInsideModal || isControlInsideSidebar) return;
                e.preventDefault();
                handleCopy();
            }
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
                if (isControlInsideModal || isControlInsideSidebar) return;
                e.preventDefault();
                handlePaste();
            }
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
                if (isControlInsideModal || isControlInsideSidebar) return;
                e.preventDefault();
                handleUndo();
            }
        }
    };

    // Global Paste Listener for external (Excel) data
    window.addEventListener('paste', (e) => {
        const isInsideModal = e.target.closest('.modal') !== null;
        const isInsideSidebar = e.target.closest('.employee-sidebar') !== null;
        const isInput = e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA';

        // If focused on an input or inside modal/sidebar, let default behavior happen
        if (isInput || isInsideModal || isInsideSidebar) return;

        // If editing a cell, let default handle it (usually)
        if (document.activeElement.contentEditable === "true") return;

        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text) return;

        // Parse Excel/TSV data
        const rows = text.split(/\r?\n/).filter(line => line.trim() !== "").map(row => row.split('\t'));
        if (rows.length > 0) {
            handleExternalPaste(rows);
        }
    });
}

function handleCopy() {
    if (selectedKeys.size === 0) return;

    // Find top-left to determine offsets
    let minR = Infinity, minC = Infinity;
    selectedKeys.forEach(key => {
        const coord = allCellCoords[key];
        if (coord.row < minR) minR = coord.row;
        if (coord.col < minC) minC = coord.col;
    });

    const shifts = {};
    selectedKeys.forEach(key => {
        const coord = allCellCoords[key];
        const val = scheduleData[key] || "";
        shifts[`${coord.row - minR},${coord.col - minC}`] = val;
    });

    clipboardBuffer = { shifts };
    console.log("Copied", Object.keys(shifts).length, "cells");
}

function handlePaste() {
    if (clipboardBuffer) {
        handleInternalPaste();
    }
}

function handleInternalPaste() {
    if (!clipboardBuffer || !activeKey) return;
    const targetBase = allCellCoords[activeKey];
    if (!targetBase) return;

    const updates = {};
    Object.entries(clipboardBuffer.shifts).forEach(([offset, value]) => {
        const [dr, dc] = offset.split(',').map(Number);
        const tr = targetBase.row + dr;
        const tc = targetBase.col + dc;

        const targetKey = coordsToKey[`${tr},${tc}`];
        if (targetKey) {
            const targetInfo = allCellCoords[targetKey];
            if (hasPermission('edit_shift', targetInfo.eId)) {
                updates[`${targetInfo.date}_${targetInfo.eId}`] = value;
            }
        }
    });
    applyShiftChanges(updates);
}

function handleExternalPaste(dataMatrix) {
    if (!activeKey) return;
    const targetBase = allCellCoords[activeKey];
    if (!targetBase) return;

    const updates = {};
    dataMatrix.forEach((row, ri) => {
        row.forEach((cellValue, ci) => {
            const tr = targetBase.row + ri;
            const tc = targetBase.col + ci;
            const targetKey = coordsToKey[`${tr},${tc}`];
            if (targetKey) {
                const targetInfo = allCellCoords[targetKey];
                if (hasPermission('edit_shift', targetInfo.eId)) {
                    updates[`${targetInfo.date}_${targetInfo.eId}`] = cellValue.trim().toUpperCase();
                }
            }
        });
    });
    applyShiftChanges(updates);
}

// --- Interaction ---
function setupEventListeners() {
    if (prevBtn) prevBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() - 1); loadEmployeesForCurrentMonth(); listenToCurrentMonthInterestFree(); renderRoster(); };
    if (nextBtn) nextBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() + 1); loadEmployeesForCurrentMonth(); listenToCurrentMonthInterestFree(); renderRoster(); };

    if (eventTypeSelect && eventTypeInput) {
        eventTypeSelect.addEventListener('change', () => {
            if (eventTypeSelect.value === '직접입력') {
                eventTypeInput.style.display = 'block';
                eventTypeInput.placeholder = '행사 종류를 직접 입력하세요';
                eventTypeInput.readOnly = false;
                eventTypeInput.style.backgroundColor = 'white';
                eventTypeInput.style.color = 'black';
                eventTypeInput.value = '';
                eventTypeInput.focus();
            } else {
                eventTypeInput.style.display = 'none';
                eventTypeInput.value = '';
            }
        });
    }

    if (saveEmpBtn) {
        saveEmpBtn.onclick = () => {
            let type = '';
            let floor = '';
            let venueDetail = '';
            if (currentCalendarType === '행사장') {
                floor = sidebarVenueFloorInput.value.trim();
                type = sidebarVenueNameInput.value.trim();
                venueDetail = sidebarVenueDetailInput ? sidebarVenueDetailInput.value.trim() : '';
            } else if (currentCalendarType === '사은행사') {
                type = eventTypeSelect.value.trim() === '직접입력' ? eventTypeInput.value.trim() : eventTypeSelect.value.trim();
            } else {
                type = eventTypeInput.value.trim();
            }
            const name = empNameInput.value.trim();
            const details = eventDetailsInput.value.trim();
            const brand = eventBrandInput.value.trim();
            const team = sidebarTeamInput ? sidebarTeamInput.value.trim() : '';
            const budget = eventBudgetInput.value.trim();
            const memo = memoArea.value.trim();
            const start = eventStartInput.value;
            const end = eventEndInput.value;
            empErrorMsg.textContent = '';

            if (currentCalendarType === '행사장') {
                if (!floor) { empErrorMsg.textContent = '층을 입력해주세요.'; return; }
                if (!type) { empErrorMsg.textContent = '행사장명을 입력해주세요.'; return; }
            } else {
                if (!type) {
                    empErrorMsg.textContent = '행사 종류를 입력해주세요.';
                    return;
                }
            }
            if (!name && currentCalendarType !== '행사장') {
                empErrorMsg.textContent = '행사명을 입력해주세요.';
                return;
            }
            if (!start || !end) {
                empErrorMsg.textContent = '행사 기간을 모두 설정해주세요.';
                return;
            }

            const newId = 'evt_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
            employeesData[newId] = { category: currentCalendarType, type: type || currentCalendarType, floor, venueDetail, name, details, startDate: start, endDate: end, budget, brand, team, memo, sortOrder: Date.now() };
            saveCurrentMonthEmployees();
            renderRoster();

            if (currentCalendarType === '행사장') {
                sidebarVenueFloorInput.value = '';
                sidebarVenueNameInput.value = '';
                if (sidebarVenueDetailInput) sidebarVenueDetailInput.value = '';
            } else if (currentCalendarType === '사은행사') {
                eventTypeSelect.value = '자사';
                eventTypeInput.style.display = 'none';
                eventTypeInput.value = '';
            } else {
                eventTypeInput.value = '이벤트';
            }
            empNameInput.value = '';
            eventDetailsInput.value = '';
            eventBrandInput.value = '';
            eventBudgetInput.value = '';
            eventStartInput.value = '';
            eventEndInput.value = '';
            memoArea.value = ''; // Clear memo field after saving
            if (sidebarTeamInput) sidebarTeamInput.value = ''; // Clear team field after saving
        };
    }

    const formatDatePicker = (e) => {
        let val = e.target.value.replace(/[^0-9]/g, '');
        if (val.length > 8) val = val.substring(0, 8);

        let formatted = '';
        if (val.length > 0) {
            formatted = val.substring(0, 4);
            if (val.length > 4) {
                formatted += '-' + val.substring(4, 6);
                if (val.length > 6) {
                    formatted += '-' + val.substring(6, 8);
                }
            }
        }
        e.target.value = formatted;
    };

    if (eventStartInput) {
        eventStartInput.addEventListener('input', formatDatePicker);
    }
    if (eventEndInput) {
        eventEndInput.addEventListener('input', formatDatePicker);
    }

    if (empNameInput) {
        empNameInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                if (editingEventId) {
                    if (updateEmpBtn && updateEmpBtn.onclick) updateEmpBtn.onclick();
                } else {
                    if (saveEmpBtn && saveEmpBtn.onclick) saveEmpBtn.onclick();
                }
            }
        };
    }

    if (eventEndInput) {
        eventEndInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                if (editingEventId) {
                    if (updateEmpBtn && updateEmpBtn.onclick) updateEmpBtn.onclick();
                } else {
                    if (saveEmpBtn && saveEmpBtn.onclick) saveEmpBtn.onclick();
                }
            }
        };
    }

    if (updateEmpBtn) {
        updateEmpBtn.onclick = () => {
            if (!editingEventId || !employeesData[editingEventId]) return;

            let type = '';
            let floor = '';
            let venueDetail = '';
            if (currentCalendarType === '행사장') {
                floor = sidebarVenueFloorInput.value.trim();
                type = sidebarVenueNameInput.value.trim();
                venueDetail = sidebarVenueDetailInput ? sidebarVenueDetailInput.value.trim() : '';
            } else if (currentCalendarType === '사은행사') {
                type = eventTypeSelect.value.trim() === '직접입력' ? eventTypeInput.value.trim() : eventTypeSelect.value.trim();
            } else {
                type = eventTypeInput.value.trim();
            }
            const name = empNameInput.value.trim();
            const details = eventDetailsInput.value.trim();
            const brand = eventBrandInput.value.trim();
            const budget = eventBudgetInput.value.trim();
            const memo = memoArea.value.trim();
            const start = eventStartInput.value;
            const end = eventEndInput.value;
            const team = sidebarTeamInput ? sidebarTeamInput.value.trim() : '';
            empErrorMsg.textContent = '';

            if (currentCalendarType === '행사장') {
                if (!floor) { empErrorMsg.textContent = '층을 입력해주세요.'; return; }
                if (!type) { empErrorMsg.textContent = '행사장명을 입력해주세요.'; return; }
            } else {
                if (!type) {
                    empErrorMsg.textContent = currentCalendarType === '행사장' ? '행사장 위치를 입력해주세요.' : '행사 종류를 입력해주세요.';
                    return;
                }
            }
            if (!name && currentCalendarType !== '행사장') {
                empErrorMsg.textContent = '행사명을 입력해주세요.';
                return;
            }
            if (!start || !end) {
                empErrorMsg.textContent = '행사 기간을 모두 설정해주세요.';
                return;
            }

            employeesData[editingEventId] = {
                ...employeesData[editingEventId],
                category: currentCalendarType,
                type: type || currentCalendarType,
                floor,
                venueDetail,
                name,
                details,
                brand,
                team,
                budget,
                memo,
                startDate: start,
                endDate: end
            };
            saveCurrentMonthEmployees();
            renderRoster();
            cancelEditEvent();
        };
    }

    if (deleteEmpBtn) {
        deleteEmpBtn.onclick = () => {
            if (!editingEventId) return;
            deleteEmployee(editingEventId);
            cancelEditEvent();
        };
    }



    if (duplicateEmpBtn) {
        duplicateEmpBtn.onclick = duplicateEvent;
    }

    if (cancelEditBtn) {
        cancelEditBtn.onclick = cancelEditEvent;
    }

    if (calTabs) {
        calTabs.forEach(tab => {
            tab.addEventListener('click', () => {
                currentCalendarType = tab.dataset.type;
                calTabs.forEach(t => t.classList.remove('active'));
                tab.classList.add('active');

                if (currentCalendarType === '행사장') {
                    eventTypeSelect.style.display = 'none';
                    eventTypeInput.style.display = 'none';

                    if (sidebarVenueFloorGroup) sidebarVenueFloorGroup.style.display = 'block';
                    if (sidebarVenueNameGroup) sidebarVenueNameGroup.style.display = 'block';
                    if (sidebarVenueDetailGroup) sidebarVenueDetailGroup.style.display = 'block';
                    if (sidebarTeamGroup) sidebarTeamGroup.style.display = 'block';

                    if (sidebarEventTypeLabel) sidebarEventTypeLabel.style.display = 'none';
                    if (sidebarEventNameGroup) sidebarEventNameGroup.style.display = 'none';
                    if (sidebarEventDetailsGroup) sidebarEventDetailsGroup.style.display = 'none';
                    if (sidebarBrandGroup) sidebarBrandGroup.style.display = 'block';
                    if (sidebarBrandLabel) sidebarBrandLabel.textContent = '브랜드';
                    if (sidebarBudgetGroup) sidebarBudgetGroup.style.display = 'none';

                    if (sidebarMemoSection) sidebarMemoSection.style.display = 'none';
                } else if (currentCalendarType === '이벤트') {
                    eventTypeSelect.style.display = 'none';
                    eventTypeInput.style.display = '';
                    eventTypeInput.value = '이벤트';
                    eventTypeInput.readOnly = true;
                    eventTypeInput.style.backgroundColor = 'var(--bg-main)';
                    eventTypeInput.style.color = 'black';

                    if (sidebarVenueFloorGroup) sidebarVenueFloorGroup.style.display = 'none';
                    if (sidebarVenueNameGroup) sidebarVenueNameGroup.style.display = 'none';
                    if (sidebarVenueDetailGroup) sidebarVenueDetailGroup.style.display = 'none';
                    if (sidebarTeamGroup) sidebarTeamGroup.style.display = 'none';

                    if (sidebarEventTypeLabel) {
                        sidebarEventTypeLabel.style.display = 'block';
                        sidebarEventTypeLabel.textContent = '행사 종류';
                    }
                    if (sidebarEventNameGroup) sidebarEventNameGroup.style.display = 'block';
                    if (sidebarEventDetailsGroup) sidebarEventDetailsGroup.style.display = 'block';
                    if (sidebarBrandGroup) sidebarBrandGroup.style.display = 'none';
                    if (sidebarBrandLabel) sidebarBrandLabel.textContent = '이벤트명';
                    if (sidebarBudgetGroup) sidebarBudgetGroup.style.display = 'none';

                    if (sidebarMemoSection) sidebarMemoSection.style.display = 'block';
                    if (eventDetailsLabel) eventDetailsLabel.textContent = '장소';
                    if (eventDetailsInput) eventDetailsInput.placeholder = '행사 장소를 입력하세요';
                } else { // 사은행사
                    eventTypeSelect.style.display = '';
                    if (eventTypeSelect.value === '직접입력') {
                        eventTypeInput.style.display = 'block';
                        eventTypeInput.placeholder = '행사 종류를 직접 입력하세요';
                        eventTypeInput.readOnly = false;
                        eventTypeInput.style.backgroundColor = 'white';
                        eventTypeInput.style.color = 'black';
                    } else {
                        eventTypeInput.style.display = 'none';
                    }

                    if (sidebarVenueFloorGroup) sidebarVenueFloorGroup.style.display = 'none';
                    if (sidebarVenueNameGroup) sidebarVenueNameGroup.style.display = 'none';
                    if (sidebarVenueDetailGroup) sidebarVenueDetailGroup.style.display = 'none';
                    if (sidebarTeamGroup) sidebarTeamGroup.style.display = 'none';

                    if (sidebarEventTypeLabel) {
                        sidebarEventTypeLabel.style.display = 'block';
                        sidebarEventTypeLabel.textContent = '행사 종류';
                    }
                    if (sidebarEventNameGroup) sidebarEventNameGroup.style.display = 'block';
                    if (sidebarEventDetailsGroup) sidebarEventDetailsGroup.style.display = 'block';
                    if (sidebarBrandGroup) sidebarBrandGroup.style.display = 'block';
                    if (sidebarBrandLabel) sidebarBrandLabel.textContent = '행사명';
                    if (sidebarBudgetGroup) sidebarBudgetGroup.style.display = 'block';

                    if (sidebarMemoSection) sidebarMemoSection.style.display = 'block';
                    if (eventDetailsLabel) eventDetailsLabel.textContent = '행사 내용';
                    if (eventDetailsInput) eventDetailsInput.placeholder = '행사 상세 내용 입력';
                }

                cancelEditEvent();
                renderRoster();
            });
        });
    }

    if (modalCloseBtn) modalCloseBtn.onclick = closeModal;
    if (excludedBrandsBtn) {
        excludedBrandsBtn.onclick = () => {
            if (excludedBrandsModal) {
                renderExcludedBrands();
                setExcludedBrandEditMode(false);
                excludedBrandsModal.classList.remove('hidden');
            }
        };
    }
    if (excludedModalCloseBtn) {
        excludedModalCloseBtn.onclick = () => {
            if (excludedBrandsModal) excludedBrandsModal.classList.add('hidden');
        };
    }

    if (excludedEditBtn) {
        excludedEditBtn.onclick = () => {
            const brands = getExcludedBrands();
            excludedBrandsInput.value = brands.join('\n');
            setExcludedBrandEditMode(true);
        };
    }

    if (excludedSaveBtn) {
        excludedSaveBtn.onclick = () => {
            const text = excludedBrandsInput.value.trim();
            const brands = text.split('\n').map(s => s.trim()).filter(s => s !== '');
            saveLocalState(LS_KEYS.EXCLUDED_BRANDS, brands);
            renderExcludedBrands();
            setExcludedBrandEditMode(false);
        };
    }

    if (excludedCancelBtn) {
        excludedCancelBtn.onclick = () => {
            setExcludedBrandEditMode(false);
        };
    }

    if (interestFreeBtn) {
        interestFreeBtn.onclick = () => {
            if (interestFreeModal) {
                renderInterestFree();
                setInterestFreeEditMode(false);
                interestFreeModal.classList.remove('hidden');
            }
        };
    }
    if (interestFreeModalCloseBtn) {
        interestFreeModalCloseBtn.onclick = () => {
            if (interestFreeModal) interestFreeModal.classList.add('hidden');
        };
    }

    if (interestFreeEditBtn) {
        interestFreeEditBtn.onclick = () => {
            const list = getInterestFree();
            interestFreeInput.value = list.join('\n');
            setInterestFreeEditMode(true);
        };
    }

    if (interestFreeSaveBtn) {
        interestFreeSaveBtn.onclick = () => {
            const text = interestFreeInput.value.trim();
            const list = text.split('\n').map(s => s.trim()).filter(s => s !== '');
            saveLocalState(`${LS_KEYS.INTEREST_FREE}_${getMonthKey()}`, list);
            renderInterestFree();
            setInterestFreeEditMode(false);
        };
    }

    if (interestFreeCancelBtn) {
        interestFreeCancelBtn.onclick = () => {
            setInterestFreeEditMode(false);
        };
    }

    window.onclick = (e) => {
        if (e.target === eventModal) closeModal();
        if (e.target === excludedBrandsModal) {
            excludedBrandsModal.classList.add('hidden');
        }
        if (e.target === interestFreeModal) {
            interestFreeModal.classList.add('hidden');
        }
        if (e.target === searchModal) {
            searchModal.classList.add('hidden');
        }
    };

    if (searchBtn) {
        searchBtn.onclick = () => {
            if (searchModal) {
                searchModal.classList.remove('hidden');
                if (searchInput) {
                    searchInput.value = '';
                    searchInput.focus();
                }
                if (searchResultsList) searchResultsList.innerHTML = '';
                if (searchNoResults) searchNoResults.style.display = 'block';
            }
        };
    }

    if (searchModalCloseBtn) {
        searchModalCloseBtn.onclick = () => {
            if (searchModal) searchModal.classList.add('hidden');
        };
    }

    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            performSearch(e.target.value.trim());
        });
    }
}

function performSearch(query) {
    if (!searchResultsList || !searchNoResults) return;

    if (!query) {
        searchResultsList.innerHTML = '';
        searchNoResults.style.display = 'block';
        return;
    }

    const results = [];
    const lowerQuery = query.toLowerCase();

    // Search across all events in employeesData
    Object.entries(employeesData).forEach(([eventId, event]) => {
        const name = (event.name || '').toLowerCase();
        const brand = (event.brand || '').toLowerCase();
        const details = (event.details || '').toLowerCase();
        const floor = (event.floor || '').toLowerCase();
        const type = (event.type || '').toLowerCase();
        const vDetail = (event.venueDetail || '').toLowerCase();
        const team = (event.team || '').toLowerCase();

        if (name.includes(lowerQuery) ||
            brand.includes(lowerQuery) ||
            details.includes(lowerQuery) ||
            floor.includes(lowerQuery) ||
            type.includes(lowerQuery) ||
            vDetail.includes(lowerQuery) ||
            team.includes(lowerQuery)) {

            // Avoid duplicates just in case
            if (!results.find(r => r.id === eventId)) {
                let monthKey = '';
                if (event.startDate) {
                    const [y, m] = event.startDate.split('-');
                    monthKey = `${parseInt(y, 10)}-${parseInt(m, 10)}`;
                } else {
                    monthKey = `${currentDate.getFullYear()}-${currentDate.getMonth() + 1}`;
                }
                results.push({ id: eventId, monthKey, ...event });
            }
        }
    });

    if (results.length === 0) {
        searchResultsList.innerHTML = '';
        searchNoResults.style.display = 'block';
        searchNoResults.innerHTML = `
            <div style="font-size: 3rem; margin-bottom: 1rem; opacity: 0.3;">❓</div>
            '${query}'에 대한 검색 결과가 없습니다.
        `;
    } else {
        searchNoResults.style.display = 'none';
        searchResultsList.innerHTML = results.map(res => `
            <li class="search-result-item" onclick="navigateToEvent('${res.monthKey}', '${res.id}')">
                <div class="search-result-header">
                    <span class="search-result-title">${res.name || res.type}</span>
                    <span class="search-result-badge badge-${res.category}">${res.category}</span>
                </div>
                <div class="search-result-info">
                    <span class="search-result-period">${res.startDate} ~ ${res.endDate}</span>
                    <span>${res.floor ? res.floor + ' ' : ''}${res.type}${res.venueDetail ? ' (' + res.venueDetail + ')' : ''}</span>
                </div>
                ${res.brand ? `<div style="font-size: 0.75rem; color: #64748b; margin-top: 2px;">브랜드: ${res.brand}</div>` : ''}
            </li>
        `).join('');
    }
}

function navigateToEvent(monthKey, eventId) {
    const [y, m] = monthKey.split('-').map(Number);
    currentDate = new Date(y, m - 1, 1);

    // Load context for that month
    loadEmployeesForCurrentMonth();

    // Find category
    const event = employeesData[eventId];
    if (event) {
        currentCalendarType = event.category || '사은행사';
        // Update tabs UI
        calTabs.forEach(t => {
            if (t.dataset.type === currentCalendarType) t.classList.add('active');
            else t.classList.remove('active');
        });
    }

    renderRoster();
    if (searchModal) searchModal.classList.add('hidden');

    // Scroll and Highlight
    setTimeout(() => {
        const eventEl = document.querySelector(`[data-key*="_E${eventId}"]`);
        if (eventEl) {
            eventEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
            eventEl.style.boxShadow = '0 0 0 4px #fbbf24';
            setTimeout(() => {
                eventEl.style.boxShadow = '';
                showEventModal(eventId);
            }, 1000);
        } else {
            showEventModal(eventId);
        }
    }, 100);
}

function getExcludedBrands() {
    const defaultList = [
        '임대 브랜드 (일부)',
        '식품관 (지급 기준 별도 확인)',
        '명품 브랜드 (일부 제외)',
        '가전/가구 (금액 합산 불가 항목)'
    ];
    return loadLocalState(LS_KEYS.EXCLUDED_BRANDS, defaultList);
}

function renderExcludedBrands() {
    if (!excludedBrandsList) return;
    const brands = getExcludedBrands();
    excludedBrandsList.innerHTML = brands.map(brand => `
        <li style="display: flex; align-items: flex-start; gap: 8px; font-weight: 600; color: var(--text-main); line-height: 1.5;">
            <span style="width: 4px; height: 4px; background: var(--pr-color); border-radius: 50%; min-width: 4px; margin-top: 0.6rem;"></span>
            <span style="flex: 1;">${brand}</span>
        </li>
    `).join('');
}

function setExcludedBrandEditMode(isEdit) {
    if (isEdit) {
        if (excludedViewContainer) excludedViewContainer.style.display = 'none';
        if (excludedEditContainer) excludedEditContainer.style.display = 'block';
        if (excludedEditBtn) excludedEditBtn.style.display = 'none';
        if (excludedSaveBtn) excludedSaveBtn.style.display = 'block';
        if (excludedCancelBtn) excludedCancelBtn.style.display = 'block';
    } else {
        if (excludedViewContainer) excludedViewContainer.style.display = 'block';
        if (excludedEditContainer) excludedEditContainer.style.display = 'none';
        if (excludedEditBtn) excludedEditBtn.style.display = 'block';
        if (excludedSaveBtn) excludedSaveBtn.style.display = 'none';
        if (excludedCancelBtn) excludedCancelBtn.style.display = 'none';
    }
}

function getInterestFree() {
    const monthKey = `${LS_KEYS.INTEREST_FREE}_${getMonthKey()}`;
    const rawMonthData = localStorage.getItem(monthKey);
    if (rawMonthData) {
        try {
            return JSON.parse(rawMonthData);
        } catch (e) {
            console.error(e);
        }
    }
    return [];
}

function renderInterestFree() {
    if (!interestFreeList) return;
    const list = getInterestFree();
    interestFreeList.innerHTML = list.map(item => `
        <li style="display: flex; align-items: flex-start; gap: 8px; font-weight: 600; color: var(--text-main); line-height: 1.5;">
            <span style="width: 4px; height: 4px; background: var(--pr-color); border-radius: 50%; min-width: 4px; margin-top: 0.6rem;"></span>
            <span style="flex: 1;">${item}</span>
        </li>
    `).join('');
}

function setInterestFreeEditMode(isEdit) {
    if (isEdit) {
        if (interestFreeViewContainer) interestFreeViewContainer.style.display = 'none';
        if (interestFreeEditContainer) interestFreeEditContainer.style.display = 'block';
        if (interestFreeEditBtn) interestFreeEditBtn.style.display = 'none';
        if (interestFreeSaveBtn) interestFreeSaveBtn.style.display = 'block';
        if (interestFreeCancelBtn) interestFreeCancelBtn.style.display = 'block';
    } else {
        if (interestFreeViewContainer) interestFreeViewContainer.style.display = 'block';
        if (interestFreeEditContainer) interestFreeEditContainer.style.display = 'none';
        if (interestFreeEditBtn) interestFreeEditBtn.style.display = 'block';
        if (interestFreeSaveBtn) interestFreeSaveBtn.style.display = 'none';
        if (interestFreeCancelBtn) interestFreeCancelBtn.style.display = 'none';
    }
}

let draggedRowType = null;

function handleRowDragStart(e) {
    draggedRowType = this.dataset.eventType;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', this.dataset.eventType);
    setTimeout(() => this.classList.add('dragging'), 0);
}

function handleRowDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleRowDragEnter(e) {
    if (this.dataset.eventType !== draggedRowType) {
        this.classList.add('drag-over');
    }
}

function handleRowDragLeave(e) {
    this.classList.remove('drag-over');
}

function handleRowDrop(e) {
    e.stopPropagation();
    this.classList.remove('drag-over');
    if (draggedRowType === null || draggedRowType === this.dataset.eventType) {
        document.querySelectorAll('.employee-name-cell').forEach(c => c.classList.remove('dragging'));
        return false;
    }

    const targetType = this.dataset.eventType;
    const currentRenderedTypes = Array.from(document.querySelectorAll('.employee-name-cell')).map(cell => cell.dataset.eventType);

    const newOrder = [...currentRenderedTypes];
    const fromIndex = newOrder.indexOf(draggedRowType);
    const toIndex = newOrder.indexOf(targetType);

    if (fromIndex !== -1 && toIndex !== -1) {
        newOrder.splice(fromIndex, 1);
        newOrder.splice(toIndex, 0, draggedRowType);
    }

    saveRowOrder(currentCalendarType, newOrder);
    renderRoster();
    return false;
}

window.deleteEmployee = deleteEmployee;

init();
