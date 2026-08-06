// ─────────────────────────────────────────────────────────────
// 상품 카테고리 분류 규칙
// 위에서부터 순서대로 검사하여 먼저 걸리는 카테고리로 확정합니다.
// 새로운 상품군이 생기면 keys 배열에 키워드만 추가하면 됩니다.
// ─────────────────────────────────────────────────────────────
const PRODUCT_CATEGORY_RULES = [
    { name: '근조화환', color: '#5f6368', keys: ['근조화환', '근조화한', '근조3단', '근조 3단', '근조용', '조화환', '근조'] },
    { name: '축하/개업화환', color: '#ea4335', keys: ['축하화환', '개업화환', '축하3단', '축하 3단', '개업 화환', '축하화한', '결혼화환', '승진화환'] },
    { name: '화환(기타)', color: '#f28b82', keys: ['화환', '스탠드화'] },
    { name: '서양란', color: '#a142f4', keys: ['서양란', '호접란', '만천홍', '덴파레', '덴드로비움', '심비디움', '카틀레아', '팔레놉시스'] },
    { name: '동양란', color: '#7b1fa2', keys: ['동양란', '한란', '보세란', '철화', '소심', '풍란', '난 ', '난('] },
    { name: '다육/분재', color: '#34a853', keys: ['다육', '분재', '테라리움'] },
    { name: '관엽식물', color: '#188038', keys: ['관엽', '안시리움', '안스리움', '안시륨', '칼라벤자민', '벤자민', '고무나무', '금전수', '스투키', '여인초', '아레카야자', '파키라', '드라세나', '산세베리아', '산세비에리아', '홍콩야자', '떡갈고무', '율마', '몬스테라', '테이블야자', '행운목', '대박나무', '알로카시아', '싱고니움', '스킨답서스', '식물', '화분'] },
    { name: '과일/식품', color: '#fbbc04', keys: ['과일', '한우', '정육', '건강식품', '홍삼', '떡', '베이커리', '케이크'] },
    { name: '꽃바구니', color: '#1a73e8', keys: ['꽃바구니', '플라워바스켓', '바구니'] },
    { name: '꽃다발', color: '#4fc3f7', keys: ['꽃다발', '부케', '다발'] },
    { name: '조화/기타상품', color: '#9aa0a6', keys: ['조화', '디퓨저', '비누꽃', '용돈', '상품권'] }
];
const CATEGORY_ETC = { name: '기타', color: '#dadce0' };

// ─────────────────────────────────────────────────────────────
// 컬럼 자동 인식 별칭 (공백 제거 후 비교)
// ─────────────────────────────────────────────────────────────
const FIELD_ALIASES = {
    usedPoint:  ['사용적립금', '적립금사용액', '적립금사용', '사용한적립금', '적립금(사용)', '사용마일리지', '마일리지사용', '사용포인트', '포인트사용', '사용예치금'],
    savedPoint: ['지급적립금', '적립예정금액', '적립예정', '발생적립금', '적립금액', '적립마일리지', '적립포인트', '적립금지급'],
    adCost:     ['광고비', '광고비용', '광고집행비', '마케팅비', '광고단가']
};

let rawData = [];
let tempSheetData = [];
let tempHeaders = [];
let tempFileName = "";
let analyzedData = {
    monthly: {},
    categories: {},
    yearly: {},
    mom: {},
    orderDist: {},
    amountDist: {},
    productCats: {},   // { 카테고리명: { total, count, months: { 'YYYY-MM': {amount, count} }, topItems: {} } }
    points: {},        // { 'YYYY-MM': { used, saved, usedOrders } }
    fieldMeta: {}      // 자동 인식된 컬럼명 표시용
};
let charts = {
    sales: null,
    category: null,
    orderDist: null,
    amountDist: null,
    memberComp: null,
    momTrend: null,
    momGrowth: null,
    catMonthly: null,
    point: null,
    adCost: null
};
let catChartMode = 'amount'; // 'amount' | 'count'

// 상품명 → 카테고리 판별
function classifyProduct(name) {
    const s = String(name || '').replace(/\s+/g, ' ').trim();
    if (!s) return CATEGORY_ETC.name;
    const plain = s.replace(/\[.*?\]/g, ' '); // 브랜드 대괄호 제거 후 검사
    for (const rule of PRODUCT_CATEGORY_RULES) {
        for (const k of rule.keys) {
            if (plain.includes(k) || s.includes(k)) return rule.name;
        }
    }
    return CATEGORY_ETC.name;
}

function categoryColor(name) {
    const r = PRODUCT_CATEGORY_RULES.find(x => x.name === name);
    return r ? r.color : CATEGORY_ETC.color;
}

// 실제 데이터의 헤더 중 별칭과 일치하는 컬럼명을 찾아 반환
function detectFields(rows) {
    const keySet = new Set();
    rows.slice(0, 50).forEach(r => Object.keys(r).forEach(k => keySet.add(String(k).replace(/\s/g, ''))));
    const keys = Array.from(keySet);

    const findBy = (aliases) => keys.find(k => aliases.some(a => k === a))
        || keys.find(k => aliases.some(a => k.includes(a)));

    let usedPoint = findBy(FIELD_ALIASES.usedPoint);
    let usedPointGuessed = false;
    if (!usedPoint) {
        // '사용'이 명시되지 않았어도 적립금/포인트/마일리지 컬럼이 하나뿐이면 사용액으로 간주
        const cand = keys.filter(k => /적립|마일리지|포인트/.test(k));
        if (cand.length === 1) { usedPoint = cand[0]; usedPointGuessed = true; }
    }
    let savedPoint = findBy(FIELD_ALIASES.savedPoint);
    if (savedPoint && savedPoint === usedPoint) savedPoint = undefined;

    return {
        usedPoint,
        usedPointGuessed,
        savedPoint,
        adCost: findBy(FIELD_ALIASES.adCost),
        productName: keys.find(k => k === '상품명') || keys.find(k => k.includes('상품명')) || keys.find(k => k.includes('상품'))
    };
}

// 1. 탭 전환 로직 (가장 안전한 방식)
function initTabs() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-tab');
            const targetContent = document.getElementById(targetId);

            if (!targetContent) return;

            // 모든 탭 비활성화
            tabButtons.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));

            // 선택된 탭 활성화
            btn.classList.add('active');
            targetContent.classList.add('active');
        });
    });
}

// 2. 상태 표시 및 로컬 저장 복구
function showStatus(message, type = 'success') {
    const uploadZone = document.querySelector('.upload-zone');
    if (!uploadZone) return;
    let statusDiv = document.getElementById('status-msg');
    if (!statusDiv) {
        statusDiv = document.createElement('div');
        statusDiv.id = 'status-msg';
        uploadZone.parentNode.insertBefore(statusDiv, uploadZone.nextSibling);
    }
    statusDiv.className = `status-area status-${type}`;
    statusDiv.innerHTML = message;
}

window.addEventListener('DOMContentLoaded', () => {
    initTabs(); // 탭 초기화
    const savedData = localStorage.getItem('excel_auto_rawData');
    const savedFileName = localStorage.getItem('excel_auto_fileName');
    if (savedData) {
        try {
            rawData = JSON.parse(savedData);
            showStatus(`<strong>복구 완료:</strong> 마지막 파일(${savedFileName || '데이터'})을 불러왔습니다.`, 'success');
            processData();
        } catch (e) {
            localStorage.removeItem('excel_auto_rawData');
        }
    }
});

function saveToStorage(name) {
    try {
        localStorage.setItem('excel_auto_rawData', JSON.stringify(rawData));
        localStorage.setItem('excel_auto_fileName', name);
    } catch (e) {
        console.warn("저장 공간 부족");
    }
}

// 3. 파일 업로드 처리
const fileInput = document.getElementById('fileInput');
const uploadZone = document.querySelector('.upload-zone');
const columnModal = document.getElementById('columnModal');
const columnList = document.getElementById('columnList');

function handleFile(file) {
    if (!file) return;

    tempFileName = file.name;
    const fNameLower = tempFileName.toLowerCase();
    if (!['.xlsx', '.xls', '.csv'].some(ext => fNameLower.endsWith(ext))) {
        showStatus("엑셀 또는 CSV 파일만 가능합니다.", "error");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, {type: 'array', cellDates: true});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
            
            let hIdx = -1;
            for (let i = 0; i < Math.min(rows.length, 20); i++) {
                if (!rows[i]) continue;
                const rowStr = JSON.stringify(rows[i]);
                if (rowStr.includes("주문일시") || rowStr.includes("결제금액") || 
                    rowStr.replace(/\s/g, "").includes("주문일시") || 
                    rowStr.replace(/\s/g, "").includes("결제금액")) { 
                    hIdx = i; 
                    break; 
                }
            }
            if (hIdx === -1) {
                throw new Error("필수 컬럼(주문일시 또는 결제금액)을 찾을 수 없습니다. 파일의 첫 행에 정확한 컬럼명이 있는지 확인해주세요.");
            }

            tempHeaders = rows[hIdx].map(h => h != null ? String(h).trim() : h);
            const rawSheetData = XLSX.utils.sheet_to_json(worksheet, {range: hIdx});
            tempSheetData = rawSheetData.map(row => {
                const normalized = {};
                Object.entries(row).forEach(([k, v]) => { normalized[String(k).trim()] = v; });
                return normalized;
            });
            openColumnModal();
        } catch (err) {
            showStatus(err.message, "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

function openColumnModal(checkedCols = null) {
    if (!columnModal || !columnList) return;

    // 표시할 헤더 목록 결정: 신규 업로드면 tempHeaders, 복구 상태면 rawData 키 사용
    const headers = tempHeaders.length > 0
        ? tempHeaders
        : (rawData.length > 0 ? Object.keys(rawData[0]) : []);

    if (headers.length === 0) return;

    columnList.innerHTML = "";
    headers.forEach((header) => {
        if (!header) return;
        const hStr = String(header).trim();
        const isChecked = checkedCols ? checkedCols.includes(hStr) : true;
        const div = document.createElement('label');
        div.className = 'column-item';
        div.innerHTML = `
            <input type="checkbox" name="col" value="${hStr}" ${isChecked ? 'checked' : ''}>
            <span>${hStr}</span>
        `;
        columnList.appendChild(div);
    });

    columnModal.style.display = "block";
}

// 모달 버튼 이벤트
document.querySelector('.close-modal')?.addEventListener('click', () => columnModal.style.display = "none");
document.getElementById('cancelModalBtn')?.addEventListener('click', () => columnModal.style.display = "none");

// 전체 선택/해제
document.getElementById('selectAllBtn')?.addEventListener('click', () => {
    const checkboxes = columnList.querySelectorAll('input[type="checkbox"]');
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !allChecked);
});

// 데이터 반영 (최종 적용)
document.getElementById('applyColumnsBtn')?.addEventListener('click', () => {
    const selectedCols = Array.from(columnList.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
    
    if (selectedCols.length === 0) {
        alert("최소 하나 이상의 컬럼을 선택해야 합니다.");
        return;
    }

    const hasDate = selectedCols.some(c => c.replace(/\s/g, "") === "주문일시");
    if (!hasDate) {
        alert("'주문일시' 컬럼은 필수입니다. 체크 후 다시 시도해주세요.");
        return;
    }

    // 선택된 컬럼만 포함하도록 데이터 필터링
    // tempSheetData가 없으면 (localStorage 복구 상태) 현재 rawData에서 컬럼 제거
    const sourceData = tempSheetData.length > 0 ? tempSheetData : rawData;
    rawData = sourceData.map(row => {
        const filteredRow = {};
        selectedCols.forEach(col => {
            if (row.hasOwnProperty(col)) {
                filteredRow[col] = row[col];
            }
        });
        return filteredRow;
    });

    saveToStorage(tempFileName);
    showStatus(`${rawData.length}건 로드 성공 (컬럼 ${selectedCols.length}개 선택)`, "success");
    processData();
    columnModal.style.display = "none";
});

// 컬럼 편집 버튼 (헤더)
document.getElementById('editColumnsBtn')?.addEventListener('click', () => {
    if (!rawData.length) return;
    const currentCols = Object.keys(rawData[0]);
    openColumnModal(currentCols);
});

// 모달 바깥 클릭 시 닫기
window.addEventListener('click', (event) => {
    if (event.target == columnModal) {
        columnModal.style.display = "none";
    }
});

if (fileInput) {
    fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
}

if (uploadZone) {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        }, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        uploadZone.addEventListener(eventName, () => uploadZone.classList.add('drag-over'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        uploadZone.addEventListener(eventName, () => uploadZone.classList.remove('drag-over'), false);
    });

    uploadZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const file = dt.files[0];
        handleFile(file);
    }, false);
}

// 고객 식별 키: 주문자명 우선 (없으면 주문자ID)
function customerKey(nr) {
    const name = nr.주문자명 != null ? String(nr.주문자명).replace(/\s+/g, '').trim() : '';
    if (name) return name;
    const id = nr.주문자ID != null ? String(nr.주문자ID).replace(/\s+/g, '').trim() : '';
    return id || '익명';
}

// 4. 데이터 집계 로직
function processData() {
    const monthly = {}; const categories = {}; const yearly = {}; const userHistory = {};
    const productCats = {}; const points = {};
    const firstMonth = {};   // 고객별 최초 주문 월

    const fields = detectFields(rawData);
    analyzedData.fieldMeta = fields;

    rawData.forEach(row => {
        const nr = normalizeRow(row);
        const uid = customerKey(nr);
        userHistory[uid] = (userHistory[uid] || 0) + 1;
        if (nr.주문일시) {
            const m = formatDate(nr.주문일시);
            if (!firstMonth[uid] || m < firstMonth[uid]) firstMonth[uid] = m;
        }
    });
    analyzedData.firstMonth = firstMonth;

    rawData.forEach(row => {
        const nr = normalizeRow(row);
        if (!nr.주문일시) return;
        const m = formatDate(nr.주문일시);
        const y = m.substring(0, 4);
        const amt = parseAmount(nr.결제금액);
        const uid = customerKey(nr);
        // 주문횟수 컬럼이 존재하면 숫자로 파싱 (0도 유효값으로 처리)
        // 컬럼 자체가 없을 때만 userHistory 누적값을 대체로 사용
        const rawOrderCnt = nr.주문횟수;
        const orderCnt = (rawOrderCnt !== undefined && rawOrderCnt !== null && rawOrderCnt !== '')
            ? parseFloat(String(rawOrderCnt).replace(/,/g, '')) || 0
            : userHistory[uid];

        const adCost = parseAmount(fields.adCost ? nr[fields.adCost] : nr.광고비);

        if (!monthly[m]) monthly[m] = { totalAmount: 0, totalCount: 0, newAmount: 0, newCount: 0, adCost: 0, users: new Set(), activeUsers: new Set(), firstTimeUsers: new Set(), memberTypes: {} };
        monthly[m].totalAmount += amt; monthly[m].totalCount += 1; monthly[m].users.add(uid);
        // 활동 회원: 해당 월 이전에 주문 이력이 있는 고객만 집계
        if (firstMonth[uid] && firstMonth[uid] < m) monthly[m].activeUsers.add(uid);
        else monthly[m].firstTimeUsers.add(uid);
        monthly[m].adCost += adCost;

        if (!yearly[y]) yearly[y] = { totalAmount: 0, totalCount: 0, newAmount: 0, newCount: 0, adCost: 0 };
        yearly[y].totalAmount += amt; yearly[y].totalCount += 1; yearly[y].adCost += adCost;

        if (orderCnt <= 1) {
            monthly[m].newAmount += amt; monthly[m].newCount += 1;
            yearly[y].newAmount += amt; yearly[y].newCount += 1;
        }

        const type = (nr.회원구분 || '미분류').trim();
        if (!monthly[m].memberTypes[type]) monthly[m].memberTypes[type] = { amount: 0, count: 0, newAmount: 0, newCount: 0 };
        monthly[m].memberTypes[type].amount += amt; monthly[m].memberTypes[type].count += 1;
        if (orderCnt <= 1) { monthly[m].memberTypes[type].newAmount += amt; monthly[m].memberTypes[type].newCount += 1; }

        const cat = ( (nr.상품명 || "").match(/\[(.*?)\]/) || [null, "기타"] )[1];
        categories[cat] = (categories[cat] || 0) + amt;

        // ── 상품 카테고리별 월별 집계 ──
        const pName = fields.productName ? nr[fields.productName] : nr.상품명;
        const pCat = classifyProduct(pName);
        if (!productCats[pCat]) productCats[pCat] = { total: 0, count: 0, months: {}, topItems: {} };
        const pc = productCats[pCat];
        pc.total += amt; pc.count += 1;
        if (!pc.months[m]) pc.months[m] = { amount: 0, count: 0 };
        pc.months[m].amount += amt; pc.months[m].count += 1;
        const itemName = String(pName || '미상').replace(/\s+/g, ' ').trim().slice(0, 40) || '미상';
        if (!pc.topItems[itemName]) pc.topItems[itemName] = { amount: 0, count: 0 };
        pc.topItems[itemName].amount += amt; pc.topItems[itemName].count += 1;

        // ── 적립금 월별 집계 ──
        const used = fields.usedPoint ? parseAmount(nr[fields.usedPoint]) : 0;
        const saved = fields.savedPoint ? parseAmount(nr[fields.savedPoint]) : 0;
        if (!points[m]) points[m] = { used: 0, saved: 0, usedOrders: 0 };
        points[m].used += used;
        points[m].saved += saved;
        if (used > 0) points[m].usedOrders += 1;
        monthly[m].usedPoint = (monthly[m].usedPoint || 0) + used;
        monthly[m].savedPoint = (monthly[m].savedPoint || 0) + saved;
    });

    analyzedData.monthly = monthly; analyzedData.categories = categories; analyzedData.yearly = yearly;
    analyzedData.productCats = productCats; analyzedData.points = points;

    // 주문건수 구간별 고객 분포 (userHistory 기준)
    const orderDist = { '0~1건': 0, '2~5건': 0, '6~10건': 0, '10건 이상': 0 };
    Object.values(userHistory).forEach(cnt => {
        if (cnt <= 1)       orderDist['0~1건']++;
        else if (cnt <= 5)  orderDist['2~5건']++;
        else if (cnt <= 10) orderDist['6~10건']++;
        else                orderDist['10건 이상']++;
    });
    analyzedData.orderDist = orderDist;

    // 결제금액대별 주문 건수 분포
    const amountDist = { '7만원 이하': 0, '7~9만원': 0, '9~12만원': 0, '12~15만원': 0, '15만원 초과': 0 };
    rawData.forEach(row => {
        const nr = normalizeRow(row);
        const amt = parseAmount(nr.결제금액);
        if (amt <= 70000)       amountDist['7만원 이하']++;
        else if (amt <= 90000)  amountDist['7~9만원']++;
        else if (amt <= 120000) amountDist['9~12만원']++;
        else if (amt <= 150000) amountDist['12~15만원']++;
        else                    amountDist['15만원 초과']++;
    });
    analyzedData.amountDist = amountDist;

    // KPI 계산
    const ms = Object.keys(monthly).sort();
    if (ms.length > 0) {
        const cur = monthly[ms[ms.length-1]];
        const pre = ms.length > 1 ? monthly[ms[ms.length-2]] : null;
        analyzedData.mom = {
            revenue: cur.totalAmount,
            revenueGrowth: pre ? (cur.totalAmount - pre.totalAmount) / pre.totalAmount * 100 : 0,
            users: cur.activeUsers.size,
            userGrowth: (pre && pre.activeUsers.size) ? (cur.activeUsers.size - pre.activeUsers.size) / pre.activeUsers.size * 100 : 0,
            totalUsers: cur.users.size,
            firstTimeUsers: cur.firstTimeUsers.size,
            isFirstMonth: ms.length === 1 || ms[ms.length-1] === ms[0],
            newRatio: (cur.newCount / cur.totalCount * 100) || 0,
            aov: (cur.totalAmount / cur.totalCount) || 0,
            adCost: cur.adCost || 0,
            roas: (cur.adCost > 0) ? (cur.newAmount / cur.adCost * 100) : null
        };
    }
    updateUI();
}

// 5. UI 업데이트
function updateUI() {
    const editBtn = document.getElementById('editColumnsBtn');
    if (editBtn) editBtn.style.display = rawData.length > 0 ? '' : 'none';

    const mom = analyzedData.mom || {};
    updateKPICard('kpi-revenue', mom.revenue, mom.revenueGrowth, '원');
    updateKPICard('kpi-users', mom.users, mom.userGrowth, '명');
    const usersSub = document.getElementById('kpi-users-sub');
    if (usersSub) {
        if (!analyzedData.mom || mom.totalUsers === undefined) {
            usersSub.innerHTML = '';
        } else if (mom.isFirstMonth) {
            usersSub.innerHTML = `분석 첫 달이라 이전 주문 이력이 없습니다 (당월 주문자 ${mom.totalUsers.toLocaleString()}명)`;
        } else {
            usersSub.innerHTML = `당월 주문자 ${mom.totalUsers.toLocaleString()}명 중 재주문 ${mom.users.toLocaleString()}명 · 첫 주문 ${mom.firstTimeUsers.toLocaleString()}명`;
        }
    }
    if (document.getElementById('kpi-new-ratio')) document.getElementById('kpi-new-ratio').innerText = (mom.newRatio || 0).toFixed(1) + '%';
    if (document.getElementById('kpi-aov')) document.getElementById('kpi-aov').innerText = Math.round(mom.aov || 0).toLocaleString() + '원';

    // ROAS KPI
    const roasEl = document.getElementById('kpi-roas');
    const roasSubEl = document.getElementById('kpi-roas-sub');
    if (roasEl) {
        if (mom.roas !== null && mom.roas !== undefined) {
            roasEl.innerText = Math.round(mom.roas).toLocaleString() + '%';
            roasEl.style.color = mom.roas >= 100 ? 'var(--success)' : 'var(--danger)';
        } else {
            roasEl.innerText = '-';
            roasEl.style.color = '';
        }
    }
    if (roasSubEl) {
        roasSubEl.innerHTML = (mom.adCost || 0) > 0
            ? `광고비 ${mom.adCost.toLocaleString()}원 | 신규매출 ${(mom.adCost && mom.roas ? Math.round(mom.adCost * mom.roas / 100) : 0).toLocaleString()}원`
            : '광고비 데이터 없음';
    }

    renderCharts();
    renderReportTable();
    renderYearlySummary();
}

function updateKPICard(id, val, growth, unit) {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerText = (val || 0).toLocaleString() + unit;
    const trendEl = el.nextElementSibling;
    if (trendEl && trendEl.classList.contains('kpi-trend')) {
        if (!growth) trendEl.innerHTML = "-";
        else {
            const isUp = growth > 0;
            trendEl.className = `kpi-trend ${isUp ? 'trend-up' : 'trend-down'}`;
            trendEl.innerHTML = `${isUp ? '▲' : '▼'} ${Math.abs(growth).toFixed(1)}% 전월대비`;
        }
    }
}

function showChartEmpty(containerId, canvasId) {
    const container = document.getElementById(containerId)?.closest('.chart-container');
    if (!container) return;
    const canvas = document.getElementById(canvasId);
    if (canvas) canvas.style.display = 'none';
    let msg = container.querySelector('.chart-empty-msg');
    if (!msg) {
        msg = document.createElement('div');
        msg.className = 'chart-empty-msg';
        container.appendChild(msg);
    }
    msg.textContent = '데이터가 없어 차트를 표시할 수 없습니다.';
}

function hideChartEmpty(containerId, canvasId) {
    const container = document.getElementById(containerId)?.closest('.chart-container');
    if (!container) return;
    const canvas = document.getElementById(canvasId);
    if (canvas) canvas.style.display = '';
    const msg = container.querySelector('.chart-empty-msg');
    if (msg) msg.remove();
}

function renderCharts() {
    const ms = Object.keys(analyzedData.monthly).sort();
    const sCanvas = document.getElementById('salesChart');
    const sCtx = sCanvas?.getContext('2d');

    if (!ms.length) {
        if (charts.sales) { charts.sales.destroy(); charts.sales = null; }
        showChartEmpty('salesChart', 'salesChart');
    } else {
        hideChartEmpty('salesChart', 'salesChart');
        if (sCanvas) {
            const minWidthPerMonth = 80;
            const totalNeededWidth = ms.length * minWidthPerMonth;
            const containerWidth = sCanvas.parentElement.clientWidth;
            sCanvas.style.width = totalNeededWidth > containerWidth ? totalNeededWidth + 'px' : '100%';
            sCanvas.style.height = '300px';
        }
        if (sCtx) {
            if (charts.sales) charts.sales.destroy();
            charts.sales = new Chart(sCtx, {
                type: 'bar',
                data: {
                    labels: ms,
                    datasets: [
                        { label: '전체 매출', data: ms.map(m => analyzedData.monthly[m].totalAmount), backgroundColor: '#1a73e8', yAxisID: 'y' },
                        { label: '신규 매출', data: ms.map(m => analyzedData.monthly[m].newAmount), type: 'line', borderColor: '#188038', backgroundColor: 'transparent', yAxisID: 'y' },
                        { label: '신규 ROAS (%)', data: ms.map(m => {
                            const d = analyzedData.monthly[m];
                            return d.adCost > 0 ? Math.round(d.newAmount / d.adCost * 100) : null;
                        }), type: 'line', borderColor: '#f9ab00', backgroundColor: 'transparent', borderDash: [5,3], yAxisID: 'y2', spanGaps: true }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: true,
                            position: 'left',
                            ticks: { callback: function(value) { return value.toLocaleString() + '원'; } }
                        },
                        y2: {
                            beginAtZero: true,
                            position: 'right',
                            grid: { drawOnChartArea: false },
                            ticks: { callback: function(value) { return value.toLocaleString() + '%'; } }
                        },
                        x: {
                            ticks: { autoSkip: false, maxRotation: 0, minRotation: 0 }
                        }
                    },
                    plugins: { legend: { position: 'top' } }
                }
            });
        }
    }

    const cCanvas = document.getElementById('categoryChart');
    const cCtx = cCanvas?.getContext('2d');
    const cData = Object.entries(analyzedData.categories).sort((a,b)=>b[1]-a[1]).slice(0,8);

    if (!cData.length) {
        if (charts.category) { charts.category.destroy(); charts.category = null; }
        showChartEmpty('categoryChart', 'categoryChart');
    } else {
        hideChartEmpty('categoryChart', 'categoryChart');
        if (cCtx) {
            if (charts.category) charts.category.destroy();
            charts.category = new Chart(cCtx, {
                type: 'doughnut',
                data: { labels: cData.map(d=>d[0]), datasets: [{data: cData.map(d=>d[1]), backgroundColor: ['#1a73e8','#34a853','#f9ab00','#ea4335','#a142f4']}] },
                options: { responsive: true, maintainAspectRatio: false }
            });
        }
    }

    renderOrderDistChart();
    renderAmountDistChart();
    renderMemberCompChart();
    renderMomCharts();
    renderCategoryMonthly();
    renderPointSection();
    renderAdCostSection();
}

// ── 공통: 월 라벨 축 설정 (라벨이 겹치지 않도록 일정 간격으로 솎아냄) ──
function monthTicks(ms, extra = {}) {
    const n = Array.isArray(ms) ? ms.length : 0;
    const step = n > 24 ? 3 : (n > 8 ? 2 : 1);
    return Object.assign({
        autoSkip: false,
        maxRotation: 0,
        minRotation: 0,
        font: { size: 11 },
        callback: function(value, index) {
            // 마지막 달은 항상 표시, 그 외에는 step 간격으로만 표시
            if (index === n - 1) return this.getLabelForValue(value);
            if ((n - 1 - index) % step !== 0) return '';
            return this.getLabelForValue(value);
        }
    }, extra);
}

// ── 공통: 숫자 축약 (1,234,567 → 123만) ──
function shortWon(v) {
    const n = Math.round(v || 0);
    if (Math.abs(n) >= 100000000) return (n / 100000000).toFixed(1) + '억';
    if (Math.abs(n) >= 10000) return Math.round(n / 10000).toLocaleString() + '만';
    return n.toLocaleString();
}

function diffBadge(cur, pre, unit, label) {
    const diff = cur - pre;
    const rate = pre ? (diff / pre * 100) : null;
    const up = diff >= 0;
    const color = up ? '#188038' : '#ea4335';
    const rateStr = rate === null ? '전월 없음' : `${up ? '▲' : '▼'} ${Math.abs(rate).toFixed(1)}%`;
    return `<div class="comp-badge" style="border-left:3px solid ${color};">
        <span class="comp-type">${label}</span>
        <span class="comp-cur">${Math.round(cur).toLocaleString()}${unit}</span>
        <span class="comp-diff" style="color:${color};">${up ? '+' : ''}${Math.round(diff).toLocaleString()}${unit} (${rateStr})</span>
    </div>`;
}

// ── 1. 전월 대비 매출금액 / 주문건수 ──
function renderMomCharts() {
    const ms = Object.keys(analyzedData.monthly).sort();
    const summary = document.getElementById('momSummary');

    if (!ms.length) {
        ['momTrend', 'momGrowth'].forEach(k => { if (charts[k]) { charts[k].destroy(); charts[k] = null; } });
        showChartEmpty('momTrendChart', 'momTrendChart');
        showChartEmpty('momGrowthChart', 'momGrowthChart');
        if (summary) summary.innerHTML = '';
        return;
    }
    hideChartEmpty('momTrendChart', 'momTrendChart');
    hideChartEmpty('momGrowthChart', 'momGrowthChart');

    const amounts = ms.map(m => analyzedData.monthly[m].totalAmount);
    const counts  = ms.map(m => analyzedData.monthly[m].totalCount);
    const aovs    = ms.map(m => Math.round(analyzedData.monthly[m].totalAmount / analyzedData.monthly[m].totalCount));

    const amtGrowth = ms.map((m, i) => i === 0 || !amounts[i-1] ? null : parseFloat(((amounts[i] - amounts[i-1]) / amounts[i-1] * 100).toFixed(1)));
    const cntGrowth = ms.map((m, i) => i === 0 || !counts[i-1]  ? null : parseFloat(((counts[i]  - counts[i-1])  / counts[i-1]  * 100).toFixed(1)));

    // 요약 배지 (당월 vs 전월)
    if (summary) {
        const li = ms.length - 1;
        if (ms.length >= 2) {
            summary.innerHTML =
                `<div class="comp-label">기준: <b>${ms[li]}</b> vs 전월 <b>${ms[li-1]}</b></div>` +
                diffBadge(amounts[li], amounts[li-1], '원', '매출액') +
                diffBadge(counts[li], counts[li-1], '건', '주문건수') +
                diffBadge(aovs[li], aovs[li-1], '원', '평균 객단가');
        } else {
            summary.innerHTML = `<p style="color:var(--secondary);font-size:13px;margin:0;">비교할 전월 데이터가 없습니다. 두 달 이상의 데이터를 올려주세요.</p>`;
        }
    }

    const tCtx = document.getElementById('momTrendChart')?.getContext('2d');
    if (tCtx) {
        if (charts.momTrend) charts.momTrend.destroy();
        charts.momTrend = new Chart(tCtx, {
            type: 'bar',
            data: {
                labels: ms,
                datasets: [
                    { label: '매출액', data: amounts, backgroundColor: '#1a73e8', borderRadius: 4, yAxisID: 'y', order: 2 },
                    { label: '주문건수', data: counts, type: 'line', borderColor: '#ea4335', backgroundColor: '#ea4335',
                      pointRadius: 4, borderWidth: 2, yAxisID: 'y1', order: 1 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    y: { beginAtZero: true, position: 'left', ticks: { callback: v => shortWon(v) },
                         title: { display: true, text: '매출액 (원)', color: '#1a73e8', font: { size: 11 } } },
                    y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false },
                          ticks: { callback: v => v.toLocaleString() + '건' },
                          title: { display: true, text: '주문건수', color: '#ea4335', font: { size: 11 } } },
                    x: { ticks: monthTicks(ms) }
                },
                plugins: {
                    legend: { position: 'top' },
                    tooltip: { callbacks: { label: c => c.dataset.yAxisID === 'y1'
                        ? ` 주문건수: ${c.parsed.y.toLocaleString()}건`
                        : ` 매출액: ${c.parsed.y.toLocaleString()}원` } }
                }
            }
        });
    }

    const gCtx = document.getElementById('momGrowthChart')?.getContext('2d');
    if (gCtx) {
        if (charts.momGrowth) charts.momGrowth.destroy();
        charts.momGrowth = new Chart(gCtx, {
            type: 'bar',
            data: {
                labels: ms,
                datasets: [
                    { label: '매출액 증감률', data: amtGrowth, borderRadius: 4,
                      backgroundColor: amtGrowth.map(v => v === null ? '#eee' : v >= 0 ? '#1a73e8' : '#f6aea9') },
                    { label: '주문건수 증감률', data: cntGrowth, borderRadius: 4,
                      backgroundColor: cntGrowth.map(v => v === null ? '#eee' : v >= 0 ? '#34a853' : '#ea4335') }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    y: { ticks: { callback: v => v + '%' }, grid: { color: ctx => ctx.tick.value === 0 ? '#5f6368' : 'rgba(0,0,0,0.06)' } },
                    x: { ticks: monthTicks(ms) }
                },
                plugins: {
                    legend: {
                        position: 'top',
                        labels: {
                            // 막대 색이 값마다 달라 범례가 회색으로 잡히는 것을 방지
                            generateLabels: (chart) => ['#1a73e8', '#34a853'].map((c, i) => ({
                                text: chart.data.datasets[i].label,
                                fillStyle: c,
                                strokeStyle: c,
                                lineWidth: 0,
                                hidden: !chart.isDatasetVisible(i),
                                datasetIndex: i
                            }))
                        }
                    },
                    tooltip: { callbacks: { label: c => c.parsed.y === null ? ' 전월 데이터 없음'
                        : ` ${c.dataset.label}: ${c.parsed.y >= 0 ? '+' : ''}${c.parsed.y}%` } }
                }
            }
        });
    }
}

// ── 2. 상품 카테고리별 월별 매출 ──
function renderCategoryMonthly() {
    const ms = Object.keys(analyzedData.monthly).sort();
    const cats = Object.entries(analyzedData.productCats).sort((a, b) => b[1].total - a[1].total);
    const tbody = document.getElementById('catMonthlyTbody');
    const thead = document.getElementById('catMonthlyHead');

    if (!ms.length || !cats.length) {
        if (charts.catMonthly) { charts.catMonthly.destroy(); charts.catMonthly = null; }
        showChartEmpty('catMonthlyChart', 'catMonthlyChart');
        if (tbody) tbody.innerHTML = `<tr><td class="text-center">데이터를 업로드하면 카테고리별 실적이 표시됩니다.</td></tr>`;
        return;
    }
    hideChartEmpty('catMonthlyChart', 'catMonthlyChart');

    const isAmt = catChartMode === 'amount';
    const ctx = document.getElementById('catMonthlyChart')?.getContext('2d');
    if (ctx) {
        if (charts.catMonthly) charts.catMonthly.destroy();
        charts.catMonthly = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ms,
                datasets: cats.map(([name, d]) => ({
                    label: name,
                    data: ms.map(m => d.months[m] ? (isAmt ? d.months[m].amount : d.months[m].count) : 0),
                    backgroundColor: categoryColor(name),
                    borderRadius: 3
                }))
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    x: { stacked: true, ticks: monthTicks(ms) },
                    y: { stacked: true, beginAtZero: true,
                         ticks: { callback: v => isAmt ? shortWon(v) : v.toLocaleString() + '건' } }
                },
                plugins: {
                    legend: { position: 'bottom', labels: { boxWidth: 12, padding: 10, font: { size: 11 } } },
                    tooltip: {
                        callbacks: {
                            label: c => ` ${c.dataset.label}: ${c.parsed.y.toLocaleString()}${isAmt ? '원' : '건'}`,
                            footer: items => {
                                const sum = items.reduce((s, i) => s + i.parsed.y, 0);
                                return `합계: ${sum.toLocaleString()}${isAmt ? '원' : '건'}`;
                            }
                        }
                    }
                }
            }
        });
    }

    // 표: 행=카테고리, 열=월
    if (thead) {
        thead.innerHTML = `<tr><th>상품 카테고리</th>${ms.map(m => `<th class="text-right">${m}</th>`).join('')}<th class="text-right">합계</th><th class="text-right">비중</th></tr>`;
    }
    const grandTotal = cats.reduce((s, [, d]) => s + d.total, 0);
    if (tbody) {
        const rows = cats.map(([name, d]) => {
            const top = Object.entries(d.topItems).sort((a, b) => b[1].amount - a[1].amount)[0];
            return `<tr>
                <td><span class="cat-dot" style="background:${categoryColor(name)}"></span><b>${name}</b>
                    <div class="cat-top">대표상품: ${top ? top[0] : '-'}</div></td>
                ${ms.map(m => `<td class="text-right">${d.months[m] ? d.months[m].amount.toLocaleString() + '원<div class="cat-sub">' + d.months[m].count + '건</div>' : '<span style="color:#ccc">-</span>'}</td>`).join('')}
                <td class="text-right"><b>${d.total.toLocaleString()}원</b><div class="cat-sub">${d.count}건</div></td>
                <td class="text-right">${(d.total / grandTotal * 100).toFixed(1)}%</td>
            </tr>`;
        }).join('');
        const totalRow = `<tr class="subtotal-row">
            <td>월 합계</td>
            ${ms.map(m => {
                const sum = cats.reduce((s, [, d]) => s + (d.months[m]?.amount || 0), 0);
                return `<td class="text-right">${sum.toLocaleString()}원</td>`;
            }).join('')}
            <td class="text-right">${grandTotal.toLocaleString()}원</td>
            <td class="text-right">100%</td>
        </tr>`;
        tbody.innerHTML = rows + totalRow;
    }
}

// ── 3. 적립금 사용 분석 ──
function renderPointSection() {
    const ms = Object.keys(analyzedData.monthly).sort();
    const meta = analyzedData.fieldMeta || {};
    const notice = document.getElementById('pointNotice');
    const summary = document.getElementById('pointSummary');
    const tbody = document.getElementById('pointTbody');

    const hasPoint = !!meta.usedPoint || !!meta.savedPoint;

    if (notice) {
        if (!hasPoint) {
            notice.innerHTML = `⚠️ 적립금 관련 컬럼을 찾지 못했습니다. 업로드 시 <b>'사용적립금'</b> 같은 컬럼을 함께 선택해 주세요.`;
            notice.className = 'field-notice warn';
        } else {
            const parts = [];
            if (meta.usedPoint) parts.push(`사용 적립금 = <b>${meta.usedPoint}</b>${meta.usedPointGuessed ? ' <span style="color:#d93025">(자동 추정)</span>' : ''}`);
            if (meta.savedPoint) parts.push(`적립(지급) = <b>${meta.savedPoint}</b>`);
            notice.innerHTML = `✅ 인식된 컬럼 — ${parts.join(' / ')}`;
            notice.className = 'field-notice ok';
        }
    }

    if (!ms.length || !hasPoint) {
        if (charts.point) { charts.point.destroy(); charts.point = null; }
        showChartEmpty('pointChart', 'pointChart');
        if (summary) summary.innerHTML = '';
        if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="text-center">적립금 데이터가 없습니다.</td></tr>`;
        return;
    }
    hideChartEmpty('pointChart', 'pointChart');

    const used  = ms.map(m => analyzedData.points[m]?.used || 0);
    const saved = ms.map(m => analyzedData.points[m]?.saved || 0);
    const orders = ms.map(m => analyzedData.points[m]?.usedOrders || 0);
    const usageRate = ms.map((m, i) => {
        const rev = analyzedData.monthly[m].totalAmount;
        return rev ? parseFloat((used[i] / rev * 100).toFixed(2)) : 0;
    });

    const li = ms.length - 1;
    const totalUsed = used.reduce((s, v) => s + v, 0);
    if (summary) {
        summary.innerHTML =
            `<div class="comp-badge" style="border-left:3px solid #1a73e8;">
                <span class="comp-type">누적 사용 적립금</span>
                <span class="comp-cur">${totalUsed.toLocaleString()}원</span>
                <span class="comp-diff" style="color:var(--secondary);">${ms[0]} ~ ${ms[li]}</span>
            </div>` +
            diffBadge(used[li], ms.length > 1 ? used[li-1] : 0, '원', `당월(${ms[li]}) 사용액`) +
            `<div class="comp-badge" style="border-left:3px solid #34a853;">
                <span class="comp-type">당월 매출 대비 사용률</span>
                <span class="comp-cur">${usageRate[li]}%</span>
                <span class="comp-diff" style="color:var(--secondary);">사용 주문 ${orders[li].toLocaleString()}건</span>
            </div>` +
            `<div class="comp-badge" style="border-left:3px solid #f9ab00;">
                <span class="comp-type">당월 건당 평균 사용액</span>
                <span class="comp-cur">${orders[li] ? Math.round(used[li] / orders[li]).toLocaleString() : 0}원</span>
                <span class="comp-diff" style="color:var(--secondary);">적립금 사용 주문 기준</span>
            </div>`;
    }

    const ctx = document.getElementById('pointChart')?.getContext('2d');
    if (ctx) {
        if (charts.point) charts.point.destroy();
        const ds = [{ label: '사용 적립금', data: used, backgroundColor: '#1a73e8', borderRadius: 4, yAxisID: 'y' }];
        if (meta.savedPoint) ds.push({ label: '적립(지급) 금액', data: saved, backgroundColor: '#a8d8ea', borderRadius: 4, yAxisID: 'y' });
        ds.push({ label: '매출 대비 사용률', data: usageRate, type: 'line', borderColor: '#f9ab00',
                  backgroundColor: 'transparent', borderDash: [5, 3], pointRadius: 4, yAxisID: 'y1' });

        charts.point = new Chart(ctx, {
            type: 'bar',
            data: { labels: ms, datasets: ds },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    y: { beginAtZero: true, ticks: { callback: v => shortWon(v) },
                         title: { display: true, text: '적립금 (원)', color: '#5f6368', font: { size: 11 } } },
                    y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false },
                          ticks: { callback: v => v + '%' },
                          title: { display: true, text: '매출 대비 사용률', color: '#f9ab00', font: { size: 11 } } },
                    x: { ticks: monthTicks(ms) }
                },
                plugins: {
                    legend: { position: 'top' },
                    tooltip: { callbacks: { label: c => c.dataset.yAxisID === 'y1'
                        ? ` 매출 대비 사용률: ${c.parsed.y}%`
                        : ` ${c.dataset.label}: ${c.parsed.y.toLocaleString()}원` } }
                }
            }
        });
    }

    if (tbody) {
        tbody.innerHTML = ms.slice().reverse().map(m => {
            const i = ms.indexOf(m);
            const rev = analyzedData.monthly[m].totalAmount;
            const prev = i > 0 ? used[i-1] : null;
            const diff = prev === null ? null : used[i] - prev;
            const rate = prev ? (diff / prev * 100) : null;
            const diffStr = diff === null ? '<span style="color:#ccc">-</span>'
                : `<span style="color:${diff >= 0 ? '#188038' : '#ea4335'};font-weight:600;">${diff >= 0 ? '+' : ''}${diff.toLocaleString()}원${rate !== null ? ` (${diff >= 0 ? '▲' : '▼'}${Math.abs(rate).toFixed(1)}%)` : ''}</span>`;
            return `<tr>
                <td class="text-center"><b>${m}</b></td>
                <td class="text-right">${used[i].toLocaleString()}원</td>
                <td class="text-right">${diffStr}</td>
                <td class="text-right">${orders[i].toLocaleString()}건</td>
                <td class="text-right">${orders[i] ? Math.round(used[i]/orders[i]).toLocaleString() : 0}원</td>
                <td class="text-right">${rev ? (used[i]/rev*100).toFixed(2) : 0}%</td>
            </tr>`;
        }).join('');
    }
}

// ── 4. 광고비 월별 분석 ──
function renderAdCostSection() {
    const ms = Object.keys(analyzedData.monthly).sort();
    const summary = document.getElementById('adSummary');
    const notice = document.getElementById('adNotice');
    const meta = analyzedData.fieldMeta || {};
    const adCosts = ms.map(m => analyzedData.monthly[m].adCost || 0);
    const hasAd = adCosts.some(v => v > 0);

    if (notice) {
        notice.innerHTML = hasAd
            ? `✅ 인식된 컬럼 — 광고비 = <b>${meta.adCost || '광고비'}</b> · 행 단위 값을 월별로 합산합니다.`
            : `⚠️ 광고비 컬럼을 찾지 못했습니다. 업로드 시 <b>'광고비'</b> 컬럼을 함께 선택해 주세요.`;
        notice.className = 'field-notice ' + (hasAd ? 'ok' : 'warn');
    }

    if (!ms.length || !hasAd) {
        if (charts.adCost) { charts.adCost.destroy(); charts.adCost = null; }
        showChartEmpty('adCostChart', 'adCostChart');
        if (summary) summary.innerHTML = '';
        return;
    }
    hideChartEmpty('adCostChart', 'adCostChart');

    const newAmounts = ms.map(m => analyzedData.monthly[m].newAmount);
    const newCounts = ms.map(m => analyzedData.monthly[m].newCount);
    const roas = ms.map((m, i) => adCosts[i] > 0 ? Math.round(newAmounts[i] / adCosts[i] * 100) : null);
    const adRatio = ms.map((m, i) => {
        const rev = analyzedData.monthly[m].totalAmount;
        return rev ? parseFloat((adCosts[i] / rev * 100).toFixed(1)) : 0;
    });

    const li = ms.length - 1;
    const total = adCosts.reduce((s, v) => s + v, 0);
    if (summary) {
        const cpa = newCounts[li] ? Math.round(adCosts[li] / newCounts[li]) : 0;
        summary.innerHTML =
            `<div class="comp-badge" style="border-left:3px solid #5f6368;">
                <span class="comp-type">누적 광고비</span>
                <span class="comp-cur">${total.toLocaleString()}원</span>
                <span class="comp-diff" style="color:var(--secondary);">${ms[0]} ~ ${ms[li]}</span>
            </div>` +
            diffBadge(adCosts[li], ms.length > 1 ? adCosts[li-1] : 0, '원', `당월(${ms[li]}) 광고비`) +
            `<div class="comp-badge" style="border-left:3px solid #f9ab00;">
                <span class="comp-type">당월 매출 대비 광고비</span>
                <span class="comp-cur">${adRatio[li]}%</span>
                <span class="comp-diff" style="color:var(--secondary);">신규 ROAS ${roas[li] !== null ? roas[li].toLocaleString() + '%' : '-'}</span>
            </div>` +
            `<div class="comp-badge" style="border-left:3px solid #ea4335;">
                <span class="comp-type">당월 신규 주문 1건당 광고비</span>
                <span class="comp-cur">${cpa.toLocaleString()}원</span>
                <span class="comp-diff" style="color:var(--secondary);">신규 주문 ${newCounts[li].toLocaleString()}건</span>
            </div>`;
    }

    const ctx = document.getElementById('adCostChart')?.getContext('2d');
    if (ctx) {
        if (charts.adCost) charts.adCost.destroy();
        charts.adCost = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ms,
                datasets: [
                    { label: '광고비', data: adCosts, backgroundColor: '#f9ab00', borderRadius: 4, yAxisID: 'y', order: 3 },
                    { label: '신규 매출', data: newAmounts, backgroundColor: '#a8e6ba', borderRadius: 4, yAxisID: 'y', order: 2 },
                    { label: '신규 ROAS (%)', data: roas, type: 'line', borderColor: '#188038',
                      backgroundColor: 'transparent', pointRadius: 4, borderWidth: 2, yAxisID: 'y1', spanGaps: true, order: 1 }
                ]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    y: { beginAtZero: true, ticks: { callback: v => shortWon(v) },
                         title: { display: true, text: '금액 (원)', color: '#5f6368', font: { size: 11 } } },
                    y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false },
                          ticks: { callback: v => v.toLocaleString() + '%' },
                          title: { display: true, text: '신규 ROAS', color: '#188038', font: { size: 11 } } },
                    x: { ticks: monthTicks(ms) }
                },
                plugins: {
                    legend: { position: 'top' },
                    tooltip: { callbacks: { label: c => c.dataset.yAxisID === 'y1'
                        ? (c.parsed.y === null ? ' ROAS: 광고비 없음' : ` 신규 ROAS: ${c.parsed.y.toLocaleString()}%`)
                        : ` ${c.dataset.label}: ${c.parsed.y.toLocaleString()}원` } }
                }
            }
        });
    }
}

function renderOrderDistChart() {
    const canvas = document.getElementById('orderDistChart');
    const ctx = canvas?.getContext('2d');
    const data = analyzedData.orderDist || {};
    const labels = Object.keys(data);
    const values = Object.values(data);
    const total = values.reduce((s, v) => s + v, 0);

    if (!total) {
        if (charts.orderDist) { charts.orderDist.destroy(); charts.orderDist = null; }
        showChartEmpty('orderDistChart', 'orderDistChart');
        return;
    }
    hideChartEmpty('orderDistChart', 'orderDistChart');
    if (ctx) {
        if (charts.orderDist) charts.orderDist.destroy();
        charts.orderDist = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: labels.map((l, i) => `${l} (${values[i]}명, ${(values[i]/total*100).toFixed(1)}%)`),
                datasets: [{
                    data: values,
                    backgroundColor: ['#1a73e8', '#34a853', '#f9ab00', '#ea4335'],
                    borderWidth: 2,
                    borderColor: '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 12 }, padding: 12 } },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                const val = ctx.parsed;
                                const pct = (val / total * 100).toFixed(1);
                                return ` ${val.toLocaleString()}명 (${pct}%)`;
                            }
                        }
                    }
                }
            }
        });
    }
}

function renderAmountDistChart() {
    const canvas = document.getElementById('amountDistChart');
    const ctx = canvas?.getContext('2d');
    const data = analyzedData.amountDist || {};
    const labels = Object.keys(data);
    const values = Object.values(data);
    const total = values.reduce((s, v) => s + v, 0);

    if (!total) {
        if (charts.amountDist) { charts.amountDist.destroy(); charts.amountDist = null; }
        showChartEmpty('amountDistChart', 'amountDistChart');
        return;
    }
    hideChartEmpty('amountDistChart', 'amountDistChart');
    if (ctx) {
        if (charts.amountDist) charts.amountDist.destroy();
        charts.amountDist = new Chart(ctx, {
            type: 'pie',
            data: {
                labels: labels.map((l, i) => `${l} (${values[i]}건, ${(values[i]/total*100).toFixed(1)}%)`),
                datasets: [{
                    data: values,
                    backgroundColor: ['#a8d8ea', '#1a73e8', '#34a853', '#f9ab00', '#ea4335'],
                    borderWidth: 2,
                    borderColor: '#fff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'bottom', labels: { font: { size: 12 }, padding: 12 } },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                const val = ctx.parsed;
                                const pct = (val / total * 100).toFixed(1);
                                return ` ${val.toLocaleString()}건 (${pct}%)`;
                            }
                        }
                    }
                }
            }
        });
    }
}

function renderMemberCompChart() {
    const canvas = document.getElementById('memberCompChart');
    const ctx = canvas?.getContext('2d');
    const ms = Object.keys(analyzedData.monthly).sort();

    if (ms.length < 1) {
        if (charts.memberComp) { charts.memberComp.destroy(); charts.memberComp = null; }
        showChartEmpty('memberCompChart', 'memberCompChart');
        return;
    }
    hideChartEmpty('memberCompChart', 'memberCompChart');

    const curKey = ms[ms.length - 1];
    const preKey = ms.length >= 2 ? ms[ms.length - 2] : null;
    const curData = analyzedData.monthly[curKey];
    const preData = preKey ? analyzedData.monthly[preKey] : null;

    // 두 달 모두에서 회원유형 수집
    const typeSet = new Set();
    Object.keys(curData.memberTypes).forEach(t => typeSet.add(t));
    if (preData) Object.keys(preData.memberTypes).forEach(t => typeSet.add(t));
    const types = Array.from(typeSet).sort();

    const curAmounts = types.map(t => curData.memberTypes[t]?.amount || 0);
    const preAmounts = types.map(t => preData ? (preData.memberTypes[t]?.amount || 0) : 0);
    const curCounts  = types.map(t => curData.memberTypes[t]?.count  || 0);

    const growthRates = types.map((t, i) => {
        if (!preData || preAmounts[i] === 0) return null;
        return parseFloat(((curAmounts[i] - preAmounts[i]) / preAmounts[i] * 100).toFixed(1));
    });

    const COLOR_CUR = ['#1a73e8', '#34a853', '#f9ab00', '#a142f4', '#ea4335'];
    const COLOR_PRE = ['#9ec5fd', '#a8e6ba', '#fde9a2', '#d8b4fe', '#f4b8b4'];

    const datasets = [];
    if (preData) {
        datasets.push({
            label: `전월 (${preKey}) 매출`,
            data: preAmounts,
            backgroundColor: types.map((_, i) => COLOR_PRE[i % COLOR_PRE.length]),
            borderRadius: 4,
            yAxisID: 'y'
        });
    }
    datasets.push({
        label: `당월 (${curKey}) 매출`,
        data: curAmounts,
        backgroundColor: types.map((_, i) => COLOR_CUR[i % COLOR_CUR.length]),
        borderRadius: 4,
        yAxisID: 'y'
    });
    if (preData) {
        datasets.push({
            label: '전월대비 증감률',
            data: growthRates,
            type: 'line',
            borderColor: '#ea4335',
            backgroundColor: 'transparent',
            pointBackgroundColor: growthRates.map(r => r === null ? 'transparent' : r >= 0 ? '#188038' : '#ea4335'),
            pointRadius: 7,
            pointHoverRadius: 9,
            borderWidth: 2,
            borderDash: [5, 3],
            yAxisID: 'y2',
            spanGaps: false
        });
    }

    // 증감 배지 렌더링
    const summary = document.getElementById('memberCompSummary');
    if (summary) {
        if (preData) {
            summary.innerHTML = types.map((t, i) => {
                const diff = curAmounts[i] - preAmounts[i];
                const rate = growthRates[i];
                const diffColor = diff >= 0 ? '#188038' : '#ea4335';
                const arrow = diff >= 0 ? '▲' : '▼';
                const rateStr = rate !== null ? `${arrow} ${Math.abs(rate)}%` : '-';
                return `<div class="comp-badge" style="border-left: 3px solid ${COLOR_CUR[i % COLOR_CUR.length]};">
                    <span class="comp-type">${t}</span>
                    <span class="comp-cur">${curAmounts[i].toLocaleString()}원</span>
                    <span class="comp-diff" style="color:${diffColor};">${diff >= 0 ? '+' : ''}${diff.toLocaleString()}원 (${rateStr})</span>
                    <span class="comp-count" style="color:#5f6368; font-size:11px;">주문 ${curCounts[i]}건</span>
                </div>`;
            }).join('');
        } else {
            summary.innerHTML = `<p style="color:var(--secondary); font-size:13px; margin:0;">전월 데이터가 없어 증감 비교를 표시할 수 없습니다.</p>`;
        }
    }

    if (ctx) {
        if (charts.memberComp) charts.memberComp.destroy();
        charts.memberComp = new Chart(ctx, {
            type: 'bar',
            data: { labels: types, datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                scales: {
                    y: {
                        beginAtZero: true,
                        position: 'left',
                        ticks: { callback: v => v.toLocaleString() + '원' },
                        title: { display: true, text: '매출액 (원)', color: '#5f6368', font: { size: 11 } }
                    },
                    y2: {
                        beginAtZero: false,
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        ticks: { callback: v => v + '%' },
                        title: { display: true, text: '증감률 (%)', color: '#ea4335', font: { size: 11 } }
                    },
                    x: { ticks: { font: { size: 13, weight: '600' } } }
                },
                plugins: {
                    legend: { position: 'top' },
                    tooltip: {
                        callbacks: {
                            label: function(ctx) {
                                if (ctx.dataset.yAxisID === 'y2') {
                                    const v = ctx.parsed.y;
                                    return v === null ? '' : ` 증감률: ${v >= 0 ? '+' : ''}${v}%`;
                                }
                                return ` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString()}원`;
                            }
                        }
                    }
                }
            }
        });
    }
}

function renderYearlySummary() {
    const tb = document.getElementById('yearlyTbody');
    if (!tb) return;
    const ys = Object.keys(analyzedData.yearly).sort().reverse();
    tb.innerHTML = ys.map(y => {
        const d = analyzedData.yearly[y];
        const roas = d.adCost > 0 ? Math.round(d.newAmount / d.adCost * 100) : null;
        const roasStr = roas !== null
            ? `<span style="color:${roas >= 100 ? 'var(--success)' : 'var(--danger)'}; font-weight:700;">${roas.toLocaleString()}%</span>`
            : '<span style="color:#ccc">-</span>';
        return `<tr>
            <td class="text-center"><b>${y}년</b></td>
            <td class="text-right">${d.totalAmount.toLocaleString()}원</td>
            <td class="text-right">${d.newAmount.toLocaleString()}원</td>
            <td class="text-right">${(d.newAmount/d.totalAmount*100).toFixed(1)}%</td>
            <td class="text-right">${d.totalCount.toLocaleString()}건</td>
            <td class="text-right">${Math.round(d.totalAmount/d.totalCount).toLocaleString()}원</td>
            <td class="text-right">${d.adCost > 0 ? d.adCost.toLocaleString() + '원' : '-'}</td>
            <td class="text-right">${roasStr}</td>
        </tr>`;
    }).join('');
    document.getElementById('yearly-summary-section').style.display = ys.length ? 'block' : 'none';
}

function renderReportTable() {
    const tb = document.getElementById('reportTbody');
    if (!tb) return;
    const ms = Object.keys(analyzedData.monthly).sort().reverse();
    if (!ms.length) { tb.innerHTML = "<tr><td colspan='8' class='text-center'>데이터가 없습니다.</td></tr>"; return; }

    tb.innerHTML = ms.map(m => {
        const data = analyzedData.monthly[m];
        const ts = Object.entries(data.memberTypes);
        const monthAdCost = data.adCost || 0;
        const monthRoas = monthAdCost > 0 ? Math.round(data.newAmount / monthAdCost * 100) : null;
        const roasCell = monthRoas !== null
            ? `<span style="color:${monthRoas >= 100 ? 'var(--success)' : 'var(--danger)'}; font-weight:700;">${monthRoas.toLocaleString()}%</span>`
            : '<span style="color:#ccc">-</span>';
        let sumA = 0; let sumC = 0; let sumNew = 0;
        const rows = ts.map(([type, s], i) => {
            sumA += s.amount; sumC += s.count; sumNew += s.newAmount;
            return `<tr>
                ${i === 0 ? `<td class="text-center" rowspan="${ts.length + 1}" style="background:#fff;font-weight:bold;">${m}</td>` : ''}
                <td>${type}</td>
                <td class="text-right">${s.count.toLocaleString()}</td>
                <td class="text-right">${s.amount.toLocaleString()}원</td>
                <td class="text-right">${s.newAmount.toLocaleString()}원</td>
                <td class="text-right">${i === 0 ? (monthAdCost > 0 ? monthAdCost.toLocaleString() + '원' : '-') : ''}</td>
                <td class="text-right">${i === 0 ? roasCell : ''}</td>
                <td class="text-right">${Math.round(s.amount/s.count).toLocaleString()}원</td>
            </tr>`;
        }).join('');
        return rows + `<tr class="subtotal-row">
            <td style="background:#f1f3f4;font-weight:bold;">${m} 합계</td>
            <td class="text-right">${sumC.toLocaleString()}</td>
            <td class="text-right">${sumA.toLocaleString()}원</td>
            <td class="text-right">${sumNew.toLocaleString()}원</td>
            <td class="text-right">${monthAdCost > 0 ? monthAdCost.toLocaleString() + '원' : '-'}</td>
            <td class="text-right">${roasCell}</td>
            <td class="text-right">${Math.round(sumA/sumC).toLocaleString()}원</td>
        </tr>`;
    }).join('');
}

// 카테고리 차트 금액/건수 전환
document.querySelectorAll('.toggle-btn[data-mode]').forEach(btn => {
    btn.addEventListener('click', () => {
        catChartMode = btn.getAttribute('data-mode');
        document.querySelectorAll('.toggle-btn[data-mode]').forEach(b => b.classList.toggle('active', b === btn));
        renderCategoryMonthly();
    });
});

// 초기화 및 기타 버튼
document.getElementById('clearBtn')?.addEventListener('click', () => { if(confirm("초기화할까요?")){ localStorage.clear(); location.reload(); } });
document.getElementById('downloadBtn')?.addEventListener('click', () => {
    if(!rawData.length) return;
    const wb = XLSX.utils.book_new();

    // 1. 원본 데이터 시트 (상단에 총건수 요약행 추가)
    const headers = Object.keys(rawData[0]);
    const dataRows = rawData.map(r => headers.map(h => r[h]));
    const ws1 = XLSX.utils.aoa_to_sheet([[rawData.length + '건', ...Array(headers.length - 1).fill(null)], headers, ...dataRows]);
    XLSX.utils.book_append_sheet(wb, ws1, "1. 원본 데이터");

    // 2. 매출 통계 시트 (신규/기존 분리, 8컬럼)
    const today = new Date();
    const dateStr = `${today.getFullYear()}. ${today.getMonth()+1}. ${today.getDate()}.`;
    const mRows = [
        ['매출 통계 보고서', null, null, null, null, null, null, null],
        ['기준일자:', dateStr, null, null, null, null, null, null],
        [],
        ['월', '회원 유형', '전체 건수', '전체 매출액', '신규 건수 (0~1회)', '신규 매출액', '기존 건수', '기존 매출액']
    ];
    Object.keys(analyzedData.monthly).sort().forEach(m => {
        Object.entries(analyzedData.monthly[m].memberTypes).forEach(([type, s]) => {
            mRows.push([m, type, s.count, s.amount, s.newCount, s.newAmount, s.count - s.newCount, s.amount - s.newAmount]);
        });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mRows), "2. 매출 통계");

    // 3. 카테고리 비중 시트
    const catEntries = Object.entries(analyzedData.categories).sort((a, b) => b[1] - a[1]);
    const catTotal = catEntries.reduce((s, [, v]) => s + v, 0);
    const cRows = [['카테고리', '매출액', '비중(%)']];
    catEntries.forEach(([cat, amt]) => cRows.push([cat, amt, parseFloat((amt / catTotal * 100).toFixed(1))]));
    cRows.push(['합계', catTotal, 100]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cRows), "3. 카테고리 비중");

    const months = Object.keys(analyzedData.monthly).sort();

    // 4. 전월 대비 매출/건수 시트
    const momRows = [['월', '매출액', '주문건수', '평균 객단가', '매출 증감액', '매출 증감률(%)', '건수 증감', '건수 증감률(%)']];
    months.forEach((m, i) => {
        const c = analyzedData.monthly[m];
        const p = i > 0 ? analyzedData.monthly[months[i-1]] : null;
        momRows.push([
            m, c.totalAmount, c.totalCount, Math.round(c.totalAmount / c.totalCount),
            p ? c.totalAmount - p.totalAmount : null,
            p && p.totalAmount ? parseFloat(((c.totalAmount - p.totalAmount) / p.totalAmount * 100).toFixed(1)) : null,
            p ? c.totalCount - p.totalCount : null,
            p && p.totalCount ? parseFloat(((c.totalCount - p.totalCount) / p.totalCount * 100).toFixed(1)) : null
        ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(momRows), "4. 전월대비");

    // 5. 상품 카테고리별 월별 매출
    const pcEntries = Object.entries(analyzedData.productCats).sort((a, b) => b[1].total - a[1].total);
    const pcRows = [['상품 카테고리', ...months.flatMap(m => [`${m} 매출액`, `${m} 건수`]), '총 매출액', '총 건수', '비중(%)']];
    const pcGrand = pcEntries.reduce((s, [, d]) => s + d.total, 0);
    pcEntries.forEach(([name, d]) => {
        pcRows.push([name, ...months.flatMap(m => [d.months[m]?.amount || 0, d.months[m]?.count || 0]),
            d.total, d.count, pcGrand ? parseFloat((d.total / pcGrand * 100).toFixed(1)) : 0]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pcRows), "5. 카테고리별 월매출");

    // 6. 적립금 사용 현황
    const ptRows = [['월', '사용 적립금', '적립(지급) 금액', '사용 주문건수', '건당 평균 사용액', '매출 대비 사용률(%)']];
    months.forEach(m => {
        const p = analyzedData.points[m] || { used: 0, saved: 0, usedOrders: 0 };
        const rev = analyzedData.monthly[m].totalAmount;
        ptRows.push([m, p.used, p.saved, p.usedOrders,
            p.usedOrders ? Math.round(p.used / p.usedOrders) : 0,
            rev ? parseFloat((p.used / rev * 100).toFixed(2)) : 0]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ptRows), "6. 적립금 사용");

    // 7. 광고비 집행 현황
    const adRows = [['월', '광고비', '총 매출액', '신규 매출액', '신규 주문건수', '매출 대비 광고비(%)', '신규 ROAS(%)', 'CPA(신규 1건당 광고비)']];
    months.forEach(m => {
        const c = analyzedData.monthly[m];
        const ad = c.adCost || 0;
        adRows.push([m, ad, c.totalAmount, c.newAmount, c.newCount,
            c.totalAmount ? parseFloat((ad / c.totalAmount * 100).toFixed(1)) : 0,
            ad > 0 ? Math.round(c.newAmount / ad * 100) : null,
            c.newCount ? Math.round(ad / c.newCount) : 0]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(adRows), "7. 광고비 집행");

    const dateFile = today.toISOString().slice(0, 10);
    XLSX.writeFile(wb, `매출보고서_${dateFile}.xlsx`);
});

function copyFormula(id) {
    const txt = document.getElementById(id)?.innerText;
    if(txt) navigator.clipboard.writeText(txt).then(() => alert("복사되었습니다."));
}

// ── 인쇄: canvas를 이미지로 교체 후 인쇄 ──
let _printBackups = [];

function setPrintMeta() {
    const el = document.getElementById('print-meta');
    if (el) {
        const now = new Date();
        const dateStr = `${now.getFullYear()}년 ${now.getMonth()+1}월 ${now.getDate()}일`;
        const ms = Object.keys(analyzedData.monthly).sort();
        const rangeStr = ms.length ? `분석 기간: ${ms[0]} ~ ${ms[ms.length-1]}` : '';
        el.innerHTML = `출력일: ${dateStr}${rangeStr ? '<br>' + rangeStr : ''}`;
    }
}

function replaceCanvasWithImage() {
    _printBackups = [];
    const canvasIds = ['salesChart', 'categoryChart', 'orderDistChart', 'amountDistChart', 'memberCompChart',
                       'momTrendChart', 'momGrowthChart', 'catMonthlyChart', 'pointChart', 'adCostChart'];

    canvasIds.forEach(id => {
        const canvas = document.getElementById(id);
        if (!canvas || canvas.style.display === 'none') return;
        const container = canvas.parentElement;
        if (!container) return;

        try {
            const dataUrl = canvas.toDataURL('image/png');
            const img = document.createElement('img');
            img.src = dataUrl;
            img.style.cssText = 'width:100%;height:100%;object-fit:contain;display:block;';
            img.className = 'print-chart-img';

            canvas.style.display = 'none';
            container.appendChild(img);
            _printBackups.push({ canvas, img, container });
        } catch (e) {
            console.warn('차트 이미지 변환 실패:', id, e);
        }
    });
}

function restoreCanvas() {
    _printBackups.forEach(({ canvas, img }) => {
        if (img.parentElement) img.parentElement.removeChild(img);
        canvas.style.display = '';
    });
    _printBackups = [];
    renderCharts();
}

function printWithCharts() {
    setPrintMeta();
    replaceCanvasWithImage();

    // 이미지가 렌더링될 시간을 확보한 후 인쇄
    setTimeout(() => {
        window.print();
        // 인쇄 다이얼로그가 닫힌 후 복구
        setTimeout(restoreCanvas, 500);
    }, 300);
}

// Ctrl+P 등 직접 인쇄 대비 (fallback)
window.addEventListener('beforeprint', () => {
    if (_printBackups.length === 0) {
        setPrintMeta();
        replaceCanvasWithImage();
    }
});
window.addEventListener('afterprint', () => {
    if (_printBackups.length > 0) {
        restoreCanvas();
    }
});

function normalizeRow(r) { const nr = {}; Object.keys(r).forEach(k => nr[k.replace(/\s/g, "")] = r[k]); return nr; }
function parseAmount(v) { return typeof v === 'number' ? v : parseFloat(String(v || 0).replace(/,/g, '')) || 0; }
function formatDate(d) {
    if (d instanceof Date) {
        const y = d.getFullYear();
        const mo = String(d.getMonth() + 1).padStart(2, '0');
        return `${y}-${mo}`;
    }
    const m = String(d).match(/(\d{4})[-. ](\d{1,2})/);
    return m ? `${m[1]}-${m[2].padStart(2, '0')}` : String(d).substring(0, 7);
}
