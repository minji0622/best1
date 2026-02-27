let rawData = [];
let tempSheetData = [];
let tempHeaders = [];
let tempFileName = "";
let analyzedData = {
    monthly: {},
    categories: {},
    yearly: {},
    mom: {}
};
let charts = {
    sales: null,
    category: null
};

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

// 4. 데이터 집계 로직
function processData() {
    const monthly = {}; const categories = {}; const yearly = {}; const userHistory = {};

    rawData.forEach(row => {
        const nr = normalizeRow(row);
        const uid = nr.주문자ID || nr.주문자명 || '익명';
        userHistory[uid] = (userHistory[uid] || 0) + 1;
    });

    rawData.forEach(row => {
        const nr = normalizeRow(row);
        if (!nr.주문일시) return;
        const m = formatDate(nr.주문일시);
        const y = m.substring(0, 4);
        const amt = parseAmount(nr.결제금액);
        const uid = nr.주문자ID || nr.주문자명 || '익명';
        const orderCnt = nr.주문횟수 || userHistory[uid];

        const adCost = parseAmount(nr.광고비);

        if (!monthly[m]) monthly[m] = { totalAmount: 0, totalCount: 0, newAmount: 0, newCount: 0, adCost: 0, users: new Set(), memberTypes: {} };
        monthly[m].totalAmount += amt; monthly[m].totalCount += 1; monthly[m].users.add(uid);
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
    });

    analyzedData.monthly = monthly; analyzedData.categories = categories; analyzedData.yearly = yearly;
    
    // KPI 계산
    const ms = Object.keys(monthly).sort();
    if (ms.length > 0) {
        const cur = monthly[ms[ms.length-1]];
        const pre = ms.length > 1 ? monthly[ms[ms.length-2]] : null;
        analyzedData.mom = {
            revenue: cur.totalAmount,
            revenueGrowth: pre ? (cur.totalAmount - pre.totalAmount) / pre.totalAmount * 100 : 0,
            users: cur.users.size,
            userGrowth: pre ? (cur.users.size - pre.users.size) / pre.users.size * 100 : 0,
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

    const mom = analyzedData.mom;
    updateKPICard('kpi-revenue', mom.revenue, mom.revenueGrowth, '원');
    updateKPICard('kpi-users', mom.users, mom.userGrowth, '명');
    if (document.getElementById('kpi-new-ratio')) document.getElementById('kpi-new-ratio').innerText = mom.newRatio.toFixed(1) + '%';
    if (document.getElementById('kpi-aov')) document.getElementById('kpi-aov').innerText = Math.round(mom.aov).toLocaleString() + '원';

    // ROAS KPI
    const roasEl = document.getElementById('kpi-roas');
    const roasSubEl = document.getElementById('kpi-roas-sub');
    if (roasEl) {
        if (mom.roas !== null) {
            roasEl.innerText = Math.round(mom.roas).toLocaleString() + '%';
            roasEl.style.color = mom.roas >= 100 ? 'var(--success)' : 'var(--danger)';
        } else {
            roasEl.innerText = '-';
            roasEl.style.color = '';
        }
    }
    if (roasSubEl) {
        roasSubEl.innerHTML = mom.adCost > 0
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

    const dateFile = today.toISOString().slice(0, 10);
    XLSX.writeFile(wb, `매출보고서_${dateFile}.xlsx`);
});

function copyFormula(id) {
    const txt = document.getElementById(id)?.innerText;
    if(txt) navigator.clipboard.writeText(txt).then(() => alert("복사되었습니다."));
}

// 인쇄 전: 메타 정보 세팅 + canvas 크기 리셋
window.addEventListener('beforeprint', () => {
    const el = document.getElementById('print-meta');
    if (el) {
        const now = new Date();
        const dateStr = `${now.getFullYear()}년 ${now.getMonth()+1}월 ${now.getDate()}일`;
        const ms = Object.keys(analyzedData.monthly).sort();
        const rangeStr = ms.length ? `분석 기간: ${ms[0]} ~ ${ms[ms.length-1]}` : '';
        el.innerHTML = `출력일: ${dateStr}${rangeStr ? '<br>' + rangeStr : ''}`;
    }

    // JavaScript로 설정된 인라인 크기를 초기화해야 A4에 맞게 렌더링됨
    const sCanvas = document.getElementById('salesChart');
    if (sCanvas) {
        sCanvas.style.width = '100%';
        sCanvas.style.height = '200px';
    }
    if (charts.sales) charts.sales.resize();
    if (charts.category) charts.category.resize();
});

// 인쇄 후: 화면용 크기로 복구
window.addEventListener('afterprint', () => {
    renderCharts();
});

function normalizeRow(r) { const nr = {}; Object.keys(r).forEach(k => nr[k.replace(/\s/g, "")] = r[k]); return nr; }
function parseAmount(v) { return typeof v === 'number' ? v : parseFloat(String(v || 0).replace(/,/g, '')) || 0; }
function formatDate(d) {
    if (d instanceof Date) return d.toISOString().substring(0, 7);
    const m = String(d).match(/(\d{4})[-. ](\d{1,2})/);
    return m ? `${m[1]}-${m[2].padStart(2, '0')}` : String(d).substring(0, 7);
}
