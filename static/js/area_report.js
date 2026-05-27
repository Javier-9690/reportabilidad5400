let areaData = null;
let mainChart = null;
let topChart = null;
const page = document.querySelector('.area-report-page');
const reportKind = page ? page.dataset.kind : 'egp';

function fmt(n) {
    const v = Number(n || 0);
    return Math.round(v).toLocaleString('es-CL');
}

function pct(n) {
    if (n === null || n === undefined || Number.isNaN(Number(n))) return '-';
    return `${Number(n).toFixed(1)}%`;
}

function getParams() {
    const params = new URLSearchParams();
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;
    const curve = document.getElementById('curveId').value;
    if (start) params.set('start_date', start);
    if (end) params.set('end_date', end);
    if (curve) params.set('curve_id', curve);
    return params;
}

function setDefaultDates(data) {
    if (data.start_date && !document.getElementById('startDate').value) document.getElementById('startDate').value = data.start_date;
    if (data.end_date && !document.getElementById('endDate').value) document.getElementById('endDate').value = data.end_date;
}

function renderSummary(data) {
    const totals = data.totals || {};
    const cards = [
        ['IDs curva', data.rows.length, 'bi-diagram-3'],
        ['Plan acumulado', fmt(totals.grand_plan), 'bi-clipboard-data'],
        ['Reservas acumuladas', fmt(totals.grand_reservas), 'bi-calendar-check'],
        ['Censo acumulado', fmt(totals.grand_censo), 'bi-house-check'],
        ['No show reservas', fmt(totals.grand_no_show_reservas), 'bi-person-dash'],
        ['Eficiencia', pct(totals.eficiencia), 'bi-percent'],
    ];
    document.getElementById('reportSummary').innerHTML = cards.map(([label, value, icon]) => `
        <div class="col-6 col-md-2">
            <div class="card shadow-sm report-kpi-card h-100">
                <div class="card-body py-2">
                    <div class="small text-muted"><i class="bi ${icon}"></i> ${label}</div>
                    <div class="fs-5 fw-bold text-primary-red">${value}</div>
                </div>
            </div>
        </div>`).join('');
}

function renderConclusions(data) {
    const items = data.conclusions || [];
    document.getElementById('conclusions').innerHTML = `
        <div class="alert alert-light border shadow-sm py-2 mb-0">
            <div class="fw-bold text-primary-red mb-1"><i class="bi bi-lightbulb"></i> Conclusiones</div>
            <ul class="mb-0 small">${items.map(x => `<li>${x}</li>`).join('')}</ul>
        </div>`;
}

function destroyCharts() {
    if (mainChart) mainChart.destroy();
    if (topChart) topChart.destroy();
    mainChart = null; topChart = null;
}

function renderCharts(data) {
    destroyCharts();
    const labels = data.date_labels || [];
    const totals = data.totals || {};
    mainChart = new Chart(document.getElementById('areaMainChart'), {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'Curva', data: totals.plan || [], borderWidth: 2.5, tension: 0.25, pointRadius: 3 },
                { label: 'Reservas', data: totals.reservas || [], borderWidth: 2.5, tension: 0.25, pointRadius: 3 },
                { label: 'Censo', data: totals.censo || [], borderWidth: 2.5, tension: 0.25, pointRadius: 3 },
                { label: 'No show reservas', data: totals.no_show_reservas || [], borderWidth: 2.5, tension: 0.25, pointRadius: 3 },
            ]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true } } }
    });

    const top = [...(data.rows || [])].sort((a, b) => (b.total_censo || 0) - (a.total_censo || 0)).slice(0, 15);
    topChart = new Chart(document.getElementById('areaTopChart'), {
        type: 'bar',
        data: { labels: top.map(r => r.id), datasets: [{ label: 'Censo acumulado', data: top.map(r => r.total_censo || 0), borderWidth: 1 }] },
        options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true } } }
    });
}

function renderSummaryTable(data) {
    document.getElementById('summaryHead').innerHTML = `<tr>
        <th>ID</th><th>Empresa</th><th>Área</th><th>Turno</th><th>Contrato</th><th>Formato</th>
        <th>Total curva</th><th>Total reservas</th><th>Total censo</th><th>No show reservas</th><th>Cumplimiento</th>
    </tr>`;
    document.getElementById('summaryBody').innerHTML = (data.rows || []).map(r => `<tr>
        <td class="fw-bold">${r.id}</td><td>${r.empresa || '-'}</td><td>${r.area || '-'}</td><td>${r.turno || '-'}</td><td>${r.tipo_contrato || '-'}</td><td>${r.formato || '-'}</td>
        <td class="text-end">${fmt(r.total_plan)}</td><td class="text-end">${fmt(r.total_reservas)}</td><td class="text-end">${fmt(r.total_censo)}</td><td class="text-end">${fmt(r.no_show_reservas)}</td><td class="text-end">${pct(r.cumplimiento)}</td>
    </tr>`).join('') || `<tr><td colspan="11" class="text-center text-muted py-3">Sin datos</td></tr>`;
}

function sectionTable(title, key, totalKey, data) {
    const dates = data.date_labels || [];
    const totalsMap = { planned_values: 'plan', reservation_values: 'reservas', census_values: 'censo' };
    const totalValues = (data.totals || {})[totalsMap[key]] || [];
    const head = `<tr><th>ID</th><th>Empresa</th><th>Área</th>${dates.map(d => `<th class="date-head">${d}</th>`).join('')}<th>Total</th></tr>`;
    const body = (data.rows || []).map(r => `<tr>
        <td class="fw-bold">${r.id}</td><td>${r.empresa || '-'}</td><td>${r.area || '-'}</td>
        ${(r[key] || []).map(v => `<td class="text-end">${fmt(v)}</td>`).join('')}
        <td class="text-end fw-bold total-cell">${fmt(r[totalKey])}</td>
    </tr>`).join('');
    const foot = `<tr class="total-row"><td colspan="3">TOTAL ${title}</td>${totalValues.map(v => `<td class="text-end">${fmt(v)}</td>`).join('')}<td class="text-end">${fmt(totalValues.reduce((a,b)=>a+Number(b||0),0))}</td></tr>`;
    return `<div class="card shadow-sm mb-3"><div class="card-header py-2 fw-bold text-primary-red">${title}</div><div class="table-responsive"><table class="table table-sm table-bordered report-table mb-0"><thead>${head}</thead><tbody>${body}${foot}</tbody></table></div></div>`;
}

function renderSections(data) {
    document.getElementById('sectionsContainer').innerHTML = [
        sectionTable('CURVA', 'planned_values', 'total_plan', data),
        sectionTable('RESERVAS', 'reservation_values', 'total_reservas', data),
        sectionTable('CENSO', 'census_values', 'total_censo', data),
    ].join('');
}

async function loadReport() {
    const params = getParams();
    const resp = await fetch(`/api/reports/${reportKind}?${params.toString()}`);
    const data = await resp.json();
    areaData = data;
    setDefaultDates(data);
    renderSummary(data);
    renderConclusions(data);
    renderCharts(data);
    renderSummaryTable(data);
    renderSections(data);
}

document.getElementById('filterForm').addEventListener('submit', e => { e.preventDefault(); loadReport(); });
document.getElementById('btnExportArea').addEventListener('click', async () => {
    const params = Object.fromEntries(getParams().entries());
    const resp = await fetch(`/api/reports/${reportKind}/export/start`, { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(params) });
    const data = await resp.json();
    if (data.page_url) window.location.href = data.page_url;
});

loadReport();
