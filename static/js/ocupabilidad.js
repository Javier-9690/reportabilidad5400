const form = document.getElementById('filterForm');
const tablesContainer = document.getElementById('tablesContainer');
const conclusions = document.getElementById('conclusions');
const summary = document.getElementById('reportSummary');
const generalCanvas = document.getElementById('generalChart');
const percentCanvas = document.getElementById('percentChart');
let generalChart = null;
let percentChart = null;
let lastData = null;

function params() {
    const p = new URLSearchParams();
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;
    const curve = document.getElementById('curveId').value;
    const bedsDefault = document.getElementById('bedsDefault').value;
    const bedsByDate = document.getElementById('bedsByDate').value;
    if (start) p.set('start_date', start);
    if (end) p.set('end_date', end);
    if (curve) p.set('curve_id', curve);
    if (bedsDefault) p.set('beds_default', bedsDefault);
    if (bedsByDate.trim()) p.set('beds_by_date', bedsByDate.trim());
    return p;
}

function num(value) {
    return Math.round(Number(value || 0)).toLocaleString('es-CL');
}

function pctValue(value) {
    return `${Math.round(Number(value || 0))}%`;
}

function metricCard(label, value, icon, extraClass = '') {
    return `
        <div class="col-6 col-lg-2-4">
            <div class="card report-metric ${extraClass}">
                <div class="card-body py-2 d-flex align-items-center justify-content-between gap-2">
                    <div>
                        <div class="small text-muted">${label}</div>
                        <div class="fs-5 fw-bold">${value}</div>
                    </div>
                    <i class="bi ${icon}"></i>
                </div>
            </div>
        </div>`;
}

function renderSummary(data) {
    const totals = data.totals || {};
    const beds = Number(totals.beds_total || 0);
    const censo = Number(totals.censo_total || 0);
    const reservas = Number(totals.reservas_total || 0);
    summary.innerHTML =
        metricCard('Curva acum.', num(totals.curva_total), 'bi-graph-up-arrow', 'metric-plan') +
        metricCard('Reservas', num(reservas), 'bi-calendar-check', 'metric-real') +
        metricCard('Censo', num(censo), 'bi-person-check-fill', 'metric-up') +
        metricCard('No presentes', num(totals.no_presentes_total), 'bi-person-dash', 'metric-down') +
        metricCard('Ocupación', beds ? pctValue(censo / beds * 100) : '0%', 'bi-percent', 'metric-compliance');
}

function renderConclusions(data) {
    const items = data.conclusions || [];
    if (!items.length) {
        conclusions.innerHTML = '';
        return;
    }
    conclusions.innerHTML = `
        <div class="card shadow-sm conclusion-card">
            <div class="card-body py-2">
                <div class="fw-bold mb-1"><i class="bi bi-lightbulb"></i> Conclusiones automáticas</div>
                <div class="row g-1">
                    ${items.map(text => `<div class="col-12 col-lg-6 small"><span class="conclusion-bullet">•</span> ${text}</div>`).join('')}
                </div>
            </div>
        </div>`;
}

function renderGeneralChart(data) {
    const labels = data.date_labels || [];
    const totals = data.totals || {};
    if (generalChart) generalChart.destroy();
    generalChart = new Chart(generalCanvas, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'Total curva', data: totals.curva || [], borderColor: '#4472C4', backgroundColor: '#4472C4', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Total reservas', data: totals.reservas || [], borderColor: '#ED7D31', backgroundColor: '#ED7D31', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Total censo', data: totals.censo || [], borderColor: '#A5A5A5', backgroundColor: '#A5A5A5', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Reservas no presentes', data: totals.no_presentes || [], borderColor: '#FFC000', backgroundColor: '#FFC000', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Camas habilitadas', data: data.beds_by_date || [], borderColor: '#5B9BD5', backgroundColor: '#5B9BD5', tension: .25, borderWidth: 2.5, pointRadius: 3 },
            ]
        },
        options: chartOptions(value => num(value))
    });
}

function renderPercentChart(data) {
    const labels = data.date_labels || [];
    const percentages = data.percentages || {};
    if (percentChart) percentChart.destroy();
    percentChart = new Chart(percentCanvas, {
        type: 'line',
        data: {
            labels,
            datasets: [
                { label: 'Ocupación', data: percentages.ocupacion || [], borderColor: '#4472C4', backgroundColor: '#4472C4', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Eficiencia', data: percentages.eficiencia || [], borderColor: '#ED7D31', backgroundColor: '#ED7D31', tension: .25, borderWidth: 2.5, pointRadius: 3 },
                { label: 'Disponibilidad', data: percentages.disponibilidad || [], borderColor: '#A5A5A5', backgroundColor: '#A5A5A5', tension: .25, borderWidth: 2.5, pointRadius: 3 },
            ]
        },
        options: chartOptions(value => `${Math.round(value)}%`, true)
    });
}

function chartOptions(tickCallback, pct = false) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
            legend: { position: 'bottom', labels: { usePointStyle: true, font: { weight: '600' } } },
            tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${pct ? pctValue(ctx.parsed.y) : num(ctx.parsed.y)}` } }
        },
        scales: {
            x: { grid: { display: false }, ticks: { maxRotation: 45 } },
            y: { beginAtZero: true, grid: { color: 'rgba(16,24,40,.08)' }, ticks: { callback: tickCallback } }
        }
    };
}

function renderSection(title, rows, totals, totalLabel, kind = 'num') {
    const dates = lastData.date_labels || [];
    return `
        <div class="table-responsive report-table-wrapper shadow-sm mb-3">
            <table class="table table-sm table-bordered report-table occupancy-table mb-0">
                <thead>
                    <tr class="report-main-title"><th colspan="${dates.length + 2}">${title}</th></tr>
                    <tr>
                        <th class="sticky-col gerencia-col">GERENCIAS</th>
                        ${dates.map(d => `<th class="date-header">${d}</th>`).join('')}
                        <th class="total-col-head">TOTAL</th>
                    </tr>
                </thead>
                <tbody>
                    ${(rows || []).map(row => `
                        <tr class="${row.gerencia === 'SIN MATCH EN CURVA' ? 'unmatched-row' : ''}">
                            <td class="sticky-col gerencia-col">${row.gerencia}</td>
                            ${(row.values || []).map(v => `<td class="actual-cell ${Number(v || 0) === 0 ? 'zero-cell' : ''}">${kind === 'pct' ? pctValue(v) : num(v)}</td>`).join('')}
                            <td class="total-col">${kind === 'pct' ? pctValue(row.total) : num(row.total)}</td>
                        </tr>`).join('')}
                    <tr class="total-row">
                        <td class="sticky-col gerencia-col">${totalLabel}</td>
                        ${(totals || []).map(v => `<td>${kind === 'pct' ? pctValue(v) : num(v)}</td>`).join('')}
                        <td>${kind === 'pct' ? '' : num((totals || []).reduce((a,b)=>a+Number(b||0),0))}</td>
                    </tr>
                </tbody>
            </table>
        </div>`;
}

function renderPercentTable(data) {
    const dates = data.date_labels || [];
    const rows = [
        ['CAMAS HABILITADAS', data.beds_by_date || [], 'num'],
        ['OCUPACIÓN', data.percentages?.ocupacion || [], 'pct'],
        ['EFICIENCIA', data.percentages?.eficiencia || [], 'pct'],
        ['DISPONIBILIDAD', data.percentages?.disponibilidad || [], 'pct'],
    ];
    return `
        <div class="table-responsive report-table-wrapper shadow-sm mb-3">
            <table class="table table-sm table-bordered report-table occupancy-table mb-0">
                <thead>
                    <tr class="report-main-title"><th colspan="${dates.length + 1}">PORCENTAJES RELEVANTES</th></tr>
                    <tr><th class="sticky-col gerencia-col"></th>${dates.map(d => `<th class="date-header">${d}</th>`).join('')}</tr>
                </thead>
                <tbody>
                    ${rows.map(([label, vals, kind]) => `
                        <tr>
                            <td class="sticky-col gerencia-col">${label}</td>
                            ${vals.map(v => `<td>${kind === 'pct' ? pctValue(v) : num(v)}</td>`).join('')}
                        </tr>`).join('')}
                </tbody>
            </table>
        </div>`;
}

function renderTables(data) {
    lastData = data;
    const sections = data.sections || {};
    const totals = data.totals || {};
    tablesContainer.innerHTML =
        renderSection('CURVA', sections.curva || [], totals.curva || [], 'TOTAL GENERAL CURVA') +
        renderSection('RESERVAS', sections.reservas || [], totals.reservas || [], 'TOTAL RESERVAS') +
        renderSection('CENSO', sections.censo || [], totals.censo || [], 'TOTAL CENSO') +
        renderSection('RESERVAS NO PRESENTES', sections.no_presentes || [], totals.no_presentes || [], 'TOTAL RESERVAS NO PRESENTES') +
        renderPercentTable(data);
}

function render(data) {
    lastData = data;
    renderSummary(data);
    renderConclusions(data);
    renderGeneralChart(data);
    renderPercentChart(data);
    renderTables(data);
}

async function load() {
    tablesContainer.innerHTML = '<div class="text-center text-muted py-4">Cargando reporte...</div>';
    try {
        const response = await fetch('/api/reports/ocupabilidad?' + params().toString());
        if (!response.ok) throw new Error('No se pudo cargar el reporte');
        render(await response.json());
    } catch (error) {
        summary.innerHTML = '';
        conclusions.innerHTML = '';
        tablesContainer.innerHTML = `<div class="text-center text-danger py-4">${error.message}</div>`;
    }
}

form.addEventListener('submit', event => {
    event.preventDefault();
    load();
});

function ensureExportStatusBox() {
    let box = document.getElementById('exportStatusBox');
    if (!box) {
        box = document.createElement('div');
        box.id = 'exportStatusBox';
        box.className = 'alert alert-info py-2 small d-none';
        const hero = document.querySelector('.report-hero');
        hero.parentNode.insertBefore(box, hero.nextSibling);
    }
    return box;
}

function setExportStatus(message, type = 'info') {
    const box = ensureExportStatusBox();
    box.className = `alert alert-${type} py-2 small`;
    box.innerHTML = message;
}

async function pollExport(jobId, button) {
    const statusUrl = `/api/exports/${jobId}/status`;
    for (let attempt = 0; attempt < 240; attempt++) {
        const response = await fetch(statusUrl);
        const job = await response.json();
        if (job.status === 'completed') {
            setExportStatus(`<i class="bi bi-check-circle"></i> ${job.message || 'Excel listo.'} Descargando...`, 'success');
            button.disabled = false;
            button.innerHTML = '<i class="bi bi-file-earmark-bar-graph"></i> Exportar Excel';
            window.location.href = job.download_url;
            return;
        }
        if (job.status === 'failed') {
            setExportStatus(`<i class="bi bi-exclamation-triangle"></i> ${job.message || 'Error al generar Excel.'}`, 'danger');
            button.disabled = false;
            button.innerHTML = '<i class="bi bi-file-earmark-bar-graph"></i> Exportar Excel';
            return;
        }
        setExportStatus(`<i class="bi bi-hourglass-split"></i> ${job.message || 'Generando Excel...'}`, 'info');
        await new Promise(resolve => setTimeout(resolve, 2500));
    }
    setExportStatus('La exportación sigue en proceso. Espera unos segundos y vuelve a intentar.', 'warning');
    button.disabled = false;
    button.innerHTML = '<i class="bi bi-file-earmark-bar-graph"></i> Exportar Excel';
}

document.getElementById('btnExportOcupabilidad').addEventListener('click', async () => {
    const button = document.getElementById('btnExportOcupabilidad');
    button.disabled = true;
    button.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Preparando Excel...';
    setExportStatus('<i class="bi bi-hourglass-split"></i> Iniciando exportación de ocupabilidad...', 'info');
    try {
        const payload = Object.fromEntries(params());
        const response = await fetch('/api/reports/ocupabilidad/export/start', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload),
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'No se pudo iniciar la exportación');
        pollExport(data.job_id, button);
    } catch (error) {
        setExportStatus(`<i class="bi bi-exclamation-triangle"></i> ${error.message}`, 'danger');
        button.disabled = false;
        button.innerHTML = '<i class="bi bi-file-earmark-bar-graph"></i> Exportar Excel';
    }
});

load();
