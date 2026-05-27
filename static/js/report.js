const form = document.getElementById('filterForm');
const head = document.getElementById('reportHead');
const body = document.getElementById('reportBody');
const conclusions = document.getElementById('conclusions');
const summary = document.getElementById('reportSummary');

function params() {
    const p = new URLSearchParams();
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;
    const curve = document.getElementById('curveId').value;
    if (start) p.set('start_date', start);
    if (end) p.set('end_date', end);
    if (curve) p.set('curve_id', curve);
    return p;
}

function num(value) {
    return Math.round(Number(value || 0)).toLocaleString('es-CL');
}

function pct(real, plan) {
    real = Number(real || 0);
    plan = Number(plan || 0);
    if (!plan) return real ? 'Sin plan' : '0%';
    return `${Math.round((real / plan) * 100)}%`;
}

function diffClass(value) {
    value = Number(value || 0);
    if (value > 0) return 'diff-positive';
    if (value < 0) return 'diff-negative';
    return 'diff-neutral';
}

function metricCard(label, value, icon, extraClass = '') {
    return `
        <div class="col-6 col-lg-3">
            <div class="card report-metric ${extraClass}">
                <div class="card-body py-2 d-flex align-items-center justify-content-between gap-2">
                    <div>
                        <div class="small text-muted">${label}</div>
                        <div class="fs-4 fw-bold">${value}</div>
                    </div>
                    <i class="bi ${icon}"></i>
                </div>
            </div>
        </div>`;
}

function renderSummary(data) {
    const real = Number(data.grand_total || 0);
    const plan = Number(data.planned_grand_total || 0);
    const diff = real - plan;
    summary.innerHTML =
        metricCard('Real acumulado', num(real), 'bi-people-fill', 'metric-real') +
        metricCard('Plan acumulado', num(plan), 'bi-graph-up-arrow', 'metric-plan') +
        metricCard('Diferencia', `${diff > 0 ? '+' : ''}${num(diff)}`, diff >= 0 ? 'bi-arrow-up-circle' : 'bi-arrow-down-circle', diff >= 0 ? 'metric-up' : 'metric-down') +
        metricCard('Cumplimiento', pct(real, plan), 'bi-percent', 'metric-compliance');
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

function renderTable(data) {
    const dates = data.date_labels || [];
    const colCount = dates.length + 5;

    head.innerHTML = `
        <tr class="report-main-title">
            <th colspan="${colCount}">
                RESUMEN DE DOTACIÓN POR GERENCIA
                <span>${data.start_date || ''} ${data.end_date ? 'al ' + data.end_date : ''}</span>
            </th>
        </tr>
        <tr>
            <th class="sticky-col gerencia-col">GERENCIAS</th>
            ${dates.map(d => `<th class="date-header">${d}</th>`).join('')}
            <th class="total-col-head">TOTAL REAL</th>
            <th class="plan-col-head">TOTAL PLAN</th>
            <th class="diff-col-head">DIF.</th>
            <th class="cumpl-col-head">CUMP.</th>
        </tr>`;

    body.innerHTML = '';
    if (!data.rows || !data.rows.length) {
        body.innerHTML = `<tr><td class="text-center text-muted" colspan="${colCount}">Sin datos para el rango seleccionado</td></tr>`;
        return;
    }

    data.rows.forEach(row => {
        const isUnmatched = row.gerencia === 'SIN MATCH EN CURVA';
        const rowDiff = Number(row.difference || 0);
        body.innerHTML += `
            <tr class="${isUnmatched ? 'unmatched-row' : ''}">
                <td class="sticky-col gerencia-col">${row.gerencia}</td>
                ${(row.values || []).map((value, idx) => {
                    const plan = (row.planned_values || [])[idx] || 0;
                    const diff = Number(value || 0) - Number(plan || 0);
                    return `<td class="actual-cell ${Number(value || 0) === 0 ? 'zero-cell' : ''}" title="Plan: ${num(plan)} | Diferencia: ${diff > 0 ? '+' : ''}${num(diff)}">${num(value)}</td>`;
                }).join('')}
                <td class="total-col">${num(row.total)}</td>
                <td class="plan-col">${num(row.planned_total)}</td>
                <td class="${diffClass(rowDiff)}">${rowDiff > 0 ? '+' : ''}${num(rowDiff)}</td>
                <td class="cumpl-col">${pct(row.total, row.planned_total)}</td>
            </tr>`;
    });

    const grandDiff = Number(data.grand_total || 0) - Number(data.planned_grand_total || 0);
    body.innerHTML += `
        <tr class="total-row">
            <td class="sticky-col gerencia-col">TOTAL</td>
            ${(data.totals_by_date || []).map(v => `<td>${num(v)}</td>`).join('')}
            <td class="total-col">${num(data.grand_total)}</td>
            <td class="plan-col">${num(data.planned_grand_total)}</td>
            <td class="${diffClass(grandDiff)}">${grandDiff > 0 ? '+' : ''}${num(grandDiff)}</td>
            <td class="cumpl-col">${pct(data.grand_total, data.planned_grand_total)}</td>
        </tr>`;
}

function render(data) {
    renderSummary(data);
    renderConclusions(data);
    renderTable(data);
}

async function load() {
    head.innerHTML = '';
    body.innerHTML = '<tr><td class="text-center text-muted py-4">Cargando reporte...</td></tr>';
    try {
        const response = await fetch('/api/reports/dotacion-gerencia?' + params().toString());
        if (!response.ok) throw new Error('No se pudo cargar el reporte');
        render(await response.json());
    } catch (error) {
        summary.innerHTML = '';
        conclusions.innerHTML = '';
        body.innerHTML = `<tr><td class="text-center text-danger py-4">${error.message}</td></tr>`;
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
            button.innerHTML = '<i class="bi bi-file-earmark-excel"></i> Exportar Excel completo';
            window.location.href = job.download_url;
            return;
        }

        if (job.status === 'failed') {
            setExportStatus(`<i class="bi bi-exclamation-triangle"></i> ${job.message || 'Error al generar Excel.'}`, 'danger');
            button.disabled = false;
            button.innerHTML = '<i class="bi bi-file-earmark-excel"></i> Exportar Excel completo';
            return;
        }

        setExportStatus(`<i class="bi bi-hourglass-split"></i> ${job.message || 'Generando Excel...'} Puedes seguir usando la aplicación mientras termina.`, 'info');
        await new Promise(resolve => setTimeout(resolve, 2500));
    }

    setExportStatus('<i class="bi bi-clock-history"></i> La exportación sigue en proceso. Espera unos segundos y vuelve a intentar descargar desde el aviso.', 'warning');
    button.disabled = false;
    button.innerHTML = '<i class="bi bi-file-earmark-excel"></i> Exportar Excel completo';
}

document.getElementById('btnExport').addEventListener('click', async () => {
    const button = document.getElementById('btnExport');
    button.disabled = true;
    button.innerHTML = '<span class="spinner-border spinner-border-sm"></span> Preparando Excel...';
    setExportStatus('<i class="bi bi-hourglass-split"></i> Iniciando exportación en segundo plano...', 'info');

    try {
        const response = await fetch('/api/reports/dotacion-gerencia/export/start', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(Object.fromEntries(params())),
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'No se pudo iniciar la exportación');
        pollExport(data.job_id, button);
    } catch (error) {
        setExportStatus(`<i class="bi bi-exclamation-triangle"></i> ${error.message}`, 'danger');
        button.disabled = false;
        button.innerHTML = '<i class="bi bi-file-earmark-excel"></i> Exportar Excel completo';
    }
});

load();
