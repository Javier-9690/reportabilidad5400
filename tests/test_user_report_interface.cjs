const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const script = fs.readFileSync(path.join(__dirname, '../static/js/user_report.js'), 'utf8');
const sample = {
  labels: ['31-ago', '1-sep', '2-sep'], dates: ['2026-08-31', '2026-09-01', '2026-09-02'],
  categories: ['Doble asignación', 'Reclamos usuarios', 'Solicitudes de usuarios', 'Solicitudes totales Samtech', 'Desviaciones clientes'],
  colors: ['#7950A3', '#B42318', '#1667A5', '#198754', '#B86A00'],
  daily: [[2, 1, 0], [2, 0, 0], [1, 2, 0], [3, 1, 0], [3, 0, 2]],
  open: [1, 1, 2, 2, 0], closed: [1, 1, 1, 1, 0], unclassified: [1, 0, 0, 1, 5]
};

function element(values = {}) {
  return Object.assign({
    events: {}, attributes: {}, hidden: true, classes: {},
    addEventListener(name, handler) { this.events[name] = handler; },
    setAttribute(name, value) { this.attributes[name] = value; },
    getAttribute(name) { return this.attributes[name]; },
    classList: { toggle() {} }
  }, values);
}

function run({ withReport = true, chartAvailable = true, restoredDate } = {}) {
  const inputs = [element({ value: restoredDate || '2026-08-31', defaultValue: '2026-08-31' }),
    element({ value: '2026-09-02', defaultValue: '2026-09-02' })];
  const exported = element();
  exported.classList.toggle = (name, state) => { exported.classes[name] = state; };
  const nodes = {
    userReportFilters: { querySelectorAll: () => inputs },
    exportUserReport: withReport ? exported : null,
    reportFiltersChanged: element(), userChartsUnavailable: element(),
    userReportChartData: withReport ? { textContent: JSON.stringify(sample) } : null,
    userDailyChart: { id: 'daily' }, userStatusChart: { id: 'status' }
  };
  const charts = [];
  const events = {};
  const context = { document: { getElementById: name => nodes[name] }, window: { addEventListener: (name, fn) => { events[name] = fn; } } };
  if (chartAvailable) context.Chart = function (canvas, config) { charts.push({canvas, config: JSON.parse(JSON.stringify(config))}); };
  vm.runInNewContext(script, context);
  return { inputs, exported, nodes, charts, events };
}

const report = run();
assert.equal(report.charts.length, 2);
assert.deepEqual(report.charts[0].config.data.datasets.map(series => series.data), sample.daily);
assert.deepEqual(report.charts[0].config.data.datasets.map(series => series.label), sample.categories);
assert.deepEqual(report.charts[0].config.data.datasets.map(series => series.borderColor), sample.colors);
assert.deepEqual(report.charts[1].config.data.labels, sample.categories);
assert.deepEqual(report.charts[1].config.data.datasets.map(series => series.data), [sample.open, sample.closed, sample.unclassified]);
assert.equal(report.nodes.reportFiltersChanged.hidden, true);
assert.equal(report.exported.getAttribute('aria-disabled'), 'false');
report.inputs[1].value = '2026-09-05';
report.inputs[1].events.input();
assert.equal(report.nodes.reportFiltersChanged.hidden, false);
assert.equal(report.exported.classes.disabled, true);
assert.equal(report.exported.tabIndex, -1);
let prevented = false;
report.exported.events.click({ preventDefault: () => { prevented = true; } });
assert.equal(prevented, true);
report.inputs[1].value = '2026-09-02';
report.events.pageshow();
assert.equal(report.exported.getAttribute('aria-disabled'), 'false');
assert.equal(report.exported.tabIndex, 0);
assert.equal(report.nodes.reportFiltersChanged.hidden, true);
const restored = run({ restoredDate: '2026-08-01' });
assert.equal(restored.exported.classes.disabled, true);
const noCharts = run({ chartAvailable: false });
assert.equal(noCharts.nodes.userChartsUnavailable.hidden, false);
assert.equal(noCharts.exported.getAttribute('aria-disabled'), 'false');
assert.equal(run({ withReport: false }).charts.length, 0);
console.log('OK: gráficos, fechas cambiadas, regreso del navegador y carga sin Chart.js.');
