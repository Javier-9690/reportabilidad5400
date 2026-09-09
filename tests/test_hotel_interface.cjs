// Verificación de los eventos de la interfaz sin dependencias de navegador.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function element(properties = {}) {
  return Object.assign({
    style: {}, handlers: {},
    addEventListener(type, callback) {
      (this.handlers[type] ||= []).push(callback);
    },
    fire(type) { (this.handlers[type] || []).forEach(callback => callback()); },
  }, properties);
}

function run(filename, context) {
  const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'js', filename), 'utf8');
  vm.runInNewContext(source, context, {filename});
}

const widgets = Object.fromEntries([
  'bulkDeleteForm', 'selectAllRecords', 'deleteSelectedRecords', 'selectedRecordCount', 'selectedButtonCount',
].map(id => [id, element({checked: false})]));
const boxes = [false, false].map(checked => {
  const row = {selected: false, classList: {toggle(name, selected) { row.selected = selected; }}};
  return element({checked, closest: () => row});
});
const table = element();
const viewport = element({scrollWidth: 1800, clientWidth: 640, scrollLeft: 0, querySelector: () => table});
const top = element({scrollLeft: 0, hidden: true});
const spacer = element();
const container = element({querySelector: selector => ({
  '[data-record-scroll]': viewport,
  '[data-record-scroll-top]': top,
  '[data-record-scroll-width]': spacer,
})[selector]});
const browserWindow = element();
const resizeCallbacks = [];
run('registros.js', {
  document: {
    getElementById: id => widgets[id],
    querySelectorAll: selector => selector === '[data-record-table]' ? [container] : boxes,
  },
  window: browserWindow,
  ResizeObserver: class {
    constructor(callback) { resizeCallbacks.push(callback); }
    observe() {}
  },
});

assert.equal(widgets.deleteSelectedRecords.disabled, true);
boxes[0].checked = true;
boxes[0].fire('change');
assert.equal(widgets.selectedRecordCount.textContent, '1 seleccionados');
assert.equal(widgets.selectAllRecords.indeterminate, true);
assert.equal(boxes[0].closest('tr').selected, true);
widgets.selectAllRecords.checked = true;
widgets.selectAllRecords.fire('change');
assert.ok(boxes.every(box => box.checked));
assert.equal(widgets.selectedButtonCount.textContent, 2);
widgets.selectAllRecords.checked = false;
widgets.selectAllRecords.fire('change');
assert.ok(boxes.every(box => !box.checked));
assert.equal(widgets.deleteSelectedRecords.disabled, true);

assert.equal(top.hidden, false);
assert.equal(spacer.style.width, '1800px');
assert.equal(top.style.width, '640px');
top.scrollLeft = 700;
top.fire('scroll');
assert.equal(viewport.scrollLeft, 700);
viewport.scrollLeft = 1100;
viewport.fire('scroll');
assert.equal(top.scrollLeft, 1100);
viewport.clientWidth = 340;
browserWindow.fire('resize');
assert.equal(top.style.width, '340px');
viewport.scrollWidth = 340;
resizeCallbacks.forEach(callback => callback());
assert.equal(top.hidden, true);
viewport.scrollWidth = 1600;
browserWindow.fire('pageshow');
assert.equal(top.hidden, false);
assert.equal(spacer.style.width, '1600px');

// Una página sin registros no necesita tabla ni casillas.
run('registros.js', {document: {getElementById: () => null, querySelectorAll: () => []}, window: element()});

const automatic = element({checked: true});
const day = element({value: '10'});
const night = element({value: '12'});
const total = element({value: ''});
const census = {querySelector: selector => ({
  '[name="censo_total_auto"]': automatic,
  '[name="censo_dia"]': day,
  '[name="censo_noche"]': night,
  '[name="total"]': total,
})[selector]};
const filename = element({value: 'Anterior.pdf'});
const picker = element({dataset: {pdfNameTarget: 'entry-archivo_pdf'}, files: []});
run('hotel_forms.js', {document: {
  querySelectorAll: selector => selector === '[data-census-form]' ? [census] : [picker],
  getElementById: id => id === 'entry-archivo_pdf' ? filename : null,
}});
assert.equal(total.value, '22');
assert.equal(total.readOnly, true);
day.value = '15';
day.fire('input');
assert.equal(total.value, '27');
automatic.checked = false;
automatic.fire('change');
total.value = '30';
night.value = '13';
night.fire('input');
assert.equal(total.value, '30');
assert.equal(total.readOnly, false);
picker.files = [{name: 'Constancia revisión.pdf'}];
picker.fire('change');
assert.equal(filename.value, 'Constancia revisión.pdf');
picker.files = [];
picker.fire('change');
assert.equal(filename.value, 'Constancia revisión.pdf');

console.log('Interfaz verificada: selección múltiple, barras sincronizadas, ajuste de ancho, total de censo y nombre de PDF.');
