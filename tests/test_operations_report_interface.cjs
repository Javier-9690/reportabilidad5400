const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const script = fs.readFileSync(path.join(__dirname, '../static/js/operations_report.js'), 'utf8');
const sample = { labels: ['01/09/2026', '02/09/2026'],
  series: [{label:'Misceláneos',color:'#916335',values:[2,3]}, {label:'Desviaciones',color:'#BD771C',values:[1,2]}],
  categories: ['Misceláneos', 'Desviaciones'],
  states: {open:[1,0],closed:[4,0],unknown:[0,0],not_applicable:[0,3]},
  companies:[{label:'Empresa',count:5}]
};
function element(values={}) {
  return Object.assign({events:{}, attributes:{}, hidden:true,
    addEventListener(name, handler) {this.events[name]=handler;},
    setAttribute(name,value) {this.attributes[name]=value;},
    getAttribute(name) {return this.attributes[name];}}, values);
}
function run({chartAvailable=true, withRows=true, category=false}={}) {
  const inputs = [element({type:'date',value:'2026-09-01',defaultValue:'2026-09-01'}),
    element({type:'search',value:'',defaultValue:''}),
    element({type:'checkbox',checked:false,defaultChecked:false})];
  const links = [element(),element(),element()];
  const nodes = {operationsFilters:{querySelectorAll:()=>inputs}, operationsFiltersChanged:element(),
    operationsChartsUnavailable:element(), operationsTrend:withRows?{id:'trend'}:null,
    operationsDistribution:{id:'distribution'},
    operationsChartData:{textContent:JSON.stringify(category?{...sample,series:[sample.series[0]]}:sample)}};
  const charts=[],events={};
  const context = {document:{getElementById:name=>nodes[name], querySelectorAll:()=>links},
    window:{addEventListener:(name,handler)=>{events[name]=handler;}}};
  if(chartAvailable) context.Chart=function(canvas,config){charts.push({canvas,config:JSON.parse(JSON.stringify(config))});};
  vm.runInNewContext(script,context);
  return {inputs,links,nodes,charts,events};
}
const app=run();
assert.equal(app.charts.length,2);
assert.deepEqual(app.charts[0].config.data.datasets.map(s=>s.data),sample.series.map(s=>s.values));
assert.deepEqual(app.charts[1].config.data.datasets.map(s=>s.data),Object.values(sample.states));
for(const input of app.inputs){
  if(input.type==='checkbox') input.checked=true; else input.value='cambiado';
  input.events.input();
  assert.equal(app.nodes.operationsFiltersChanged.hidden,false);
  for(const link of app.links){
    assert.equal(link.getAttribute('aria-disabled'),'true');
    assert.equal(link.tabIndex,-1);
    let blocked=false;link.events.click({preventDefault:()=>{blocked=true;}});assert.ok(blocked);
  }
  input.value=input.defaultValue;input.checked=input.defaultChecked;
  app.events.pageshow();
  assert.equal(app.nodes.operationsFiltersChanged.hidden,true);
  assert.ok(app.links.every(link=>link.getAttribute('aria-disabled')==='false'));
}
const category=run({category:true});
assert.deepEqual(category.charts[1].config.data.labels,['Empresa']);
assert.deepEqual(category.charts[1].config.data.datasets[0].data,[5]);
assert.equal(run({chartAvailable:false}).nodes.operationsChartsUnavailable.hidden,false);
assert.equal(run({withRows:false}).charts.length,0);
console.log('OK: gráficos, filtros, fechas vacías, regreso del navegador, sin Chart.js y sin datos.');
