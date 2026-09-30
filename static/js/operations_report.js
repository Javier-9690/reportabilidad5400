(() => {
  const form = document.getElementById('operationsFilters');
  if (form) {
    const inputs = [...form.querySelectorAll('input:not([type="hidden"])')];
    const links = [...document.querySelectorAll('[data-report-link]')];
    const notice = document.getElementById('operationsFiltersChanged');
    const update = () => {
      const changed = inputs.some(input => input.type === 'checkbox'
        ? input.checked !== input.defaultChecked : input.value !== input.defaultValue);
      notice.hidden = !changed;
      links.forEach(link => { link.setAttribute('aria-disabled', String(changed)); link.tabIndex = changed ? -1 : 0; });
    };
    inputs.forEach(input => ['input', 'change'].forEach(event => input.addEventListener(event, update)));
    links.forEach(link => link.addEventListener('click', event => {
      if (link.getAttribute('aria-disabled') === 'true') event.preventDefault();
    }));
    window.addEventListener('pageshow', update);
    update();
  }
  const source = document.getElementById('operationsChartData');
  const trend = document.getElementById('operationsTrend');
  if (!source || !trend) return;
  const warning = document.getElementById('operationsChartsUnavailable');
  if (typeof Chart === 'undefined') { warning.hidden = false; return; }
  const data = JSON.parse(source.textContent);
  const common = {
    locale: 'es-CL',
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { position: 'bottom', labels: { boxWidth: 9, boxHeight: 9, font: { size: 11 } } } },
    scales: { x: { ticks: { maxTicksLimit: 9, font: { size: 10 } }, grid: { display: false } },
      y: { beginAtZero: true, ticks: { precision: 0, font: { size: 11 } }, grid: { color: '#eef1f5' } } }
  };
  new Chart(trend, {
    type: 'line', data: { labels: data.labels, datasets: data.series.map(series => ({
      label: series.label, data: series.values, borderColor: series.color, backgroundColor: series.color,
      borderWidth: 2, tension: 0, pointRadius: data.labels.length <= 31 ? 2 : 0
    })) }, options: common
  });
  const distribution = document.getElementById('operationsDistribution');
  if (data.series.length === 1) {
    new Chart(distribution, { type: 'bar', data: {
      labels: data.companies.map(company => company.label),
      datasets: [{ label: 'Registros', data: data.companies.map(company => company.count), backgroundColor: data.series[0].color, borderRadius: 3 }]
    }, options: { ...common, indexAxis: 'y', plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true, ticks: { precision: 0 } }, y: { grid: { display: false }, ticks: { font: { size: 11 } } } } } });
  } else {
    new Chart(distribution, { type: 'bar', data: { labels: data.categories, datasets: [
      { label: 'Abiertos', data: data.states.open, backgroundColor: '#CF973C' },
      { label: 'Cerrados', data: data.states.closed, backgroundColor: '#278269' },
      { label: 'Por revisar', data: data.states.unknown, backgroundColor: '#8290A0' },
      { label: 'Sin campo de estado', data: data.states.not_applicable, backgroundColor: '#C5CDD7' }
    ] }, options: { ...common, indexAxis: 'y', scales: {
      x: { stacked: true, beginAtZero: true, ticks: { precision: 0 } },
      y: { stacked: true, grid: { display: false }, ticks: { font: { size: 11 } } }
    } } });
  }
})();
