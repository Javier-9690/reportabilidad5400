(() => {
  const form = document.getElementById('userReportFilters');
  const exportLink = document.getElementById('exportUserReport');
  if (form && exportLink) {
    const inputs = [...form.querySelectorAll('input[type="date"]')];
    const original = inputs.map(input => input.defaultValue);
    const notice = document.getElementById('reportFiltersChanged');
    const update = () => {
      const changed = inputs.some((input, index) => input.value !== original[index]);
      notice.hidden = !changed;
      exportLink.classList.toggle('disabled', changed);
      exportLink.setAttribute('aria-disabled', String(changed));
      exportLink.tabIndex = changed ? -1 : 0;
    };
    inputs.forEach(input => { input.addEventListener('input', update); input.addEventListener('change', update); });
    exportLink.addEventListener('click', event => {
      if (exportLink.getAttribute('aria-disabled') === 'true') event.preventDefault();
    });
    window.addEventListener('pageshow', update);
    update();
  }

  const source = document.getElementById('userReportChartData');
  if (!source) return;
  if (typeof Chart === 'undefined') {
    document.getElementById('userChartsUnavailable').hidden = false;
    return;
  }
  const data = JSON.parse(source.textContent);
  const colors = ['#b42318', '#1667a5', '#198754'];
  new Chart(document.getElementById('userDailyChart'), {
    type: 'line',
    data: { labels: data.labels, datasets: data.categories.map((label, index) => ({
      label, data: data.daily[index], borderColor: colors[index], backgroundColor: colors[index],
      tension: 0, pointRadius: data.labels.length <= 31 ? 3 : 0, borderWidth: 2
    })) },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' }, tooltip: { callbacks: {
        title: items => data.dates[items[0].dataIndex]
      } } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } }, x: { ticks: { maxTicksLimit: 14 } } }
    }
  });
  new Chart(document.getElementById('userStatusChart'), {
    type: 'bar',
    data: { labels: data.categories, datasets: [
      { label: 'Abiertos', data: data.open, backgroundColor: '#ca8100' },
      { label: 'Cerrados', data: data.closed, backgroundColor: '#198754' },
      { label: 'Sin clasificar', data: data.unclassified, backgroundColor: '#777f87' }
    ] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: { x: { stacked: true, beginAtZero: true, ticks: { precision: 0 } }, y: { stacked: true } }
    }
  });
})();
