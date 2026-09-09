(() => {
  document.querySelectorAll('[data-census-form]').forEach(form => {
    const automatic = form.querySelector('[name="censo_total_auto"]');
    const day = form.querySelector('[name="censo_dia"]');
    const night = form.querySelector('[name="censo_noche"]');
    const total = form.querySelector('[name="total"]');
    function updateTotal() {
      total.readOnly = automatic.checked;
      if (automatic.checked) {
        const sum = Number(day.value || 0) + Number(night.value || 0);
        total.value = Number.isFinite(sum) ? String(sum) : '';
      }
    }
    automatic.addEventListener('change', updateTotal);
    day.addEventListener('input', updateTotal);
    night.addEventListener('input', updateTotal);
    updateTotal();
  });

  document.querySelectorAll('[data-pdf-name-target]').forEach(picker => {
    picker.addEventListener('change', () => {
      const target = document.getElementById(picker.dataset.pdfNameTarget);
      if (target && picker.files.length) target.value = picker.files[0].name;
    });
  });
})();
