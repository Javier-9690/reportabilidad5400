(() => {
  const form = document.getElementById('bulkDeleteForm');
  if (!form) return;
  const boxes = [...document.querySelectorAll('.record-select[form="bulkDeleteForm"]')];
  const selectAll = document.getElementById('selectAllRecords');
  const deleteSelected = document.getElementById('deleteSelectedRecords');
  const countLabel = document.getElementById('selectedRecordCount');
  const buttonCount = document.getElementById('selectedButtonCount');

  function updateSelection() {
    const count = boxes.filter(box => box.checked).length;
    selectAll.checked = boxes.length > 0 && count === boxes.length;
    selectAll.indeterminate = count > 0 && count < boxes.length;
    deleteSelected.disabled = count === 0;
    countLabel.textContent = `${count} seleccionados`;
    buttonCount.textContent = count;
    boxes.forEach(box => box.closest('tr').classList.toggle('is-selected', box.checked));
  }

  selectAll.addEventListener('change', () => {
    boxes.forEach(box => { box.checked = selectAll.checked; });
    updateSelection();
  });
  boxes.forEach(box => box.addEventListener('change', updateSelection));
  window.addEventListener('pageshow', updateSelection);
  updateSelection();
})();

(() => {
  document.querySelectorAll('[data-record-table]').forEach(container => {
    const viewport = container.querySelector('[data-record-scroll]');
    const top = container.querySelector('[data-record-scroll-top]');
    const spacer = container.querySelector('[data-record-scroll-width]');
    const table = viewport.querySelector('table');
    function updateWidth() {
      spacer.style.width = `${viewport.scrollWidth}px`;
      top.style.width = `${viewport.clientWidth}px`;
      top.hidden = viewport.scrollWidth <= viewport.clientWidth + 1;
      top.scrollLeft = viewport.scrollLeft;
    }
    top.addEventListener('scroll', () => {
      if (viewport.scrollLeft !== top.scrollLeft) viewport.scrollLeft = top.scrollLeft;
    });
    viewport.addEventListener('scroll', () => {
      if (top.scrollLeft !== viewport.scrollLeft) top.scrollLeft = viewport.scrollLeft;
    });
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(updateWidth);
      observer.observe(viewport);
      observer.observe(table);
    }
    window.addEventListener('resize', updateWidth);
    window.addEventListener('pageshow', updateWidth);
    updateWidth();
  });
})();
