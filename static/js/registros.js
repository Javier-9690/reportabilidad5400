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
