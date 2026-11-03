/* Pide contraseña por operación destructiva; la validación real ocurre en Flask. */
(() => {
  'use strict';
  const approved = new WeakSet();
  const protectedOperation = (form) => {
    if (form.method.toUpperCase() !== 'POST') return false;
    const pathname = new URL(form.action, location.href).pathname
      .replace(/^\/gestion-5s(?=\/)/, '');
    return /^\/api\/files\/\d+\/delete$/.test(pathname) ||
      /^\/delete\/[^/]+\/(?:\d+|bulk)$/.test(pathname) ||
      pathname === '/import/ordenes/confirm';
  };

  const dialog = document.createElement('dialog');
  dialog.style.cssText = 'border:0;border-radius:14px;max-width:420px;width:calc(100% - 36px);padding:28px;box-shadow:0 20px 70px #0007';
  dialog.innerHTML = `<form method="dialog" id="r5400-dialog-form">
    <h3 style="color:#a6192e;font-size:1.25rem">Autorización de administrador</h3>
    <p>Para eliminar o reemplazar estos registros, ingresa la clave nuevamente.</p>
    <label for="r5400-confirm-password">Contraseña</label>
    <input id="r5400-confirm-password" type="password" autocomplete="off" required
      style="width:100%;padding:10px;border:1px solid #aaa;border-radius:6px;box-sizing:border-box">
    <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:20px">
      <button type="button" id="r5400-cancel" style="padding:10px">Cancelar</button>
      <button type="submit" style="padding:10px;background:#a6192e;color:white;border:0;border-radius:6px">Autorizar</button>
    </div>
  </form>`;
  document.body.appendChild(dialog);
  const input = dialog.querySelector('#r5400-confirm-password');
  let pendingForm = null;
  let pendingSubmitter = null;
  dialog.querySelector('#r5400-cancel').addEventListener('click', () => dialog.close());
  dialog.addEventListener('close', () => { input.value = ''; pendingForm = null; pendingSubmitter = null; });
  dialog.querySelector('form').addEventListener('submit', (event) => {
    event.preventDefault();
    if (!pendingForm || !input.value) return;
    const form = pendingForm;
    const submitter = pendingSubmitter;
    let field = form.querySelector('input[name="admin_password"]');
    if (!field) {
      field = document.createElement('input');
      field.type = 'hidden';
      field.name = 'admin_password';
      form.appendChild(field);
    }
    field.value = input.value;
    dialog.close();
    approved.add(form);
    try {
      if (submitter && submitter.form === form) form.requestSubmit(submitter);
      else form.requestSubmit();
    } finally {
      approved.delete(form);
      // Campo queda presente durante el POST; la página se recarga después.
    }
  });
  // Captura el submit antes de otros listeners; preserva botón y campos.
  document.addEventListener('submit', (event) => {
    const form = event.target;
    if (!(form instanceof HTMLFormElement) || !protectedOperation(form)) return;
    if (approved.has(form)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    pendingForm = form;
    pendingSubmitter = event.submitter || null;
    dialog.showModal();
    input.focus();
  }, true);
})();
