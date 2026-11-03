"""Aplicar parche al repositorio ORIGINAL. Ejecutar: python instalar_seguridad.py"""
from pathlib import Path
import shutil
import sys

ROOT = Path(__file__).resolve().parent


def replace_once(contents, original, replacement, file):
    if original not in contents:
        raise RuntimeError(f'No se encontró el bloque esperado en {file}; puede ser otra versión del repositorio.')
    if contents.count(original) != 1:
        raise RuntimeError(f'El bloque aparece varias veces en {file}. No se modificó.')
    return contents.replace(original, replacement, 1)


def modify(path, func):
    file = ROOT / path
    original = file.read_text(encoding='utf-8')
    changed = func(original)
    if changed == original:
        print(f'Sin cambios: {path}')
        return
    backup = file.with_suffix(file.suffix + '.antes_seguridad.bak')
    if not backup.exists():
        shutil.copyfile(file, backup)
    file.write_text(changed, encoding='utf-8')
    print(f'Actualizado: {path} (respaldo {backup.name})')


def modify_app(src):
    if 'install_security(app, gestion5s_app)' in src:
        return src
    needle = 'app.wsgi_app = DispatcherMiddleware(\n'
    replacement = ('from security import install_security\n'
                   'install_security(app, gestion5s_app)\n\n'
                   + needle)
    return replace_once(src, needle, replacement, 'app.py')


def modify_base(src):
    if '/static/js/admin-security.js' in src:
        return src
    script = '<script src="/static/js/admin-security.js"></script>\n'
    src = replace_once(src, '</body>', script + '</body>', 'base.html')
    logout = '''
    <form action="/logout" method="post" class="text-center my-2">
        <input type="hidden" name="csrf_token" value="{{ admin_logout_csrf }}">
        <button class="btn btn-sm btn-outline-danger" type="submit">Cerrar sesión</button>
    </form>
'''
    src = replace_once(src, '</header>', '</header>' + logout, 'base.html')
    return src


def main():
    if not (ROOT / 'app.py').exists() or not (ROOT / 'gestion5s/web.py').exists():
        sys.exit('ERROR: ejecuta este instalador en la raíz del repositorio original completo.')
    modify('app.py', modify_app)
    modify('templates/base.html', modify_base)
    modify('gestion5s/templates/base.html', modify_base)
    print('\nActualización aplicada. Configura ADMIN_PASSWORD y SECRET_KEY en Render.')
    print('No subas archivos *.bak a GitHub.')


if __name__ == '__main__':
    main()
