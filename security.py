"""Acceso administrador y autorización por operación para Reportabilidad 5400."""
import hmac
import os
import re
import secrets
from datetime import timedelta
from urllib.parse import urlsplit, quote

from flask import abort, redirect, render_template, request, session, url_for

ADMIN_SESSION_KEY = 'r5400_admin_authenticated'


def verify_password(candidate):
    expected = os.environ.get('ADMIN_PASSWORD', '')
    if not expected:
        raise RuntimeError('Configura ADMIN_PASSWORD en las variables de entorno de Render.')
    return hmac.compare_digest(str(candidate).encode('utf-8'), expected.encode('utf-8'))


def safe_next(target):
    target = (target or '/').strip()
    parsed = urlsplit(target)
    if (not target.startswith('/') or target.startswith('//') or
            parsed.scheme or parsed.netloc or '\\' in target or
            any(ord(c) < 32 for c in target)):
        return '/'
    return target


def destructive_path(path):
    # Las rutas de hotelería llegan sin el prefijo /gestion-5s a su sub-app.
    return bool(
        re.fullmatch(r'/api/files/\d+/delete', path) or
        re.fullmatch(r'/delete/[^/]+/(?:\d+|bulk)', path) or
        path == '/import/ordenes/confirm'
    )


def install_security(main_app, hotel_app, checklist_app=None):
    if not os.getenv('ADMIN_PASSWORD'):
        raise RuntimeError('ADMIN_PASSWORD es obligatorio; no se inicia sin esa variable.')
    if main_app.secret_key in ('dev-change-me', 'dev-secret', None, ''):
        raise RuntimeError('Configura SECRET_KEY con un valor aleatorio y seguro.')

    # La cookie debe ser idéntica para las dos aplicaciones (ambas montadas en el mismo dominio).
    hotel_app.secret_key = main_app.secret_key
    if checklist_app is not None:
        checklist_app.secret_key = main_app.secret_key

    applications = [main_app, hotel_app]
    if checklist_app is not None:
        applications.append(checklist_app)
    for application in applications:
        application.config.update(
            SESSION_COOKIE_NAME='r5400_session',
            SESSION_COOKIE_HTTPONLY=True,
            SESSION_COOKIE_SAMESITE='Lax',
            SESSION_COOKIE_PATH='/',
            SESSION_COOKIE_SECURE=(os.getenv('RENDER') is not None or
                                   os.getenv('SESSION_COOKIE_SECURE', '').lower() == 'true'),
            PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
        )

        def make_guard(app_instance):
            def guard():
                if app_instance is main_app and request.path == '/login':
                    return None
                if request.path.startswith('/static/'):
                    return None
                if not session.get(ADMIN_SESSION_KEY):
                    if request.method == 'GET':
                        destination = request.full_path.rstrip('?')
                        # El endpoint admin_login vive en la app principal, no en hotel_app.
                        # Usar ruta absoluta también preserva /gestion-5s en el retorno.
                        if app_instance is hotel_app:
                            destination = '/gestion-5s' + destination
                        elif app_instance is checklist_app:
                            destination = '/camp-checklist' + destination
                        return redirect('/login?next=' + quote(destination, safe=''))
                    abort(401, description='Inicia sesión para continuar.')
                if request.method == 'POST' and destructive_path(request.path):
                    if not verify_password(request.form.get('admin_password', '')):
                        abort(403, description='Clave administrativa incorrecta. No se eliminó ningún registro.')
                return None
            return guard

        application.before_request(make_guard(application))

    @main_app.route('/login', methods=['GET', 'POST'])
    def admin_login():
        destination = safe_next(request.values.get('next'))
        if destination.startswith('/login') or destination.startswith('/logout'):
            destination = '/'
        if 'r5400_login_csrf' not in session:
            session['r5400_login_csrf'] = secrets.token_urlsafe(32)
        error = None
        if request.method == 'POST':
            token = request.form.get('csrf_token', '')
            if not hmac.compare_digest(token, session['r5400_login_csrf']):
                abort(400, description='Formulario de acceso expirado.')
            username = request.form.get('username', '').strip()
            if (hmac.compare_digest(username, 'admin') and
                    verify_password(request.form.get('password', ''))):
                session.clear()
                session[ADMIN_SESSION_KEY] = True
                session.permanent = True
                return redirect(destination)
            error = 'Usuario o contraseña incorrectos.'
        elif session.get(ADMIN_SESSION_KEY):
            return redirect(destination)
        return render_template('admin_login.html', error=error,
                               csrf_token=session['r5400_login_csrf'],
                               next_url=destination), (401 if error else 200)

    @main_app.post('/logout')
    def admin_logout():
        token = request.form.get('csrf_token', '')
        if not hmac.compare_digest(token, session.get('r5400_logout_csrf', '')):
            abort(400, description='Solicitud inválida.')
        session.clear()
        return redirect('/login')

    @main_app.context_processor
    def inject_logout_token():
        if 'r5400_logout_csrf' not in session:
            session['r5400_logout_csrf'] = secrets.token_urlsafe(32)
        return {'admin_logout_csrf': session['r5400_logout_csrf']}

    @hotel_app.context_processor
    def inject_hotel_logout_token():
        if 'r5400_logout_csrf' not in session:
            session['r5400_logout_csrf'] = secrets.token_urlsafe(32)
        return {'admin_logout_csrf': session['r5400_logout_csrf']}

    if checklist_app is not None:
        @checklist_app.context_processor
        def inject_checklist_logout_token():
            if 'r5400_logout_csrf' not in session:
                session['r5400_logout_csrf'] = secrets.token_urlsafe(32)
            return {'admin_logout_csrf': session['r5400_logout_csrf']}

    return main_app
