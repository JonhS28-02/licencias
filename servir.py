"""Servidor local de desarrollo; no sustituye autenticación para producción."""
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, unquote

ROOT = Path(__file__).resolve().parent
PUBLIC_FILES = {
    '/', '/index.html', '/styles.css', '/cs.js', '/seguridad.js',
    '/eventos.js', '/constancia.js', '/data.json', '/TEMPLATE_GLOBALES.xlsx',
}


class LocalHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def send_head(self):
        if unquote(urlsplit(self.path).path) not in PUBLIC_FILES:
            self.send_error(404)
            return None
        return super().send_head()

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('X-Frame-Options', 'DENY')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('Cross-Origin-Resource-Policy', 'same-origin')
        super().end_headers()

    def log_message(self, format, *args):
        # No registrar rutas, parámetros de búsqueda ni información de personas.
        pass


if __name__ == '__main__':
    print('Consulta local: http://127.0.0.1:8000 · Ctrl+C para detener')
    with ThreadingHTTPServer(('127.0.0.1', 8000), LocalHandler) as server:
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
