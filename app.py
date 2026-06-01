"""
ASI QuickBooks Middleware
Receives invoice jobs from FileMaker and submits them to QuickBooks via COM/QBFC.
"""

from flask import Flask
from middleware.fm_routes import bp as fm_bp
from middleware.config import PORT
from middleware import logger as _logger
from version import VERSION, BUILD

try:
    import pythoncom
    pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
except ImportError:
    pass


def create_app():
    app = Flask(__name__)

    # Send UTF-8 characters directly rather than \uXXXX JSON escape sequences.
    # FileMaker 19 mishandles Unicode escapes for non-ASCII characters (e.g.
    # ’ for the right single quote decodes to the Windows-1252 byte 0x92,
    # a C1 control character, rather than the proper UTF-8 sequence). FM handles
    # raw UTF-8 from Insert from URL correctly, so this avoids the mangling.
    app.json.ensure_ascii = False

    _logger.ensure_logs_dir()
    _logger.trim_all_logs()
    app.register_blueprint(fm_bp)

    @app.get("/health")
    def health():
        return {"status": "ok", "version": VERSION, "build": BUILD}

    return app


if __name__ == "__main__":
    app = create_app()
    # threaded=False: all requests on the main thread, same as standalone Python.
    # debug=False: disables Werkzeug reloader subprocess which complicates COM access.
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=False)
