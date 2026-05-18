"""
ASI QuickBooks Middleware
Receives invoice jobs from FileMaker and submits them to QuickBooks via COM/QBFC.
"""

from flask import Flask
import middleware.db as db
from middleware.fm_routes import bp as fm_bp
from middleware.config import PORT
from version import VERSION, BUILD

try:
    import pythoncom
    pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
except ImportError:
    pass


def create_app():
    app = Flask(__name__)

    db.init_db()
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
