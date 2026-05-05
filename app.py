"""
ASI QuickBooks Middleware
Receives invoice jobs from FileMaker, queues them for QuickBooks Web Connector.
"""

from flask import Flask, request, Response
import middleware.db as db
from middleware.wc_soap import handle_soap
from middleware.fm_routes import bp as fm_bp
from middleware.config import PORT


def create_app():
    app = Flask(__name__)

    db.init_db()
    app.register_blueprint(fm_bp)

    @app.post("/wc")
    def web_connector():
        """SOAP endpoint for QuickBooks Web Connector."""
        result = handle_soap(request.data.decode("utf-8"))
        return Response(result, content_type="text/xml; charset=utf-8")

    @app.get("/health")
    def health():
        return {"status": "ok"}

    return app


if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=PORT, debug=True)
