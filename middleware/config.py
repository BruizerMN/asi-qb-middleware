import os
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.environ["API_KEY"]
PORT = int(os.environ.get("PORT", 5100))

# Web Connector credentials keyed by company slug
WC_CREDENTIALS = {
    "acoustical": {
        "user": os.environ["WC_USER_ACOUSTICAL"],
        "password": os.environ["WC_PASS_ACOUSTICAL"],
    },
    "architectural": {
        "user": os.environ["WC_USER_ARCHITECTURAL"],
        "password": os.environ["WC_PASS_ARCHITECTURAL"],
    },
}

# Map WC username → company slug (built from WC_CREDENTIALS)
WC_USER_TO_COMPANY = {v["user"]: k for k, v in WC_CREDENTIALS.items()}

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "jobs.db")
