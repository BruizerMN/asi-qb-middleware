import os
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.environ["API_KEY"]
PORT = int(os.environ.get("PORT", 5100))

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "jobs.db")

# QB company file paths, one per company slug.
# These are set in .env and point to the .QBW files on the Q: drive.
# To switch from test to production company files, update .env only -- no code changes needed.
QB_COMPANY_FILES = {
    "acoustical":    os.environ.get("QB_FILE_ACOUSTICAL", ""),
    "architectural": os.environ.get("QB_FILE_ARCHITECTURAL", ""),
}
