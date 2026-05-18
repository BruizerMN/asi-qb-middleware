import os
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.environ["API_KEY"]
PORT = int(os.environ.get("PORT", 5100))

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "jobs.db")
