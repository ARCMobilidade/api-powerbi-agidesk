import dotenv
import os

dotenv.load_dotenv()

DB_USER = os.getenv("DB_USERNAME")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("HOSTNAME")
KEY = os.getenv("API_KEY")