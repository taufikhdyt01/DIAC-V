# config.py - Konfigurasi aplikasi DIAC-V

import os

# Konfigurasi jalur
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
DATA_DIR = os.path.join(BASE_DIR, "data")

# File data
USERS_DB = os.path.join(DATA_DIR, "users.xlsx")

# Konfigurasi aplikasi
APP_NAME = "DIAC-V"
APP_VERSION = "1.0.0"
APP_LOGO = os.path.join(ASSETS_DIR, "icons", "logo.png")

# Konfigurasi warna tema
PRIMARY_COLOR = "#2C3E50"  # Dark blue
SECONDARY_COLOR = "#3498DB"  # Light blue
ACCENT_COLOR = "#E74C3C"  # Red
BG_COLOR = "#ECF0F1"  # Light grey
TEXT_COLOR = "#2C3E50"  # Dark blue

# Konfigurasi departemen - menggunakan emoji sebagai pengganti ikon
DEPARTMENTS = [
    {"id": "ADE", "name": "ADE Group", "color": "#3498DB", "emoji": "👥"},
    {"id": "BDU", "name": "BDU Group", "color": "#2ECC71", "emoji": "📊"},
    {"id": "MAR", "name": "MAR Group", "color": "#E74C3C", "emoji": "📢"},
    {"id": "MAN", "name": "MAN Group", "color": "#F39C12", "emoji": "⚙️"},
    {"id": "PRJ", "name": "PRJ Group", "color": "#9B59B6", "emoji": "📋"},
    {"id": "FIN", "name": "FIN Group", "color": "#16A085", "emoji": "💰"},
    {"id": "LEG", "name": "LEG Group", "color": "#8E44AD", "emoji": "⚖️"}
]

# Konfigurasi hak akses
ACCESS_LEVELS = {
    "user": 1,  # Hanya akses ke grup sendiri
    "manager": 2,  # Akses ke grup sendiri dan beberapa grup terkait
    "director": 3,  # Akses ke banyak grup
    "admin": 4,  # Akses ke semua grup + fitur admin
    "ceo": 5  # Akses penuh
}