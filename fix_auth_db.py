"""
fix_auth_db.py
Letakkan file ini di folder yang SAMA dengan auth.py dan app.py
Lalu jalankan: python fix_auth_db.py
"""

import sqlite3
import hashlib
import os
from datetime import datetime, timedelta

_HERE        = os.path.dirname(os.path.abspath(__file__))
AUTH_DB_FILE = os.path.join(_HERE, "auth_users.db")

def _hash(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def fix():
    print(f"\n{'='*55}")
    print(f"  FIX AUTH DATABASE")
    print(f"  Path: {AUTH_DB_FILE}")
    print(f"{'='*55}\n")

    conn = sqlite3.connect(AUTH_DB_FILE)
    c    = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            username      TEXT    UNIQUE NOT NULL,
            password_hash TEXT    NOT NULL,
            email         TEXT    DEFAULT '',
            role          TEXT    DEFAULT 'trial',
            trial_start   TEXT,
            trial_end     TEXT,
            is_active     INTEGER DEFAULT 1,
            created_at    TEXT
        )
    """)
    conn.commit()
    print("[OK] Tabel users siap")

    now        = datetime.now()
    admin_hash = _hash("admin123")
    far_future = (now + timedelta(days=36500)).isoformat()

    c.execute("SELECT id FROM users WHERE username='admin'")
    if c.fetchone():
        c.execute("""
            UPDATE users
            SET password_hash=?, role='admin', is_active=1, trial_end=?
            WHERE username='admin'
        """, (admin_hash, far_future))
        print("[OK] Admin ditemukan - password + role di-reset")
    else:
        c.execute("""
            INSERT INTO users
                (username, password_hash, email, role,
                 trial_start, trial_end, is_active, created_at)
            VALUES (?,?,?,?,?,?,1,?)
        """, ("admin", admin_hash, "admin@fnb.com", "admin",
              now.isoformat(), far_future, now.isoformat()))
        print("[OK] Admin dibuat baru")

    conn.commit()

    c.execute("SELECT password_hash FROM users WHERE username='admin'")
    stored = c.fetchone()[0]
    conn.close()

    if stored == admin_hash:
        print("\n" + "="*55)
        print("  BERHASIL! Login dengan:")
        print("     Username : admin")
        print("     Password : admin123")
        print("="*55 + "\n")
    else:
        print("\n  GAGAL - hash tidak cocok\n")

if __name__ == "__main__":
    fix()