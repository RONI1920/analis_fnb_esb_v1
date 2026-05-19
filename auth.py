"""
auth.py — Sistem Auth Login + Trial 3 Hari
Letakkan di folder yang sama dengan app.py.
"""

from __future__ import annotations

import os
import sqlite3
import hashlib
import streamlit as st

from datetime import datetime, timedelta


# ─────────────────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────────────────

_HERE = os.path.dirname(os.path.abspath(__file__))

AUTH_DB_FILE = os.path.join(_HERE, "auth_users.db")

TRIAL_DAYS = 3
SESSION_KEY = "auth_session"

# Toggle fitur register
ENABLE_REGISTER = False

# Tampilkan akun admin default
SHOW_DEFAULT_ADMIN = True


# ─────────────────────────────────────────────────────────
# HELPER
# ─────────────────────────────────────────────────────────

def _hash_password(password: str) -> str:
    return hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()


def _get_auth_conn() -> sqlite3.Connection:
    return sqlite3.connect(AUTH_DB_FILE)


# ─────────────────────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────────────────────

def init_auth_db() -> None:

    try:

        conn = _get_auth_conn()
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                email TEXT DEFAULT '',
                role TEXT DEFAULT 'trial',
                trial_start TEXT,
                trial_end TEXT,
                is_active INTEGER DEFAULT 1,
                created_at TEXT
            );
        """)

        conn.commit()

        # Seed admin default
        cursor.execute(
            "SELECT id FROM users WHERE username='admin'"
        )

        if not cursor.fetchone():

            now = datetime.now()

            cursor.execute("""
                INSERT INTO users (
                    username,
                    password_hash,
                    email,
                    role,
                    trial_start,
                    trial_end,
                    is_active,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, 1, ?)
            """, (
                "admin",
                _hash_password("admin123"),
                "admin@fnb.com",
                "admin",
                now.isoformat(),
                (now + timedelta(days=36500)).isoformat(),
                now.isoformat(),
            ))

            conn.commit()

        conn.close()

    except Exception as e:

        st.error(
            f"❌ Gagal membuat database auth.\n\n"
            f"Path: {AUTH_DB_FILE}\n\n"
            f"Error: {e}"
        )

        st.stop()


# ─────────────────────────────────────────────────────────
# CRUD USERS
# ─────────────────────────────────────────────────────────

def register_user(
    username: str,
    password: str,
    email: str = "",
    role: str = "trial"
) -> dict:

    try:

        conn = _get_auth_conn()
        cursor = conn.cursor()

        cursor.execute(
            "SELECT id FROM users WHERE username=?",
            (username.strip(),)
        )

        if cursor.fetchone():

            conn.close()

            return {
                "ok": False,
                "msg": "Username sudah digunakan."
            }

        now = datetime.now()

        cursor.execute("""
            INSERT INTO users (
                username,
                password_hash,
                email,
                role,
                trial_start,
                trial_end,
                is_active,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, 1, ?)
        """, (
            username.strip(),
            _hash_password(password),
            email.strip(),
            role,
            now.isoformat(),
            (now + timedelta(days=TRIAL_DAYS)).isoformat(),
            now.isoformat(),
        ))

        conn.commit()
        conn.close()

        return {
            "ok": True,
            "msg": "Registrasi berhasil!"
        }

    except Exception as e:

        return {
            "ok": False,
            "msg": f"Error database: {e}"
        }


def get_user(username: str) -> dict | None:

    try:

        conn = _get_auth_conn()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                id,
                username,
                password_hash,
                email,
                role,
                trial_start,
                trial_end,
                is_active
            FROM users
            WHERE username=?
        """, (username,))

        row = cursor.fetchone()

        conn.close()

        if not row:
            return None

        return {
            "id": row[0],
            "username": row[1],
            "password_hash": row[2],
            "email": row[3] or "",
            "role": row[4],
            "trial_start": row[5],
            "trial_end": row[6],
            "is_active": row[7],
        }

    except Exception:
        return None


def upgrade_user_to_premium(username: str) -> None:

    try:

        conn = _get_auth_conn()

        conn.execute(
            "UPDATE users SET role='premium' WHERE username=?",
            (username,)
        )

        conn.commit()
        conn.close()

    except Exception:
        pass


def deactivate_user(username: str) -> None:

    try:

        conn = _get_auth_conn()

        conn.execute(
            "UPDATE users SET is_active=0 WHERE username=?",
            (username,)
        )

        conn.commit()
        conn.close()

    except Exception:
        pass


def activate_user(username: str) -> None:

    try:

        conn = _get_auth_conn()

        conn.execute(
            "UPDATE users SET is_active=1 WHERE username=?",
            (username,)
        )

        conn.commit()
        conn.close()

    except Exception:
        pass


def extend_trial(
    username: str,
    extra_days: int = 3
) -> None:

    try:

        conn = _get_auth_conn()
        cursor = conn.cursor()

        cursor.execute(
            "SELECT trial_end FROM users WHERE username=?",
            (username,)
        )

        row = cursor.fetchone()

        if row:

            try:
                base = max(
                    datetime.fromisoformat(row[0]),
                    datetime.now()
                )

            except Exception:
                base = datetime.now()

            new_end = base + timedelta(days=extra_days)

            conn.execute(
                "UPDATE users SET trial_end=? WHERE username=?",
                (new_end.isoformat(), username)
            )

            conn.commit()

        conn.close()

    except Exception:
        pass


def list_all_users() -> list:

    try:

        conn = _get_auth_conn()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                id,
                username,
                email,
                role,
                trial_start,
                trial_end,
                is_active,
                created_at
            FROM users
            ORDER BY created_at DESC
        """)

        rows = cursor.fetchall()

        conn.close()

        return [
            {
                "id": r[0],
                "username": r[1],
                "email": r[2] or "",
                "role": r[3],
                "trial_start": r[4],
                "trial_end": r[5],
                "is_active": r[6],
                "created_at": r[7],
            }
            for r in rows
        ]

    except Exception:
        return []


# ─────────────────────────────────────────────────────────
# AKSES
# ─────────────────────────────────────────────────────────

def check_access(user: dict) -> dict:

    if not user.get("is_active"):

        return {
            "allowed": False,
            "reason": "⛔ Akun dinonaktifkan.",
            "days_left": 0
        }

    if user.get("role") in ("premium", "admin"):

        return {
            "allowed": True,
            "reason": "Akses penuh.",
            "days_left": 9999
        }

    try:

        trial_end = datetime.fromisoformat(
            user["trial_end"]
        )

    except Exception:

        return {
            "allowed": False,
            "reason": "Data trial tidak valid.",
            "days_left": 0
        }

    days_left = (trial_end - datetime.now()).days

    if datetime.now() > trial_end:

        return {
            "allowed": False,
            "reason": "⏰ Masa trial berakhir.",
            "days_left": 0
        }

    return {
        "allowed": True,
        "reason": f"Trial aktif — {days_left + 1} hari tersisa.",
        "days_left": days_left + 1
    }


def authenticate(
    username: str,
    password: str
) -> dict:

    user = get_user(username.strip())

    if not user:

        return {
            "ok": False,
            "msg": "Username tidak ditemukan."
        }

    if user["password_hash"] != _hash_password(password):

        return {
            "ok": False,
            "msg": "Password salah."
        }

    access = check_access(user)

    if not access["allowed"]:

        return {
            "ok": False,
            "msg": access["reason"]
        }

    return {
        "ok": True,
        "user": user,
        "access": access
    }


# ─────────────────────────────────────────────────────────
# UI LOGIN
# ─────────────────────────────────────────────────────────

def _render_login_page() -> None:

    st.markdown(
        """
        <div style='text-align:center;padding:2rem 0 1rem'>
            <h1>🍽️ Data Driven F&B Analyst</h1>
            <p style='color:#aaa;font-size:1.1rem'>
                Silakan login untuk mengakses sistem
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ─────────────────────────────────────
    # LOGIN ONLY
    # ─────────────────────────────────────

    if not ENABLE_REGISTER:

        with st.form(
            "form_login",
            clear_on_submit=False
        ):

            uname = st.text_input(
                "Username",
                placeholder="Masukkan username Anda"
            )

            pwd = st.text_input(
                "Password",
                type="password",
                placeholder="Masukkan password Anda"
            )

            ok = st.form_submit_button(
                "🔓 Login",
                use_container_width=True
            )

        if ok:

            if not uname or not pwd:

                st.error(
                    "Username dan password tidak boleh kosong."
                )

            else:

                result = authenticate(uname, pwd)

                if result["ok"]:

                    st.session_state[SESSION_KEY] = {
                        "logged_in": True,
                        "username": result["user"]["username"],
                        "role": result["user"]["role"],
                        "access": result["access"],
                    }

                    st.success(
                        f"✅ Selamat datang "
                        f"{result['user']['username']}"
                    )

                    st.rerun()

                else:
                    st.error(result["msg"])

    # ─────────────────────────────────────
    # LOGIN + REGISTER
    # ─────────────────────────────────────

    else:

        tab_login, tab_register = st.tabs([
            "🔑 Login",
            "📝 Daftar Akun Baru"
        ])

        # LOGIN TAB
        with tab_login:

            with st.form(
                "form_login",
                clear_on_submit=False
            ):

                uname = st.text_input(
                    "Username",
                    placeholder="Masukkan username Anda"
                )

                pwd = st.text_input(
                    "Password",
                    type="password",
                    placeholder="Masukkan password Anda"
                )

                ok = st.form_submit_button(
                    "🔓 Login",
                    use_container_width=True
                )

            if ok:

                if not uname or not pwd:

                    st.error(
                        "Username dan password tidak boleh kosong."
                    )

                else:

                    result = authenticate(uname, pwd)

                    if result["ok"]:

                        st.session_state[SESSION_KEY] = {
                            "logged_in": True,
                            "username": result["user"]["username"],
                            "role": result["user"]["role"],
                            "access": result["access"],
                        }

                        st.success(
                            f"✅ Selamat datang "
                            f"{result['user']['username']}"
                        )

                        st.rerun()

                    else:
                        st.error(result["msg"])

        # REGISTER TAB
        with tab_register:

            st.info(
                f"🎁 Trial gratis "
                f"{TRIAL_DAYS} hari"
            )

            with st.form(
                "form_register",
                clear_on_submit=True
            ):

                nu = st.text_input(
                    "Username baru"
                )

                ne = st.text_input(
                    "Email"
                )

                np = st.text_input(
                    "Password",
                    type="password"
                )

                np2 = st.text_input(
                    "Konfirmasi Password",
                    type="password"
                )

                reg = st.form_submit_button(
                    "🚀 Daftar",
                    use_container_width=True
                )

            if reg:

                err = None

                if not nu or not np:
                    err = "Username dan password wajib."

                elif len(nu.strip()) < 3:
                    err = "Username minimal 3 karakter."

                elif len(np) < 6:
                    err = "Password minimal 6 karakter."

                elif np != np2:
                    err = "Password tidak cocok."

                if err:

                    st.error(err)

                else:

                    res = register_user(
                        nu.strip(),
                        np,
                        ne
                    )

                    if res["ok"]:

                        st.success(
                            f"Akun {nu} berhasil dibuat."
                        )

                    else:

                        st.error(res["msg"])

    # ─────────────────────────────────────
    # BAGIAN BAWAH
    # ─────────────────────────────────────

    if SHOW_DEFAULT_ADMIN:

        with st.expander("ℹ️ Akun Default"):

            st.info(
                "**Username:** `admin`  \n"
                "**Password:** `admin123`"
            )

    st.markdown("---")

    st.caption(
        "💬 Butuh upgrade Premium? Hubungi admin."
    )


# ─────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────

def _render_trial_banner(session: dict) -> None:

    role = session.get("role", "trial")

    days = session.get(
        "access",
        {}
    ).get("days_left", 0)

    name = session.get("username", "")

    if role == "admin":

        st.success(
            "🛡️ Login sebagai Admin Trail"
        )

    elif role == "premium":

        st.success(
            f"⭐ Premium aktif — {name}"
        )

    elif days <= 1:

        st.error(
            f"⏰ Trial tinggal "
            f"{days} hari lagi"
        )

    elif days <= 3:

        st.warning(
            f"🕐 Trial aktif — "
            f"{days} hari tersisa"
        )

    else:

        st.info(
            f"🎉 Trial aktif — "
            f"{days} hari tersisa"
        )


# ─────────────────────────────────────────────────────────
# REQUIRE AUTH
# ─────────────────────────────────────────────────────────

def require_auth() -> dict:

    init_auth_db()

    session = st.session_state.get(
        SESSION_KEY
    )

    if not session or not session.get("logged_in"):

        _render_login_page()

        st.stop()

        return None

    user = get_user(session["username"])

    if not user:

        st.session_state.pop(
            SESSION_KEY,
            None
        )

        st.error(
            "Sesi tidak valid."
        )

        _render_login_page()

        st.stop()

        return None

    access = check_access(user)

    if not access["allowed"]:

        st.session_state.pop(
            SESSION_KEY,
            None
        )

        st.error(access["reason"])

        _render_login_page()

        st.stop()

        return None

    # update session terbaru
    session["access"] = access
    session["role"] = user["role"]

    st.session_state[SESSION_KEY] = session

    _render_trial_banner(session)

    # SIDEBAR
    with st.sidebar:

        st.markdown("---")

        icons = {
            "admin": "🛡️",
            "premium": "⭐",
            "trial": "🕐"
        }

        st.caption(
            f"{icons.get(session['role'], '👤')} "
            f"{session['username']} "
            f"({session['role']})"
        )

        if session["role"] == "trial":

            st.caption(
                f"Sisa trial: "
                f"{access.get('days_left', 0)} hari"
            )

        if st.button(
            "🚪 Logout",
            use_container_width=True
        ):

            st.session_state.pop(
                SESSION_KEY,
                None
            )

            st.rerun()

    return session