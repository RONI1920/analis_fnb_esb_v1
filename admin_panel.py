"""
admin_panel.py — Panel Admin Kelola User
Panggil build_admin_panel() di halaman Admin pada app.py.
"""

from __future__ import annotations

import os
import sqlite3
import streamlit as st
import pandas as pd
from datetime import datetime
from auth import (
    AUTH_DB_FILE,
    SESSION_KEY,
    TRIAL_DAYS,
    _hash_password,
    list_all_users,
    upgrade_user_to_premium,
    deactivate_user,
    activate_user,
    extend_trial,
    register_user,
)


# ─────────────────────────────────────────────────────────
# HELPER INTERNAL
# ─────────────────────────────────────────────────────────

def _conn() -> sqlite3.Connection:
    return sqlite3.connect(AUTH_DB_FILE)


def _change_password(username: str, new_password: str) -> None:
    try:
        conn = _conn()
        conn.execute(
            "UPDATE users SET password_hash=? WHERE username=?",
            (_hash_password(new_password), username)
        )
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Gagal ganti password: {e}")


def _delete_user(username: str) -> None:
    try:
        conn = _conn()
        conn.execute("DELETE FROM users WHERE username=?", (username,))
        conn.commit()
        conn.close()
    except Exception as e:
        st.error(f"Gagal hapus user: {e}")


def _fmt_dt(dt_str: str) -> str:
    if not dt_str:
        return "-"
    try:
        return datetime.fromisoformat(dt_str).strftime("%d %b %Y  %H:%M")
    except Exception:
        return str(dt_str)


def _sisa_hari(u: dict) -> str:
    if u["role"] in ("premium", "admin"):
        return "∞  Penuh"
    if not u["is_active"]:
        return "🚫 Nonaktif"
    try:
        end  = datetime.fromisoformat(u["trial_end"])
        sisa = (end - datetime.now()).days
        return "⏰ Expired" if sisa < 0 else f"{sisa + 1} hari"
    except Exception:
        return "-"


# ─────────────────────────────────────────────────────────
# PANEL UTAMA
# ─────────────────────────────────────────────────────────

def build_admin_panel() -> None:
    """Hanya bisa diakses jika role = 'admin'."""
    session = st.session_state.get(SESSION_KEY, {})

    if session.get("role") != "admin":
        st.warning("🔒 Halaman ini hanya untuk Admin.")
        return

    st.header("🛠️ Panel Admin — Manajemen User")

    raw_users = list_all_users()

    # ── STATISTIK ──────────────────────────────────────────
    st.subheader("📊 Statistik User")

    total    = len(raw_users)
    premiums = sum(1 for u in raw_users if u["role"] == "premium")
    admins   = sum(1 for u in raw_users if u["role"] == "admin")
    nonaktif = sum(1 for u in raw_users if not u["is_active"])

    expired = 0
    for u in raw_users:
        if u["role"] == "trial" and u["is_active"] and u["trial_end"]:
            try:
                if datetime.fromisoformat(u["trial_end"]) < datetime.now():
                    expired += 1
            except Exception:
                pass

    trials_aktif = sum(1 for u in raw_users if u["role"] == "trial") - expired

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("👥 Total",        total)
    c2.metric("⭐ Premium",       premiums)
    c3.metric("🕐 Trial Aktif",  trials_aktif)
    c4.metric("⏰ Trial Expired", expired)
    c5.metric("🚫 Nonaktif",     nonaktif)

    st.markdown("---")

    # ── TABEL USER ─────────────────────────────────────────
    st.subheader("📋 Daftar Semua User")

    if raw_users:
        rows = [
            {
                "username":       u["username"],
                "email":          u["email"] or "-",
                "role":           u["role"],
                "status":         "✅ Aktif" if u["is_active"] else "❌ Nonaktif",
                "Sisa Trial":     _sisa_hari(u),
                "Tgl Daftar":     _fmt_dt(u["created_at"]),
                "Trial Berakhir": _fmt_dt(u["trial_end"]),
                "_role":          u["role"],
                "_active":        u["is_active"],
            }
            for u in raw_users
        ]
        df = pd.DataFrame(rows)

        col1, col2 = st.columns(2)
        with col1:
            fr = st.selectbox("Filter Role:",
                              ["Semua", "trial", "premium", "admin"],
                              key="adm_fr")
        with col2:
            fs = st.selectbox("Filter Status:",
                              ["Semua", "Aktif", "Nonaktif"],
                              key="adm_fs")

        dv = df.copy()
        if fr != "Semua":
            dv = dv[dv["_role"] == fr]
        if fs == "Aktif":
            dv = dv[dv["_active"] == 1]
        elif fs == "Nonaktif":
            dv = dv[dv["_active"] == 0]

        st.dataframe(
            dv[["username", "email", "role", "status",
                "Sisa Trial", "Tgl Daftar", "Trial Berakhir"]],
            use_container_width=True,
            hide_index=True,
        )

        csv = dv[["username", "email", "role", "status",
                "Sisa Trial", "Tgl Daftar", "Trial Berakhir"]
            ].to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download CSV",
            data=csv,
            file_name=f"users_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )
    else:
        st.info("Belum ada user.")

    st.markdown("---")

    # ── AKSI PER USER ──────────────────────────────────────
    st.subheader("⚡ Aksi per User")

    usernames = [u["username"] for u in raw_users if u["username"] != "admin"]
    if not usernames:
        st.info("Belum ada user selain admin.")
        return

    sel = st.selectbox("Pilih Username:", usernames, key="adm_sel")
    sel_data = next((u for u in raw_users if u["username"] == sel), None)

    if sel_data:
        with st.container(border=True):
            a, b, c, d = st.columns(4)
            a.markdown(f"**Role:** `{sel_data['role']}`")
            b.markdown(f"**Status:** {'✅' if sel_data['is_active'] else '❌'}")
            c.markdown(f"**Email:** {sel_data['email'] or '-'}")
            d.markdown(f"**Expired:** {_fmt_dt(sel_data['trial_end'])}")

    t1, t2, t3, t4, t5 = st.tabs([
        "⬆️ Upgrade", "📅 Perpanjang Trial",
        "🔘 Aktif/Non", "🔑 Password", "🗑️ Hapus",
    ])

    with t1:
        st.write(f"Upgrade **{sel}** → Premium (tanpa batas waktu).")
        if st.button("⬆️ Upgrade ke Premium", key="adm_upgrade"):
            upgrade_user_to_premium(sel)
            st.success(f"✅ **{sel}** berhasil di-upgrade ke Premium!")
            st.rerun()

    with t2:
        extra = st.number_input("Tambah berapa hari?", 1, 365, 3, key="adm_days")
        if st.button("📅 Perpanjang Trial", key="adm_extend"):
            extend_trial(sel, extra)
            st.success(f"✅ Trial **{sel}** diperpanjang {extra} hari.")
            st.rerun()

    with t3:
        if sel_data and sel_data["is_active"]:
            st.warning(f"Nonaktifkan **{sel}**? User tidak bisa login.")
            if st.button("🚫 Nonaktifkan", key="adm_deact"):
                deactivate_user(sel)
                st.success(f"**{sel}** dinonaktifkan.")
                st.rerun()
        else:
            if st.button("✅ Aktifkan Kembali", key="adm_act"):
                activate_user(sel)
                st.success(f"**{sel}** diaktifkan kembali.")
                st.rerun()

    with t4:
        with st.form("form_pwd"):
            np  = st.text_input("Password Baru", type="password",
                                placeholder="min. 6 karakter")
            np2 = st.text_input("Konfirmasi", type="password")
            ok  = st.form_submit_button("🔑 Ganti Password")
        if ok:
            if not np:
                st.error("Password kosong.")
            elif len(np) < 6:
                st.error("Minimal 6 karakter.")
            elif np != np2:
                st.error("Konfirmasi tidak cocok.")
            else:
                _change_password(sel, np)
                st.success(f"✅ Password **{sel}** berhasil diubah.")

    with t5:
        st.error("⚠️ Hapus permanen — tidak bisa dibatalkan.")
        konfirm = st.text_input(f"Ketik **{sel}** untuk konfirmasi:",
                                key="adm_konfirm")
        if st.button("🗑️ Hapus Permanen", key="adm_del"):
            if konfirm == sel:
                _delete_user(sel)
                st.success(f"✅ **{sel}** dihapus.")
                st.rerun()
            else:
                st.error("Konfirmasi tidak cocok. Batal.")

    st.markdown("---")

    # ── TAMBAH USER MANUAL ─────────────────────────────────
    st.subheader("➕ Tambah User Baru (Manual)")

    with st.form("form_add", clear_on_submit=True):
        c1, c2 = st.columns(2)
        un   = c1.text_input("Username", placeholder="min. 3 karakter")
        em   = c2.text_input("Email (opsional)")
        pw   = c1.text_input("Password", type="password",
                        placeholder="min. 6 karakter")
        role = c2.selectbox("Role", ["trial", "premium", "admin"])
        add  = st.form_submit_button("➕ Tambah User", use_container_width=True)

    if add:
        if not un or not pw:
            st.error("Username dan password wajib.")
        elif len(un.strip()) < 3:
            st.error("Username minimal 3 karakter.")
        elif len(pw) < 6:
            st.error("Password minimal 6 karakter.")
        else:
            res = register_user(un.strip(), pw, em, role)
            if res["ok"]:
                st.success(f"✅ User **{un}** ({role}) ditambahkan.")
                st.rerun()
            else:
                st.error(res["msg"])