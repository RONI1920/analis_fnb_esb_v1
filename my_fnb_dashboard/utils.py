# utils.py
import pandas as pd


def format_rupiah(amount):
    """Format angka menjadi string Rupiah DENGAN titik pemisah ribuan."""
    if pd.isna(amount):
        amount = 0
    return f"Rp {amount:,.0f}".replace(",", ".")


def format_angka_bulat(number):
    """Format angka kuantitas menjadi bulat tanpa desimal."""
    if pd.isna(number):
        number = 0
    return f"{number:.0f}"


def format_persen(number):
    """Format angka menjadi string persentase."""
    if pd.isna(number):
        number = 0
    return f"{number:,.1f}%"


def clean_payment_method(method_str):
    """Mengelompokkan metode pembayaran yang berantakan."""
    method_str = str(method_str).upper()
    if "VOUCHER" in method_str or "," in method_str:
        return "Voucher / Split"
    if "QRIS" in method_str:
        return "QRIS"
    if "VISA" in method_str:
        return "VISA"
    if "DEBIT CARD" in method_str:
        return "DEBIT CARD"
    if "BCA CARD" in method_str:
        return "BCA CARD"
    if "MASTER" in method_str:
        return "MASTER"
    if "CC " in method_str or "CREDIT CARD" in method_str:
        return "CREDIT CARD (Lainnya)"
    if "CASH" in method_str:
        return "CASH"
    if "TRANSFER" in method_str:
        return "TRANSFER"
    if "BRI CARD" in method_str:
        return "BRI CARD"
    if "BNI CARD" in method_str:
        return "BNI CARD"
    return "Lainnya"


def calculate_delta(value_A, value_B, formatter_func, higher_is_better=True):
    """
    Menghitung delta antara A dan B, mengembalikan string dan warna.
    """
    delta_abs = value_A - value_B

    # Tentukan fungsi format
    if formatter_func == format_rupiah:
        delta_abs_formatted = formatter_func(delta_abs)
    elif formatter_func == format_angka_bulat:
        delta_abs_formatted = formatter_func(delta_abs)
    else:  # Asumsi format_persen atau float
        delta_abs_formatted = f"{delta_abs:,.2f}"

    delta_pct_str = ""
    delta_color = "off"  # Default abu-abu

    if value_B != 0:
        delta_pct = (delta_abs / value_B) * 100
        arrow = "🔼" if delta_pct > 0 else "🔽"
        delta_pct_str = f"{arrow} {delta_pct:.1f}%"

        if delta_abs > 0:
            delta_color = "normal" if higher_is_better else "inverse"
        elif delta_abs < 0:
            delta_color = "inverse" if higher_is_better else "normal"
    elif value_A != 0:
        delta_pct_str = "🔼 100% +"
        delta_color = "normal" if higher_is_better else "inverse"
    else:
        delta_pct_str = "-"

    return delta_abs_formatted, delta_pct_str, delta_color
