import streamlit as st
import pandas as pd

from itertools import groupby
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
from st_aggrid.shared import ColumnsAutoSizeMode

from utils.exporters import create_download_button_excel, create_download_button_csv
from datetime import date, timedelta, datetime
import pytz

LOCAL_TIMEZONE = pytz.timezone('Asia/Jakarta')

from typing import Optional, List, Dict, Any, Tuple
from utils.validators import validate_quantity, validate_required_fields, validate_price
import re
import time


def normalize_for_comparison(text: str) -> str:
    """
    Enhanced normalization for strict duplicate detection.
    Handles special characters, spacing, and case.
    """
    if not text or not isinstance(text, str):
        return ""
    normalized = text.lower()
    normalized = ''.join(normalized.split())
    special_chars = ['(', ')', '-', '/', ',', '.', ':', ';', '_', '[', ']', '{', '}', '"', "'"]
    for char in special_chars:
        normalized = normalized.replace(char, '')
    normalized = ''.join(c for c in normalized if c.isalnum())
    return normalized.strip()


from config.settings import *
from models.inventory import (
    add_inventory_item, reduce_inventory_quantity,
    get_all_inventory, update_inventory_item, delete_inventory_item
)
from models.history import (
    log_transaction, get_history, get_summary_stats,
    delete_history_record, update_history_record
)
from models.reference import add_reference_item
from utils.helpers import (
    terbilang, format_date_for_display, standardize_display,
    is_pabrik_source, format_currency
)

try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


# ==================================================
# VALIDATION STYLING
# ==================================================

def inject_validation_css():
    st.markdown("""
    <style>
    .valid-field input, .valid-field select {
        border: 2px solid #10b981 !important;
        background-color: rgba(16, 185, 129, 0.05) !important;
    }
    .invalid-field input, .invalid-field select {
        border: 2px solid #ef4444 !important;
        background-color: rgba(239, 68, 68, 0.05) !important;
    }
    .field-error {
        color: #ef4444;
        font-size: 0.85rem;
        margin-top: -0.5rem;
        margin-bottom: 0.5rem;
        font-weight: 500;
    }
    .item-validation {
        padding: 0.75rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border-left: 4px solid;
    }
    .item-valid {
        background-color: rgba(16, 185, 129, 0.1);
        border-left-color: #10b981;
        color: #10b981;
    }
    .item-invalid {
        background-color: rgba(239, 68, 68, 0.1);
        border-left-color: #ef4444;
        color: #ef4444;
    }
    .item-partial {
        background-color: rgba(245, 158, 11, 0.1);
        border-left-color: #f59e0b;
        color: #f59e0b;
    }
    .optional-text {
        color: rgba(128, 128, 128, 0.6);
        font-weight: normal;
        font-size: 0.9em;
    }
    </style>
    """, unsafe_allow_html=True)


def inject_aggrid_merge_css():
    st.markdown("""
    <style>
    .ag-cell { position: relative; }
    div.ag-cell[data-value="↑"] {
        color: transparent !important;
        visibility: hidden !important;
        border: none !important;
    }
    .ag-theme-streamlit .ag-cell,
    .ag-theme-alpine .ag-cell,
    .ag-theme-balham .ag-cell { border-color: #4a5568; }
    .ag-row .ag-cell { border-right: 1px solid #4a5568; }
    </style>
    """, unsafe_allow_html=True)


def validate_field(value: Any, field_type: str = "text", required: bool = True) -> Tuple[bool, str]:
    if not required and not value:
        return True, ""
    if required and not value:
        return False, "Field ini wajib diisi"
    if field_type == "text" and isinstance(value, str):
        if value.strip() == "" or value.strip() == "-- Pilih --":
            return False, "Field ini wajib diisi"
    if field_type == "number":
        if value is None or value <= 0:
            return False, "Harus lebih besar dari 0"
    return True, ""


def show_field_validation(is_valid: bool, error_msg: str = ""):
    if not is_valid and error_msg:
        st.markdown(f'<div class="field-error">⚠️ {error_msg}</div>', unsafe_allow_html=True)


# ==================================================
# CACHED DATA LOADING
# ==================================================

@st.cache_data(ttl=300, show_spinner="Loading inventory...")
def get_cached_inventory():
    return get_all_inventory()


@st.cache_data(ttl=300, show_spinner="Loading history...")
def get_cached_history(start_date=None, end_date=None, item_name=None):
    return get_history(start_date=start_date, end_date=end_date, item_name=item_name)


def invalidate_caches():
    st.cache_data.clear()
    st.session_state.force_refresh_dashboard = True
    if 'dashboard_last_picked_cache' in st.session_state:
        del st.session_state['dashboard_last_picked_cache']


# ==================================================
# CONSTANTS
# ==================================================

DATE_FORMAT = "%d/%m/%y"
VALIDATION_MESSAGES = {
    'empty_field': "⚠️ Field ini wajib diisi",
    'duplicate': "⚠️ '{value}' sudah ada dalam daftar",
    'invalid_quantity': "⚠️ Jumlah harus lebih besar dari 0",
    'invalid_price': "⚠️ Harga harus lebih besar dari 0"
}


def render_instruction_box():
    with st.expander("📋 Panduan Pengisian Data Barang", expanded=False):
        st.markdown("#### 📝 Cara Pengisian Data Barang")
        st.write("Isi data barang sesuai format berikut:")
        st.markdown("---")
        st.markdown("##### 💡 Contoh Format:")
        st.success("**1 bh matras Kingbreeze 6k**")
        st.write("→ Kategori: `Matras` | Nama Barang: `Kingbreeze` | Ukuran: `6k` | Supplier: `Modis` | Qty: `1` | Warna: `Kosongin`")
        st.success("**2 bh b/g modis**")
        st.write("→ Kategori: `Bonus` | Nama Barang: `Bantal Guling` | Ukuran: `Kosongin` | Supplier: `Modis` | Qty: `2` | Warna: `Kosongin`")
        st.success("**2 bh mp modis 6k**")
        st.write("→ Kategori: `Bonus` | Nama Barang: `Matras Protector` | Ukuran: `6k` | Supplier: `Modis` | Qty: `2` | Warna: `Kosongin`")
        st.success("**1 tb Max Foam 3k (14)**")
        st.write("→ Kategori: `Tilam Busa` | Nama Barang: `Max Foam` | Ukuran: `3k (14)` | Supplier: `Modis` | Qty: `1` | Warna: `Kosongin`")
        st.success("**2 bh kursi makan 182 D Grey**")
        st.write("→ Kategori: `Kursi Makan` | Nama Barang: `182 D/nama lengkap` | Ukuran: `Kosongin/ ketik manual` | Supplier: `Deca Jaya` | Qty: `2` | Warna: `Abu-abu`")
        st.markdown("---")
        st.warning("⚠️ **Ketentuan Penting:**")
        st.markdown("""
        - ✅ Harap **cari pilihan** sebelum pilih "Tambah Baru"
        - ✅ **Ukuran HARAP diisi jika ada**
        - ✅ **HARAP tulis nama panjang** (contoh b/g -> bantal guling)
        - ✅ **Warna dan Ukuran boleh kosong**, akan otomatis terisi "Tidak Ada"
        - 🏭 Jika lokasi dipilih **Pabrik**, stok TIDAK tercatat sebagai stok aktif
        - 📦 Jika barang dari pabrik lalu **disimpan di toko/gudang**, pilih lokasi yang sesuai
        - 🎁 Untuk Bantal Guling, Guling, Bantal, Matras Protector → pilih kategori **Bonus**
        """)


# ==================================================
# UTILITY FUNCTIONS
# ==================================================

def convert_to_local_time(dt):
    if dt is None:
        return None
    if isinstance(dt, date) and not isinstance(dt, datetime):
        dt = datetime.combine(dt, datetime.min.time())
    if isinstance(dt, str):
        try:
            dt = datetime.strptime(dt[:26], '%Y-%m-%d %H:%M:%S.%f')
        except Exception:
            try:
                dt = datetime.strptime(dt[:19], '%Y-%m-%d %H:%M:%S')
            except Exception:
                return None
    if dt.tzinfo is None:
        dt = pytz.UTC.localize(dt)
    return dt.astimezone(LOCAL_TIMEZONE)


def format_local_datetime(dt, format_str="%d/%m/%y %H:%M"):
    local_dt = convert_to_local_time(dt)
    if local_dt is None:
        return "—"
    return local_dt.strftime(format_str)


def expand_combo_item(item_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    item_name_lower = item_data.get('name', '').strip().lower()
    if item_name_lower in COMBO_ITEM:
        sub_items = COMBO_ITEM[item_name_lower]
        expanded_items = []
        for sub_item in sub_items:
            new_item = item_data.copy()
            new_item['name'] = sub_item
            if 'details' not in new_item:
                new_item['details'] = {}
            new_item['details']['is_combo'] = True
            new_item['details']['combo_parent'] = item_name_lower
            expanded_items.append(new_item)
        return expanded_items
    return [item_data]


def format_date_dd_mm_yy(date_obj: Any) -> str:
    try:
        if isinstance(date_obj, (date, datetime)):
            if isinstance(date_obj, datetime):
                local_dt = convert_to_local_time(date_obj)
                return local_dt.strftime("%d/%m/%y") if local_dt else "—"
            return date_obj.strftime("%d/%m/%y")
        if isinstance(date_obj, str):
            local_dt = convert_to_local_time(date_obj)
            return local_dt.strftime("%d/%m/%y") if local_dt else "—"
    except Exception as e:
        print(f"Date formatting error: {e}")
    return "—"


def safe_get(dictionary: Dict, key: str, default: Any = None) -> Any:
    try:
        return dictionary.get(key, default)
    except (AttributeError, TypeError):
        return default


def format_item_combined(name: str, size: str, color: str) -> str:
    name = name.title() if name else "—"
    size = size.strip() if size and size.lower() != "tidak ada" else ""
    color = color.strip() if color and color.lower() != "tidak ada" else ""
    if size and color:
        return f"{name} ({size}) - {color.title()}"
    elif size:
        return f"{name} ({size})"
    elif color:
        return f"{name} - {color.title()}"
    return name


def select_or_text(
    label: str,
    options: List[str],
    key_prefix: str,
    default: str = "",
    manual_label: str = "Lainnya (ketik manual)",
    container=None,
    session_values_key: Optional[str] = None,
    required: bool = False
) -> str:
    if container is None:
        container = st

    from utils.validators import find_similar_entries, comprehensive_field_validation, get_all_reference_data_for_validation

    base_options = []
    if session_values_key and session_values_key in st.session_state:
        session_data = st.session_state[session_values_key]
        if session_data and len(session_data) > 0:
            base_options = list(session_data)
    if not base_options and options:
        base_options = [standardize_display(o) for o in options if o]
    if not isinstance(base_options, list):
        base_options = []

    base_options = sorted(set(base_options))
    normalized_lookup = {normalize_for_comparison(opt): opt for opt in base_options}

    value_key = f"{key_prefix}_value"
    confirm_key = f"{key_prefix}_confirm_add"
    ver_key = f"{key_prefix}_ver"

    if value_key not in st.session_state:
        st.session_state[value_key] = default
    if confirm_key not in st.session_state:
        st.session_state[confirm_key] = False
    if ver_key not in st.session_state:
        st.session_state[ver_key] = 0

    input_key = f"{key_prefix}_input_v{st.session_state[ver_key]}"

    if required:
        container.markdown(f"**{label}** *", unsafe_allow_html=True)
    else:
        container.markdown(
            f"**{label}** <span class='optional-text'>(optional)</span>",
            unsafe_allow_html=True
        )

    user_input = container.text_input(
        f"{label}_input",
        value=st.session_state[value_key],
        key=input_key,
        placeholder="Ketik untuk mencari atau pilih dari saran...",
        label_visibility="collapsed"
    )

    if user_input != st.session_state[value_key]:
        st.session_state[value_key] = user_input
        st.session_state[confirm_key] = False

    def _pick(chosen: str):
        st.session_state[value_key] = chosen
        st.session_state[ver_key] += 1
        st.session_state[confirm_key] = False
        st.rerun()

    if user_input and user_input.strip():
        search_term = user_input.strip().lower()
        normalized_input = normalize_for_comparison(user_input)

        if normalized_input in normalized_lookup:
            existing_match = normalized_lookup[normalized_input]
            if user_input.strip() == existing_match:
                return existing_match
            container.info(f"ℹ️ '{user_input}' sama dengan '{existing_match}' yang sudah ada")
            btn_key_use = f"{key_prefix}_use_existing_{hash(existing_match) % 100000}_{st.session_state[ver_key]}"
            if container.button(f"✓ Gunakan '{existing_match}'", key=btn_key_use, use_container_width=True):
                _pick(existing_match)
            return existing_match

        similar_entries = find_similar_entries(
            user_input, base_options, levenshtein_threshold=3, check_phonetic=True
        )

        if similar_entries:
            critical_matches = [e for e in similar_entries if e[1] <= 2 or e[2]]
            if critical_matches:
                container.warning(f"⚠️ **'{user_input}' mirip dengan data yang sudah ada:**")
                for idx, (entry, distance, is_phonetic) in enumerate(critical_matches[:5]):
                    match_type = "🔊 Bunyi sama" if is_phonetic else f"📏 Jarak: {distance}"
                    col1, col2 = container.columns([3, 1])
                    col1.markdown(
                        f"<div style='background: rgba(255, 165, 0, 0.15); padding: 0.5rem; border-radius: 4px; margin: 0.25rem 0;'>"
                        f"<strong>'{entry}'</strong> → {match_type}</div>",
                        unsafe_allow_html=True
                    )
                    btn_key_quick = f"{key_prefix}_quick_{idx}_{hash(entry) % 100000}_{st.session_state[ver_key]}"
                    if col2.button("✓ Pilih", key=btn_key_quick, use_container_width=True):
                        _pick(entry)
                container.info(
                    f"💡 **Saran:** Klik '✓ Pilih' untuk menggunakan pilihan yang ada, "
                    f"atau lanjut ke bawah jika '{user_input}' memang berbeda."
                )

        matching_options = []
        for opt in base_options:
            opt_lower = opt.lower()
            if search_term in opt_lower or opt_lower in search_term:
                matching_options.append(opt)
                continue
            opt_words = opt_lower.split()
            for word in opt_words:
                if word.startswith(search_term) and len(search_term) >= 2:
                    matching_options.append(opt)
                    break
        matching_options = list(dict.fromkeys(matching_options))

        if matching_options:
            top_matches = matching_options[:10]
            container.markdown(
                "<div style='background: #1a1a1a; padding: 0.5rem; border-radius: 8px; margin-top: 0.5rem; margin-bottom: 0.5rem;'>",
                unsafe_allow_html=True
            )
            container.caption(f"💡 {len(matching_options)} pilihan cocok - klik untuk memilih:")
            cols_per_row = 2
            for i in range(0, len(top_matches), cols_per_row):
                cols = container.columns(cols_per_row)
                for j, col in enumerate(cols):
                    idx = i + j
                    if idx < len(top_matches):
                        suggestion = top_matches[idx]
                        btn_key_select = f"{key_prefix}_sel_{i}_{j}_{hash(suggestion) % 100000}_{st.session_state[ver_key]}"
                        if col.button(f"✓ {suggestion}", key=btn_key_select, use_container_width=True):
                            _pick(suggestion)
            container.markdown("</div>", unsafe_allow_html=True)

        if not matching_options or not any(opt.lower() == search_term for opt in matching_options):
            container.markdown("---")

            should_block_add = False
            block_reason = ""

            field_type_mapping = {
                'dynamic_kategori': 'kategori',
                'dynamic_barang': 'barang',
                'dynamic_ukuran': 'ukuran',
                'dynamic_warna': 'warna',
                'dynamic_supplier': 'supplier',
                'dynamic_via': 'via',
                'dynamic_toko': 'toko'
            }

            field_type = field_type_mapping.get(session_values_key)

            if field_type:
                all_reference_data = get_all_reference_data_for_validation()
                typed_std = standardize_display(user_input)
                item_num_match = re.search(r'_(\d+)$', key_prefix)
                item_num = item_num_match.group(1) if item_num_match else "1"

                should_block, error_msg = comprehensive_field_validation(
                    field_type=field_type,
                    user_input=typed_std,
                    item_num=item_num,
                    all_reference_data=all_reference_data
                )
                if should_block:
                    should_block_add = True
                    block_reason = error_msg

                if field_type == 'barang' and not should_block_add:
                    all_kategori = all_reference_data.get('kategori', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for kat in all_kategori:
                        kat_normalized = normalize_for_comparison(kat)
                        if typed_normalized == kat_normalized or kat_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Kategori**, bukan Nama Barang!\n\n"
                                f"💡 **Anda mencoba memasukkan '{kat}' (Kategori) ke field Nama Barang**\n\n"
                                f"✅ **Yang benar:**\n• Kategori: `{kat}` → Nama Barang: `Emily`, `Kingbreeze`, dsb\n\n"
                                f"❌ **Salah:**\n• Nama Barang: `{typed_std}` ❌"
                            )
                            break

                if field_type in ['ukuran', 'warna', 'via', 'toko'] and not should_block_add:
                    all_barang = all_reference_data.get('barang', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for barang in all_barang:
                        barang_normalized = normalize_for_comparison(barang)
                        if typed_normalized == barang_normalized or barang_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Barang**, bukan {field_type.title()}!\n\n"
                                f"💡 **Anda mencoba memasukkan '{barang}' (Nama Barang) ke field {field_type.title()}**"
                            )
                            break

                if field_type in ['ukuran', 'warna', 'supplier', 'via', 'toko'] and not should_block_add:
                    all_kategori = all_reference_data.get('kategori', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for kat in all_kategori:
                        kat_normalized = normalize_for_comparison(kat)
                        if typed_normalized == kat_normalized or kat_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Kategori**, bukan {field_type.title()}!"
                            )
                            break

                if field_type == 'ukuran' and not should_block_add:
                    all_warna = all_reference_data.get('warna', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for warna in all_warna:
                        warna_normalized = normalize_for_comparison(warna)
                        if typed_normalized == warna_normalized or warna_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Warna**, bukan Ukuran!\n\n"
                                f"✅ **Yang benar:**\n• Warna: `{warna}` | Ukuran: `3k`, `6k`, `180x200`, dll"
                            )
                            break

                if field_type == 'warna' and not should_block_add:
                    all_ukuran = all_reference_data.get('ukuran', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for ukuran in all_ukuran:
                        ukuran_normalized = normalize_for_comparison(ukuran)
                        if typed_normalized == ukuran_normalized or ukuran_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Ukuran**, bukan Warna!\n\n"
                                f"✅ **Yang benar:**\n• Ukuran: `{ukuran}` | Warna: `Merah`, `Abu-abu`, dll"
                            )
                            break

                if field_type in ['via', 'toko'] and not should_block_add:
                    all_supplier = all_reference_data.get('supplier', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for supplier in all_supplier:
                        supplier_normalized = normalize_for_comparison(supplier)
                        if typed_normalized == supplier_normalized or supplier_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Supplier**, bukan {field_type.title()}!"
                            )
                            break

                if field_type == 'toko' and not should_block_add:
                    all_via = all_reference_data.get('via', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for via in all_via:
                        via_normalized = normalize_for_comparison(via)
                        if typed_normalized == via_normalized or via_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Via Pengiriman**, bukan Kirim Ke Toko!"
                            )
                            break

                if field_type == 'via' and not should_block_add:
                    all_toko = all_reference_data.get('toko', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    for toko in all_toko:
                        toko_normalized = normalize_for_comparison(toko)
                        if typed_normalized == toko_normalized or toko_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Toko**, bukan Via Pengiriman!"
                            )
                            break

            if should_block_add:
                container.error(f"🚫 {block_reason}")
                container.warning(
                    "❌ **DATA DITOLAK!**\n\n"
                    "💡 **Pastikan Anda memasukkan data di field yang benar.**"
                )
                st.session_state[confirm_key] = False
            else:
                if not st.session_state[confirm_key]:
                    container.info(f"💡 Tidak menemukan '{user_input}' dalam daftar")
                    btn_key_add = f"{key_prefix}_add_new_{hash(user_input) % 100000}_{st.session_state[ver_key]}"
                    if container.button(
                        f"➕ Tambahkan '{user_input}' sebagai pilihan baru",
                        key=btn_key_add,
                        use_container_width=True,
                        type="primary"
                    ):
                        st.session_state[confirm_key] = True
                        st.rerun()
                else:
                    typed_std = standardize_display(user_input)
                    typed_normalized = normalize_for_comparison(typed_std)
                    if typed_normalized in normalized_lookup:
                        existing = normalized_lookup[typed_normalized]
                        container.warning(f"⚠️ '{typed_std}' sama dengan '{existing}'")
                        _pick(existing)
                    else:
                        if session_values_key:
                            table_mapping = {
                                'dynamic_supplier': 'suppliers',
                                'dynamic_kategori': 'kategori',
                                'dynamic_barang': 'barang',
                                'dynamic_warna': 'warna',
                                'dynamic_ukuran': 'ukuran',
                                'dynamic_toko': 'toko',
                                'dynamic_via': 'via_pengiriman'
                            }
                            table_name = table_mapping.get(session_values_key)
                            if table_name:
                                success, msg = add_reference_item(table_name, typed_std)
                                if success:
                                    if session_values_key not in st.session_state:
                                        st.session_state[session_values_key] = []
                                    if typed_std not in st.session_state[session_values_key]:
                                        st.session_state[session_values_key].append(typed_std)
                                        st.session_state[session_values_key] = sorted(
                                            st.session_state[session_values_key]
                                        )
                                    st.session_state.reference_data_updated = True
                                    container.success(f"✅ '{typed_std}' berhasil ditambahkan!")
                                    _pick(typed_std)
                                else:
                                    container.error(f"❌ Gagal menambahkan: {msg}")
                                    st.session_state[confirm_key] = False

    current_value = st.session_state[value_key]
    return current_value if current_value and current_value.strip() else "-"


def create_dataframe_from_records(records: List[Dict[str, Any]], columns_map: Dict[str, str]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()
    data = []
    for rec in records:
        row = {}
        for key, col_name in columns_map.items():
            value = safe_get(rec, key, '—')
            if 'date' in key.lower() and isinstance(value, (date, datetime, str)):
                value = format_date_dd_mm_yy(value)
            elif 'price' in key.lower() or 'harga' in col_name.lower():
                value = format_currency(value) if isinstance(value, (int, float)) else value
            elif isinstance(value, str):
                value = value.title() if key in ['name', 'item_name'] else value
            row[col_name] = value
        data.append(row)
    return pd.DataFrame(data)


# ==================================================
# DASHBOARD TAB
# ==================================================

def render_dashboard_tab(today: date):
    cache_key = 'dashboard_last_picked_cache'
    force_refresh_key = 'force_refresh_dashboard'

    if force_refresh_key not in st.session_state:
        st.session_state[force_refresh_key] = False

    rebuild_cache = (
        cache_key not in st.session_state or
        st.session_state.get(force_refresh_key, False) or
        not isinstance(st.session_state.get(cache_key), dict)
    )

    if rebuild_cache:
        try:
            all_data = get_history()
            category_times = {
                'MASUK': "Belum Ada Data",
                'LUAR_KOTA': "Belum Ada Data",
                'ECERAN': "Belum Ada Data"
            }
            latest_overall = None
            latest_time = datetime.min.replace(tzinfo=pytz.UTC)

            if all_data and len(all_data) > 0:
                for action_type in ['MASUK', 'LUAR_KOTA', 'ECERAN']:
                    category_records = [r for r in all_data if r.get('action_type') == action_type]
                    if category_records:
                        latest_in_category = max(
                            category_records,
                            key=lambda r: convert_to_local_time(r.get('created_at')) or datetime.min.replace(tzinfo=pytz.UTC)
                        )
                        created_at = latest_in_category.get('created_at')
                        time_str = format_local_datetime(created_at, "%d/%m/%y %H:%M")
                        if time_str != "—":
                            category_times[action_type] = time_str
                            local_dt = convert_to_local_time(created_at)
                            if local_dt and local_dt > latest_time:
                                latest_time = local_dt
                                latest_overall = latest_in_category

                if latest_overall:
                    action_type = latest_overall.get('action_type', '')
                    category_map = {'MASUK': 'Barang Masuk', 'LUAR_KOTA': 'Luar Kota', 'ECERAN': 'Eceran'}
                    last_category = category_map.get(action_type, action_type)
                    trans_date = latest_overall.get('transaction_date')
                    last_trans_date = format_local_datetime(trans_date, "%d/%m/%Y")
                    last_picked = format_item_combined(
                        latest_overall.get('item_name', ''),
                        latest_overall.get('size', ''),
                        latest_overall.get('color', '')
                    )
                    last_picked = f"{last_picked} | {last_trans_date}"
                else:
                    last_category = "—"
                    last_picked = "Belum Ada Data"

                st.session_state[cache_key] = {
                    'masuk_time': category_times['MASUK'],
                    'luar_time': category_times['LUAR_KOTA'],
                    'eceran_time': category_times['ECERAN'],
                    'last_category': last_category,
                    'last_picked': last_picked
                }
            else:
                st.session_state[cache_key] = {
                    'masuk_time': "Belum Ada Data",
                    'luar_time': "Belum Ada Data",
                    'eceran_time': "Belum Ada Data",
                    'last_category': "—",
                    'last_picked': "Belum Ada Data"
                }
            st.session_state[force_refresh_key] = False
        except Exception as e:
            st.error(f"❌ Error loading dashboard: {e}")
            import traceback
            st.code(traceback.format_exc())

    cache_data = st.session_state.get(cache_key, {
        'masuk_time': "Belum Ada Data",
        'luar_time': "Belum Ada Data",
        'eceran_time': "Belum Ada Data",
        'last_category': "—",
        'last_picked': "Belum Ada Data"
    })

    if not isinstance(cache_data, dict):
        cache_data = {
            'masuk_time': "Belum Ada Data",
            'luar_time': "Belum Ada Data",
            'eceran_time': "Belum Ada Data",
            'last_category': "—",
            'last_picked': "Belum Ada Data"
        }

    masuk_time = cache_data.get('masuk_time', "Belum Ada Data")
    luar_time = cache_data.get('luar_time', "Belum Ada Data")
    eceran_time = cache_data.get('eceran_time', "Belum Ada Data")
    last_category = cache_data.get('last_category', "—")
    last_picked = cache_data.get('last_picked', "Belum Ada Data")

    current_local_time = datetime.now(LOCAL_TIMEZONE).strftime("%d/%m/%y %H:%M:%S")
    st.caption(f"🕐 Waktu Sekarang: {current_local_time}")

    col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
    with col1:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px; font-size:0.75rem;'>TERAKHIR INPUT</small>"
            f"<h3 style='margin:8px 0; font-size:1rem; color:#3B82F6;'>📥 Barang Masuk</h3>"
            f"<p style='margin:5px 0; font-size:0.95rem; font-weight:600;'>{masuk_time}</p>"
            f"</div>", unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px; font-size:0.75rem;'>TERAKHIR INPUT</small>"
            f"<h3 style='margin:8px 0; font-size:1rem; color:#F59E0B;'>🚚 Luar Kota</h3>"
            f"<p style='margin:5px 0; font-size:0.95rem; font-weight:600;'>{luar_time}</p>"
            f"</div>", unsafe_allow_html=True
        )
    with col3:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px; font-size:0.75rem;'>TERAKHIR INPUT</small>"
            f"<h3 style='margin:8px 0; font-size:1rem; color:#10B981;'>🛒 Eceran</h3>"
            f"<p style='margin:5px 0; font-size:0.95rem; font-weight:600;'>{eceran_time}</p>"
            f"</div>", unsafe_allow_html=True
        )
    with col4:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px;'>BARANG TERAKHIR DIUPDATE</small>"
            f"<h2 style='margin:10px 0 0 0; font-size:1.3rem;'>{last_picked}</h2>"
            f"<p style='color:#888; margin:5px 0 0 0; font-size:0.9rem;'>Kategori: {last_category}</p>"
            f"</div>", unsafe_allow_html=True
        )

    st.markdown("---")

    try:
        all_inventory = get_cached_inventory()
        inventory = [item for item in all_inventory if not is_pabrik_source(item.get('location', ''))]

        if not inventory:
            st.info("ℹ️ Tidak ada stok aktif.")
            return

        low_stock = [item for item in inventory if 0 < item.get('quantity', 0) < 3]
        critical_stock = [item for item in inventory if item.get('quantity', 0) == 0]
        medium_stock = [item for item in inventory if 4 <= item.get('quantity', 0) <= 10]
        high_stock = [item for item in inventory if item.get('quantity', 0) > 10]

        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("⚠️ Stock Rendah", len(low_stock))
        col_b.metric("🚨 Kehabisan Stock", len(critical_stock))
        col_c.metric("📦 Stock Cukup", len(medium_stock))
        col_d.metric("✅ Stock Banyak", len(high_stock))

        if critical_stock or low_stock:
            with st.expander("⚠️ Peringatan Stok", expanded=True):
                for item in (critical_stock + low_stock)[:10]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    badge_color = "#dc2626" if qty == 0 else "#d97706"
                    badge_text = "HABIS" if qty == 0 else "RENDAH"
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'),
                        item.get('size', ''),
                        item.get('color', '')
                    )
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid {badge_color};'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:{badge_color}; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>{badge_text}</div>"
                        f"</div></div>", unsafe_allow_html=True
                    )

        if medium_stock:
            with st.expander("📦 Barang Stock Cukup (4-10)", expanded=False):
                for item in medium_stock[:15]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'), item.get('size', ''), item.get('color', '')
                    )
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid #2563eb;'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:#2563eb; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>CUKUP</div>"
                        f"</div></div>", unsafe_allow_html=True
                    )

        if high_stock:
            with st.expander("✅ Barang Stock Banyak (>10)", expanded=False):
                for item in high_stock[:15]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'), item.get('size', ''), item.get('color', '')
                    )
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid #059669;'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:#059669; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>BANYAK</div>"
                        f"</div></div>", unsafe_allow_html=True
                    )

        if not critical_stock and not low_stock:
            st.success("✅ Semua stok dalam kondisi baik")

    except Exception as e:
        st.error(f"❌ Error loading dashboard: {str(e)}")


# ==================================================
# BARANG MASUK TAB
# ==================================================

def render_masuk_barang_tab(today: date):
    inject_validation_css()
    render_instruction_box()

    tanggal_masuk = st.date_input("Tanggal Masuk", value=today, max_value=today, key="masuk_tgl", format="DD/MM/YYYY")

    col1, col2 = st.columns([1, 2])
    nomor_po = col1.text_input("Nomor PO *", key="masuk_po", placeholder="Wajib diisi").strip()
    supplier = select_or_text("Supplier", SUPPLIERS, key_prefix="masuk_sup", session_values_key="dynamic_supplier", container=col2, required=True)

    st.markdown("---")
    st.markdown("### 📍 Lokasi Penyimpanan")
    same_location = st.radio(
        "Apakah semua barang ditaruh di tempat yang sama?",
        options=["Iya", "Tidak"], key="masuk_same_loc", horizontal=True
    )

    shared_location = None
    if same_location == "Iya":
        shared_location = st.selectbox("Lokasi untuk Semua Barang *", STANDARD_LOCATIONS, key="masuk_shared_loc")
        if shared_location == "Pabrik":
            st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak masuk Stok Aktif")

    st.markdown("---")
    st.markdown("### 📦 Daftar Barang")

    if 'masuk_item_count' not in st.session_state:
        st.session_state.masuk_item_count = 1

    items_to_add = []
    item_validations = []

    for i in range(1, st.session_state.masuk_item_count + 1):
        with st.expander(f"Item {i}", expanded=(i == st.session_state.masuk_item_count)):
            col1, col2 = st.columns(2)
            cat = select_or_text("Kategori", STANDARD_CATEGORIES, f"add_cat_{i}", container=col1, session_values_key="dynamic_kategori", required=True).strip().lower()
            name = select_or_text("Nama Barang", STANDARD_BARANG, f"add_name_{i}", container=col2, session_values_key="dynamic_barang", required=True).strip().lower()

            col3, col4, col5 = st.columns([2, 2, 1.5])
            size = select_or_text("Ukuran", STANDARD_SIZES, f"add_sz_{i}", container=col3, session_values_key="dynamic_ukuran", required=True).strip().lower()
            color = select_or_text("Warna", STANDARD_COLORS, f"add_col_{i}", container=col4, session_values_key="dynamic_warna").strip().lower()
            qty = col5.number_input("Jumlah *", min_value=0, step=1, key=f"add_qty_{i}")

            if same_location == "Tidak":
                loc = st.selectbox("Lokasi *", STANDARD_LOCATIONS, key=f"add_loc_{i}")
                if loc == "Pabrik":
                    st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak masuk Stok Aktif")
            else:
                loc = shared_location

            komen = st.text_input("Komen", placeholder="Optional", key=f"add_komen_{i}").strip()

            missing_fields = []
            validation_errors = []

            if not cat or cat == "-":
                missing_fields.append("Kategori")
            if not name or name == "-":
                missing_fields.append("Nama Barang")
            if not size or size == "-":
                missing_fields.append("Ukuran")
            if qty <= 0:
                missing_fields.append("Jumlah (harus > 0)")

            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">❌ {overlap_msg}</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, [overlap_msg]))
                    continue

            if cat or name or size or qty > 0:
                if missing_fields or validation_errors:
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    error_display.extend(validation_errors)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Item {i}: <strong>{" | ".join(error_display)}</strong></div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">✅ Item {i}: Lengkap dan siap disimpan</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, True, []))
                    items_to_add.append({
                        "name": name, "category": cat, "color": color or "Tidak Ada",
                        "size": size, "quantity": qty, "location": loc, "supplier": supplier,
                        "nomor_po": nomor_po or None,
                        "details": {"komen": komen if komen else None}
                    })

    col_add, col_delete = st.columns([3, 1])
    with col_add:
        if st.session_state.masuk_item_count < MAX_ITEMS_PER_FORM:
            if st.button("➕ Tambah Item Baru", key="add_masuk_item_btn", use_container_width=True):
                st.session_state.masuk_item_count += 1
                st.rerun()
        else:
            st.info(f"ℹ️ Maksimal {MAX_ITEMS_PER_FORM} item per form")
    with col_delete:
        if st.session_state.masuk_item_count > 1:
            if st.button("🗑️ Hapus Item Terakhir", key="delete_masuk_item_btn", use_container_width=True, type="secondary"):
                st.session_state.masuk_item_count -= 1
                st.rerun()

    st.markdown("---")

    if items_to_add:
        expanded_preview = []
        for item in items_to_add:
            expanded_preview.extend(expand_combo_item(item))
        st.success(f"✅ {len(expanded_preview)} item siap disimpan")
        col1, col2 = st.columns(2)
        col1.metric("Total Items", len(expanded_preview))
        col2.metric("Total Quantity", sum(item['quantity'] for item in expanded_preview))

    if st.button("💾 Simpan Barang Masuk", type="primary", key="save_masuk_btn"):
        validation_errors = []
        if not nomor_po:
            validation_errors.append("❌ **Nomor PO** wajib diisi!")
        if not supplier:
            validation_errors.append("❌ **Supplier** wajib diisi!")
        if not items_to_add:
            validation_errors.append("❌ **Belum ada item yang lengkap**.")
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, errors_list in incomplete_items:
                if errors_list:
                    validation_errors.append(f"❌ **Item {item_num}**: {', '.join(str(e) for e in errors_list)}")
        if validation_errors:
            st.error("### ⚠️ Validasi Gagal:")
            for error in validation_errors:
                st.markdown(error)
            return

        success_count = 0
        all_items = []
        for item in items_to_add:
            all_items.extend(expand_combo_item(item))

        for item in all_items:
            try:
                if not is_pabrik_source(item['location']):
                    add_inventory_item(
                        name=item['name'], category=item['category'], color=item['color'],
                        size=item['size'], quantity=item['quantity'], location=item['location'],
                        supplier=item['supplier'], nomor_po=item['nomor_po'], date_added=tanggal_masuk
                    )
                transaction_details = item.get('details', {}).copy()
                if item.get('nomor_po'):
                    transaction_details['nomor_po'] = item['nomor_po']
                log_transaction(
                    action_type="MASUK", item_name=item['name'], quantity=item['quantity'],
                    category=item['category'], color=item['color'], size=item['size'],
                    location=item['location'], supplier=item['supplier'],
                    details=transaction_details, transaction_date=tanggal_masuk
                )
                success_count += 1
            except Exception as e:
                st.error(f"❌ Error menyimpan {item['name']}: {str(e)}")

        if success_count > 0:
            st.success(f"✅ Berhasil menyimpan {success_count} item!")
            st.balloons()
            invalidate_caches()
            st.session_state.force_refresh_dashboard = True
            st.session_state.masuk_item_count = 1
            keys_to_clear = [k for k in st.session_state.keys() if k.startswith('masuk_') or k.startswith('add_')]
            for key in keys_to_clear:
                del st.session_state[key]
            time.sleep(1)
            st.rerun()


# ==================================================
# LUAR KOTA TAB
# ==================================================

def render_luar_kota_tab(today: date):
    inject_validation_css()
    render_instruction_box()

    tanggal_kirim = st.date_input("Tanggal Kirim", value=today, max_value=today, key="luar_tgl", format="DD/MM/YYYY")

    col1, col2 = st.columns(2)
    purpose = select_or_text("Kirim Ke (Toko)", DAFTAR_TOKO, "luar_tuj", container=col1, session_values_key="dynamic_toko", required=True)
    via = select_or_text("Via Pengiriman", VIA_OPTIONS, "luar_via", container=col2, session_values_key="dynamic_via", required=True)

    st.markdown("---")
    st.markdown("### 📍 Sumber Barang")
    same_source = st.radio(
        "Apakah semua barang diambil dari tempat yang sama?",
        options=["Iya", "Tidak"], key="luar_same_src", horizontal=True
    )

    shared_source = None
    if same_source == "Iya":
        shared_source = st.selectbox("Sumber untuk Semua Barang *", STANDARD_LOCATIONS, key="luar_shared_src")
        if shared_source == "Pabrik":
            st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak kurangi Stok Aktif")

    st.markdown("---")
    st.markdown("### 📦 Barang yang Dikirim")

    luar_items = []
    item_validations = []

    if 'luar_item_count' not in st.session_state:
        st.session_state.luar_item_count = 1

    for i in range(1, st.session_state.luar_item_count + 1):
        with st.expander(f"Item {i}", expanded=(i == st.session_state.luar_item_count)):
            col1, col2 = st.columns(2)
            cat = select_or_text("Kategori", STANDARD_CATEGORIES, f"luar_cat_{i}", container=col1, session_values_key="dynamic_kategori", required=True).strip().lower()
            name = select_or_text("Nama Barang", STANDARD_BARANG, f"luar_name_{i}", container=col2, session_values_key="dynamic_barang", required=True).strip().lower()

            col3, col4, col5 = st.columns([2, 2, 1])
            sz = select_or_text("Ukuran", STANDARD_SIZES, f"luar_sz_{i}", container=col3, session_values_key="dynamic_ukuran", required=True).strip().lower()
            colr = select_or_text("Warna", STANDARD_COLORS, f"luar_col_{i}", container=col4, session_values_key="dynamic_warna").strip().lower()
            qty = col5.number_input("Qty *", min_value=0, key=f"luar_qty_{i}")

            supplier = select_or_text("Supplier", SUPPLIERS, f"luar_sup_{i}", session_values_key="dynamic_supplier", required=True).strip().lower()

            if same_source == "Tidak":
                loc = st.selectbox("Dari *", STANDARD_LOCATIONS, key=f"luar_loc_{i}")
                if loc == "Pabrik":
                    st.info("ℹ️ Dari Pabrik: tidak kurangi Stok Aktif")
            else:
                loc = shared_source

            komen = st.text_input("Komen", placeholder="Optional", key=f"luar_komen_{i}").strip()

            missing_fields = []
            validation_errors = []

            if not cat or cat == "-":
                missing_fields.append("Kategori")
            if not name or name == "-":
                missing_fields.append("Nama Barang")
            if not sz or sz == "-":
                missing_fields.append("Ukuran")
            if not supplier or supplier == "-":
                missing_fields.append("Supplier")
            if qty <= 0:
                missing_fields.append("Qty (harus > 0)")

            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">❌ {overlap_msg}</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, [overlap_msg]))
                    continue

            if cat or name or sz or supplier or qty > 0:
                if missing_fields or validation_errors:
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    error_display.extend(validation_errors)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Item {i}: <strong>{" | ".join(error_display)}</strong></div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">✅ Item {i}: Lengkap dan siap dikirim</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, True, []))
                    luar_items.append({
                        "name": name, "category": cat, "color": colr or "Tidak Ada",
                        "size": sz, "quantity": qty, "source": loc, "location": loc,
                        "supplier": supplier,
                        "details": {"komen": komen if komen else None, "purpose": purpose, "via": via}
                    })

    col_add_luar, col_delete_luar = st.columns([3, 1])
    with col_add_luar:
        if st.session_state.luar_item_count < MAX_ITEMS_PER_FORM:
            if st.button("➕ Tambah Item Baru", key="add_luar_item_btn", use_container_width=True):
                st.session_state.luar_item_count += 1
                st.rerun()
        else:
            st.info(f"ℹ️ Maksimal {MAX_ITEMS_PER_FORM} item per form")
    with col_delete_luar:
        if st.session_state.luar_item_count > 1:
            if st.button("🗑️ Hapus Item Terakhir", key="delete_luar_item_btn", use_container_width=True, type="secondary"):
                st.session_state.luar_item_count -= 1
                st.rerun()

    st.markdown("---")

    if luar_items:
        expanded_preview = []
        for item in luar_items:
            expanded_preview.extend(expand_combo_item(item))
        st.success(f"✅ {len(expanded_preview)} item siap dikirim")
        col1, col2 = st.columns(2)
        col1.metric("📦 Total Items", len(expanded_preview))
        col2.metric("📊 Total Qty", sum(item['quantity'] for item in expanded_preview))

    if st.button("🚚 Simpan Pengiriman", type="primary", use_container_width=True):
        validation_errors = []
        if not purpose:
            validation_errors.append("❌ **Kirim Ke (Toko)** wajib diisi!")
        if not via:
            validation_errors.append("❌ **Via Pengiriman** wajib diisi!")
        if not luar_items:
            validation_errors.append("❌ **Belum ada item yang lengkap**.")
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, missing in incomplete_items:
                validation_errors.append(f"❌ **Item {item_num}**: {', '.join(missing)}")
        if validation_errors:
            st.error("### ⚠️ Validasi Gagal:")
            for error in validation_errors:
                st.markdown(error)
            return

        success_count = 0
        all_items = []
        for item in luar_items:
            all_items.extend(expand_combo_item(item))

        for item in all_items:
            try:
                if not is_pabrik_source(item['location']):
                    try:
                        reduce_inventory_quantity(
                            name=item['name'], category=item['category'], color=item['color'],
                            size=item['size'], location=item['location'], quantity=item['quantity']
                        )
                    except Exception:
                        pass
                log_transaction(
                    action_type="LUAR_KOTA", item_name=item['name'], quantity=-item['quantity'],
                    category=item['category'], color=item['color'], size=item['size'],
                    location=item['location'], supplier=item['supplier'],
                    details=item['details'], transaction_date=tanggal_kirim
                )
                success_count += 1
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

        if success_count > 0:
            st.success(f"✅ Berhasil menyimpan {success_count} item!")
            st.balloons()
            invalidate_caches()
            st.session_state.force_refresh_dashboard = True
            keys_to_clear = [k for k in st.session_state.keys() if k.startswith('luar_')]
            for key in keys_to_clear:
                del st.session_state[key]
            time.sleep(1)
            st.rerun()


# ==================================================
# ECERAN TAB
# ==================================================

def render_eceran_tab(today: date):
    inject_validation_css()
    render_instruction_box()

    ecer_items = []
    item_validations = []

    tanggal_jual = st.date_input("Tanggal Jual", today, max_value=today, key="ecer_tgl_date", format="DD/MM/YYYY")

    st.markdown("#### 👤 Informasi Pelanggan")
    col1, col2 = st.columns([3, 1.5])
    cust = col1.text_input("Nama Pelanggan *", key="ecer_cust", placeholder="Nama lengkap").strip()
    phone = col2.text_input("Telepon", key="ecer_phone", placeholder="08xxx").strip()
    alamat = st.text_area("Alamat *", key="ecer_alamat", placeholder="Alamat lengkap").strip()

    st.markdown("#### 💳 Detail Transaksi")
    cols = st.columns(4)

    spg_options = ["-- Pilih SPG --"] + SPG_LIST
    spg_selected = cols[0].selectbox("SPG *", spg_options, key="ecer_spg")
    if spg_selected == "Lainnya (ketik manual)":
        spg = cols[0].text_input("SPG (manual)", key="ecer_spg_man").strip() or "Manual"
    elif spg_selected == "-- Pilih SPG --":
        spg = None
    else:
        spg = spg_selected

    pay_type_options = ["-- Pilih Pembayaran --"] + PAYMENT_TYPES
    pay_type_selected = cols[1].selectbox("Pembayaran *", pay_type_options, key="ecer_paytype")
    if pay_type_selected == "Lainnya (ketik manual)":
        pay_type = cols[1].text_input("Pembayaran (manual)", key="ecer_paytype_man").strip() or "Manual"
    elif pay_type_selected == "-- Pilih Pembayaran --":
        pay_type = None
    else:
        pay_type = pay_type_selected

    rekening = None
    if pay_type == "Transfer" or pay_type_selected == "Transfer":
        rek_options = ["-- Pilih Rekening --"] + REKENING_OPTIONS
        rek_opt = cols[2].selectbox("Rekening *", rek_options, key="ecer_rek")
        if rek_opt == "Lainnya (ketik manual)":
            rekening = cols[2].text_input("Rekening manual", key="ecer_rek_man").strip() or "Manual"
        elif rek_opt == "-- Pilih Rekening --":
            rekening = None
        else:
            rekening = rek_opt

    pay_stat_options = ["-- Pilih Status --"] + PAYMENT_STATUS
    pay_stat_selected = cols[3].selectbox("Status Bayar *", pay_stat_options, key="ecer_stat")
    if pay_stat_selected == "Lainnya (ketik manual)":
        pay_stat = cols[3].text_input("Status Bayar (manual)", key="ecer_stat_man").strip() or "Manual"
    elif pay_stat_selected == "-- Pilih Status --":
        pay_stat = None
    else:
        pay_stat = pay_stat_selected

    kirim_options = ["-- Pilih Status --"] + SHIPPING_STATUS
    kirim_selected = st.selectbox("Status Pengiriman *", kirim_options, key="ecer_kirim")
    if kirim_selected == "Lainnya (ketik manual)":
        kirim = st.text_input("Status Pengiriman (manual)", key="ecer_kirim_man").strip() or "Manual"
    elif kirim_selected == "-- Pilih Status --":
        kirim = None
    else:
        kirim = kirim_selected

    kirim_tgl = st.date_input("Tgl Kirim", today, key="ecer_kirim_tgl_date", format="DD/MM/YYYY") if kirim == "Konfirmasi" else ""

    st.markdown("---")
    st.markdown("### 📍 Sumber Barang")
    same_source = st.radio(
        "Apakah semua barang diambil dari tempat yang sama?",
        options=["Iya", "Tidak"], key="ecer_same_src", horizontal=True
    )

    shared_source = None
    if same_source == "Iya":
        shared_source_options = ["-- Pilih Sumber --"] + SALES_SOURCE_OPTIONS
        shared_source_selected = st.selectbox("Sumber untuk Semua Barang *", shared_source_options, key="ecer_shared_src")
        if shared_source_selected == "Lainnya (ketik manual)":
            shared_source = st.text_input("Sumber (manual)", key="ecer_shared_src_man").strip() or "Manual"
        elif shared_source_selected == "-- Pilih Sumber --":
            shared_source = None
        else:
            shared_source = shared_source_selected
        if shared_source == "Pabrik":
            st.info("ℹ️ Dari Pabrik: tidak kurangi Stok Aktif")

    st.markdown("---")
    st.markdown("#### 🛍️ Barang Dijual")

    if 'ecer_item_count' not in st.session_state:
        st.session_state.ecer_item_count = 1

    for i in range(1, st.session_state.ecer_item_count + 1):
        with st.expander(f"Barang {i}", expanded=(i == st.session_state.ecer_item_count)):
            c1, c2 = st.columns(2)
            cat = select_or_text("Kategori", STANDARD_CATEGORIES, f"ec_cat_{i}", container=c1, session_values_key="dynamic_kategori", required=True).strip().lower()
            name = select_or_text("Nama Barang", STANDARD_BARANG, f"ec_name_{i}", container=c2, session_values_key="dynamic_barang", required=True).strip().lower()

            c3, c4, c5, c6 = st.columns([2, 2, 1, 1.5])
            sz = select_or_text("Ukuran", STANDARD_SIZES, f"ec_sz_{i}", container=c3, session_values_key="dynamic_ukuran", required=True).strip().lower()
            colr = select_or_text("Warna", STANDARD_COLORS, f"ec_col_{i}", container=c4, session_values_key="dynamic_warna").strip().lower()
            qty = c5.number_input("Qty", min_value=0, key=f"ec_qty_{i}")
            harga = c6.number_input("Harga Total (Rp)", min_value=0, step=1000, key=f"ec_hrg_{i}")

            c7, c8 = st.columns(2)
            sup = select_or_text("Supplier", SUPPLIERS, f"ec_sup_{i}", container=c7, session_values_key="dynamic_supplier", required=True).strip().lower()

            if same_source == "Tidak":
                src_options = ["-- Pilih Sumber --"] + SALES_SOURCE_OPTIONS
                src_selected = c8.selectbox("Sumber *", src_options, key=f"ec_src_{i}")
                if src_selected == "Lainnya (ketik manual)":
                    src = c8.text_input("Sumber (manual)", key=f"ec_src_{i}_man").strip() or "Manual"
                elif src_selected == "-- Pilih Sumber --":
                    src = None
                else:
                    src = src_selected
            else:
                src = shared_source

            is_bonus = (cat == "bonus")

            missing_fields = []
            validation_errors = []

            if not cat:
                missing_fields.append("Kategori")
            if not name:
                missing_fields.append("Nama Barang")
            if not sz:
                missing_fields.append("Ukuran")
            if not sup:
                missing_fields.append("Supplier")
            if qty <= 0:
                missing_fields.append("Qty (harus > 0)")
            if not is_bonus and harga <= 0:
                missing_fields.append("Harga (harus > 0 kecuali Bonus)")

            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">❌ {overlap_msg}</div>',
                        unsafe_allow_html=True
                    )

            if cat or name or sz or sup or qty > 0 or harga > 0:
                if missing_fields or validation_errors:
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    error_display.extend(validation_errors)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Barang {i}: <strong>{" | ".join(error_display)}</strong></div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">✅ Barang {i}: Lengkap dan siap dijual</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, True, []))
                    ecer_items.append({
                        "name": name, "quantity": qty, "source": src, "location": src,
                        "price": harga, "category": cat, "color": colr or "Tidak Ada",
                        "size": sz, "supplier": sup or None,
                        "details": {
                            "total_price": harga, "is_bonus": is_bonus, "customer": cust,
                            "phone": phone, "address": alamat, "spg": spg, "payment_type": pay_type,
                            "payment_status": pay_stat, "rekening": rekening, "shipping_status": kirim,
                            "shipping_date": str(kirim_tgl) if kirim_tgl else None
                        }
                    })

    col_add_ecer, col_delete_ecer = st.columns([3, 1])
    with col_add_ecer:
        if st.session_state.ecer_item_count < MAX_ITEMS_PER_FORM:
            if st.button("➕ Tambah Barang Baru", key="add_ecer_item_btn", use_container_width=True):
                st.session_state.ecer_item_count += 1
                st.rerun()
        else:
            st.info(f"ℹ️ Maksimal {MAX_ITEMS_PER_FORM} barang per form")
    with col_delete_ecer:
        if st.session_state.ecer_item_count > 1:
            if st.button("🗑️ Hapus Barang Terakhir", key="delete_ecer_item_btn", use_container_width=True, type="secondary"):
                st.session_state.ecer_item_count -= 1
                st.rerun()

    st.markdown("---")

    total = sum(item['price'] for item in ecer_items)
    if ecer_items:
        expanded_preview = []
        for item in ecer_items:
            expanded_preview.extend(expand_combo_item(item))
        st.success(f"✅ {len(expanded_preview)} barang siap dijual")

    st.markdown("### 💰 Total Penjualan")
    col_total1, col_total2 = st.columns([3, 2])
    with col_total1:
        st.markdown(
            f"<div style='background: linear-gradient(135deg, #1a1a1a 0%, #2a2a2a 100%); "
            f"padding: 2rem; border-radius: 16px; border-left: 4px solid #0d7377; margin-bottom: 1rem;'>"
            f"<h1 style='color: #14ffec; margin: 0; font-size: 3rem;'>{format_currency(total)}</h1></div>",
            unsafe_allow_html=True
        )
    with col_total2:
        if total > 0:
            st.markdown(
                f"<div style='padding: 2rem 1rem;'><strong style='color: #14ffec;'>Terbilang:</strong><br>"
                f"<span style='font-size: 0.95rem;'>{terbilang(total)} rupiah</span></div>",
                unsafe_allow_html=True
            )

    st.markdown("---")

    jumlah_dp = 0
    sisa_pembayaran = 0
    metode_sisa = None
    rekening_sisa = None

    if pay_stat == "DP":
        if total > 0:
            st.markdown("### 💳 Detail Pembayaran DP")
            jumlah_dp = st.number_input(
                "💵 Jumlah DP (Rp)", min_value=0, max_value=total, step=1000, key="ecer_dp"
            )
            sisa_pembayaran = total - jumlah_dp
            col_dp1, col_dp2 = st.columns(2)
            col_dp1.metric("✅ DP Dibayar", format_currency(jumlah_dp))
            col_dp2.metric("⏳ Sisa Pembayaran", format_currency(sisa_pembayaran))

            st.markdown("---")
            st.markdown("#### 💳 Info Pembayaran Sisa")
            col_sisa1, col_sisa2 = st.columns(2)

            metode_sisa_options = ["-- Pilih Metode --"] + PAYMENT_TYPES
            metode_sisa_selected = col_sisa1.selectbox("Metode Pembayaran Sisa *", metode_sisa_options, key="ecer_metode_sisa")
            if metode_sisa_selected == "Lainnya (ketik manual)":
                metode_sisa = col_sisa1.text_input("Metode Sisa (manual)", key="ecer_metode_sisa_man").strip() or "Manual"
            elif metode_sisa_selected == "-- Pilih Metode --":
                metode_sisa = None
            else:
                metode_sisa = metode_sisa_selected

            if metode_sisa == "Transfer" or metode_sisa_selected == "Transfer":
                rek_sisa_options = ["-- Pilih Rekening --"] + REKENING_OPTIONS
                rek_sisa_opt = col_sisa2.selectbox("Rekening untuk Sisa *", rek_sisa_options, key="ecer_rek_sisa")
                if rek_sisa_opt == "Lainnya (ketik manual)":
                    rekening_sisa = col_sisa2.text_input("Rekening Sisa manual", key="ecer_rek_sisa_man").strip() or "Manual"
                elif rek_sisa_opt == "-- Pilih Rekening --":
                    rekening_sisa = None
                else:
                    rekening_sisa = rek_sisa_opt

            for item in ecer_items:
                item['details']['jumlah_dp'] = jumlah_dp
                item['details']['sisa_pembayaran'] = sisa_pembayaran
                item['details']['metode_sisa'] = metode_sisa
                item['details']['rekening_sisa'] = rekening_sisa

            st.markdown("---")

    if st.button("💳 Simpan Penjualan", type="primary", use_container_width=True):
        validation_errors = []
        if not cust:
            validation_errors.append("❌ **Nama Pelanggan** wajib diisi!")
        if not alamat:
            validation_errors.append("❌ **Alamat** wajib diisi!")
        if not spg or spg == "-- Pilih SPG --":
            validation_errors.append("❌ **SPG** wajib dipilih!")
        if not pay_type or pay_type == "-- Pilih Pembayaran --":
            validation_errors.append("❌ **Tipe Pembayaran** wajib dipilih!")
        if pay_type == "Transfer" and (not rekening or rekening == "-- Pilih Rekening --"):
            validation_errors.append("❌ **Rekening** wajib dipilih untuk Transfer!")
        if not pay_stat or pay_stat == "-- Pilih Status --":
            validation_errors.append("❌ **Status Pembayaran** wajib dipilih!")
        if not kirim or kirim == "-- Pilih Status --":
            validation_errors.append("❌ **Status Pengiriman** wajib dipilih!")
        if not ecer_items:
            validation_errors.append("❌ **Belum ada barang yang lengkap**.")
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, missing in incomplete_items:
                validation_errors.append(f"❌ **Barang {item_num}**: {', '.join(missing)}")
        if pay_stat == "DP" and jumlah_dp <= 0:
            validation_errors.append("❌ **Jumlah DP** harus > 0!")

        if validation_errors:
            st.error("### ⚠️ Validasi Gagal:")
            for error in validation_errors:
                st.markdown(error)
            return

        success_count = 0
        all_items = []
        for item in ecer_items:
            all_items.extend(expand_combo_item(item))

        for item in all_items:
            try:
                if not is_pabrik_source(item['location']):
                    try:
                        reduce_inventory_quantity(
                            name=item['name'], category=item['category'], color=item['color'],
                            size=item['size'], location=item['location'], quantity=item['quantity']
                        )
                    except Exception:
                        pass
                log_transaction(
                    action_type="ECERAN", item_name=item['name'], quantity=-item['quantity'],
                    category=item['category'], color=item['color'], size=item['size'],
                    location=item['location'], supplier=item['supplier'],
                    details=item['details'], transaction_date=tanggal_jual
                )
                success_count += 1
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

        if success_count > 0:
            st.success(f"✅ Berhasil menyimpan {success_count} item!")
            st.balloons()
            invalidate_caches()
            st.session_state.force_refresh_dashboard = True
            keys_to_clear = [k for k in st.session_state.keys() if k.startswith('ecer_') or k.startswith('ec_')]
            for key in keys_to_clear:
                del st.session_state[key]
            time.sleep(1)
            st.rerun()


# ==================================================
# STOK TAB
# ==================================================

def render_stok_tab():
    st.caption("⚠️ Stok negatif menunjukkan item terjual sebelum tercatat masuk")

    col_f1, col_f2 = st.columns(2)
    search_query = col_f1.text_input(
        "🔍 Cari", placeholder="Ketik nama barang, supplier, kategori, ukuran, warna..."
    ).strip().lower()

    location_options = ["Semua Lokasi", "Toko Only", "Gudang Only", "Pabrik Only"]
    location_filter = col_f2.multiselect(
        "📍 Filter Lokasi", location_options, default=["Semua Lokasi"], key="stok_location_filter"
    )

    try:
        inventory = get_cached_inventory()

        if "Semua Lokasi" not in location_filter and location_filter:
            filtered_inventory = []
            for item in inventory:
                loc = item.get('location', '').lower()
                if "Toko Only" in location_filter and "toko" in loc:
                    filtered_inventory.append(item)
                elif "Gudang Only" in location_filter and "gudang" in loc:
                    filtered_inventory.append(item)
                elif "Pabrik Only" in location_filter and "pabrik" in loc:
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        else:
            inventory = [item for item in inventory if not is_pabrik_source(item.get('location', ''))]

        if inventory:
            df_data = []
            for item in inventory:
                tgl_isi = item.get('created_at')
                tgl_isi_str = "—"
                if tgl_isi:
                    tgl_isi_local = convert_to_local_time(tgl_isi)
                    tgl_isi_str = tgl_isi_local.strftime("%d/%m/%y %H:%M") if tgl_isi_local else "—"

                tgl_update = item.get('date_added') or item.get('updated_at')
                tgl_update_str = format_date_dd_mm_yy(tgl_update)

                item_name = item['name'].title()
                size = item.get('size', '').strip()
                color = item.get('color', '').strip()
                merged_name = item_name
                if size and size != '-' and size.lower() != 'tidak ada':
                    merged_name += f" ({size})"
                if color and color != '-' and color.lower() != 'tidak ada':
                    merged_name += f" - {color}"

                df_data.append({
                    "Tanggal Input": tgl_isi_str,
                    "Terakhir Update": tgl_update_str,
                    "Supplier": item.get('supplier', '—'),
                    "Kategori": item.get('category', '—').upper(),
                    "Nama Barang": merged_name,
                    "Qty": item['quantity'],
                    "Lokasi": item.get('location', '—')
                })

            df = pd.DataFrame(df_data)

            if search_query:
                mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                df = df[mask]

            col1, col2, col3 = st.columns(3)
            col1.metric("📦 Total Items", len(df))
            col2.metric("📊 Total Qty", f"{df['Qty'].sum():,}")
            negative_count = len(df[df['Qty'] < 0])
            if negative_count > 0:
                col3.metric("⚠️ Stok Negatif", negative_count, delta=f"-{negative_count}", delta_color="inverse")
            else:
                col3.metric("✅ Stok Normal", len(df))

            st.markdown("---")

            def highlight_negative(row):
                if row['Qty'] < 0:
                    return ['background-color: #442222; color: #ffaaaa'] * len(row)
                return [''] * len(row)

            if not df.empty:
                st.dataframe(df.style.apply(highlight_negative, axis=1), use_container_width=True, hide_index=True, height=1500)
            else:
                st.info(f"🔭 Tidak ada data yang cocok dengan pencarian '{search_query}'")
        else:
            st.info("🔭 Tidak ada stok yang ditemukan")

    except Exception as e:
        st.error(f"❌ Error saat memuat data stok: {str(e)}")


# ==================================================
# EDIT STOK TAB
# ==================================================

def render_edit_stok_tab():
    if not st.session_state.get('stock_edit_unlocked', False):
        st.warning("🔒 Mode Edit terkunci - masukkan password")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            pwd = st.text_input("Password Edit Stok", type="password", key="edit_pwd_input")
            if st.button("🔓 Buka Mode Edit", key="unlock_edit_btn", use_container_width=True):
                if pwd == STOCK_EDIT_PASSWORD:
                    st.session_state.stock_edit_unlocked = True
                    st.rerun()
                else:
                    st.error("❌ Password salah!")
        st.stop()

    st.success("✅ Mode Edit Spreadsheet AKTIF")

    col_f1, col_f2 = st.columns(2)
    search_query = col_f1.text_input(
        "🔍 Cari", placeholder="Ketik kata kunci..."
    ).strip().lower()

    location_options = ["Semua Lokasi", "Toko Only", "Gudang Only", "Pabrik Only"]
    location_filter = col_f2.multiselect(
        "📍 Filter Lokasi", location_options, default=["Semua Lokasi"], key="edit_location_filter"
    )

    try:
        inventory = get_cached_inventory()

        if "Semua Lokasi" not in location_filter and location_filter:
            filtered_by_location = []
            for item in inventory:
                loc = item.get('location', '').lower()
                if "Toko Only" in location_filter and "toko" in loc:
                    filtered_by_location.append(item)
                elif "Gudang Only" in location_filter and "gudang" in loc:
                    filtered_by_location.append(item)
                elif "Pabrik Only" in location_filter and "pabrik" in loc:
                    filtered_by_location.append(item)
            inventory = filtered_by_location
        else:
            inventory = [item for item in inventory if not is_pabrik_source(item.get('location', ''))]

        df_data = []
        for i, item in enumerate(inventory):
            details = item.get('details', {})
            komen_val = details.get('komen', '—') if isinstance(details, dict) else '—'

            tgl_input = item.get('created_at')
            tgl_input_str = "—"
            if tgl_input:
                tgl_input_local = convert_to_local_time(tgl_input)
                tgl_input_str = tgl_input_local.strftime("%d/%m/%y %H:%M") if tgl_input_local else "—"

            tgl_update = item.get('date_added') or item.get('updated_at')
            tgl_update_str = format_date_dd_mm_yy(tgl_update)

            item_name = str(item.get('name', '')).title()
            size = str(item.get('size') or '').strip()
            color = str(item.get('color') or '').strip()
            merged_name = item_name
            if size and size != '-' and size.lower() != 'tidak ada':
                merged_name += f" ({size})"
            if color and color != '-' and color.lower() != 'tidak ada':
                merged_name += f" - {color}"

            df_data.append({
                "ID": i + 1,
                "Tanggal Input": tgl_input_str,
                "Terakhir Update": tgl_update_str,
                "PO": str(item.get('nomor_po') or '—'),
                "Supplier": str(item.get('supplier') or '—'),
                "Kategori": str(item.get('category', '—')).upper(),
                "Nama Barang": merged_name,
                "Qty": int(item['quantity']),
                "Lokasi": str(item.get('location', '—')),
                "Komen": str(komen_val),
                "UUID": item.get('id')
            })

        df = pd.DataFrame(df_data)

        if search_query:
            mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
            df_filtered = df[mask]
        else:
            df_filtered = df

        st.info("💡 Klik sel untuk edit, '+' untuk tambah baris, atau pilih baris & tekan 'Delete' untuk hapus.")

        edited_df = st.data_editor(
            df_filtered,
            column_config={
                "ID": st.column_config.NumberColumn(disabled=True),
                "Tanggal Input": st.column_config.TextColumn("📅 Tgl Input Form", disabled=True),
                "Terakhir Update": st.column_config.TextColumn("🔄 Terakhir Update", disabled=True),
                "UUID": None,
                "Lokasi": st.column_config.SelectboxColumn("Lokasi", options=STANDARD_LOCATIONS),
                "Qty": st.column_config.NumberColumn("Qty", min_value=0),
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="stock_editor_spreadsheet"
        )

        st.markdown("---")
        st.markdown("#### 📥 Download Data")
        create_download_button_excel(
            df_filtered,
            f"stok_edit_{date.today().strftime('%d_%m_%y')}.xlsx",
            "📥 Download Excel",
            "Stock Edit Report"
        )
        st.markdown("---")

        if st.button("💾 Simpan Perubahan Spreadsheet", type="primary", use_container_width=True):
            state = st.session_state["stock_editor_spreadsheet"]
            success_total = 0

            def to_python_int(val):
                if val is None or pd.isna(val):
                    return 0
                try:
                    return int(float(val))
                except Exception:
                    return 0

            def to_python_str(val):
                if val is None or pd.isna(val):
                    return ""
                return str(val).strip()

            for idx in state.get("deleted_rows", []):
                try:
                    db_id = df_filtered.iloc[idx]["UUID"]
                    delete_inventory_item(db_id)
                    success_total += 1
                except Exception:
                    pass

            for idx, changes in state.get("edited_rows", {}).items():
                try:
                    row_original = df_filtered.iloc[idx]
                    db_id = row_original["UUID"]
                    update_inventory_item(
                        item_id=db_id,
                        name=to_python_str(changes.get("Barang", row_original.get("Barang", ""))).lower(),
                        category=to_python_str(changes.get("Kategori", row_original["Kategori"])).lower(),
                        color=to_python_str(changes.get("Warna", row_original.get("Warna", ""))),
                        size=to_python_str(changes.get("Ukuran", row_original.get("Ukuran", ""))),
                        quantity=to_python_int(changes.get("Qty", row_original["Qty"])),
                        location=to_python_str(changes.get("Lokasi", row_original["Lokasi"])),
                        supplier=to_python_str(changes.get("Supplier", row_original["Supplier"]))
                    )
                    success_total += 1
                except Exception as e:
                    st.error(f"Gagal update baris {idx + 1}: {e}")

            for row_new in state.get("added_rows", []):
                nama_item = to_python_str(row_new.get("Barang", ""))
                if nama_item:
                    try:
                        add_inventory_item(
                            name=nama_item.lower(),
                            category=to_python_str(row_new.get("Kategori", "LAINNYA")).lower(),
                            color=to_python_str(row_new.get("Warna", "Tidak Ada")),
                            size=to_python_str(row_new.get("Ukuran", "-")),
                            quantity=to_python_int(row_new.get("Qty", 0)),
                            location=to_python_str(row_new.get("Lokasi", "Gudang")),
                            supplier=to_python_str(row_new.get("Supplier", "—")),
                            nomor_po=to_python_str(row_new.get("PO", None))
                        )
                        success_total += 1
                    except Exception as e:
                        st.error(f"Gagal tambah barang: {e}")

            if success_total > 0:
                st.success(f"✅ Berhasil memproses {success_total} perubahan!")
                invalidate_caches()
                time.sleep(1)
                st.rerun()

    except Exception as e:
        st.error(f"❌ Terjadi kesalahan sistem: {str(e)}")


# ==================================================
# HISTORY HTML TABLE HELPERS
# ==================================================

def inject_merged_table_css():
    st.markdown("""
    <style>
    .merged-table-container {
        max-height: 600px;
        overflow-y: auto;
        overflow-x: auto;
        border-radius: 8px;
        border: 1px solid #4a5568;
        margin: 1rem 0;
        width: 100%;
    }
    .merged-table {
        width: auto;
        min-width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        table-layout: auto;
    }
    .merged-table thead {
        position: sticky;
        top: 0;
        z-index: 10;
        background-color: #2d3748;
    }
    .merged-table th {
        background-color: #2d3748;
        color: #ffffff;
        padding: 10px 8px;
        text-align: left;
        font-weight: 600;
        border: 1px solid #4a5568;
        font-size: 0.8rem;
        white-space: nowrap;
        min-width: 60px;
    }
    .merged-table tbody { background-color: #1a202c; color: #e2e8f0; }
    .merged-table td {
        padding: 8px 6px;
        border: 1px solid #4a5568;
        vertical-align: middle;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 200px;
    }
    .merged-table tbody tr:hover td { background-color: #2d3748; }
    .merged-cell {
        background-color: #2a3441 !important;
        text-align: center;
        font-weight: 500;
    }
    .merged-table-container::-webkit-scrollbar { width: 10px; height: 10px; }
    .merged-table-container::-webkit-scrollbar-track { background: #1a202c; border-radius: 5px; }
    .merged-table-container::-webkit-scrollbar-thumb { background: #4a5568; border-radius: 5px; }
    .merged-table-container::-webkit-scrollbar-thumb:hover { background: #718096; }
    </style>
    """, unsafe_allow_html=True)


def _get_time_bucket(record, tolerance_minutes=2):
    """Round timestamp to nearest N minutes for grouping."""
    dt = convert_to_local_time(record.get('created_at'))
    if not dt:
        return "unknown"
    minutes = (dt.minute // tolerance_minutes) * tolerance_minutes
    rounded = dt.replace(minute=minutes, second=0, microsecond=0)
    return rounded.strftime("%d/%m/%y %H:%M")


def render_merged_table_masuk(records: List[Dict]) -> str:
    """Generate HTML table for Barang Masuk with merged cells."""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"

    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())

    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))

    html = '<div class="merged-table-container"><table class="merged-table">'
    html += '''<thead><tr>
        <th>ID</th><th>UUID</th><th>Tanggal Input Form</th><th>Masuk Tanggal</th>
        <th>Nomor PO</th><th>Supplier</th><th>Kategori</th>
        <th>Nama Barang, Ukuran, Warna</th>
        <th style="text-align:center;">Quantity</th>
        <th style="text-align:center;">Total Qty</th>
        <th>Lokasi</th><th>Komen</th>
    </tr></thead><tbody>'''

    record_id = 1

    for created_at_key, grp in groupby(
        sorted_records, key=lambda x: _get_time_bucket(x, 2)
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)

        first = group_list[0]
        det_first = first.get('details') or {}

        masuk_tanggal_first = format_date_dd_mm_yy(first['transaction_date'])
        nomor_po_first = det_first.get('nomor_po') or '—'
        supplier_first = first.get('supplier') or '—'
        lokasi_first = first.get('location') or '—'

        same_date = all(format_date_dd_mm_yy(r['transaction_date']) == masuk_tanggal_first for r in group_list)
        same_po = all(((r.get('details') or {}).get('nomor_po') or '—') == nomor_po_first for r in group_list)
        same_supplier = all((r.get('supplier') or '—') == supplier_first for r in group_list)
        same_location = all((r.get('location') or '—') == lokasi_first for r in group_list)

        for idx, rec in enumerate(group_list):
            det = rec.get('details') or {}
            html += '<tr>'
            html += f'<td>{record_id}</td>'
            uuid_val = rec.get('uuid') or rec.get('id') or 'N/A'
            html += f'<td style="font-family:monospace;font-size:0.75rem;">{str(uuid_val)[:8]}</td>'

            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}">{created_at_key}</td>'

            if same_date:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{masuk_tanggal_first}</td>'
            else:
                html += f'<td>{format_date_dd_mm_yy(rec["transaction_date"])}</td>'

            if same_po:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{nomor_po_first}</td>'
            else:
                html += f'<td>{det.get("nomor_po") or "—"}</td>'

            if same_supplier:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{supplier_first}</td>'
            else:
                html += f'<td>{rec.get("supplier") or "—"}</td>'

            html += f'<td>{(rec.get("category") or "—").upper()}</td>'

            combined_name = format_item_combined(
                rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')
            )
            html += f'<td>{combined_name}</td>'
            html += f'<td style="text-align:center;">{abs(rec["quantity"])}</td>'

            if idx == 0:
                html += (
                    f'<td class="merged-cell" rowspan="{group_size}" '
                    f'style="font-weight:bold;text-align:center;">{total_qty}</td>'
                )

            if same_location:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{lokasi_first}</td>'
            else:
                html += f'<td>{rec.get("location") or "—"}</td>'

            html += f'<td>{det.get("komen") or "—"}</td>'
            html += '</tr>'
            record_id += 1

    html += '</tbody></table></div>'
    return html


def render_merged_table_luar_kota(records: List[Dict]) -> str:
    """Generate HTML table for Luar Kota with merged cells."""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"

    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())

    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))

    html = '<div class="merged-table-container"><table class="merged-table">'
    html += '''<thead><tr>
        <th>ID</th><th>UUID</th><th>Tanggal Input Form</th><th>Pengiriman Tanggal</th>
        <th>Tujuan</th><th>Via</th><th>Supplier</th><th>Kategori</th>
        <th>Nama Barang, Ukuran, Warna</th>
        <th style="text-align:center;">Quantity</th>
        <th style="text-align:center;">Total Qty</th>
        <th>Lokasi</th><th>Komen</th>
    </tr></thead><tbody>'''

    record_id = 1

    for created_at_key, grp in groupby(
        sorted_records, key=lambda x: _get_time_bucket(x, 2)
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)

        first = group_list[0]
        det_first = first.get('details', {})

        tanggal_kirim_first = format_date_dd_mm_yy(first['transaction_date'])
        tujuan_first = det_first.get('purpose', '—')
        via_first = det_first.get('via', '—')
        supplier_first = first.get('supplier', '—')
        lokasi_first = first.get('location', '—')

        same_date = all(format_date_dd_mm_yy(r['transaction_date']) == tanggal_kirim_first for r in group_list)
        same_tujuan = all((r.get('details', {}).get('purpose') or '—') == tujuan_first for r in group_list)
        same_via = all((r.get('details', {}).get('via') or '—') == via_first for r in group_list)
        same_supplier = all((r.get('supplier') or '—') == supplier_first for r in group_list)
        same_location = all((r.get('location') or '—') == lokasi_first for r in group_list)

        for idx, rec in enumerate(group_list):
            det = rec.get('details', {})
            html += '<tr>'
            html += f'<td>{record_id}</td>'
            html += f'<td style="font-family:monospace;font-size:0.75rem;">{str(rec.get("uuid", "N/A"))[:8]}</td>'

            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}">{created_at_key}</td>'

            if same_date:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{tanggal_kirim_first}</td>'
            else:
                html += f'<td>{format_date_dd_mm_yy(rec["transaction_date"])}</td>'

            if same_tujuan:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{tujuan_first}</td>'
            else:
                html += f'<td>{det.get("purpose", "—")}</td>'

            if same_via:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{via_first}</td>'
            else:
                html += f'<td>{det.get("via", "—")}</td>'

            if same_supplier:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{supplier_first}</td>'
            else:
                html += f'<td>{rec.get("supplier", "—")}</td>'

            html += f'<td>{(rec.get("category") or "—").upper()}</td>'

            combined_name = format_item_combined(
                rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')
            )
            html += f'<td>{combined_name}</td>'
            html += f'<td style="text-align:center;">{abs(rec["quantity"])}</td>'

            if idx == 0:
                html += (
                    f'<td class="merged-cell" rowspan="{group_size}" '
                    f'style="font-weight:bold;text-align:center;">{total_qty}</td>'
                )

            if same_location:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{lokasi_first}</td>'
            else:
                html += f'<td>{rec.get("location", "—")}</td>'

            html += f'<td>{det.get("komen", "—")}</td>'
            html += '</tr>'
            record_id += 1

    html += '</tbody></table></div>'
    return html


def render_merged_table_eceran(records: List[Dict]) -> str:
    """Generate HTML table for Eceran with merged cells. (FIXED - was incomplete)"""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"

    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())

    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))

    html = '<div class="merged-table-container"><table class="merged-table">'
    html += '''<thead><tr>
        <th>ID</th><th>UUID</th><th>Tanggal Input Form</th><th>Tanggal Penjualan</th>
        <th>Pelanggan</th><th>Supplier</th><th>Kategori</th>
        <th>Nama Barang, Ukuran, Warna</th>
        <th style="text-align:center;">Quantity</th>
        <th style="text-align:center;">Total Qty</th>
        <th style="text-align:right;">Harga Per Item</th>
        <th style="text-align:right;">Total Penjualan</th>
        <th>Status</th><th>Pembayaran</th><th>Rekening</th>
        <th style="text-align:right;">Total DP</th>
        <th style="text-align:right;">Sisa Pembayaran</th>
        <th>Metode Sisa</th><th>Rekening Sisa</th>
        <th>Lokasi</th><th>Tgl Kirim</th><th>Komen</th>
    </tr></thead><tbody>'''

    record_id = 1

    for created_at_key, grp in groupby(
        sorted_records, key=lambda x: _get_time_bucket(x, 2)
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)
        total_sales = sum((r.get('details') or {}).get('total_price', 0) for r in group_list)

        first = group_list[0]
        det_first = first.get('details') or {}

        tanggal_jual_first = format_date_dd_mm_yy(first['transaction_date'])
        pelanggan_first = det_first.get('customer', '—')
        supplier_first = first.get('supplier', '—')
        lokasi_first = first.get('location', '—')
        status_first = det_first.get('payment_status', '—')
        pembayaran_first = det_first.get('payment_type', '—')
        rekening_first = det_first.get('rekening', '—')

        same_date = all(format_date_dd_mm_yy(r['transaction_date']) == tanggal_jual_first for r in group_list)
        same_pelanggan = all(((r.get('details') or {}).get('customer') or '—') == pelanggan_first for r in group_list)
        same_supplier = all((r.get('supplier') or '—') == supplier_first for r in group_list)
        same_location = all((r.get('location') or '—') == lokasi_first for r in group_list)
        same_status = all(((r.get('details') or {}).get('payment_status') or '—') == status_first for r in group_list)
        same_pembayaran = all(((r.get('details') or {}).get('payment_type') or '—') == pembayaran_first for r in group_list)
        same_rekening = all(((r.get('details') or {}).get('rekening') or '—') == rekening_first for r in group_list)

        for idx, rec in enumerate(group_list):
            det = rec.get('details') or {}
            html += '<tr>'
            html += f'<td>{record_id}</td>'
            html += f'<td style="font-family:monospace;font-size:0.75rem;">{str(rec.get("uuid", "N/A"))[:8]}</td>'

            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}">{created_at_key}</td>'

            if same_date:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{tanggal_jual_first}</td>'
            else:
                html += f'<td>{format_date_dd_mm_yy(rec["transaction_date"])}</td>'

            if same_pelanggan:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{pelanggan_first}</td>'
            else:
                html += f'<td>{det.get("customer", "—")}</td>'

            if same_supplier:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{supplier_first}</td>'
            else:
                html += f'<td>{rec.get("supplier", "—")}</td>'

            html += f'<td>{(rec.get("category") or "—").upper()}</td>'

            combined_name = format_item_combined(
                rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')
            )
            html += f'<td>{combined_name}</td>'
            html += f'<td style="text-align:center;">{abs(rec["quantity"])}</td>'

            if idx == 0:
                html += (
                    f'<td class="merged-cell" rowspan="{group_size}" '
                    f'style="font-weight:bold;text-align:center;">{total_qty}</td>'
                )

            html += f'<td style="text-align:right;">{format_currency(det.get("total_price", 0))}</td>'

            if idx == 0:
                html += (
                    f'<td class="merged-cell" rowspan="{group_size}" '
                    f'style="font-weight:bold;text-align:right;">{format_currency(total_sales)}</td>'
                )

            if same_status:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{status_first}</td>'
            else:
                html += f'<td>{det.get("payment_status", "—")}</td>'

            if same_pembayaran:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{pembayaran_first}</td>'
            else:
                html += f'<td>{det.get("payment_type", "—")}</td>'

            if same_rekening:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{rekening_first}</td>'
            else:
                html += f'<td>{det.get("rekening", "—")}</td>'

            # DP fields - per row (not merged, can differ per item)
            html += f'<td style="text-align:right;">{format_currency(det.get("jumlah_dp", 0))}</td>'
            html += f'<td style="text-align:right;">{format_currency(det.get("sisa_pembayaran", 0))}</td>'
            html += f'<td>{det.get("metode_sisa", "—")}</td>'
            html += f'<td>{det.get("rekening_sisa", "—")}</td>'

            if same_location:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{lokasi_first}</td>'
            else:
                html += f'<td>{rec.get("location", "—")}</td>'

            html += f'<td>{det.get("shipping_date", "—")}</td>'
            html += f'<td>{det.get("komen", "—")}</td>'
            html += '</tr>'
            record_id += 1

    html += '</tbody></table></div>'
    return html


# ==================================================
# INLINE EDITOR FOR HISTORY
# ==================================================

def render_inline_editor(records: list, action_type: str, icon: str):
    """Render a contextual editor for each history section."""
    with st.expander(f"{icon} Kelola / Edit Transaksi {action_type.replace('_', ' ').title()}", expanded=False):
        record_map = {
            f"No: {i + 1} | {r['item_name'].title()} ({r['quantity']}) | {r['uuid'][:8]}": r
            for i, r in enumerate(records)
        }
        selected_key = st.selectbox(
            f"Pilih item {action_type} untuk diubah:",
            ["-- Pilih --"] + list(record_map.keys()),
            key=f"sel_{action_type}"
        )

        if selected_key != "-- Pilih --":
            rec = record_map[selected_key]
            st.markdown(f"**🛠️ Mode Edit Transaksi: {rec['uuid'][:8]}**")

            with st.form(key=f"form_hist_{rec['uuid']}"):
                col1, col2 = st.columns(2)
                new_name = col1.text_input("Nama Barang", value=rec['item_name']).strip().lower()
                new_cat = col2.text_input("Kategori", value=rec.get('category', '')).strip().lower()

                col3, col4, col5 = st.columns(3)
                new_qty = col3.number_input("Quantity", value=float(rec['quantity']))
                new_sz = col4.text_input("Ukuran", value=rec.get('size') or "")
                new_col = col5.text_input("Warna", value=rec.get('color', ''))

                new_date = st.date_input("Tanggal Transaksi", value=rec['transaction_date'])

                c_edit, c_del = st.columns([2, 1])

                if c_edit.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
                    success, msg = update_history_record(
                        rec['uuid'],
                        item_name=new_name,
                        category=new_cat,
                        quantity=new_qty,
                        size=new_sz if new_sz else None,
                        color=new_col,
                        transaction_date=new_date
                    )
                    if success:
                        st.success("✅ Riwayat diperbarui!")
                        invalidate_caches()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)

                if c_del.form_submit_button("🗑️ Hapus Transaksi", type="secondary", use_container_width=True):
                    if rec['action_type'] == "MASUK" and not is_pabrik_source(rec.get('location', '')):
                        try:
                            reduce_inventory_quantity(
                                name=rec['item_name'], category=rec.get('category'),
                                color=rec.get('color'), size=rec.get('size'),
                                location=rec.get('location'), quantity=abs(rec['quantity'])
                            )
                        except Exception:
                            pass
                    elif rec['action_type'] in ["LUAR_KOTA", "ECERAN"] and not is_pabrik_source(rec.get('location', '')):
                        try:
                            add_inventory_item(
                                name=rec['item_name'], category=rec.get('category'),
                                color=rec.get('color'), size=rec.get('size'),
                                quantity=abs(rec['quantity']), location=rec.get('location'),
                                supplier=rec.get('supplier'), date_added=date.today()
                            )
                        except Exception:
                            pass

                    success, msg = delete_history_record(rec['uuid'])
                    if success:
                        st.success("✅ Transaksi dihapus & Stok dibalikkan!")
                        invalidate_caches()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)


# ==================================================
# HISTORY TAB
# ==================================================

def render_history_tab(today: date) -> None:
    """History tab with HTML merged-cell tables + filters + Excel download."""

    if not st.session_state.get('history_view_unlocked', False):
        st.warning("🔒 Riwayat terkunci - masukkan password")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            pwd = st.text_input("Password Riwayat", type="password", key="hist_pwd_input")
            if st.button("🔓 Buka Riwayat", key="unlock_history_btn", use_container_width=True):
                if pwd == HISTORY_VIEW_PASSWORD:
                    st.session_state.history_view_unlocked = True
                    st.success("✅ Akses diberikan!")
                    st.rerun()
                else:
                    st.error("❌ Password salah!")
        st.stop()

    st.success("✅ Mode Riwayat AKTIF")
    st.markdown("---")

    st.markdown("#### 🔍 Filter Data")
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Dari", today - timedelta(days=30), key="hist_from", format="DD/MM/YYYY")
    end_date = col2.date_input("Sampai", today, min_value=start_date, key="hist_to", format="DD/MM/YYYY")

    col3, col4 = st.columns(2)
    activity_options = ["Semua", "Masuk Barang", "Luar Kota", "Eceran"]
    activity_selected = col3.selectbox("Jenis Aktivitas", activity_options, key="hist_activity_filter")
    global_search = col4.text_input(
        "🔍 Cari", key="hist_global_search", placeholder="Ketik kata kunci..."
    ).strip().lower()

    activity_map = {"Semua": None, "Masuk Barang": "MASUK", "Luar Kota": "LUAR_KOTA", "Eceran": "ECERAN"}
    activity_filter = activity_map[activity_selected]

    with st.expander("🎯 Filter Lanjutan", expanded=False):
        col_adv1, col_adv2, col_adv3 = st.columns(3)
        all_suppliers = sorted(list(set(st.session_state.get("dynamic_supplier", SUPPLIERS))))
        all_kategori = sorted(list(set(st.session_state.get("dynamic_kategori", STANDARD_CATEGORIES))))
        all_locations = sorted(list(set(STANDARD_LOCATIONS)))

        supplier_selected = col_adv1.selectbox("Supplier", ["Semua"] + all_suppliers, key="hist_supplier_filter")
        supplier_filter = None if supplier_selected == "Semua" else supplier_selected

        kategori_selected = col_adv2.selectbox("Kategori", ["Semua"] + all_kategori, key="hist_kategori_filter")
        kategori_filter = None if kategori_selected == "Semua" else kategori_selected

        location_selected = col_adv3.selectbox("Lokasi", ["Semua"] + all_locations, key="hist_location_filter")
        location_filter = None if location_selected == "Semua" else location_selected

    kirim_ke_filter = None
    via_filter = None
    payment_type_filter = None
    payment_status_filter = None
    use_price_filter = False
    price_min = None
    price_max = None

    if activity_filter == "LUAR_KOTA":
        st.markdown("##### 🚚 Filter Luar Kota")
        col_lk1, col_lk2 = st.columns(2)
        toko_selected = col_lk1.selectbox("Kirim Ke (Toko)", ["Semua"] + sorted(DAFTAR_TOKO), key="hist_lk_toko")
        if toko_selected != "Semua":
            kirim_ke_filter = toko_selected
        via_selected = col_lk2.selectbox("Via Pengiriman", ["Semua"] + sorted(VIA_OPTIONS), key="hist_lk_via")
        if via_selected != "Semua":
            via_filter = via_selected

    elif activity_filter == "ECERAN":
        st.markdown("##### 🛒 Filter Eceran")
        col_ec1, col_ec2 = st.columns(2)
        payment_type_selected = col_ec1.selectbox("Tipe Pembayaran", ["Semua"] + sorted(PAYMENT_TYPES), key="hist_ec_payment_type")
        if payment_type_selected != "Semua":
            payment_type_filter = payment_type_selected
        payment_status_selected = col_ec2.selectbox("Status Pembayaran", ["Semua"] + sorted(PAYMENT_STATUS), key="hist_ec_payment_status")
        if payment_status_selected != "Semua":
            payment_status_filter = payment_status_selected

        st.markdown("**💰 Filter Total Pembayaran (Rp)**")
        col_price1, col_price2 = st.columns(2)
        use_price_filter = col_price1.checkbox("Aktifkan filter harga", key="hist_ec_use_price")
        if use_price_filter:
            price_min = col_price1.number_input("Harga Minimum (Rp)", min_value=0, step=10000, key="hist_ec_price_min")
            price_max = col_price2.number_input("Harga Maximum (Rp)", min_value=0, step=10000, key="hist_ec_price_max")
            if price_max > 0 and price_max < price_min:
                st.warning("⚠️ Harga maximum harus lebih besar dari minimum")

    try:
        all_history = get_history(start_date=start_date, end_date=end_date)
        filtered_history = []

        for rec in all_history:
            if activity_filter and rec['action_type'] != activity_filter:
                continue

            if global_search:
                searchable_text = (
                    f"{rec.get('item_name', '')} {rec.get('supplier', '')} "
                    f"{rec.get('category', '')} {rec.get('size', '')} {rec.get('color', '')} "
                    f"{rec.get('location', '')} "
                    f"{rec.get('details', {}).get('nomor_po', '')} "
                    f"{rec.get('details', {}).get('customer', '')} "
                    f"{rec.get('details', {}).get('purpose', '')} "
                    f"{rec.get('details', {}).get('via', '')}"
                ).lower()
                if global_search not in searchable_text:
                    continue

            if supplier_filter:
                if not rec.get('supplier') or rec.get('supplier', '').lower() != supplier_filter.lower():
                    continue

            if kategori_filter:
                if not rec.get('category') or rec.get('category', '').lower() != kategori_filter.lower():
                    continue

            if location_filter:
                if not rec.get('location') or rec.get('location', '').lower() != location_filter.lower():
                    continue

            details = rec.get('details', {})

            if activity_filter == "LUAR_KOTA":
                if kirim_ke_filter:
                    purpose = details.get('purpose', '')
                    if not purpose or purpose.lower() != kirim_ke_filter.lower():
                        continue
                if via_filter:
                    via = details.get('via', '')
                    if not via or via.lower() != via_filter.lower():
                        continue

            if activity_filter == "ECERAN":
                if payment_type_filter:
                    if not details.get('payment_type') or details.get('payment_type', '').lower() != payment_type_filter.lower():
                        continue
                if payment_status_filter:
                    if not details.get('payment_status') or details.get('payment_status', '').lower() != payment_status_filter.lower():
                        continue
                if use_price_filter:
                    total_price = details.get('total_price', 0)
                    if price_min and total_price < price_min:
                        continue
                    if price_max and total_price > price_max:
                        continue

            filtered_history.append(rec)

        if not filtered_history:
            st.info("🔭 Tidak ada data yang sesuai dengan filter")
            return

        st.markdown("---")
        st.success(f"✅ Menampilkan {len(filtered_history)} transaksi")
        st.markdown("---")

        inject_merged_table_css()

        # ── BARANG MASUK ──
        st.markdown("### 📥 Barang Masuk")
        masuk_records = [r for r in filtered_history if r['action_type'] == 'MASUK']
        if masuk_records:
            excel_masuk = []
            for rec in masuk_records:
                details = rec.get('details', {})
                excel_masuk.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Masuk': format_date_dd_mm_yy(rec['transaction_date']),
                    'Nomor PO': details.get('nomor_po', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': format_item_combined(rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')),
                    'Qty': abs(rec['quantity']),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            create_download_button_excel(
                pd.DataFrame(excel_masuk),
                f"barang_masuk_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Barang Masuk", "Barang Masuk"
            )
            st.markdown("---")
            st.markdown(render_merged_table_masuk(masuk_records), unsafe_allow_html=True)
            st.markdown("---")
            render_inline_editor(masuk_records, "MASUK", "📥")
        else:
            st.info("Tidak ada riwayat barang masuk.")

        # ── LUAR KOTA ──
        st.markdown("---")
        st.markdown("### 🚚 Pengeluaran Luar Kota")
        luar_records = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA']
        if luar_records:
            excel_luar = []
            for rec in luar_records:
                details = rec.get('details', {})
                excel_luar.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Kirim': format_date_dd_mm_yy(rec['transaction_date']),
                    'Tujuan': details.get('purpose', '—'),
                    'Via': details.get('via', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': format_item_combined(rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')),
                    'Qty': abs(rec['quantity']),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            create_download_button_excel(
                pd.DataFrame(excel_luar),
                f"luar_kota_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Luar Kota", "Luar Kota"
            )
            st.markdown("---")
            st.markdown(render_merged_table_luar_kota(luar_records), unsafe_allow_html=True)
            st.markdown("---")
            render_inline_editor(luar_records, "LUAR_KOTA", "🚚")
        else:
            st.info("Tidak ada riwayat pengeluaran luar kota.")

        # ── ECERAN ──
        st.markdown("---")
        st.markdown("### 🛒 Penjualan Eceran")
        ecer_records = [r for r in filtered_history if r['action_type'] == 'ECERAN']
        if ecer_records:
            excel_eceran = []
            for rec in ecer_records:
                details = rec.get('details', {})
                excel_eceran.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Jual': format_date_dd_mm_yy(rec['transaction_date']),
                    'Pelanggan': details.get('customer', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': format_item_combined(rec.get('item_name', ''), rec.get('size', ''), rec.get('color', '')),
                    'Qty': abs(rec['quantity']),
                    'Harga': details.get('total_price', 0),
                    'Status Bayar': details.get('payment_status', '—'),
                    'Tipe Bayar': details.get('payment_type', '—'),
                    'Rekening': details.get('rekening', '—'),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            create_download_button_excel(
                pd.DataFrame(excel_eceran),
                f"eceran_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Eceran", "Penjualan Eceran"
            )
            st.markdown("---")
            st.markdown(render_merged_table_eceran(ecer_records), unsafe_allow_html=True)
            st.markdown("---")
            render_inline_editor(ecer_records, "ECERAN", "🛒")
        else:
            st.info("Tidak ada riwayat penjualan eceran.")

    except Exception as e:
        st.error(f"❌ Error loading history: {str(e)}")
        import traceback
        st.code(traceback.format_exc())


# ==================================================
# SUMMARY TAB
# ==================================================

def render_summary_tab(today: date):

    if not st.session_state.get('summary_unlocked', False):
        st.warning("🔒 Ringkasan terkunci - masukkan password")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            pwd = st.text_input("Password Ringkasan", type="password", key="summary_pwd_input")
            if st.button("🔓 Buka Ringkasan", key="unlock_summary_btn", use_container_width=True):
                if pwd == SUMMARY_PASSWORD:
                    st.session_state.summary_unlocked = True
                    st.success("✅ Akses diberikan!")
                    st.rerun()
                else:
                    st.error("❌ Password salah!")
        st.stop()

    st.success("✅ Mode Ringkasan AKTIF")
    st.markdown("---")

    col1, col2 = st.columns(2)
    def_start = today - timedelta(days=90)
    sum_start = col1.date_input("Dari", def_start, key="sum_from", format="DD/MM/YYYY")
    sum_end = col2.date_input("Sampai", today, min_value=sum_start, key="sum_to", format="DD/MM/YYYY")

    st.markdown("#### 🔎 Filter Data")
    col_f1, col_f2, col_f3 = st.columns(3)

    activity_filter = col_f1.multiselect(
        "Jenis Aktivitas", ["MASUK", "LUAR_KOTA", "ECERAN"], key="sum_activity_multi"
    )
    all_suppliers = st.session_state.get("dynamic_supplier", SUPPLIERS)
    supplier_filter = col_f2.multiselect("Supplier", all_suppliers, key="sum_supplier_multi")
    all_kategori = st.session_state.get("dynamic_kategori", STANDARD_CATEGORIES)
    kategori_filter = col_f3.multiselect("Kategori", all_kategori, key="sum_category_multi")

    try:
        all_history = get_cached_history(start_date=sum_start, end_date=sum_end, item_name=None)

        filtered_history = []
        for rec in all_history:
            if activity_filter and rec['action_type'] not in activity_filter:
                continue
            if supplier_filter and rec.get('supplier') not in supplier_filter:
                continue
            if kategori_filter and rec.get('category') not in kategori_filter:
                continue
            filtered_history.append(rec)

        if not filtered_history:
            st.info("🔭 Tidak ada data yang sesuai dengan filter")
            return

        st.markdown("---")
        st.success(f"✅ Menampilkan {len(filtered_history)} transaksi")

        # ── TOP 10 TABLES ──
        st.markdown("### 📊 Top 10 - Berdasarkan Filter (Bonus Dikecualikan)")

        masuk_data = [r for r in filtered_history if r['action_type'] == 'MASUK' and r.get('category', '').lower() != 'bonus']
        luar_kota_data = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA' and r.get('category', '').lower() != 'bonus']
        eceran_data = [r for r in filtered_history if r['action_type'] == 'ECERAN' and r.get('category', '').lower() != 'bonus']

        if masuk_data:
            with st.expander("📥 Top 10 Barang Masuk", expanded=False):
                masuk_items = {}
                for r in masuk_data:
                    item = r['item_name'].title()
                    masuk_items[item] = masuk_items.get(item, 0) + abs(r['quantity'])
                top_masuk = sorted(masuk_items.items(), key=lambda x: x[1], reverse=True)[:10]
                masuk_df = pd.DataFrame(top_masuk, columns=['Barang', 'Total Qty'])
                masuk_df.index = range(1, len(masuk_df) + 1)
                st.dataframe(masuk_df, use_container_width=True)

        if luar_kota_data:
            with st.expander("🚚 Top 10 Barang Luar Kota", expanded=False):
                luar_items = {}
                for r in luar_kota_data:
                    item = r['item_name'].title()
                    luar_items[item] = luar_items.get(item, 0) + abs(r['quantity'])
                top_luar = sorted(luar_items.items(), key=lambda x: x[1], reverse=True)[:10]
                luar_df = pd.DataFrame(top_luar, columns=['Barang', 'Total Qty'])
                luar_df.index = range(1, len(luar_df) + 1)
                st.dataframe(luar_df, use_container_width=True)

        if eceran_data:
            with st.expander("🛒 Top 10 Barang Eceran", expanded=False):
                eceran_items = {}
                for r in eceran_data:
                    item = r['item_name'].title()
                    eceran_items[item] = eceran_items.get(item, 0) + abs(r['quantity'])
                top_eceran = sorted(eceran_items.items(), key=lambda x: x[1], reverse=True)[:10]
                eceran_df = pd.DataFrame(top_eceran, columns=['Barang', 'Total Qty'])
                eceran_df.index = range(1, len(eceran_df) + 1)
                st.dataframe(eceran_df, use_container_width=True)

        st.markdown("---")

        if not PLOTLY_AVAILABLE:
            st.warning("⚠️ Install plotly: `pip install plotly`")
            return

        st.markdown("""
        <style>
        .chart-label { color: #FFFFFF; font-weight: 600; font-size: 14px; }
        .chart-legend { color: #FFFFFF; font-size: 13px; }
        </style>
        """, unsafe_allow_html=True)

        DARK_MODE_PALETTE = [
            "#29B6F6", "#66BB6A", "#FFCA28", "#FF7043", "#AB47BC",
            "#26C6DA", "#D4E157", "#EC407A", "#00897B", "#FFA726",
            "#5C6BC0", "#F06292", "#00FFFF", "#FFD700", "#DC143C",
        ]

        # ── PIE CHARTS ──
        st.markdown("### 📊 Distribusi Data")

        # Activity distribution
        if filtered_history:
            st.markdown("#### 📊 Distribusi Berdasarkan Jenis Aktivitas")
            non_bonus_history = [r for r in filtered_history if r.get('category', '').lower() != 'bonus']
            activity_counts = {}
            for r in non_bonus_history:
                activity_counts[r['action_type']] = activity_counts.get(r['action_type'], 0) + abs(r['quantity'])
            if activity_counts:
                activity_df = pd.DataFrame([{'Aktivitas': k, 'Jumlah': v} for k, v in activity_counts.items()])
                fig = px.pie(
                    activity_df, values='Jumlah', names='Aktivitas',
                    title='Distribusi Berdasarkan Jenis Aktivitas (Tanpa Bonus)',
                    hole=0.35, color_discrete_sequence=DARK_MODE_PALETTE[:3]
                )
                fig.update_traces(
                    textposition="inside", textinfo="percent+label",
                    texttemplate="%{label}<br>%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                                font=dict(color='#FFFFFF', size=13, family='Arial')),
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)

        # Detailed charts only when filters active
        if activity_filter or supplier_filter or kategori_filter:

            # Supplier distribution
            non_bonus_history = [r for r in filtered_history if r.get('category', '').lower() != 'bonus']
            if non_bonus_history:
                st.markdown("#### 🏭 Distribusi Berdasarkan Supplier")
                supplier_counts = {}
                for r in non_bonus_history:
                    s = r.get('supplier', 'Unknown')
                    supplier_counts[s] = supplier_counts.get(s, 0) + abs(r['quantity'])
                supplier_sorted = sorted(supplier_counts.items(), key=lambda x: x[1], reverse=True)
                if len(supplier_sorted) > 12:
                    top_12 = dict(supplier_sorted[:12])
                    others = sum(v for _, v in supplier_sorted[12:])
                    if others > 0:
                        top_12['Lainnya'] = others
                    supplier_counts = top_12
                supplier_df = pd.DataFrame([{'Supplier': k, 'Jumlah': v} for k, v in supplier_counts.items()])
                fig = px.pie(
                    supplier_df, values='Jumlah', names='Supplier',
                    title='Distribusi Berdasarkan Supplier (Tanpa Bonus)',
                    hole=0.35, color_discrete_sequence=DARK_MODE_PALETTE
                )
                fig.update_traces(
                    textposition="inside", textinfo="percent+label",
                    texttemplate="%{label}<br>%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                                font=dict(color='#FFFFFF', size=13, family='Arial')),
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)

            # Luar Kota toko distribution
            luar_kota_records = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA']
            if luar_kota_records:
                st.markdown("#### 🏪 Distribusi Kirim Ke Toko (Luar Kota)")
                toko_counts = {}
                for r in luar_kota_records:
                    details = r.get('details', {})
                    toko = None
                    if isinstance(details, dict):
                        toko = details.get('purpose') or details.get('kirim_ke') or details.get('tujuan')
                    if not toko or toko == '':
                        toko = 'Data Lama (Toko Tidak Tercatat)'
                    toko_counts[toko] = toko_counts.get(toko, 0) + abs(r['quantity'])

                toko_sorted = sorted(toko_counts.items(), key=lambda x: x[1], reverse=True)
                if len(toko_sorted) > 12:
                    top_12 = dict(toko_sorted[:12])
                    others = sum(v for _, v in toko_sorted[12:])
                    if others > 0:
                        top_12['Lainnya'] = others
                    toko_counts = top_12

                toko_df = pd.DataFrame([{'Toko': k, 'Jumlah': v} for k, v in toko_counts.items()])
                if not toko_df.empty:
                    fig = px.pie(
                        toko_df, values='Jumlah', names='Toko',
                        title='Distribusi Pengiriman per Toko Tujuan',
                        hole=0.35, color_discrete_sequence=DARK_MODE_PALETTE
                    )
                    fig.update_traces(
                        textposition="inside", textinfo="percent+label",
                        texttemplate="%{label}<br>%{percent:.1%}",
                        hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                        textfont=dict(color='#FFFFFF', size=14, family='Arial')
                    )
                    fig.update_layout(
                        showlegend=True,
                        legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                                    font=dict(color='#FFFFFF', size=13, family='Arial')),
                        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#FFFFFF', size=14, family='Arial'),
                        title=dict(font=dict(color='#FFFFFF', size=16))
                    )
                    st.plotly_chart(fig, use_container_width=True)

            # Category distribution (detailed, includes bonus)
            st.markdown("#### 🏷️ Distribusi Berdasarkan Kategori (Detail)")
            category_counts = {}
            for r in filtered_history:
                category = r.get('category', 'unknown').title()
                name = r['item_name'].title()
                size = r.get('size', '').title() if r.get('size') else ''
                color = r.get('color', '').title() if r.get('color') and r.get('color').lower() != 'tidak ada' else ''
                parts = [category, name]
                if size:
                    parts.append(size)
                if color:
                    parts.append(color)
                label = ' - '.join(parts)
                category_counts[label] = category_counts.get(label, 0) + abs(r['quantity'])

            category_sorted = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
            if len(category_sorted) > 15:
                top_15 = dict(category_sorted[:15])
                others = sum(v for _, v in category_sorted[15:])
                if others > 0:
                    top_15['Lainnya'] = others
                category_counts = top_15

            category_df = pd.DataFrame([{'Kategori': k, 'Jumlah': v} for k, v in category_counts.items()])
            if not category_df.empty:
                fig = px.pie(
                    category_df, values='Jumlah', names='Kategori',
                    title='Distribusi Berdasarkan Kategori (Detail)',
                    hole=0.35, color_discrete_sequence=DARK_MODE_PALETTE
                )
                fig.update_traces(
                    textposition="inside", textinfo="percent",
                    texttemplate="%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
                                font=dict(color='#FFFFFF', size=13, family='Arial')),
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # ── MONTHLY STACKED BAR CHARTS ──
        st.markdown("### 📊 Analisis Bulanan Per Supplier (Bonus Dikecualikan)")

        st.markdown("**📅 Pilih Rentang Bulan untuk Chart:**")
        col_month1, col_year1, col_month2, col_year2 = st.columns(4)

        current_date_dt = datetime.now()
        current_year = current_date_dt.year

        months_id = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        available_years = list(range(2020, current_year + 2))
        default_start_dt = current_date_dt - timedelta(days=365)

        start_month_idx = col_month1.selectbox(
            "Dari Bulan", range(12), format_func=lambda x: months_id[x],
            index=default_start_dt.month - 1, key="chart_start_month"
        )
        start_year = col_year1.selectbox(
            "Tahun", available_years,
            index=available_years.index(default_start_dt.year) if default_start_dt.year in available_years else len(available_years) - 2,
            key="chart_start_year"
        )
        end_month_idx = col_month2.selectbox(
            "Sampai Bulan", range(12), format_func=lambda x: months_id[x],
            index=current_date_dt.month - 1, key="chart_end_month"
        )
        end_year = col_year2.selectbox(
            "Tahun", available_years,
            index=available_years.index(current_year) if current_year in available_years else len(available_years) - 1,
            key="chart_end_year"
        )

        start_month = date(start_year, start_month_idx + 1, 1)
        end_month = date(end_year, end_month_idx + 1, 1)

        if start_month > end_month:
            st.error("⚠️ 'Dari Bulan' harus lebih awal dari 'Sampai Bulan'")
            return

        st.caption(f"📊 Menampilkan data dari **{months_id[start_month_idx]} {start_year}** sampai **{months_id[end_month_idx]} {end_year}**")

        # Build all months in range
        all_months_in_range = []
        cur = start_month
        while cur <= end_month:
            all_months_in_range.append((cur.strftime('%Y-%m'), cur.strftime('%b %Y')))
            if cur.month == 12:
                cur = cur.replace(year=cur.year + 1, month=1)
            else:
                cur = cur.replace(month=cur.month + 1)

        all_data_for_charts = [r for r in get_cached_history() if r.get('category', '').lower() != 'bonus']

        # Filter by month range
        month_filtered_data = []
        for r in all_data_for_charts:
            trans_date = r['transaction_date']
            if isinstance(trans_date, str):
                trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
            trans_month = trans_date.replace(day=1)
            if start_month <= trans_month <= end_month:
                month_filtered_data.append(r)

        def create_supplier_stacked_chart(data_filtered, title):
            monthly_supplier_data = {}
            suppliers_set = set()

            for r in data_filtered:
                trans_date = r['transaction_date']
                if isinstance(trans_date, str):
                    trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
                month_year = trans_date.strftime('%Y-%m')
                supplier = r.get('supplier', 'Unknown') or 'Unknown'
                suppliers_set.add(supplier)
                key = (month_year, supplier)
                monthly_supplier_data[key] = monthly_supplier_data.get(key, 0) + abs(r['quantity'])

            plot_data = []
            if suppliers_set:
                for month_year, month_display in all_months_in_range:
                    for supplier in suppliers_set:
                        qty = monthly_supplier_data.get((month_year, supplier), 0)
                        plot_data.append({
                            'Bulan': month_display,
                            'Supplier': supplier,
                            'Jumlah': qty,
                            'sort_key': month_year
                        })
            else:
                for month_year, month_display in all_months_in_range:
                    plot_data.append({'Bulan': month_display, 'Supplier': 'No Data', 'Jumlah': 0, 'sort_key': month_year})

            if not plot_data:
                return None

            chart_df = pd.DataFrame(plot_data).sort_values('sort_key')
            month_order = [m[1] for m in all_months_in_range]

            fig = px.bar(
                chart_df, x='Bulan', y='Jumlah', color='Supplier',
                barmode='stack', title=title,
                labels={'Jumlah': 'Total Quantity', 'Bulan': 'Month'},
                text='Jumlah', color_discrete_sequence=DARK_MODE_PALETTE,
                category_orders={'Bulan': month_order}
            )
            fig.update_traces(
                texttemplate='%{text}', textposition='inside',
                textfont=dict(color='#FFFFFF', size=12, family='Arial')
            )
            fig.update_layout(
                xaxis_title='Bulan', yaxis_title='Total Quantity', height=500,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02,
                            title="Supplier", font=dict(color='#FFFFFF', size=13)),
                xaxis={'categoryorder': 'array', 'categoryarray': month_order, 'tickangle': -45},
                bargap=0.15,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#FFFFFF', size=14, family='Arial'),
                title=dict(font=dict(color='#FFFFFF', size=16)),
                xaxis_gridcolor='rgba(128,128,128,0.2)', yaxis_gridcolor='rgba(128,128,128,0.2)'
            )
            return fig

        st.markdown("#### 📥 Barang Masuk per Supplier (Bulanan)")
        masuk_chart_data = [r for r in month_filtered_data if r['action_type'] == 'MASUK']
        fig = create_supplier_stacked_chart(masuk_chart_data, 'Barang Masuk - Breakdown per Supplier')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 🚚 Pengiriman Luar Kota per Supplier (Bulanan)")
        luar_chart_data = [r for r in month_filtered_data if r['action_type'] == 'LUAR_KOTA']
        fig = create_supplier_stacked_chart(luar_chart_data, 'Luar Kota - Breakdown per Supplier')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 🛒 Penjualan Eceran per Sumber (Bulanan)")
        eceran_chart_data = [r for r in month_filtered_data if r['action_type'] == 'ECERAN']
        for r in eceran_chart_data:
            if not r.get('supplier'):
                r['supplier'] = r.get('location', 'Unknown')
        fig = create_supplier_stacked_chart(eceran_chart_data, 'Penjualan Eceran - Breakdown per Sumber')
        if fig:
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")

        # ── MONTHLY METRICS ──
        st.markdown("### 📊 Persentase Penjualan Bulanan")

        def calculate_monthly_metrics(data_filtered, metric_type="quantity"):
            monthly_data = {}

            for r in data_filtered:
                trans_date = r['transaction_date']
                if isinstance(trans_date, str):
                    trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
                month_key = trans_date.strftime('%Y-%m')
                month_name = trans_date.strftime('%B %Y')

                if month_key not in monthly_data:
                    monthly_data[month_key] = {
                        'name': month_name, 'total': 0,
                        'suppliers': {}, 'kirim_toko': {}
                    }

                supplier = r.get('supplier', 'Unknown')
                details = r.get('details', {})

                if metric_type in ("quantity", "quantity_supplier", "quantity_luar_kota"):
                    qty = abs(r['quantity'])
                    monthly_data[month_key]['total'] += qty
                    monthly_data[month_key]['suppliers'][supplier] = (
                        monthly_data[month_key]['suppliers'].get(supplier, 0) + qty
                    )
                    if metric_type == "quantity_luar_kota":
                        kirim_ke = None
                        if isinstance(details, dict):
                            kirim_ke = details.get('purpose') or details.get('kirim_ke') or details.get('tujuan')
                        if not kirim_ke or kirim_ke == '':
                            kirim_ke = 'Data Lama (Toko Tidak Tercatat)'
                        monthly_data[month_key]['kirim_toko'][kirim_ke] = (
                            monthly_data[month_key]['kirim_toko'].get(kirim_ke, 0) + qty
                        )

                elif metric_type in ("sales", "sales_supplier"):
                    sales = details.get('total_price', 0)
                    monthly_data[month_key]['total'] += sales
                    monthly_data[month_key]['suppliers'][supplier] = (
                        monthly_data[month_key]['suppliers'].get(supplier, 0) + sales
                    )

            sorted_months = sorted(monthly_data.keys())
            metrics = []

            for i, month_key in enumerate(sorted_months):
                current = monthly_data[month_key]['total']
                if i > 0:
                    prev_month_key = sorted_months[i - 1]
                    previous = monthly_data[prev_month_key]['total']
                    if previous > 0:
                        growth_pct = ((current - previous) / previous) * 100
                        growth_str = f'{growth_pct:,.1f}%'
                    else:
                        growth_str = 'n/a'
                else:
                    growth_str = 'n/a'

                prev_suppliers = monthly_data[sorted_months[i - 1]]['suppliers'] if i > 0 else {}
                supplier_breakdown = []
                for supplier, value in sorted(monthly_data[month_key]['suppliers'].items(), key=lambda x: x[1], reverse=True):
                    pct = (value / current * 100) if current > 0 else 0
                    if i == 0:
                        supplier_growth_str = 'n/a'
                    else:
                        prev_value = prev_suppliers.get(supplier, 0)
                        if prev_value > 0:
                            sg_pct = ((value - prev_value) / prev_value) * 100
                            if sg_pct > 999.9:
                                supplier_growth_str = '>999.9%'
                            elif sg_pct < -999.9:
                                supplier_growth_str = '<-999.9%'
                            else:
                                supplier_growth_str = f'{sg_pct:,.1f}%'
                        elif prev_value == 0 and value > 0:
                            supplier_growth_str = 'New'
                        else:
                            supplier_growth_str = 'n/a'
                    supplier_breakdown.append({
                        'name': supplier, 'value': value,
                        'percentage': pct, 'growth_str': supplier_growth_str
                    })

                # Kirim toko breakdown (for luar kota)
                prev_kirim_toko = monthly_data[sorted_months[i - 1]]['kirim_toko'] if i > 0 else {}
                kirim_toko_breakdown = []
                for toko, value in sorted(monthly_data[month_key]['kirim_toko'].items(), key=lambda x: x[1], reverse=True):
                    pct = (value / current * 100) if current > 0 else 0
                    if i == 0:
                        toko_growth_str = 'n/a'
                    else:
                        prev_toko_value = prev_kirim_toko.get(toko, 0)
                        if prev_toko_value > 0:
                            tg_pct = ((value - prev_toko_value) / prev_toko_value) * 100
                            if tg_pct > 999.9:
                                toko_growth_str = '>999.9%'
                            elif tg_pct < -999.9:
                                toko_growth_str = '<-999.9%'
                            else:
                                toko_growth_str = f'{tg_pct:,.1f}%'
                        elif prev_toko_value == 0 and value > 0:
                            toko_growth_str = 'New'
                        else:
                            toko_growth_str = 'n/a'
                    kirim_toko_breakdown.append({
                        'name': toko, 'value': value,
                        'percentage': pct, 'growth_str': toko_growth_str
                    })

                metrics.append({
                    'month_key': month_key,
                    'month_name': monthly_data[month_key]['name'],
                    'total': current,
                    'growth': growth_str,
                    'suppliers': supplier_breakdown,
                    'kirim_toko': kirim_toko_breakdown
                })

            return metrics

        # Barang Masuk metrics
        st.markdown("#### 📥 Barang Masuk (Bulanan)")
        if masuk_data:
            masuk_metrics = calculate_monthly_metrics(masuk_data, "quantity")
            for metric in masuk_metrics[-12:]:
                with st.expander(f"**{metric['month_name']}**", expanded=False):
                    st.metric("📊 Total Unit", f"{metric['total']:,} unit", delta=metric['growth'])
                    st.markdown("---")
                    st.markdown("**📦 Breakdown per Supplier:**")
                    num_s = len(metric['suppliers'])
                    if num_s > 0:
                        cols = st.columns(min(num_s, 4))
                        for idx, si in enumerate(metric['suppliers']):
                            cols[idx % min(num_s, 4)].metric(si['name'], f"{si['value']:,} unit", delta=si['growth_str'])
        else:
            st.info("Tidak ada data")

        st.markdown("---")

        # Luar Kota metrics
        st.markdown("#### 🚚 Luar Kota (Bulanan)")
        if luar_kota_data:
            luar_metrics = calculate_monthly_metrics(luar_kota_data, "quantity_luar_kota")
            for metric in luar_metrics[-12:]:
                with st.expander(f"**{metric['month_name']}**", expanded=False):
                    st.metric("📊 Total Unit", f"{metric['total']:,} unit", delta=metric['growth'])
                    st.markdown("---")
                    st.markdown("**📦 Per Supplier:**")
                    num_s = len(metric['suppliers'])
                    if num_s > 0:
                        cols = st.columns(min(num_s, 4))
                        for idx, si in enumerate(metric['suppliers']):
                            cols[idx % min(num_s, 4)].metric(si['name'], f"{si['value']:,} unit", delta=si['growth_str'])
                    st.markdown("---")
                    st.markdown("**🏪 Kirim ke Toko:**")
                    num_t = len(metric['kirim_toko'])
                    if num_t > 0:
                        cols = st.columns(min(num_t, 4))
                        for idx, ti in enumerate(metric['kirim_toko']):
                            cols[idx % min(num_t, 4)].metric(ti['name'], f"{ti['value']:,} unit", delta=ti['growth_str'])
        else:
            st.info("Tidak ada data")

        st.markdown("---")

        # Eceran quantity metrics
        st.markdown("#### 🛒 Eceran - Jumlah Barang (Bulanan)")
        if eceran_data:
            eceran_qty_metrics = calculate_monthly_metrics(eceran_data, "quantity_supplier")
            for metric in eceran_qty_metrics[-12:]:
                with st.expander(f"**{metric['month_name']}**", expanded=False):
                    st.metric("📊 Total Unit", f"{metric['total']:,} unit", delta=metric['growth'])
                    st.markdown("---")
                    st.markdown("**📦 Per Supplier:**")
                    num_s = len(metric['suppliers'])
                    if num_s > 0:
                        cols = st.columns(min(num_s, 4))
                        for idx, si in enumerate(metric['suppliers']):
                            cols[idx % min(num_s, 4)].metric(si['name'], f"{si['value']:,} unit", delta=si['growth_str'])
        else:
            st.info("Tidak ada data")

        st.markdown("---")

        # Eceran sales metrics
        st.markdown("#### 💰 Eceran - Total Penjualan (Bulanan)")
        if eceran_data:
            eceran_sales_metrics = calculate_monthly_metrics(eceran_data, "sales_supplier")
            for metric in eceran_sales_metrics[-12:]:
                with st.expander(f"**{metric['month_name']}**", expanded=False):
                    st.metric("💰 Total Penjualan", format_currency(metric['total']), delta=metric['growth'])
                    st.markdown("---")
                    st.markdown("**📦 Per Supplier:**")
                    num_s = len(metric['suppliers'])
                    if num_s > 0:
                        cols = st.columns(min(num_s, 4))
                        for idx, si in enumerate(metric['suppliers']):
                            cols[idx % min(num_s, 4)].metric(si['name'], format_currency(si['value']), delta=si['growth_str'])
        else:
            st.info("Tidak ada data")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
