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

# ===== ADD THIS FUNCTION HERE =====
def normalize_for_comparison(text: str) -> str:
    """
    Enhanced normalization for strict duplicate detection.
    Handles special characters, spacing, and case.
    
    Args:
        text: Input text
    
    Returns:
        str: Normalized text for comparison
    """
    if not text or not isinstance(text, str):
        return ""
    
    # Convert to lowercase
    normalized = text.lower()
    
    # Remove all whitespace (including tabs, newlines)
    normalized = ''.join(normalized.split())
    
    # Remove special characters and punctuation
    special_chars = ['(', ')', '-', '/', ',', '.', ':', ';', '_', '[', ']', '{', '}', '"', "'"]
    for char in special_chars:
        normalized = normalized.replace(char, '')
    
    # Remove any remaining non-alphanumeric characters
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
    """Inject custom CSS for validation styling"""
    st.markdown("""
    <style>
    /* Valid field styling */
    .valid-field input,
    .valid-field select {
        border: 2px solid #10b981 !important;
        background-color: rgba(16, 185, 129, 0.05) !important;
    }
    
    /* Invalid field styling */
    .invalid-field input,
    .invalid-field select {
        border: 2px solid #ef4444 !important;
        background-color: rgba(239, 68, 68, 0.05) !important;
    }
    
    /* Error message styling */
    .field-error {
        color: #ef4444;
        font-size: 0.85rem;
        margin-top: -0.5rem;
        margin-bottom: 0.5rem;
        font-weight: 500;
    }
    
    /* Item validation summary */
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
    
    /* Optional text styling */
    .optional-text {
        color: rgba(128, 128, 128, 0.6);
        font-weight: normal;
        font-size: 0.9em;
    }
    </style>
    """, unsafe_allow_html=True)

def inject_aggrid_merge_css():
    """Make AgGrid cells with arrows appear completely merged - ENHANCED VERSION"""
    st.markdown("""
    <style>
    /* Hide cells that contain only the arrow symbol */
    .ag-cell:has-text("↑") {
        color: transparent !important;
        border-top: none !important;
        border-bottom: none !important;
        background-color: inherit !important;
    }
    
    /* Alternative approach: target cells with specific content */
    .ag-cell {
        position: relative;
    }
    
    /* Make arrow cells invisible and remove borders */
    div.ag-cell[data-value="↑"] {
        color: transparent !important;
        visibility: hidden !important;
        border: none !important;
    }
    
    /* Remove all borders for arrow cells */
    .ag-theme-streamlit .ag-cell,
    .ag-theme-alpine .ag-cell,
    .ag-theme-balham .ag-cell {
        border-color: #4a5568;
    }
    
    /* Specific styling for merged appearance */
    .ag-row .ag-cell {
        border-right: 1px solid #4a5568;
    }
    
    /* Hide the actual arrow content */
    span.ag-cell-value {
        display: inline-block;
    }
    
    /* Target cells with arrow and make them seamless */
    .ag-cell-wrapper > span.ag-cell-value:contains("↑") {
        display: none !important;
    }
    </style>
    
    <script>
    // JavaScript to enhance cell merging after AgGrid renders
    setTimeout(function() {
        // Find all cells
        const allCells = document.querySelectorAll('.ag-cell');
        
        allCells.forEach(cell => {
            const cellValue = cell.textContent.trim();
            
            // If cell contains only arrow, make it invisible and borderless
            if (cellValue === '↑' || cellValue === '') {
                cell.style.color = 'transparent';
                cell.style.borderTop = 'none';
                cell.style.borderBottom = 'none';
                cell.style.visibility = 'hidden';
                
                // Get the cell above to extend its border
                const cellRect = cell.getBoundingClientRect();
                const cellAbove = document.elementFromPoint(
                    cellRect.left + cellRect.width / 2,
                    cellRect.top - 5
                );
                
                if (cellAbove && cellAbove.classList.contains('ag-cell')) {
                    cellAbove.style.borderBottom = 'none';
                }
            }
        });
    }, 100);
    
    // Re-run on scroll to handle virtual scrolling
    const gridContainer = document.querySelector('.ag-body-viewport');
    if (gridContainer) {
        gridContainer.addEventListener('scroll', function() {
            setTimeout(function() {
                const allCells = document.querySelectorAll('.ag-cell');
                allCells.forEach(cell => {
                    const cellValue = cell.textContent.trim();
                    if (cellValue === '↑' || cellValue === '') {
                        cell.style.color = 'transparent';
                        cell.style.borderTop = 'none';
                        cell.style.borderBottom = 'none';
                        cell.style.visibility = 'hidden';
                    }
                });
            }, 50);
        });
    }
    </script>
    """, unsafe_allow_html=True)
    
def validate_field(value: Any, field_type: str = "text", required: bool = True) -> Tuple[bool, str]:
    """
    Validate a single field
    Returns: (is_valid, error_message)
    """
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
    """Display validation feedback for a field"""
    if not is_valid and error_msg:
        st.markdown(f'<div class="field-error">⚠️ {error_msg}</div>', unsafe_allow_html=True)

# ==================================================
# CACHED DATA LOADING FUNCTIONS (Performance Boost)
# ==================================================

@st.cache_data(ttl=300, show_spinner="Loading inventory...")
def get_cached_inventory():
    """Cache inventory for 5 minutes"""
    return get_all_inventory()

@st.cache_data(ttl=300, show_spinner="Loading history...")
def get_cached_history(start_date=None, end_date=None, item_name=None):
    """Cache history for 5 minutes"""
    return get_history(start_date=start_date, end_date=end_date, item_name=item_name)

def invalidate_caches():
    """Clear all caches when data changes"""
    st.cache_data.clear()
    st.session_state.force_refresh_dashboard = True
    if 'dashboard_last_picked_cache' in st.session_state:
        del st.session_state['dashboard_last_picked_cache']

# ==================================================
# COMMON CONSTANTS
# ==================================================

DATE_FORMAT = "%d/%m/%y"
VALIDATION_MESSAGES = {
    'empty_field': "⚠️ Field ini wajib diisi",
    'duplicate': "⚠️ '{value}' sudah ada dalam daftar",
    'invalid_quantity': "⚠️ Jumlah harus lebih besar dari 0",
    'invalid_price': "⚠️ Harga harus lebih besar dari 0"
}

# Helper function to render collapsible instruction box
def render_instruction_box():
    """Render a collapsible instruction box - clean version"""
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
        - ✅ Harap **cari pilihan** sebelum pilih "Tambah Baru" untuk menstandarisasi data
        - ✅ **Ukuran HARAP diisi jika ada** untuk semua jenis transaksi (Barang Masuk, Luar Kota, Eceran)
        - ✅ **HARAP tulis nama panjang daripada nama singkat "(contoh b/g -> bantal guling, dsb)"
        - ✅ **Warna dan Ukuran boleh kosong**, akan otomatis terisi "Tidak Ada"
        - 🏭 Jika lokasi dipilih **Pabrik**, stok TIDAK tercatat sebagai stok aktif (dianggap langsung ke konsumen/luar kota)
        - 📦 Jika barang dari pabrik lalu **disimpan di toko/gudang**, pilih lokasi penyimpanan yang sesuai
        - 🎁 Untuk Bantal Guling (b/g), Guling (g), Bantal (b), Matras Protector (Mp) → pilih kategori **Bonus**
          - Eceran: jika bonus → harga = 0 | jika diperjualkan → harga > 0
        """)

# ==================================================
# UTILITY FUNCTIONS
# ==================================================

def convert_to_local_time(dt):
    """Convert UTC datetime to local timezone"""
    if dt is None:
        return None
    
    if isinstance(dt, date) and not isinstance(dt, datetime):
        dt = datetime.combine(dt, datetime.min.time())
    
    if isinstance(dt, str):
        try:
            dt = datetime.strptime(dt[:26], '%Y-%m-%d %H:%M:%S.%f')
        except:
            try:
                dt = datetime.strptime(dt[:19], '%Y-%m-%d %H:%M:%S')
            except:
                return None
    
    if dt.tzinfo is None:
        dt = pytz.UTC.localize(dt)
    
    local_dt = dt.astimezone(LOCAL_TIMEZONE)
    return local_dt

def format_local_datetime(dt, format_str="%d/%m/%y %H:%M"):
    """Format datetime in local timezone"""
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
    else:
        return [item_data]

def format_date_dd_mm_yy(date_obj: Any) -> str:
    """Updated to use local timezone"""
    try:
        if isinstance(date_obj, (date, datetime)):
            if isinstance(date_obj, datetime):
                local_dt = convert_to_local_time(date_obj)
                return local_dt.strftime("%d/%m/%y") if local_dt else "—"
            else:
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
    """Format combined item name with size and color"""
    name = name.title() if name else "—"
    size = size.strip() if size and size.lower() != "tidak ada" else ""
    color = color.strip() if color and color.lower() != "tidak ada" else ""
    
    if size and color:
        return f"{name} ({size}) - {color.title()}"
    elif size:
        return f"{name} ({size})"
    elif color:
        return f"{name} - {color.title()}"
    else:
        return name

def select_or_text(
    label: str, 
    options: List[str], 
    key_prefix: str, 
    default: str = "",
    manual_label: str = "Lainnya (ketik manual)", 
    container = None, 
    session_values_key: Optional[str] = None,
    required: bool = False
) -> str:
    """
    Enhanced autocomplete with comprehensive cross-field validation.
    Blocks inappropriate data entry across ALL fields.
    """
    if container is None: 
        container = st
    
    from utils.validators import find_similar_entries, comprehensive_field_validation, get_all_reference_data_for_validation
    
    # --------------- base options -------------------------------
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
    
    normalized_lookup = {
        normalize_for_comparison(opt): opt 
        for opt in base_options
    }
    
    # --------------- session-state keys --------------------------
    value_key   = f"{key_prefix}_value"
    confirm_key = f"{key_prefix}_confirm_add"
    ver_key     = f"{key_prefix}_ver"

    if value_key not in st.session_state:
        st.session_state[value_key] = default
    if confirm_key not in st.session_state:
        st.session_state[confirm_key] = False
    if ver_key not in st.session_state:
        st.session_state[ver_key] = 0

    input_key = f"{key_prefix}_input_v{st.session_state[ver_key]}"

    # --------------- label ----------------------------------------
    if required:
        container.markdown(f"**{label}** *", unsafe_allow_html=True)
    else:
        container.markdown(
            f"**{label}** <span class='optional-text'>(optional)</span>", 
            unsafe_allow_html=True
        )
    
    # --------------- text input ------------------------------------
    user_input = container.text_input(
        f"{label}_input",
        value=st.session_state[value_key],
        key=input_key,
        placeholder="Ketik untuk mencari atau pilih dari saran...",
        label_visibility="collapsed"
    )
    
    # --------------- sync: typing vs button-pick ------------------
    if user_input != st.session_state[value_key]:
        st.session_state[value_key] = user_input
        st.session_state[confirm_key] = False

    # --------------- helper picked by every ✓ button --------------
    def _pick(chosen: str):
        st.session_state[value_key]   = chosen
        st.session_state[ver_key]    += 1
        st.session_state[confirm_key] = False
        st.rerun()

    # --------------- suggestion UI (only when text is non-empty) --
    if user_input and user_input.strip():
        search_term      = user_input.strip().lower()
        normalized_input = normalize_for_comparison(user_input)
        
        # --- exact-match check ---
        if normalized_input in normalized_lookup:
            existing_match = normalized_lookup[normalized_input]

            if user_input.strip() == existing_match:
                return existing_match

            container.info(
                f"ℹ️ '{user_input}' sama dengan '{existing_match}' yang sudah ada"
            )
            
            btn_key_use = f"{key_prefix}_use_existing_{hash(existing_match) % 100000}_{st.session_state[ver_key]}"
            
            if container.button(
                f"✓ Gunakan '{existing_match}'",
                key=btn_key_use,
                use_container_width=True
            ):
                _pick(existing_match)
            
            return existing_match
        
        # --- fuzzy / phonetic similar entries -----------------------
        similar_entries = find_similar_entries(
            user_input, 
            base_options, 
            levenshtein_threshold=3,
            check_phonetic=True
        )
        
        if similar_entries:
            critical_matches = [
                entry for entry in similar_entries 
                if entry[1] <= 2 or entry[2]
            ]
            
            if critical_matches:
                container.warning(
                    f"⚠️ **'{user_input}' mirip dengan data yang sudah ada:**"
                )
                
                for idx, (entry, distance, is_phonetic) in enumerate(critical_matches[:5]):
                    match_type = "🔊 Bunyi sama" if is_phonetic else f"📏 Jarak: {distance}"
                    
                    col1, col2 = container.columns([3, 1])
                    col1.markdown(
                        f"<div style='background: rgba(255, 165, 0, 0.15); padding: 0.5rem; border-radius: 4px; margin: 0.25rem 0;'>"
                        f"<strong>'{entry}'</strong> → {match_type}"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                    
                    btn_key_quick = f"{key_prefix}_quick_{idx}_{hash(entry) % 100000}_{st.session_state[ver_key]}"
                    
                    if col2.button(
                        "✓ Pilih",
                        key=btn_key_quick,
                        use_container_width=True
                    ):
                        _pick(entry)
                
                container.info(
                    f"💡 **Saran:** Klik '✓ Pilih' untuk menggunakan pilihan yang ada, "
                    f"atau lanjut ke bawah jika '{user_input}' memang berbeda."
                )
        
        # --- substring matching ------------------------------------
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
                        
                        if col.button(
                            f"✓ {suggestion}", 
                            key=btn_key_select,
                            use_container_width=True
                        ):
                            _pick(suggestion)
            
            container.markdown("</div>", unsafe_allow_html=True)

        # ========================================================
        # ADD-NEW FLOW WITH COMPREHENSIVE VALIDATION
        # ========================================================
        if not matching_options or not any(opt.lower() == search_term for opt in matching_options):
            container.markdown("---")
            
            # **COMPREHENSIVE CROSS-FIELD VALIDATION**
            should_block_add = False
            block_reason = ""
            
            # Map session_values_key to field type
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
                # Get all reference data
                all_reference_data = get_all_reference_data_for_validation()
                
                # Run comprehensive validation
                typed_std = standardize_display(user_input)
                
                # **FIX: Extract item number from key_prefix**
                item_num_match = re.search(r'_(\d+)$', key_prefix)
                item_num = item_num_match.group(1) if item_num_match else "1"
                
                # **NEW: Get the form prefix (add, luar, ec)**
                form_prefix_match = re.match(r'^([a-z]+)_', key_prefix)
                form_prefix = form_prefix_match.group(1) if form_prefix_match else "add"
                
                should_block, error_msg = comprehensive_field_validation(
                    field_type=field_type,
                    user_input=typed_std,
                    item_num=item_num,
                    all_reference_data=all_reference_data
                )
                
                if should_block:
                    should_block_add = True
                    block_reason = error_msg
                
                # **KATEGORI IN BARANG CHECK**
                if field_type == 'barang' and not should_block:
                    all_kategori = all_reference_data.get('kategori', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for kat in all_kategori:
                        kat_normalized = normalize_for_comparison(kat)
                        if typed_normalized == kat_normalized or kat_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Kategori**, bukan Nama Barang!\n\n"
                                f"💡 **Anda mencoba memasukkan '{kat}' (Kategori) ke field Nama Barang**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Kategori: `{kat}` → Nama Barang: `Emily`, `Kingbreeze`, dsb\n\n"
                                f"❌ **Salah:**\n"
                                f"• Kategori: `{kat}` → Nama Barang: `{kat}` atau `{typed_std}` ❌"
                            )
                            break
                
                # **BARANG IN WRONG FIELD CHECK**
                if field_type in ['ukuran', 'warna', 'via', 'toko'] and not should_block:
                    all_barang = all_reference_data.get('barang', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for barang in all_barang:
                        barang_normalized = normalize_for_comparison(barang)
                        if typed_normalized == barang_normalized or barang_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Barang**, bukan {field_type.title()}!\n\n"
                                f"💡 **Anda mencoba memasukkan '{barang}' (Nama Barang) ke field {field_type.title()}**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Nama Barang: `{barang}` → {field_type.title()}: `(nilai yang sesuai)`\n\n"
                                f"❌ **Salah:**\n"
                                f"• {field_type.title()}: `{typed_std}` ❌"
                            )
                            break
                
                # **KATEGORI IN WRONG FIELD CHECK**
                if field_type in ['ukuran', 'warna', 'supplier', 'via', 'toko'] and not should_block:
                    all_kategori = all_reference_data.get('kategori', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for kat in all_kategori:
                        kat_normalized = normalize_for_comparison(kat)
                        if typed_normalized == kat_normalized or kat_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Kategori**, bukan {field_type.title()}!\n\n"
                                f"💡 **Anda mencoba memasukkan '{kat}' (Kategori) ke field {field_type.title()}**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Kategori: `{kat}` → {field_type.title()}: `(nilai yang sesuai)`\n\n"
                                f"❌ **Salah:**\n"
                                f"• {field_type.title()}: `{typed_std}` ❌"
                            )
                            break
                
                # **WARNA IN UKURAN CHECK**
                if field_type == 'ukuran' and not should_block:
                    all_warna = all_reference_data.get('warna', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for warna in all_warna:
                        warna_normalized = normalize_for_comparison(warna)
                        if typed_normalized == warna_normalized or warna_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Warna**, bukan Ukuran!\n\n"
                                f"💡 **Anda mencoba memasukkan '{warna}' (Warna) ke field Ukuran**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Warna: `{warna}` | Ukuran: `3k`, `6k`, `180x200`, dll\n\n"
                                f"❌ **Salah:**\n"
                                f"• Ukuran: `{typed_std}` ❌"
                            )
                            break
                
                # **UKURAN IN WARNA CHECK**
                if field_type == 'warna' and not should_block:
                    all_ukuran = all_reference_data.get('ukuran', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for ukuran in all_ukuran:
                        ukuran_normalized = normalize_for_comparison(ukuran)
                        if typed_normalized == ukuran_normalized or ukuran_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Ukuran**, bukan Warna!\n\n"
                                f"💡 **Anda mencoba memasukkan '{ukuran}' (Ukuran) ke field Warna**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Ukuran: `{ukuran}` | Warna: `Merah`, `Abu-abu`, dll\n\n"
                                f"❌ **Salah:**\n"
                                f"• Warna: `{typed_std}` ❌"
                            )
                            break
                
                # **NEW: SUPPLIER IN VIA/TOKO CHECK**
                if field_type in ['via', 'toko'] and not should_block:
                    all_supplier = all_reference_data.get('supplier', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for supplier in all_supplier:
                        supplier_normalized = normalize_for_comparison(supplier)
                        if typed_normalized == supplier_normalized or supplier_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Supplier**, bukan {field_type.title()}!\n\n"
                                f"💡 **Anda mencoba memasukkan '{supplier}' (Supplier) ke field {field_type.title()}**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Supplier: `{supplier}` → {field_type.title()}: `(nilai yang sesuai)`\n\n"
                                f"❌ **Salah:**\n"
                                f"• {field_type.title()}: `{typed_std}` ❌"
                            )
                            break
                
                # **NEW: VIA IN TOKO CHECK**
                if field_type == 'toko' and not should_block:
                    all_via = all_reference_data.get('via', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for via in all_via:
                        via_normalized = normalize_for_comparison(via)
                        if typed_normalized == via_normalized or via_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah **Via Pengiriman**, bukan Kirim Ke Toko!\n\n"
                                f"💡 **Anda mencoba memasukkan '{via}' (Via Pengiriman) ke field Kirim Ke Toko**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Via Pengiriman: `{via}` | Kirim Ke Toko: `Toko A`, `Toko B`, dll\n\n"
                                f"❌ **Salah:**\n"
                                f"• Kirim Ke Toko: `{typed_std}` ❌"
                            )
                            break
                
                # **NEW: TOKO IN VIA CHECK**
                if field_type == 'via' and not should_block:
                    all_toko = all_reference_data.get('toko', [])
                    typed_normalized = normalize_for_comparison(typed_std)
                    
                    for toko in all_toko:
                        toko_normalized = normalize_for_comparison(toko)
                        if typed_normalized == toko_normalized or toko_normalized in typed_normalized:
                            should_block_add = True
                            block_reason = (
                                f"🚫 **'{typed_std}'** adalah nama **Toko**, bukan Via Pengiriman!\n\n"
                                f"💡 **Anda mencoba memasukkan '{toko}' (Nama Toko) ke field Via Pengiriman**\n\n"
                                f"✅ **Yang benar:**\n"
                                f"• Kirim Ke Toko: `{toko}` | Via Pengiriman: `JNE`, `Grab`, dll\n\n"
                                f"❌ **Salah:**\n"
                                f"• Via Pengiriman: `{typed_std}` ❌"
                            )
                            break
            
            if should_block_add:
                # **BLOCK: Show error, DON'T show "add" button**
                container.error(f"🚫 {block_reason}")
                container.warning(
                    "❌ **DATA DITOLAK!**\n\n"
                    "💡 **Pastikan Anda memasukkan data di field yang benar:**\n"
                    "• **Kategori** → Jenis barang (Matras, Bed Dorong, dll)\n"
                    "• **Nama Barang** → Nama spesifik (Emily, Kingbreeze, dll)\n"
                    "• **Ukuran** → Dimensi (3k, 6k, 180x200, dll)\n"
                    "• **Warna** → Warna produk (Merah, Abu-abu, dll)\n"
                    "• **Supplier** → Nama supplier\n"
                    "• **Via Pengiriman** → Metode kirim (JNE, Grab, dll)\n"
                    "• **Kirim ke Toko** → Nama toko tujuan"
                )
                # Reset confirm state
                st.session_state[confirm_key] = False
            else:
                # **ALLOW: Show normal "add new" flow**
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
                        container.warning(
                            f"⚠️ '{typed_std}' sama dengan '{existing}' (mengabaikan huruf besar/kecil)"
                        )
                        _pick(existing)
                    else:
                        # All validations passed - proceed to save
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
    if not records: return pd.DataFrame()
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
# DASHBOARD TAB (Keep original - no changes needed)
# ==================================================

# REPLACE the render_dashboard_tab function starting around line 550
# This fixes the stock warning cards to show "Nama Barang (Ukuran) - Warna" format

def render_dashboard_tab(today: date):
    """Dashboard with "Terakhir Input" labels showing timestamp of last form submission"""
    
    cache_key = 'dashboard_last_picked_cache'
    force_refresh_key = 'force_refresh_dashboard'
    
    if force_refresh_key not in st.session_state: 
        st.session_state[force_refresh_key] = False
    
    rebuild_cache = (cache_key not in st.session_state or 
                     st.session_state.get(force_refresh_key, False) or
                     not isinstance(st.session_state.get(cache_key), dict))
    
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
                        latest_in_category = max(category_records, 
                                                key=lambda r: convert_to_local_time(r.get('created_at')) or datetime.min.replace(tzinfo=pytz.UTC))
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
                    category_map = {
                        'MASUK': 'Barang Masuk',
                        'LUAR_KOTA': 'Luar Kota',
                        'ECERAN': 'Eceran'
                    }
                    last_category = category_map.get(action_type, action_type)
                    trans_date = latest_overall.get('transaction_date')
                    last_trans_date = format_local_datetime(trans_date, "%d/%m/%Y")
                    
                    # **FIX: Use format_item_combined for last picked item**
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
            f"</div>", 
            unsafe_allow_html=True
        )
    
    with col2:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px; font-size:0.75rem;'>TERAKHIR INPUT</small>"
            f"<h3 style='margin:8px 0; font-size:1rem; color:#F59E0B;'>🚚 Luar Kota</h3>"
            f"<p style='margin:5px 0; font-size:0.95rem; font-weight:600;'>{luar_time}</p>"
            f"</div>", 
            unsafe_allow_html=True
        )
    
    with col3:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px; font-size:0.75rem;'>TERAKHIR INPUT</small>"
            f"<h3 style='margin:8px 0; font-size:1rem; color:#10B981;'>🛒 Eceran</h3>"
            f"<p style='margin:5px 0; font-size:0.95rem; font-weight:600;'>{eceran_time}</p>"
            f"</div>", 
            unsafe_allow_html=True
        )
    
    with col4:
        st.markdown(
            f"<div class='metric-box' style='min-height:160px;'>"
            f"<small style='color:gray; letter-spacing:1px;'>BARANG TERAKHIR DIUPDATE</small>"
            f"<h2 style='margin:10px 0 0 0; font-size:1.3rem;'>{last_picked}</h2>"
            f"<p style='color:#888; margin:5px 0 0 0; font-size:0.9rem;'>Kategori: {last_category}</p>"
            f"</div>", 
            unsafe_allow_html=True
        )

    st.markdown("---")
    
    try:
        all_inventory = get_cached_inventory()
        inventory = [item for item in all_inventory if not is_pabrik_source(item.get('location', ''))]
        
        if not inventory:
            st.info("ℹ️ Tidak ada stok aktif. Semua barang dari Pabrik (tidak masuk stok aktif).")
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
        
        # **CRITICAL FIX: Stock warning cards now use format_item_combined**
        if critical_stock or low_stock:
            with st.expander("⚠️ Peringatan Stok", expanded=True):
                for item in (critical_stock + low_stock)[:10]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    badge_color = "#dc2626" if qty == 0 else "#d97706"
                    border_color = badge_color
                    badge_text = "HABIS" if qty == 0 else "RENDAH"
                    
                    # **FIX: Use format_item_combined instead of just .title()**
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'),
                        item.get('size', ''),
                        item.get('color', '')
                    )
                    
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid {border_color};'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:{badge_color}; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>{badge_text}</div>"
                        f"</div></div>", 
                        unsafe_allow_html=True
                    )
        
        if medium_stock:
            with st.expander("📦 Barang Stock Cukup (4-10)", expanded=False):
                for item in medium_stock[:15]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    badge_color = "#2563eb"
                    border_color = badge_color
                    
                    # **FIX: Use format_item_combined**
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'),
                        item.get('size', ''),
                        item.get('color', '')
                    )
                    
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid {border_color};'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:{badge_color}; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>CUKUP</div>"
                        f"</div></div>", 
                        unsafe_allow_html=True
                    )
        
        if high_stock:
            with st.expander("✅ Barang Stock Banyak (>10)", expanded=False):
                for item in high_stock[:15]:
                    qty = item.get('quantity', 0)
                    location = item.get('location', '-')
                    badge_color = "#059669"
                    border_color = badge_color
                    
                    # **FIX: Use format_item_combined**
                    item_display = format_item_combined(
                        item.get('name', 'Unknown'),
                        item.get('size', ''),
                        item.get('color', '')
                    )
                    
                    st.markdown(
                        f"<div style='background:#1a1a1a; padding:1rem; margin:0.5rem 0; "
                        f"border-radius:8px; border-left:4px solid {border_color};'>"
                        f"<div style='display:flex; justify-content:space-between;'>"
                        f"<div><strong>{item_display}</strong>"
                        f"<small style='display:block; color:#888;'>Qty: {qty} | Lokasi: {location}</small></div>"
                        f"<div style='background:{badge_color}; color:white; padding:0.25rem 0.75rem; "
                        f"border-radius:12px; font-size:0.85rem; font-weight:600;'>BANYAK</div>"
                        f"</div></div>", 
                        unsafe_allow_html=True
                    )
        
        if not critical_stock and not low_stock:
            st.success("✅ Semua stok dalam kondisi baik")
            
    except Exception as e: 
        st.error(f"❌ Error loading dashboard: {str(e)}")
        
# ==================================================
# BARANG MASUK TAB - WITH IMPROVED VALIDATION
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
        options=["Iya", "Tidak"],
        key="masuk_same_loc",
        horizontal=True
    )
    
    shared_location = None
    if same_location == "Iya":
        shared_location = st.selectbox("Lokasi untuk Semua Barang *", STANDARD_LOCATIONS, key="masuk_shared_loc")
        if shared_location == "Pabrik": 
            st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak masuk Stok Aktif")
    
    st.markdown("---")
    st.markdown("### 📦 Daftar Barang")
    
    # Initialize item count in session state
    if 'masuk_item_count' not in st.session_state:
        st.session_state.masuk_item_count = 1
    
    items_to_add = []
    item_validations = []
    
    # Render items based on current count
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
                if loc == "Pabrik": st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak masuk Stok Aktif")
            else:
                loc = shared_location
            
            komen = st.text_input("Komen", placeholder="Optional", key=f"add_komen_{i}").strip()
            
            # REAL-TIME VALIDATION FEEDBACK
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
            
            # SEMANTIC OVERLAP CHECK
            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ {overlap_msg}'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, [overlap_msg]))
                    continue  # **CRITICAL: Skip this item**
            
            if cat or name or size or qty > 0:  # Only show validation if user started filling
                if missing_fields or validation_errors:
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    if validation_errors:
                        error_display.extend(validation_errors)
                    
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Item {i}: <strong>{" | ".join(error_display)}</strong>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">'
                        f'✅ Item {i}: Lengkap dan siap disimpan'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, True, []))
                    
                    items_to_add.append({
                        "name": name, "category": cat, "color": color or "Tidak Ada", 
                        "size": size, "quantity": qty, "location": loc, "supplier": supplier, 
                        "nomor_po": nomor_po or None, 
                        "details": {"komen": komen if komen else None}
                    })
    
    # ========================================
    # ADD/DELETE ITEM BUTTONS
    # ========================================
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
    
    # PREVIEW SUMMARY
    if items_to_add:
        expanded_preview = []
        for item in items_to_add: 
            expanded_preview.extend(expand_combo_item(item))
        st.success(f"✅ {len(expanded_preview)} item siap disimpan")
        col1, col2 = st.columns(2)
        col1.metric("Total Items", len(expanded_preview))
        col2.metric("Total Quantity", sum(item['quantity'] for item in expanded_preview))
    
    # SUBMIT BUTTON WITH COMPREHENSIVE VALIDATION
    if st.button("💾 Simpan Barang Masuk", type="primary", key="save_masuk_btn"):
        validation_errors = []
        
        # Validate Nomor PO
        if not nomor_po:
            validation_errors.append("❌ **Nomor PO** wajib diisi!")
        
        # Validate Supplier
        if not supplier:
            validation_errors.append("❌ **Supplier** wajib diisi!")
        
        # Validate items
        if not items_to_add:
            validation_errors.append("❌ **Belum ada item yang lengkap**. Pastikan setiap item memiliki: Kategori, Nama Barang, Ukuran, dan Jumlah > 0")
        
        # Show specific item errors
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, errors_list in incomplete_items:
                if errors_list:
                    validation_errors.append(f"❌ **Item {item_num}**: {', '.join(str(e) for e in errors_list)}")
        
        if validation_errors:
            st.error("### ⚠️ Validasi Gagal - Perbaiki Error Berikut:")
            for error in validation_errors:
                st.markdown(error)
            return
        
        # All validations passed - proceed with save
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
            
            # Reset item count
            st.session_state.masuk_item_count = 1
            
            # Force form refresh by clearing all form-related session state
            keys_to_clear = [k for k in st.session_state.keys() if k.startswith('masuk_') or k.startswith('add_')]
            for key in keys_to_clear:
                del st.session_state[key]
            
            time.sleep(1)
            st.rerun()

# ==================================================
# LUAR KOTA TAB - WITH IMPROVED VALIDATION
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
        options=["Iya", "Tidak"],
        key="luar_same_src",
        horizontal=True
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
    
    # Initialize item count for Luar Kota
    if 'luar_item_count' not in st.session_state:
        st.session_state.luar_item_count = 1  # Default 5 items
    
    for i in range(1, st.session_state.luar_item_count + 1):
        with st.expander(f"Item {i}", expanded=(i == st.session_state.luar_item_count)):  # ← SESUDAHNYA
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
                if loc == "Pabrik": st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak kurangi Stok Aktif")
            else:
                loc = shared_source
            
            komen = st.text_input("Komen", placeholder="Optional", key=f"luar_komen_{i}").strip()
            
            # REAL-TIME VALIDATION
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
            
            # SEMANTIC OVERLAP CHECK
            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ {overlap_msg}'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, [overlap_msg]))
                    continue  # **CRITICAL: Skip this item**
            
            if cat or name or sz or supplier or qty > 0:
                if missing_fields or validation_errors:
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    if validation_errors:
                        error_display.extend(validation_errors)
                    
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Item {i}: <strong>{" | ".join(error_display)}</strong>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">'
                        f'✅ Item {i}: Lengkap dan siap dikirim'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, True, []))
                    luar_items.append({
                        "name": name, "category": cat, "color": colr or "Tidak Ada", 
                        "size": sz, "quantity": qty, "source": loc, "location": loc, 
                        "supplier": supplier, 
                        "details": {"komen": komen if komen else None, "purpose": purpose, "via": via}
                    })
    
    # ========================================
    # ADD/DELETE ITEM BUTTONS - LUAR KOTA
    # ========================================
    col_add_luar, col_delete_luar = st.columns([3, 1])
    
    with col_add_luar:
        # Initialize luar_item_count if not exists
        if 'luar_item_count' not in st.session_state:
            st.session_state.luar_item_count = MAX_ITEMS_PER_FORM
        
        if st.session_state.luar_item_count < MAX_ITEMS_PER_FORM:
            if st.button("➕ Tambah Item Baru", key="add_luar_item_btn", use_container_width=True):
                st.session_state.luar_item_count += 1
                st.rerun()
        else:
            st.info(f"ℹ️ Maksimal {MAX_ITEMS_PER_FORM} item per form")
    
    with col_delete_luar:
        if 'luar_item_count' not in st.session_state:
            st.session_state.luar_item_count = MAX_ITEMS_PER_FORM
        
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
            validation_errors.append("❌ **Belum ada item yang lengkap**. Pastikan setiap item memiliki: Kategori, Nama Barang, Ukuran, Supplier, dan Qty > 0")
        
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, missing in incomplete_items:
                validation_errors.append(f"❌ **Item {item_num}**: Belum diisi - {', '.join(missing)}")
        
        if validation_errors:
            st.error("### ⚠️ Validasi Gagal - Perbaiki Error Berikut:")
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
                    except: 
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
            
            # Force form refresh
            keys_to_clear = [k for k in st.session_state.keys() if k.startswith('luar_')]
            for key in keys_to_clear:
                del st.session_state[key]
            
            time.sleep(1)
            st.rerun()
# ==================================================
# ECERAN TAB - WITH IMPROVED VALIDATION
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
        options=["Iya", "Tidak"],
        key="ecer_same_src",
        horizontal=True
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
            st.info("ℹ️ Dari Pabrik: Dicatat dalam Riwayat, tidak kurangi Stok Aktif")
    
    st.markdown("---")
    st.markdown("#### 🛍️ Barang Dijual")
    
    # Initialize item count for Eceran
    if 'ecer_item_count' not in st.session_state:
        st.session_state.ecer_item_count = 1  # Default 5 items
    
    for i in range(1, st.session_state.ecer_item_count + 1):
        with st.expander(f"Barang {i}", expanded=(i == st.session_state.ecer_item_count)):  # ← SESUDAHNYA
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
            
            # REAL-TIME VALIDATION
            missing_fields = []
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
            
            # **NEW: Add semantic overlap check**
            validation_errors = []
            
            if cat and cat != "-" and name and name != "-":
                from utils.validators import semantic_overlap_check
                has_overlap, overlap_msg = semantic_overlap_check(cat, name)
                if has_overlap:
                    validation_errors.append(overlap_msg)
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ {overlap_msg}'
                        f'</div>',
                        unsafe_allow_html=True
                    )
            
            if cat or name or sz or sup or qty > 0 or harga > 0:
                if missing_fields or validation_errors:  # **CHANGED: Include validation_errors**
                    error_display = []
                    if missing_fields:
                        error_display.append(f"Belum lengkap: {', '.join(missing_fields)}")
                    if validation_errors:
                        error_display.extend(validation_errors)
                    
                    st.markdown(
                        f'<div class="item-validation item-invalid">'
                        f'❌ Barang {i}: <strong>{" | ".join(error_display)}</strong>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    item_validations.append((i, False, missing_fields + validation_errors))
                else:
                    st.markdown(
                        f'<div class="item-validation item-valid">'
                        f'✅ Barang {i}: Lengkap dan siap dijual'
                        f'</div>',
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
    
    # ========================================
    # ADD/DELETE ITEM BUTTONS - ECERAN
    # ========================================
    col_add_ecer, col_delete_ecer = st.columns([3, 1])
    
    with col_add_ecer:
        # Initialize ecer_item_count if not exists
        if 'ecer_item_count' not in st.session_state:
            st.session_state.ecer_item_count = MAX_ITEMS_PER_FORM
        
        if st.session_state.ecer_item_count < MAX_ITEMS_PER_FORM:
            if st.button("➕ Tambah Barang Baru", key="add_ecer_item_btn", use_container_width=True):
                st.session_state.ecer_item_count += 1
                st.rerun()
        else:
            st.info(f"ℹ️ Maksimal {MAX_ITEMS_PER_FORM} barang per form")
    
    with col_delete_ecer:
        if 'ecer_item_count' not in st.session_state:
            st.session_state.ecer_item_count = MAX_ITEMS_PER_FORM
        
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
    
    # DP PAYMENT HANDLING
    jumlah_dp = 0
    sisa_pembayaran = 0
    metode_sisa = None
    rekening_sisa = None
    
    if pay_stat == "DP":
        if total > 0:
            st.markdown("### 💳 Detail Pembayaran DP")
            jumlah_dp = st.number_input(
                "💵 Jumlah DP (Rp)", min_value=0, max_value=total, step=1000, 
                key="ecer_dp", help="Masukkan jumlah uang muka yang dibayarkan"
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
                metode_sisa = col_sisa1.text_input("Metode Sisa (manual)", key="ecer_metode_sisa_man", placeholder="Ketik metode pembayaran").strip()
                if not metode_sisa:
                    metode_sisa = "Manual"
            elif metode_sisa_selected == "-- Pilih Metode --":
                metode_sisa = None
            else:
                metode_sisa = metode_sisa_selected
            
            if metode_sisa == "Transfer" or metode_sisa_selected == "Transfer":
                rek_sisa_options = ["-- Pilih Rekening --"] + REKENING_OPTIONS
                rek_sisa_opt = col_sisa2.selectbox("Rekening untuk Sisa Pembayaran *", rek_sisa_options, key="ecer_rek_sisa")
                
                if rek_sisa_opt == "Lainnya (ketik manual)":
                    rekening_sisa = col_sisa2.text_input("Rekening Sisa manual", key="ecer_rek_sisa_man", placeholder="Ketik rekening").strip()
                    if not rekening_sisa:
                        rekening_sisa = "Manual"
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
        
        # Validate customer info
        if not cust:
            validation_errors.append("❌ **Nama Pelanggan** wajib diisi!")
        if not alamat:
            validation_errors.append("❌ **Alamat** wajib diisi!")
        
        # Validate transaction details
        if not spg or spg == "-- Pilih SPG --":
            validation_errors.append("❌ **SPG** wajib dipilih!")
        if not pay_type or pay_type == "-- Pilih Pembayaran --":
            validation_errors.append("❌ **Tipe Pembayaran** wajib dipilih!")
        if pay_type == "Transfer" and (not rekening or rekening == "-- Pilih Rekening --"):
            validation_errors.append("❌ **Rekening** wajib dipilih untuk pembayaran Transfer!")
        if not pay_stat or pay_stat == "-- Pilih Status --":
            validation_errors.append("❌ **Status Pembayaran** wajib dipilih!")
        if not kirim or kirim == "-- Pilih Status --":
            validation_errors.append("❌ **Status Pengiriman** wajib dipilih!")
        
        # Validate items
        if not ecer_items:
            validation_errors.append("❌ **Belum ada barang yang lengkap**. Pastikan setiap barang memiliki: Nama Barang, Ukuran, Qty > 0, dan Harga > 0 (kecuali Bonus)")
        
        incomplete_items = [v for v in item_validations if not v[1]]
        if incomplete_items:
            for item_num, is_valid, missing in incomplete_items:
                validation_errors.append(f"❌ **Barang {item_num}**: Belum diisi - {', '.join(missing)}")
        
        # Validate DP
        if pay_stat == "DP" and jumlah_dp <= 0:
            validation_errors.append("❌ **Jumlah DP** harus diisi dan lebih besar dari 0!")
        
        if validation_errors:
            st.error("### ⚠️ Validasi Gagal - Perbaiki Error Berikut:")
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
                    except: 
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
            
            # Force form refresh
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
    
    # **NEW: Enhanced filters with location multi-select**
    col_f1, col_f2 = st.columns(2)
    
    # 1. Global Search
    search_query = col_f1.text_input(
        "🔍 Cari", 
        placeholder="Ketik nama barang, supplier, kategori, ukuran, warna..."
    ).strip().lower()
    
    # 2. Location Filter - Multi-select
    location_options = ["Semua Lokasi", "Toko Only", "Gudang Only", "Pabrik Only"]
    location_filter = col_f2.multiselect(
        "📍 Filter Lokasi",
        location_options,
        default=["Semua Lokasi"],
        key="stok_location_filter"
    )
    
    try:
        # Load data
        inventory = get_cached_inventory() 
        
        # Apply location filter FIRST
        if "Semua Lokasi" not in location_filter and location_filter:
            filtered_inventory = []
            for item in inventory:
                loc = item.get('location', '').lower()
                
                # Check if location matches any selected filter
                if "Toko Only" in location_filter and "toko" in loc:
                    filtered_inventory.append(item)
                elif "Gudang Only" in location_filter and "gudang" in loc:
                    filtered_inventory.append(item)
                elif "Pabrik Only" in location_filter and "pabrik" in loc:
                    filtered_inventory.append(item)
            
            inventory = filtered_inventory
        else:
            # If "Semua Lokasi" selected or no filter, exclude Pabrik only
            inventory = [item for item in inventory if not is_pabrik_source(item.get('location', ''))]
        
        if inventory:
            df_data = []
            for item in inventory:
                # Tanggal Input (Saat form diisi pertama kali) - CONVERT TO JAKARTA TIME
                tgl_isi = item.get('created_at')
                if tgl_isi:
                    tgl_isi_local = convert_to_local_time(tgl_isi)
                    tgl_isi_str = tgl_isi_local.strftime("%d/%m/%y %H:%M") if tgl_isi_local else "—"
                else:
                    tgl_isi_str = "—"
                
                # Terakhir Update (Tanggal operasional dari form eceran/masuk)
                tgl_update = item.get('date_added') or item.get('updated_at')
                tgl_update_str = format_date_dd_mm_yy(tgl_update)
                
                # Build merged name
                item_name = item['name'].title()
                size = item.get('size', '').strip()
                color = item.get('color', '').strip()
                
                # Merge format: "Name (Size) - Color" or variations
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

            # 2. Logika Pencarian Global (Multi-Kolom)
            if search_query:
                # Mencari kata kunci di semua kolom yang ada di DataFrame
                mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
                df = df[mask]

            # 3. Metrics (Dihitung berdasarkan hasil pencarian)
            col1, col2, col3 = st.columns(3)
            col1.metric("📦 Total Items", len(df))
            col2.metric("📊 Total Qty", f"{df['Qty'].sum():,}")
            negative_count = len(df[df['Qty'] < 0])
            if negative_count > 0: 
                col3.metric("⚠️ Stok Negatif", negative_count, delta=f"-{negative_count}", delta_color="inverse")
            else: 
                col3.metric("✅ Stok Normal", len(df))
            
            st.markdown("---")
            
            # 4. Fungsi Styling untuk Stok Negatif
            def highlight_negative(row):
                if row['Qty'] < 0: 
                    return ['background-color: #442222; color: #ffaaaa'] * len(row)
                return [''] * len(row)
            
            # 5. Tampilkan Tabel
            if not df.empty:
                st.dataframe(
                    df.style.apply(highlight_negative, axis=1), 
                    use_container_width=True, 
                    hide_index=True, 
                    height=1500
                )
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
    
    # 1. Proteksi Password
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

    # **NEW: Enhanced filters**
    col_f1, col_f2 = st.columns(2)
    
    # 2. Global Search Filter
    search_query = col_f1.text_input(
        "🔍 Cari (Barang, Supplier, PO, Kategori, Ukuran, Warna, Lokasi)", 
        placeholder="Ketik kata kunci..."
    ).strip().lower()
    
    # 3. Location Filter - Multi-select
    location_options = ["Semua Lokasi", "Toko Only", "Gudang Only", "Pabrik Only"]
    location_filter = col_f2.multiselect(
        "📍 Filter Lokasi",
        location_options,
        default=["Semua Lokasi"],
        key="edit_location_filter"
    )
    
    try:
        # Load data
        inventory = get_cached_inventory()
        
        # Apply location filter FIRST
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
            # If "Semua Lokasi" selected, exclude Pabrik only
            inventory = [item for item in inventory if not is_pabrik_source(item.get('location', ''))]

        # 3. Siapkan Dataframe
        df_data = []
        for i, item in enumerate(inventory):
            details = item.get('details', {})
            komen_val = details.get('komen', '—') if isinstance(details, dict) else '—'
            
            # Tanggal Input - CONVERT TO JAKARTA TIME
            tgl_input = item.get('created_at')
            if tgl_input:
                tgl_input_local = convert_to_local_time(tgl_input)
                tgl_input_str = tgl_input_local.strftime("%d/%m/%y %H:%M") if tgl_input_local else "—"
            else:
                tgl_input_str = "—"
            
            # Terakhir Update
            tgl_update = item.get('date_added') or item.get('updated_at')
            tgl_update_str = format_date_dd_mm_yy(tgl_update)
            
            # Build merged name for edit view
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

        # 4. Logika Pencarian
        if search_query:
            mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)
            df_filtered = df[mask]
        else:
            df_filtered = df

        st.info("💡 Klik sel untuk edit, '+' untuk tambah baris, atau pilih baris & tekan 'Delete' untuk hapus.")
        
        # 5. Data Editor
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

        # Add download buttons
        st.markdown("---")
        st.markdown("#### 📥 Download Data")
        create_download_button_excel(
            df_filtered, 
            f"stok_edit_{date.today().strftime('%d_%m_%y')}.xlsx",
            "📥 Download Excel",
            "Stock Edit Report"
        )
        st.markdown("---")
        
        # 6. Save Button
        if st.button("💾 Simpan Perubahan Spreadsheet", type="primary", use_container_width=True):
            state = st.session_state["stock_editor_spreadsheet"]
            success_total = 0
            
            def to_python_int(val):
                if val is None or pd.isna(val): return 0
                try: return int(float(val))
                except: return 0

            def to_python_str(val):
                if val is None or pd.isna(val): return ""
                return str(val).strip()

            # A. Delete
            for idx in state.get("deleted_rows", []):
                try:
                    db_id = df_filtered.iloc[idx]["UUID"]
                    delete_inventory_item(db_id)
                    success_total += 1
                except: pass

            # B. Edit
            for idx, changes in state.get("edited_rows", {}).items():
                try:
                    row_original = df_filtered.iloc[idx]
                    db_id = row_original["UUID"]
                    
                    update_inventory_item(
                        item_id=db_id,
                        name=to_python_str(changes.get("Barang", row_original["Barang"])).lower(),
                        category=to_python_str(changes.get("Kategori", row_original["Kategori"])).lower(),
                        color=to_python_str(changes.get("Warna", row_original["Warna"])),
                        size=to_python_str(changes.get("Ukuran", row_original["Ukuran"])),
                        quantity=to_python_int(changes.get("Qty", row_original["Qty"])),
                        location=to_python_str(changes.get("Lokasi", row_original["Lokasi"])),
                        supplier=to_python_str(changes.get("Supplier", row_original["Supplier"]))
                    )
                    success_total += 1
                except Exception as e:
                    st.error(f"Gagal update baris {idx+1}: {e}")

            # C. Add New
            for row_new in state.get("added_rows", []):
                nama_item = to_python_str(row_new.get("Barang"))
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
# HISTORY TAB (REVISED WITH MERGED CELLS)
# ==================================================

# ==================================================
# HISTORY TAB (HTML VERSION WITH TRUE MERGED CELLS)
# ==================================================

def create_merged_html_table(df, merge_columns, group_by_column):
    """
    Create HTML table with merged cells (rowspan) for specified columns.
    
    Args:
        df: DataFrame to display
        merge_columns: List of column names to merge when values are same
        group_by_column: Column to group by for merging
    
    Returns:
        HTML string
    """
    # Start table
    html = """
    <style>
        .merged-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
            font-family: 'Source Sans Pro', sans-serif;
        }
        .merged-table th {
            background-color: #1f1f1f;
            color: #ffffff;
            padding: 12px 8px;
            text-align: left;
            border: 1px solid #444;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        .merged-table td {
            padding: 10px 8px;
            border: 1px solid #444;
            color: #ffffff;
            background-color: #0e1117;
            vertical-align: middle;
        }
        .merged-table tr:hover td {
            background-color: #1a1a1a;
        }
        .merged-cell {
            background-color: #151515;
            font-weight: 500;
        }
    </style>
    """
    
    html += '<div style="max-height: 600px; overflow-y: auto;"><table class="merged-table">'
    
    # Table header
    html += '<thead><tr>'
    for col in df.columns:
        if col not in ['UUID', 'ID']:  # Hide technical columns
            html += f'<th>{col}</th>'
    html += '</tr></thead><tbody>'
    
    # Group data
    grouped = df.groupby(group_by_column, sort=False)
    
    for group_name, group in grouped:
        group_size = len(group)
        
        for row_idx, (idx, row) in enumerate(group.iterrows()):
            html += '<tr>'
            
            for col in df.columns:
                if col in ['UUID', 'ID']:  # Skip technical columns
                    continue
                
                # Check if this column should be merged
                if col in merge_columns:
                    # Only add cell on first row of group
                    if row_idx == 0:
                        cell_value = str(row[col])
                        if cell_value and cell_value != '' and cell_value != 'nan':
                            html += f'<td class="merged-cell" rowspan="{group_size}">{cell_value}</td>'
                        else:
                            html += f'<td class="merged-cell" rowspan="{group_size}">—</td>'
                    # Skip for other rows (cell is merged)
                else:
                    # Regular cell (not merged)
                    cell_value = str(row[col])
                    if cell_value and cell_value != '' and cell_value != 'nan':
                        html += f'<td>{cell_value}</td>'
                    else:
                        html += f'<td>—</td>'
            
            html += '</tr>'
    
    html += '</tbody></table></div>'
    return html


# ==================================================
# HISTORY TAB (UPDATED WITH COMPREHENSIVE FILTERS)
# ==================================================

def inject_merged_table_css():
    """Inject CSS for merged tables with correct styling and horizontal scroll"""
    st.markdown("""
    <style>
    /* Merged table container with horizontal scroll */
    .merged-table-container {
        max-height: 600px;
        overflow-y: auto;
        overflow-x: auto;
        border-radius: 8px;
        border: 1px solid #4a5568;
        margin: 1rem 0;
        width: 100%;
    }
    
    /* Merged table styles */
    .merged-table {
        width: auto;
        min-width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        table-layout: auto;
    }
    
    /* Table headers - sticky and auto-width */
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
    
    /* Table body */
    .merged-table tbody {
        background-color: #1a202c;
        color: #e2e8f0;
    }
    
    .merged-table td {
        padding: 8px 6px;
        border: 1px solid #4a5568;
        vertical-align: middle;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 200px;
    }
    
    /* Allow text wrapping for specific columns */
    .merged-table td:nth-child(8),  /* Nama Barang column */
    .merged-table td:last-child {    /* Komen column */
        white-space: normal;
        word-wrap: break-word;
        max-width: 250px;
    }
    
    /* Hover effect */
    .merged-table tbody tr:hover td {
        background-color: #2d3748;
    }
    
    /* ============================================ */
    /* CRITICAL: Merged cells styling - HIDE BORDERS */
    /* ============================================ */
    .merged-cell {
        background-color: #2a3441 !important;
        text-align: center;
        font-weight: 500;
        border-top: 1px solid #4a5568 !important;
        border-bottom: 1px solid #4a5568 !important;
    }
    
    /* Remove inner borders between merged cells */
    .merged-cell + td {
        border-left: none !important;
    }
    
    tr:has(.merged-cell) + tr td:nth-child(n) {
        border-top: none !important;
    }
    
    /* ID column - fixed small width */
    .merged-table th:first-child,
    .merged-table td:first-child {
        min-width: 40px;
        max-width: 50px;
        text-align: center;
    }
    
    /* UUID column - fixed small width */
    .merged-table th:nth-child(2),
    .merged-table td:nth-child(2) {
        min-width: 60px;
        max-width: 80px;
        font-family: monospace;
        font-size: 0.75rem;
    }
    
    /* Date columns - medium width */
    .merged-table th:nth-child(3),
    .merged-table th:nth-child(4),
    .merged-table td:nth-child(3),
    .merged-table td:nth-child(4) {
        min-width: 90px;
        max-width: 120px;
    }
    
    /* Scrollbar styling for dark theme */
    .merged-table-container::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }
    
    .merged-table-container::-webkit-scrollbar-track {
        background: #1a202c;
        border-radius: 5px;
    }
    
    .merged-table-container::-webkit-scrollbar-thumb {
        background: #4a5568;
        border-radius: 5px;
    }
    
    .merged-table-container::-webkit-scrollbar-thumb:hover {
        background: #718096;
    }
    </style>
    """, unsafe_allow_html=True)
    
def render_merged_table_masuk(records: List[Dict]) -> str:
    """Generate HTML table for Barang Masuk with merged cells — Image 1 format."""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"

    # Ensure created_at exists
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())

    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))

    html = '<div class="merged-table-container">'
    html += '<table class="merged-table">'
    html += '''
    <thead>
        <tr>
            <th>ID</th>
            <th>UUID</th>
            <th>Tanggal Input Form</th>
            <th>Masuk Tanggal</th>
            <th>Nomor PO</th>
            <th>Supplier</th>
            <th>Kategori</th>
            <th>Nama Barang, Ukuran, Warna</th>
            <th style="text-align: center;">Quantity</th>
            <th style="text-align: center;">Total Quantity</th>
            <th>Lokasi</th>
            <th>Komen</th>
        </tr>
    </thead>
    <tbody>
    '''

    record_id = 1

    # Group with 2-minute tolerance
    def get_time_bucket(record, tolerance_minutes=2):
        """Round timestamp to nearest N minutes for grouping"""
        dt = convert_to_local_time(record.get('created_at'))
        if not dt:
            return "unknown"
        
        # Round down to nearest tolerance_minutes
        minutes = (dt.minute // tolerance_minutes) * tolerance_minutes
        rounded = dt.replace(minute=minutes, second=0, microsecond=0)
        return rounded.strftime("%d/%m/%y %H:%M")

    # Group strictly by created_at timestamp with tolerance
    for created_at_key, grp in groupby(
        sorted_records,
        key=lambda x: get_time_bucket(x, tolerance_minutes=2)
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty  = sum(abs(r['quantity']) for r in group_list)

        first     = group_list[0]
        det_first = first.get('details') or {}

        # Reference values from first record
        masuk_tanggal_first = format_date_dd_mm_yy(first['transaction_date'])
        nomor_po_first      = det_first.get('nomor_po') or '—'
        supplier_first      = first.get('supplier') or '—'
        lokasi_first        = first.get('location') or '—'

        # FIX: proper None-safe comparisons for merge checks
        same_date = all(
            format_date_dd_mm_yy(r['transaction_date']) == masuk_tanggal_first
            for r in group_list
        )
        same_po = all(
            ((r.get('details') or {}).get('nomor_po') or '—') == nomor_po_first
            for r in group_list
        )
        same_supplier = all(
            (r.get('supplier') or '—') == supplier_first
            for r in group_list
        )
        same_location = all(
            (r.get('location') or '—') == lokasi_first
            for r in group_list
        )

        for idx, rec in enumerate(group_list):
            det  = rec.get('details') or {}
            html += '<tr>'

            # ── ID ──
            html += f'<td>{record_id}</td>'

            # ── UUID (first 8 chars) ──
            uuid_val = rec.get('uuid') or rec.get('id') or 'N/A'
            html += f'<td style="font-family: monospace; font-size: 0.75rem;">{str(uuid_val)[:8]}</td>'

            # ── Tanggal Input Form — ALWAYS merged for the group ──
            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}">{created_at_key}</td>'

            # ── Masuk Tanggal — merged if all same ──
            if same_date:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{masuk_tanggal_first}</td>'
            else:
                html += f'<td>{format_date_dd_mm_yy(rec["transaction_date"])}</td>'

            # ── Nomor PO — merged if all same ──
            if same_po:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{nomor_po_first}</td>'
            else:
                html += f'<td>{det.get("nomor_po") or "—"}</td>'

            # ── Supplier — merged if all same ──
            if same_supplier:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{supplier_first}</td>'
            else:
                html += f'<td>{rec.get("supplier") or "—"}</td>'

            # ── Kategori — NEVER merged, each row shows its own ──
            html += f'<td>{(rec.get("category") or "—").upper()}</td>'

            # ── Nama Barang, Ukuran, Warna — combined, never merged ──
            combined_name = format_item_combined(
                rec.get('item_name', ''),
                rec.get('size', ''),
                rec.get('color', '')
            )
            html += f'<td>{combined_name}</td>'

            # ── Quantity — per row ──
            html += f'<td style="text-align: center;">{abs(rec["quantity"])}</td>'

            # ── Total Quantity — ALWAYS merged (group sum) ──
            if idx == 0:
                html += (
                    f'<td class="merged-cell" rowspan="{group_size}" '
                    f'style="font-weight: bold; text-align: center;">{total_qty}</td>'
                )

            # ── Lokasi — merged if all same ──
            if same_location:
                if idx == 0:
                    html += f'<td class="merged-cell" rowspan="{group_size}">{lokasi_first}</td>'
            else:
                html += f'<td>{rec.get("location") or "—"}</td>'

            # ── Komen — per row ──
            html += f'<td>{det.get("komen") or "—"}</td>'

            html += '</tr>'
            record_id += 1

    html += '</tbody></table></div>'
    return html


def render_merged_table_luar_kota(records: List[Dict]) -> str:
    """Generate HTML table for Luar Kota with merged cells"""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"
    
    # **CRITICAL FIX: Ensure created_at exists**
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())
    
    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))
    
    html = '<div class="merged-table-container">'
    html += '<table class="merged-table">'
    html += '''
    <thead>
        <tr>
            <th>ID</th>
            <th>UUID</th>
            <th>Tanggal Input Form</th>
            <th>Pengiriman Tanggal</th>
            <th>Tujuan</th>
            <th>Via</th>
            <th>Supplier</th>
            <th>Kategori</th>
            <th>Nama Barang, Ukuran, Warna</th>
            <th style="text-align: center;">Quantity</th>
            <th style="text-align: center;">Total Quantity</th>
            <th>Lokasi</th>
            <th>Komen</th>
        </tr>
    </thead>
    <tbody>
    '''
    
    record_id = 1
    
    # Group with 2-minute tolerance
    def get_time_bucket_luar(record, tolerance_minutes=2):
        dt = convert_to_local_time(record.get('created_at'))
        if not dt:
            return "unknown"
        minutes = (dt.minute // tolerance_minutes) * tolerance_minutes
        rounded = dt.replace(minute=minutes, second=0, microsecond=0)
        return rounded.strftime("%d/%m/%y %H:%M")
    
    for created_at, group in groupby(
        sorted_records, 
        key=lambda x: get_time_bucket_luar(x, tolerance_minutes=2)
    ):
        group_list = list(group)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)
        
        first_record = group_list[0]
        tanggal_input = created_at
        pengiriman_tanggal = format_date_dd_mm_yy(first_record['transaction_date'])
        
        det_first = first_record.get('details', {})
        tujuan = det_first.get('purpose', '—')
        via = det_first.get('via', '—')
        supplier = first_record.get('supplier', '—')
        lokasi = first_record.get('location', '—')
        kategori = first_record.get('category', '—').upper()
        
        same_date = all(
            format_date_dd_mm_yy(r['transaction_date']) == pengiriman_tanggal 
            for r in group_list
        )
        same_tujuan = all(
            r.get('details', {}).get('purpose', '—') == tujuan 
            for r in group_list
        )
        same_via = all(
            r.get('details', {}).get('via', '—') == via 
            for r in group_list
        )
        same_supplier = all(
            r.get('supplier', '—') == supplier 
            for r in group_list
        )
        same_category = all(
            r.get('category', '—').upper() == kategori 
            for r in group_list
        )
        same_location = all(
            r.get('location', '—') == lokasi 
            for r in group_list
        )
        
        for idx, rec in enumerate(group_list):
            det = rec.get('details', {})
            html += '<tr>'
            
            html += f'<td>{record_id}</td>'
            uuid_short = rec.get('uuid', 'N/A')[:8]
            html += f'<td style="font-family: monospace; font-size: 0.75rem;">{uuid_short}</td>'
            
            # Tanggal Input Form (merged)
            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}">{tanggal_input}</td>'
            
            # Pengiriman Tanggal (merged if same)
            if idx == 0 and same_date:
                html += f'<td class="merged-cell" rowspan="{group_size}">{pengiriman_tanggal}</td>'
            elif not same_date:
                html += f'<td>{format_date_dd_mm_yy(rec["transaction_date"])}</td>'
            
            # Tujuan (merged if same)
            if idx == 0 and same_tujuan:
                html += f'<td class="merged-cell" rowspan="{group_size}">{tujuan}</td>'
            elif not same_tujuan:
                html += f'<td>{det.get("purpose", "—")}</td>'
            
            # Via (merged if same)
            if idx == 0 and same_via:
                html += f'<td class="merged-cell" rowspan="{group_size}">{via}</td>'
            elif not same_via:
                html += f'<td>{det.get("via", "—")}</td>'
            
            # Supplier (merged if same)
            if idx == 0 and same_supplier:
                html += f'<td class="merged-cell" rowspan="{group_size}">{supplier}</td>'
            elif not same_supplier:
                html += f'<td>{rec.get("supplier", "—")}</td>'
            
            # Kategori (merged if same)
            if idx == 0 and same_category:
                html += f'<td class="merged-cell" rowspan="{group_size}">{kategori}</td>'
            elif not same_category:
                html += f'<td>{rec.get("category", "—").upper()}</td>'
            
            # Nama Barang, Ukuran, Warna (combined)
            combined_name = format_item_combined(
                rec['item_name'],
                rec.get('size', ''),
                rec.get('color', '')
            )
            html += f'<td>{combined_name}</td>'
            
            # Quantity
            html += f'<td style="text-align: center;">{abs(rec["quantity"])}</td>'
            
            # Total Quantity (merged)
            if idx == 0:
                html += f'<td class="merged-cell" rowspan="{group_size}" style="font-weight: bold;">{total_qty}</td>'
            
            # Lokasi (merged if same)
            if idx == 0 and same_location:
                html += f'<td class="merged-cell" rowspan="{group_size}">{lokasi}</td>'
            elif not same_location:
                html += f'<td>{rec.get("location", "—")}</td>'
            
            # Komen (separate)
            komen = det.get('komen', '—')
            html += f'<td>{komen}</td>'
            
            html += '</tr>'
            record_id += 1
    
    html += '</tbody></table></div>'
    return html


def render_merged_table_eceran(records: List[Dict]) -> str:
    """Generate HTML table for Eceran with merged cells"""
    if not records:
        return "<p style='color: #888;'>Tidak ada data</p>"
    
    # **CRITICAL FIX: Ensure created_at exists**
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())
    
    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))
    
    html = '<div class="merged-table-container">'
    html += '<table class="merged-table">'
    html += '''
    <thead>
        <tr>
            <th>ID</th>
            <th>UUID</th>
            <th>Tanggal Input Form</th>
            <th>Tanggal Penjualan</th>
            <th>Pelanggan</th>
            <th>Supplier</th>
            <th>Kategori</th>
            <th>Nama Barang, Ukuran, Warna</th>
            <th style="text-align: center;">Quantity</th>
            <th style="text-align: center;">Total Quantity</th>
            <th style="text-align: right;">Harga Per Item</th>
            <th style="text-align: right;">Total Penjualan</th>
            <th>Status</th>
            <th>Pembayaran</th>
            <th>Rekening</th>
            <th style="text-align: right;">Total DP</th>
            <th style="text-align: right;">Sisa Pembayaran</th>
            <th>Metode Sisa</th>
            <th>Rekening Sisa</th>
            <th>Lokasi</th>
            <th>Tgl Kirim</th>
            <th>Komen</th>
        </tr>
    </thead>
    <tbody>
    '''
    
    record_id = 1

# ==================================================
# COMPLETE HISTORY TAB FUNCTION (REPLACE YOUR EXISTING ONE)
# ==================================================

def prepare_masuk_aggrid_data(records):
    """
    Prepare Barang Masuk data for AgGrid with row spanning.
    Returns DataFrame with special markers for merged cells.
    """
    if not records:
        return pd.DataFrame()
    
    # Ensure created_at exists
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())
    
    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))
    
    rows = []
    record_id = 1
    
    for created_at_key, grp in groupby(
        sorted_records,
        key=lambda x: format_local_datetime(x.get('created_at'), "%d/%m/%y %H:%M")
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)
        
        first_record = group_list[0]
        det_first = first_record.get('details', {})
        
        # Get values from first record for merged columns
        tanggal_masuk_first = format_date_dd_mm_yy(first_record['transaction_date'])
        nomor_po_first = det_first.get('nomor_po', '—')
        supplier_first = first_record.get('supplier', '—')
        lokasi_first = first_record.get('location', '—')
        
        # Check if values are same across group
        same_date = all(
            format_date_dd_mm_yy(r['transaction_date']) == tanggal_masuk_first 
            for r in group_list
        )
        same_po = all(
            (r.get('details', {}).get('nomor_po') or '—') == nomor_po_first 
            for r in group_list
        )
        same_supplier = all(
            (r.get('supplier') or '—') == supplier_first 
            for r in group_list
        )
        same_location = all(
            (r.get('location') or '—') == lokasi_first 
            for r in group_list
        )
        
        for idx, rec in enumerate(group_list):
            det = rec.get('details', {})
            
            combined_name = format_item_combined(
                rec.get('item_name', ''),
                rec.get('size', ''),
                rec.get('color', '')
            )
            
            row = {
                'ID': record_id,
                'UUID': str(rec.get('uuid', 'N/A'))[:8],
                
                # Tanggal Input - Always merged for group
                'Tanggal Input': created_at_key if idx == 0 else '↑',
                '_span_tanggal_input': group_size if idx == 0 else 0,
                
                # Tanggal Masuk - Merged if all same
                'Tanggal Masuk': tanggal_masuk_first if (idx == 0 or not same_date) else ('↑' if same_date else format_date_dd_mm_yy(rec['transaction_date'])),
                '_span_tanggal_masuk': group_size if (idx == 0 and same_date) else (1 if not same_date else 0),
                
                # Nomor PO - Merged if all same
                'Nomor PO': nomor_po_first if (idx == 0 or not same_po) else ('↑' if same_po else det.get('nomor_po', '—')),
                '_span_nomor_po': group_size if (idx == 0 and same_po) else (1 if not same_po else 0),
                
                # Supplier - Merged if all same
                'Supplier': supplier_first if (idx == 0 or not same_supplier) else ('↑' if same_supplier else rec.get('supplier', '—')),
                '_span_supplier': group_size if (idx == 0 and same_supplier) else (1 if not same_supplier else 0),
                
                # Non-merged columns
                'Kategori': (rec.get('category', '—')).upper(),
                'Nama Barang': combined_name,
                'Qty': abs(rec['quantity']),
                
                # Total Qty - Always merged for group
                'Total Qty': total_qty if idx == 0 else '↑',
                '_span_total_qty': group_size if idx == 0 else 0,
                
                # Lokasi - Merged if all same
                'Lokasi': lokasi_first if (idx == 0 or not same_location) else ('↑' if same_location else rec.get('location', '—')),
                '_span_lokasi': group_size if (idx == 0 and same_location) else (1 if not same_location else 0),
                
                'Komen': det.get('komen', '—'),
                
                # Hidden columns for filtering/editing
                '_created_at': created_at_key,
                '_uuid_full': rec.get('uuid', 'N/A'),
                '_record': rec  # Store full record for editing
            }
            rows.append(row)
            record_id += 1
    
    return pd.DataFrame(rows)


def prepare_luar_kota_aggrid_data(records):
    """Prepare Luar Kota data for AgGrid with row spanning"""
    if not records:
        return pd.DataFrame()
    
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())
    
    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))
    
    rows = []
    record_id = 1
    
    for created_at_key, grp in groupby(
        sorted_records,
        key=lambda x: format_local_datetime(x.get('created_at'), "%d/%m/%y %H:%M")
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)
        
        first_record = group_list[0]
        det_first = first_record.get('details', {})
        
        tanggal_kirim_first = format_date_dd_mm_yy(first_record['transaction_date'])
        tujuan_first = det_first.get('purpose', '—')
        via_first = det_first.get('via', '—')
        supplier_first = first_record.get('supplier', '—')
        lokasi_first = first_record.get('location', '—')
        
        same_date = all(format_date_dd_mm_yy(r['transaction_date']) == tanggal_kirim_first for r in group_list)
        same_tujuan = all((r.get('details', {}).get('purpose') or '—') == tujuan_first for r in group_list)
        same_via = all((r.get('details', {}).get('via') or '—') == via_first for r in group_list)
        same_supplier = all((r.get('supplier') or '—') == supplier_first for r in group_list)
        same_location = all((r.get('location') or '—') == lokasi_first for r in group_list)
        
        for idx, rec in enumerate(group_list):
            det = rec.get('details', {})
            combined_name = format_item_combined(
                rec.get('item_name', ''),
                rec.get('size', ''),
                rec.get('color', '')
            )
            
            row = {
                'ID': record_id,
                'UUID': str(rec.get('uuid', 'N/A'))[:8],
                'Tanggal Input': created_at_key if idx == 0 else '↑',
                '_span_tanggal_input': group_size if idx == 0 else 0,
                
                'Tanggal Kirim': tanggal_kirim_first if (idx == 0 or not same_date) else ('↑' if same_date else format_date_dd_mm_yy(rec['transaction_date'])),
                '_span_tanggal_kirim': group_size if (idx == 0 and same_date) else (1 if not same_date else 0),
                
                'Tujuan': tujuan_first if (idx == 0 or not same_tujuan) else ('↑' if same_tujuan else det.get('purpose', '—')),
                '_span_tujuan': group_size if (idx == 0 and same_tujuan) else (1 if not same_tujuan else 0),
                
                'Via': via_first if (idx == 0 or not same_via) else ('↑' if same_via else det.get('via', '—')),
                '_span_via': group_size if (idx == 0 and same_via) else (1 if not same_via else 0),
                
                'Supplier': supplier_first if (idx == 0 or not same_supplier) else ('↑' if same_supplier else rec.get('supplier', '—')),
                '_span_supplier': group_size if (idx == 0 and same_supplier) else (1 if not same_supplier else 0),
                
                'Kategori': (rec.get('category', '—')).upper(),
                'Nama Barang': combined_name,
                'Qty': abs(rec['quantity']),
                
                'Total Qty': total_qty if idx == 0 else '',
                '_span_total_qty': group_size if idx == 0 else 0,
                
                'Lokasi': lokasi_first if (idx == 0 or not same_location) else ('↑' if same_location else rec.get('location', '—')),
                '_span_lokasi': group_size if (idx == 0 and same_location) else (1 if not same_location else 0),
                
                'Komen': det.get('komen', '—'),
                '_created_at': created_at_key,
                '_uuid_full': rec.get('uuid', 'N/A'),
                '_record': rec
            }
            rows.append(row)
            record_id += 1
    
    return pd.DataFrame(rows)


def prepare_eceran_aggrid_data(records):
    """Prepare Eceran data for AgGrid with row spanning"""
    if not records:
        return pd.DataFrame()
    
    for rec in records:
        if 'created_at' not in rec or not rec['created_at']:
            rec['created_at'] = rec.get('transaction_date', datetime.now())
    
    sorted_records = sorted(records, key=lambda x: x.get('created_at', ''))
    
    rows = []
    record_id = 1
    
    for created_at_key, grp in groupby(
        sorted_records,
        key=lambda x: format_local_datetime(x.get('created_at'), "%d/%m/%y %H:%M")
    ):
        group_list = list(grp)
        group_size = len(group_list)
        total_qty = sum(abs(r['quantity']) for r in group_list)
        total_sales = sum(r.get('details', {}).get('total_price', 0) for r in group_list)
        
        first_record = group_list[0]
        det_first = first_record.get('details', {})
        
        tanggal_jual_first = format_date_dd_mm_yy(first_record['transaction_date'])
        pelanggan_first = det_first.get('customer', '—')
        supplier_first = first_record.get('supplier', '—')
        lokasi_first = first_record.get('location', '—')
        status_first = det_first.get('payment_status', '—')
        pembayaran_first = det_first.get('payment_type', '—')
        rekening_first = det_first.get('rekening', '—')
        
        same_date = all(format_date_dd_mm_yy(r['transaction_date']) == tanggal_jual_first for r in group_list)
        same_pelanggan = all((r.get('details', {}).get('customer') or '—') == pelanggan_first for r in group_list)
        same_supplier = all((r.get('supplier') or '—') == supplier_first for r in group_list)
        same_location = all((r.get('location') or '—') == lokasi_first for r in group_list)
        same_status = all((r.get('details', {}).get('payment_status') or '—') == status_first for r in group_list)
        same_pembayaran = all((r.get('details', {}).get('payment_type') or '—') == pembayaran_first for r in group_list)
        same_rekening = all((r.get('details', {}).get('rekening') or '—') == rekening_first for r in group_list)
        
        for idx, rec in enumerate(group_list):
            det = rec.get('details', {})
            combined_name = format_item_combined(
                rec.get('item_name', ''),
                rec.get('size', ''),
                rec.get('color', '')
            )
            
            row = {
                'ID': record_id,
                'UUID': str(rec.get('uuid', 'N/A'))[:8],
                
                'Tanggal Input': created_at_key if idx == 0 else '↑',
                '_span_tanggal_input': group_size if idx == 0 else 0,
                
                'Tanggal Jual': tanggal_jual_first if (idx == 0 or not same_date) else ('↑' if same_date else format_date_dd_mm_yy(rec['transaction_date'])),
                '_span_tanggal_jual': group_size if (idx == 0 and same_date) else (1 if not same_date else 0),
                
                'Pelanggan': pelanggan_first if (idx == 0 or not same_pelanggan) else ('↑' if same_pelanggan else det.get('customer', '—')),
                '_span_pelanggan': group_size if (idx == 0 and same_pelanggan) else (1 if not same_pelanggan else 0),
                
                'Supplier': supplier_first if (idx == 0 or not same_supplier) else ('↑' if same_supplier else rec.get('supplier', '—')),
                '_span_supplier': group_size if (idx == 0 and same_supplier) else (1 if not same_supplier else 0),
                
                'Kategori': (rec.get('category', '—')).upper(),
                'Nama Barang': combined_name,
                'Qty': abs(rec['quantity']),
                
                'Total Qty': total_qty if idx == 0 else '↑',
                '_span_total_qty': group_size if idx == 0 else 0,
                
                'Harga Item': format_currency(det.get('total_price', 0)),
                
                'Total Penjualan': format_currency(total_sales) if idx == 0 else '↑',
                '_span_total_penjualan': group_size if idx == 0 else 0,
                
                'Status': status_first if (idx == 0 or not same_status) else ('↑' if same_status else det.get('payment_status', '—')),
                '_span_status': group_size if (idx == 0 and same_status) else (1 if not same_status else 0),
                
                'Pembayaran': pembayaran_first if (idx == 0 or not same_pembayaran) else ('↑' if same_pembayaran else det.get('payment_type', '—')),
                '_span_pembayaran': group_size if (idx == 0 and same_pembayaran) else (1 if not same_pembayaran else 0),
                
                'Rekening': rekening_first if (idx == 0 or not same_rekening) else ('↑' if same_rekening else det.get('rekening', '—')),
                '_span_rekening': group_size if (idx == 0 and same_rekening) else (1 if not same_rekening else 0),
                
                'Lokasi': lokasi_first if (idx == 0 or not same_location) else ('↑' if same_location else rec.get('location', '—')),
                '_span_lokasi': group_size if (idx == 0 and same_location) else (1 if not same_location else 0),
                
                'Komen': det.get('komen', '—'),
                '_created_at': created_at_key,
                '_uuid_full': rec.get('uuid', 'N/A'),
                '_record': rec
            }
            rows.append(row)
            record_id += 1
    
    return pd.DataFrame(rows)


def build_aggrid_config(df, height=600):
    """
    Build AgGrid configuration with:
    - Row spanning for merged cells
    - Filtering, sorting enabled
    - Dark theme
    """
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Enable features
    gb.configure_default_column(
        filterable=True,
        sortable=True,
        resizable=True,
        wrapText=True,
        autoHeight=True
    )
    
    # Hide internal columns
    hidden_cols = [col for col in df.columns if col.startswith('_')]
    for col in hidden_cols:
        gb.configure_column(col, hide=True)
    
    # Configure ID column
    gb.configure_column('ID', width=60, pinned='left')
    gb.configure_column('UUID', width=80, pinned='left')
    
    # Enable grid features
    gb.configure_grid_options(
        enableRangeSelection=True,
        rowHeight=35,
        headerHeight=40,
        suppressRowHoverHighlight=False,
        suppressCellFocus=False
    )
    
    # Enable pagination
    gb.configure_pagination(
        enabled=True,
        paginationAutoPageSize=False,
        paginationPageSize=50
    )
    
    # Side bar for filters
    gb.configure_side_bar(
        filters_panel=True,
        columns_panel=True,
        defaultToolPanel=""
    )
    
    grid_options = gb.build()
    
    return grid_options


def render_aggrid_table(df, height=600):
    """Render AgGrid table with all features enabled"""
    if df.empty:
        st.info("Tidak ada data")
        return None
    
    grid_options = build_aggrid_config(df, height)
    
    # ✅ ADD THIS LINE HERE:
    inject_aggrid_merge_css()
    
    # Custom CSS for dark theme
    custom_css = {
        ".ag-header-cell-text": {"color": "#ffffff"},
        ".ag-cell": {"color": "#e0e0e0"},
        ".ag-row": {"background-color": "#1a1a1a"},
        ".ag-row-odd": {"background-color": "#0e1117"},
        ".ag-row-hover": {"background-color": "#2d3748"}
    }
    
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        height=height,
        theme='streamlit',  # Use 'alpine' for light theme, 'balham' for professional
        custom_css=custom_css,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
        allow_unsafe_jscode=True
    )
    
    return grid_response


def render_history_tab(today: date) -> None:
    """History tab using HTML tables with TRUE merged cells + full filtering + Excel download"""
    
    # ========================================
    # PASSWORD PROTECTION
    # ========================================
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
    
    # ========================================
    # FILTER CONTROLS - ENHANCED
    # ========================================
    st.markdown("#### 🔍 Filter Data")
    
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Dari", today - timedelta(days=30), key="hist_from", format="DD/MM/YYYY")
    end_date = col2.date_input("Sampai", today, min_value=start_date, key="hist_to", format="DD/MM/YYYY")
    
    col3, col4 = st.columns(2)
    activity_options = ["Semua", "Masuk Barang", "Luar Kota", "Eceran"]
    activity_selected = col3.selectbox("Jenis Aktivitas", activity_options, key="hist_activity_filter")
    
    # Global search across ALL fields
    global_search = col4.text_input(
        "🔍 Cari (Barang/Supplier/PO/Kategori/Ukuran/Warna/Lokasi)", 
        key="hist_global_search", 
        placeholder="Ketik kata kunci..."
    ).strip().lower()
    
    activity_map = {
        "Semua": None,
        "Masuk Barang": "MASUK",
        "Luar Kota": "LUAR_KOTA",
        "Eceran": "ECERAN"
    }
    activity_filter = activity_map[activity_selected]
    
    # ========================================
    # ADVANCED FILTERS (EXPANDABLE)
    # ========================================
    with st.expander("🎯 Filter Lanjutan", expanded=False):
        col_adv1, col_adv2, col_adv3 = st.columns(3)
        
        # Get all unique values for filters
        all_suppliers = sorted(list(set(st.session_state.get("dynamic_supplier", SUPPLIERS))))
        all_kategori = sorted(list(set(st.session_state.get("dynamic_kategori", STANDARD_CATEGORIES))))
        all_locations = sorted(list(set(STANDARD_LOCATIONS)))
        
        # Supplier filter
        supplier_options = ["Semua"] + all_suppliers
        supplier_selected = col_adv1.selectbox("Supplier", supplier_options, key="hist_supplier_filter")
        supplier_filter = None if supplier_selected == "Semua" else supplier_selected
        
        # Kategori filter
        kategori_options = ["Semua"] + all_kategori
        kategori_selected = col_adv2.selectbox("Kategori", kategori_options, key="hist_kategori_filter")
        kategori_filter = None if kategori_selected == "Semua" else kategori_selected
        
        # Location filter
        location_options = ["Semua"] + all_locations
        location_selected = col_adv3.selectbox("Lokasi", location_options, key="hist_location_filter")
        location_filter = None if location_selected == "Semua" else location_selected
    
    # ========================================
    # CONDITIONAL FILTERS (Based on Activity Type)
    # ========================================
    kirim_ke_filter = None
    via_filter = None
    payment_type_filter = None
    payment_status_filter = None
    price_min = None
    price_max = None
    
    if activity_filter == "LUAR_KOTA":
        st.markdown("##### 🚚 Filter Luar Kota")
        col_lk1, col_lk2 = st.columns(2)
        
        toko_options = ["Semua"] + sorted(DAFTAR_TOKO)
        toko_selected = col_lk1.selectbox("Kirim Ke (Toko)", toko_options, key="hist_lk_toko")
        if toko_selected != "Semua":
            kirim_ke_filter = toko_selected
        
        via_options = ["Semua"] + sorted(VIA_OPTIONS)
        via_selected = col_lk2.selectbox("Via Pengiriman", via_options, key="hist_lk_via")
        if via_selected != "Semua":
            via_filter = via_selected
    
    elif activity_filter == "ECERAN":
        st.markdown("##### 🛒 Filter Eceran")
        col_ec1, col_ec2 = st.columns(2)
        
        payment_type_options = ["Semua"] + sorted(PAYMENT_TYPES)
        payment_type_selected = col_ec1.selectbox("Tipe Pembayaran", payment_type_options, key="hist_ec_payment_type")
        if payment_type_selected != "Semua":
            payment_type_filter = payment_type_selected
        
        payment_status_options = ["Semua"] + sorted(PAYMENT_STATUS)
        payment_status_selected = col_ec2.selectbox("Status Pembayaran", payment_status_options, key="hist_ec_payment_status")
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
    
    # ========================================
    # LOAD AND FILTER DATA
    # ========================================
    try:
        all_history = get_history(start_date=start_date, end_date=end_date)
        
        filtered_history = []
        for rec in all_history:
            # Activity type filter
            if activity_filter and rec['action_type'] != activity_filter:
                continue
            
            # Global search - search across ALL fields
            if global_search:
                searchable_text = (
                    f"{rec.get('item_name', '')} "
                    f"{rec.get('supplier', '')} "
                    f"{rec.get('category', '')} "
                    f"{rec.get('size', '')} "
                    f"{rec.get('color', '')} "
                    f"{rec.get('location', '')} "
                    f"{rec.get('details', {}).get('nomor_po', '')} "
                    f"{rec.get('details', {}).get('customer', '')} "
                    f"{rec.get('details', {}).get('purpose', '')} "
                    f"{rec.get('details', {}).get('via', '')}"
                ).lower()
                
                if global_search not in searchable_text:
                    continue
            
            # Supplier filter
            if supplier_filter:
                rec_supplier = rec.get('supplier', '')
                if not rec_supplier or rec_supplier.lower() != supplier_filter.lower():
                    continue
            
            # Kategori filter
            if kategori_filter:
                rec_kategori = rec.get('category', '')
                if not rec_kategori or rec_kategori.lower() != kategori_filter.lower():
                    continue
            
            # Location filter
            if location_filter:
                rec_location = rec.get('location', '')
                if not rec_location or rec_location.lower() != location_filter.lower():
                    continue
            
            details = rec.get('details', {})
            
            # Luar Kota specific filters
            if activity_filter == "LUAR_KOTA":
                if kirim_ke_filter:
                    purpose = details.get('purpose', '')
                    if not purpose or purpose.lower() != kirim_ke_filter.lower():
                        continue
                
                if via_filter:
                    via = details.get('via', '')
                    if not via or via.lower() != via_filter.lower():
                        continue
            
            # Eceran specific filters
            if activity_filter == "ECERAN":
                if payment_type_filter:
                    pay_type = details.get('payment_type', '')
                    if not pay_type or pay_type.lower() != payment_type_filter.lower():
                        continue
                
                if payment_status_filter:
                    pay_status = details.get('payment_status', '')
                    if not pay_status or pay_status.lower() != payment_status_filter.lower():
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
        
        # ========================================
        # INJECT CSS FOR HTML TABLES
        # ========================================
        inject_merged_table_css()
        
        # ========================================
        # 1. BARANG MASUK - HTML TABLE WITH TRUE MERGED CELLS
        # ========================================
        st.markdown("### 📥 Barang Masuk")
        masuk_records = [r for r in filtered_history if r['action_type'] == 'MASUK']
        
        if masuk_records:
            # **Excel Download Button**
            excel_masuk = []
            for rec in masuk_records:
                details = rec.get('details', {})
                combined_name = format_item_combined(
                    rec.get('item_name', ''),
                    rec.get('size', ''),
                    rec.get('color', '')
                )
                
                excel_masuk.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Masuk': format_date_dd_mm_yy(rec['transaction_date']),
                    'Nomor PO': details.get('nomor_po', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': combined_name,
                    'Qty': abs(rec['quantity']),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            
            df_excel_masuk = pd.DataFrame(excel_masuk)
            
            create_download_button_excel(
                df_excel_masuk,
                f"barang_masuk_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Barang Masuk",
                "Barang Masuk"
            )
            
            st.markdown("---")
            
            # **HTML Table with TRUE merged cells**
            html_table = render_merged_table_masuk(masuk_records)
            st.markdown(html_table, unsafe_allow_html=True)
            
            st.markdown("---")
            render_inline_editor(masuk_records, "MASUK", "📥")
        else:
            st.info("Tidak ada riwayat barang masuk.")
        
        # ========================================
        # 2. LUAR KOTA - HTML TABLE WITH TRUE MERGED CELLS
        # ========================================
        st.markdown("---")
        st.markdown("### 🚚 Pengeluaran Luar Kota")
        luar_records = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA']
        
        if luar_records:
            # **Excel Download Button**
            excel_luar = []
            for rec in luar_records:
                details = rec.get('details', {})
                combined_name = format_item_combined(
                    rec.get('item_name', ''),
                    rec.get('size', ''),
                    rec.get('color', '')
                )
                
                excel_luar.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Kirim': format_date_dd_mm_yy(rec['transaction_date']),
                    'Tujuan': details.get('purpose', '—'),
                    'Via': details.get('via', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': combined_name,
                    'Qty': abs(rec['quantity']),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            
            df_excel_luar = pd.DataFrame(excel_luar)
            
            create_download_button_excel(
                df_excel_luar,
                f"luar_kota_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Luar Kota",
                "Luar Kota"
            )
            
            st.markdown("---")
            
            # **HTML Table with TRUE merged cells**
            html_table = render_merged_table_luar_kota(luar_records)
            st.markdown(html_table, unsafe_allow_html=True)
            
            st.markdown("---")
            render_inline_editor(luar_records, "LUAR_KOTA", "🚚")
        else:
            st.info("Tidak ada riwayat pengeluaran luar kota.")
        
        # ========================================
        # 3. ECERAN - HTML TABLE WITH TRUE MERGED CELLS
        # ========================================
        st.markdown("---")
        st.markdown("### 🛒 Penjualan Eceran")
        ecer_records = [r for r in filtered_history if r['action_type'] == 'ECERAN']
        
        if ecer_records:
            # **Excel Download Button**
            excel_eceran = []
            for rec in ecer_records:
                details = rec.get('details', {})
                combined_name = format_item_combined(
                    rec.get('item_name', ''),
                    rec.get('size', ''),
                    rec.get('color', '')
                )
                
                excel_eceran.append({
                    'Tanggal Input': format_local_datetime(rec.get('created_at'), "%d/%m/%y %H:%M"),
                    'Tanggal Jual': format_date_dd_mm_yy(rec['transaction_date']),
                    'Pelanggan': details.get('customer', '—'),
                    'Supplier': rec.get('supplier', '—'),
                    'Kategori': rec.get('category', '—').upper(),
                    'Nama Barang': combined_name,
                    'Qty': abs(rec['quantity']),
                    'Harga': details.get('total_price', 0),
                    'Status Bayar': details.get('payment_status', '—'),
                    'Tipe Bayar': details.get('payment_type', '—'),
                    'Rekening': details.get('rekening', '—'),
                    'Lokasi': rec.get('location', '—'),
                    'Komen': details.get('komen', '—')
                })
            
            df_excel_eceran = pd.DataFrame(excel_eceran)
            
            create_download_button_excel(
                df_excel_eceran,
                f"eceran_{start_date.strftime('%d%m%y')}_to_{end_date.strftime('%d%m%y')}.xlsx",
                "📥 Download Excel - Eceran",
                "Penjualan Eceran"
            )
            
            st.markdown("---")
            
            # **HTML Table with TRUE merged cells**
            html_table = render_merged_table_eceran(ecer_records)
            st.markdown(html_table, unsafe_allow_html=True)
            
            st.markdown("---")
            render_inline_editor(ecer_records, "ECERAN", "🛒")
        else:
            st.info("Tidak ada riwayat penjualan eceran.")
    
    except Exception as e:
        st.error(f"❌ Error loading history: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

def render_inline_editor(records: list, action_type: str, icon: str):
    """Helper function to render a contextual editor for each section"""
    with st.expander(f"{icon} Kelola / Edit Transaksi {action_type.replace('_', ' ').title()}", expanded=False):
        record_map = {
            f"No: {i+1} | {r['item_name'].title()} ({r['quantity']}) | {r['uuid'][:8]}": r 
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
                
                # UPDATE BUTTON
                if c_edit.form_submit_button("💾 Simpan Perubahan", use_container_width=True):
                    from models.history import update_history_record
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
                
                # DELETE BUTTON
                if c_del.form_submit_button("🗑️ Hapus Transaksi", type="secondary", use_container_width=True):
                    # Logic to reverse stock
                    if rec['action_type'] == "MASUK" and not is_pabrik_source(rec.get('location', '')):
                        try:
                            reduce_inventory_quantity(
                                name=rec['item_name'], 
                                category=rec.get('category'), 
                                color=rec.get('color'), 
                                size=rec.get('size'), 
                                location=rec.get('location'), 
                                quantity=abs(rec['quantity'])
                            )
                        except:
                            pass
                    elif rec['action_type'] in ["LUAR_KOTA", "ECERAN"] and not is_pabrik_source(rec.get('location', '')):
                        try:
                            add_inventory_item(
                                name=rec['item_name'], 
                                category=rec.get('category'), 
                                color=rec.get('color'), 
                                size=rec.get('size'), 
                                quantity=abs(rec['quantity']), 
                                location=rec.get('location'), 
                                supplier=rec.get('supplier'), 
                                date_added=date.today()
                            )
                        except:
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
# SUMMARY TAB
# ==================================================

# This file contains ONLY the updated render_summary_tab function
# Replace the existing function in your app_helpers.py with this version

# This file contains ONLY the updated render_summary_tab function
# Replace the existing function in your app_helpers.py with this version

# This file contains ONLY the updated render_summary_tab function
# Replace the existing function in your app_helpers.py with this version

# This file contains ONLY the updated render_summary_tab function
# Replace the existing function in your app_helpers.py with this version

def render_summary_tab(today: date):
    
    # ========================================
    # PASSWORD PROTECTION - FIRST THING
    # ========================================
    if not st.session_state.get('summary_unlocked', False):
        st.warning("🔒 Ringkasan terkunci - masukkan password untuk akses")
        
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

    # Date range selection
    col1, col2 = st.columns(2)
    def_start = today - timedelta(days=90)
    sum_start = col1.date_input("Dari", def_start, key="sum_from", format="DD/MM/YYYY")
    sum_end = col2.date_input("Sampai", today, min_value=sum_start, key="sum_to", format="DD/MM/YYYY")

    # MULTI-SELECT FILTERS
    st.markdown("#### 🔎 Filter Data")
    col_f1, col_f2, col_f3 = st.columns(3)

    activity_filter = col_f1.multiselect(
        "Jenis Aktivitas", 
        ["MASUK", "LUAR_KOTA", "ECERAN"],
        key="sum_activity_multi"
    )

    all_suppliers = st.session_state.get("dynamic_supplier", SUPPLIERS)
    supplier_filter = col_f2.multiselect(
        "Supplier", 
        all_suppliers,
        key="sum_supplier_multi"
    )

    all_kategori = st.session_state.get("dynamic_kategori", STANDARD_CATEGORIES)
    kategori_filter = col_f3.multiselect(
        "Kategori", 
        all_kategori,
        key="sum_category_multi"
    )

    try:
        # Get all history with date filter
        all_history = get_cached_history(start_date=sum_start, end_date=sum_end, item_name=None)
        
        # Apply multi-select filters
        filtered_history = []
        for rec in all_history:
            # Activity type filter
            if activity_filter and rec['action_type'] not in activity_filter:
                continue
            
            # Supplier filter
            if supplier_filter and rec.get('supplier') not in supplier_filter:
                continue
            
            # Category filter
            if kategori_filter and rec.get('category') not in kategori_filter:
                continue
            
            filtered_history.append(rec)
        
        if not filtered_history:
            st.info("🔭 Tidak ada data yang sesuai dengan filter")
            return

        st.markdown("---")
        st.success(f"✅ Menampilkan {len(filtered_history)} transaksi")

        # ========================================
        # SECTION 1: TOP 10 TABLES (COLLAPSIBLE, EXCLUDE BONUS)
        # ========================================
        st.markdown("### 📊 Top 10 - Berdasarkan Filter (Bonus Dikecualikan)")
        
        # Separate data by type and EXCLUDE BONUS
        masuk_data = [r for r in filtered_history if r['action_type'] == 'MASUK' and r.get('category', '').lower() != 'bonus']
        luar_data = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA' and r.get('category', '').lower() != 'bonus']
        eceran_data = [r for r in filtered_history if r['action_type'] == 'ECERAN' and r.get('category', '').lower() != 'bonus']
        
        # Top 10: Masuk Barang (COLLAPSIBLE)
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
        
        # Top 10: Luar Kota (COLLAPSIBLE)
        if luar_data:
            with st.expander("🚚 Top 10 Barang Luar Kota", expanded=False):
                luar_items = {}
                for r in luar_data:
                    item = r['item_name'].title()
                    luar_items[item] = luar_items.get(item, 0) + abs(r['quantity'])
                
                top_luar = sorted(luar_items.items(), key=lambda x: x[1], reverse=True)[:10]
                luar_df = pd.DataFrame(top_luar, columns=['Barang', 'Total Qty'])
                luar_df.index = range(1, len(luar_df) + 1)
                st.dataframe(luar_df, use_container_width=True)
        
        # Top 10: Eceran (COLLAPSIBLE)
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
            st.warning("⚠️ Install plotly untuk visualisasi: `pip install plotly`")
            return
        
        # ========================================
        # CUSTOM CSS FOR CHARTS
        # ========================================
        st.markdown("""
        <style>
            .chart-label {
                color: #FFFFFF;
                font-weight: 600;
                font-size: 14px;
            }
            .chart-legend {
                color: #FFFFFF;
                font-size: 13px;
            }
        </style>
        """, unsafe_allow_html=True)
        
        # ========================================
        # DARK MODE COLOR PALETTE
        # ========================================
        
        # 15-color palette for dark background
        DARK_MODE_PALETTE = [
            "#29B6F6",  # Sky Blue
            "#66BB6A",  # Emerald Green
            "#FFCA28",  # Amber
            "#FF7043",  # Coral
            "#AB47BC",  # Violet
            "#26C6DA",  # Aqua
            "#D4E157",  # Lime
            "#EC407A",  # Magenta
            "#00897B",  # Teal
            "#FFA726",  # Bright Orange
            "#5C6BC0",  # Indigo
            "#F06292",  # Pink
            "#00FFFF",  # Cyan
            "#FFD700",  # Gold
            "#DC143C",  # Crimson
        ]
        
        # ========================================
        # PIE CHARTS - DARK MODE PALETTE
        # ========================================
        
        st.markdown("### 📊 Distribusi Data")
        
        # Pie Chart: Distribution by Activity Type (EXCLUDE BONUS) - ALWAYS SHOW
        if filtered_history:
            st.markdown("#### 📊 Distribusi Berdasarkan Jenis Aktivitas")
            non_bonus_history = [r for r in filtered_history if r.get('category', '').lower() != 'bonus']
            
            activity_counts = {}
            for r in non_bonus_history:
                activity_counts[r['action_type']] = activity_counts.get(r['action_type'], 0) + abs(r['quantity'])
            
            activity_df = pd.DataFrame([
                {'Aktivitas': k, 'Jumlah': v} for k, v in activity_counts.items()
            ])
            
            if not activity_df.empty:
                fig = px.pie(
                    activity_df,
                    values='Jumlah',
                    names='Aktivitas',
                    title='Distribusi Berdasarkan Jenis Aktivitas (Tanpa Bonus)',
                    hole=0.35,
                    color_discrete_sequence=DARK_MODE_PALETTE[:3]  # Use first 3 colors
                )
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    texttemplate="%{label}<br>%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')  # Pure white text
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(
                        orientation="h", 
                        yanchor="bottom", 
                        y=-0.15, 
                        xanchor="center", 
                        x=0.5,
                        font=dict(color='#FFFFFF', size=13, family='Arial')
                    ),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # ONLY SHOW DETAILED PIE CHARTS IF ANY FILTER IS SELECTED
        if activity_filter or supplier_filter or kategori_filter:
            # Pie Chart: Distribution by Supplier (EXCLUDE BONUS)
            if filtered_history:
                st.markdown("#### 🏭 Distribusi Berdasarkan Supplier")
                non_bonus_history = [r for r in filtered_history if r.get('category', '').lower() != 'bonus']
            
                supplier_counts = {}
                for r in non_bonus_history:
                    supplier = r.get('supplier', 'Unknown')
                    supplier_counts[supplier] = supplier_counts.get(supplier, 0) + abs(r['quantity'])
            
            # Limit to top 12, group others
            supplier_sorted = sorted(supplier_counts.items(), key=lambda x: x[1], reverse=True)
            if len(supplier_sorted) > 12:
                top_12 = dict(supplier_sorted[:12])
                others = sum([v for k, v in supplier_sorted[12:]])
                if others > 0:
                    top_12['Lainnya'] = others
                supplier_counts = top_12
            
            supplier_df = pd.DataFrame([
                {'Supplier': k, 'Jumlah': v} for k, v in supplier_counts.items()
            ])
            
            if not supplier_df.empty:
                fig = px.pie(
                    supplier_df,
                    values='Jumlah',
                    names='Supplier',
                    title='Distribusi Berdasarkan Supplier (Tanpa Bonus)',
                    hole=0.35,
                    color_discrete_sequence=DARK_MODE_PALETTE
                )
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    texttemplate="%{label}<br>%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')  # Pure white text
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(
                        orientation="h", 
                        yanchor="bottom", 
                        y=-0.15, 
                        xanchor="center", 
                        x=0.5,
                        font=dict(color='#FFFFFF', size=13, family='Arial')
                    ),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)
    # Pie Chart: Distribution by Kirim Ke Toko (LUAR KOTA ONLY)
        if activity_filter or supplier_filter or kategori_filter:
            luar_kota_records = [r for r in filtered_history if r['action_type'] == 'LUAR_KOTA']
            if luar_kota_records:
                st.markdown("#### 🏪 Distribusi Kirim Ke Toko (Luar Kota)")
                st.caption("📌 Distribusi pengiriman berdasarkan toko tujuan")
                
                toko_counts = {}  # ✅ FIXED: Properly indented inside the if block
                for r in luar_kota_records:
                    details = r.get('details', {})
                    
                    # Try multiple field names for backwards compatibility
                    toko = None
                    if isinstance(details, dict):
                        toko = details.get('purpose') or details.get('kirim_ke') or details.get('tujuan')
                    
                    # If still no toko found, try to extract from other fields
                    if not toko or toko == '':
                        # Check if there's a 'via' field that might contain toko info
                        via = details.get('via', '') if isinstance(details, dict) else ''
                        
                        # If nothing found, mark as "Data Lama (Toko Tidak Tercatat)"
                        toko = 'Data Lama (Toko Tidak Tercatat)'
                    
                    toko_counts[toko] = toko_counts.get(toko, 0) + abs(r['quantity'])
                
                # Limit to top 12, group others
                toko_sorted = sorted(toko_counts.items(), key=lambda x: x[1], reverse=True)
                if len(toko_sorted) > 12:
                    top_12 = dict(toko_sorted[:12])
                    others = sum([v for k, v in toko_sorted[12:]])
                    if others > 0:
                        top_12['Lainnya'] = others
                    toko_counts = top_12
                
                toko_df = pd.DataFrame([
                    {'Toko': k, 'Jumlah': v} for k, v in toko_counts.items()
                ])
                
                if not toko_df.empty:
                    fig = px.pie(
                        toko_df,
                        values='Jumlah',
                        names='Toko',
                        title='Distribusi Pengiriman per Toko Tujuan',
                        hole=0.35,
                        color_discrete_sequence=DARK_MODE_PALETTE
                    )
                    fig.update_traces(
                        textposition="inside",
                        textinfo="percent+label",
                        texttemplate="%{label}<br>%{percent:.1%}",
                        hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                        textfont=dict(color='#FFFFFF', size=14, family='Arial')
                    )
                    fig.update_layout(
                        showlegend=True,
                        legend=dict(
                            orientation="h", 
                            yanchor="bottom", 
                            y=-0.15, 
                            xanchor="center", 
                            x=0.5,
                            font=dict(color='#FFFFFF', size=13, family='Arial')
                        ),
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#FFFFFF', size=14, family='Arial'),
                        title=dict(font=dict(color='#FFFFFF', size=16))
                    )
                    st.plotly_chart(fig, use_container_width=True)
        
        # Pie Chart: Distribution by Category (DETAILED - INCLUDES BONUS)
        if (activity_filter or supplier_filter or kategori_filter) and filtered_history:
            st.markdown("#### 🏷️ Distribusi Berdasarkan Kategori (Detail)")
            st.caption("📌 Termasuk Bonus | Format: Kategori - Nama Barang - Ukuran - Warna")
            
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
            
            # Limit to top 15, group others
            category_sorted = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
            if len(category_sorted) > 15:
                top_15 = dict(category_sorted[:15])
                others = sum([v for k, v in category_sorted[15:]])
                if others > 0:
                    top_15['Lainnya'] = others
                category_counts = top_15
            
            category_df = pd.DataFrame([
                {'Kategori': k, 'Jumlah': v} for k, v in category_counts.items()
            ])
            
            if not category_df.empty:
                fig = px.pie(
                    category_df,
                    values='Jumlah',
                    names='Kategori',
                    title='Distribusi Berdasarkan Kategori (Detail)',
                    hole=0.35,
                    color_discrete_sequence=DARK_MODE_PALETTE
                )
                fig.update_traces(
                    textposition="inside",
                    textinfo="percent",
                    texttemplate="%{percent:.1%}",
                    hovertemplate="<b>%{label}</b><br>Jumlah: %{value:,}<br>Persentase: %{percent:.1%}",
                    textfont=dict(color='#FFFFFF', size=14, family='Arial')  # Pure white text
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(
                        orientation="v",
                        yanchor="top",
                        y=1,
                        xanchor="left",
                        x=1.02,
                        font=dict(color='#FFFFFF', size=13, family='Arial')
                    ),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#FFFFFF', size=14, family='Arial'),
                    title=dict(font=dict(color='#FFFFFF', size=16))
                )
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        
        # ========================================
        # MONTHLY STACKED BAR CHARTS - DARK MODE PALETTE
        # ========================================
        
        st.markdown("---")
        
        # ========================================
        # MONTHLY STACKED BAR CHARTS - DARK MODE PALETTE
        # ========================================
        
        st.markdown("### 📊 Analisis Bulanan Per Supplier (Bonus Dikecualikan)")
        st.caption("📌 Stacked bar charts showing monthly trends by supplier for each transaction type")
        
        # **NEW: Month-Year Range Filter for Bar Charts**
        col_date1, col_date2 = st.columns(2)
        
        current_date = datetime.now().date()
        
       # **NEW: Month-Year Range Filter for Bar Charts**
        st.markdown("**📅 Pilih Rentang Bulan untuk Chart:**")
        
        col_month1, col_year1, col_month2, col_year2 = st.columns(4)
        
        current_date = datetime.now()
        current_year = current_date.year
        
        # Month names in Indonesian
        months_id = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]
        
        # Available years (2020 to current year + 1)
        available_years = list(range(2020, current_year + 2))
        
        # Default: from 12 months ago
        default_start = current_date - timedelta(days=365)
        default_end = current_date
        
        # FROM Month/Year
        start_month_idx = col_month1.selectbox(
            "Dari Bulan",
            range(12),
            format_func=lambda x: months_id[x],
            index=default_start.month - 1,
            key="chart_start_month"
        )
        
        start_year = col_year1.selectbox(
            "Tahun",
            available_years,
            index=available_years.index(default_start.year) if default_start.year in available_years else len(available_years) - 2,
            key="chart_start_year"
        )
        
        # TO Month/Year
        end_month_idx = col_month2.selectbox(
            "Sampai Bulan",
            range(12),
            format_func=lambda x: months_id[x],
            index=default_end.month - 1,
            key="chart_end_month"
        )
        
        end_year = col_year2.selectbox(
            "Tahun",
            available_years,
            index=available_years.index(default_end.year) if default_end.year in available_years else len(available_years) - 1,
            key="chart_end_year"
        )
        
        # Convert to date objects for comparison
        start_month = date(start_year, start_month_idx + 1, 1)
        end_month = date(end_year, end_month_idx + 1, 1)
        
        # Validate date range
        if start_month > end_month:
            st.error("⚠️ 'Dari Bulan' harus lebih awal dari 'Sampai Bulan'")
            st.stop()
        
        st.caption(f"📊 Menampilkan data dari **{months_id[start_month_idx]} {start_year}** sampai **{months_id[end_month_idx]} {end_year}**")
        
        # Get ALL history data and EXCLUDE BONUS
        all_data = [r for r in get_cached_history() if r.get('category', '').lower() != 'bonus']
        
        # **NEW: Filter by month range**
        month_filtered_data = []
        for r in all_data:
            trans_date = r['transaction_date']
            if isinstance(trans_date, str):
                trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
            
            # Compare year-month only
            trans_month = trans_date.replace(day=1)
            start_filter = start_month.replace(day=1)
            end_filter = end_month.replace(day=1)
            
            if start_filter <= trans_month <= end_filter:
                month_filtered_data.append(r)
        
        all_data = month_filtered_data
        
        if all_data:
            months_order = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ]
            
            # Helper function to create stacked bar chart with dark mode colors
            def create_supplier_stacked_chart(data_filtered, title, color_palette):
                """Create stacked bar chart for supplier analysis - shows ALL months in selected range"""
                monthly_supplier_data = {}
                suppliers_set = set()
                
                # Collect data from filtered records
                for r in data_filtered:
                    trans_date = r['transaction_date']
                    if isinstance(trans_date, str):
                        trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
                    
                    # Include year in month key
                    month_year = trans_date.strftime('%Y-%m')  # e.g., "2024-01"
                    supplier = r.get('supplier', 'Unknown')
                    
                    if not supplier or supplier.strip() == '':
                        supplier = 'Unknown'
                    
                    suppliers_set.add(supplier)
                    
                    key = (month_year, supplier)
                    if key not in monthly_supplier_data:
                        monthly_supplier_data[key] = 0
                    
                    monthly_supplier_data[key] += abs(r['quantity'])
                
                # **NEW: Generate ALL months in the selected range**
                all_months = []
                current = start_month.replace(day=1)
                end = end_month.replace(day=1)
                
                while current <= end:
                    month_year = current.strftime('%Y-%m')
                    month_display = current.strftime('%b %Y')  # e.g., "Jan 2024"
                    all_months.append((month_year, month_display))
                    
                    # Move to next month
                    if current.month == 12:
                        current = current.replace(year=current.year + 1, month=1)
                    else:
                        current = current.replace(month=current.month + 1)
                
                # Create complete dataset with ALL months
                plot_data = []
                
                if suppliers_set:
                    # For each month in range, create entries for all suppliers (even if 0)
                    for month_year, month_display in all_months:
                        for supplier in suppliers_set:
                            qty = monthly_supplier_data.get((month_year, supplier), 0)
                            plot_data.append({
                                'Bulan': month_display,
                                'Supplier': supplier,
                                'Jumlah': qty,
                                'sort_key': month_year  # For proper sorting
                            })
                else:
                    # No data at all
                    for month_year, month_display in all_months:
                        plot_data.append({
                            'Bulan': month_display,
                            'Supplier': 'No Data',
                            'Jumlah': 0,
                            'sort_key': month_year
                        })
                
                if plot_data:
                    chart_df = pd.DataFrame(plot_data)
                    
                    # Sort by the actual date (sort_key)
                    chart_df = chart_df.sort_values('sort_key')
                    
                    # Get unique months in order
                    month_order = [m[1] for m in all_months]
                    
                    fig = px.bar(
                        chart_df,
                        x='Bulan',
                        y='Jumlah',
                        color='Supplier',
                        barmode='stack',
                        title=title,
                        labels={'Jumlah': 'Total Quantity', 'Bulan': 'Month'},
                        text='Jumlah',
                        color_discrete_sequence=color_palette,
                        category_orders={'Bulan': month_order}  # Force correct order
                    )
                    
                    fig.update_traces(
                        texttemplate='%{text}',
                        textposition='inside',
                        textfont=dict(color='#FFFFFF', size=12, family='Arial')
                    )
                    
                    fig.update_layout(
                        xaxis_title='Bulan',
                        yaxis_title='Total Quantity',
                        height=500,
                        legend=dict(
                            orientation="v",
                            yanchor="top",
                            y=1,
                            xanchor="left",
                            x=1.02,
                            title="Supplier",
                            font=dict(color='#FFFFFF', size=13, family='Arial')
                        ),
                        xaxis={
                            'categoryorder': 'array',
                            'categoryarray': month_order,
                            'tickangle': -45  # Angle labels for better readability
                        },
                        bargap=0.15,
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        font=dict(color='#FFFFFF', size=14, family='Arial'),
                        title=dict(font=dict(color='#FFFFFF', size=16)),
                        xaxis_gridcolor='rgba(128,128,128,0.2)',
                        yaxis_gridcolor='rgba(128,128,128,0.2)'
                    )
                    
                    return fig
                return None
            
            # 1. MASUK BARANG - Using dark mode palette
            st.markdown("#### 📥 Barang Masuk per Supplier (Bulanan)")
            masuk_data = [r for r in all_data if r['action_type'] == 'MASUK']
            fig_masuk = create_supplier_stacked_chart(
                masuk_data,
                'Barang Masuk - Breakdown per Supplier',
                DARK_MODE_PALETTE
            )
            if fig_masuk:
                st.plotly_chart(fig_masuk, use_container_width=True)
            
            # 2. LUAR KOTA - Using dark mode palette
            st.markdown("#### 🚚 Pengiriman Luar Kota per Supplier (Bulanan)")
            luar_kota_data = [r for r in all_data if r['action_type'] == 'LUAR_KOTA']
            fig_luar = create_supplier_stacked_chart(
                luar_kota_data,
                'Luar Kota - Breakdown per Supplier',
                DARK_MODE_PALETTE
            )
            if fig_luar:
                st.plotly_chart(fig_luar, use_container_width=True)
            
            # 3. ECERAN - Using dark mode palette
            st.markdown("#### 🛒 Penjualan Eceran per Sumber (Bulanan)")
            eceran_data = [r for r in all_data if r['action_type'] == 'ECERAN']
            for r in eceran_data:
                if not r.get('supplier'):
                    r['supplier'] = r.get('location', 'Unknown')
            
            fig_eceran = create_supplier_stacked_chart(
                eceran_data,
                'Penjualan Eceran - Breakdown per Sumber',
                DARK_MODE_PALETTE
            )
            if fig_eceran:
                st.plotly_chart(fig_eceran, use_container_width=True)
        else:
            st.info("📭 Tidak ada data untuk ditampilkan")
            st.markdown("---")
        
        st.markdown("---")
        
      # ========================================
        # MONTHLY METRICS WITH GROWTH CALCULATION (COLLAPSIBLE)
        # ========================================
        
        st.markdown("### 📊 Persentase Penjualan Bulanan")
        st.caption("📌 Pertumbuhan bulanan dibandingkan dengan bulan sebelumnya (Bonus Dikecualikan)")
        
        def calculate_monthly_metrics_detailed(data_filtered, metric_type="quantity"):
            """Calculate monthly totals with supplier/delivery breakdown"""
            monthly_data = {}
            
            for r in data_filtered:
                trans_date = r['transaction_date']
                if isinstance(trans_date, str):
                    trans_date = datetime.strptime(trans_date[:10], '%Y-%m-%d').date()
                
                month_key = trans_date.strftime('%Y-%m')
                month_name = trans_date.strftime('%B %Y')
                
                if month_key not in monthly_data:
                    monthly_data[month_key] = {
                        'name': month_name,
                        'total': 0,
                        'suppliers': {},
                        'deliveries': {},
                        'kirim_toko': {}
                    }
                
                # Get supplier info
                supplier = r.get('supplier', 'Unknown')
                details = r.get('details', {})
                
                if metric_type == "quantity":
                    qty = abs(r['quantity'])
                    monthly_data[month_key]['total'] += qty
                    
                    # Track by supplier
                    if supplier not in monthly_data[month_key]['suppliers']:
                        monthly_data[month_key]['suppliers'][supplier] = 0
                    monthly_data[month_key]['suppliers'][supplier] += qty
                    
                elif metric_type == "sales":
                    sales = details.get('total_price', 0)
                    monthly_data[month_key]['total'] += sales
                    
                    # Track by supplier
                    if supplier not in monthly_data[month_key]['suppliers']:
                        monthly_data[month_key]['suppliers'][supplier] = 0
                    monthly_data[month_key]['suppliers'][supplier] += sales
                
                elif metric_type == "quantity_supplier":
                    # Eceran - only supplier breakdown
                    qty = abs(r['quantity'])
                    monthly_data[month_key]['total'] += qty
                    
                    if supplier not in monthly_data[month_key]['suppliers']:
                        monthly_data[month_key]['suppliers'][supplier] = 0
                    monthly_data[month_key]['suppliers'][supplier] += qty
                
                elif metric_type == "sales_supplier":
                    # Eceran sales - only supplier breakdown
                    sales = details.get('total_price', 0)
                    monthly_data[month_key]['total'] += sales
                    
                    if supplier not in monthly_data[month_key]['suppliers']:
                        monthly_data[month_key]['suppliers'][supplier] = 0
                    monthly_data[month_key]['suppliers'][supplier] += sales
                
                elif metric_type == "quantity_luar_kota":
                    # Luar kota - supplier and kirim ke toko
                    qty = abs(r['quantity'])
                    monthly_data[month_key]['total'] += qty
                    
                    # Track by supplier
                    if supplier not in monthly_data[month_key]['suppliers']:
                        monthly_data[month_key]['suppliers'][supplier] = 0
                    monthly_data[month_key]['suppliers'][supplier] += qty
                    
                    # Track by kirim ke toko - ROBUST HANDLING
                    kirim_ke = None
                    if isinstance(details, dict):
                        kirim_ke = details.get('purpose') or details.get('kirim_ke') or details.get('tujuan')
                    
                    if not kirim_ke or kirim_ke == '':
                        kirim_ke = 'Data Lama (Toko Tidak Tercatat)'
                    
                    if kirim_ke not in monthly_data[month_key]['kirim_toko']:
                        monthly_data[month_key]['kirim_toko'][kirim_ke] = 0
                    monthly_data[month_key]['kirim_toko'][kirim_ke] += qty
            
            # Sort by date
            sorted_months = sorted(monthly_data.keys())
            
            # Calculate growth and percentages
            metrics = []
            for i, month_key in enumerate(sorted_months):
                current = monthly_data[month_key]['total']
                
                # Calculate growth
                if i > 0:
                    prev_month_key = sorted_months[i - 1]
                    previous = monthly_data[prev_month_key]['total']
                    
                    if previous > 0:
                        growth_pct = ((current - previous) / previous) * 100
                        growth_str = f'{growth_pct:,.1f}%'
                        growth_value = current - previous
                        delta_color = 'normal'
                    else:
                        growth_str = 'n/a'
                        growth_value = 0
                        delta_color = 'off'
                else:
                    growth_str = 'n/a'
                    growth_value = 0
                    delta_color = 'off'
                
                # Calculate supplier percentages and growth
                supplier_breakdown = []
                prev_suppliers = monthly_data[sorted_months[i-1]]['suppliers'] if i > 0 else {}
                
                for supplier, value in sorted(monthly_data[month_key]['suppliers'].items(), key=lambda x: x[1], reverse=True):
                    pct = (value / current * 100) if current > 0 else 0
                    
                    # Calculate supplier growth
                    if i == 0:
                        # First month - no previous data
                        supplier_growth_str = 'n/a'
                        supplier_growth_value = 0
                    else:
                        prev_value = prev_suppliers.get(supplier, 0)
                        if prev_value > 0:
                            supplier_growth_pct = ((value - prev_value) / prev_value) * 100
                            # Cap display at ±999.9% for readability
                            if supplier_growth_pct > 999.9:
                                supplier_growth_str = '>999.9%'
                            elif supplier_growth_pct < -999.9:
                                supplier_growth_str = '<-999.9%'
                            else:
                                supplier_growth_str = f'{supplier_growth_pct:,.1f}%'
                            supplier_growth_value = value - prev_value
                        elif prev_value == 0 and value > 0:
                            # New supplier or supplier that had 0 last month
                            supplier_growth_str = 'New'
                            supplier_growth_value = value
                        else:
                            supplier_growth_str = 'n/a'
                            supplier_growth_value = 0
                    
                    supplier_breakdown.append({
                        'name': supplier,
                        'value': value,
                        'percentage': pct,
                        'growth_str': supplier_growth_str,
                        'growth_value': supplier_growth_value
                    })
                
               # Calculate kirim toko percentages and growth
                kirim_toko_breakdown = []
                prev_kirim_toko = monthly_data[sorted_months[i-1]]['kirim_toko'] if i > 0 else {}
                
                for toko, value in sorted(monthly_data[month_key]['kirim_toko'].items(), key=lambda x: x[1], reverse=True):
                    pct = (value / current * 100) if current > 0 else 0
                    
                    # Calculate toko growth
                    if i == 0:
                        # First month - no previous data
                        toko_growth_str = 'n/a'
                        toko_growth_value = 0
                    else:
                        prev_toko_value = prev_kirim_toko.get(toko, 0)
                        if prev_toko_value > 0:
                            toko_growth_pct = ((value - prev_toko_value) / prev_toko_value) * 100
                            # Cap display at ±999.9% for readability
                            if toko_growth_pct > 999.9:
                                toko_growth_str = '>999.9%'
                            elif toko_growth_pct < -999.9:
                                toko_growth_str = '<-999.9%'
                            else:
                                toko_growth_str = f'{toko_growth_pct:,.1f}%'
                            toko_growth_value = value - prev_toko_value
                        elif prev_toko_value == 0 and value > 0:
                            # New toko destination or had 0 last month
                            toko_growth_str = 'New'
                            toko_growth_value = value
                        else:
                            toko_growth_str = 'n/a'
                            toko_growth_value = 0
                    
                    kirim_toko_breakdown.append({
                        'name': toko,
                        'value': value,
                        'percentage': pct,
                        'growth_str': toko_growth_str,
                        'growth_value': toko_growth_value
                    })
                
                metrics.append({
                    'month_key': month_key,
                    'month_name': monthly_data[month_key]['name'],
                    'total': current,
                    'growth': growth_str,
                    'growth_value': growth_value,
                    'delta_color': delta_color,
                    'suppliers': supplier_breakdown,
                    'kirim_toko': kirim_toko_breakdown
                })
            
            return metrics
        
        # ========================================
        # 1. BARANG MASUK METRICS (COLLAPSIBLE)
        # ========================================
        st.markdown("#### 📥 Barang Masuk (Bulanan)")
        if masuk_data:
            masuk_metrics = calculate_monthly_metrics_detailed(masuk_data, "quantity")
            
            if masuk_metrics:
                # Show only last 12 months
                recent_masuk = masuk_metrics[-12:]
                
                for metric in recent_masuk:
                    # Only show month name in header
                    expander_label = f"**{metric['month_name']}**"
                    with st.expander(expander_label, expanded=False):
                        # Show total and growth INSIDE the expander
                        col_summary1, col_summary2 = st.columns(2)
                        col_summary1.metric(
                            label="📊 Total Unit",
                            value=f"{metric['total']:,} unit",
                            delta=metric['growth']
                        )
                        
                        st.markdown("---")
                        
                        # Supplier breakdown in cards
                        st.markdown("**📦 Breakdown per Supplier:**")
                        
                        # Calculate number of columns (max 4)
                        num_suppliers = len(metric['suppliers'])
                        num_cols = min(num_suppliers, 4)
                        
                        if num_suppliers > 0:
                            cols = st.columns(num_cols)
                            for idx, supplier_info in enumerate(metric['suppliers']):
                                col = cols[idx % num_cols]
                                with col:
                                    st.metric(
                                        label=f"{supplier_info['name']}",
                                        value=f"{supplier_info['value']:,} unit",
                                        delta=supplier_info['growth_str']  # Changed from percentage to growth_str
                                    )
            else:
                st.info("Tidak ada data")
        else:
            st.info("Tidak ada data")
        
        st.markdown("---")
        
        # ========================================
        # 2. LUAR KOTA METRICS (COLLAPSIBLE)
        # ========================================
        st.markdown("#### 🚚 Luar Kota (Bulanan)")
        if luar_kota_data:
            luar_metrics = calculate_monthly_metrics_detailed(luar_kota_data, "quantity_luar_kota")
            
            if luar_metrics:
                # Show only last 12 months
                recent_luar = luar_metrics[-12:]
                
                for metric in recent_luar:
                    # Only show month name in header
                    expander_label = f"**{metric['month_name']}**"
                    with st.expander(expander_label, expanded=False):
                        # Show total and growth INSIDE the expander
                        col_summary1, col_summary2 = st.columns(2)
                        col_summary1.metric(
                            label="📊 Total Unit",
                            value=f"{metric['total']:,} unit",
                            delta=metric['growth']
                        )
                        
                        st.markdown("---")
                        
                        # Supplier breakdown
                        st.markdown("**📦 Per Supplier:**")
                        num_suppliers = len(metric['suppliers'])
                        if num_suppliers > 0:
                            cols = st.columns(min(num_suppliers, 4))
                            for idx, supplier_info in enumerate(metric['suppliers']):
                                col = cols[idx % min(num_suppliers, 4)]
                                with col:
                                    st.metric(
                                        label=f"{supplier_info['name']}",
                                        value=f"{supplier_info['value']:,} unit",
                                        delta=supplier_info['growth_str']  # Changed from percentage to growth_str
                                    )
                        
                        st.markdown("---")
                        
                        # Kirim ke toko breakdown
                        st.markdown("**🏪 Kirim ke Toko:**")
                        num_toko = len(metric['kirim_toko'])
                        if num_toko > 0:
                            cols = st.columns(min(num_toko, 4))
                            for idx, toko_info in enumerate(metric['kirim_toko']):
                                col = cols[idx % min(num_toko, 4)]
                                with col:
                                    st.metric(
                                        label=f"{toko_info['name']}",
                                        value=f"{toko_info['value']:,} unit",
                                        delta=toko_info['growth_str']  # Changed to use growth_str
                                    )
        st.markdown("---")
        
        # ========================================
        # 3. ECERAN QUANTITY METRICS (COLLAPSIBLE)
        # ========================================
        st.markdown("#### 🛒 Eceran - Jumlah Barang (Bulanan)")
        if eceran_data:
            eceran_qty_metrics = calculate_monthly_metrics_detailed(eceran_data, "quantity_supplier")
            
            if eceran_qty_metrics:
                # Show only last 12 months
                recent_eceran_qty = eceran_qty_metrics[-12:]
                
                for metric in recent_eceran_qty:
                    # Only show month name in header
                    expander_label = f"**{metric['month_name']}**"
                    with st.expander(expander_label, expanded=False):
                        # Show total and growth INSIDE the expander
                        col_summary1, col_summary2 = st.columns(2)
                        col_summary1.metric(
                            label="📊 Total Unit",
                            value=f"{metric['total']:,} unit",
                            delta=metric['growth']
                        )
                        
                        st.markdown("---")
                        
                        # Supplier breakdown with growth
                        st.markdown("**📦 Per Supplier:**")
                        num_suppliers = len(metric['suppliers'])
                        if num_suppliers > 0:
                            cols = st.columns(min(num_suppliers, 4))
                            for idx, supplier_info in enumerate(metric['suppliers']):
                                col = cols[idx % min(num_suppliers, 4)]
                                with col:
                                    # Show growth in quantity compared to last month
                                    delta_display = supplier_info['growth_str']
                                    st.metric(
                                        label=f"{supplier_info['name']}",
                                        value=f"{supplier_info['value']:,} unit",
                                        delta=delta_display
                                    )
            else:
                st.info("Tidak ada data")
        else:
            st.info("Tidak ada data")
        
        st.markdown("---")
        
        # ========================================
        # 4. ECERAN SALES VALUE METRICS (COLLAPSIBLE)
        # ========================================
        st.markdown("#### 💰 Eceran - Total Penjualan (Bulanan)")
        if eceran_data:
            eceran_sales_metrics = calculate_monthly_metrics_detailed(eceran_data, "sales_supplier")
            
            if eceran_sales_metrics:
                # Show only last 12 months
                recent_eceran_sales = eceran_sales_metrics[-12:]
                
                for metric in recent_eceran_sales:
                    # Only show month name in header
                    expander_label = f"**{metric['month_name']}**"
                    with st.expander(expander_label, expanded=False):
                        # Show total and growth INSIDE the expander
                        col_summary1, col_summary2 = st.columns(2)
                        col_summary1.metric(
                            label="💰 Total Penjualan",
                            value=format_currency(metric['total']),
                            delta=metric['growth']
                        )
                        
                        st.markdown("---")
                        
                        # Supplier breakdown with growth in Rp
                        st.markdown("**📦 Per Supplier:**")
                        num_suppliers = len(metric['suppliers'])
                        if num_suppliers > 0:
                            cols = st.columns(min(num_suppliers, 4))
                            for idx, supplier_info in enumerate(metric['suppliers']):
                                col = cols[idx % min(num_suppliers, 4)]
                                with col:
                                    # Show growth in Rp compared to last month
                                    delta_display = supplier_info['growth_str']
                                    st.metric(
                                        label=f"{supplier_info['name']}",
                                        value=format_currency(supplier_info['value']),
                                        delta=delta_display
                                    )
            else:
                st.info("Tidak ada data")
        else:
            st.info("Tidak ada data")
    
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
