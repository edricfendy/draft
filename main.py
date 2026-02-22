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
