import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

# --- GÖRSEL AYARLAR (Sade ve Kurumsal) ---
st.set_page_config(page_title="Hakediş Yönetim Sistemi", layout="centered")

st.title("Hakediş Yönetim Sistemi")
st.markdown("Şablon hakediş dosyasını, park ve bulvar metrajlarınıza göre otomatik olarak filtreleyip güncelleyin.")
st.markdown("---")

class HakedisEngine:
    def __init__(self, park_file, bulvar_file):
        self.col_map = {
            'çim': 'ÇİM  (m2)', 'çalı': 'ÇALI (m2)', 'çiçek': 'ÇİÇEK (m2)',
            'sert': 'SERT (m2)', 'çoa': 'ÇOA (m2)', 'ağaçlık': 'AĞAÇLIK (m2)',
            'spor': 'SPOR(m2)', 'toprak': 'TOPRAK (m2)', 'tırpanlık': 'TIRPANLIK  (m2)'
        }
        # Park ve Bulvar verilerini yükle ve sözlüğe çevir
        self.park_data = self._load_data(park_file) if park_file else {}
        self.bulvar_data = self._load_data(bulvar_file) if bulvar_file else {}

    def _normalize(self, text):
        return re.sub(r'\s+', ' ', str(text).upper().strip())

    def _load_data(self, uploaded_file):
        df = pd.read_excel(uploaded_file)
        data_dict = {}
        for _, row in df.iterrows():
            name = self._normalize(row.get('MAHAL ADI', ''))
            if name:
                data_dict[name] = {
                    k: (row[v] if v in df.columns and pd.notna(row[v]) else 0) 
                    for k, v in self.col_map.items()
                }
        return data_dict

    def _find_data(self, h_name, data_source):
        """Kısmi eşleşmeleri ve Gökçek gibi istisnaları yakalayan arama motoru"""
        if h_name in data_source: return data_source[h_name]
        for key, data in data_source.items():
            if h_name in key or key in h_name:
                if len(key) > 5 and len(h_name) > 5: return data
        return None

    def process_hakedis(self, template_bytes, sheet_name, target_type="PARK"):
        wb = openpyxl.load_workbook(BytesIO(template_bytes))
        sheet = wb[sheet_name]
        
        target_data = self.park_data if target_type == "PARK" else self.bulvar_data

        for row in range(1, sheet.max_row + 1):
            poz = sheet.cell(row=row, column=1).value
            imalat = sheet.cell(row=row, column=2).value
            
            # 1. Yapısal Satırları Koru (Poz No veya Toplam)
            if poz or (imalat and "TOPLAM" in str(imalat).upper()):
                continue

            if not isinstance(imalat, str):
                continue

            # 2. İmalat adını ve türünü ayır
            match = re.search(r'\(([^)]+)\)$', imalat.strip())
            if match:
                tur = match.group(1).lower()
                clean_name = self._normalize(imalat.replace(f"({match.group(1)})", ""))
            else:
                tur = None
                clean_name = self._normalize(imalat)

            # 3. Hedef veride bu mahal var mı? (Akıllı arama)
            data = self._find_data(clean_name, target_data)

            # 4. Karar Mekanizması
            if data:
                # Mahal bu dosyaya (Park veya Bulvar) ait
                if tur: # Çim, Çalı vb. alt satır
                    val = data.get(tur, 0)
                    if val > 0:
                        sheet.cell(row=row, column=5).value = val
                        sheet.row_dimensions[row].hidden = False 
                    else:
                        # Veri 0 ise EN değerini SİL ve gizle
                        sheet.cell(row=row, column=5).value = None
                        sheet.row_dimensions[row].hidden = True
                else: 
                    # Sarı Ana Başlık satırı, göster ve koru
                    sheet.cell(row=row, column=5).value = None
                    sheet.row_dimensions[row].hidden = False
            else:
                # Mahal bu gruba ait değil, EN değerini SİL ve gizle
                sheet.cell(row=row, column=5).value = None
                sheet.row_dimensions[row].hidden = True

        out = BytesIO()
        wb.save(out)
        return out.getvalue()


# --- KULLANICI ARAYÜZÜ (UI) ---
st.subheader("1. Veri Kaynakları")
col1, col2, col3 = st.columns(3)
with col1:
    h_file = st.file_uploader("Şablon Hakediş", type=['xlsx'])
with col2:
    p_file = st.file_uploader("Parklar Metrajı", type=['xlsx'])
with col3:
    b_file = st.file_uploader("Bulvarlar Metrajı", type=['xlsx'])

if h_file and p_file and b_file:
    
    # Şablonu belleğe al
    template_bytes = h_file.getvalue()
    wb_temp = openpyxl.load_workbook(BytesIO(template_bytes))
    
    st.subheader("2. İşlem Ayarları")
    selected_sheet = st.selectbox("İşlem Yapılacak Sayfa Seçiniz:", wb_temp.sheetnames)
    
    st.info("Onayladığınızda, şablonunuz kullanılarak Parklar ve Bulvarlar için iki ayrı hakediş dosyası eşzamanlı üretilecektir.")
    
    if st.button("Hakedişleri Hazırla", type="primary", use_container_width=True):
        with st.spinner("Dosyalar işleniyor, formüller güncelleniyor..."):
            
            engine = HakedisEngine(p_file, b_file)
            
            # İki dosyayı (Park ve Bulvar) ayrı ayrı üret
            park_result = engine.process_hakedis(template_bytes, selected_sheet, target_type="PARK")
            bulvar_result = engine.process_hakedis(template_bytes, selected_sheet, target_type="BULVAR")
            
            st.success("İşlem başarıyla tamamlandı. Dosyaları aşağıdan indirebilirsiniz.")
            
            d1, d2 = st.columns(2)
            with d1:
                st.download_button("📥 Parklar Hakedişini İndir", data=park_result, file_name="Hakedis_Parklar.xlsx", use_container_width=True)
            with d2:
                st.download_button("📥 Bulvarlar Hakedişini İndir", data=bulvar_result, file_name="Hakedis_Bulvarlar.xlsx", use_container_width=True)