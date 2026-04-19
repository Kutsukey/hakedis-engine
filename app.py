import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

class HakedisPro:
    def __init__(self, hakedis_file, snc_file, hakedis_sheet, snc_sheet):
        self.wb = openpyxl.load_workbook(hakedis_file)
        self.sheet = self.wb[hakedis_sheet]
        self.df_snc = pd.read_excel(snc_file, sheet_name=snc_sheet)
        self.col_map = {
            'çim': 'ÇİM  (m2)', 'çalı': 'ÇALI (m2)', 'çiçek': 'ÇİÇEK (m2)',
            'sert': 'SERT (m2)', 'çoa': 'ÇOA (m2)', 'ağaçlık': 'AĞAÇLIK (m2)',
            'spor': 'SPOR(m2)', 'toprak': 'TOPRAK (m2)', 'tırpanlık': 'TIRPANLIK  (m2)'
        }
        self.snc_dict = self._prepare_dict()

    def _normalize(self, text):
        return re.sub(r'\s+', ' ', str(text).upper().strip())

    def _prepare_dict(self):
        d = {}
        for _, row in self.df_snc.iterrows():
            name = self._normalize(row['MAHAL ADI'])
            d[name] = {k: (row[v] if v in self.df_snc.columns and pd.notna(row[v]) else 0) for k, v in self.col_map.items()}
        return d

    def process(self, filter_list):
        for row in range(1, self.sheet.max_row + 1):
            # Yapısal satırları koru (Poz No varsa veya TOPLAM yazıyorsa)
            poz = self.sheet.cell(row=row, column=1).value
            imalat = self.sheet.cell(row=row, column=2).value
            if poz or (imalat and "TOPLAM" in str(imalat).upper()): continue
            if not isinstance(imalat, str): continue

            match = re.search(r'\(([^)]+)\)$', imalat.strip())
            if match:
                tur = match.group(1).lower()
                clean_name = self._normalize(imalat.replace(f"({match.group(1)})", ""))
                
                # Akıllı Eşleştirme
                is_target = any(f.upper() in clean_name for f in filter_list)
                val = self.snc_dict.get(clean_name, {}).get(tur, 0)

                if is_target and val > 0:
                    self.sheet.cell(row=row, column=5).value = val
                else:
                    self.sheet.cell(row=row, column=5).value = None
                    self.sheet.row_dimensions[row].hidden = True
            else:
                # Verisi olmayan boş satırları gizle
                if not self.sheet.cell(row=row, column=3).value:
                    self.sheet.row_dimensions[row].hidden = True
        
        out = BytesIO()
        self.wb.save(out)
        return out.getvalue()

# --- UI Arayüzü ---
st.set_page_config(page_title="Hakediş Otomasyonu", layout="wide")
st.title("🚀 Profesyonel Hakediş Filtreleme Sistemi")

col1, col2 = st.columns(2)
with col1:
    h_file = st.file_uploader("Hakediş Dosyasını Seç (xlsx)", type=['xlsx'])
with col2:
    s_file = st.file_uploader("SNC/Metraj Dosyasını Seç (xlsx)", type=['xlsx'])

if h_file and s_file:
    h_sheets = openpyxl.load_workbook(h_file).sheetnames
    s_sheets = pd.ExcelFile(s_file).sheet_names
    
    sel_h = st.selectbox("Hakediş Sayfası:", h_sheets)
    sel_s = st.selectbox("Metraj Veri Sayfası:", s_sheets)
    
    filter_input = st.text_area("İşlem yapılacak park/cadde isimlerini girin (Virgül ile ayırın):", 
                                "GOP PARKI, FATİH PARKI, DEMOKRASİ")
    
    if st.button("Hakedişi Hazırla"):
        f_list = [i.strip() for i in filter_input.split(",")]
        engine = HakedisEngine(h_file, s_file, sel_h, sel_s)
        result = engine.process(f_list)
        
        st.success("İşlem Tamamlandı!")
        st.download_button(label="📁 Güncel Hakedişi İndir", 
                           data=result, 
                           file_name="Guncel_Hakedis.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")