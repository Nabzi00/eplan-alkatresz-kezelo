import streamlit as st
import pandas as pd
import io

# Oldal beállításai
st.set_page_config(page_title="EPLAN Alkatrészlista Kezelő", layout="wide")

# --- DINAMIKUS FEJLÉC (Cím és Projekt infó) ---
header_container = st.container()
with header_container:
    col_title, col_info = st.columns([3, 1])
    with col_title:
        st.title("⚙️ EPLAN Alkatrészlista Kezelő")
        st.write("Töltsd fel az EPLAN-ból kimentett Excel fájlt, szerkeszd a tételeket, majd töltsd le az eredményt!")
    with col_info:
        info_placeholder = st.empty()

# --- TO-DO LISTA RÉSZ ---
if 'todo_list' not in st.session_state:
    st.session_state.todo_list = []

with st.expander("📝 Napi Teendők (To-Do lista)", expanded=False):
    col1, col2 = st.columns([4, 1])
    with col1:
        uj_feladat = st.text_input("Új feladat:", label_visibility="collapsed", placeholder="Ide írd a feladatot...", key="uj_feladat_input")
    with col2:
        if st.button("Hozzáadás", use_container_width=True, key="add_button"):
            if uj_feladat.strip():
                st.session_state.todo_list.append(uj_feladat.strip())
                st.rerun()
    st.write("---")
    
    if not st.session_state.todo_list:
        st.caption("Nincsenek aktuális teendők.")
    else:
        megmarado_feladatok = []
        for i, feladat in enumerate(st.session_state.todo_list):
            c1, c2 = st.columns([9, 1])
            with c1:
                st.write(f"• {feladat}")
            with c2:
                if st.button("❌", key=f"del_{i}"):
                    continue
            megmarado_feladatok.append(feladat)
        
        if len(megmarado_feladatok) != len(st.session_state.todo_list):
            st.session_state.todo_list = megmarado_feladatok
            st.rerun()
            
st.divider()

# --- SEGÉDFÜGGVÉNY AZ OSZLOPOK BIZTONSÁGOS KIOLVASÁSÁHOZ ---
def get_column_safe(df, col_index):
    if col_index < len(df.columns):
        s = df.iloc[:, col_index].astype(str).str.strip()
        return s.replace(['nan', 'None', '', 'NaN', '<NA>'], pd.NA)
    return pd.Series([pd.NA] * len(df), index=df.index)

# --- FÁJLOK FELTÖLTÉSE ---
st.subheader("📁 Fájlok feltöltése")
col1, col2 = st.columns(2)

with col1:
    eplan_file = st.file_uploader("1. EPLAN Excel feltöltése", type=["xlsx"])
    
with col2:
    raktar_file = st.file_uploader("2. Raktárkészlet Excel feltöltése (Opcionális)", type=["xlsx"])

# --- ADATFELDOLGOZÁS ---
if eplan_file is not None:
    try:
        df_eplan_raw = pd.read_excel(eplan_file, header=None)
        
        # --- PROJEKT ÉS ÁLLOMÁS ADATOK KINYERÉSE ÉS CSERÉS FORMÁZÁSA ---
        nyers_projekt = str(df_eplan_raw.iloc[0, 1]).strip() if pd.notna(df_eplan_raw.iloc[0, 1]) else "Nincs megadva"
        
        if nyers_projekt != "Nincs megadva":
            # 1. Levágjuk az utolsó 5 karaktert
            projekt_szam = nyers_projekt[:-5] if len(nyers_projekt) > 5 else nyers_projekt
            
            # 2. LECSERÉLJÜK az 5. (index 4) és a 12. (index 11) karaktert
            if len(projekt_szam) > 11:
                # [0-tól 3-ig] + '/' + [5-től 10-ig] + '-' + [12-től a végéig]
                projekt_szam = projekt_szam[:4] + '/' + projekt_szam[5:11] + '-' + projekt_szam[12:]
            elif len(projekt_szam) > 4:
                # Ha rövidebb, mint 12, de hosszabb, mint 4
                projekt_szam = projekt_szam[:4] + '/' + projekt_szam[5:]
        else:
            projekt_szam = "Nincs megadva"
            
        # Állomás kinyerése
        allomas_szam = str(df_eplan_raw.iloc[1, 1]).strip() if pd.notna(df_eplan_raw.iloc[1, 1]) else "Nincs megadva"
        
        info_placeholder.markdown(
            f"""
            <div style="background-color: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 5px solid #005A9C;">
                <div style="font-size: 14px; color: #555;">📁 <b>Projekt:</b> {projekt_szam}</div>
                <div style="font-size: 14px; color: #555; margin-top: 5px;">📍 <b>Állomás:</b> {allomas_szam}</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        # --- OSZLOPNEVEK BEÁLLÍTÁSA AZ 5. SORBÓL ---
        oszlop_fejlecek = df_eplan_raw.iloc[4]
        df_eplan = df_eplan_raw.iloc[5:].reset_index(drop=True)
        df_eplan.columns = oszlop_fejlecek
        
        # --- ÜRES SOROK TISZTÍTÁSA ---
        df_eplan.dropna(how='all', inplace=True)
        df_eplan = df_eplan.reset_index(drop=True)
        
        # --- OSZLOPOK BIZTONSÁGOS KINYERÉSE ---
        eplan_a_col_name = df_eplan.columns[0] if len(df_eplan.columns) > 0 else None
        eplan_c_series = get_column_safe(df_eplan, 2) 
        eplan_d_series = get_column_safe(df_eplan, 3) 
        eplan_qty_series = get_column_safe(df_eplan, 4) 
        
        # --- CIKKSZÁM NÉLKÜLI SOROK SZŰRÉSE ---
        temp_cikkszam = eplan_d_series.fillna(eplan_c_series)
        valid_eplan_mask = temp_cikkszam.notna() & (temp_cikkszam != "")
        
        df_eplan = df_eplan[valid_eplan_mask].reset_index(drop=True)
        eplan_c_series = eplan_c_series[valid_eplan_mask].reset_index(drop=True)
        eplan_d_series = eplan_d_series[valid_eplan_mask].reset_index(drop=True)
        eplan_qty_series = eplan_qty_series[valid_eplan_mask].reset_index(drop=True)
        
        # --- A FELESLEGES 'A' OSZLOP ELDOBÁSA ---
        if eplan_a_col_name and eplan_a_col_name in df_eplan.columns:
            df_eplan = df_eplan.drop(columns=[eplan_a_col_name])
            
        df_eplan = df_eplan.fillna("")
        
    except Exception as e:
        st.error(f"❌ Hiba az EPLAN fájl beolvasásakor! Kérlek ellenőrizd a fájlt. Részletek: {e}")
        st.stop()
    
    if "Beszerzés státusza" not in df_eplan.columns:
        df_eplan.insert(0, "Beszerzés státusza", "Kiválasztandó...")
    
    # 2. RAKTÁRI ÖSSZEVETÉS
    if raktar_file is not None:
        try:
            df_raktar = pd.read_excel(raktar_file)
            
            df_raktar.dropna(how='all', inplace=True)
            df_raktar.dropna(how='all', axis=1, inplace=True)
            df_raktar = df_raktar.reset_index(drop=True)
            
            raktar_a_series = get_column_safe(df_raktar, 0)
            raktar_qty_series = get_column_safe(df_raktar, 1)
            raktar_e_series = get_column_safe(df_raktar, 4)
            raktar_f_hely_series = get_column_safe(df_raktar, 5)
            
            df_raktar = df_raktar.fillna("")
            
            eplan_vegleges_cikkszam = eplan_d_series.fillna(eplan_c_series).fillna("-")
            raktar_vegleges_cikkszam = raktar_e_series.fillna(raktar_a_series).fillna("-")
            
            raktar_f_helyek = raktar_f_hely_series.fillna("Nincs megadva")
            
            tiszta_eplan_qty = pd.to_numeric(eplan_qty_series.astype(str).str.replace(r'[^\d\.]', '', regex=True), errors='coerce').fillna(0)
            tiszta_raktar_qty = pd.to_numeric(raktar_qty_series.astype(str).str.replace(r'[^\d\.]', '', regex=True), errors='coerce').fillna(0)
            
            valid_mask = raktar_vegleges_cikkszam != "-"
            raktar_hely_dict = dict(zip(raktar_vegleges_cikkszam[valid_mask], raktar_f_helyek[valid_mask]))
            raktar_qty_dict = dict(zip(raktar_vegleges_cikkszam[valid_mask], tiszta_raktar_qty[valid_mask]))
            
            for col in ['Raktáron találva', 'Raktárhely', 'Raktári készlet']:
                if col in df_eplan.columns:
                    df_eplan = df_eplan.drop(columns=[col])
            
            df_eplan['Raktárhely'] = eplan_vegleges_cikkszam.map(raktar_hely_dict).fillna("-")
            df_eplan['Raktári készlet'] = eplan_vegleges_cikkszam.map(raktar_qty_dict).fillna(0)
            
            def ertekel_mennyiseg(row, idx):
                if row['Raktárhely'] == "-":
                    return "❌ Nem"
                
                e_qty = tiszta_eplan_qty.iloc[idx]
                r_qty = row['Raktári készlet']
                
                if r_qty <= 0:
                    return "❌ Nem"
                elif r_qty >= e_qty:
                    return "✅ Igen"
                else:
                    return "⚠️ Részleges"
            
            df_eplan['Raktáron találva'] = [ertekel_mennyiseg(row, idx) for idx, row in df_eplan.iterrows()]
            
            df_eplan.loc[df_eplan['Raktáron találva'] == "✅ Igen", "Beszerzés státusza"] = "Áttároltatható"
            df_eplan.loc[df_eplan['Raktáron találva'] == "⚠️ Részleges", "Beszerzés státusza"] = "Rendelni"
            df_eplan.loc[df_eplan['Raktáron találva'] == "❌ Nem", "Beszerzés státusza"] = "Rendelni"
            
            col_keszlet = df_eplan.pop('Raktári készlet')
            col_hely = df_eplan.pop('Raktárhely')
            col_talalat = df_eplan.pop('Raktáron találva')
            df_eplan.insert(0, 'Raktáron találva', col_talalat)
            df_eplan.insert(1, 'Raktárhely', col_hely)
            df_eplan.insert(2, 'Raktári készlet', col_keszlet)
            
            st.success("✅ A rendszer automatikusan beolvasta a Projekt adatait, és elvégezte az összevetést!")
                
        except Exception as e:
            st.error(f"❌ Hiba történt a Raktár Excel beolvasásakor! Részletek: {e}")

    # --- TÖBBSZÖRÖS SZŰRŐ GOMBOK ---
    st.divider()
    st.subheader("📝 Adatok szerkesztése és szűrése")
    
    df_to_display = df_eplan.copy()
    
    if 'Raktáron találva' in df_to_display.columns:
        szuro_opciok = ["✅ Igen", "⚠️ Részleges", "❌ Nem"]
        
        kivalasztott_szurok = st.multiselect(
            "Szűrés raktári találat alapján (többet is választhatsz):",
            options=szuro_opciok,
            default=szuro_opciok,
            help="Kattints az 'X'-re egy feltétel elrejtéséhez, vagy a mezőre új hozzáadásához."
        )
        
        df_to_display = df_to_display[df_to_display['Raktáron találva'].isin(kivalasztott_szurok)]

    # --- AUTOMATIKUS SZÍNEZÉS FEKETE BETŰKKEL ---
    def color_rows(row):
        if 'Raktáron találva' in row and row['Raktáron találva'] in ["✅ Igen", "⚠️ Részleges"]:
            hely = str(row.get('Raktárhely', '')).strip()
            if hely != "Projektek":
                return ['background-color: #C6E0B4; color: black;'] * len(row) 
            else:
                return ['background-color: #FFF2CC; color: black;'] * len(row) 
        return [''] * len(row)

    styled_eplan = df_to_display.style.apply(color_rows, axis=1)

    # --- TÁBLÁZAT MEGJELENÍTÉSE ---
    edited_df = st.data_editor(
        styled_eplan,
        column_config={
            "Beszerzés státusza": st.column_config.SelectboxColumn(
                "Beszerzés státusza",
                help="Válaszd ki, mi történjen az alkatrésszel",
                options=[
                    "Rendelni", 
                    "Raktárból", 
                    "Áttároltatható", 
                    "Közös rendelés", 
                    "Már megrendelve", 
                    "Nem kell"
                ],
                required=True
            )
        },
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic"
    )

    # --- EXPORTÁLÁS ---
    st.divider()
    st.subheader("📥 Kész fájl exportálása")
    st.info("💡 A rendszer mindig pontosan azokat a sorokat menti ki Excelbe, amiket jelenleg a fenti táblázatban látsz!")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False, sheet_name='Feldolgozott_Lista')
    
    st.download_button(
        label="Letöltés új Excelként",
        data=buffer.getvalue(),
        file_name="EPLAN_rendelesi_lista_kesz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )