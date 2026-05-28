import streamlit as st
import pandas as pd
import io
import openpyxl
import json

# Oldal beállításai
st.set_page_config(page_title="EPLAN Alkatrészlista Kezelő", layout="wide")

# --- SEGÉDFÜGGVÉNY A DÁTUM KINYERÉSÉHEZ ---
def get_excel_modified_date(uploaded_file):
    try:
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file, read_only=True)
        mod_time = wb.properties.modified
        uploaded_file.seek(0)
        
        if mod_time:
            return mod_time.strftime("%Y. %m. %d. %H:%M")
        return "Ismeretlen dátum"
    except Exception:
        uploaded_file.seek(0)
        return "Ismeretlen dátum"

# --- DINAMIKUS FEJLÉC (Cím és Projekt infó) ---
header_container = st.container()
with header_container:
    col_title, col_info = st.columns([3, 1])
    with col_title:
        st.title("⚙️ EPLAN Alkatrészlista Kezelő")
        st.write("Töltsd fel az EPLAN-ból kimentett Excel fájlt, szerkeszd a tételeket, majd töltsd le az eredményt!")
    with col_info:
        info_placeholder = st.empty()

# --- TO-DO LISTA RÉSZ (Szerkeszthető verzió + TXT letöltés) ---
if 'todo_list' not in st.session_state:
    saved_todo = st.query_params.get("todo_data", None)
    if saved_todo:
        st.session_state.todo_list = json.loads(saved_todo)
    else:
        st.session_state.todo_list = []

# Szerkesztési állapot kezelése
if 'editing_index' not in st.session_state:
    st.session_state.editing_index = None

def save_todo():
    st.query_params["todo_data"] = json.dumps(st.session_state.todo_list)

with st.expander("📝 Napi Teendők (To-Do lista)", expanded=False):
    col1, col2 = st.columns([4, 1])
    with col1:
        uj_feladat = st.text_input("Új feladat:", label_visibility="collapsed", placeholder="Ide írd a feladatot...", key="uj_feladat_input")
    with col2:
        if st.button("Hozzáadás", use_container_width=True, key="add_button"):
            if uj_feladat.strip():
                st.session_state.todo_list.append(uj_feladat.strip())
                save_todo()
                st.rerun()
    st.write("---")
    
    if not st.session_state.todo_list:
        st.caption("Nincsenek aktuális teendők.")
    else:
        for i, feladat in enumerate(st.session_state.todo_list):
            if st.session_state.editing_index == i:
                # Szerkesztő mód
                c1, c2 = st.columns([8, 2])
                with c1:
                    edit_val = st.text_input(f"edit_{i}", value=feladat, label_visibility="collapsed")
                with c2:
                    if st.button("✅ OK", key=f"save_{i}"):
                        st.session_state.todo_list[i] = edit_val
                        st.session_state.editing_index = None
                        save_todo()
                        st.rerun()
            else:
                # Olvasó mód
                c1, c2, c3 = st.columns([8, 1, 1])
                with c1:
                    st.write(f"• {feladat}")
                with c2:
                    if st.button("✏️", key=f"edit_{i}"):
                        st.session_state.editing_index = i
                        st.rerun()
                with c3:
                    if st.button("❌", key=f"del_{i}"):
                        st.session_state.todo_list.pop(i)
                        save_todo()
                        st.rerun()

        # ÚJ FUNKCIÓ: Teendők letöltése TXT-be
        st.write("---")
        todo_text = "\n".join(st.session_state.todo_list)
        st.download_button(
            label="📥 Teendők mentése TXT-be",
            data=todo_text,
            file_name="teendok.txt",
            mime="text/plain"
        )
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
    eplan_ph = st.empty()
    eplan_file = st.file_uploader("eplan_upload", type=["xlsx", "xls"], label_visibility="collapsed")
    
    if eplan_file is not None:
        eplan_date = get_excel_modified_date(eplan_file)
        eplan_ph.markdown(f"""
            <div style='display: flex; justify-content: space-between; align-items: center; padding-bottom: 5px;'>
                <span style='font-size: 14px;'>1. EPLAN Excel feltöltése</span>
                <span style='color: #888; font-size: 12px;'>🕒 <b>Utolsó mentés:</b> {eplan_date}</span>
            </div>
        """, unsafe_allow_html=True)
    else:
        eplan_ph.markdown("<div style='font-size: 14px; padding-bottom: 5px;'>1. EPLAN Excel feltöltése</div>", unsafe_allow_html=True)
    
with col2:
    raktar_ph = st.empty()
    raktar_file = st.file_uploader("raktar_upload", type=["xlsx", "xls"], label_visibility="collapsed")
    
    if raktar_file is not None:
        raktar_date = get_excel_modified_date(raktar_file)
        raktar_ph.markdown(f"""
            <div style='display: flex; justify-content: space-between; align-items: center; padding-bottom: 5px;'>
                <span style='font-size: 14px;'>2. DM aktuális készlet</span>
                <span style='color: #888; font-size: 12px;'>🕒 <b>Utolsó mentés:</b> {raktar_date}</span>
            </div>
        """, unsafe_allow_html=True)
    else:
        raktar_ph.markdown("<div style='font-size: 14px; padding-bottom: 5px;'>2. DM aktuális készlet (Opcionális)</div>", unsafe_allow_html=True)

# --- ADATFELDOLGOZÁS ---
# Projektváltozók előkészítése (hogy az exportnál is meglegyenek)
projekt_szam = "Nincs megadva"
allomas_szam = "Nincs megadva"

if eplan_file is not None:
    try:
        # ELLENŐRIZZÜK, HOGY EZ EGY FOLYTATANDÓ (KORÁBBAN KIMENTETT) FÁJL-E
        df_check = pd.read_excel(eplan_file, nrows=5)
        is_resumed = "Beszerzés státusza" in df_check.columns

        if is_resumed:
            # --- FÉLBEHAGYOTT MUNKA BETÖLTÉSE ---
            df_eplan = pd.read_excel(eplan_file)
            df_eplan = df_eplan.fillna("")
            
            info_placeholder.markdown(
                f"""
                <div style="background-color: #f0f2f6; padding: 15px; border-radius: 10px; border-left: 5px solid #28a745;">
                    <div style="font-size: 14px; color: #555;">🔄 <b>Állapot:</b> Folytatott munka betöltve</div>
                    <div style="font-size: 14px; color: #555; margin-top: 5px;">Minden korábbi módosításod visszaállítva.</div>
                </div>
                """, 
                unsafe_allow_html=True
            )
            st.success("🔄 Sikeresen betöltöttük a korábban kimentett listát! Folytathatod a munkát.")
            
        else:
            # --- NYERS EPLAN FÁJL FELDOLGOZÁSA ---
            df_eplan_raw = pd.read_excel(eplan_file, header=None)
            
            # Projekt adatok kinyerése
            nyers_projekt = str(df_eplan_raw.iloc[0, 1]).strip() if pd.notna(df_eplan_raw.iloc[0, 1]) else "Nincs megadva"
            if nyers_projekt != "Nincs megadva":
                projekt_szam = nyers_projekt[:-5] if len(nyers_projekt) > 5 else nyers_projekt
                if len(projekt_szam) > 11:
                    projekt_szam = projekt_szam[:4] + '/' + projekt_szam[5:11] + '-' + projekt_szam[12:]
                elif len(projekt_szam) > 4:
                    projekt_szam = projekt_szam[:4] + '/' + projekt_szam[5:]
            
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
            
            # Oszlopnevek és tisztítás
            oszlop_fejlecek = df_eplan_raw.iloc[4]
            df_eplan = df_eplan_raw.iloc[5:].reset_index(drop=True)
            df_eplan.columns = oszlop_fejlecek
            
            df_eplan.dropna(how='all', inplace=True)
            df_eplan = df_eplan.reset_index(drop=True)
            
            eplan_a_col_name = df_eplan.columns[0] if len(df_eplan.columns) > 0 else None
            eplan_c_series = get_column_safe(df_eplan, 2) 
            eplan_d_series = get_column_safe(df_eplan, 3) 
            eplan_qty_series = get_column_safe(df_eplan, 4) 
            
            temp_cikkszam = eplan_d_series.fillna(eplan_c_series)
            valid_eplan_mask = temp_cikkszam.notna() & (temp_cikkszam != "")
            
            df_eplan = df_eplan[valid_eplan_mask].reset_index(drop=True)
            eplan_c_series = eplan_c_series[valid_eplan_mask].reset_index(drop=True)
            eplan_d_series = eplan_d_series[valid_eplan_mask].reset_index(drop=True)
            eplan_qty_series = eplan_qty_series[valid_eplan_mask].reset_index(drop=True)
            
            if eplan_a_col_name and eplan_a_col_name in df_eplan.columns:
                df_eplan = df_eplan.drop(columns=[eplan_a_col_name])
                
            df_eplan = df_eplan.fillna("")
            df_eplan.insert(0, "Beszerzés státusza", "Kiválasztandó...")

    except Exception as e:
        st.error(f"❌ Hiba az EPLAN fájl beolvasásakor! Kérlek ellenőrizd a fájlt. Részletek: {e}")
        st.stop()
    
  # --- RAKTÁRI ÖSSZEVETÉS ---
    if raktar_file is not None:
        try:
            # 1. Beolvassuk "nyersen"
            df_raktar_raw = pd.read_excel(raktar_file, header=None)
            
            # 2. Kulcsszavak (Sarzs benne van!)
            kulcsszavak = {
                'cikkszam': ['cikkszám', 'típus', 'cikk', 'megrendelési szám', 'anyagszám', 'material', 'azonosító'],
                'mennyiseg': ['mennyiség', 'készlet', 'db', 'stock', 'szabad', 'raktárkészlet', 'menny.'],
                'hely': ['raktárhely', 'hely', 'fiók', 'polc', 'bin', 'storage', 'tárhely', 'sarzs']
            }
            
            fejlec_sor_idx = -1
            c_idx, q_idx, h_idx = -1, -1, -1

            # 3. Keresés az első 20 sorban
            for row_idx, row in df_raktar_raw.head(20).iterrows():
                temp_c, temp_q, temp_h = -1, -1, -1
                for col_idx, cell_value in enumerate(row):
                    cell_str = str(cell_value).lower().strip()
                    if temp_c == -1 and any(k in cell_str for k in kulcsszavak['cikkszam']):
                        temp_c = col_idx
                    elif temp_q == -1 and any(k in cell_str for k in kulcsszavak['mennyiseg']):
                        temp_q = col_idx
                    elif temp_h == -1 and any(k in cell_str for k in kulcsszavak['hely']):
                        temp_h = col_idx
                        
                if temp_c != -1 and temp_q != -1:
                    fejlec_sor_idx = row_idx
                    c_idx = temp_c
                    q_idx = temp_q
                    h_idx = temp_h
                    break
            
            # 4. Ha megvan a fejléc, kinyerjük az adatokat
            if fejlec_sor_idx != -1:
                # Biztosíték: ha nincs meg a Sarzs szó, kényszerítjük az A oszlopot (0)
                if h_idx == -1:
                    h_idx = 0 
                
                # --- HELYKÓD MÁSOLÓ ---
                actual_helyek = df_raktar_raw.iloc[:, h_idx].astype(str).str.strip()
                # Kiterjesztettük az üres cellák vizsgálatát
                mask_empty = actual_helyek.str.lower().isin(['nan', 'none', 'null', '<na>', '', '0', '0.0'])
                
                if c_idx == h_idx:
                    qty_col = df_raktar_raw.iloc[:, q_idx].astype(str).str.lower().str.strip()
                    qty_empty = qty_col.isin(['nan', 'none', 'null', '<na>', '', '0', '0.0'])
                    is_location = (~mask_empty) & qty_empty
                    actual_helyek[~is_location] = None
                else:
                    actual_helyek[mask_empty] = None
                
                actual_helyek.iloc[fejlec_sor_idx] = None
                actual_helyek = actual_helyek.ffill().fillna("Nincs megadva")
                
                df_raktar = df_raktar_raw.iloc[fejlec_sor_idx + 1:].reset_index(drop=True)
                raktar_f_helyek = actual_helyek.iloc[fejlec_sor_idx + 1:].reset_index(drop=True)
                
                # --- CIKKSZÁM ÉS MENNYISÉG TISZTÍTÁS (Itt a .0 javítás!) ---
                raktar_vegleges_cikkszam = df_raktar.iloc[:, c_idx].astype(str).str.strip()
                # Eltávolítjuk a rejtett ".0" tizedeseket a végéről
                raktar_vegleges_cikkszam = raktar_vegleges_cikkszam.str.replace(r'\.0$', '', regex=True)
                raktar_vegleges_cikkszam = raktar_vegleges_cikkszam.replace(['nan', 'None', '', 'NaN', '<NA>'], "-")
                
                tiszta_raktar_qty = pd.to_numeric(df_raktar.iloc[:, q_idx].astype(str).str.replace(r'[^\d\.]', '', regex=True), errors='coerce').fillna(0)
                
                # --- TÖBB TÁRHELY ÖSSZEVONÁSA ---
                df_raktar_clean = pd.DataFrame({
                    'Cikkszám': raktar_vegleges_cikkszam,
                    'Mennyiség': tiszta_raktar_qty,
                    'Hely': raktar_f_helyek
                })
                df_raktar_clean = df_raktar_clean[df_raktar_clean['Cikkszám'] != "-"]
                
                df_shelf_totals = df_raktar_clean.groupby(['Cikkszám', 'Hely'], as_index=False)['Mennyiség'].sum()
                
                def format_shelf_qty(v):
                    return str(int(v)) if v == int(v) else str(v)
                df_shelf_totals['Hely_Kombinalt'] = df_shelf_totals['Hely'].astype(str) + " (" + df_shelf_totals['Mennyiség'].apply(format_shelf_qty) + " db)"
                
                aggregated_raktar = df_shelf_totals.groupby('Cikkszám').agg({
                    'Mennyiség': 'sum',
                    'Hely_Kombinalt': lambda x: ", ".join(x)
                })
                
                raktar_hely_dict = aggregated_raktar['Hely_Kombinalt'].to_dict()
                raktar_qty_dict = aggregated_raktar['Mennyiség'].to_dict()
                
                # --- 🛠️ DEBUG (Röntgen nézet) MEGJEELNÍTÉSE ---
                with st.expander("🛠️ Fejlesztői Debug: Hogyan látja a program a raktári adatokat?", expanded=False):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.write(f"**Fejléc sor:** {fejlec_sor_idx}")
                    c2.write(f"**Cikkszám oszlop:** {c_idx}")
                    c3.write(f"**Mennyiség oszlop:** {q_idx}")
                    c4.write(f"**Sarzs/Hely oszlop:** {h_idx}")
                    st.write("---")
                    st.write("**Első 15 feldolgozott (Cikkszám, Mennyiség, Kitalált Helykód):**")
                    st.dataframe(df_raktar_clean.head(15), use_container_width=True)

            else:
                st.error("❌ Nem találtam Cikkszám és Mennyiség oszlopokat a raktári fájlban!")
                st.stop()

            # --- 5. EPLAN CIKKSZÁM PÁROSÍTÁS (.0 TISZTÍTÁSSAL EGYÜTT!) ---
            if 'is_resumed' in locals() and is_resumed:
                cikkszam_col = next((col for col in df_eplan.columns if 'cikkszám' in str(col).lower() or 'típus' in str(col).lower()), df_eplan.columns[3])
                eplan_vegleges_cikkszam = df_eplan[cikkszam_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).replace(['nan', 'None', '', 'NaN'], "-")
                qty_col = next((col for col in df_eplan.columns if 'mennyiség' in str(col).lower() or 'db' in str(col).lower()), df_eplan.columns[4])
                tiszta_eplan_qty = pd.to_numeric(df_eplan[qty_col].astype(str).str.replace(r'[^\d\.]', '', regex=True), errors='coerce').fillna(0)
            else:
                eplan_vegleges_cikkszam = eplan_d_series.fillna(eplan_c_series).astype(str).str.strip().str.replace(r'\.0$', '', regex=True).fillna("-")
                tiszta_eplan_qty = pd.to_numeric(eplan_qty_series.astype(str).str.replace(r'[^\d\.]', '', regex=True), errors='coerce').fillna(0)
                
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
                
                if r_qty <= 0: return "❌ Nem"
                elif r_qty >= e_qty: return "✅ Igen"
                else: return "⚠️ Részleges"
            
            df_eplan['Raktáron találva'] = [ertekel_mennyiseg(row, idx) for idx, row in df_eplan.iterrows()]
            
            mask_igen = (df_eplan['Raktáron találva'] == "✅ Igen") & (df_eplan['Beszerzés státusza'] == "Kiválasztandó...")
            df_eplan.loc[mask_igen, "Beszerzés státusza"] = "Áttároltatható"
            mask_nem = (df_eplan['Raktáron találva'].isin(["⚠️ Részleges", "❌ Nem"])) & (df_eplan['Beszerzés státusza'] == "Kiválasztandó...")
            df_eplan.loc[mask_nem, "Beszerzés státusza"] = "Rendelni"
            
            col_keszlet = df_eplan.pop('Raktári készlet')
            col_hely = df_eplan.pop('Raktárhely')
            col_talalat = df_eplan.pop('Raktáron találva')
            df_eplan.insert(0, 'Raktáron találva', col_talalat)
            df_eplan.insert(1, 'Raktárhely', col_hely)
            df_eplan.insert(2, 'Raktári készlet', col_keszlet)
            
            st.success("✅ Raktárkészlet és Sarzs kódok sikeresen beolvasva és összevonva!")
                
        except Exception as e:
            st.error(f"❌ Hiba történt a DM aktuális készlet beolvasásakor! Részletek: {e}")
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
                    "Kiválasztandó...",
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

    # --- EXPORTÁLÓ SZEKCIÓ ---
    st.divider()
    st.subheader("📥 Exportálás és Mentés")
    
    col_exp1, col_exp2 = st.columns(2)
    
    # 1. GOMB: Saját mentés (folytatáshoz)
    with col_exp1:
        st.info("💡 **Munka mentése:** Ha később folytatnád, ezt mentsd le.")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Feldolgozott_Lista')
        
        st.download_button(
            label="💾 Mentés (Munka folytatásához)",
            data=buffer.getvalue(),
            file_name="EPLAN_munka_mentese.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ÚJ FUNKCIÓ: 2. GOMB (Csillagpont automata projekt + A/B felcserélve)
    with col_exp2:
        if st.button("📤 Csillagpont ajánlatkérő generálása"):
            try:
                wb = openpyxl.load_workbook("sablon.xlsx")
                ws = wb["VM_BOM"]
                
                # AUTOMATIKUS KITÖLTÉS (a nyers fájlból kinyert adatokkal)
                ws["B1"] = projekt_szam
                ws["B2"] = allomas_szam
                
                # Szűrés a FORGALMAZÓ alapján
                csillagpont_df = edited_df[
                    edited_df["FORGALMAZÓ"].astype(str).str.contains("Csillagpont", case=False, na=False)
                ]
                
                if csillagpont_df.empty:
                    st.warning("⚠️ Nincs találat a 'Csillagpont' forgalmazónál.")
                else:
                    # A 6. sortól kezdjük az adatírást
                    for r_idx, (index, row) in enumerate(csillagpont_df.iterrows(), start=6):
                        # A6: Megnevezés (felcserélve az eredeti B-vel)
                        ws.cell(row=r_idx, column=1, value=row["MEGNEVEZÉS 1"])
                        # B6: Cikkszám (felcserélve az eredeti A-val)
                        ws.cell(row=r_idx, column=2, value=row["MEGRENDELÉSI SZÁM"])
                        # C6: Mennyiség
                        ws.cell(row=r_idx, column=3, value=row["MENNYISÉG"])
                        # D6: Mértékegység
                        ws.cell(row=r_idx, column=4, value=row["MÉRTÉKEGYSÉG"])
                        # E6: Gyártó
                        ws.cell(row=r_idx, column=5, value=row["GYÁRTÓ"])
                    
                    # Mentés és letöltés
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Letöltés: Csillagpont ajánlatkérő",
                        data=output.getvalue(),
                        file_name="Csillagpont_ajanlatkeres.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success(f"✅ Siker! {len(csillagpont_df)} tétel exportálva.")
            except Exception as e:
                st.error(f"❌ Hiba: {e}")