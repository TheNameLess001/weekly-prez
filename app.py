import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# --- 1. CONFIGURATION & DESIGN ---
st.set_page_config(page_title="WBR Generator | LMD", page_icon="💎", layout="wide")

custom_css = """
<style>
    :root { --primary-purple: #4C1D95; --accent-red: #E11D48; --bg-light: #F8FAFC; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    h1, h2, h3 { color: var(--primary-purple) !important; font-family: 'Helvetica Neue', sans-serif; font-weight: 800; }
    .stButton>button { background: linear-gradient(90deg, #4C1D95 0%, #E11D48 100%); color: white !important; border: none; border-radius: 8px; padding: 10px 24px; font-weight: bold; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(225, 29, 72, 0.3); }
    [data-testid="stMetricValue"] { color: var(--accent-red) !important; font-weight: 900; }
    [data-testid="stMetricLabel"] { color: var(--primary-purple) !important; font-weight: bold; font-size: 16px; }
    .streamlit-expanderHeader { background-color: var(--bg-light); color: var(--primary-purple) !important; font-weight: bold; border-radius: 8px; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)
st.title("💎 Générateur Weekly Business Review (Yassir / Kool)")

# --- 2. FONCTIONS DATA ---
@st.cache_data
def process_admin_file(file):
    df = pd.read_csv(file)
    df.columns = df.columns.str.strip().str.lower()
    if 'order day' in df.columns:
        df['order day'] = pd.to_datetime(df['order day'], errors='coerce')
    return df

@st.cache_data
def process_aux_files(files):
    df_sales_list, df_cons_list = [], []
    for f in files:
        try:
            df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
            cols = df.columns.str.strip().str.lower()
            df.columns = cols
            if 'statut' in cols or 'segment ' in cols: df_sales_list.append(df)
            elif "nom de l'établissement" in cols or "nom de l'etablissement" in cols: df_cons_list.append(df)
        except: continue
    df_sales = pd.concat(df_sales_list, ignore_index=True) if df_sales_list else pd.DataFrame()
    df_cons = pd.concat(df_cons_list, ignore_index=True) if df_cons_list else pd.DataFrame()
    return df_sales, df_cons

def calculate_kpis(df):
    if df is None or df.empty:
        return {'orders': 0, 'gmv': 0, 'aov': 0, 'cancel_rate': 0, 'avg_time': 0, 'top_partners': [], 'cancel_reasons': []}
    
    df_delivered = df[df['status'] == 'Delivered']
    orders = len(df_delivered)
    col_gmv = 'item total' if 'item total' in df.columns else 'payable amount'
    gmv = df_delivered[col_gmv].sum() if col_gmv in df.columns else 0
    aov = gmv / orders if orders > 0 else 0
    
    col_time = 'delivery time(m)'
    avg_time = df_delivered[col_time].mean() if col_time in df_delivered.columns else 0
    
    df_cancelled = df[df['status'].astype(str).str.contains('Cancel', na=False, case=False)]
    cancel_rate = (len(df_cancelled) / len(df)) * 100 if len(df) > 0 else 0
    
    cancel_reasons = []
    if 'cancellation reason' in df.columns and len(df_cancelled) > 0:
        cancel_reasons = [(r, c) for r, c in df_cancelled['cancellation reason'].value_counts().items()][:3]

    top_partners = []
    if 'restaurant name' in df.columns and not df_delivered.empty:
        grouped = df_delivered.groupby('restaurant name').agg(orders=('status','count'), gmv=(col_gmv,'sum')).reset_index()
        grouped['aov'] = grouped['gmv'] / grouped['orders']
        grouped['weight'] = (grouped['gmv'] / gmv * 100) if gmv > 0 else 0
        top_partners = grouped.sort_values(by='gmv', ascending=False).head(10).to_dict('records')

    return {'orders': orders, 'gmv': gmv, 'aov': aov, 'avg_time': avg_time, 'cancel_rate': cancel_rate, 'top_partners': top_partners, 'cancel_reasons': cancel_reasons}

# --- 3. FIX: FONCTIONS PPTX SECURISÉES ---
def add_slide_with_bullet_points(prs, title_text, text_lines):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title_text
    
    # Solution robuste pour trouver le bloc de texte (placeholder)
    body_shape = None
    for shape in slide.placeholders:
        if shape.placeholder_format.idx != 0: # 0 = Titre
            body_shape = shape
            break
            
    if body_shape and body_shape.has_text_frame:
        tf = body_shape.text_frame
        tf.text = text_lines[0] if text_lines else ""
        for line in text_lines[1:]:
            p = tf.add_paragraph()
            p.text = line

def add_table_to_slide(slide, table_data, headers, left, top, width, height):
    rows, cols = len(table_data) + 1, len(headers)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    for i, h in enumerate(headers): table.cell(0, i).text = str(h)
    for r_idx, row in enumerate(table_data):
        for c_idx, val in enumerate(row):
            table.cell(r_idx + 1, c_idx).text = str(val)

# --- 4. INTERFACE ---
st.sidebar.markdown("### 📥 1. Imports")
admin_file = st.sidebar.file_uploader("1. Export Admin (CSV)*", type=['csv'])
aux_files = st.sidebar.file_uploader("2. Sales / Consolider (XLSX)", type=['csv', 'xlsx'], accept_multiple_files=True)
template_file = st.sidebar.file_uploader("3. Template PPTX", type=['pptx'])

if admin_file:
    with st.spinner("Traitement des données..."):
        df_admin = process_admin_file(admin_file)
        df_sales, df_cons = process_aux_files(aux_files) if aux_files else (pd.DataFrame(), pd.DataFrame())
        
        min_d = df_admin['order day'].min().date() if not pd.isna(df_admin['order day'].min()) else datetime.date.today()
        max_d = df_admin['order day'].max().date() if not pd.isna(df_admin['order day'].max()) else datetime.date.today()
    
    st.sidebar.markdown("### 📅 2. Période")
    start_date = st.sidebar.date_input("Date de début", min_d)
    end_date = st.sidebar.date_input("Date de fin", max_d)
    
    duration = (end_date - start_date).days + 1
    prev_start = start_date - datetime.timedelta(days=duration)
    prev_end = start_date - datetime.timedelta(days=1)
    
    mask_curr = (df_admin['order day'].dt.date >= start_date) & (df_admin['order day'].dt.date <= end_date)
    mask_prev = (df_admin['order day'].dt.date >= prev_start) & (df_admin['order day'].dt.date <= prev_end)
    
    kpis = calculate_kpis(df_admin[mask_curr])
    kpis_prev = calculate_kpis(df_admin[mask_prev])
    
    signed_count = len(df_sales[df_sales['statut'].astype(str).str.contains('Signé', case=False, na=False)]) if not df_sales.empty and 'statut' in df_sales.columns else 146
    onboarded_count = len(df_cons[(df_cons['date'].dt.date >= start_date) & (df_cons['date'].dt.date <= end_date)]) if not df_cons.empty and 'date' in df_cons.columns else 110

    tab1, tab2, tab3 = st.tabs(["📊 Dashboard KPIs", "✍️ Saisie Manuelle", "📤 Export PPTX"])

    with tab1:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Commandes Livrées", f"{kpis['orders']:,}", f"{(kpis['orders']-kpis_prev['orders'])/kpis_prev['orders']*100:.1f}% vs S-1" if kpis_prev['orders'] else "N/A")
        c2.metric("GMV (MAD)", f"{kpis['gmv']:,.0f}", f"{(kpis['gmv']-kpis_prev['gmv'])/kpis_prev['gmv']*100:.1f}% vs S-1" if kpis_prev['gmv'] else "N/A")
        c3.metric("AOV (MAD)", f"{kpis['aov']:.1f}")
        c4.metric("Taux d'annulation", f"{kpis['cancel_rate']:.1f}%")

    with tab2:
        with st.expander("📝 Slide 1: Executive Summary"):
            c1, c2, c3 = st.columns(3)
            in_req = c1.number_input("Total Requests", value=16800)
            in_budget_gmv = c2.number_input("Budget GMV", value=1150000)
            in_target_ord = c3.number_input("Target Orders", value=15000)
            c4, c5 = st.columns(2)
            in_freq = c4.number_input("Fréquence (cmds/user)", value=2.8)
            in_weekdays = c5.number_input("Répartition Weekdays (%)", value=65)
            in_high = st.text_area("Highlights", "Lancement réussi Market Zone Nord.\nRecord commandes ce weekend.")

        with st.expander("🚚 Slide 2: Opérations & Delivery"):
            c1, c2, c3, c4 = st.columns(4)
            in_t_food = c1.number_input("Tps Food (m)", value=42)
            in_t_groc = c2.number_input("Tps Groceries (m)", value=35)
            in_batch = c3.number_input("Batching Rate", value=1.3)
            in_auto = c4.number_input("Auto-dispatch (%)", value=94)
            
            c5, c6, c7, c8 = st.columns(4)
            in_fleet = c5.number_input("Coursiers Actifs", value=450)
            in_mix = c6.number_input("Mix 3PL (%)", value=65)
            in_cpo = c7.number_input("CPO (MAD)", value=12.5)
            in_calls = c8.number_input("Appels Support", value=1250)
            
            in_t_acc = st.number_input("Temps Acceptation (m)", value=3)
            in_t_pick = st.number_input("Temps Pickup (m)", value=12)
            in_t_traj = st.number_input("Temps Trajet (m)", value=27)
            in_reclamations = st.text_input("Ranking Réclamations", "1. Retard, 2. Missing item")

        with st.expander("🤝 Slide 3: Sales, AM & Forecast"):
            c1, c2, c3 = st.columns(3)
            in_signed = c1.number_input("Signés", value=signed_count)
            in_onboarded = c2.number_input("Onboardés", value=onboarded_count)
            in_churn = c3.number_input("Taux Churn (%)", value=2.1)
            
            in_topdeal = st.text_input("Top Deal", "McDonald's")
            in_kool = st.text_input("Performance KOOL", "42% du GMV Global / 38% du Volume")
            in_am = st.text_area("Actions AM", "Optimisation des menus Top 50.")
            
            if 'df_forecast' not in st.session_state:
                st.session_state.df_forecast = pd.DataFrame({"Restaurant": ["Tacos de Lyon", "Luigi"], "Glovo Orders": [185, 142], "Yassir Forecast": [100, 80]})
            edited_forecast = st.data_editor(st.session_state.df_forecast, num_rows="dynamic", use_container_width=True)

        with st.expander("📈 Slide 4: Performance Marketing"):
            c1, c2, c3 = st.columns(3)
            in_installs = c1.number_input("Installs", value=4500)
            in_ftu = c2.number_input("First Users", value=2100)
            in_cac = c3.number_input("CAC (MAD)", value=22)
            
            in_visits = st.number_input("Visites App", value=18500)
            in_campaign = st.text_input("Perf Campagnes", "Push notifications: CTR 12%")
            
            c4, c5 = st.columns(2)
            in_budget_promo = c4.number_input("Budget Promo", value=15000)
            in_gmv_promo = c5.number_input("GMV Promo", value=185000)
            in_mktg_budget = st.text_input("Expenses vs Budget", "Dépensé: 45K MAD / Budget: 60K MAD")

        with st.expander("📢 Slides 5, 6 & 7: Comms, Blocages, Roadmap"):
            in_social = st.text_area("Slide 5 - Social & PR", "Croissance commu: +12%\nTop post: Lancement resto\nRP: 3 articles presse")
            in_corp = st.text_area("Slide 5 - Corp & M+2", "Newsletter interne envoyée.\nCalendrier M+2: Activation Plages")
            in_bloc = st.text_area("Slide 6 - Blocages", "Bug technique temporaire.\nManque coursiers (pluie).\nActions: Surge pricing activé.")
            in_road = st.text_area("Slide 7 - Roadmap S+1", "1. Nouvelles zones.\n2. Grosse campagne marketing.")

    with tab3:
        st.markdown("### 📥 Génération PPTX")
        
        def generate_pptx():
            prs = Presentation(template_file) if template_file else Presentation()
            
            # Slide 1: Exec Summary
            vs_budget_gmv = ((kpis['gmv'] - in_budget_gmv) / in_budget_gmv * 100) if in_budget_gmv else 0
            req_vs_del = (kpis['orders'] / in_req * 100) if in_req else 0
            add_slide_with_bullet_points(prs, "01. Executive Summary", [
                f"Commandes: {kpis['orders']:,} | GMV: {kpis['gmv']:,.0f} MAD",
                f"Request vs Delivered: {req_vs_del:.1f}%",
                f"AOV: {kpis['aov']:.1f} MAD",
                f"Fréquence: {in_freq} | Weekdays: {in_weekdays}% / Weekend: {100-in_weekdays}%",
                f"Vs Budget (GMV): {vs_budget_gmv:+.1f}%",
                "--- HIGHLIGHTS ---"
            ] + in_high.split('\n'))

            # Slide 2: Opérations
            cancel_str = [f"{r[0]} ({r[1]})" for r in kpis['cancel_reasons']]
            add_slide_with_bullet_points(prs, "02. Operations & Delivery", [
                f"Temps moyen Click-to-Door: {kpis['avg_time']:.0f}m (Food: {in_t_food}m / Groceries: {in_t_groc}m)",
                f"Détail: Acc: {in_t_acc}m | Pick: {in_t_pick}m | Traj: {in_t_traj}m",
                f"Batching Rate: {in_batch} | Auto-dispatch: {in_auto}%",
                f"Taux d'annulation: {kpis['cancel_rate']:.1f}%",
                "Principales raisons: " + ", ".join(cancel_str),
                f"Coursiers actifs: {in_fleet} | Mix: {in_mix}% 3PL vs Indep",
                f"CPO: {in_cpo} MAD",
                f"Support Client: {in_calls} appels | Remb: ~1.2% | Réclamations: {in_reclamations}"
            ])

            # Slide 3: Sales
            slide3 = prs.slides.add_slide(prs.slide_layouts[1])
            slide3.shapes.title.text = "03. Sales & Account Management"
            tf3 = slide3.placeholders[1].text_frame if len(slide3.placeholders) > 1 else slide3.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(2)).text_frame
            tf3.text = f"Acquisition: {in_signed} Signés | {in_onboarded} Onboardés"
            tf3.add_paragraph().text = f"Top Deal: {in_topdeal}"
            tf3.add_paragraph().text = f"KOOL Exclusives: {in_kool}"
            tf3.add_paragraph().text = f"AM: Churn {in_churn}% | Actions: {in_am}"
            tf3.add_paragraph().text = "Top 10 Partenaires (Tableau ci-dessous):"
            
            t10_data = [[p['restaurant name'][:20], f"{p['orders']:,}", f"{p['gmv']:,.0f}"] for p in kpis['top_partners']]
            add_table_to_slide(slide3, t10_data, ["Partenaire", "Orders", "GMV"], Inches(0.5), Inches(3.5), Inches(9), Inches(1.5))

            # Slide 3b: Forecast
            slide3b = prs.slides.add_slide(prs.slide_layouts[1])
            slide3b.shapes.title.text = "03b. Orders Forecast (Nouvelles Signatures)"
            f_data = edited_forecast.dropna(how='all').values.tolist()
            add_table_to_slide(slide3b, f_data, ["Restaurant", "Glovo", "Yassir Forecast"], Inches(1), Inches(2), Inches(8), Inches(3))

            # Slide 4: Growth
            conv_rate = (kpis['orders'] / in_visits * 100) if in_visits else 0
            add_slide_with_bullet_points(prs, "04. Performance Marketing (Growth)", [
                f"Acquisition (Funnel): Visites {in_visits:,} -> Installs {in_installs:,} -> FTU {in_ftu:,}",
                f"Taux de Conversion (Visite->Cmd): {conv_rate:.1f}%",
                f"CAC Estimé: {in_cac} MAD",
                f"Campagnes: {in_campaign}",
                f"Codes Promo: {in_budget_promo:,} MAD alloués vs {in_gmv_promo:,} MAD générés",
                f"Expenses vs Budget: {in_mktg_budget}"
            ])

            # Slide 5: Comms
            add_slide_with_bullet_points(prs, "05. Communication & Brand", 
                ["--- RESEAUX & RP ---"] + in_social.split('\n') + 
                ["--- CORPORATE & M+2 ---"] + in_corp.split('\n')
            )

            # Slide 6: Blocages
            add_slide_with_bullet_points(prs, "06. Blocages & Difficultés", in_bloc.split('\n'))

            # Slide 7: Roadmap
            add_slide_with_bullet_points(prs, "07. Priorités & Roadmap S+1", in_road.split('\n'))

            binary_output = BytesIO()
            prs.save(binary_output)
            return binary_output.getvalue()

        if st.button("🎁 Générer la présentation PowerPoint (.pptx)"):
            pptx_data = generate_pptx()
            st.download_button("📥 Télécharger PPTX", data=pptx_data, file_name=f"WBR_LMD.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
